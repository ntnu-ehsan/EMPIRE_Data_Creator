"""
Region-specific solar and onshore-wind capacity factors from ERA5 via Atlite.

Why this exists
---------------
``generate_scenarios.py`` currently disaggregates ``solar.csv`` and
``windonshore.csv`` by *broadcast*: every NUTS region of a country receives the
**identical** national capacity-factor (CF) time series.  That removes all
intra-country spatial diversity (no coastal-vs-inland wind gradient, no
north-south solar gradient) and makes every region of a country perfectly
correlated.

This script replaces that with physically distinct, weather-consistent CF series
per region, computed from ERA5 reanalysis with `Atlite
<https://atlite.readthedocs.io>`_.  For each country it:

  1. loads the region polygons (NUTS2, or NUTS3 for selected countries),
  2. builds a tight ERA5 *cutout* around just that country (+ a small buffer),
  3. computes area-weighted PV and onshore-wind CF series per region,
  4. optionally **reconciles** the national average back to the original EMPIRE
     national CF (bias-correction that preserves spatial + temporal shape),
  5. **deletes the cutout** before moving to the next country/year.

Only **onshore** technologies are produced: ``solar`` and ``windonshore``.
Offshore wind is intentionally out of scope here.

Disk strategy
-------------
Cutouts are the only large artifact and are processed one country (and, by
default, one year) at a time, then deleted.  Peak scratch disk therefore holds a
single country-year — roughly:

  * NL / DK : <0.2 GB      * UK / FI : ~0.6 GB
  * SE / ES : ~1 GB        * NO      : ~2 GB

The per-region CF CSVs that are kept are only a few MB.  Pass
``--combined-years`` to build all years in one cutout per country (faster, more
CDS-friendly, but higher peak disk).

Spatial resolution
-------------------
Default is NUTS2.  Append ``:3`` to a country token to emit NUTS3 instead, e.g.
``ES:3`` produces Spanish CFs at NUTS3 (ES111, ES112, ...).

Prerequisites
-------------
  * Run under an environment with the Atlite stack, e.g. ``conda activate pypsa-eur``.
  * A Copernicus CDS account + ``~/.cdsapirc`` credentials
    (https://cds.climate.copernicus.eu/how-to-api).  Without it, only
    ``--dry-run`` works (it never downloads).

Usage
-----
    conda activate pypsa-eur

    # Inspect what would be downloaded (regions, boxes, size estimates) — no download:
    python Scripts/preprocess_renewables_atlite.py --dry-run

    # Full run for the North Sea set plus Spain at NUTS3:
    python Scripts/preprocess_renewables_atlite.py \
        --countries UK NO DK NL DE ES:3 \
        --years 2015-2019 \
        --output-dir data/Atlite/NorthSea

Output
------
    <output-dir>/solar.csv
    <output-dir>/windonshore.csv

Each has a ``time`` column (``DD/MM/YYYY HH:MM``, matching the EMPIRE scenario
CSVs) and one column per region code across all processed countries.  These can
then be consumed by ``generate_scenarios.py`` in place of the broadcast step.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

# Heavy geospatial / weather deps are imported lazily inside functions so that
# --help and argument parsing work even outside the Atlite environment.

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_SHP_DIR = _PROJECT_ROOT / 'data' / 'shp'

# Same multi-file fallback as config_loader: the 2024 vintage has revised NL
# codes but dropped all UK regions after Brexit; the 2021 vintage still carries
# the 41 UK NUTS2 regions.  First file to provide a (country, level) wins.
_SHAPEFILE_PATHS = [
    _SHP_DIR / 'NUTS_RG_20M_2024_3035.shp',
    _SHP_DIR / 'NUTS_RG_60M_2021_3035.geojson',
]

# EMPIRE scenario CSVs label some countries differently from their NUTS prefix.
# Used only for reconciliation against the original national CF series.
PREFIX_TO_CSV_COLUMN = {'UK': 'GB', 'NO': 'Norway'}

# -----------------------------------------------------------------------------
# Conversion assumptions (edit here to change technology characteristics)
# -----------------------------------------------------------------------------
WIND_TURBINE = 'Vestas_V112_3MW'          # representative onshore turbine
SOLAR_PANEL = 'CSi'                        # crystalline-silicon PV
SOLAR_ORIENTATION = {'slope': 35.0, 'azimuth': 180.0}  # fixed, south-facing
ERA5_MODULE = 'era5'
# Features needed for onshore PV + wind only (no runoff, no offshore):
#   influx + temperature -> PV;  wind -> onshore wind.
ERA5_FEATURES = ['influx', 'temperature', 'wind']
# Approx. number of time-varying variables stored, for the size estimate only.
_N_VARS_ESTIMATE = 9

# Remote/outermost territories excluded by default: they are irrelevant to the
# energy model but their geographic extent inflates the cutout bounding box by
# huge factors (e.g. Svalbard pushes Norway's box to 81 N; the Canaries push
# Spain's to 27 N / -18 E).  Matched by NUTS-code prefix (any level).
#   NO0B = Svalbard og Jan Mayen      ES70 = Canarias
#   ES63 = Ceuta   ES64 = Melilla     PT20 = Azores   PT30 = Madeira
#   FRY  = French overseas regions
DEFAULT_EXCLUDE_PREFIXES = ['NO0B', 'ES70', 'ES63', 'ES64', 'PT20', 'PT30', 'FRY']

# Default country set: North Sea project (NUTS2) + Spain at NUTS3.
DEFAULT_COUNTRIES = ['UK', 'NO', 'DK', 'NL', 'DE', 'ES:3']

# Countries flagged "large" only for the dry-run size warning.
_LARGE_COUNTRIES = {'NO', 'SE', 'ES', 'FR', 'UK', 'FI'}


# =============================================================================
# Country / region helpers
# =============================================================================

def parse_countries(tokens: list[str]) -> list[tuple[str, int]]:
    """Parse ``['UK', 'ES:3']`` -> ``[('UK', 2), ('ES', 3)]`` (level defaults to 2)."""
    out: list[tuple[str, int]] = []
    for tok in tokens:
        if ':' in tok:
            prefix, lvl = tok.split(':', 1)
            out.append((prefix.strip().upper(), int(lvl)))
        else:
            out.append((tok.strip().upper(), 2))
    return out


def load_regions(prefix: str, level: int, exclude_prefixes: list[str] | None = None):
    """Return a GeoDataFrame (index=region code, geometry in EPSG:4326) for one
    country at the requested NUTS level, using the 2024->2021 file fallback.

    Regions whose code starts with any entry in ``exclude_prefixes`` are dropped
    (used to strip remote/outermost territories that would inflate the bbox).
    """
    import geopandas as gpd

    exclude = tuple(exclude_prefixes or ())
    existing = [p for p in _SHAPEFILE_PATHS if p.exists()]
    if not existing:
        raise FileNotFoundError(
            f"No NUTS shapefile found. Looked for: {[str(p) for p in _SHAPEFILE_PATHS]}\n"
            "Download from: https://gisco-services.ec.europa.eu/distribution/v2/nuts/shp/"
        )

    for shp_path in existing:
        gdf = gpd.read_file(shp_path)
        sub = gdf[(gdf['CNTR_CODE'] == prefix) & (gdf['LEVL_CODE'] == level)]
        if not sub.empty:
            sub = sub[['NUTS_ID', 'geometry']].copy().to_crs(epsg=4326)
            sub = sub.set_index('NUTS_ID').sort_index()
            sub.index.name = 'region'
            if exclude:
                dropped = [c for c in sub.index if c.startswith(exclude)]
                if dropped:
                    print(f"  excluding remote regions: {', '.join(dropped)}")
                    sub = sub.drop(index=dropped)
            return sub
    raise ValueError(
        f"No NUTS{level} regions found for country '{prefix}' in any shapefile."
    )


def bbox_with_buffer(regions, buffer_deg: float) -> tuple[float, float, float, float]:
    """Return (x0, y0, x1, y1) lon/lat bounds of the regions, padded by buffer_deg."""
    minx, miny, maxx, maxy = regions.total_bounds
    return (minx - buffer_deg, miny - buffer_deg,
            maxx + buffer_deg, maxy + buffer_deg)


def estimate_gb(box: tuple[float, float, float, float], hours: int,
                dx: float = 0.25, dy: float = 0.25) -> float:
    """Rough netCDF size estimate (GB): cells * hours * vars * 4 bytes."""
    x0, y0, x1, y1 = box
    n_cells = max(1, round((x1 - x0) / dx)) * max(1, round((y1 - y0) / dy))
    return n_cells * hours * _N_VARS_ESTIMATE * 4 / 1e9


def parse_years(spec: str) -> list[int]:
    """Parse '2015-2019' or '2015,2017' or '2016' into a list of ints."""
    spec = spec.strip()
    if '-' in spec and ',' not in spec:
        a, b = spec.split('-', 1)
        return list(range(int(a), int(b) + 1))
    return [int(y) for y in spec.split(',')]


def _hours_in(year: int) -> int:
    return 8784 if (year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)) else 8760


# =============================================================================
# Atlite computation
# =============================================================================

def _compute_one_cutout(regions, x_slice, y_slice, time_spec, cutout_path: Path,
                        concurrent_requests: bool = False):
    """Build (download+prepare) a cutout for one time span, return per-region
    PV and onshore-wind CF DataFrames (index = UTC hourly, cols = region codes).
    The cutout netCDF is left on disk for the caller to delete.
    """
    import atlite

    cutout = atlite.Cutout(
        path=str(cutout_path),
        module=ERA5_MODULE,
        x=x_slice,
        y=y_slice,
        time=time_spec,
    )
    # Downloads from CDS on first call; cached in cutout_path thereafter.
    # monthly_requests=True splits each year into 12 per-month CDS requests so
    # none exceeds the CDS-Beta per-request cost limit (a single full-year
    # request returns "403 cost limits exceeded").
    cutout.prepare(
        features=ERA5_FEATURES,
        monthly_requests=True,
        concurrent_requests=concurrent_requests,
    )

    shapes = regions.geometry  # GeoSeries indexed by region code (EPSG:4326)

    # per_unit=True -> capacity factor (generation per unit installed capacity),
    # area-weighted across the grid cells overlapping each region polygon.
    solar = cutout.pv(
        panel=SOLAR_PANEL,
        orientation=SOLAR_ORIENTATION,
        shapes=shapes,
        per_unit=True,
    ).to_pandas()

    wind = cutout.wind(
        turbine=WIND_TURBINE,
        # Extend the power curve to the cutout's max wind speed rather than
        # extrapolating (Atlite-recommended; affects high-wind hours).
        add_cutout_windspeed=True,
        shapes=shapes,
        per_unit=True,
    ).to_pandas()

    # Guard against the single-region case returning a Series.
    if isinstance(solar, pd.Series):
        solar = solar.to_frame().T
    if isinstance(wind, pd.Series):
        wind = wind.to_frame().T

    return solar, wind


def _chunks(years: list[int], combined_years: bool):
    """Yield (time_spec, hours, label) per download chunk."""
    if combined_years:
        return [(slice(f"{years[0]}-01-01", f"{years[-1]}-12-31"),
                 sum(_hours_in(y) for y in years),
                 f"{years[0]}-{years[-1]}")]
    return [(str(y), _hours_in(y), str(y)) for y in years]


def _part_path(parts_dir: Path, tech: str, prefix: str, label: str) -> Path:
    return parts_dir / f"{tech}_{prefix}_{label}.csv"


def _write_part(df: pd.DataFrame, path: Path) -> None:
    """Write a per-(country, chunk) CF part. Index is kept as ISO timestamps so
    parts re-read and sort chronologically when assembled."""
    out = df.copy()
    out.index.name = 'time'
    path.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(path)


def _read_part(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, index_col='time', parse_dates=['time'])


def compute_country_parts(prefix: str, level: int, years: list[int],
                          combined_years: bool, scratch_dir: Path, buffer_deg: float,
                          parts_dir: Path, exclude_prefixes: list[str] | None = None,
                          concurrent_requests: bool = False,
                          overwrite: bool = False) -> list[str]:
    """Ensure per-(country, chunk) CF part files exist, computing missing ones.

    Each chunk is written to ``parts_dir`` as soon as it is computed and the
    cutout is deleted, so an interrupted run resumes from the last completed
    chunk (existing parts are skipped unless ``overwrite``).  Returns the list
    of chunk labels for this country (used to assemble the full series).
    """
    regions = load_regions(prefix, level, exclude_prefixes)
    box = bbox_with_buffer(regions, buffer_deg)
    x_slice = slice(box[0], box[2])
    y_slice = slice(box[1], box[3])

    print(f"\n=== {prefix} (NUTS{level}): {len(regions)} regions ===")
    print(f"  bbox lon[{box[0]:.2f}, {box[2]:.2f}] lat[{box[1]:.2f}, {box[3]:.2f}]")

    labels: list[str] = []
    for time_spec, hours, label in _chunks(years, combined_years):
        labels.append(label)
        sp = _part_path(parts_dir, 'solar', prefix, label)
        wp = _part_path(parts_dir, 'windonshore', prefix, label)
        if sp.exists() and wp.exists() and not overwrite:
            print(f"  chunk {label}: parts already exist -> skip (resume)")
            continue

        size = estimate_gb(box, hours)
        print(f"  chunk {label}: ~{hours} h, est. cutout ~{size:.2f} GB ... ",
              end='', flush=True)

        cutout_path = scratch_dir / f"cutout_{prefix}_{label}.nc"
        try:
            solar, wind = _compute_one_cutout(regions, x_slice, y_slice,
                                              time_spec, cutout_path,
                                              concurrent_requests)
            _write_part(solar, sp)
            _write_part(wind, wp)
            print("done")
        finally:
            # Reclaim disk before the next country/year regardless of success.
            if cutout_path.exists():
                try:
                    cutout_path.unlink()
                except OSError as e:
                    print(f"\n  WARNING: could not delete {cutout_path}: {e}")

    return labels


def assemble_country(parts_dir: Path, prefix: str, labels: list[str]):
    """Concatenate a country's per-chunk parts into full solar/wind DataFrames
    (DatetimeIndex spanning all years, one column per region)."""
    def _load(tech: str) -> pd.DataFrame:
        parts = [_read_part(_part_path(parts_dir, tech, prefix, lbl)) for lbl in labels]
        df = pd.concat(parts).sort_index() if len(parts) > 1 else parts[0]
        return df
    return _load('solar'), _load('windonshore')


# =============================================================================
# Reconciliation against the original EMPIRE national CF
# =============================================================================

def reconcile_to_national(df: pd.DataFrame, prefix: str, national_csv: Path) -> pd.DataFrame:
    """Scale every region series by one factor so the country mean reproduces
    the original EMPIRE national mean CF.

    This corrects ERA5/turbine level bias while preserving the spatial pattern
    (relative differences between regions) and the temporal shape.  A single
    multiplicative factor is used, so annual energy matches but no information is
    lost beyond an overall level adjustment.  Values are clipped to [0, 1].
    """
    if not national_csv.exists():
        print(f"  reconcile: national file {national_csv.name} not found -> skipped")
        return df

    col = PREFIX_TO_CSV_COLUMN.get(prefix, prefix)
    nat = pd.read_csv(national_csv)
    if col not in nat.columns:
        print(f"  reconcile: column '{col}' not in {national_csv.name} -> skipped")
        return df

    nat_mean = float(nat[col].astype(float).mean())
    atlite_mean = float(df.values.mean())
    if atlite_mean <= 0:
        print(f"  reconcile: Atlite mean is 0 for {prefix} -> skipped")
        return df

    factor = nat_mean / atlite_mean
    print(f"  reconcile {prefix} via '{col}': national={nat_mean:.4f} "
          f"atlite={atlite_mean:.4f} -> factor={factor:.4f}")
    return (df * factor).clip(lower=0.0, upper=1.0)


# =============================================================================
# Output assembly
# =============================================================================

def to_empire_time_format(df: pd.DataFrame) -> pd.DataFrame:
    """Move the DatetimeIndex into a 'time' column formatted DD/MM/YYYY HH:MM."""
    out = df.copy()
    out.insert(0, 'time', out.index.strftime('%d/%m/%Y %H:%M'))
    return out.reset_index(drop=True)


# =============================================================================
# Main
# =============================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description='Region-specific solar & onshore-wind CFs from ERA5 (Atlite).')
    p.add_argument('--countries', nargs='+', default=DEFAULT_COUNTRIES,
                   help="Country tokens; append ':3' for NUTS3 "
                        f"(default: {' '.join(DEFAULT_COUNTRIES)}).")
    p.add_argument('--years', default='2015-2019',
                   help="Weather years, e.g. '2015-2019' or '2016' (default: 2015-2019).")
    p.add_argument('--combined-years', action='store_true',
                   help='Build all years in one cutout per country (faster, more disk).')
    p.add_argument('--buffer-deg', type=float, default=0.5,
                   help='Bounding-box padding in degrees (default: 0.5).')
    p.add_argument('--exclude', nargs='*', default=None,
                   help='Extra NUTS code prefixes to drop, added to the remote-territory '
                        f'defaults ({", ".join(DEFAULT_EXCLUDE_PREFIXES)}).')
    p.add_argument('--include-remote', action='store_true',
                   help='Keep remote/outermost territories (Svalbard, Canaries, ...) '
                        'that are excluded by default.')
    p.add_argument('--output-dir', default='data/Atlite/NorthSea',
                   help='Directory for solar.csv / windonshore.csv (default: data/Atlite/NorthSea).')
    p.add_argument('--scratch-dir', default='data/Atlite/_cutouts',
                   help='Where cutouts are built and then deleted (default: data/Atlite/_cutouts).')
    p.add_argument('--empire-scenario-dir', default=None,
                   help='ScenarioData dir with the original national solar.csv/windonshore.csv '
                        'for reconciliation. Default: derived from config.yaml empire_folder.')
    p.add_argument('--no-reconcile', action='store_true',
                   help='Skip national-mean reconciliation (keep raw Atlite levels).')
    p.add_argument('--concurrent-requests', action='store_true',
                   help='Submit the 12 monthly CDS requests in parallel (faster, but '
                        'more likely to hit CDS queue limits). Default: sequential.')
    p.add_argument('--overwrite', action='store_true',
                   help='Recompute and overwrite cached per-year CF parts instead of '
                        'resuming from them.')
    p.add_argument('--dry-run', action='store_true',
                   help='List regions, boxes and size estimates without downloading.')
    return p.parse_args()


def _default_empire_scenario_dir() -> Path | None:
    """Derive ScenarioData dir from config.yaml empire_folder (best effort)."""
    try:
        import yaml
    except ImportError:
        return None
    cfg_path = _PROJECT_ROOT / 'config.yaml'
    if not cfg_path.exists():
        return None
    with open(cfg_path, encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}
    folder = cfg.get('empire_folder')
    if not folder:
        return None
    cand = _PROJECT_ROOT / 'data' / 'EMPIRE' / folder / 'ScenarioData'
    return cand if cand.exists() else None


def main() -> None:
    args = parse_args()
    countries = parse_countries(args.countries)
    years = parse_years(args.years)

    if args.include_remote:
        exclude_prefixes: list[str] = []
    else:
        exclude_prefixes = list(DEFAULT_EXCLUDE_PREFIXES) + list(args.exclude or [])

    output_dir = (_PROJECT_ROOT / args.output_dir
                  if not Path(args.output_dir).is_absolute() else Path(args.output_dir))
    scratch_dir = (_PROJECT_ROOT / args.scratch_dir
                   if not Path(args.scratch_dir).is_absolute() else Path(args.scratch_dir))

    if args.empire_scenario_dir:
        empire_dir = Path(args.empire_scenario_dir)
    else:
        empire_dir = _default_empire_scenario_dir()

    print("=" * 64)
    print("Atlite renewable CF preprocessing (onshore solar + wind)")
    print("=" * 64)
    print(f"  Countries  : {', '.join(f'{p}(NUTS{l})' for p, l in countries)}")
    print(f"  Years      : {years[0]}-{years[-1]}  "
          f"({'combined' if args.combined_years else 'year-by-year'})")
    print(f"  Output dir : {output_dir}")
    print(f"  Scratch dir: {scratch_dir}")
    print(f"  Reconcile  : {'OFF' if args.no_reconcile else f'ON (vs {empire_dir})'}")
    print(f"  Turbine    : {WIND_TURBINE}   Panel: {SOLAR_PANEL}")

    # --- credential preflight ----------------------------------------------
    if not args.dry_run:
        cdsrc = Path.home() / '.cdsapirc'
        if not cdsrc.exists():
            print(f"\nERROR: CDS credentials not found at {cdsrc}.\n"
                  "Set up an account and API key: https://cds.climate.copernicus.eu/how-to-api\n"
                  "Re-run with --dry-run to validate everything else without downloading.")
            sys.exit(1)

    # --- dry run: just report geometry + size ------------------------------
    if args.dry_run:
        total_hours = sum(_hours_in(y) for y in years)
        print("\n--- DRY RUN (no download) ---")
        grand = 0.0
        for prefix, level in countries:
            regions = load_regions(prefix, level, exclude_prefixes)
            box = bbox_with_buffer(regions, args.buffer_deg)
            peak = (estimate_gb(box, total_hours) if args.combined_years
                    else estimate_gb(box, 8784))  # worst single (leap) year
            grand += estimate_gb(box, total_hours)
            flag = '  <-- large' if prefix in _LARGE_COUNTRIES else ''
            print(f"  {prefix} NUTS{level}: {len(regions):3d} regions  "
                  f"bbox[{box[0]:6.1f},{box[1]:5.1f} .. {box[2]:6.1f},{box[3]:5.1f}]  "
                  f"peak ~{peak:5.2f} GB{flag}")
        mode = 'combined-years' if args.combined_years else 'year-by-year'
        print(f"\n  Peak scratch disk ({mode}) = the largest single value above.")
        print(f"  Total download volume across all chunks ~ {grand:.1f} GB "
              "(transient; cutouts are deleted as they are consumed).")
        print("\n  Dry run OK. Add CDS credentials and drop --dry-run to compute CFs.")
        return

    # --- real run ----------------------------------------------------------
    scratch_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    parts_dir = output_dir / '_parts'
    print(f"  Parts cache: {parts_dir} (per-year; enables resume)")

    solar_all: list[pd.DataFrame] = []
    wind_all: list[pd.DataFrame] = []

    for prefix, level in countries:
        labels = compute_country_parts(
            prefix, level, years, args.combined_years, scratch_dir, args.buffer_deg,
            parts_dir, exclude_prefixes, args.concurrent_requests, args.overwrite)
        solar_df, wind_df = assemble_country(parts_dir, prefix, labels)

        if not args.no_reconcile and empire_dir is not None:
            solar_df = reconcile_to_national(solar_df, prefix, empire_dir / 'solar.csv')
            wind_df = reconcile_to_national(wind_df, prefix, empire_dir / 'windonshore.csv')

        solar_all.append(solar_df)
        wind_all.append(wind_df)

    # Concatenate regions side by side (same time index across countries).
    solar_out = pd.concat(solar_all, axis=1)
    wind_out = pd.concat(wind_all, axis=1)

    solar_path = output_dir / 'solar.csv'
    wind_path = output_dir / 'windonshore.csv'
    to_empire_time_format(solar_out).to_csv(solar_path, index=False)
    to_empire_time_format(wind_out).to_csv(wind_path, index=False)

    print(f"\nSaved {solar_out.shape[1]} region cols x {len(solar_out)} h -> {solar_path}")
    print(f"Saved {wind_out.shape[1]} region cols x {len(wind_out)} h -> {wind_path}")
    print("\nDone. Point generate_scenarios.py at these files to replace the broadcast step.")


if __name__ == '__main__':
    main()
