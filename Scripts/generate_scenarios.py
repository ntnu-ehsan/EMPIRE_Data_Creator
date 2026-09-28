"""
Generate EMPIRE scenario CSV files with NUTS2 spatial disaggregation.

This script reads the raw EMPIRE scenario time-series CSV files and replaces
selected country-level columns with their constituent NUTS2 regions.

Two disaggregation methods are supported:

  electricload.csv
      Load is distributed using the ``Final Energy|Electricity [Share]``
      variable from an IAMC data file.  Values are in percent; they are
      normalised within each country before use.

  hydroror.csv  /  solar.csv  /  windonshore.csv
    These files contain capacity factors, so every NUTS2 region of each
    selected country receives the **same** value as its parent country column
      (broadcast/replicate).  No normalisation or weighting is applied.
    For windonshore.csv all NUTS2 regions of selected countries are included.

    solar.csv and windonshore.csv can instead be spatially differentiated, for
    the countries named in the ``renewable_cf`` section of config.yaml, by
    rescaling the national series per region (data/Renewable/...) — or replaced
    outright by per-region Atlite series when ``--atlite-dir`` is given.
    Offshore wind is always broadcast; there is no per-region CF source for it.

  hydroseasonal.csv
      Absolute reservoir-hydro inflow is distributed by explicit regional
      ``Capacity|Electricity|Hydro|Reservoir`` capacity. Bare ``Hydro`` is used
      only as a legacy fallback for a country with no explicit reservoir data.

  windoffshore.csv
      Capacity factors, broadcast like above, but only NUTS2 regions that
      have Offshore wind capacity according to ``wind_nuts2_DK_DE.csv``
      are included as output columns.

In both cases NUTS2 columns are inserted **in place** of the original country
column, preserving the exact positions of every other column (including
metadata columns that may appear mid-row in some files).

Usage
-----
    python generate_scenarios.py [options]

Options
-------
    --iamc        Path to the IAMC CSV file.
                  Default: data/IAMC/NorthSea/Denmark_Germany.csv
    --scenario    Scenario name inside the IAMC file.
                  Default: Transmission v1.0
    --input-dir   Directory that contains the raw EMPIRE scenario CSV files.
                  Default: data/EMPIRE/v_52/ScenarioData
    --output-dir  Directory where disaggregated CSV files are written.
                  Default: data/EMPIRE/v_52/ScenarioData_NUTS2

Example
-------
    conda activate IAMC_format
    python generate_scenarios.py

Notes
-----
- Shares are always normalised within each country so that rounding errors
  never introduce energy imbalances in the output.
- The column order of every other country (AT, BE, ...) and all metadata
  columns (time, hour, dayofweek, month) is preserved exactly.
- hydroror.csv places GR, HR, LU *after* the metadata columns.  The generic
  disaggregation engine handles this correctly without any special casing.
"""

import argparse
import shutil
import sys
from pathlib import Path

import pandas as pd

try:
    import pyam
except ImportError:
    print("ERROR: pyam-iamc not installed. Run: pip install pyam-iamc")
    import sys; sys.exit(1)

# =============================================================================
# Constants
# =============================================================================

IAMC_VAR_SHARE = 'Final Energy|Electricity [Share]'

# NUTS2 configuration — loaded from config.yaml in the project root.
# Edit config.yaml to change which countries are represented at NUTS2 level.
from config_loader import load_nuts2_config as _load_nuts2_config
_cfg = _load_nuts2_config()
NUTS2_COUNTRY_PREFIXES = _cfg['NUTS2_COUNTRY_PREFIXES']

# Some EMPIRE scenario CSVs label a country column differently from its NUTS2
# prefix.  Great Britain is 'GB' in the scenario CSVs, but its NUTS2 regions use
# the 'UK' prefix (UKC1, UKD1, ...).  Map NUTS2 prefix -> CSV column name here.
PREFIX_TO_CSV_COLUMN = {'UK': 'GB', 'NO': 'Norway'}

# Resource-potential / capacity-factor data (data/Renewable/...).  When a
# configured country has data there, solar.csv and windonshore.csv are
# differentiated across its NUTS regions by rescaling the national CF series
# instead of broadcasting it verbatim.  See config_loader.load_renewable_cf_config
# for the 'relative' vs 'absolute' modes.
from config_loader import load_renewable_cf_config as _load_renewable_cf_config
_rc_cfg = _load_renewable_cf_config()
RENEWABLE_DIR = _rc_cfg['dir']
RENEWABLE_MODE = _rc_cfg['mode']
RENEWABLE_COUNTRIES = _rc_cfg['countries']

# Per-country onshore-wind CF source.  Both layouts carry the same columns
# (NUTS_ID, country, ref_MW, ref_GWh); 'Wind.csv' is the Spain NUTS3 export and
# 'onshore.csv' the NESP NUTS2 one.  First file present wins.
WIND_CF_FILENAMES = ('onshore.csv', 'Wind.csv')
SOLAR_CF_FILENAME = 'Solar.xlsx'


def load_nuts2_regions_from_sets(
    sets_file: str,
    country_prefixes: list[str],
) -> dict[str, list[str]]:
    """Read NUTS2 regions per selected country from Sets.xlsx Nodes sheet.

    This is used as a robust fallback when IAMC regional share rows are missing
    for a selected country.
    """
    nodes_df = pd.read_excel(sets_file, sheet_name='Nodes')
    if 'Node' not in nodes_df.columns:
        raise ValueError("Nodes sheet in Sets file must contain a 'Node' column")

    node_values = [str(v) for v in nodes_df['Node'].dropna().tolist()]
    result: dict[str, list[str]] = {}
    for prefix in country_prefixes:
        result[prefix] = sorted(
            n for n in node_values
            if n.startswith(prefix) and len(n) >= 4
        )
    return result

# Files that require disaggregation (others are copied unchanged).
DISAGGREGATE_FILES = {
    'electricload.csv':   'iamc_share',
    'hydroror.csv':       'hydro_ror_sets',
    'hydroseasonal.csv':  'hydro_capacity',
    'solar.csv':          'broadcast',
    'windoffshore.csv':   'wind_offshore_broadcast',
    'windonshore.csv':    'broadcast',
}


# =============================================================================
# Share-loading helpers
# =============================================================================

def _iamc_long(iamc_file: str, scenario: str, variable: str) -> 'pd.DataFrame':
    """Load an IAMC file (CSV or Excel) and return long-format rows filtered
    to the given scenario and variable, excluding hierarchical region entries."""
    from iamc_utils import load_iamc
    idf = load_iamc(iamc_file)
    data = idf.filter(scenario=scenario, variable=variable).data
    return data[~data['region'].str.contains('>', na=False)].copy()


def load_iamc_shares(
    iamc_file: str,
    scenario: str,
    nuts2_regions_by_country: dict[str, list[str]],
) -> dict[str, dict[str, float]]:
    """Return normalised fractional shares for selected NUTS2 countries.

    Shares are derived from ``Final Energy|Electricity [Share]`` in the IAMC
    file (percent values, one reference year).

    Returns
    -------
    ``{'DE': {nuts2: fraction, ...}, 'DK': {...}, 'NL': {...}}``
    """
    filtered = _iamc_long(iamc_file, scenario, IAMC_VAR_SHARE)

    if filtered.empty:
        print(f"ERROR: No data found for variable='{IAMC_VAR_SHARE}' "
              f"in scenario='{scenario}'.")
        sys.exit(1)

    shares: dict[str, dict[str, float]] = {}

    for country_col in NUTS2_COUNTRY_PREFIXES:
        nuts_prefix = country_col
        rows = filtered[filtered['region'].str.startswith(nuts_prefix)]
        if rows.empty:
            fallback_regions = nuts2_regions_by_country.get(country_col, [])
            if fallback_regions:
                equal = 1.0 / len(fallback_regions)
                shares[country_col] = {r: equal for r in fallback_regions}
                print(f"  WARNING: No IAMC share data for prefix='{nuts_prefix}'. "
                      f"Using equal split across {len(fallback_regions)} "
                      "NUTS2 regions from Sets file.")
            else:
                print(f"  WARNING: No IAMC share data for prefix='{nuts_prefix}', "
                      "and no fallback NUTS2 regions in Sets file.")
            continue

        raw = dict(zip(rows['region'], rows['value'].astype(float)))
        total = sum(raw.values())
        if total == 0:
            fallback_regions = nuts2_regions_by_country.get(country_col, [])
            if fallback_regions:
                equal = 1.0 / len(fallback_regions)
                shares[country_col] = {r: equal for r in fallback_regions}
                print(f"  WARNING: IAMC shares for {country_col} sum to 0. "
                      f"Using equal split across {len(fallback_regions)} "
                      "NUTS2 regions from Sets file.")
            else:
                print(f"  ERROR: IAMC shares for {country_col} sum to 0 and "
                      "no fallback NUTS2 regions are available.")
                sys.exit(1)
        else:
            shares[country_col] = {r: v / total for r, v in sorted(raw.items())}
            print(f"  {country_col}: {len(raw)} NUTS2 regions "
                  f"(raw sum={total:.4f} %, normalised to 1.0)")

    return shares


def load_broadcast_shares(
    iamc_file: str,
    scenario: str,
    nuts2_regions_by_country: dict[str, list[str]],
) -> dict[str, dict[str, float]]:
    """Return shares of 1.0 for every NUTS2 region of selected countries.

    Used for files that contain capacity factors (e.g. hydroror.csv,
    solar.csv): every NUTS2 region should receive the **same** value as
    its parent country column rather than a fraction of it.

    The NUTS2 region list is derived from the
    ``Final Energy|Electricity [Share]`` variable in the IAMC file.

    Returns
    -------
    ``{'DE': {nuts2: 1.0, ...}, 'DK': {...}, 'NL': {...}}``
    """
    df = _iamc_long(iamc_file, scenario, IAMC_VAR_SHARE)

    result: dict[str, dict[str, float]] = {}
    for country_col in NUTS2_COUNTRY_PREFIXES:
        nuts_prefix = country_col
        rows = df[df['region'].str.startswith(nuts_prefix)]
        if rows.empty:
            nuts2 = nuts2_regions_by_country.get(country_col, [])
            if not nuts2:
                print(f"  WARNING: No NUTS2 region list found for prefix='{nuts_prefix}'; "
                      "column will be skipped if present.")
                continue
            print(f"  WARNING: No IAMC regional rows for prefix='{nuts_prefix}'. "
                  f"Using {len(nuts2)} NUTS2 regions from Sets file.")
        else:
            nuts2 = sorted(rows['region'].unique())
        result[country_col] = {r: 1.0 for r in nuts2}
        print(f"  {country_col}: {len(nuts2)} NUTS2 regions -- "
              "each region copies the country capacity factor (share=1.0)")

    return result


def _resolve_wind_cf_file(renewable_dir: Path) -> Path | None:
    """Return the onshore-wind CF source inside ``renewable_dir``, or None.

    Accepts either supported layout name (see :data:`WIND_CF_FILENAMES`); both
    carry the same columns, so the caller needs no further branching.
    """
    for name in WIND_CF_FILENAMES:
        candidate = renewable_dir / name
        if candidate.exists():
            return candidate
    return None


def load_cf_targets(
    regions_by_country: dict[str, list[str]],
    renewable_dir: Path | None = RENEWABLE_DIR,
    countries: list[str] | None = None,
) -> dict[str, dict[str, float]]:
    """Return per-region capacity factors, keyed by scenario CSV file.

    ``{'solar.csv': {region: cf, ...}, 'windonshore.csv': {region: cf, ...}}``

    * windonshore.csv : CF derived from the wind file as
                        ``ref_GWh*1000 / (ref_MW*8760)`` (full-load hours / 8760),
                        per region directly.  Regions with ``ref_MW == 0``
                        (Svalbard, Åland) yield no CF and stay broadcast.
    * solar.csv       : NUTS2 ``Capacity_Factor`` from Solar.xlsx assigned to the
                        region itself (NUTS2 runs) or to each of its modeled
                        NUTS3 children (CF is intensive -> no split).

    Only countries listed in ``countries`` are considered.  Regions without data
    are simply omitted; the rescaling step falls back to the national broadcast
    value for them.  Empty dicts are returned when the source files are missing
    so callers can fall back to broadcast entirely.
    """
    if renewable_dir is None or not countries:
        return {'solar.csv': {}, 'windonshore.csv': {}}

    # Region -> country, restricted to the countries this feature covers.
    region_country: dict[str, str] = {}
    for country in countries:
        for region in regions_by_country.get(country, []):
            region_country[str(region)] = country
    if not region_country:
        print(f"  [scenarios] note: none of {', '.join(countries)} are modeled "
              "at NUTS level — CF rescaling is inactive.")
        return {'solar.csv': {}, 'windonshore.csv': {}}

    solar_cf: dict[str, float] = {}
    wind_cf: dict[str, float] = {}

    wind_path = _resolve_wind_cf_file(renewable_dir)
    if wind_path is not None:
        w = pd.read_csv(wind_path)
        for _, r in w.iterrows():
            rid = str(r['NUTS_ID'])
            if rid not in region_country:
                continue
            mw = pd.to_numeric(r.get('ref_MW'), errors='coerce')
            gwh = pd.to_numeric(r.get('ref_GWh'), errors='coerce')
            if pd.notna(mw) and mw > 0 and pd.notna(gwh):
                wind_cf[rid] = float(gwh) * 1000.0 / (float(mw) * 8760.0)
        print(f"  windonshore.csv: {len(wind_cf)} regions with CF data "
              f"from {wind_path.name}")
    else:
        print(f"  [scenarios] note: none of {'/'.join(WIND_CF_FILENAMES)} found "
              f"in {renewable_dir} — windonshore stays broadcast.")

    solar_path = renewable_dir / SOLAR_CF_FILENAME
    if solar_path.exists():
        s = pd.read_excel(solar_path)
        # Keyed by (country, NUTS2) so a region only ever picks up its own
        # country's row even if two classifications share a code.
        n2_cf: dict[tuple[str, str], float] = {}
        for _, r in s.iterrows():
            code = str(r.get('Code'))
            cf = pd.to_numeric(r.get('Capacity_Factor'), errors='coerce')
            if code in countries and pd.notna(cf):
                n2_cf[(code, str(r['NUTS2 NEW']))] = float(cf)
        for rid, country in region_country.items():
            # NUTS2 codes are 4 chars and map to themselves; NUTS3 codes are 5
            # and map to their NUTS2 parent.
            parent = rid[:4]
            if (country, parent) in n2_cf:
                solar_cf[rid] = n2_cf[(country, parent)]
        print(f"  solar.csv: {len(solar_cf)} regions with CF data "
              f"from {solar_path.name}")
        no_solar = sorted(
            c for c in countries
            if any(region_country.get(r) == c for r in region_country)
            and not any(region_country[r] == c for r in solar_cf)
        )
        if no_solar:
            print(f"  [scenarios] note: {SOLAR_CF_FILENAME} has no rows for "
                  f"{', '.join(no_solar)} — their solar stays broadcast.")
    else:
        print(f"  [scenarios] note: {solar_path} not found — solar "
              "stays broadcast.")

    return {'solar.csv': solar_cf, 'windonshore.csv': wind_cf}


# Generator type in the produced Generator.xlsx per capacity-factor file, used
# to capacity-weight the CF reconciliation (RefInitialCap per NUTS region).
ATLITE_RECONCILE_TECH = {'solar.csv': 'Solar', 'windonshore.csv': 'Wind onshore'}


def load_capacity_weights(
    gen_file: Path,
    tech: str,
    regions: list[str],
) -> dict[str, float]:
    """Normalised per-region capacity weights from Generator.xlsx.

    Weights are ``RefInitialCap(region, tech) / Σ RefInitialCap`` for the given
    generator type across ``regions`` — the same reference installed capacity
    that :mod:`preprocess_eu` uses to build the EU node.  These weight each
    region's contribution when regional CFs are reconciled back to the EMPIRE
    national level, so annual energy is conserved given the model's own capacity
    split.  Falls back to an equal split ({r: 1/N}) when the sheet is missing or
    the technology total is zero.
    """
    equal = {r: 1.0 / len(regions) for r in regions} if regions else {}
    if not gen_file or not Path(gen_file).exists():
        return equal
    try:
        df = pd.read_excel(gen_file, sheet_name='RefInitialCap')
    except Exception:
        return equal
    if df.shape[1] < 3:
        return equal
    nc, gc, vc = df.columns[:3]
    reg_set = set(regions)
    sub = df[(df[gc] == tech) & (df[nc].astype(str).isin(reg_set))]
    caps = {
        str(r): float(v)
        for r, v in zip(sub[nc], pd.to_numeric(sub[vc], errors='coerce').fillna(0.0))
    }
    total = sum(caps.get(r, 0.0) for r in regions)
    if total <= 0:
        return equal
    return {r: caps.get(r, 0.0) / total for r in regions}


def load_wind_offshore_shares_from_sets(
    sets_file: str,
) -> dict[str, dict[str, float]]:
    """Return shares of 1.0 for NUTS2 regions that have an offshore-wind
    generator registered in the Sets file's GeneratorsOfNode sheet.

    Only nodes listed there receive a column in the disaggregated
    windoffshore.csv, so the scenario output is always consistent with
    what EMPIRE's GeneratorsOfNode recognises.

    Returns
    -------
    ``{'DE': {nuts2: 1.0, ...}, 'DK': {...}, 'NL': {...}}``
    """
    raw = pd.read_excel(sets_file, sheet_name='GeneratorsOfNode', header=None)
    headers = raw.iloc[2].tolist()
    data = raw.iloc[3:].copy()
    data.columns = headers[:data.shape[1]]
    data = data.reset_index(drop=True)

    # Keep only rows where the generator name contains 'offshore' (case-insensitive)
    offshore_mask = data['Generator'].str.lower().str.contains('offshore', na=False)
    offshore_nodes = data.loc[offshore_mask, 'Node'].unique()

    result: dict[str, dict[str, float]] = {}
    for country_col in NUTS2_COUNTRY_PREFIXES:
        nuts_prefix = country_col
        regions = sorted(
            r for r in offshore_nodes if str(r).startswith(nuts_prefix)
        )
        if not regions:
            print(f"  {country_col}: no offshore-wind NUTS2 nodes found in "
                  f"GeneratorsOfNode of {sets_file} -- column will be skipped.")
            continue
        result[country_col] = {r: 1.0 for r in regions}
        print(f"  {country_col}: {len(regions)} offshore NUTS2 nodes "
              f"(from GeneratorsOfNode): {', '.join(regions)}")

    return result


def _read_sets_generators(sets_file: str) -> pd.DataFrame:
    """Read GeneratorsOfNode from a Sets.xlsx, returning a DataFrame with
    columns ['Node', 'Generator']."""
    raw = pd.read_excel(sets_file, sheet_name='GeneratorsOfNode', header=None)
    headers = raw.iloc[2].tolist()
    data = raw.iloc[3:].copy()
    data.columns = headers[:data.shape[1]]
    return data.reset_index(drop=True)


def load_hydro_ror_shares_from_sets(
    sets_file: str,
) -> dict[str, dict[str, float]]:
    """Return shares of 1.0 for NUTS2 nodes that have 'Hydro run-of-the-river'
    in GeneratorsOfNode.  Only those nodes receive a column in hydroror.csv.

    Returns
    -------
    ``{'DE': {nuts2: 1.0, ...}, 'DK': {...}, 'NL': {...}}``
    """
    data = _read_sets_generators(sets_file)
    ror_nodes = data.loc[
        data['Generator'] == 'Hydro run-of-the-river', 'Node'
    ].unique()

    result: dict[str, dict[str, float]] = {}
    for country_col in NUTS2_COUNTRY_PREFIXES:
        nuts_prefix = country_col
        regions = sorted(r for r in ror_nodes if str(r).startswith(nuts_prefix))
        if not regions:
            print(f"  {country_col}: no 'Hydro run-of-the-river' NUTS2 nodes in "
                  f"GeneratorsOfNode -- column will be skipped if present.")
            continue
        result[country_col] = {r: 1.0 for r in regions}
        print(f"  {country_col}: {len(regions)} hydro-ROR NUTS2 nodes "
              f"(from GeneratorsOfNode): {', '.join(regions)}")
    return result


def load_hydro_capacity_shares(
    iamc_file: str,
    scenario: str,
    sets_file: str,
) -> dict[str, dict[str, float]]:
    """Return capacity-weighted fractional shares for hydroseasonal disaggregation.

    Only NUTS2 nodes that have 'Hydro regulated' in GeneratorsOfNode receive
    a column — this keeps the scenario output consistent with the Sets file and
    prevents Pyomo from rejecting invalid (node, generator) indices.

    Returns
    -------
    ``{'DE': {nuts2: fraction, ...}, 'DK': {...}, 'NL': {...}}``
    """
    # Valid NUTS2 nodes from GeneratorsOfNode (Hydro regulated only)
    gon = _read_sets_generators(sets_file)
    reg_nodes = set(gon.loc[gon['Generator'] == 'Hydro regulated', 'Node'].unique())

    regions_by_country = {
        country_col: sorted(
            n for n in reg_nodes if str(n).startswith(country_col)
        )
        for country_col in NUTS2_COUNTRY_PREFIXES
    }
    from iamc_utils import load_hydro_capacity_weights
    hydro_weights, hydro_sources = load_hydro_capacity_weights(
        iamc_file, scenario, regions_by_country, 'seasonal_reservoir')

    result: dict[str, dict[str, float]] = {}

    for country_col in NUTS2_COUNTRY_PREFIXES:
        # Only include nodes registered in GeneratorsOfNode for this country.
        all_nuts2 = regions_by_country[country_col]
        if not all_nuts2:
            print(f"  {country_col}: no 'Hydro regulated' NUTS2 nodes in "
                  "GeneratorsOfNode -- column will be skipped if present.")
            continue

        # Unit-normalised reservoir-capacity values from IAMC.
        cap_map = {r: hydro_weights.get(r, 0.0) for r in all_nuts2}
        cap_total = sum(cap_map.get(r, 0.0) for r in all_nuts2)

        if cap_total == 0:
            # No hydro capacity data for this country -> broadcast evenly.
            n = len(all_nuts2)
            shares = {r: 1.0 / n for r in all_nuts2}
            print(f"  {country_col}: {len(shares)} NUTS2 nodes (Sets-filtered) -- "
                  f"{hydro_sources[country_col]} -> equal split")
        else:
            shares = {r: cap_map.get(r, 0.0) / cap_total for r in all_nuts2}
            n_nonzero = sum(1 for v in shares.values() if v > 0)
            print(f"  {country_col}: {len(shares)} NUTS2 nodes (Sets-filtered), "
                  f"{n_nonzero} with non-zero hydro capacity "
                  f"({hydro_sources[country_col]}, total={cap_total:.1f} MW, "
                  "normalised to 1.0)")

        result[country_col] = shares

    return result


# =============================================================================
# Generic CSV disaggregation engine
# =============================================================================

def disaggregate_csv(
    input_path: Path,
    output_path: Path,
    shares: dict[str, dict[str, float]],
    *,
    always_drop: tuple[str, ...] = tuple(NUTS2_COUNTRY_PREFIXES),
) -> None:
    """Replace country columns with NUTS2 columns, preserving column order.

    Parameters
    ----------
    input_path:
        Raw EMPIRE scenario CSV.
    output_path:
        Destination path for the disaggregated CSV.
    shares:
        ``{country_col: {nuts2_code: fraction, ...}, ...}``
        Fractions must sum to 1 (or all be 0 for zero-capacity countries).
    always_drop:
        Country column codes that must be removed even when ``shares`` has
        no entry for them (the parent country node no longer exists in the
        model).
    """
    print(f"\nProcessing: {input_path.name}")
    df = pd.read_csv(input_path)

    # Map each NUTS2 prefix to the column name actually used in this CSV
    # (e.g. prefix 'UK' is labelled 'GB' in the scenario CSVs).
    prefixes = set(shares.keys()) | set(always_drop)
    csvcol_to_prefix = {PREFIX_TO_CSV_COLUMN.get(p, p): p for p in prefixes}

    # Country columns present in this CSV that must be expanded/dropped.
    cols_to_replace = {c for c in csvcol_to_prefix if c in df.columns}

    # Build the new column order by expanding country columns in place.
    new_col_order: list[str] = []
    for col in df.columns:
        if col in cols_to_replace:
            prefix = csvcol_to_prefix[col]
            nuts2_cols = list(shares.get(prefix, {}).keys())
            if nuts2_cols:
                print(f"  Replacing '{col}' with {len(nuts2_cols)} NUTS2 columns: "
                      f"{nuts2_cols[0]} ... {nuts2_cols[-1]}")
            else:
                print(f"  Dropping '{col}' (no qualifying NUTS2 regions)")
            new_col_order.extend(nuts2_cols)
        else:
            new_col_order.append(col)

    # Compute new NUTS2 columns and drop the original country columns.
    for csv_col in cols_to_replace:
        prefix = csvcol_to_prefix[csv_col]
        region_shares = shares.get(prefix, {})
        country_series = df[csv_col]
        for region, fraction in region_shares.items():
            df[region] = country_series * fraction
        df.drop(columns=[csv_col], inplace=True)

    # Apply the derived column order.
    df = df[new_col_order]

    # Sanity check on row count (2015-2019 = 43 800 h + 2 leap days = 43 824).
    n_rows = len(df)
    expected = 5 * 365 * 24
    if abs(n_rows - expected) > 48:
        print(f"  WARNING: Unexpected row count {n_rows} (expected ~{expected}).")
    else:
        print(f"  Rows: {n_rows}  Columns: {len(df.columns)}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)
    print(f"  Saved -> {output_path}")


def _cf_multipliers(
    regions: list[str],
    cf_targets: dict[str, float],
    national_cf: float,
    mode: str,
    weights: dict[str, float] | None,
) -> tuple[dict[str, float], int]:
    """Return ``({region: multiplier}, n_regions_with_data)`` for one country.

    In ``absolute`` mode the multiplier is ``regional_CF / national_CF``, so each
    region's annual-mean CF becomes exactly its ``cf_targets`` value and the
    national level follows the source data.

    In ``relative`` mode the multipliers are divided by the capacity-weighted
    mean of ``cf_targets`` across the regions that have data, so that weighted
    mean becomes 1.0 and the country's national CF — hence its annual energy —
    is preserved.  Only the spatial pattern of the source data is used.

    Regions absent from ``cf_targets`` get a multiplier of 1.0 (plain broadcast).
    Because the relative normaliser is computed over the data-regions' weights
    renormalised among themselves, those broadcast regions sit exactly at the
    national CF and do not disturb the conservation.
    """
    with_data = [r for r in regions if r in cf_targets]
    if not with_data or national_cf <= 0:
        return {r: 1.0 for r in regions}, 0

    if mode == 'absolute':
        divisor = national_cf
    else:
        w = weights or {}
        # Renormalise the weights within the data-regions subset.
        sub_total = sum(w.get(r, 0.0) for r in with_data)
        if sub_total > 0:
            divisor = sum(
                (w.get(r, 0.0) / sub_total) * cf_targets[r] for r in with_data
            )
        else:
            divisor = sum(cf_targets[r] for r in with_data) / len(with_data)
        if divisor <= 0:
            return {r: 1.0 for r in regions}, 0

    return (
        {r: (cf_targets[r] / divisor if r in cf_targets else 1.0) for r in regions},
        len(with_data),
    )


def disaggregate_csv_cf_rescale(
    input_path: Path,
    output_path: Path,
    broadcast_shares: dict[str, dict[str, float]],
    cf_targets: dict[str, float],
    *,
    rescale_prefixes: tuple[str, ...] = ('ES',),
    mode: str = 'absolute',
    gen_file: Path | None = None,
    tech: str | None = None,
    clip_max: float = 1.0,
    always_drop: tuple[str, ...] = tuple(NUTS2_COUNTRY_PREFIXES),
) -> None:
    """Disaggregate a capacity-factor CSV with per-region CF rescaling.

    Behaves like :func:`disaggregate_csv` in broadcast mode (each NUTS region
    replaces its country column in place, preserving column order) except that
    for every country in ``rescale_prefixes`` each region's series is the
    national series scaled by a per-region multiplier, so the hourly shape is
    preserved but the annual level differs region to region.  See
    :func:`_cf_multipliers` for how ``mode`` ('absolute' vs 'relative') sets that
    multiplier.  Values are clipped to ``[0, clip_max]`` so scaling up can never
    exceed the physical maximum.

    ``gen_file``/``tech`` supply the RefInitialCap capacity weights used by
    'relative' mode; without them it falls back to an unweighted mean.

    Regions absent from ``cf_targets`` — and every country outside
    ``rescale_prefixes`` — keep the plain broadcast (copy) behaviour.
    """
    print(f"\nProcessing (CF rescale, {mode}): {input_path.name}")
    df = pd.read_csv(input_path)

    prefixes = set(broadcast_shares.keys()) | set(always_drop)
    csvcol_to_prefix = {PREFIX_TO_CSV_COLUMN.get(p, p): p for p in prefixes}
    cols_to_replace = {c for c in csvcol_to_prefix if c in df.columns}
    rescale_set = set(rescale_prefixes)

    # Per (csv column) multiplier for each of its NUTS regions.
    multipliers: dict[str, dict[str, float]] = {}
    for csv_col in cols_to_replace:
        prefix = csvcol_to_prefix[csv_col]
        regions = list(broadcast_shares.get(prefix, {}).keys())
        if prefix in rescale_set and cf_targets:
            nat_mean = float(pd.to_numeric(df[csv_col], errors='coerce').mean())
            weights = (
                load_capacity_weights(gen_file, tech, regions)
                if mode == 'relative' and gen_file and tech else None
            )
            reg_mult, n_data = _cf_multipliers(
                regions, cf_targets, nat_mean, mode, weights)
            multipliers[prefix] = reg_mult
            if n_data:
                spread = max(reg_mult.values()) / min(reg_mult.values())
                print(f"  {csv_col}: rescaling {len(regions)} regions "
                      f"(national mean={nat_mean:.4f}); {n_data} with data, "
                      f"{len(regions) - n_data} broadcast; "
                      f"multiplier {min(reg_mult.values()):.3f}-"
                      f"{max(reg_mult.values()):.3f} ({spread:.2f}x)")
            else:
                print(f"  {csv_col}: no CF data for any region -> broadcast.")
        else:
            multipliers[prefix] = {r: 1.0 for r in regions}

    # Build new column order by expanding country columns in place.
    new_col_order: list[str] = []
    for col in df.columns:
        if col in cols_to_replace:
            new_col_order.extend(multipliers.get(csvcol_to_prefix[col], {}).keys())
        else:
            new_col_order.append(col)

    # Compute region columns and drop the originals.
    for csv_col in cols_to_replace:
        prefix = csvcol_to_prefix[csv_col]
        country_series = df[csv_col]
        do_clip = (prefix in rescale_set)
        for region, mult in multipliers.get(prefix, {}).items():
            series = country_series * mult
            if do_clip:
                series = series.clip(lower=0.0, upper=clip_max)
            df[region] = series
        df.drop(columns=[csv_col], inplace=True)

    df = df[new_col_order]
    print(f"  Rows: {len(df)}  Columns: {len(df.columns)}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)
    print(f"  Saved -> {output_path}")


def disaggregate_csv_atlite(
    input_path: Path,
    output_path: Path,
    atlite_path: Path,
    *,
    prefixes: tuple[str, ...] = tuple(NUTS2_COUNTRY_PREFIXES),
    reconcile_prefix: str | None = None,
    gen_file: Path | None = None,
    tech: str | None = None,
    clip_max: float = 1.0,
) -> None:
    """Replace country columns with region-specific capacity factors taken from
    an Atlite preprocessing CSV (Scripts/preprocess_renewables_atlite.py).

    Unlike :func:`disaggregate_csv` in 'broadcast' mode — where every NUTS2
    region of a country gets a *copy* of the national capacity-factor column —
    here each region receives its **own** weather-consistent CF time series, so
    intra-country spatial diversity (coastal-vs-inland wind, etc.) is preserved.

    Rows are aligned on the ``time`` column, so the Atlite file must cover every
    timestamp present in the scenario CSV.  Column order is preserved: each
    country column is expanded in place into its region columns, every other
    column is left untouched.

    Reconciliation
    --------------
    When ``reconcile_prefix`` (e.g. ``'ES'``) is set and ``gen_file``/``tech``
    are given, that country's Atlite series are rescaled by a **single**
    multiplicative factor so their capacity-weighted national mean reproduces the
    EMPIRE national CF (the country's original column mean in this file).  The
    factor preserves every region's shape and their relative differences; only
    the overall level is corrected.  Weights come from
    :func:`load_capacity_weights` (RefInitialCap per NUTS region).  Values are
    clipped to ``[0, clip_max]``.  Other prefixes are always inserted verbatim
    (assumed already at the correct level, e.g. reconciled at Atlite-build time).
    """
    print(f"\nProcessing (Atlite CFs): {input_path.name}")
    df = pd.read_csv(input_path)
    cf = pd.read_csv(atlite_path)

    if 'time' not in df.columns or 'time' not in cf.columns:
        raise ValueError(
            f"Both {input_path.name} and {atlite_path.name} must have a 'time' column "
            "for Atlite CF substitution.")

    cf_idx = cf.set_index('time')

    # Every scenario timestamp must be present in the Atlite file.
    missing_times = set(df['time']) - set(cf_idx.index)
    if missing_times:
        raise ValueError(
            f"Atlite file {atlite_path.name} is missing {len(missing_times)} timestamps "
            f"present in {input_path.name} (e.g. {sorted(missing_times)[:3]}). "
            "Generate the Atlite CFs for the same weather years as the scenario data.")
    # Align the Atlite rows to the scenario rows (positional after reindex).
    cf_aligned = cf_idx.reindex(df['time']).reset_index(drop=True)

    # Map each NUTS2 prefix to the column name used in this CSV (UK -> GB, ...).
    csvcol_to_prefix = {PREFIX_TO_CSV_COLUMN.get(p, p): p for p in prefixes}
    cols_to_replace = {c for c in csvcol_to_prefix if c in df.columns}

    # Assemble the output by expanding each country column in place into its
    # region columns; build the frame in one concat to avoid fragmentation.
    pieces: list[pd.DataFrame] = []
    for col in df.columns:
        if col in cols_to_replace:
            prefix = csvcol_to_prefix[col]
            region_cols = [c for c in cf_aligned.columns if str(c).startswith(prefix)]
            if not region_cols:
                raise ValueError(
                    f"No regions for prefix '{prefix}' found in {atlite_path.name}. "
                    f"Include {prefix} in the Atlite run, or drop --atlite-dir to broadcast.")
            block = cf_aligned[region_cols].reset_index(drop=True)

            if reconcile_prefix and prefix == reconcile_prefix and gen_file and tech:
                target = float(pd.to_numeric(df[col], errors='coerce').mean())
                weights = load_capacity_weights(gen_file, tech, region_cols)
                atl_nat = sum(weights.get(r, 0.0) * float(block[r].mean())
                              for r in region_cols)
                if target > 0 and atl_nat > 0:
                    factor = target / atl_nat
                    block = (block * factor).clip(lower=0.0, upper=clip_max)
                    print(f"  reconcile '{prefix}' [{tech}]: national target="
                          f"{target:.4f}, atlite cap-weighted={atl_nat:.4f} "
                          f"-> factor={factor:.4f} (clip {clip_max})")
                else:
                    print(f"  reconcile '{prefix}' skipped "
                          f"(target={target:.4f}, atlite={atl_nat:.4f}).")

            print(f"  Replacing '{col}' with {len(region_cols)} Atlite CF columns: "
                  f"{region_cols[0]} ... {region_cols[-1]}")
            pieces.append(block)
        else:
            pieces.append(df[[col]].reset_index(drop=True))

    df = pd.concat(pieces, axis=1)

    print(f"  Rows: {len(df)}  Columns: {len(df.columns)}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)
    print(f"  Saved -> {output_path}")


def copy_unchanged(input_path: Path, output_path: Path) -> None:
    """Copy a scenario file to the output directory without modification."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(input_path, output_path)
    print(f"  Copied (unchanged): {output_path.name}")


def resolve_atlite_file(atlite_dir: Path, scenario_name: str) -> Path | None:
    """Find the Atlite CF file for a scenario CSV inside ``atlite_dir``.

    Accepts the canonical name (``solar.csv``) and any single stem-suffixed
    variant (``solar_ES_NUTS3.csv``), so country/level-tagged Atlite exports can
    be used without renaming.  Returns the matching path, or ``None`` when no
    file (or an ambiguous set) is found — callers then fall back to broadcast.
    """
    exact = atlite_dir / scenario_name
    if exact.exists():
        return exact
    stem = Path(scenario_name).stem            # 'solar', 'windonshore'
    matches = sorted(atlite_dir.glob(f"{stem}_*.csv"))
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        print(f"  Note: multiple Atlite candidates for '{scenario_name}' in "
              f"{atlite_dir} ({', '.join(m.name for m in matches)}); "
              "skipping Atlite substitution.")
    return None


# =============================================================================
# Main
# =============================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description='Disaggregate EMPIRE scenario CSVs to NUTS2 for selected countries.')
    p.add_argument('--iamc',
                   default='data/IAMC/NorthSea/Denmark_Germany.csv',
                   help='Path to IAMC CSV file.')
    p.add_argument('--scenario',
                   default='Transmission v1.0',
                   help='Scenario name in the IAMC file.')
    p.add_argument('--sets-file',
                   default='Sets_NUTS2_DK_DE.xlsx',
                   help='Generated Sets.xlsx; used to derive which NUTS2 nodes '
                        'receive offshore-wind columns (default: Sets_NUTS2_DK_DE.xlsx).')
    p.add_argument('--input-dir',
                   default='data/EMPIRE/v_52/ScenarioData',
                   help='Directory containing raw EMPIRE scenario CSV files.')
    p.add_argument('--output-dir',
                   default='data/EMPIRE/v_52/ScenarioData_NUTS2',
                   help='Directory for disaggregated output CSV files.')
    p.add_argument('--atlite-dir',
                   default=None,
                   help='Directory with region-specific CF CSVs (solar.csv, windonshore.csv) '
                        'from preprocess_renewables_atlite.py. When given, those files use the '
                        'per-region Atlite series instead of broadcasting the national CF. '
                        'Files not present in this directory fall back to broadcast.')
    p.add_argument('--no-es-reconcile', action='store_true',
                   help='Do NOT rescale the Spain Atlite CFs to the EMPIRE national level '
                        '(by default ES solar/windonshore Atlite series are reconciled with a '
                        'single capacity-weighted factor so annual energy matches the original).')
    return p.parse_args()


def main() -> None:
    args = parse_args()

    iamc_file  = args.iamc
    scenario   = args.scenario
    input_dir  = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    atlite_dir = Path(args.atlite_dir) if args.atlite_dir else None

    print("=" * 60)
    print("EMPIRE Scenario Disaggregation -- selected countries -> NUTS2")
    print("=" * 60)
    print(f"  IAMC file   : {iamc_file}")
    print(f"  Scenario    : {scenario}")
    print(f"  Sets file   : {args.sets_file}")
    print(f"  Input dir   : {input_dir}")
    print(f"  Output dir  : {output_dir}")
    print(f"  Atlite dir  : {atlite_dir if atlite_dir else '(none -- broadcast solar/wind)'}")
    print(f"  NUTS2 countries: {', '.join(NUTS2_COUNTRY_PREFIXES)}")
    print()

    print("Loading NUTS2 region lists from Sets file ...")
    nuts2_regions_by_country = load_nuts2_regions_from_sets(
        args.sets_file,
        NUTS2_COUNTRY_PREFIXES,
    )
    for country_col in NUTS2_COUNTRY_PREFIXES:
        regions = nuts2_regions_by_country.get(country_col, [])
        print(f"  {country_col}: {len(regions)} NUTS2 regions from Sets")

    # ------------------------------------------------------------------
    # 1. Pre-load shares (done once, reused across files).
    # ------------------------------------------------------------------
    print("Loading NUTS2 load shares from IAMC file ...")
    iamc_shares = load_iamc_shares(
        iamc_file,
        scenario,
        nuts2_regions_by_country,
    )

    print("\nLoading NUTS2 broadcast shares (capacity factors) ...")
    broadcast_shares = load_broadcast_shares(
        iamc_file,
        scenario,
        nuts2_regions_by_country,
    )

    # Per-region CF targets to rescale solar/windonshore instead of broadcasting
    # the national CF.  Empty unless a configured country is modeled at NUTS
    # level and its source files are present.
    cf_countries = [
        c for c in RENEWABLE_COUNTRIES if nuts2_regions_by_country.get(c)
    ]
    cf_targets_by_file: dict[str, dict[str, float]] = {}
    if cf_countries and RENEWABLE_DIR:
        print(f"\nLoading capacity-factor targets from {RENEWABLE_DIR} "
              f"for {', '.join(cf_countries)} (mode={RENEWABLE_MODE}) ...")
        cf_targets_by_file = load_cf_targets(
            nuts2_regions_by_country, RENEWABLE_DIR, cf_countries)

    print("\nLoading NUTS2 hydro ROR shares from GeneratorsOfNode ...")
    hydro_ror_shares = load_hydro_ror_shares_from_sets(args.sets_file)

    print("\nLoading NUTS2 hydro capacity shares from IAMC file (Sets-filtered) ...")
    hydro_shares = load_hydro_capacity_shares(iamc_file, scenario, args.sets_file)

    print("\nLoading NUTS2 wind offshore shares from GeneratorsOfNode ...")
    wind_offshore_shares = load_wind_offshore_shares_from_sets(args.sets_file)

    # ------------------------------------------------------------------
    # 2. Process all CSVs in the input directory.
    # ------------------------------------------------------------------
    all_input_files = sorted(input_dir.glob('*.csv'))

    print("\n" + "-" * 60)
    for csv_path in all_input_files:
        output_path = output_dir / csv_path.name
        method = DISAGGREGATE_FILES.get(csv_path.name)

        if method == 'iamc_share':
            disaggregate_csv(csv_path, output_path, iamc_shares)
        elif method == 'broadcast':
            atlite_file = resolve_atlite_file(atlite_dir, csv_path.name) if atlite_dir else None
            file_cf = cf_targets_by_file.get(csv_path.name)
            if atlite_file and atlite_file.exists():
                reconcile_prefix = (
                    'ES' if nuts2_regions_by_country.get('ES')
                    and not args.no_es_reconcile else None)
                disaggregate_csv_atlite(
                    csv_path, output_path, atlite_file,
                    reconcile_prefix=reconcile_prefix,
                    gen_file=output_dir.parent / 'Generator.xlsx',
                    tech=ATLITE_RECONCILE_TECH.get(csv_path.name))
            elif file_cf:
                if atlite_dir:
                    print(f"\n  Note: {csv_path.name} not found in {atlite_dir}; "
                          "using CF rescaling.")
                disaggregate_csv_cf_rescale(
                    csv_path, output_path, broadcast_shares, file_cf,
                    rescale_prefixes=tuple(cf_countries),
                    mode=RENEWABLE_MODE,
                    gen_file=output_dir.parent / 'Generator.xlsx',
                    tech=ATLITE_RECONCILE_TECH.get(csv_path.name))
            else:
                if atlite_dir:
                    print(f"\n  Note: {csv_path.name} not found in {atlite_dir}; "
                          "falling back to broadcast.")
                disaggregate_csv(csv_path, output_path, broadcast_shares)
        elif method == 'hydro_ror_sets':
            disaggregate_csv(csv_path, output_path, hydro_ror_shares)
        elif method == 'hydro_capacity':
            disaggregate_csv(csv_path, output_path, hydro_shares)
        elif method == 'wind_offshore_broadcast':
            disaggregate_csv(csv_path, output_path, wind_offshore_shares)
        else:
            copy_unchanged(csv_path, output_path)

    print("\nDone.")


if __name__ == '__main__':
    main()
