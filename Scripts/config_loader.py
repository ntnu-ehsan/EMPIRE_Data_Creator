"""
Load configuration from config.yaml in the project root.

All generate_*.py scripts import from this module so that the country list,
NUTS2 region codes, EMPIRE node names, and NUTS2 coordinates come from one
place.

NUTS2 region codes are derived automatically from the IAMC input file: any
standalone region in the file whose code is exactly 4 characters and starts
with a configured country prefix is included.

Coordinates are computed as polygon centroids from the NUTS shapefile at
data/shp/NUTS_RG_20M_2024_3035.shp and reprojected to WGS84 (lat/lon).
"""

import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    print("ERROR: PyYAML not installed. Run: pip install pyyaml")
    sys.exit(1)

_CONFIG_PATH = Path(__file__).parent.parent / 'config.yaml'
_SHP_DIR = Path(__file__).parent.parent / 'data' / 'shp'

# No single NUTS vintage covers every region the model needs:
#   * the 2024 classification has the revised NL codes (NL35, NL36) but dropped
#     all UK regions after Brexit;
#   * the 2021 classification still has the 41 UK NUTS2 regions but lacks the
#     newer NL codes.
# So centroids are resolved from multiple files in order, first match wins.
# Keep the 2024 file first so DK/NL/NO centroids match earlier outputs, and use
# the 2021 file only to fill the UK regions the 2024 file is missing.
_SHAPEFILE_PATHS = [
    _SHP_DIR / 'NUTS_RG_20M_2024_3035.shp',
    _SHP_DIR / 'NUTS_RG_60M_2021_3035.geojson',
]


def _derive_regions_from_iamc(iamc_file: str, prefixes: list) -> dict:
    """Return {prefix: [nuts2_code, ...]} by scanning the IAMC region column."""
    try:
        import pyam  # noqa: F401  (ensures a clear error if pyam is missing)
        from iamc_utils import load_iamc
    except ImportError:
        print("ERROR: pyam-iamc not installed. Run: pip install pyam-iamc")
        sys.exit(1)

    df = load_iamc(iamc_file)
    all_regions = set(df.region)

    regions_by_country = {}
    for prefix in prefixes:
        # NUTS2 codes are 4 chars (e.g. ES41); NUTS3 codes are 5 chars
        # (e.g. ES411). Accept both so a country can be modelled at either
        # spatial resolution depending on what the IAMC file provides.
        nuts2 = sorted(
            r for r in all_regions
            if '>' not in r and len(r) in (4, 5) and r.startswith(prefix)
        )
        if not nuts2:
            print(f"  WARNING [config]: no NUTS2 regions found for '{prefix}' in {iamc_file}")
        regions_by_country[prefix] = nuts2

    return regions_by_country


def _derive_coords_from_shapefile(all_nuts2: list) -> dict:
    """Return {nuts2_code: (lat, lon)} from shapefile polygon centroids."""
    try:
        import geopandas as gpd
    except ImportError:
        print("ERROR: geopandas not installed. Run: pip install geopandas")
        sys.exit(1)

    existing = [p for p in _SHAPEFILE_PATHS if p.exists()]
    if not existing:
        raise FileNotFoundError(
            f"No NUTS shapefile found. Looked for: {[str(p) for p in _SHAPEFILE_PATHS]}\n"
            "Download from: https://gisco-services.ec.europa.eu/distribution/v2/nuts/shp/"
        )

    nuts2_set = set(all_nuts2)
    coords = {}

    # Read each file in order; the first file to provide a code wins.
    for shp_path in existing:
        remaining = nuts2_set - set(coords)
        if not remaining:
            break

        gdf = gpd.read_file(shp_path)
        if 'NUTS_ID' not in gdf.columns:
            raise KeyError(
                f"Expected 'NUTS_ID' column not found in {shp_path.name}. "
                f"Available columns: {list(gdf.columns)}"
            )

        gdf_nuts2 = gdf[gdf['NUTS_ID'].isin(remaining)].copy()
        if gdf_nuts2.empty:
            continue

        # Compute centroids in native CRS, then reproject to WGS84
        centroids = gdf_nuts2.geometry.centroid
        gdf_centroids = gdf_nuts2[['NUTS_ID']].copy()
        gdf_centroids = gdf_centroids.set_geometry(centroids)
        gdf_centroids = gdf_centroids.set_crs(gdf.crs).to_crs(epsg=4326)

        added = 0
        for _, row in gdf_centroids.iterrows():
            code = row['NUTS_ID']
            if code not in coords:
                coords[code] = (round(row.geometry.y, 6), round(row.geometry.x, 6))
                added += 1
        if added:
            print(f"  [config] {added} centroids resolved from {shp_path.name}")

    missing = nuts2_set - set(coords)
    if missing:
        print(f"  WARNING [config]: no shapefile entry for: {sorted(missing)}")

    return coords


def load_nuts2_polygons(all_nuts2: list) -> dict:
    """Return {nuts2_code: shapely geometry in EPSG:3035} for the given codes.

    Uses the same multi-file fallback as centroid resolution (first match wins),
    so UK regions come from the 2021 file and NL/DK/NO from the 2024 file.
    Geometries are returned in EPSG:3035 so distances are metres.
    """
    try:
        import geopandas as gpd
    except ImportError:
        print("ERROR: geopandas not installed. Run: pip install geopandas")
        sys.exit(1)

    existing = [p for p in _SHAPEFILE_PATHS if p.exists()]
    if not existing:
        raise FileNotFoundError(
            f"No NUTS shapefile found. Looked for: {[str(p) for p in _SHAPEFILE_PATHS]}\n"
            "Download from: https://gisco-services.ec.europa.eu/distribution/v2/nuts/shp/"
        )

    wanted = set(all_nuts2)
    polys = {}
    for shp_path in existing:
        remaining = wanted - set(polys)
        if not remaining:
            break
        gdf = gpd.read_file(shp_path).to_crs(epsg=3035)
        if 'NUTS_ID' not in gdf.columns:
            raise KeyError(
                f"Expected 'NUTS_ID' column not found in {shp_path.name}. "
                f"Available columns: {list(gdf.columns)}"
            )
        sub = gdf[gdf['NUTS_ID'].isin(remaining)]
        for _, row in sub.iterrows():
            polys.setdefault(row['NUTS_ID'], row.geometry)

    missing = wanted - set(polys)
    if missing:
        print(f"  WARNING [config]: no shapefile polygon for: {sorted(missing)}")

    return polys


def nearest_nuts2_by_distance(lon: float, lat: float,
                              candidate_codes: list, polygons: dict):
    """Return (best_code, distance_m) for the candidate NUTS2 region whose
    polygon boundary is nearest to the WGS84 point (lon, lat).

    Because offshore points lie at sea, the point-to-polygon distance equals the
    distance to that region's coastline. Distance is in metres (EPSG:3035).
    Returns (None, inf) if no candidate has a polygon.
    """
    try:
        import geopandas as gpd
        from shapely.geometry import Point
    except ImportError:
        print("ERROR: geopandas/shapely not installed. Run: pip install geopandas")
        sys.exit(1)

    pt = gpd.GeoSeries([Point(lon, lat)], crs='EPSG:4326').to_crs(epsg=3035).iloc[0]
    best, best_d = None, float('inf')
    for code in candidate_codes:
        geom = polygons.get(code)
        if geom is None:
            continue
        d = geom.distance(pt)
        if d < best_d:
            best, best_d = code, d
    return best, best_d


def load_nuts2_config(config_path=None):
    """Load configuration from config.yaml.

    NUTS2 region lists and coordinates are derived automatically from the
    IAMC file and the NUTS shapefile — they no longer need to be listed
    manually in config.yaml.

    Returns a dict with keys:
        IAMC_FILE                    str
        SCENARIO                     str
        NUTS2_COUNTRY_PREFIXES       list[str]
        NUTS2_COUNTRY_TO_EMPIRE_NODE dict[str, str]
        NUTS2_REGIONS_BY_COUNTRY     dict[str, list[str]]
        NUTS2_COORDS                 dict[str, tuple[float, float]]
        ALL_NUTS2                    list[str]  (sorted)
        ALL_NUTS2_SET                set[str]
        REPLACED_COUNTRY_NODES       list[str]
    """
    path = Path(config_path) if config_path else _CONFIG_PATH
    if not path.exists():
        raise FileNotFoundError(
            f"config.yaml not found at {path}. "
            "Make sure it exists in the project root."
        )

    with open(path, encoding='utf-8') as f:
        cfg = yaml.safe_load(f)

    if 'nuts2_countries' not in cfg:
        raise KeyError("config.yaml must contain a 'nuts2_countries' section.")

    # PyYAML (YAML 1.1) parses bare keys like NO, YES, TRUE as booleans.
    # Normalise all country keys to strings before any further use.
    countries = {str(k): v for k, v in cfg['nuts2_countries'].items()}
    prefixes = list(countries.keys())
    empire_nodes = {k: v['empire_node'] for k, v in countries.items()}

    iamc_file = cfg.get('iamc_file', '')
    # Allow CLI arguments to override the IAMC file for NUTS2 region detection.
    # Two calling conventions are supported:
    #   generate_scenarios.py   uses: --iamc <path>  (named flag)
    #   generate_sets_nuts2.py  uses: <path> as first positional arg
    #   generate_generation_nuts2.py, generate_node.py, generate_storage.py
    #       pass <path> as a later positional arg
    # Both are detected here so that running with --iamc on main.py correctly
    # propagates to all sub-scripts regardless of how they accept it.
    argv = sys.argv
    _override = None
    for i, arg in enumerate(argv):
        if arg == '--iamc' and i + 1 < len(argv):
            _override = argv[i + 1]
            break
    if _override is None:
        # Fall back: any positional .xlsx argument whose path contains 'iamc'
        # (case-insensitive) is treated as the IAMC file override.
        for arg in argv[1:]:
            if not arg.startswith('-') and arg.lower().endswith('.xlsx') and 'iamc' in arg.lower():
                _override = arg
                break
    if _override is not None:
        iamc_file = _override
    if not iamc_file:
        raise ValueError(
            "iamc_file must be set in config.yaml.\n"
            "Example:  iamc_file: data/IAMC/NorthSea/Denmark_Germany_Netherlands.xlsx"
        )

    print(f"  [config] Scanning IAMC file for NUTS2 regions: {iamc_file}")
    regions_by_country = _derive_regions_from_iamc(iamc_file, prefixes)

    all_nuts2 = sorted(r for regions in regions_by_country.values() for r in regions)
    total = sum(len(v) for v in regions_by_country.values())
    print(f"  [config] Found {total} NUTS2 regions "
          f"({', '.join(f'{p}: {len(regions_by_country[p])}' for p in prefixes)})")

    print(f"  [config] Computing centroids from shapefile...")
    coords = _derive_coords_from_shapefile(all_nuts2)
    print(f"  [config] Coordinates resolved for {len(coords)}/{total} regions")

    return {
        'IAMC_FILE': iamc_file,
        'SCENARIO': cfg.get('scenario', ''),
        'NUTS2_COUNTRY_PREFIXES': prefixes,
        'NUTS2_COUNTRY_TO_EMPIRE_NODE': empire_nodes,
        'NUTS2_REGIONS_BY_COUNTRY': regions_by_country,
        'NUTS2_COORDS': coords,
        'ALL_NUTS2': all_nuts2,
        'ALL_NUTS2_SET': set(all_nuts2),
        'REPLACED_COUNTRY_NODES': [empire_nodes[p] for p in prefixes],
    }


def load_load_flow_config(config_path=None):
    """Return the load-flow settings from config.yaml.

    Returns a dict with keys:
        enabled            bool   — whether to generate load-flow sheets
        reactance_per_km   float  — Ohm/km for computed reactances
        block_capacity_mw  float  — constant candidate block capacity (MW)
        nominal_voltage_kv float  — nominal grid voltage (kV)

    Missing keys fall back to the documented defaults, and an absent
    ``load_flow`` section disables the feature.
    """
    path = Path(config_path) if config_path else _CONFIG_PATH
    if not path.exists():
        raise FileNotFoundError(
            f"config.yaml not found at {path}. "
            "Make sure it exists in the project root."
        )
    with open(path, encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}

    lf = cfg.get('load_flow') or {}
    return {
        'enabled': bool(lf.get('enabled', False)),
        'reactance_per_km': float(lf.get('reactance_per_km', 0.246)),
        'block_capacity_mw': float(lf.get('block_capacity_mw', 1500)),
        'nominal_voltage_kv': float(lf.get('nominal_voltage_kv', 400)),
    }


def load_expansion_config(config_path=None):
    """Return the planned-transmission-expansion settings from config.yaml.

    Returns a dict with keys:
        enabled  bool — whether to fold Spain's planned lines into the pipeline
                        (adds new corridors to Sets and injects their capacity
                        into Transmission InitialCapacity).

    An absent ``expansion_plans`` section disables the feature.
    """
    path = Path(config_path) if config_path else _CONFIG_PATH
    with open(path, encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}
    ep = cfg.get('expansion_plans') or {}
    return {'enabled': bool(ep.get('enabled', False))}


def load_renewable_cf_config(config_path=None):
    """Return the regional capacity-factor settings from config.yaml.

    Returns a dict with keys:
        dir        Path  — folder holding the CF source files (Wind.csv or
                           onshore.csv, and Solar.xlsx)
        mode       str   — 'relative' or 'absolute' (see below)
        countries  list  — country prefixes the rescaling applies to

    The two modes differ in what they do to a country's *national* capacity
    factor:

    ``absolute``
        Each region's series is scaled to the capacity factor read from the
        source file (``multiplier = regional_CF / national_CF``).  The national
        CF is therefore replaced by the source data's own level.  This is the
        original Spain behaviour, which is safe there because the Spanish atlas
        mean matches the EMPIRE national CF to within 1%.

    ``relative``
        Only the *spatial pattern* of the source data is used: multipliers are
        normalised so the capacity-weighted mean across a country's regions is
        1.0, leaving the EMPIRE national CF — and hence national annual energy —
        untouched.  Use this whenever the source atlas and EMPIRE disagree on
        level, as they do for the NESP countries (Denmark's atlas CF is 1.66x
        EMPIRE's).

    An absent ``renewable_cf`` section reproduces the historical Spain defaults
    (data/Renewable, absolute, ES only) so existing runs are unaffected.
    """
    path = Path(config_path) if config_path else _CONFIG_PATH
    with open(path, encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}

    rc = cfg.get('renewable_cf')
    if rc is None:
        return {
            'dir': Path(__file__).parent.parent / 'data' / 'Renewable',
            'mode': 'absolute',
            'countries': ['ES'],
        }

    raw_dir = str(rc.get('dir', '') or '')
    cf_dir = Path(raw_dir)
    if raw_dir and not cf_dir.is_absolute():
        cf_dir = Path(__file__).parent.parent / cf_dir

    mode = str(rc.get('mode', 'relative')).strip().lower()
    if mode not in ('relative', 'absolute'):
        raise ValueError(
            f"renewable_cf.mode must be 'relative' or 'absolute', got '{mode}'.")

    # PyYAML turns a bare NO into False (YAML 1.1); normalise back to strings.
    countries = [
        'NO' if c is False else str(c)
        for c in (rc.get('countries') or [])
    ]

    return {
        'dir': cf_dir if raw_dir else None,
        'mode': mode,
        'countries': countries,
    }


def load_nuts0_polygons(country_codes: list) -> dict:
    """Return {cc: shapely geometry in EPSG:3035} for NUTS0 (2-letter) codes.

    Uses the same multi-file fallback as the NUTS2 polygon loader. NUTS0
    country polygons are used to test land adjacency between a fine-resolution
    region and a neighbouring country modelled as a single node.
    """
    try:
        import geopandas as gpd
    except ImportError:
        print("ERROR: geopandas not installed. Run: pip install geopandas")
        sys.exit(1)

    existing = [p for p in _SHAPEFILE_PATHS if p.exists()]
    if not existing:
        raise FileNotFoundError(
            f"No NUTS shapefile found. Looked for: {[str(p) for p in _SHAPEFILE_PATHS]}\n"
            "Download from: https://gisco-services.ec.europa.eu/distribution/v2/nuts/shp/"
        )

    wanted = set(country_codes)
    polys = {}
    for shp_path in existing:
        remaining = wanted - set(polys)
        if not remaining:
            break
        gdf = gpd.read_file(shp_path).to_crs(epsg=3035)
        if 'NUTS_ID' not in gdf.columns:
            raise KeyError(
                f"Expected 'NUTS_ID' column not found in {shp_path.name}. "
                f"Available columns: {list(gdf.columns)}"
            )
        sub = gdf[gdf['NUTS_ID'].isin(remaining)]
        for _, row in sub.iterrows():
            polys.setdefault(row['NUTS_ID'], row.geometry)

    missing = wanted - set(polys)
    if missing:
        print(f"  WARNING [config]: no NUTS0 polygon for: {sorted(missing)}")
    return polys
