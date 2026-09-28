"""
Generate Sets.xlsx for EMPIRE with NUTS2 resolution for selected countries.

All European countries are modeled as single nodes EXCEPT countries configured
in NUTS2_COUNTRY_PREFIXES, which are modeled at NUTS2 resolution.

Usage:
    python generate_sets_nuts2.py <iamc_file> <scenario> <output_file> [--reference <existing_sets.xlsx>]

Example:
    python generate_sets_nuts2.py data/IAMC/NorthSea/Denmark_Germany.csv "Transmission v1.0" Sets_NUTS2_DK_DE.xlsx
"""

import sys
import argparse
import math
from pathlib import Path
import pandas as pd
import numpy as np

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

try:
    import pyam
except ImportError:
    print("ERROR: pyam-iamc not installed. Run: pip install pyam-iamc")
    sys.exit(1)


# =============================================================================
# NUTS2 configuration — loaded from config.yaml in the project root.
# Edit config.yaml to change which countries are represented at NUTS2 level.
# =============================================================================
from config_loader import load_nuts2_config as _load_nuts2_config
from config_loader import load_load_flow_config as _load_load_flow_config
_cfg = _load_nuts2_config()
NUTS2_COUNTRY_PREFIXES = _cfg['NUTS2_COUNTRY_PREFIXES']
NUTS2_COUNTRY_TO_EMPIRE_NODE = _cfg['NUTS2_COUNTRY_TO_EMPIRE_NODE']
NUTS2_REGIONS_BY_COUNTRY = _cfg['NUTS2_REGIONS_BY_COUNTRY']
NUTS2_COORDS = _cfg['NUTS2_COORDS']
ALL_NUTS2 = _cfg['ALL_NUTS2']
REPLACED_COUNTRY_NODES = _cfg['REPLACED_COUNTRY_NODES']
_LF = _load_load_flow_config()

# =============================================================================
# Node name normalisation: the reference Sets.xlsx DirectionalLines and Coords
# sheets use no-space variants (e.g. "DoggerBank") while the Nodes sheet uses
# the canonical spaced variants ("Dogger Bank"). Apply this mapping whenever
# node names are read from those affected reference sheets so that every output
# sheet is consistent with the Nodes sheet.
# =============================================================================
NODE_NAME_FIXES = {
    'DoggerBank':       'Dogger Bank',
    'EastAnglia':       'East Anglia',
    'FirthofForth':     'Firth of Forth',
    'HelgolanderBucht': 'Helgolander Bucht',
    'HollandseeKust':   'Hollandsee Kust',
    'MorayFirth':       'Moray Firth',
    'OuterDowsing':     'Outer Dowsing',
    'SorligeNordsjoI':  'Sorlige Nordsjo I',
    'SorligeNordsjoII': 'Sorlige Nordsjo II',
    # Trinity's Nodes sheet spells these with a stray space, its
    # DirectionalLines sheet without one; both must land on the same name or
    # the hubs end up in the Node set with no corridors at all.
    'Sorlige NordsjoI':  'Sorlige Nordsjo I',
    'Sorlige NordsjoII': 'Sorlige Nordsjo II',
    'UtsiraNord':       'Utsira Nord',
    'BosniaH':          'Bosnia H',
    'CzechR':           'Czech R',
    'GreatBrit.':       'Great Brit.',
    # IE01/IE02 appear in the Nodes sheet of v1.3.0 reference files,
    # but all other sheets use 'Ireland' — normalise back to Ireland.
    'IE01':             'Ireland',
    'IE02':             'Ireland',
}


def fix_node_name(name):
    """Return the canonical node name, correcting legacy no-space variants."""
    if isinstance(name, str):
        return NODE_NAME_FIXES.get(name.strip(), name)
    return name


# =============================================================================
# Mapping from IAMC external codes to EMPIRE reference node names.
# External endpoints in the IAMC file that are outside DE/DK must be mapped
# to their corresponding EMPIRE country-level node name.
# =============================================================================
IAMC_CODE_TO_EMPIRE_NODE = {
    'AT': 'Austria',
    'DE': 'Germany',
    'BE': 'Belgium',
    'BE33': 'Belgium',
    'BG': 'Bulgaria',
    'CH': 'Switzerland',
    'CZ': 'Czech R',
    'EE': 'Estonia',
    'ES': 'Spain',
    'FI': 'Finland',
    'FR': 'France',
    'FRF1': 'France',
    'FRF3': 'France',
    'GB': 'Great Brit.',
    'GR': 'Greece',
    'HR': 'Croatia',
    'HU': 'Hungary',
    'IE': 'Ireland',
    'IT': 'Italy',
    'LT': 'Lithuania',
    'LU': 'Luxemb.',
    'LV': 'Latvia',
    'MK': 'Macedonia',
    'NL': 'Netherlands',
    'NL11': 'Netherlands',
    'NL21': 'Netherlands',
    'NL22': 'Netherlands',
    'NL42': 'Netherlands',
    'NO': 'NO1',
    'NO01': 'NO1',
    'NO02': 'NO2',
    'NO03': 'NO3',
    'NO04': 'NO4',
    'NO05': 'NO5',
    'NO09': 'NO2',
    'PL': 'Poland',
    'PT': 'Portugal',
    'RO': 'Romania',
    'RS': 'Serbia',
    'SE': 'Sweden',
    'SI': 'Slovenia',
    'SK': 'Slovakia',
    'UK': 'Great Brit.',
    'UKF3': 'Great Brit.',
}

# windoffshore.csv abbreviates the dedicated offshore hub nodes (MF, DB, SN1...)
# rather than naming them, and the abbreviations are not derivable by rule --
# 'Hornsea' -> HS and 'Norfolk' -> NF are single words, 'Moray Firth' -> MF is
# initials.  Without this map the phantom-offshore filter below cannot tell that
# a hub is profiled and strips its generators from GeneratorsOfNode, which
# aborts Pyomo model construction on genCapAvailStochRaw.
OFFSHORE_HUB_CODE_TO_NODE = {
    'MF':  'MorayFirth',
    'FF':  'FirthofForth',
    'DB':  'DoggerBank',
    'HS':  'Hornsea',
    'OD':  'OuterDowsing',
    'NF':  'Norfolk',
    'EA':  'EastAnglia',
    'BS':  'Borssele',
    'HK':  'HollandseeKust',
    'HB':  'HelgolanderBucht',
    'NS':  'Nordsoen',
    'UN':  'UtsiraNord',
    'SN1': 'SorligeNordsjoI',
    'SN2': 'SorligeNordsjoII',
}


def _despace(name):
    """Node key ignoring spaces.

    The reference datasets are internally inconsistent -- the Nodes sheet says
    'MorayFirth' while OffshoreNodes says 'Moray Firth' -- and fix_node_name
    rewrites toward the spaced form.  EMPIRE strips spaces from both node and
    generator names when it builds its index (the reported failure index was
    ('MorayFirth', 'Windoffshoregrounded', ...)), so matching on the despaced
    key is what actually reflects the model's view.
    """
    return str(name).replace(' ', '')

# =============================================================================
# IAMC Capacity|Electricity variable → EMPIRE Generator mapping
# =============================================================================
CAPACITY_TO_GENERATOR = {
    'Capacity|Electricity|Biomass|w/o CCS': 'Bio',
    'Capacity|Electricity|Biogas|w/o CCS': 'Bio',
    'Capacity|Electricity|Coal|w/o CCS': 'Coal',
    'Capacity|Electricity|Gas|CCGT|w/o CCS': 'Gas CCGT',
    'Capacity|Electricity|Gas|Steam Turbine|w/o CCS': 'Gas OCGT',
    'Capacity|Electricity|Gas|w/o CCS': 'Gas CCGT',
    # Reservoir = regulated/dammed hydro. The bare 'Hydro' key is the
    # unclassified/non-reservoir residual in the supported inputs and maps to
    # run-of-river; explicit Run of River rows, when present, are additional.
    'Capacity|Electricity|Hydro': 'Hydro run-of-the-river',
    'Capacity|Electricity|Hydro|Reservoir': 'Hydro regulated',
    'Capacity|Electricity|Hydro|Run of River': 'Hydro run-of-the-river',
    'Capacity|Electricity|Hydro|Pumped Storage': None,  # Storage, not generator
    'Capacity|Electricity|Nuclear': 'Nuclear',
    'Capacity|Electricity|Ocean': 'Wave',
    'Capacity|Electricity|Oil|CCGT|w/o CCS': 'Oil',
    'Capacity|Electricity|Oil|Steam Turbine|w/o CCS': 'Oil',
    'Capacity|Electricity|Oil|w/o CCS': 'Oil',
    'Capacity|Electricity|Other': None,
    'Capacity|Electricity|Solar|CSP': 'Solar',
    'Capacity|Electricity|Solar|PV': 'Solar',
    'Capacity|Electricity|Waste|w/o CCS': 'Waste',
    'Capacity|Electricity|Wind': None,  # legacy aggregate key, skip
    'Capacity|Electricity|Wind|Offshore': 'Wind offshore grounded',
    'Capacity|Electricity|Wind|Onshore': 'Wind onshore',
}


def resolve_iamc_endpoint(code):
    """Map an IAMC endpoint code to an EMPIRE node name.

    If the code is a DE/DK NUTS2 region, keep it as-is.
    Otherwise, map it to the EMPIRE country-level name.
    """
    if code in ALL_NUTS2:
        return code
    if code in IAMC_CODE_TO_EMPIRE_NODE:
        return IAMC_CODE_TO_EMPIRE_NODE[code]
    # Fallback: try 2-letter country code
    cc = code[:2]
    if cc in IAMC_CODE_TO_EMPIRE_NODE:
        return IAMC_CODE_TO_EMPIRE_NODE[cc]
    print(f"  ⚠ Unknown IAMC endpoint code: {code}")
    return code


def load_iamc_data(iamc_file, scenario):
    """Load and filter IAMC data.

    Returns (df_filt, nodes, links, ac_links, dc_links). The two link sets are
    kept separate because a node pair can carry parallel HVAC and HVDC
    capacity; collapsing them to one type silently discards one circuit.
    """
    print(f"[INFO] Loading IAMC data from {iamc_file}...")
    from iamc_utils import load_iamc
    df = load_iamc(iamc_file)
    df_filt = df.filter(scenario=scenario)
    regions = sorted(list(set(df_filt.region)))
    nodes = sorted([r for r in regions if '>' not in r])
    links = sorted([r for r in regions if '>' in r])

    ac_df = df_filt.filter(variable='Network|Electricity|Capacity').as_pandas()
    dc_df = df_filt.filter(variable='Network|Electricity|DC|Capacity').as_pandas()
    ac_links = set(
        ac_df.loc[ac_df['region'].str.contains('>', na=False), 'region'].unique()
    )
    dc_links = set(
        dc_df.loc[dc_df['region'].str.contains('>', na=False), 'region'].unique()
    )

    print(f"  [OK] {len(nodes)} IAMC nodes, {len(links)} IAMC links "
          f"({len(ac_links)} AC, {len(dc_links)} DC capacity links)")
    return df_filt, nodes, links, ac_links, dc_links


def add_metadata_header(df, source='-', description=''):
    """Add metadata header rows matching existing Sets.xlsx format.

    Row 0: Source: <source>
    Row 1: Description: <desc>
    Row 2: Column headers
    Row 3+: Data
    """
    cols = df.columns.tolist()
    n_cols = len(cols)

    source_row = [f'Source: {source}'] + [np.nan] * (n_cols - 1)
    desc_row = [f'Description: {description}'] + [np.nan] * (n_cols - 1)
    header_row = cols

    rows = [source_row, desc_row, header_row]
    for _, row in df.iterrows():
        rows.append(row.tolist())

    result = pd.DataFrame(rows, columns=[f'col_{i}' for i in range(n_cols)])
    return result


# =========================================================================
# Sheet builders
# =========================================================================

def make_nodes_sheet(reference_sets):
    """Sheet: Nodes

    Start from reference nodes, remove country nodes replaced by NUTS2,
    then add configured NUTS2 regions.
    """
    print("[1/11] Creating Nodes sheet...")
    ref_nodes = pd.read_excel(reference_sets, sheet_name='Nodes')
    ref_list = [fix_node_name(n) for n in ref_nodes['Node'].tolist()]

    new_nodes = list(dict.fromkeys(n for n in ref_list if n not in REPLACED_COUNTRY_NODES))
    # Add NUTS2
    new_nodes.extend(ALL_NUTS2)
    new_nodes = sorted(new_nodes)

    df = pd.DataFrame({'Node': new_nodes})
    print(f"  [OK] {len(new_nodes)} nodes (removed {', '.join(REPLACED_COUNTRY_NODES)}, "
          f"added {len(ALL_NUTS2)} NUTS2 regions)")
    return df, new_nodes


def make_offshore_nodes(reference_sets):
    """Sheet: OffshoreNodes - copy from reference as-is."""
    print("[2/11] Copying OffshoreNodes from reference...")
    df = pd.read_excel(reference_sets, sheet_name='OffshoreNodes')
    print(f"  [OK] {len(df)} offshore nodes")
    return df


def make_horizon(reference_sets):
    """Sheet: Horizon - copy from reference as-is."""
    print("[3/11] Copying Horizon from reference...")
    df = pd.read_excel(reference_sets, sheet_name='Horizon')
    print(f"  [OK] {len(df)} periods")
    return df


def make_countries():
    """Sheet: Countries - one row per NUTS-disaggregated country.

    Lists only countries that carry at least one national-level limit, i.e.
    the countries configured for NUTS2/NUTS3 disaggregation. See
    NUTS_national_limits_data_spec.md.
    """
    print("[INFO] Creating Countries sheet...")
    names = [NUTS2_COUNTRY_TO_EMPIRE_NODE[p] for p in NUTS2_COUNTRY_PREFIXES]
    df = pd.DataFrame({'Country': names})
    print(f"  [OK] {len(df)} countries: {', '.join(names)}")
    return df


def make_nodes_of_country(all_nodes):
    """Sheet: NodesOfCountry - maps each disaggregated country to its NUTS
    regions, so national-level Generator.xlsx limits know which nodes to sum
    over. See NUTS_national_limits_data_spec.md.
    """
    print("[INFO] Creating NodesOfCountry sheet...")
    node_set = set(all_nodes)
    rows = []
    for prefix in NUTS2_COUNTRY_PREFIXES:
        country = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        regions = [r for r in NUTS2_REGIONS_BY_COUNTRY.get(prefix, []) if r in node_set]
        for region in regions:
            rows.append({'Country': country, 'Node': region})
    df = pd.DataFrame(rows)
    print(f"  [OK] {len(df)} node assignments")
    return df


def copy_sheet_simple(reference_sets, sheet_name):
    """Copy a sheet as-is from reference (with column headers)."""
    return pd.read_excel(reference_sets, sheet_name=sheet_name)


def copy_sheet_raw(reference_sets, sheet_name):
    """Copy a sheet as-is from reference (raw, preserving metadata rows)."""
    return pd.read_excel(reference_sets, sheet_name=sheet_name, header=None)


def make_storage_of_nodes(all_nodes, reference_sets):
    """Sheet: StorageOfNodes - cross-product of all nodes × all storage types."""
    print("[4/11] Creating StorageOfNodes...")
    storage_df = pd.read_excel(reference_sets, sheet_name='Storage')
    storage_types = storage_df['Storage'].dropna().tolist()

    rows = []
    for node in all_nodes:
        for st in storage_types:
            rows.append({'Node': node, 'Storage': st})
    df = pd.DataFrame(rows)
    print(f"  [OK] {len(df)} rows ({len(all_nodes)} nodes × {len(storage_types)} storage types)")
    return df


def make_offshore_reconnections(reference_sets, ref_dl):
    """Re-attach offshore nodes whose onshore landing was a disaggregated country.

    In the reference each offshore node lands on one or more mainland country
    nodes. When such a country is replaced by its NUTS2 regions, the
    offshore<->country line is dropped (the country node no longer exists),
    leaving the offshore node disconnected from that shore. This restores the
    link by connecting the offshore node to the *nearest coastal NUTS2 region of
    the same country* — measured as the point-to-polygon distance from the
    offshore platform to each candidate NUTS2 region (EPSG:3035). Since the
    platform is at sea, that distance is the distance to the region's coastline.

    Returns a DataFrame of directional (both-direction) line pairs to add.
    """
    from config_loader import load_nuts2_polygons, nearest_nuts2_by_distance

    empty = pd.DataFrame(columns=['NodeFrom', 'NodeTo'])

    # Offshore nodes (canonical, spaced names) and their coordinates.
    off_df = pd.read_excel(reference_sets, sheet_name='OffshoreNodes')
    offshore_nodes = {fix_node_name(str(n)) for n in off_df.iloc[:, 0].dropna()}

    coords = pd.read_excel(reference_sets, sheet_name='Coords', header=2)
    coords.columns = ['Location', 'Latitude', 'Longitude']
    coords['Location'] = coords['Location'].map(fix_node_name)
    coord_lookup = {r['Location']: (r['Latitude'], r['Longitude'])
                    for _, r in coords.iterrows()}

    # empire node name -> country prefix, restricted to disaggregated countries.
    node_to_prefix = {NUTS2_COUNTRY_TO_EMPIRE_NODE[p]: p
                      for p in NUTS2_COUNTRY_PREFIXES}

    # Severed landings: offshore node directly linked to a replaced country node.
    landings = set()  # {(offshore_node, prefix)}
    for _, row in ref_dl.iterrows():
        a, b = row['NodeFrom'], row['NodeTo']
        for off, onshore in ((a, b), (b, a)):
            if off in offshore_nodes and onshore in node_to_prefix:
                landings.add((off, node_to_prefix[onshore]))

    if not landings:
        print("  Offshore reconnection: no severed offshore<->country landings")
        return empty

    needed = sorted({c for _, pfx in landings
                     for c in NUTS2_REGIONS_BY_COUNTRY.get(pfx, [])})
    polygons = load_nuts2_polygons(needed)

    rows = []
    for off, pfx in sorted(landings):
        if off not in coord_lookup:
            print(f"  WARNING: offshore node '{off}' has no coordinates; "
                  f"cannot reconnect to {NUTS2_COUNTRY_TO_EMPIRE_NODE[pfx]}")
            continue
        lat, lon = coord_lookup[off]
        candidates = NUTS2_REGIONS_BY_COUNTRY.get(pfx, [])
        best, dist = nearest_nuts2_by_distance(lon, lat, candidates, polygons)
        if best is None:
            print(f"  WARNING: no NUTS2 polygon for '{pfx}' to reconnect '{off}'")
            continue
        rows.append({'NodeFrom': off, 'NodeTo': best})
        rows.append({'NodeFrom': best, 'NodeTo': off})
        print(f"  Reconnect: {off} <-> {best} "
              f"(was {NUTS2_COUNTRY_TO_EMPIRE_NODE[pfx]}; nearest coast {dist/1000:.0f} km)")

    return pd.DataFrame(rows) if rows else empty


def make_reference_corridor_fallbacks(reference_sets, ref_dl, iamc_links):
    """Restore national reference borders for which IAMC has no spatial link.

    IAMC is normally the allocation pattern. If it contains no corridor at all
    for a non-zero national reference border, dropping the border is worse than
    using a documented fallback. The replaced endpoint is connected to the
    nearest configured NUTS centroid and a warning is printed. Offshore hubs are
    handled separately by make_offshore_reconnections.
    """
    off_df = pd.read_excel(reference_sets, sheet_name='OffshoreNodes')
    offshore = {fix_node_name(str(n)) for n in off_df.iloc[:, 0].dropna()}
    node_to_prefix = {NUTS2_COUNTRY_TO_EMPIRE_NODE[p]: p
                      for p in NUTS2_COUNTRY_PREFIXES}
    region_to_country = {
        region: NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        for prefix, regions in NUTS2_REGIONS_BY_COUNTRY.items()
        for region in regions
    }

    def national(node):
        return region_to_country.get(node, node)

    iamc_national_pairs = set()
    for link in iamc_links:
        a, b = (resolve_iamc_endpoint(part.strip()) for part in link.split('>'))
        iamc_national_pairs.add(frozenset((national(a), national(b))))

    # Restore borders with positive existing capacity or positive expansion
    # headroom. A corridor can start at zero in InitialCapacity but become
    # buildable through MaxInstallCapacityRaw in later periods.
    positive_ref_pairs = set()
    trans_path = Path(reference_sets).with_name('Transmission.xlsx')
    if trans_path.exists():
        for sheet in ('InitialCapacity', 'MaxInstallCapacityRaw'):
            cap = pd.read_excel(
                trans_path, sheet_name=sheet, header=2).iloc[:, :4]
            cap.columns = ['FromNode', 'ToNode', 'Period', 'Value']
            cap['FromNode'] = cap['FromNode'].map(fix_node_name)
            cap['ToNode'] = cap['ToNode'].map(fix_node_name)
            cap['Value'] = pd.to_numeric(
                cap['Value'], errors='coerce').fillna(0.0)
            positive_ref_pairs |= {
                frozenset((r['FromNode'], r['ToNode']))
                for _, r in cap[cap['Value'] > 0].iterrows()
            }

    coords = pd.read_excel(reference_sets, sheet_name='Coords', header=2)
    coords.columns = ['Location', 'Latitude', 'Longitude']
    coords['Location'] = coords['Location'].map(fix_node_name)
    coord_lookup = {r['Location']: (float(r['Latitude']), float(r['Longitude']))
                    for _, r in coords.dropna().iterrows()}
    coord_lookup.update({r: (float(lat), float(lon))
                         for r, (lat, lon) in NUTS2_COORDS.items()})

    def distance(a, b):
        lat1, lon1 = coord_lookup[a]
        lat2, lon2 = coord_lookup[b]
        x = math.radians(lon2 - lon1) * math.cos(
            math.radians((lat1 + lat2) / 2))
        y = math.radians(lat2 - lat1)
        return 6371.0 * math.sqrt(x * x + y * y)

    def replacement(node, other):
        prefix = node_to_prefix.get(node)
        if prefix is None:
            return node
        candidates = list(NUTS2_REGIONS_BY_COUNTRY.get(prefix, []))
        if not candidates:
            return None
        if other in coord_lookup:
            candidates = [c for c in candidates if c in coord_lookup]
            if candidates:
                return min(candidates, key=lambda c: distance(c, other))
        return sorted(candidates)[0]

    ref_pairs = set()
    for _, row in ref_dl.iterrows():
        a, b = row['NodeFrom'], row['NodeTo']
        if a in offshore or b in offshore:
            continue
        if a not in node_to_prefix and b not in node_to_prefix:
            continue
        if positive_ref_pairs and frozenset((a, b)) not in positive_ref_pairs:
            continue
        ref_pairs.add(tuple(sorted((a, b))))

    rows = []
    for a, b in sorted(ref_pairs):
        if frozenset((a, b)) in iamc_national_pairs:
            continue
        new_a = replacement(a, b)
        new_b = replacement(b, new_a or a)
        if new_a is None or new_b is None or new_a == new_b:
            print(f"  [DirectionalLines][WARNING] could not restore reference "
                  f"corridor {a} <-> {b}; no regional fallback endpoint.")
            continue
        rows.extend([
            {'NodeFrom': new_a, 'NodeTo': new_b},
            {'NodeFrom': new_b, 'NodeTo': new_a},
        ])
        print(f"  [DirectionalLines][WARNING] IAMC has no spatial corridor for "
              f"reference border {a} <-> {b}; using nearest fallback "
              f"{new_a} <-> {new_b} so the national capacity is preserved.")
    return pd.DataFrame(rows, columns=['NodeFrom', 'NodeTo'])


def make_directional_lines(reference_sets, iamc_links):
    """Sheet: DirectionalLines

    1. Take all lines from reference that do NOT have Germany or Denmark as endpoint.
    2. Add all bidirectional lines derived from the IAMC file. For IAMC links,
       endpoints outside DE/DK are mapped to EMPIRE country names.
    3. Deduplicate.
    """
    print("[5/11] Creating DirectionalLines...")

    # --- Reference lines (remove any involving Germany or Denmark) ---
    ref_dl = pd.read_excel(reference_sets, sheet_name='DirectionalLines', header=2)
    ref_dl.columns = ['NodeFrom', 'NodeTo']  # normalise column names
    # Fix legacy no-space node names (e.g. "DoggerBank" -> "Dogger Bank")
    ref_dl['NodeFrom'] = ref_dl['NodeFrom'].map(fix_node_name)
    ref_dl['NodeTo']   = ref_dl['NodeTo'].map(fix_node_name)
    mask_de_dk = (
        ref_dl['NodeFrom'].isin(REPLACED_COUNTRY_NODES) |
        ref_dl['NodeTo'].isin(REPLACED_COUNTRY_NODES)
    )
    ref_kept = ref_dl[~mask_de_dk].copy()
    n_removed = mask_de_dk.sum()
    print(f"  Reference: kept {len(ref_kept)} lines, removed {n_removed} involving replaced country nodes")

    # --- IAMC lines (bidirectional, with endpoint resolution) ---
    iamc_pairs = set()
    for link in iamc_links:
        parts = link.split('>')
        node_from = resolve_iamc_endpoint(parts[0])
        node_to = resolve_iamc_endpoint(parts[1])
        if node_from == node_to:
            continue  # skip self-loops (shouldn't happen)
        iamc_pairs.add((node_from, node_to))
        iamc_pairs.add((node_to, node_from))

    iamc_rows = []
    for nf, nt in sorted(iamc_pairs):
        iamc_rows.append({'NodeFrom': nf, 'NodeTo': nt})
    iamc_df = pd.DataFrame(iamc_rows) if iamc_rows else pd.DataFrame(columns=['NodeFrom', 'NodeTo'])
    print(f"  IAMC: {len(iamc_df)} directional lines (bidirectional from {len(iamc_links)} IAMC links)")

    # --- Reconnect offshore nodes whose onshore landing was a disaggregated country ---
    offshore_df = make_offshore_reconnections(reference_sets, ref_dl)
    print(f"  Offshore: {len(offshore_df)} directional lines "
          f"({len(offshore_df) // 2} reconnected landings)")

    fallback_df = make_reference_corridor_fallbacks(
        reference_sets, ref_dl, iamc_links)
    print(f"  Reference fallback: {len(fallback_df)} directional lines "
          f"({len(fallback_df) // 2} borders absent from IAMC)")

    # --- Merge & deduplicate ---
    combined = pd.concat(
        [ref_kept, iamc_df, offshore_df, fallback_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=['NodeFrom', 'NodeTo']).reset_index(drop=True)
    combined = combined.sort_values(['NodeFrom', 'NodeTo']).reset_index(drop=True)
    print(f"  [OK] Total: {len(combined)} unique directional lines")
    return combined, ref_kept


def make_line_type_of_directional_lines(directional_lines_df, ref_kept_df,
                                        reference_sets, ac_links, dc_links):
    """Sheet: LineTypeOfDirectionalLines

    - Lines from the reference keep their original type.
    - New lines touching an offshore node (e.g. reconnected landings) → HVDC_Cable.
    - IAMC AC and DC capacities are represented independently. If both occur
      on one node pair, both type assignments are emitted in each direction.
    - New untyped lines touching an offshore node become HVDC_Cable.
    - All other new untyped lines become HVAC_OverheadLine.
    """
    print("[6/11] Creating LineTypeOfDirectionalLines...")

    # Offshore nodes — any line touching one is a subsea HVDC cable.
    off_df = pd.read_excel(reference_sets, sheet_name='OffshoreNodes')
    offshore_nodes = {fix_node_name(str(n)) for n in off_df.iloc[:, 0].dropna()}

    # Load original line types from reference
    ref_lt = pd.read_excel(reference_sets, sheet_name='LineTypeOfDirectionalLines', header=2)
    ref_lt.columns = ['FromNode', 'ToNode', 'LineType']
    ref_lt['FromNode'] = ref_lt['FromNode'].map(fix_node_name)
    ref_lt['ToNode'] = ref_lt['ToNode'].map(fix_node_name)
    ref_all_type_lookup = {}
    for _, row in ref_lt.iterrows():
        ref_all_type_lookup.setdefault(
            (row['FromNode'], row['ToNode']), set()).add(row['LineType'])

    # Remove lines involving Germany/Denmark from reference types
    mask_de_dk = (
        ref_lt['FromNode'].isin(REPLACED_COUNTRY_NODES) |
        ref_lt['ToNode'].isin(REPLACED_COUNTRY_NODES)
    )
    ref_lt_kept = ref_lt[~mask_de_dk].copy()

    # Build lookup for reference line types. Use a set so parallel reference
    # technologies on one endpoint pair are retained too.
    ref_type_lookup = {}
    for _, row in ref_lt_kept.iterrows():
        ref_type_lookup.setdefault(
            (row['FromNode'], row['ToNode']), set()).add(row['LineType'])

    # Build resolved AC/DC pairs in both orientations; IAMC normally stores an
    # undirected corridor only once.
    ac_pairs = set()
    for link in ac_links:
        parts = link.split('>')
        nf = resolve_iamc_endpoint(parts[0].strip())
        nt = resolve_iamc_endpoint(parts[1].strip())
        ac_pairs.add((nf, nt))
        ac_pairs.add((nt, nf))

    dc_pairs = set()
    for link in dc_links:
        parts = link.split('>')
        nf = resolve_iamc_endpoint(parts[0].strip())
        nt = resolve_iamc_endpoint(parts[1].strip())
        dc_pairs.add((nf, nt))
        dc_pairs.add((nt, nf))

    region_to_country = {
        region: NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        for prefix, regions in NUTS2_REGIONS_BY_COUNTRY.items()
        for region in regions
    }

    # Assign every applicable type. Do not use if/elif: HVAC and HVDC may
    # coexist on the same corridor.
    rows = []
    for _, row in directional_lines_df.iterrows():
        nf, nt = row['NodeFrom'], row['NodeTo']
        key = (nf, nt)
        types = set(ref_type_lookup.get(key, set()))
        if key in ac_pairs:
            types.add('HVAC_OverheadLine')
        if key in dc_pairs:
            types.add('HVDC_Cable')
        # A reference border missing from IAMC is restored with a nearest-region
        # fallback. In that one case inherit the original national line type.
        if not types:
            national_key = (region_to_country.get(nf, nf),
                            region_to_country.get(nt, nt))
            types |= ref_all_type_lookup.get(national_key, set())
            types |= ref_all_type_lookup.get(
                (national_key[1], national_key[0]), set())
        if not types:
            types.add('HVDC_Cable' if nf in offshore_nodes or nt in offshore_nodes
                      else 'HVAC_OverheadLine')
        for line_type in sorted(types):
            rows.append({'FromNode': nf, 'ToNode': nt,
                         'LineType': line_type})

    df = pd.DataFrame(rows).drop_duplicates().sort_values(
        ['FromNode', 'ToNode', 'LineType']).reset_index(drop=True)

    hvac = (df['LineType'] == 'HVAC_OverheadLine').sum()
    hvdc = (df['LineType'] == 'HVDC_Cable').sum()
    parallel = (df.groupby(['FromNode', 'ToNode']).size() > 1).sum()
    print(f"  [OK] {len(df)} type assignments ({hvac} HVAC, {hvdc} HVDC; "
          f"{parallel} directional endpoint pair(s) carry parallel types)")
    return df


def compute_adjacency_candidate_pairs(all_nodes):
    """Return a set of undirected frozensets {a, b} of adjacent modelled regions.

    Adjacency is land-border contact derived from the NUTS shapefile polygons:
      * fine region <-> fine region  (both endpoints in ALL_NUTS2)
      * fine region <-> neighbouring country node (country modelled as one node)

    Only nodes present in ``all_nodes`` are considered. Country nodes are mapped
    to their 2-letter NUTS0 code via IAMC_CODE_TO_EMPIRE_NODE. Used by the
    load-flow path to add HVAC candidate corridors between adjacent regions that
    have no existing line.
    """
    from config_loader import load_nuts2_polygons, load_nuts0_polygons

    node_set = set(all_nodes)
    fine = [c for c in ALL_NUTS2 if c in node_set]
    if not fine:
        return set()

    fine_polys = load_nuts2_polygons(fine)
    fine_with_poly = [c for c in fine if c in fine_polys]

    # Map each modelled country node to a 2-letter NUTS0 code (exclude fine regions).
    cc_by_node = {}
    for code, node in IAMC_CODE_TO_EMPIRE_NODE.items():
        if len(code) == 2 and node in node_set and node not in fine:
            cc_by_node.setdefault(node, code)
    country_ccs = sorted(set(cc_by_node.values()))
    country_polys = load_nuts0_polygons(country_ccs) if country_ccs else {}
    cc_to_node = {cc: node for node, cc in cc_by_node.items()}

    TOL = 1000.0  # metres — adjacent NUTS polygons share a boundary (distance ~0)

    pairs = set()

    # fine <-> fine
    for i, a in enumerate(fine_with_poly):
        ga = fine_polys[a]
        for b in fine_with_poly[i + 1:]:
            if ga.distance(fine_polys[b]) <= TOL:
                pairs.add(frozenset((a, b)))

    # fine <-> neighbouring country node
    for cc, cgeom in country_polys.items():
        node = cc_to_node.get(cc)
        if node is None:
            continue
        for a in fine_with_poly:
            if fine_polys[a].distance(cgeom) <= TOL:
                pairs.add(frozenset((a, node)))

    return pairs


def make_line_type_of_candidate_lines(bidirectional_lines_df, line_type_df):
    """Sheet: LineTypeOfCandidateLines

    Every bidirectional corridor is an expansion candidate under the load-flow
    path. Each candidate inherits the type of its directional line (HVAC or
    HVDC); the orientation matches BidirectionalLines.
    """
    print("[load-flow] Creating LineTypeOfCandidateLines...")
    type_lookup = {}
    for _, r in line_type_df.iterrows():
        type_lookup.setdefault((r['FromNode'], r['ToNode']), set()).add(r['LineType'])
    rows = []
    for _, r in bidirectional_lines_df.iterrows():
        nf, nt = r['NodeFrom'], r['NodeTo']
        types = (type_lookup.get((nf, nt), set()) |
                 type_lookup.get((nt, nf), set())) or {'HVAC_OverheadLine'}
        for lt in sorted(types):
            rows.append({'FromNode': nf, 'ToNode': nt, 'LineType': lt})
    df = pd.DataFrame(rows)
    hvdc = (df['LineType'] == 'HVDC_Cable').sum()
    print(f"  [OK] {len(df)} typed candidate corridors "
          f"({len(df) - hvdc} HVAC, {hvdc} HVDC)")
    return df


def make_generators_of_node(all_nodes, reference_sets, df_iamc):
    """Sheet: GeneratorsOfNode

    - For NUTS2 regions of DE/DK: derive from IAMC Capacity|Electricity data.
    - For all other nodes: copy from reference Sets.xlsx (excluding Germany/Denmark rows).
    """
    print("[7/11] Creating GeneratorsOfNode...")

    # --- Reference generators (non-DE/DK) ---
    ref_gon = pd.read_excel(reference_sets, sheet_name='GeneratorsOfNode', header=2)
    ref_gon.columns = ['Node', 'Generator']
    # Normalise legacy no-space node names (e.g. "GreatBrit." -> "Great Brit.")
    # so replaced-country rows are removed and the rest match the Nodes sheet.
    ref_gon['Node'] = ref_gon['Node'].map(fix_node_name)
    ref_gon_kept = ref_gon[~ref_gon['Node'].isin(REPLACED_COUNTRY_NODES)].drop_duplicates().copy()

    # Fix typo: 'Liginite' -> 'Lignite'
    ref_gon_kept['Generator'] = ref_gon_kept['Generator'].str.replace(
        r'^Liginite$', 'Lignite', regex=True)

    # For any node that has 'Gas', also add 'Gas CCGT' and 'Gas OCGT' if not
    # already present (v1.3.0 reference files use 'Gas' as a legacy name but
    # Generator.xlsx uses 'Gas CCGT' / 'Gas OCGT').
    existing_pairs = set(zip(ref_gon_kept['Node'], ref_gon_kept['Generator'].str.strip()))

    gas_nodes = ref_gon_kept.loc[ref_gon_kept['Generator'].str.strip() == 'Gas', 'Node'].tolist()
    if gas_nodes:
        extra = [
            {'Node': n, 'Generator': g}
            for n in gas_nodes
            for g in ('Gas CCGT', 'Gas OCGT')
            if (n, g) not in existing_pairs
        ]
        # Remove the legacy 'Gas' entries — not a valid EMPIRE generator name
        ref_gon_kept = ref_gon_kept[ref_gon_kept['Generator'].str.strip() != 'Gas'].copy()
        if extra:
            ref_gon_kept = pd.concat([ref_gon_kept, pd.DataFrame(extra)], ignore_index=True)
            existing_pairs.update((r['Node'], r['Generator']) for r in extra)
        print(f"  Replaced legacy 'Gas' with 'Gas CCGT' + 'Gas OCGT' for {len(set(gas_nodes))} nodes")


    print(f"  Reference: {len(ref_gon_kept)} generator assignments (replaced country nodes removed)")

    # --- IAMC generators for NUTS2 nodes ---
    iamc_rows = []
    for node in ALL_NUTS2:
        try:
            df_node = df_iamc.filter(region=node)
            node_vars = list(set(df_node.variable))
        except Exception:
            node_vars = []

        generators_at_node = set()
        for var in node_vars:
            if var in CAPACITY_TO_GENERATOR and CAPACITY_TO_GENERATOR[var] is not None:
                generators_at_node.add(CAPACITY_TO_GENERATOR[var])

        for gen in sorted(generators_at_node):
            iamc_rows.append({'Node': node, 'Generator': gen})

    iamc_gon = pd.DataFrame(iamc_rows) if iamc_rows else pd.DataFrame(columns=['Node', 'Generator'])
    print(f"  IAMC NUTS2: {len(iamc_gon)} generator assignments for {len(ALL_NUTS2)} NUTS2 regions")

    # IAMC describes the current fleet, not a siting ban for future investment.
    # Only hydro, offshore wind and wave retain an IAMC-derived regional
    # eligibility mask. Every other generator in the reference Generator set
    # is available at every NUTS node.
    resource_limited = {
        'Hydro regulated', 'Hydro run-of-the-river',
        'Wind offshore grounded', 'Wind offshore floating',
        'Wave',
    }
    ref_generators = pd.read_excel(reference_sets, sheet_name='Generators')
    unrestricted = sorted({
        str(v).strip()
        for col in ref_generators.columns
        for v in ref_generators[col].dropna()
        if str(v).strip() and str(v).strip() not in resource_limited
    })
    existing = set(zip(iamc_gon['Node'], iamc_gon['Generator']))
    unrestricted_rows = [
        {'Node': node, 'Generator': gen}
        for node in ALL_NUTS2
        for gen in unrestricted
        if (node, gen) not in existing
    ]
    if unrestricted_rows:
        iamc_gon = pd.concat(
            [iamc_gon, pd.DataFrame(unrestricted_rows)], ignore_index=True)
        print(f"  Added {len(unrestricted_rows)} unrestricted investment assignments "
              "for generators without spatial eligibility limits at NUTS nodes")

    # Ensure Wind onshore and Solar are assigned to every NUTS2 node.
    # generate_scenarios.py broadcasts windonshore.csv and solar.csv to ALL NUTS2
    # regions; EMPIRE requires every node in the scenario data to be in
    # GeneratorsOfNode, so we must guarantee these two technologies are present
    # even for regions with zero current IAMC capacity.
    mandatory_gens = ['Solar', 'Wind onshore']
    mandatory_rows = []
    existing = set(zip(iamc_gon['Node'], iamc_gon['Generator']))
    for node in ALL_NUTS2:
        for gen in mandatory_gens:
            if (node, gen) not in existing:
                mandatory_rows.append({'Node': node, 'Generator': gen})
    if mandatory_rows:
        mandatory_df = pd.DataFrame(mandatory_rows)
        iamc_gon = pd.concat([iamc_gon, mandatory_df], ignore_index=True)
        print(f"  Added {len(mandatory_rows)} mandatory Wind onshore/Solar assignments "
              f"for NUTS2 nodes missing them in IAMC data")

    # Ensure every NUTS2 node with 'Wind offshore grounded' also has
    # 'Wind offshore floating'. EMPIRE's north_sea mode maps windoffshore.csv
    # to Windoffshorefloating; both technologies must be in GeneratorsOfNode
    # for any coastal node or Pyomo rejects the stochastic availability index.
    floating_rows = []
    existing2 = set(zip(iamc_gon['Node'], iamc_gon['Generator']))
    grounded_nodes = {node for node, gen in existing2
                      if gen == 'Wind offshore grounded'}
    for node in grounded_nodes:
        if (node, 'Wind offshore floating') not in existing2:
            floating_rows.append({'Node': node, 'Generator': 'Wind offshore floating'})
    if floating_rows:
        iamc_gon = pd.concat([iamc_gon, pd.DataFrame(floating_rows)], ignore_index=True)
        print(f"  Added {len(floating_rows)} 'Wind offshore floating' assignments "
              f"to match 'Wind offshore grounded' nodes")

    # Ensure every NUTS2 node with 'Hydro regulated' also has
    # 'Hydro run-of-the-river'. The IAMC file merges both hydro types into a
    # single Capacity|Electricity|Hydro variable, so any node assigned
    # 'Hydro regulated' from IAMC data could also have run-of-river capacity.
    # hydroror.csv is broadcast to nodes with this generator, so they must
    # be registered here or Pyomo rejects the stochastic availability index.
    ror_rows = []
    existing3 = set(zip(iamc_gon['Node'], iamc_gon['Generator']))
    hydro_reg_nodes = {node for node, gen in existing3
                       if gen == 'Hydro regulated'}
    for node in hydro_reg_nodes:
        if (node, 'Hydro run-of-the-river') not in existing3:
            ror_rows.append({'Node': node, 'Generator': 'Hydro run-of-the-river'})
    if ror_rows:
        iamc_gon = pd.concat([iamc_gon, pd.DataFrame(ror_rows)], ignore_index=True)
        print(f"  Added {len(ror_rows)} 'Hydro run-of-the-river' assignments "
              f"to match 'Hydro regulated' nodes")

    # A reference country can contain existing ROR capacity even when IAMC has
    # no hydro observation for any of its regions (Denmark in NESP). Capacity
    # must still have a legal NUTS destination. Generator.xlsx will issue a
    # second warning and preserve the reference total using fallback weights.
    node_to_prefix = {v: k for k, v in NUTS2_COUNTRY_TO_EMPIRE_NODE.items()}
    ref_ror_countries = {
        node_to_prefix.get(r['Node'])
        for _, r in ref_gon.iterrows()
        if str(r['Generator']).strip() == 'Hydro run-of-the-river'
        and r['Node'] in REPLACED_COUNTRY_NODES
    }
    existing4 = set(zip(iamc_gon['Node'], iamc_gon['Generator']))
    fallback_ror_rows = []
    for prefix in sorted(p for p in ref_ror_countries if p):
        regions = NUTS2_REGIONS_BY_COUNTRY.get(prefix, [])
        if any((rg, 'Hydro run-of-the-river') in existing4 for rg in regions):
            continue
        country = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        print(f"  [GeneratorsOfNode][WARNING] {country} has run-of-river in the "
              "national reference but no IAMC hydro region; enabling ROR in all "
              f"{len(regions)} NUTS region(s) so the reference capacity can be "
              "distributed instead of dropped.")
        for rg in regions:
            fallback_ror_rows.append(
                {'Node': rg, 'Generator': 'Hydro run-of-the-river'})
            existing4.add((rg, 'Hydro run-of-the-river'))
    if fallback_ror_rows:
        iamc_gon = pd.concat(
            [iamc_gon, pd.DataFrame(fallback_ror_rows)], ignore_index=True)

    # Ensure every dispatchable generator a modeled country has in the reference
    # also exists in its NUTS2 nodes, so generate_generation_nuts2 can place that
    # country's reference capacity there (national totals are conserved).  Only
    # dispatchable technologies are added: renewables (Solar / Wind / Hydro) are
    # driven by per-node capacity-factor CSVs and are handled by the dedicated
    # blocks above, so adding them to regions that lack a source CF column would
    # break generate_scenarios.  Any reference capacity left without a slot is
    # reported by generate_generation_nuts2.
    SAFE_ORPHAN_GENERATORS = {
        'Bio', 'Coal', 'Gas CCGT', 'Gas OCGT', 'Oil',
        'Nuclear', 'Lignite', 'Waste', 'Geo',
        # CCS variants are dispatchable thermal units too (no CF profile
        # needed). Without them a disaggregated country whose reference node
        # offers CCS (e.g. Spain: zero initial capacity but a real
        # MaxBuiltCapacity budget) would lose the option to invest in CCS.
        'Bio CCS', 'Coal CCS', 'Gas CCS', 'Lignite CCS',
    }
    node_to_prefix = {v: k for k, v in NUTS2_COUNTRY_TO_EMPIRE_NODE.items()}
    # Generators each modeled country has in the reference GeneratorsOfNode
    # (apply the same normalisation used for the kept reference rows).
    country_ref_gens: dict[str, set] = {}
    for _, r in ref_gon.iterrows():
        node = r['Node']
        if node not in REPLACED_COUNTRY_NODES:
            continue
        prefix = node_to_prefix.get(node)
        gen = str(r['Generator']).strip()
        if gen == 'Liginite':
            gen = 'Lignite'
        if gen == 'Gas':
            country_ref_gens.setdefault(prefix, set()).update({'Gas CCGT', 'Gas OCGT'})
            continue
        country_ref_gens.setdefault(prefix, set()).add(gen)

    existing_nuts2 = set(zip(iamc_gon['Node'], iamc_gon['Generator']))
    orphan_rows = []
    for prefix, gens in country_ref_gens.items():
        regions = NUTS2_REGIONS_BY_COUNTRY.get(prefix, [])
        for gen in sorted(gens & SAFE_ORPHAN_GENERATORS):
            already = any((rg, gen) in existing_nuts2 for rg in regions)
            if already:
                continue
            for rg in regions:
                orphan_rows.append({'Node': rg, 'Generator': gen})
                existing_nuts2.add((rg, gen))
    if orphan_rows:
        iamc_gon = pd.concat([iamc_gon, pd.DataFrame(orphan_rows)], ignore_index=True)
        added = sorted({r['Generator'] for r in orphan_rows})
        print(f"  Added {len(orphan_rows)} dispatchable orphan assignments "
              f"({', '.join(added)}) so reference capacity has a NUTS2 home")

    # --- Combine ---
    combined = pd.concat([ref_gon_kept, iamc_gon], ignore_index=True)

    # Drop offshore-wind generators for nodes that have no windoffshore.csv
    # profile column. EMPIRE maps the offshore availability CSV column-by-column
    # to nodes; a node with a 'Wind offshore *' generator but no column would get
    # zero availability and could never generate (a silent modelling gap). We
    # remove the phantom generator instead of inventing a profile.
    windoff_path = Path(reference_sets).parent / 'ScenarioData' / 'windoffshore.csv'
    if windoff_path.exists():
        header = pd.read_csv(windoff_path, nrows=0).columns.tolist()
        profile_codes = {c for c in header if c.lower() != 'time'}
        # Nodes that will end up with a profile column:
        #  * a dedicated offshore hub whose abbreviation is a column (MF, SN1...)
        #  * a country/aggregate node whose code is a column (FR->France, EU->EU)
        #  * a NUTS2 fine region whose country prefix is a column (broadcast source)
        resolved = {}
        for code in profile_codes:
            if code in OFFSHORE_HUB_CODE_TO_NODE:
                resolved[code] = OFFSHORE_HUB_CODE_TO_NODE[code]
            else:
                resolved[code] = IAMC_CODE_TO_EMPIRE_NODE.get(code, code)
        profiled_nodes = {_despace(n) for n in resolved.values()}
        # A column may name a country node that NUTS disaggregation replaced
        # rather than its two-letter code: preprocess_norway.py renames the 'NO'
        # column to 'Norway' so EMPIRE's scenario_random.py will not re-split it
        # back into NO1-NO5.  Translate such a name to its prefix so the rule
        # below reaches that country's NUTS regions, exactly as a bare 'DK'
        # does; without it every Norwegian region silently loses its offshore
        # generators a few lines further down.
        # Resolve through the node name, not the raw column code: the column is
        # 'GB' but Great Britain's NUTS2 regions carry the 'UK' prefix, so
        # node_to_prefix['GB'] misses and every UK region would lose its
        # offshore generators.  resolved['GB'] is already 'Great Brit.', which
        # node_to_prefix does map to 'UK'.  Bare codes that name no EMPIRE node
        # ('DK') still fall through to themselves, as before.
        node_to_prefix = {v: k for k, v in NUTS2_COUNTRY_TO_EMPIRE_NODE.items()}

        def _code_to_prefix(code):
            return node_to_prefix.get(resolved[code], code)

        profile_prefixes = {_code_to_prefix(c) for c in profile_codes}
        profiled_nodes |= {
            _despace(region) for region in ALL_NUTS2
            if region[:2] in profile_prefixes
        }

        # A column that reaches no node at all is how this bug hid: the hub
        # columns were silently treated as unprofiled and their generators
        # stripped.  Say so — but only for codes that really reach nothing.  A
        # code naming a replaced country (DK -> Denmark) still profiles that
        # country's NUTS regions through the prefix rule above.
        all_node_keys = {_despace(n) for n in combined['Node'].astype(str)}
        unmatched = sorted(
            code for code, node in resolved.items()
            if _despace(node) not in all_node_keys
            and not any(r[:2] == _code_to_prefix(code) for r in ALL_NUTS2)
        )
        if unmatched:
            print(f"  [GeneratorsOfNode][WARNING] {len(unmatched)} windoffshore.csv "
                  f"column(s) reach no node: {', '.join(unmatched)}. If one names an "
                  "offshore hub, add it to OFFSHORE_HUB_CODE_TO_NODE; if it names a "
                  "country whose node was replaced by NUTS regions, those regions "
                  "are about to lose their offshore generators.")

        off_mask = combined['Generator'].astype(str).str.lower().str.contains('offshore')
        node_keys = combined['Node'].astype(str).map(_despace)
        drop_mask = off_mask & ~node_keys.isin(profiled_nodes)
        if drop_mask.any():
            dropped = sorted(set(combined.loc[drop_mask, 'Node']))
            combined = combined[~drop_mask].copy()
            print(f"  Removed offshore-wind generators for {len(dropped)} node(s) "
                  f"with no windoffshore.csv profile: {', '.join(dropped)}")

    combined = combined.sort_values(['Node', 'Generator']).reset_index(drop=True)
    print(f"  [OK] Total: {len(combined)} generator-node assignments")
    return combined


def make_bidirectional_lines(directional_lines_df):
    """Sheet: BidirectionalLines - unique undirected corridors.

    Derived from DirectionalLines by keeping only one direction per pair
    (the one where NodeFrom < NodeTo alphabetically).
    """
    print("[9/11] Creating BidirectionalLines...")
    seen = set()
    rows = []
    for _, row in directional_lines_df.iterrows():
        nf, nt = row['NodeFrom'], row['NodeTo']
        key = tuple(sorted([nf, nt]))
        if key not in seen:
            seen.add(key)
            rows.append({'NodeFrom': key[0], 'NodeTo': key[1]})
    df = pd.DataFrame(rows).sort_values(['NodeFrom', 'NodeTo']).reset_index(drop=True)
    print(f"  [OK] {len(df)} bidirectional corridors")
    return df


def make_coords(all_nodes, reference_sets):
    """Sheet: Coords

    - For nodes from the reference (except Germany/Denmark): use reference coords.
    - For DE/DK NUTS2 regions: use the hardcoded NUTS2_COORDS table.
    """
    print("[8/11] Creating Coords...")

    ref_coords = pd.read_excel(reference_sets, sheet_name='Coords', header=2)
    ref_coords.columns = ['Location', 'Latitude', 'Longitude']

    # Fix legacy no-space node names before filtering
    ref_coords['Location'] = ref_coords['Location'].map(fix_node_name)
    # Remove Germany and Denmark rows from reference
    ref_kept = ref_coords[~ref_coords['Location'].isin(REPLACED_COUNTRY_NODES)].copy()

    # Add NUTS2 coords
    nuts2_rows = []
    for code in ALL_NUTS2:
        if code in NUTS2_COORDS:
            lat, lon = NUTS2_COORDS[code]
            nuts2_rows.append({'Location': code, 'Latitude': lat, 'Longitude': lon})
        else:
            print(f"  ⚠ No coordinates for {code}")

    nuts2_df = pd.DataFrame(nuts2_rows)

    combined = pd.concat([ref_kept, nuts2_df], ignore_index=True)
    combined = combined.sort_values('Location').reset_index(drop=True)
    print(f"  [OK] {len(combined)} nodes with coordinates")
    return combined


# =========================================================================
# Main
# =========================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Generate Sets.xlsx for EMPIRE with NUTS2 resolution for selected countries'
    )
    parser.add_argument('iamc_file', help='Path to IAMC data file (CSV or Excel)')
    parser.add_argument('scenario', help='Scenario name to filter')
    parser.add_argument('output_file', help='Output Sets.xlsx path')
    parser.add_argument('--reference', default='data/EMPIRE/v_52/Sets.xlsx',
                        help='Reference Sets.xlsx (default: data/EMPIRE/v_52/Sets.xlsx)')
    args = parser.parse_args()

    print("=" * 80)
    print("SETS.XLSX GENERATOR — NUTS2 for selected countries")
    print("=" * 80)
    print(f"Input:     {args.iamc_file}")
    print(f"Scenario:  {args.scenario}")
    print(f"Reference: {args.reference}")
    print(f"Output:    {args.output_file}")
    print(f"NUTS2 countries: {', '.join(NUTS2_COUNTRY_PREFIXES)}")
    print()

    # --- Load IAMC data ---
    df_iamc, iamc_nodes, iamc_links, ac_links, dc_links = load_iamc_data(
        args.iamc_file, args.scenario)

    # --- Build sheets ---
    simple_sheets = {}

    # 1. Nodes
    nodes_df, all_nodes = make_nodes_sheet(args.reference)
    simple_sheets['Nodes'] = nodes_df

    # 2. OffshoreNodes
    simple_sheets['OffshoreNodes'] = make_offshore_nodes(args.reference)

    # Countries / NodesOfCountry — national-level limits (NUTS_national_limits_data_spec.md)
    simple_sheets['Countries'] = make_countries()
    non_df = make_nodes_of_country(all_nodes)

    # 3. Horizon
    simple_sheets['Horizon'] = make_horizon(args.reference)

    # 4-7. Storage, Technology, Generators, LineType — copy as-is
    print("[INFO] Copying Storage, Technology, Generators, LineType from reference...")
    for sheet in ['Storage', 'Technology', 'Generators', 'LineType']:
        df = copy_sheet_simple(args.reference, sheet)
        if sheet == 'Generators':
            # Strip trailing/leading whitespace and deduplicate each column
            # (reference may have 'Bio ', 'Lignite CCS ', etc. alongside 'Bio', 'Lignite CCS')
            cleaned = {}
            for col in df.columns:
                series = df[col].astype(str).str.strip().replace('nan', float('nan'))
                series = series.dropna().drop_duplicates().reset_index(drop=True)
                cleaned[col] = series
            df = pd.DataFrame(cleaned)
            print(f"  [OK] Generators: cleaned (stripped whitespace, deduplicated each column)")
        simple_sheets[sheet] = df
        print(f"  [OK] {sheet}: {len(simple_sheets[sheet])} rows")

    # 8. HourOfSeason — raw copy (preserves metadata)
    print("[INFO] Copying HourOfSeason from reference (raw)...")
    hour_df = copy_sheet_raw(args.reference, 'HourOfSeason')

    # 9. GeneratorsOfTechnology — raw copy (preserves metadata), deduplicated
    print("[INFO] Copying GeneratorsOfTechnology from reference (raw)...")
    gen_tech_df = copy_sheet_raw(args.reference, 'GeneratorsOfTechnology')
    gen_tech_df = gen_tech_df.drop_duplicates()

    # --- Metadata sheets ---
    metadata_sheets = {}

    # NodesOfCountry
    metadata_sheets['NodesOfCountry'] = add_metadata_header(
        non_df, source='-',
        description='Assignment of nodes to countries for national-level limits')

    # StorageOfNodes
    son_df = make_storage_of_nodes(all_nodes, args.reference)
    metadata_sheets['StorageOfNodes'] = add_metadata_header(
        son_df, source='-', description='Assignment of storage types to nodes')

    # DirectionalLines
    dl_df, ref_kept_dl = make_directional_lines(args.reference, iamc_links)

    # Cross-check: drop any line whose endpoint is not in the Node set
    node_set = set(all_nodes)
    orphan_mask = (
        ~dl_df['NodeFrom'].isin(node_set) |
        ~dl_df['NodeTo'].isin(node_set)
    )
    if orphan_mask.any():
        orphan_lines = dl_df[orphan_mask][['NodeFrom', 'NodeTo']].values.tolist()
        missing_nodes = sorted(
            n for pair in orphan_lines for n in pair if n not in node_set
        )
        print(f"\n  WARNING: {orphan_mask.sum()} directional line(s) reference node(s) "
              f"not in the Node set and will be removed:")
        for nf, nt in orphan_lines:
            print(f"    {nf} -> {nt}")
        print(f"  Missing nodes: {', '.join(dict.fromkeys(missing_nodes))}")
        dl_df = dl_df[~orphan_mask].reset_index(drop=True)

    # Load-flow: inject adjacency-only HVAC candidates as new 0-capacity arcs so
    # EMPIRE can build them. Every adjacent region pair with no existing corridor
    # becomes a new bidirectional HVAC line.
    if _LF['enabled']:
        existing_undirected = {
            frozenset((r['NodeFrom'], r['NodeTo'])) for _, r in dl_df.iterrows()
        }
        adj_pairs = compute_adjacency_candidate_pairs(all_nodes)
        new_pairs = [p for p in adj_pairs if p not in existing_undirected]
        if new_pairs:
            new_rows = []
            for p in new_pairs:
                a, b = tuple(p)
                new_rows.append({'NodeFrom': a, 'NodeTo': b})
                new_rows.append({'NodeFrom': b, 'NodeTo': a})
            dl_df = pd.concat([dl_df, pd.DataFrame(new_rows)], ignore_index=True)
            dl_df = dl_df.drop_duplicates(subset=['NodeFrom', 'NodeTo'])
            dl_df = dl_df.sort_values(['NodeFrom', 'NodeTo']).reset_index(drop=True)
        print(f"[load-flow] Added {len(new_pairs)} adjacency-only HVAC candidate "
              f"corridors ({len(adj_pairs)} adjacent pairs, "
              f"{len(adj_pairs) - len(new_pairs)} already existing)")

    metadata_sheets['DirectionalLines'] = add_metadata_header(
        dl_df, source=f'{args.iamc_file} + reference',
        description='Decleration of arcs in both directions')

    # LineTypeOfDirectionalLines
    lt_df = make_line_type_of_directional_lines(
        dl_df, ref_kept_dl, args.reference, ac_links, dc_links)
    metadata_sheets['LineTypeOfDirectionalLines'] = add_metadata_header(
        lt_df, source=f'{args.iamc_file} + reference',
        description='Assignment of line type to transmission lines')

    # GeneratorsOfNode
    gon_df = make_generators_of_node(all_nodes, args.reference, df_iamc)
    metadata_sheets['GeneratorsOfNode'] = add_metadata_header(
        gon_df, source=f'{args.iamc_file} + reference',
        description='Assignment of generator types to nodes')

    # BidirectionalLines
    bl_df = make_bidirectional_lines(dl_df)
    metadata_sheets['BidirectionalLines'] = add_metadata_header(
        bl_df, source=f'{args.iamc_file} + reference',
        description='List of unique undirected transmission corridors')

    # LineTypeOfCandidateLines (load-flow only): every corridor is a candidate.
    if _LF['enabled']:
        ltc_df = make_line_type_of_candidate_lines(bl_df, lt_df)
        metadata_sheets['LineTypeOfCandidateLines'] = add_metadata_header(
            ltc_df, source=f'{args.iamc_file} + reference',
            description='Assignment of line type to candidate transmission lines')

    # Coords
    coords_df = make_coords(all_nodes, args.reference)
    metadata_sheets['Coords'] = add_metadata_header(
        coords_df, source='reference + hardcoded NUTS2 centroids',
        description='Location of nodes for plotting')

    # --- Write output ---
    print(f"\n[WRITE] Writing {args.output_file}...")
    with pd.ExcelWriter(args.output_file, engine='openpyxl') as writer:
        # Simple sheets (with proper headers)
        for name, df in simple_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

        # HourOfSeason: raw copy
        hour_df.to_excel(writer, sheet_name='HourOfSeason', index=False, header=False)

        # Metadata sheets
        for name, df in metadata_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False, header=False)

        # GeneratorsOfTechnology: raw copy
        gen_tech_df.to_excel(writer, sheet_name='GeneratorsOfTechnology', index=False, header=False)

    # --- Summary ---
    print()
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Output: {args.output_file}")
    print()
    print("Sheets created:")
    print(f"  Nodes:                       {len(simple_sheets['Nodes'])} nodes")
    print(f"  OffshoreNodes:               {len(simple_sheets['OffshoreNodes'])} offshore nodes")
    print(f"  Countries:                   {len(simple_sheets['Countries'])} countries")
    print(f"  NodesOfCountry:              {len(non_df)} node assignments")
    print(f"  Horizon:                     {len(simple_sheets['Horizon'])} periods")
    print(f"  Storage:                     {len(simple_sheets['Storage'])} types")
    print(f"  Technology:                  {len(simple_sheets['Technology'])} types")
    print(f"  Generators:                  {len(simple_sheets['Generators'])} generators")
    print(f"  LineType:                    {len(simple_sheets['LineType'])} types")
    print(f"  HourOfSeason:                (copied from reference)")
    print(f"  StorageOfNodes:              {len(son_df)} assignments")
    print(f"  DirectionalLines:            {len(dl_df)} links")
    print(f"  LineTypeOfDirectionalLines:  {len(lt_df)} assignments")
    print(f"  BidirectionalLines:          {len(bl_df)} corridors")
    if _LF['enabled']:
        print(f"  LineTypeOfCandidateLines:    {len(bl_df)} candidate corridors")
    print(f"  GeneratorsOfNode:            {len(gon_df)} assignments")
    print(f"  GeneratorsOfTechnology:      (copied from reference)")
    print(f"  Coords:                      {len(coords_df)} coordinates")
    print()
    print("[OK] Done!")
    print("=" * 80)


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Usage: python generate_sets_nuts2.py <iamc_file> <scenario> <output_file> [--reference <sets.xlsx>]")
        print()
        print("Example:")
        print('  python generate_sets_nuts2.py data/IAMC/NorthSea/Denmark_Germany.csv "Transmission v1.0" Sets_NUTS2_DK_DE.xlsx')
        sys.exit(1)
    main()
