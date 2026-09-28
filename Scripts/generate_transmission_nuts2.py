"""
Generate Transmission.xlsx for EMPIRE with NUTS2 resolution for selected countries.

Reads the generated Sets_NUTS2_DK_DE.xlsx for node lists, directional/bidirectional
lines and periods. Combines reference Transmission.xlsx data with IAMC data for new
NUTS2 corridors.

Usage:
    python generate_transmission_nuts2.py <iamc_file> <scenario> <sets_file> <output_file>
                                          [--reference <Transmission.xlsx>]

Example:
    python generate_transmission_nuts2.py data/IAMC/NorthSea/Denmark_Germany.csv
        "Transmission v1.0" Sets_NUTS2_DK_DE.xlsx Transmission_NUTS2_DK_DE.xlsx
"""

import sys
import argparse
import math
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
ALL_NUTS2 = _cfg['ALL_NUTS2']
_LF = _load_load_flow_config()

REPLACED_COUNTRY_NODES_LOWER = {
    NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix].lower()
    for prefix in NUTS2_COUNTRY_PREFIXES
}

IAMC_CODE_TO_EMPIRE_NODE = {
    'AT': 'Austria',   'BE': 'Belgium',  'BE33': 'Belgium',
    'BG': 'Bulgaria',  'CH': 'Switzerland', 'CZ': 'Czech R',
    'DE': 'Germany',
    'EE': 'Estonia',   'ES': 'Spain',    'FI': 'Finland',
    'FR': 'France',    'FRF1': 'France', 'FRF3': 'France',
    'GB': 'Great Brit.', 'GR': 'Greece', 'HR': 'Croatia',
    'HU': 'Hungary',   'IE': 'Ireland',  'IT': 'Italy',
    'LT': 'Lithuania', 'LU': 'Luxemb.',  'LV': 'Latvia',
    'MK': 'Macedonia', 'NL': 'Netherlands',
    'NL11': 'Netherlands', 'NL21': 'Netherlands',
    'NL22': 'Netherlands', 'NL42': 'Netherlands',
    'NO': 'NO1', 'NO01': 'NO1', 'NO02': 'NO2', 'NO03': 'NO3',
    'NO04': 'NO4', 'NO05': 'NO5', 'NO09': 'NO2',
    'PL': 'Poland',    'PT': 'Portugal', 'RO': 'Romania',
    'RS': 'Serbia',    'SE': 'Sweden',   'SI': 'Slovenia',
    'SK': 'Slovakia',  'UK': 'Great Brit.', 'UKF3': 'Great Brit.',
}


def resolve_iamc_endpoint(code):
    """Map an IAMC endpoint code to an EMPIRE node name."""
    if code in ALL_NUTS2:
        return code
    if code in IAMC_CODE_TO_EMPIRE_NODE:
        return IAMC_CODE_TO_EMPIRE_NODE[code]
    cc = code[:2]
    if cc in IAMC_CODE_TO_EMPIRE_NODE:
        return IAMC_CODE_TO_EMPIRE_NODE[cc]
    return code


# =============================================================================
# Helpers
# =============================================================================

def add_metadata_header(df, source='-', description=''):
    """Wrap a DataFrame with Source / Description metadata rows."""
    cols = df.columns.tolist()
    n = len(cols)
    rows = [
        [f'Source: {source}'] + [np.nan] * (n - 1),
        [f'Description: {description}'] + [np.nan] * (n - 1),
        cols,
    ]
    for _, row in df.iterrows():
        rows.append(row.tolist())
    return pd.DataFrame(rows, columns=[f'c{i}' for i in range(n)])


def haversine_km(lat1, lon1, lat2, lon2):
    """Great-circle distance in km between two (lat, lon) points."""
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def normalise(name):
    """Lowercase, strip spaces — used for fuzzy key matching across files."""
    return str(name).strip().lower().replace(' ', '')


def build_lookup(df, col_from, col_to, *value_cols):
    """Build dict (norm_from, norm_to) -> {value_col: value, ...}."""
    lookup = {}
    for _, row in df.iterrows():
        key = (normalise(row[col_from]), normalise(row[col_to]))
        lookup[key] = {c: row[c] for c in value_cols}
    return lookup


# Map every NUTS2 code to the (normalised) EMPIRE country node it belongs to.
# Used to transfer the capacity of a severed offshore<->country line onto the
# offshore<->NUTS2 corridor that replaces it.
NUTS2_TO_COUNTRY_NORM = {
    normalise(code): normalise(NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix])
    for prefix, codes in NUTS2_REGIONS_BY_COUNTRY.items()
    for code in codes
}

# Normalised names of the disaggregated country nodes. Use normalise() (not the
# space-preserving REPLACED_COUNTRY_NODES_LOWER) so multi-word nodes such as
# 'Great Brit.' match the normalised reference endpoints.
REPLACED_COUNTRY_NORM = set(NUTS2_TO_COUNTRY_NORM.values())


def load_offshore_landing_capacity(reference_trans, offshore_nodes, sheet):
    """Build {(offshore_norm, country_norm, period): value} for reference rows
    linking an offshore node to a country that is being disaggregated to NUTS2.

    These offshore<->country corridors disappear when the country splits into
    NUTS2 nodes (generate_sets_nuts2 reconnects the offshore node to its nearest
    coastal NUTS2 region instead). This lookup lets the new offshore<->NUTS2
    corridor inherit the reference capacity of the original landing.
    """
    ref = pd.read_excel(reference_trans, sheet_name=sheet, header=2).iloc[:, :4]
    ref.columns = ['From', 'To', 'Period', 'Value']
    off_norm = {normalise(n) for n in offshore_nodes}
    lookup = {}
    for _, row in ref.iterrows():
        nf, nt = normalise(row['From']), normalise(row['To'])
        period = int(row['Period'])
        if nf in off_norm and nt in REPLACED_COUNTRY_NORM:
            lookup[(nf, nt, period)] = row['Value']
        elif nt in off_norm and nf in REPLACED_COUNTRY_NORM:
            lookup[(nt, nf, period)] = row['Value']
    return lookup, off_norm


def offshore_landing_value(knf, knt, period, off_norm, landing_lookup):
    """If (knf, knt) is an offshore<->NUTS2 corridor whose NUTS2 endpoint belongs
    to a disaggregated country, return the inherited reference landing capacity
    for that period (else None)."""
    if knf in off_norm and knt in NUTS2_TO_COUNTRY_NORM:
        off, country = knf, NUTS2_TO_COUNTRY_NORM[knt]
    elif knt in off_norm and knf in NUTS2_TO_COUNTRY_NORM:
        off, country = knt, NUTS2_TO_COUNTRY_NORM[knf]
    else:
        return None
    return landing_lookup.get((off, country, period))


# =============================================================================
# Data loaders
# =============================================================================

def load_sets(sets_file):
    """Load key sheets from the generated Sets_NUTS2_DK_DE.xlsx."""
    print(f"[INFO] Reading Sets file: {sets_file}")
    dir_lines = pd.read_excel(sets_file, sheet_name='DirectionalLines', header=2)
    dir_lines.columns = ['NodeFrom', 'NodeTo']

    bi_lines = pd.read_excel(sets_file, sheet_name='BidirectionalLines', header=2)
    bi_lines.columns = ['NodeFrom', 'NodeTo']

    horizon = pd.read_excel(sets_file, sheet_name='Horizon')
    periods = sorted(horizon.iloc[:, 0].dropna().astype(int).tolist())

    coords = pd.read_excel(sets_file, sheet_name='Coords', header=2)
    coords.columns = ['Location', 'Latitude', 'Longitude']
    coords_dict = {normalise(r['Location']): (r['Latitude'], r['Longitude'])
                   for _, r in coords.iterrows()}

    # One endpoint pair may carry both HVAC and HVDC.
    lt_df = pd.read_excel(sets_file, sheet_name='LineTypeOfDirectionalLines', header=2)
    lt_df.columns = ['FromNode', 'ToNode', 'LineType']
    linetype_lookup = {}
    for _, row in lt_df.iterrows():
        key = (normalise(row['FromNode']), normalise(row['ToNode']))
        linetype_lookup.setdefault(key, set()).add(str(row['LineType']).strip())

    print(f"  ✓ {len(dir_lines)} directional lines, {len(bi_lines)} bidirectional corridors, "
          f"{len(periods)} periods")
    return dir_lines, bi_lines, periods, coords_dict, linetype_lookup


_CAPACITY_VAR_AC = 'Network|Electricity|Capacity'
_CAPACITY_VAR_DC = 'Network|Electricity|DC|Capacity'
_REACTANCE_VAR = 'Network|Electricity|Reactance'


def load_iamc_capacity(iamc_file, scenario):
    """Load AC and DC transmission capacity from IAMC.

    Returns two separate dicts, each mapping (norm_from, norm_to) -> MW:
        ac_lookup  — from Network|Electricity|Capacity       (HVAC corridors)
        dc_lookup  — from Network|Electricity|DC|Capacity    (HVDC corridors)

    HVAC and HVDC are distinct line types in EMPIRE and must not be merged.
    """
    print(f"[INFO] Loading IAMC data from {iamc_file} (scenario: {scenario})...")
    from iamc_utils import load_iamc
    df_iamcobj = load_iamc(iamc_file)
    df = df_iamcobj.filter(
        scenario=scenario,
        variable=[_CAPACITY_VAR_AC, _CAPACITY_VAR_DC],
    ).as_pandas()

    ac_lookup = {}
    dc_lookup = {}

    if not df.empty:
        df = df[df['region'].str.contains('>', na=False)].copy()
        for _, row in df.iterrows():
            parts = row['region'].split('>')
            nf = resolve_iamc_endpoint(parts[0].strip())
            nt = resolve_iamc_endpoint(parts[1].strip())
            val_mw = row['value'] * 1000 if row['unit'] == 'GW' else row['value']
            key = (normalise(nf), normalise(nt))
            if row['variable'] == _CAPACITY_VAR_DC:
                dc_lookup[key] = val_mw
            else:
                ac_lookup[key] = val_mw

    print(f"  ✓ {len(ac_lookup)} HVAC records ('{_CAPACITY_VAR_AC}')")
    print(f"  ✓ {len(dc_lookup)} HVDC records ('{_CAPACITY_VAR_DC}')")
    print(f"  ✓ {len(ac_lookup) + len(dc_lookup)} IAMC link capacity records (total)")
    return ac_lookup, dc_lookup


def line_types_for_pair(linetype_lookup, knf, knt):
    """Return all transmission types assigned to an undirected node pair."""
    return (set(linetype_lookup.get((knf, knt), set())) |
            set(linetype_lookup.get((knt, knf), set()))) or {
                'HVAC_OverheadLine'}


def iamc_capacity_for_type(ac_lookup, dc_lookup, knf, knt, line_type):
    """Return one undirected IAMC capacity value for a typed corridor.

    If IAMC contains both orientations, average them instead of double-counting
    the same physical capacity.
    """
    lookup = dc_lookup if 'HVDC' in str(line_type) else ac_lookup
    values = [lookup[key] for key in ((knf, knt), (knt, knf)) if key in lookup]
    return float(sum(values) / len(values)) if values else 0.0


def make_typed_corridors(bi_lines, linetype_lookup, ac_lookup, dc_lookup):
    """Expand each endpoint pair into one row per HVAC/HVDC technology."""
    rows = []
    for _, r in bi_lines.iterrows():
        nf, nt = r['NodeFrom'], r['NodeTo']
        knf, knt = normalise(nf), normalise(nt)
        for line_type in sorted(line_types_for_pair(linetype_lookup, knf, knt)):
            rows.append({
                'NodeFrom': nf,
                'NodeTo': nt,
                'LineType': line_type,
                'IAMCCapacity': iamc_capacity_for_type(
                    ac_lookup, dc_lookup, knf, knt, line_type),
            })
    return pd.DataFrame(rows)


def national_endpoint(node):
    """Map a NUTS endpoint back to its national reference node key."""
    key = normalise(node)
    return NUTS2_TO_COUNTRY_NORM.get(key, key)


def national_pair(node_from, node_to):
    """Canonical national endpoint pair used to match the reference workbook."""
    return tuple(sorted((national_endpoint(node_from), national_endpoint(node_to))))


def read_reference_period_totals(reference_trans, sheet_name):
    """Read a national transmission sheet as (undirected pair, period) -> MW."""
    ref = pd.read_excel(reference_trans, sheet_name=sheet_name, header=2).iloc[:, :4]
    ref.columns = ['FromNode', 'ToNode', 'Period', 'Value']
    ref['Value'] = pd.to_numeric(ref['Value'], errors='coerce').fillna(0.0)
    lookup = {}
    for _, row in ref.iterrows():
        key = (national_pair(row['FromNode'], row['ToNode']), int(row['Period']))
        if key in lookup and not math.isclose(lookup[key], float(row['Value'])):
            raise ValueError(
                f"{sheet_name} has conflicting duplicate rows for {key}: "
                f"{lookup[key]} and {row['Value']}")
        lookup[key] = float(row['Value'])
    return lookup


def allocate_total(total, records):
    """Allocate an extensive national total by IAMC spatial/type weights."""
    weights = [max(float(r['IAMCCapacity']), 0.0) for r in records]
    weight_sum = sum(weights)
    used_fallback = weight_sum <= 0
    if used_fallback:
        weights = [1.0] * len(records)
        weight_sum = float(len(records))
    values = [total * weight / weight_sum for weight in weights]
    if values:
        values[-1] = total - sum(values[:-1])
    return values, used_fallback


def make_capacity_sheets(bi_lines, periods, reference_trans, sheet_name,
                         value_col, ac_lookup, dc_lookup, linetype_lookup,
                         iamc_multiplier):
    """Build aggregate and type-aware transmission-capacity sheets.

    A period-specific national reference value is authoritative whenever it
    exists; IAMC supplies only regional/type allocation weights. Corridors with
    no national counterpart, notably internal NUTS links, retain IAMC levels.
    """
    typed_corridors = make_typed_corridors(
        bi_lines, linetype_lookup, ac_lookup, dc_lookup)
    ref_lookup = read_reference_period_totals(reference_trans, sheet_name)

    groups = {}
    for record in typed_corridors.to_dict('records'):
        groups.setdefault(
            national_pair(record['NodeFrom'], record['NodeTo']), []).append(record)

    # Never silently drop a positive reference corridor touching a country that
    # was replaced by NUTS nodes. generate_sets_nuts2 normally creates a nearest
    # fallback when IAMC has no spatial pattern; this assertion protects direct
    # calls with an inconsistent Sets workbook too.
    missing_reference = [
        (pair, period, value)
        for (pair, period), value in ref_lookup.items()
        if value > 0
        and (pair[0] in REPLACED_COUNTRY_NORM or
             pair[1] in REPLACED_COUNTRY_NORM)
        and pair not in groups
    ]
    if missing_reference:
        sample = ', '.join(
            f'{pair[0]}<->{pair[1]} P{period}: {value:,.3f} MW'
            for pair, period, value in missing_reference[:5])
        raise ValueError(
            f"{sheet_name}: {len(missing_reference)} positive national "
            f"reference row(s) have no NUTS corridor in Sets.xlsx: {sample}")

    typed_rows = []
    reference_rows = iamc_rows = fallback_groups = 0
    for pair, records in sorted(groups.items()):
        for period in periods:
            ref_value = ref_lookup.get((pair, period))
            if ref_value is not None and pair[0] != pair[1]:
                values, fallback = allocate_total(ref_value, records)
                reference_rows += len(records)
                if fallback and ref_value and len(records) > 1:
                    fallback_groups += 1
                    print(f"  [{sheet_name}][WARNING] no IAMC capacity weights for "
                          f"national corridor {pair[0]} <-> {pair[1]} in P{period}; "
                          f"splitting {ref_value:,.3f} MW equally over "
                          f"{len(records)} typed regional corridor(s).")
            else:
                values = [float(r['IAMCCapacity']) * iamc_multiplier
                          for r in records]
                iamc_rows += len(records)

            for record, value in zip(records, values):
                typed_rows.append({
                    'InterconnectorLinks': record['NodeFrom'],
                    'ToNode': record['NodeTo'],
                    'LineType': record['LineType'],
                    'Period': period,
                    value_col: value,
                })

    typed_df = pd.DataFrame(typed_rows)
    aggregate_df = (typed_df.groupby(
        ['InterconnectorLinks', 'ToNode', 'Period'], as_index=False)[value_col]
        .sum())
    print(f"  {len(aggregate_df)} aggregate rows and {len(typed_df)} typed rows "
          f"({reference_rows} allocated from period-specific reference, "
          f"{iamc_rows} from IAMC; {fallback_groups} equal-weight fallback groups)")
    return aggregate_df, typed_df


# =============================================================================
# Sheet builders
# =============================================================================

def make_line_efficiency(dir_lines):
    """lineEfficiency: 0.97 for every directional line."""
    print("[1/8] Creating lineEfficiency...")
    rows = [{'FromNode': r['NodeFrom'], 'ToNode': r['NodeTo'], 'lineEfficiency': 0.97}
            for _, r in dir_lines.iterrows()]
    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} rows")
    return df


def make_max_built_capacity(bi_lines, periods):
    """MaxBuiltCapacity: 20 000 MW for every bidirectional corridor × period."""
    print("[2/8] Creating MaxBuiltCapacity...")
    rows = []
    for _, r in bi_lines.iterrows():
        for p in periods:
            rows.append({
                'InterconnectorLinks': r['NodeFrom'],
                'ToNode': r['NodeTo'],
                'Period': p,
                'TransmissionMaxBuiltCapacity in MW': 20000,
            })
    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} rows ({len(bi_lines)} corridors × {len(periods)} periods)")
    return df


def make_length(bi_lines, reference_trans, coords_dict):
    """Length: reference value if corridor exists there; else haversine; else 0."""
    print("[3/8] Creating Length...")

    ref_ln = pd.read_excel(reference_trans, sheet_name='Length', header=2)
    ref_ln = ref_ln.iloc[:, :3]
    ref_ln.columns = ['FromNode', 'ToNode', 'LengthKm']
    # Remove rows that contain replaced country nodes (old NUTS0 names)
    ref_ln = ref_ln[
        ~ref_ln['FromNode'].str.lower().isin(REPLACED_COUNTRY_NODES_LOWER) &
        ~ref_ln['ToNode'].str.lower().isin(REPLACED_COUNTRY_NODES_LOWER)
    ]
    ref_lookup = build_lookup(ref_ln, 'FromNode', 'ToNode', 'LengthKm')

    rows = []
    from_ref, from_calc, from_zero = 0, 0, 0
    for _, r in bi_lines.iterrows():
        nf, nt = r['NodeFrom'], r['NodeTo']
        key  = (normalise(nf), normalise(nt))
        rkey = (normalise(nt), normalise(nf))   # bidirectional: try both orderings

        if key in ref_lookup:
            length = ref_lookup[key]['LengthKm']
            from_ref += 1
        elif rkey in ref_lookup:
            length = ref_lookup[rkey]['LengthKm']
            from_ref += 1
        elif normalise(nf) in coords_dict and normalise(nt) in coords_dict:
            lat1, lon1 = coords_dict[normalise(nf)]
            lat2, lon2 = coords_dict[normalise(nt)]
            length = round(haversine_km(lat1, lon1, lat2, lon2), 1)
            from_calc += 1
        else:
            length = 0
            from_zero += 1

        rows.append({'FromNode': nf, 'ToNode': nt, 'Length in km': length})

    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} rows  "
          f"({from_ref} from reference, {from_calc} calculated, {from_zero} set to 0)")
    return df


def copy_type_sheet(reference_trans, sheet_name, label):
    """Copy TypeCapitalCost / TypeFixedOMCost as-is (raw, preserves metadata)."""
    print(f"[INFO] Copying {sheet_name} from reference...")
    df = pd.read_excel(reference_trans, sheet_name=sheet_name, header=None)
    print(f"  ✓ {sheet_name}: {df.shape[0]} rows")
    return df


def make_initial_capacity(bi_lines, periods, reference_trans, ac_lookup, dc_lookup,
                          linetype_lookup, landing_lookup, off_norm):
    """InitialCapacity: reference value if present; else inherited offshore landing
    capacity; else IAMC value (AC or DC per line type); else skip."""
    print("[4/8] Creating InitialCapacity...")

    ref_ic = pd.read_excel(reference_trans, sheet_name='InitialCapacity', header=2)
    ref_ic = ref_ic.iloc[:, :4]
    ref_ic.columns = ['InterconnectorLinks', 'ToNode', 'Period', 'TransmissionInitialCapacity']
    # Remove replaced-country rows (replaced by NUTS2)
    ref_ic = ref_ic[
        ~ref_ic['InterconnectorLinks'].str.lower().isin(REPLACED_COUNTRY_NODES_LOWER) &
        ~ref_ic['ToNode'].str.lower().isin(REPLACED_COUNTRY_NODES_LOWER)
    ]
    # Build lookup: (norm_from, norm_to, period) -> capacity
    ref_ic_lookup = {}
    for _, row in ref_ic.iterrows():
        key = (normalise(row['InterconnectorLinks']), normalise(row['ToNode']), int(row['Period']))
        ref_ic_lookup[key] = row['TransmissionInitialCapacity']

    rows = []
    from_ref, from_iamc, from_landing, skipped = 0, 0, 0, 0
    for _, r in bi_lines.iterrows():
        nf, nt = r['NodeFrom'], r['NodeTo']
        knf, knt = normalise(nf), normalise(nt)

        # Select IAMC lookup based on corridor line type (HVAC or HVDC)
        line_type = linetype_lookup.get((knf, knt)) or linetype_lookup.get((knt, knf), 'HVAC_OverheadLine')
        iamc_lookup = dc_lookup if 'HVDC' in str(line_type) else ac_lookup

        for p in periods:
            key1 = (knf, knt, p)
            key2 = (knt, knf, p)
            landing = offshore_landing_value(knf, knt, p, off_norm, landing_lookup)

            if key1 in ref_ic_lookup:
                cap = ref_ic_lookup[key1]
                rows.append({'InterconnectorLinks': nf, 'ToNode': nt,
                             'Period': p, 'TransmissionInitialCapacity': cap})
                from_ref += 1
            elif key2 in ref_ic_lookup:
                cap = ref_ic_lookup[key2]
                rows.append({'InterconnectorLinks': nf, 'ToNode': nt,
                             'Period': p, 'TransmissionInitialCapacity': cap})
                from_ref += 1
            elif landing is not None:
                rows.append({'InterconnectorLinks': nf, 'ToNode': nt,
                             'Period': p, 'TransmissionInitialCapacity': landing})
                from_landing += 1
            elif (knf, knt) in iamc_lookup:
                cap = iamc_lookup[(knf, knt)]
                rows.append({'InterconnectorLinks': nf, 'ToNode': nt,
                             'Period': p, 'TransmissionInitialCapacity': cap})
                from_iamc += 1
            elif (knt, knf) in iamc_lookup:
                cap = iamc_lookup[(knt, knf)]
                rows.append({'InterconnectorLinks': nf, 'ToNode': nt,
                             'Period': p, 'TransmissionInitialCapacity': cap})
                from_iamc += 1
            else:
                skipped += 1  # skip this corridor/period

    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=['InterconnectorLinks', 'ToNode', 'Period', 'TransmissionInitialCapacity'])

    total_corridors = len(bi_lines) * len(periods)
    print(f"  ✓ {len(df)} rows  "
          f"({from_ref} from reference, {from_landing} offshore landings, "
          f"{from_iamc} from IAMC, {skipped}/{total_corridors} skipped)")
    return df


def make_max_install_capacity_raw(bi_lines, periods, reference_trans, ac_lookup, dc_lookup,
                                  linetype_lookup, landing_lookup, off_norm):
    """MaxInstallCapacityRaw: reference if exists; else inherited offshore landing
    capacity; else IAMC×2 (AC or DC per line type); else 0."""
    print("[5/8] Creating MaxInstallCapacityRaw...")

    ref_mr = pd.read_excel(reference_trans, sheet_name='MaxInstallCapacityRaw', header=2)
    ref_mr = ref_mr.iloc[:, :4]
    ref_mr.columns = ['InterconnectorLinks', 'ToNode', 'Period', 'MaxRawNotAdjustWithInitCap']
    ref_mr = ref_mr[
        ~ref_mr['InterconnectorLinks'].str.lower().isin(REPLACED_COUNTRY_NODES_LOWER) &
        ~ref_mr['ToNode'].str.lower().isin(REPLACED_COUNTRY_NODES_LOWER)
    ]
    ref_mr_lookup = {}
    for _, row in ref_mr.iterrows():
        key = (normalise(row['InterconnectorLinks']), normalise(row['ToNode']), int(row['Period']))
        ref_mr_lookup[key] = row['MaxRawNotAdjustWithInitCap']

    rows = []
    from_ref, from_landing, from_iamc, zero_count = 0, 0, 0, 0
    for _, r in bi_lines.iterrows():
        nf, nt = r['NodeFrom'], r['NodeTo']
        knf, knt = normalise(nf), normalise(nt)

        # Select IAMC lookup based on corridor line type (HVAC or HVDC)
        line_type = linetype_lookup.get((knf, knt)) or linetype_lookup.get((knt, knf), 'HVAC_OverheadLine')
        iamc_lookup = dc_lookup if 'HVDC' in str(line_type) else ac_lookup

        for p in periods:
            key1 = (knf, knt, p)
            key2 = (knt, knf, p)
            landing = offshore_landing_value(knf, knt, p, off_norm, landing_lookup)

            if key1 in ref_mr_lookup:
                val = ref_mr_lookup[key1]
                from_ref += 1
            elif key2 in ref_mr_lookup:
                val = ref_mr_lookup[key2]
                from_ref += 1
            elif landing is not None:
                val = landing
                from_landing += 1
            elif (knf, knt) in iamc_lookup:
                val = iamc_lookup[(knf, knt)] * 2
                from_iamc += 1
            elif (knt, knf) in iamc_lookup:
                val = iamc_lookup[(knt, knf)] * 2
                from_iamc += 1
            else:
                val = 0
                zero_count += 1

            rows.append({'InterconnectorLinks': nf, 'ToNode': nt,
                         'Period': p, 'MaxRawNotAdjustWithInitCap in MW': val})

    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} rows  "
          f"({from_ref} from reference, {from_landing} offshore landings, "
          f"{from_iamc} from IAMC×2, {zero_count} set to 0)")
    return df


def make_lifetime(bi_lines):
    """Lifetime: 40 years for every bidirectional corridor."""
    print("[6/8] Creating Lifetime...")
    rows = [{'InterconnectorLinks': r['NodeFrom'], 'To Node': r['NodeTo'],
             'transmissionLifetime': 40}
            for _, r in bi_lines.iterrows()]
    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} rows")
    return df


# =============================================================================
# Load-flow sheet builders
# =============================================================================

def load_iamc_reactance(iamc_file, scenario):
    """Load corridor reactance (Ohm) from the IAMC 'Network|Electricity|Reactance'
    variable. Returns {(norm_from, norm_to): ohm}."""
    print(f"[load-flow] Loading IAMC reactance from {iamc_file}...")
    from iamc_utils import load_iamc
    df = load_iamc(iamc_file).filter(
        scenario=scenario, variable=_REACTANCE_VAR).as_pandas()
    lookup = {}
    if not df.empty:
        df = df[df['region'].str.contains('>', na=False)].copy()
        for _, row in df.iterrows():
            parts = row['region'].split('>')
            nf = resolve_iamc_endpoint(parts[0].strip())
            nt = resolve_iamc_endpoint(parts[1].strip())
            lookup[(normalise(nf), normalise(nt))] = float(row['value'])
    print(f"  ✓ {len(lookup)} reactance records ('{_REACTANCE_VAR}')")
    return lookup


def build_length_lookup(len_df):
    """Return {(norm_from, norm_to): length_km} (both orientations) from Length."""
    lookup = {}
    for _, r in len_df.iterrows():
        length = r['Length in km']
        lookup[(normalise(r['FromNode']), normalise(r['ToNode']))] = length
        lookup[(normalise(r['ToNode']), normalise(r['FromNode']))] = length
    return lookup


def make_line_reactance(bi_lines, length_lookup, linetype_lookup, reactance_lookup, per_km):
    """lineReactance (Ohm) for AC corridors only. Uses the IAMC reactance when
    available, otherwise length_km * reactance_per_km. HVDC corridors are
    excluded (they do not obey Ohm's law)."""
    print("[load-flow] Creating lineReactance...")
    rows = []
    from_iamc, from_len = 0, 0
    for _, r in bi_lines.iterrows():
        nf, nt = r['NodeFrom'], r['NodeTo']
        knf, knt = normalise(nf), normalise(nt)
        line_types = line_types_for_pair(linetype_lookup, knf, knt)
        if not any('HVAC' in str(line_type) for line_type in line_types):
            continue
        if (knf, knt) in reactance_lookup:
            x = reactance_lookup[(knf, knt)]
            from_iamc += 1
        elif (knt, knf) in reactance_lookup:
            x = reactance_lookup[(knt, knf)]
            from_iamc += 1
        else:
            length = length_lookup.get((knf, knt), 0) or 0
            x = round(length * per_km, 4)
            from_len += 1
        rows.append({'FromNode': nf, 'ToNode': nt, 'lineReactance': x})
    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} AC corridors ({from_iamc} from IAMC, {from_len} from length×{per_km})")
    return df


def make_candidate_transmission(bi_lines):
    """CandidateTransmission: every bidirectional corridor is an expansion
    candidate (orientation matches BidirectionalLines / InitialCapacity)."""
    print("[load-flow] Creating CandidateTransmission...")
    df = pd.DataFrame([{'NodeFrom': r['NodeFrom'], 'NodeTo': r['NodeTo']}
                       for _, r in bi_lines.iterrows()])
    print(f"  ✓ {len(df)} candidate corridors")
    return df


def make_line_block_capacity(bi_lines, block_mw):
    """LineBlockCapacity: constant block size (MW) per candidate."""
    print("[load-flow] Creating LineBlockCapacity...")
    df = pd.DataFrame([{'NodeFrom': r['NodeFrom'], 'NodeTo': r['NodeTo'],
                        'LineBlockCapacity': block_mw}
                       for _, r in bi_lines.iterrows()])
    print(f"  ✓ {len(df)} rows ({block_mw} MW each)")
    return df


def make_line_block_reactance(bi_lines, length_lookup, per_km):
    """LineBlockReactance (Ohm): length_km * reactance_per_km per candidate."""
    print("[load-flow] Creating LineBlockReactance...")
    rows = []
    for _, r in bi_lines.iterrows():
        knf, knt = normalise(r['NodeFrom']), normalise(r['NodeTo'])
        length = length_lookup.get((knf, knt), 0) or 0
        rows.append({'NodeFrom': r['NodeFrom'], 'NodeTo': r['NodeTo'],
                     'LineBlockReactance': round(length * per_km, 4)})
    df = pd.DataFrame(rows)
    print(f"  ✓ {len(df)} rows (length×{per_km})")
    return df


def enforce_expansion_headroom(mr_df, ic_df, block_mw):
    """Raise MaxInstallCapacityRaw so every corridor/period has room for at least
    one candidate block: MaxInstallCapacityRaw >= InitialCapacity + block_mw."""
    print("[load-flow] Enforcing MaxInstallCapacityRaw expansion headroom...")
    ic_lookup = {}
    for _, r in ic_df.iterrows():
        ic_lookup[(normalise(r['InterconnectorLinks']), normalise(r['ToNode']),
                   int(r['Period']))] = r['TransmissionInitialCapacity']
    col = 'MaxRawNotAdjustWithInitCap in MW'
    out = mr_df.copy()
    bumped, vals = 0, []
    for _, r in out.iterrows():
        key = (normalise(r['InterconnectorLinks']), normalise(r['ToNode']), int(r['Period']))
        ic = ic_lookup.get(key, 0) or 0
        needed = ic + block_mw
        v = r[col]
        if pd.isna(v) or v < needed:
            v = needed
            bumped += 1
        vals.append(v)
    out[col] = vals
    print(f"  ✓ {bumped} rows raised to Initial + {block_mw}")
    return out


# =============================================================================
# Main
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Generate Transmission.xlsx for EMPIRE (NUTS2 selected countries)'
    )
    parser.add_argument('iamc_file', help='Path to IAMC data file (CSV or Excel)')
    parser.add_argument('scenario', help='Scenario name')
    parser.add_argument('sets_file', help='Sets_NUTS2_DK_DE.xlsx (generated by generate_sets_nuts2.py)')
    parser.add_argument('output_file', help='Output Transmission.xlsx path')
    parser.add_argument('--reference', default='data/EMPIRE/v_52/Transmission.xlsx',
                        help='Reference Transmission.xlsx (default: data/EMPIRE/v_52/Transmission.xlsx)')
    args = parser.parse_args()

    print("=" * 80)
    print("TRANSMISSION.XLSX GENERATOR — NUTS2 for selected countries")
    print("=" * 80)
    print(f"IAMC file:  {args.iamc_file}")
    print(f"Scenario:   {args.scenario}")
    print(f"Sets file:  {args.sets_file}")
    print(f"Reference:  {args.reference}")
    print(f"Output:     {args.output_file}")
    print(f"NUTS2 countries: {', '.join(NUTS2_COUNTRY_PREFIXES)}")
    print()

    # --- Load inputs ---
    dir_lines, bi_lines, periods, coords_dict, linetype_lookup = load_sets(args.sets_file)
    ac_lookup, dc_lookup = load_iamc_capacity(args.iamc_file, args.scenario)

    # --- Build sheets ---
    eff_df   = make_line_efficiency(dir_lines)
    mbc_df   = make_max_built_capacity(bi_lines, periods)
    len_df   = make_length(bi_lines, args.reference, coords_dict)
    tc_raw   = copy_type_sheet(args.reference, 'TypeCapitalCost', 'TypeCapitalCost')
    tfc_raw  = copy_type_sheet(args.reference, 'TypeFixedOMCost', 'TypeFixedOMCost')
    print("[4/8] Creating InitialCapacity (national reference totals, IAMC weights)...")
    ic_df, ic_typed_df = make_capacity_sheets(
        bi_lines, periods, args.reference, 'InitialCapacity',
        'TransmissionInitialCapacity', ac_lookup, dc_lookup,
        linetype_lookup, iamc_multiplier=1.0)
    print("[5/8] Creating MaxInstallCapacityRaw (national reference totals, IAMC weights)...")
    mr_df, mr_typed_df = make_capacity_sheets(
        bi_lines, periods, args.reference, 'MaxInstallCapacityRaw',
        'MaxRawNotAdjustWithInitCap in MW', ac_lookup, dc_lookup,
        linetype_lookup, iamc_multiplier=2.0)
    lt_df    = make_lifetime(bi_lines)

    # --- Load-flow sheets (optional) ---
    lf_lr_df = lf_cand_df = lf_lbc_df = lf_lbr_df = None
    if _LF['enabled']:
        per_km = _LF['reactance_per_km']
        block_mw = _LF['block_capacity_mw']
        length_lookup = build_length_lookup(len_df)
        reactance_lookup = load_iamc_reactance(args.iamc_file, args.scenario)
        lf_lr_df   = make_line_reactance(bi_lines, length_lookup, linetype_lookup,
                                         reactance_lookup, per_km)
        lf_cand_df = make_candidate_transmission(bi_lines)
        lf_lbc_df  = make_line_block_capacity(bi_lines, block_mw)
        lf_lbr_df  = make_line_block_reactance(bi_lines, length_lookup, per_km)
        mr_df      = enforce_expansion_headroom(mr_df, ic_df, block_mw)

    # --- Write Excel ---
    print(f"\n[WRITE] Writing {args.output_file}...")
    with pd.ExcelWriter(args.output_file, engine='openpyxl') as writer:

        # lineEfficiency — metadata
        add_metadata_header(eff_df,
            source='Assumed',
            description='Efficiency of transmission exchange (default: 0.97)'
        ).to_excel(writer, sheet_name='lineEfficiency', index=False, header=False)

        # MaxBuiltCapacity — metadata
        add_metadata_header(mbc_df,
            source='Assumed',
            description='Maximum capacity expansion per corridor per period (default: 20 000 MW)'
        ).to_excel(writer, sheet_name='MaxBuiltCapacity', index=False, header=False)

        # Length — metadata
        add_metadata_header(len_df,
            source=f'{args.reference} + haversine from coords',
            description='Assumed length of transmission corridors in km'
        ).to_excel(writer, sheet_name='Length', index=False, header=False)

        # TypeCapitalCost — raw copy
        tc_raw.to_excel(writer, sheet_name='TypeCapitalCost', index=False, header=False)

        # TypeFixedOMCost — raw copy
        tfc_raw.to_excel(writer, sheet_name='TypeFixedOMCost', index=False, header=False)

        # InitialCapacity — metadata
        add_metadata_header(ic_df,
            source=f'{args.reference} + {args.iamc_file}',
            description='Aggregate HVAC+HVDC initial capacity; national reference totals distributed by IAMC weights (MW)'
        ).to_excel(writer, sheet_name='InitialCapacity', index=False, header=False)

        add_metadata_header(ic_typed_df,
            source=f'{args.reference} + {args.iamc_file}',
            description='Initial capacity by corridor, line type and period; use for parallel HVAC/HVDC (MW)'
        ).to_excel(writer, sheet_name='InitialCapacityByType', index=False, header=False)

        # MaxInstallCapacityRaw — metadata
        add_metadata_header(mr_df,
            source=f'{args.reference} + {args.iamc_file}',
            description='Aggregate HVAC+HVDC maximum installed capacity; national reference totals distributed by IAMC weights (MW)'
        ).to_excel(writer, sheet_name='MaxInstallCapacityRaw', index=False, header=False)

        add_metadata_header(mr_typed_df,
            source=f'{args.reference} + {args.iamc_file}',
            description='Maximum installed capacity by corridor, line type and period; use for parallel HVAC/HVDC (MW)'
        ).to_excel(writer, sheet_name='MaxInstallCapacityRawByType', index=False, header=False)

        # Lifetime — metadata
        add_metadata_header(lt_df,
            source='Assumed',
            description='Lifetime of transmission lines in years (default: 40)'
        ).to_excel(writer, sheet_name='Lifetime', index=False, header=False)

        # Load-flow sheets (optional)
        if _LF['enabled']:
            add_metadata_header(lf_lr_df,
                source=f'{args.iamc_file} (Network|Electricity|Reactance) + length×{_LF["reactance_per_km"]}',
                description='Reactance of AC transmission corridors (Ohm)'
            ).to_excel(writer, sheet_name='lineReactance', index=False, header=False)

            add_metadata_header(lf_cand_df,
                source='Existing corridors + adjacency',
                description='Candidate lines for transmission expansion planning'
            ).to_excel(writer, sheet_name='CandidateTransmission', index=False, header=False)

            add_metadata_header(lf_lbc_df,
                source='Assumed',
                description='Candidate lines capacity for transmission expansion planning. MW'
            ).to_excel(writer, sheet_name='LineBlockCapacity', index=False, header=False)

            add_metadata_header(lf_lbr_df,
                source=f'length×{_LF["reactance_per_km"]}',
                description='Candidate lines reactance for transmission expansion planning. Ohm'
            ).to_excel(writer, sheet_name='LineBlockReactance', index=False, header=False)

    # --- Summary ---
    print()
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Output: {args.output_file}")
    print()
    print(f"  lineEfficiency:          {len(eff_df)} directional lines")
    print(f"  MaxBuiltCapacity:        {len(mbc_df)} rows "
          f"({len(bi_lines)} corridors × {len(periods)} periods)")
    print(f"  Length:                  {len(len_df)} corridors")
    print(f"  TypeCapitalCost:         (copied from reference)")
    print(f"  TypeFixedOMCost:         (copied from reference)")
    print(f"  InitialCapacity:         {len(ic_df)} rows")
    print(f"  MaxInstallCapacityRaw:   {len(mr_df)} rows "
          f"({len(bi_lines)} corridors × {len(periods)} periods)")
    print(f"  Lifetime:                {len(lt_df)} corridors")
    if _LF['enabled']:
        print(f"  lineReactance:           {len(lf_lr_df)} AC corridors")
        print(f"  CandidateTransmission:   {len(lf_cand_df)} corridors")
        print(f"  LineBlockCapacity:       {len(lf_lbc_df)} rows ({_LF['block_capacity_mw']} MW)")
        print(f"  LineBlockReactance:      {len(lf_lbr_df)} rows")
    print()
    print("✓ Done!")
    print("=" * 80)


if __name__ == '__main__':
    if len(sys.argv) < 5:
        print("Usage: python generate_transmission_nuts2.py <iamc_file> <scenario> "
              "<sets_file> <output_file> [--reference <Transmission.xlsx>]")
        print()
        print("Example:")
        print('  python generate_transmission_nuts2.py '
              'data/IAMC/NorthSea/Denmark_Germany.csv '
              '"Transmission v1.0" Sets_NUTS2_DK_DE.xlsx Transmission_NUTS2_DK_DE.xlsx')
        sys.exit(1)
    main()
