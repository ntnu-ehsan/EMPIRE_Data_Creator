"""
Generate Node.xlsx for EMPIRE from generated Sets.xlsx and IAMC data.

Usage:
    python generate_node.py <sets_file> <iamc_file> <scenario> <output_file>
                            [--reference <existing_node.xlsx>]

Example:
    python generate_node.py Sets_NUTS2_DK_DE.xlsx \\
        data/IAMC/NorthSea/Denmark_Germany.csv \\
        "Transmission v1.0" Node_NUTS2_DK_DE.xlsx

Sheets generated
----------------
ElectricAnnualDemand
    Annual electricity demand per node per investment period (MWh/h).
    External nodes: copied from reference.
    DE / DK NUTS2: country reference value for each period is distributed
    across NUTS2 regions using the ``Final Energy|Electricity [Share]``
    variable from the IAMC file (values are percentages, applied uniformly
    across all periods because the IAMC file contains only a single year).

NodeLostLoadCost
    Cost of unserved energy (EUR/MWh) per node.
    External nodes: copied from reference.
    DE / DK NUTS2: the Germany / Denmark reference value is assigned
    identically to every NUTS2 region in that country.

HydroGenMaxAnnualProduction
    Maximum annual non-pumped hydro production (MWh/yr) per node. EMPIRE's
    node-level constraint covers both regulated and run-of-river hydro.
    External nodes: copied from reference.
    NUTS regions: the country reference value is distributed proportionally
    to total non-pumped IAMC hydro capacity (Reservoir + Run of River + bare
    Hydro residual). Falls back to an even split when none is available.

BiomassMaxAnnualActivity (optional)
    Preserved whenever present in the reference workbook. Nodes that stay whole
    keep their reference row unchanged. A NUTS-disaggregated country emits no
    nodal rows at all — its annual biomass-production limit is national and
    moves to BiomassMaxAnnualActivityCountry.

BiomassMaxAnnualActivityCountry (optional)
    National counterpart of the above, one row per (Country, Period) with the
    reference country value unchanged. Enforced over the sum of the country's
    nodes (Sets.xlsx NodesOfCountry), like the *Country sheets in Generator.xlsx.
"""

import sys
import argparse
import pandas as pd
import numpy as np

try:
    import pyam
except ImportError:
    print("ERROR: pyam-iamc not installed. Run: pip install pyam-iamc")
    sys.exit(1)


# =============================================================================
# Constants — loaded from config.yaml in the project root.
# Edit config.yaml to change which countries are represented at NUTS2 level.
# =============================================================================
from config_loader import load_nuts2_config as _load_nuts2_config
_cfg = _load_nuts2_config()
NUTS2_COUNTRY_PREFIXES = _cfg['NUTS2_COUNTRY_PREFIXES']
NUTS2_COUNTRY_TO_EMPIRE_NODE = _cfg['NUTS2_COUNTRY_TO_EMPIRE_NODE']

IAMC_VAR_DEMAND_SHARE = 'Final Energy|Electricity [Share]'


# =============================================================================
# Helpers
# =============================================================================


def add_metadata_header(df, source='-', description=''):
    """Prepend two metadata rows + a column-header row (EMPIRE format)."""
    cols = df.columns.tolist()
    n = len(cols)
    rows = [
        [f'Source: {source}'] + [np.nan] * (n - 1),
        [f'Description: {description}'] + [np.nan] * (n - 1),
        cols,
    ]
    for _, row in df.iterrows():
        rows.append(row.tolist())
    return pd.DataFrame(rows, columns=[f'col_{i}' for i in range(n)])


def read_sets_sheet_simple(sets_file, sheet_name):
    """Read a Sets.xlsx sheet with a single header row."""
    return pd.read_excel(sets_file, sheet_name=sheet_name)


def classify_nodes(all_nodes):
    """Split node list into ({prefix: [nuts2_nodes]}, external_list)."""
    nodes_by_prefix = {
        prefix: [n for n in all_nodes if str(n).startswith(prefix)]
        for prefix in NUTS2_COUNTRY_PREFIXES
    }
    nuts2_set = set(n for nodes in nodes_by_prefix.values() for n in nodes)
    external = [n for n in all_nodes if n not in nuts2_set]
    return nodes_by_prefix, external


def load_iamc_single_value(iamc_file, scenario, variable):
    """Return {region: value} for a IAMC variable (single year assumed)."""
    from iamc_utils import load_iamc
    df = load_iamc(iamc_file)
    flt = df.filter(scenario=scenario, variable=variable)
    if len(flt.data) == 0:
        print(f'  WARNING: no IAMC data for variable={variable!r}')
        return {}
    lookup = {}
    for _, r in flt.data.iterrows():
        region = r['region']
        if '>' not in region:
            lookup[region] = r['value']
    return lookup


# =============================================================================
# Node name normalisation
#
# Node.xlsx is inconsistent with itself: NodeLostLoadCost and
# HydroGenMaxAnnualProduction spell the affected nodes 'Great Brit.' / 'Czech R'
# / 'Bosnia H', while ElectricAnnualDemand and BiomassMaxAnnualActivity use the
# closed-up 'GreatBrit.' / 'CzechR' / 'BosniaH'.  generate_sets_nuts2.py
# normalises the Nodes set to the spaced form, so every sheet here has to agree
# or it silently detaches from the node it describes: biomass rows filtered
# against the node list vanish outright, and demand rows survive under a name no
# node carries.  Same map as generate_sets_nuts2.py / generate_generation_nuts2.py.
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
    'Sorlige NordsjoI':  'Sorlige Nordsjo I',
    'Sorlige NordsjoII': 'Sorlige Nordsjo II',
    'UtsiraNord':       'Utsira Nord',
    'BosniaH':          'Bosnia H',
    'CzechR':           'Czech R',
    'GreatBrit.':       'Great Brit.',
    'IE01':             'Ireland',
    'IE02':             'Ireland',
}


def fix_node_name(name):
    """Return the canonical node name, correcting legacy no-space variants."""
    if isinstance(name, str):
        return NODE_NAME_FIXES.get(name.strip(), name)
    return name


def read_ref_sheet_raw(ref_file, sheet_name):
    """Return (header_cols, data_df) for a reference Node.xlsx sheet.

    Row 0 = source, row 1 = description, row 2 = column headers, rows 3+ = data.
    The node column is normalised here so every caller sees the same spelling
    the Nodes set uses.
    """
    raw = pd.read_excel(ref_file, sheet_name=sheet_name, header=None)
    header_cols = [str(c) if pd.notna(c) else '' for c in raw.iloc[2].tolist()]
    data = raw.iloc[3:].copy()
    data.columns = list(range(data.shape[1]))
    data = data.reset_index(drop=True)
    data[0] = data[0].map(fix_node_name)
    return header_cols, data


# =============================================================================
# Sheet builders
# =============================================================================


def make_electric_annual_demand(all_nodes, nodes_by_prefix,
                                periods, demand_shares,
                                ref_file):
    """ElectricAnnualDemand: Nodes x Period -> MWh/h.

    Uses ``Final Energy|Electricity [Share]`` values (%) from the IAMC file.
    The same regional shares are applied to every investment period because
    the IAMC file contains only a single reference year.
    """
    COL = 'ElectricAdjustment in MWh per hour'

    header_cols, ref_data = read_ref_sheet_raw(ref_file, 'ElectricAnnualDemand')
    # Positional: col0=Nodes, col1=Period, col2=ElectricAdjustment
    ref_data.columns = list(range(ref_data.shape[1]))

    # Build reference lookup: (country_name, period) -> value
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        node = str(r.iloc[0])
        try:
            period = int(float(str(r.iloc[1])))
        except (ValueError, TypeError):
            continue
        try:
            val = float(r.iloc[2])
        except (ValueError, TypeError):
            continue
        ref_lookup[(node, period)] = val

    nuts2_set = set(n for nodes in nodes_by_prefix.values() for n in nodes)
    replaced_names = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())
    period_set = set(periods)

    # Per-country share lookups (node -> fraction summing to 1.0)
    country_shares = {
        prefix: _compute_shares(
            nuts2_nodes, demand_shares,
            label=f'{prefix} Final Energy|Electricity [Share]')
        for prefix, nuts2_nodes in nodes_by_prefix.items()
    }

    rows = []

    # External nodes: copy from reference for the periods in Horizon
    for _, r in ref_data.iterrows():
        node = str(r.iloc[0])
        if node in nuts2_set or node in replaced_names:
            continue
        try:
            period = int(float(str(r.iloc[1])))
        except (ValueError, TypeError):
            continue
        if period not in period_set:
            continue
        try:
            val = float(r.iloc[2])
        except (ValueError, TypeError):
            val = 0.0
        rows.append({'Nodes': node, 'Period': period, COL: round(val, 6)})

    # NUTS2 nodes — one loop per configured country
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        shares = country_shares[prefix]
        for node in nuts2_nodes:
            share = shares.get(node, 0.0)
            for p in periods:
                country_val = ref_lookup.get((country_name, p), 0.0)
                rows.append({
                    'Nodes': node,
                    'Period': p,
                    COL: round(share * country_val, 6),
                })

    return pd.DataFrame(rows)


def make_node_lost_load_cost(all_nodes, nodes_by_prefix, ref_file):
    """NodeLostLoadCost: Nodes x Period -> EUR/MWh.

    One row per node (period 1 only, matching reference convention).
    NUTS2 nodes receive the same single-period value as their country.
    """
    COL = 'NodeLostLoadCost'

    header_cols, ref_data = read_ref_sheet_raw(ref_file, 'NodeLostLoadCost')

    # Positional: col0=Nodes, col1=Period, col2=NodeLostLoadCost
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        node = str(r.iloc[0])
        try:
            period = int(float(str(r.iloc[1])))
        except (ValueError, TypeError):
            continue
        try:
            val = float(r.iloc[2])
        except (ValueError, TypeError):
            val = 22000.0
        ref_lookup[(node, period)] = val

    nuts2_set = set(n for nodes in nodes_by_prefix.values() for n in nodes)
    replaced_names = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())

    rows = []

    # External nodes: copy as-is
    for _, r in ref_data.iterrows():
        node = str(r.iloc[0])
        if node in nuts2_set or node in replaced_names:
            continue
        try:
            period = int(float(str(r.iloc[1])))
            val = float(r.iloc[2])
        except (ValueError, TypeError):
            continue
        rows.append({'Nodes': node, 'Period': period, COL: val})

    # NUTS2 nodes — use the country's reference value (period 1)
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        cost = ref_lookup.get((country_name, 1), 22000.0)
        for node in nuts2_nodes:
            rows.append({'Nodes': node, 'Period': 1, COL: cost})

    return pd.DataFrame(rows)


def make_hydro_gen_max_annual_production(all_nodes, nodes_by_prefix,
                                         hydro_caps, ref_file):
    """HydroGenMaxAnnualProduction: Nodes -> MWh/yr  (no Period dimension).

    NUTS nodes: country reference value distributed proportionally to each
    region's total non-pumped hydro capacity from the IAMC file. Falls back
    to an even split when no regional hydro data are available.
    """
    COL = 'HydroGenMaxAnnualProduction in MWh per year'

    header_cols, ref_data = read_ref_sheet_raw(
        ref_file, 'HydroGenMaxAnnualProduction')

    # Positional: col0=Nodes, col1=HydroGenMaxAnnualProduction
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        node = str(r.iloc[0])
        try:
            val = float(r.iloc[1])
        except (ValueError, TypeError):
            val = 0.0
        ref_lookup[node] = val

    nuts2_set = set(n for nodes in nodes_by_prefix.values() for n in nodes)
    replaced_names = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())

    # Per-country hydro share lookups
    country_shares = {
        prefix: _compute_shares(
            nuts2_nodes, hydro_caps,
            label=f'{prefix} total non-pumped hydro capacity')
        for prefix, nuts2_nodes in nodes_by_prefix.items()
    }

    rows = []

    # External nodes: copy as-is
    for _, r in ref_data.iterrows():
        node = str(r.iloc[0])
        if node in nuts2_set or node in replaced_names:
            continue
        try:
            val = float(r.iloc[1])
        except (ValueError, TypeError):
            val = 0.0
        rows.append({'Nodes': node, COL: round(val, 3)})

    # NUTS2 nodes — one loop per configured country
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        country_total = ref_lookup.get(country_name, 0.0)
        shares = country_shares[prefix]
        for node in nuts2_nodes:
            rows.append({
                'Nodes': node,
                COL: round(shares.get(node, 0.0) * country_total, 3),
            })

    return pd.DataFrame(rows)


# =============================================================================
# Optional reference sheets
# =============================================================================


def _read_biomass_reference(ref_file, periods):
    '''Return (raw_sheet, tidy_data) for the reference BiomassMaxAnnualActivity.

    ``tidy_data`` is filtered to the investment periods of this run; ``raw_sheet``
    is kept so the caller can reuse the original metadata rows.
    '''
    raw = pd.read_excel(
        ref_file, sheet_name='BiomassMaxAnnualActivity', header=None
    )
    if raw.shape[0] < 3 or raw.shape[1] < 3:
        raise ValueError(
            'BiomassMaxAnnualActivity must contain two metadata rows, a header '
            'row, and three columns'
        )

    data = raw.iloc[3:, :3].copy()
    data.columns = ['Node', 'Period', 'maxBiomassNode']
    data['Node'] = data['Node'].map(fix_node_name)
    data['Period'] = pd.to_numeric(data['Period'], errors='coerce')
    data['maxBiomassNode'] = pd.to_numeric(
        data['maxBiomassNode'], errors='coerce'
    )
    data = data.dropna(subset=['Node', 'Period', 'maxBiomassNode'])
    return raw, data[data['Period'].isin(set(periods))].copy()


def make_biomass_max_annual_activity(all_nodes, periods, ref_file):
    '''BiomassMaxAnnualActivity (nodal) — pass-through for nodes that stay whole.

    The limit caps annual biomass *production* in a node.  For a country split
    into NUTS regions the cap is national: the biomass resource is traded within
    the country, so neither replicating the national total to every region (each
    region could then burn the whole national budget) nor splitting it by
    installed capacity (which strands the cap in the few regions the IAMC file
    happens to cover) reflects the constraint.  Those countries' rows move to
    ``BiomassMaxAnnualActivityCountry`` instead — the same treatment
    MaxBuiltCapacity / MinBuiltCapacity get in generate_generation_nuts2.py.
    '''
    raw, data = _read_biomass_reference(ref_file, periods)

    country_nodes = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())
    output_data = data[
        data['Node'].isin(set(all_nodes)) & ~data['Node'].isin(country_nodes)
    ].copy()
    output_data['Period'] = output_data['Period'].astype(int)

    metadata = raw.iloc[:3, :3].copy()
    metadata.columns = range(3)
    output_data.columns = range(3)
    return pd.concat([metadata, output_data], ignore_index=True)


def make_biomass_max_annual_activity_country(nodes_by_prefix, periods,
                                             ref_file):
    '''BiomassMaxAnnualActivityCountry — national biomass production limit.

    One row per (Country, Period), value taken unchanged from the reference
    workbook's country-level row, for every NUTS-disaggregated country.  The
    constraint is meant to be enforced over the sum of the country's NUTS nodes
    (``NodesOfCountry`` in Sets.xlsx); a missing country means unconstrained.
    '''
    _, data = _read_biomass_reference(ref_file, periods)

    rows = []
    for prefix, regions in nodes_by_prefix.items():
        country_node = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        country_rows = data[data['Node'] == country_node]
        if country_rows.empty:
            print(
                f'  [BiomassMaxAnnualActivityCountry][WARNING] {country_node}: '
                'no reference row; its biomass production stays unconstrained.'
            )
            continue
        if not regions:
            print(
                f'  [BiomassMaxAnnualActivityCountry][WARNING] {country_node}: '
                'no NUTS regions in Sets; dropping its biomass-limit rows.'
            )
            continue
        for _, row in country_rows.iterrows():
            rows.append({
                'Country': country_node,
                'Period': int(row['Period']),
                'maxBiomassCountry': row['maxBiomassNode'],
            })

    return pd.DataFrame(
        rows, columns=['Country', 'Period', 'maxBiomassCountry']
    )


# =============================================================================
# Internal helpers
# =============================================================================


def _compute_shares(nodes, value_lookup, label=''):
    """Compute fractional shares from a {node: raw_value} lookup.

    If the value is already a percentage (%), converts to fraction.
    Falls back to equal weighting if no data exist for any node.
    """
    raw = {n: value_lookup.get(n, 0.0) for n in nodes}
    total = sum(raw.values())
    if total == 0:
        n = len(nodes)
        print(f'  WARNING: no {label} data found for nodes '
              f'{nodes[:3]}... — falling back to equal weights.')
        return {n_: 1.0 / len(nodes) if len(nodes) > 0 else 0.0
                for n_ in nodes}
    return {n_: v / total for n_, v in raw.items()}


# =============================================================================
# Main
# =============================================================================


def main():
    parser = argparse.ArgumentParser(
        description='Generate Node.xlsx for EMPIRE')
    parser.add_argument('sets_file', help='Generated Sets.xlsx')
    parser.add_argument('iamc_file', help='IAMC data file (CSV / Excel)')
    parser.add_argument('scenario', help='IAMC scenario name')
    parser.add_argument('output_file', help='Output Node.xlsx path')
    parser.add_argument('--reference',
                        default='data/EMPIRE/v_52/Node.xlsx',
                        help='Reference Node.xlsx '
                             '(default: data/EMPIRE/v_52/Node.xlsx)')
    args = parser.parse_args()

    print('=' * 80)
    print('NODE.XLSX GENERATOR FOR EMPIRE')
    print('=' * 80)
    print(f'Sets:      {args.sets_file}')
    print(f'IAMC:      {args.iamc_file}')
    print(f'Scenario:  {args.scenario}')
    print(f'Reference: {args.reference}')
    print(f'Output:    {args.output_file}')
    print()

    # ------------------------------------------------------------------
    # Read Sets.xlsx
    # ------------------------------------------------------------------
    all_nodes = read_sets_sheet_simple(args.sets_file, 'Nodes')['Node'].tolist()
    periods = read_sets_sheet_simple(args.sets_file, 'Horizon')['Horizon'].tolist()
    nodes_by_prefix, ext_nodes = classify_nodes(all_nodes)

    print(f'Total nodes:    {len(all_nodes)}')
    for prefix, nodes in nodes_by_prefix.items():
        print(f'{prefix} NUTS2 nodes: {len(nodes)}')
    print(f'External nodes: {len(ext_nodes)}')
    print(f'Periods:        {periods}')
    print()

    # ------------------------------------------------------------------
    # Load IAMC data
    # ------------------------------------------------------------------
    print('Loading IAMC data...')
    demand_shares = load_iamc_single_value(
        args.iamc_file, args.scenario, IAMC_VAR_DEMAND_SHARE)
    from iamc_utils import load_hydro_capacity_weights
    hydro_caps, hydro_sources = load_hydro_capacity_weights(
        args.iamc_file, args.scenario, nodes_by_prefix, 'annual_total')

    reference_sheets = pd.ExcelFile(args.reference).sheet_names
    has_biomass_limit = 'BiomassMaxAnnualActivity' in reference_sheets

    for prefix, nuts2_nodes in nodes_by_prefix.items():
        dc = sum(1 for n in nuts2_nodes if n in demand_shares)
        hc = sum(1 for n in nuts2_nodes if hydro_caps.get(n, 0.0) > 0)
        print(f'  {prefix}: {dc}/{len(nuts2_nodes)} nodes have demand share, '
              f'{hc}/{len(nuts2_nodes)} have hydro capacity '
              f'[{hydro_sources[prefix]}]')
    print()

    # ------------------------------------------------------------------
    # Build sheets
    # ------------------------------------------------------------------
    sheet_total = 5 if has_biomass_limit else 3
    print(f'[1/{sheet_total}] ElectricAnnualDemand '
          '(external from reference; NUTS2 weighted by IAMC share)...')
    ead_df = make_electric_annual_demand(
        all_nodes, nodes_by_prefix, periods, demand_shares, args.reference)
    print(f'  -> {len(ead_df)} rows '
          f'({len(ead_df["Nodes"].unique())} nodes x ~{len(periods)} periods)')

    print(f'[2/{sheet_total}] NodeLostLoadCost '
          '(external from reference; NUTS2 use country value)...')
    nllc_df = make_node_lost_load_cost(
        all_nodes, nodes_by_prefix, args.reference)
    print(f'  -> {len(nllc_df)} rows')

    print(f'[3/{sheet_total}] HydroGenMaxAnnualProduction '
          '(NUTS: total non-pumped hydro weighted; even split if absent)...')
    hgmap_df = make_hydro_gen_max_annual_production(
        all_nodes, nodes_by_prefix, hydro_caps, args.reference)
    print(f'  -> {len(hgmap_df)} rows')

    biomass_raw = None
    biomass_country_df = None
    if has_biomass_limit:
        print(
            f'[4/{sheet_total}] BiomassMaxAnnualActivity '
            '(nodal; NUTS-disaggregated countries move to the national sheet)...'
        )
        biomass_raw = make_biomass_max_annual_activity(
            all_nodes, periods, args.reference,
        )
        print(f'  -> {len(biomass_raw) - 3} rows')

        print(
            f'[5/{sheet_total}] BiomassMaxAnnualActivityCountry '
            '(national limit, reference country value unchanged)...'
        )
        biomass_country_df = make_biomass_max_annual_activity_country(
            nodes_by_prefix, periods, args.reference,
        )
        print(f'  -> {len(biomass_country_df)} rows')
    else:
        # Non-fatal: some reference datasets genuinely lack the sheet.  Labelled
        # WARNING so it surfaces in a log scan — an upstream step silently
        # dropping the sheet looks identical to it never being there, and the
        # country's biomass production then goes unconstrained.
        print(
            '  WARNING: optional BiomassMaxAnnualActivity sheet not present in '
            'reference; skipping. Biomass production will be unconstrained — '
            'check this is expected for this dataset.'
        )
    print()

    # ------------------------------------------------------------------
    # Add metadata headers
    # ------------------------------------------------------------------
    ead_meta = add_metadata_header(
        ead_df,
        source=f'{args.iamc_file} + reference',
        description='Annual demand in nodes used to adjust hourly load profile '
                    '(such that the integral/sum equals the annual demand) '
                    '(default: 0)')

    nllc_meta = add_metadata_header(
        nllc_df,
        source='Reference + assumed',
        description='The cost of not generating electricity in an hour '
                    '(default: 22 000)')

    hgmap_meta = add_metadata_header(
        hgmap_df,
        source=f'{args.iamc_file} + reference',
        description='HydroGenerator With Reservoir Maximum Expected Annual '
                    'Production (MWh/yr)')

    biomass_country_meta = None
    if biomass_country_df is not None:
        biomass_country_meta = add_metadata_header(
            biomass_country_df,
            source='Reference Node.xlsx (country-level row)',
            description='Annual electricity production from biomass (Bio, '
                        'BioCCS) summed over all NUTS nodes of a country, per '
                        'period. National counterpart of '
                        'BiomassMaxAnnualActivity for NUTS-disaggregated '
                        'countries (only listed countries constrained)')

    # ------------------------------------------------------------------
    # Write output
    # ------------------------------------------------------------------
    print(f'Writing {args.output_file}...')
    with pd.ExcelWriter(args.output_file, engine='openpyxl') as writer:
        ead_meta.to_excel(writer, sheet_name='ElectricAnnualDemand',
                          index=False, header=False)
        nllc_meta.to_excel(writer, sheet_name='NodeLostLoadCost',
                           index=False, header=False)
        hgmap_meta.to_excel(writer, sheet_name='HydroGenMaxAnnualProduction',
                            index=False, header=False)
        if biomass_raw is not None:
            biomass_raw.to_excel(
                writer,
                sheet_name='BiomassMaxAnnualActivity',
                index=False,
                header=False,
            )
        if biomass_country_meta is not None:
            biomass_country_meta.to_excel(
                writer,
                sheet_name='BiomassMaxAnnualActivityCountry',
                index=False,
                header=False,
            )

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    print()
    print('=' * 80)
    print('NODE.XLSX GENERATION SUMMARY')
    print('=' * 80)
    print(f'Output: {args.output_file}')
    print()
    print('Sheets:')
    if biomass_raw is not None:
        print('  {:35s} {:>6} rows  (reference, non-disaggregated nodes)'.format(
            'BiomassMaxAnnualActivity', len(biomass_raw) - 3
        ))
        print('  {:35s} {:>6} rows  (national limit)'.format(
            'BiomassMaxAnnualActivityCountry', len(biomass_country_df)
        ))
    print(f'  {"ElectricAnnualDemand":35s} {len(ead_df):>6} rows  (generated)')
    print(f'  {"NodeLostLoadCost":35s} {len(nllc_df):>6} rows  (generated)')
    print(f'  {"HydroGenMaxAnnualProduction":35s} {len(hgmap_df):>6} rows  (generated)')
    print()

    # Per-country distribution summaries
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        hydro_rows = hgmap_df[hgmap_df['Nodes'].isin(nuts2_nodes)]
        hydro_total = hydro_rows['HydroGenMaxAnnualProduction in MWh per year'].sum()
        print(f'{prefix} NUTS2 HydroGenMaxAnnualProduction: {hydro_total:,.0f} MWh/yr')

        demand_rows = ead_df[ead_df['Nodes'].isin(nuts2_nodes) &
                             (ead_df['Period'] == periods[0])]
        demand_total = demand_rows['ElectricAdjustment in MWh per hour'].sum()
        print(f'{prefix} ElectricAnnualDemand period {periods[0]} sum: {demand_total:,.0f}')
    print()
    print('OK Node.xlsx generation complete!')
    print('=' * 80)


if __name__ == '__main__':
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    main()
