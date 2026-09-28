"""
Generate Storage.xlsx for EMPIRE from generated Sets.xlsx and IAMC data.

Usage:
    python generate_storage.py <sets_file> <iamc_file> <scenario> <output_file> [--reference <existing_storage.xlsx>]

Example:
    python generate_storage.py Sets_NUTS2_DK_DE.xlsx data/IAMC/NorthSea/Denmark_Germany.csv "Transmission v1.0" Storage_NUTS2_DK_DE.xlsx
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

# IAMC variable for pumped hydro storage capacity
IAMC_PUMPED_STORAGE_VAR = 'Capacity|Electricity|Hydro|Pumped Storage'


# =============================================================================
# Helpers
# =============================================================================


def add_metadata_header(df, source='-', description=''):
    """Add 2 metadata rows + column header row (matching EMPIRE format)."""
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


def read_sets_sheet(sets_file, sheet_name):
    """Read a Sets.xlsx sheet with 2 metadata rows then a header row."""
    df = pd.read_excel(sets_file, sheet_name=sheet_name, header=None)
    headers = df.iloc[2].tolist()
    data = df.iloc[3:].copy()
    data.columns = headers[:data.shape[1]]
    data = data.reset_index(drop=True)
    return data



def read_sets_sheet_simple(sets_file, sheet_name):
    """Read a Sets.xlsx sheet that has a single header row (e.g. Horizon, Nodes)."""
    return pd.read_excel(sets_file, sheet_name=sheet_name)


FALLBACK_REFERENCE = 'data/EMPIRE/v_52/Storage.xlsx'


def _resolve_ref_file(ref_file, sheet_name):
    """Return (file, sheet_name) to use, falling back to v_52 if sheet is missing."""
    import openpyxl
    wb = openpyxl.load_workbook(ref_file, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    wb.close()
    if sheet_name in sheet_names:
        return ref_file
    print(f'  [FALLBACK] Sheet "{sheet_name}" not found in reference — '
          f'using {FALLBACK_REFERENCE}')
    return FALLBACK_REFERENCE


def read_ref_data(ref_file, sheet_name):
    """Read a reference Storage.xlsx data sheet, returning the data portion.

    If the sheet is missing from ref_file, falls back to FALLBACK_REFERENCE.
    """
    src = _resolve_ref_file(ref_file, sheet_name)
    raw = pd.read_excel(src, sheet_name=sheet_name, header=None)
    cols = raw.iloc[2].tolist()
    data = raw.iloc[3:].copy()
    # Only keep columns that actually have headers
    n_cols = sum(1 for c in cols if pd.notna(c))
    data = data.iloc[:, :n_cols].copy()
    data.columns = cols[:n_cols]
    data = data.reset_index(drop=True)
    return data


def classify_nodes(all_nodes):
    """Classify nodes into ({prefix: [nuts2_nodes]}, external_list)."""
    nodes_by_prefix = {
        prefix: [n for n in all_nodes if str(n).startswith(prefix)]
        for prefix in NUTS2_COUNTRY_PREFIXES
    }
    nuts2_set = set(n for nodes in nodes_by_prefix.values() for n in nodes)
    external = [n for n in all_nodes if n not in nuts2_set]
    return nodes_by_prefix, external


def _safe_div(value, count):
    """Return value/count, or 0 when count is zero."""
    return value / count if count > 0 else 0.0


def _pumped_storage_shares(nodes_by_prefix, iamc_ps_mw):
    """Return regional pumped-storage shares for each disaggregated country.

    IAMC capacities are used only as spatial weights. When a country has no
    positive IAMC pumped-storage capacity, fall back to an equal split so that
    the national reference total is never lost.
    """
    shares = {}
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        weights = {
            node: max(float(iamc_ps_mw.get(node, 0) or 0), 0.0)
            for node in nuts2_nodes
        }
        total = sum(weights.values())
        if total > 0:
            shares[prefix] = {
                node: weights[node] / total for node in nuts2_nodes
            }
        else:
            equal_share = _safe_div(1.0, len(nuts2_nodes))
            shares[prefix] = {
                node: equal_share for node in nuts2_nodes
            }
    return shares


def load_iamc_pumped_storage(iamc_file, scenario):
    """Load IAMC pumped storage capacity data.

    Returns dict: NUTS2 region -> capacity in MW (float).
    Handles both 'GW' and 'MW' unit entries from the IAMC file.
    """
    from iamc_utils import load_iamc
    df = load_iamc(iamc_file)
    df_filt = df.filter(scenario=scenario,
                        variable=IAMC_PUMPED_STORAGE_VAR)

    lookup = {}
    if len(df_filt.data) > 0:
        for _, r in df_filt.data.iterrows():
            region = r['region']
            if '>' not in region:  # Skip transmission links
                unit = r.get('unit', 'GW')
                val_mw = r['value'] * 1000 if unit == 'GW' else r['value']
                lookup[region] = val_mw  # MW
    return lookup


# =============================================================================
# Sheet builders — node-dependent sheets
# =============================================================================


def make_initial_power_capacity(storage_of_nodes, periods,
                                iamc_ps_mw, ref_data,
                                nodes_by_prefix):
    """InitialPowerCapacity: Nodes x StorageTypes x Period.

    - External nodes: from reference
    - NUTS2, Hydro Pump: reference country total distributed by each region's
      share of country IAMC pumped-storage capacity
    - NUTS2, Li-Ion_BESS: 0
    """
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        key = (str(r.iloc[0]), str(r.iloc[1]), int(r.iloc[2]))
        ref_lookup[key] = float(r.iloc[3])

    node_to_prefix = {n: p for p, nodes in nodes_by_prefix.items() for n in nodes}
    country_ps_share = _pumped_storage_shares(nodes_by_prefix, iamc_ps_mw)

    rows = []
    for _, son in storage_of_nodes.iterrows():
        node = str(son.iloc[0])
        stype = str(son.iloc[1])
        prefix = node_to_prefix.get(node)
        for p in periods:
            if prefix is not None:
                if stype == 'Hydro Pump Storage':
                    country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
                    country_val = ref_lookup.get((country_name, stype, int(p)), 0)
                    val = country_ps_share[prefix][node] * country_val
                else:
                    val = 0
            else:
                val = ref_lookup.get((node, stype, int(p)), 0)
            rows.append({
                'Nodes': node,
                'StorageTypes': stype,
                'Period': p,
                'InitialPowerCapacity': round(val, 3),
            })
    return pd.DataFrame(rows)


def make_energy_initial_capacity(storage_of_nodes, periods,
                                 ref_data, nodes_by_prefix, iamc_ps_mw):
    """EnergyInitialCapacity: Nodes x StorageTypes x Period.

    - External nodes: from reference
    - NUTS2, Hydro Pump Storage: reference country value weighted by each
      region's share of total country IAMC pumped storage capacity.
    - NUTS2, Li-Ion_BESS: reference country value / n_nuts2 (even split)
      # TODO: Replace even Li-Ion split with proper regional data.
    """
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        key = (str(r.iloc[0]), str(r.iloc[1]), int(r.iloc[2]))
        ref_lookup[key] = float(r.iloc[3])

    # Per-country: IAMC pumped-storage share per node and region count
    country_ps_share = _pumped_storage_shares(nodes_by_prefix, iamc_ps_mw)
    country_n = {}
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        country_n[prefix] = len(nuts2_nodes)

    node_to_prefix = {n: p for p, nodes in nodes_by_prefix.items() for n in nodes}

    rows = []
    for _, son in storage_of_nodes.iterrows():
        node = str(son.iloc[0])
        stype = str(son.iloc[1])
        prefix = node_to_prefix.get(node)
        for p in periods:
            if prefix is not None:
                country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
                n_nuts2 = country_n[prefix]
                if stype == 'Hydro Pump Storage':
                    country_val = ref_lookup.get((country_name, stype, int(p)), 0)
                    val = country_ps_share[prefix][node] * country_val
                else:  # Li-Ion_BESS: even split
                    val = _safe_div(ref_lookup.get((country_name, stype, int(p)), 0), n_nuts2)
            else:
                val = ref_lookup.get((node, stype, int(p)), 0)
            rows.append({
                'Nodes': node,
                'StorageTypes': stype,
                'Period': p,
                'EnergyInitialCapacity': round(val, 3),
            })
    return pd.DataFrame(rows)


def make_max_built_capacity(storage_of_nodes, periods,
                            ref_data, nodes_by_prefix, iamc_ps_mw, col_name):
    """PowerMaxBuiltCapacity or EnergyMaxBuiltCapacity.

    - External: from reference
    - NUTS2, Hydro Pump: reference country total distributed by IAMC shares
    - NUTS2, Li-Ion_BESS: reference country total / n_nuts2 (even split)

    # TODO: Replace the even Li-Ion split with proper regional data.
    """
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        key = (str(r.iloc[0]), str(r.iloc[1]))
        ref_lookup[key] = float(r.iloc[3])

    default_val = 500000  # EMPIRE default

    node_to_prefix = {n: p for p, nodes in nodes_by_prefix.items() for n in nodes}
    country_ps_share = _pumped_storage_shares(nodes_by_prefix, iamc_ps_mw)

    rows = []
    for _, son in storage_of_nodes.iterrows():
        node = str(son.iloc[0])
        stype = str(son.iloc[1])
        prefix = node_to_prefix.get(node)
        for p in periods:
            if prefix is not None:
                country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
                country_val = ref_lookup.get((country_name, stype), default_val)
                if stype == 'Hydro Pump Storage':
                    val = country_ps_share[prefix][node] * country_val
                else:
                    val = _safe_div(country_val, len(nodes_by_prefix[prefix]))
            else:
                val = ref_lookup.get((node, stype), default_val)
            rows.append({
                'Nodes': node,
                'StorageTypes': stype,
                'Period': p,
                col_name: round(val, 3),
            })
    return pd.DataFrame(rows)


def make_power_max_installed_capacity(storage_of_nodes,
                                      iamc_ps_mw, ref_data,
                                      nodes_by_prefix):
    """PowerMaxInstalledCapacity: Nodes x StorageTypes (no Period).

    - External: from reference
    - NUTS2, Hydro Pump: reference country maximum distributed by each region's
      share of country IAMC pumped-storage capacity
    - NUTS2, Li-Ion_BESS: reference country Li-Ion value / n_nuts2

    # TODO: Replace the even Li-Ion split with proper regional data.
    """
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        key = (str(r.iloc[0]), str(r.iloc[1]))
        ref_lookup[key] = float(r.iloc[2])  # 3rd column is value (no Period)

    # Per-country Li-Ion per node (even split)
    country_liion = {
        prefix: _safe_div(
            ref_lookup.get((NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix], 'Li-Ion_BESS'), 0),
            len(nuts2_nodes))
        for prefix, nuts2_nodes in nodes_by_prefix.items()
    }

    node_to_prefix = {n: p for p, nodes in nodes_by_prefix.items() for n in nodes}
    country_ps_share = _pumped_storage_shares(nodes_by_prefix, iamc_ps_mw)

    rows = []
    for _, son in storage_of_nodes.iterrows():
        node = str(son.iloc[0])
        stype = str(son.iloc[1])
        prefix = node_to_prefix.get(node)
        if prefix is not None:
            if stype == 'Hydro Pump Storage':
                country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
                country_val = ref_lookup.get((country_name, stype), 0)
                val = country_ps_share[prefix][node] * country_val
            else:
                val = country_liion[prefix]
        else:
            val = ref_lookup.get((node, stype), 0)
        rows.append({
            'Nodes': node,
            'StorageTypes': stype,
            'PowerMaxInstalledCapacity': round(val, 3),
        })
    return pd.DataFrame(rows)


def make_energy_max_installed_capacity(storage_of_nodes,
                                       ref_data, nodes_by_prefix, iamc_ps_mw):
    """EnergyMaxInstalledCapacity: Nodes x StorageTypes (no Period).

    - External: from reference
    - NUTS2, Hydro Pump Storage: reference country value distributed
      proportionally by IAMC pumped storage capacity (same weighting as
      EnergyInitialCapacity), guaranteeing MaxInstalled >= Initial.
    - NUTS2, Li-Ion_BESS: reference country value / n_nuts2 (even split)
    """
    ref_lookup = {}
    for _, r in ref_data.iterrows():
        key = (str(r.iloc[0]), str(r.iloc[1]))
        ref_lookup[key] = float(r.iloc[2])  # 3rd column is value (no Period)

    # Per-country IAMC-capacity-weighted shares and region counts
    country_ps_share = _pumped_storage_shares(nodes_by_prefix, iamc_ps_mw)
    country_n = {}
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        country_n[prefix] = len(nuts2_nodes)

    node_to_prefix = {n: p for p, nodes in nodes_by_prefix.items() for n in nodes}

    rows = []
    for _, son in storage_of_nodes.iterrows():
        node = str(son.iloc[0])
        stype = str(son.iloc[1])
        prefix = node_to_prefix.get(node)
        if prefix is not None:
            country_name = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
            n_nuts2 = country_n[prefix]
            if stype == 'Hydro Pump Storage':
                country_val = ref_lookup.get((country_name, stype), 0)
                val = country_ps_share[prefix][node] * country_val
            else:  # Li-Ion_BESS: even split
                val = _safe_div(ref_lookup.get((country_name, stype), 0), n_nuts2)
        else:
            val = ref_lookup.get((node, stype), 0)
        rows.append({
            'Nodes': node,
            'StorageTypes': stype,
            'EnergyMaxInstalledCapacity': round(val, 3),
        })
    return pd.DataFrame(rows)


# =============================================================================
# Main
# =============================================================================


def main():
    parser = argparse.ArgumentParser(
        description='Generate Storage.xlsx for EMPIRE')
    parser.add_argument('sets_file', help='Generated Sets.xlsx')
    parser.add_argument('iamc_file', help='IAMC data file (CSV / Excel)')
    parser.add_argument('scenario', help='IAMC scenario name')
    parser.add_argument('output_file', help='Output Storage.xlsx path')
    parser.add_argument('--reference',
                        default='data/EMPIRE/v_52/Storage.xlsx',
                        help='Reference Storage.xlsx '
                             '(default: data/EMPIRE/v_52/Storage.xlsx)')
    args = parser.parse_args()

    print('=' * 80)
    print('STORAGE.XLSX GENERATOR FOR EMPIRE')
    print('=' * 80)
    print(f'Sets:      {args.sets_file}')
    print(f'IAMC:      {args.iamc_file}')
    print(f'Scenario:  {args.scenario}')
    print(f'Reference: {args.reference}')
    print(f'Output:    {args.output_file}')
    print()

    # ------------------------------------------------------------------
    # Read inputs from generated Sets.xlsx
    # ------------------------------------------------------------------
    all_nodes = read_sets_sheet_simple(
        args.sets_file, 'Nodes')['Node'].tolist()
    storage_of_nodes = read_sets_sheet(args.sets_file, 'StorageOfNodes')
    periods = read_sets_sheet_simple(
        args.sets_file, 'Horizon')['Horizon'].tolist()

    nodes_by_prefix, ext_nodes = classify_nodes(all_nodes)

    print(f'Total nodes:    {len(all_nodes)}')
    for prefix, nodes in nodes_by_prefix.items():
        print(f'{prefix} NUTS2 nodes: {len(nodes)}')
    print(f'External nodes: {len(ext_nodes)}')
    print(f'StorageOfNodes: {len(storage_of_nodes)} entries')
    print(f'Periods:        {periods}')
    print()

    # ------------------------------------------------------------------
    # Load IAMC pumped storage data
    # ------------------------------------------------------------------
    print('Loading IAMC pumped storage capacity data...')
    iamc_ps_mw = load_iamc_pumped_storage(args.iamc_file, args.scenario)
    print(f'  IAMC pumped storage entries: {len(iamc_ps_mw)}')
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        total_ps = sum(iamc_ps_mw.get(n, 0) for n in nuts2_nodes)
        print(f'  Total {prefix} pumped storage: {total_ps:.1f} MW')
    print()

    # ------------------------------------------------------------------
    # Load reference data for node-dependent sheets
    # ------------------------------------------------------------------
    print('Loading reference data...')
    ref_ipc = read_ref_data(args.reference, 'InitialPowerCapacity')
    ref_eic = read_ref_data(args.reference, 'EnergyInitialCapacity')
    ref_pmbc = read_ref_data(args.reference, 'PowerMaxBuiltCapacity')
    ref_embc = read_ref_data(args.reference, 'EnergyMaxBuiltCapacity')
    ref_pmic = read_ref_data(args.reference, 'PowerMaxInstalledCapacity')
    ref_emic = read_ref_data(args.reference, 'EnergyMaxInstalledCapacity')
    print('  Reference sheets loaded.')
    print()

    # ------------------------------------------------------------------
    # Sheets copied as-is from reference (no node dimension)
    # ------------------------------------------------------------------
    copy_sheets = [
        'StorageChargeEff',
        'StorageDischargeEff',
        'StoragePowToEnergy',
        'StorageBleedEfficiency',
        'StorageInitialEnergyLevel',
        'Lifetime',
        'PowerCapitalCost',
        'PowerFixedOMCost',
        'EnergyCapitalCost',
        'EnergyFixedOMCost',
    ]

    raw_copies = {}
    print('Copying unchanged sheets from reference:')
    for name in copy_sheets:
        src = _resolve_ref_file(args.reference, name)
        raw_copies[name] = pd.read_excel(src, sheet_name=name, header=None)
        print(f'  {name}: {raw_copies[name].shape[0] - 3} data rows')
    print()

    # ------------------------------------------------------------------
    # Build node-dependent sheets
    # ------------------------------------------------------------------
    print('[1/6] InitialPowerCapacity '
          '(Hydro Pump: IAMC-weighted share of country reference)...')
    ipc_df = make_initial_power_capacity(
        storage_of_nodes, periods, iamc_ps_mw, ref_ipc, nodes_by_prefix)
    nonzero_ipc = (ipc_df['InitialPowerCapacity'] > 0).sum()
    print(f'  -> {len(ipc_df)} rows ({nonzero_ipc} with capacity > 0)')

    print('[2/6] EnergyInitialCapacity '
          '(Hydro Pump: IAMC-weighted share of country ref; Li-Ion: even split)...')
    eic_df = make_energy_initial_capacity(
        storage_of_nodes, periods, ref_eic, nodes_by_prefix, iamc_ps_mw)
    nonzero_eic = (eic_df['EnergyInitialCapacity'] > 0).sum()
    print(f'  -> {len(eic_df)} rows ({nonzero_eic} with capacity > 0)')

    print('[3/6] PowerMaxBuiltCapacity '
          '(Hydro Pump: IAMC-weighted reference; Li-Ion: even split)...')
    pmbc_df = make_max_built_capacity(
        storage_of_nodes, periods, ref_pmbc, nodes_by_prefix, iamc_ps_mw,
        'PowerMaxBuiltCapacity')
    print(f'  -> {len(pmbc_df)} rows')

    print('[4/6] EnergyMaxBuiltCapacity '
          '(Hydro Pump: IAMC-weighted reference; Li-Ion: even split)...')
    embc_df = make_max_built_capacity(
        storage_of_nodes, periods, ref_embc, nodes_by_prefix, iamc_ps_mw,
        'EnergyMaxBuiltCapacity')
    print(f'  -> {len(embc_df)} rows')

    print('[5/6] PowerMaxInstalledCapacity '
          '(Hydro Pump: IAMC-weighted reference; Li-Ion: even split)...')
    pmic_df = make_power_max_installed_capacity(
        storage_of_nodes, iamc_ps_mw, ref_pmic, nodes_by_prefix)
    nonzero_pmic = (pmic_df['PowerMaxInstalledCapacity'] > 0).sum()
    print(f'  -> {len(pmic_df)} rows ({nonzero_pmic} with capacity > 0)')

    print('[6/6] EnergyMaxInstalledCapacity '
          '(IAMC-weighted for NUTS2 Hydro Pump, even split for Li-Ion)...')
    emic_df = make_energy_max_installed_capacity(
        storage_of_nodes, ref_emic, nodes_by_prefix, iamc_ps_mw)
    nonzero_emic = (emic_df['EnergyMaxInstalledCapacity'] > 0).sum()
    print(f'  -> {len(emic_df)} rows ({nonzero_emic} with capacity > 0)')
    print()

    # ------------------------------------------------------------------
    # Add metadata headers
    # ------------------------------------------------------------------
    meta = {}
    meta['InitialPowerCapacity'] = add_metadata_header(
        ipc_df, source=f'Reference + {args.iamc_file} spatial weights',
        description='Initial installed power capacity of storage (MW) '
                    '(default: 0)')
    meta['EnergyInitialCapacity'] = add_metadata_header(
        eic_df, source=f'Reference + {args.iamc_file} spatial weights',
        description='Initial installed energy capacity of storage (MWh) '
                    '(default: 0)')
    meta['PowerMaxBuiltCapacity'] = add_metadata_header(
        pmbc_df, source=f'Reference + {args.iamc_file} spatial weights',
        description='Maximum power capacity expansion of storage per node '
                    'per period (MW) (default: 500 000)')
    meta['EnergyMaxBuiltCapacity'] = add_metadata_header(
        embc_df, source=f'Reference + {args.iamc_file} spatial weights',
        description='Maximum energy capacity expansion of storage per node '
                    'per period (MWh) (default: 500 000)')
    meta['PowerMaxInstalledCapacity'] = add_metadata_header(
        pmic_df, source=f'Reference + {args.iamc_file} spatial weights',
        description='Maximum total power capacity of storage (MW) '
                    '(default: 0)')
    meta['EnergyMaxInstalledCapacity'] = add_metadata_header(
        emic_df, source=f'Reference + {args.iamc_file} spatial weights',
        description='Maximum total energy capacity of storage (MWh) '
                    '(default: 0)')

    # ------------------------------------------------------------------
    # Write output — preserve reference sheet order
    # ------------------------------------------------------------------
    # Reference order: InitialPowerCapacity, PowerCapitalCost, PowerFixedOMCost,
    #   PowerMaxBuiltCapacity, EnergyCapitalCost, EnergyFixedOMCost,
    #   EnergyInitialCapacity, EnergyMaxBuiltCapacity, EnergyMaxInstalledCapacity,
    #   PowerMaxInstalledCapacity, StorageInitialEnergyLevel, StorageChargeEff,
    #   StorageDischargeEff, StoragePowToEnergy, StorageBleedEfficiency, Lifetime
    sheet_order = [
        ('InitialPowerCapacity', 'meta'),
        ('PowerCapitalCost', 'copy'),
        ('PowerFixedOMCost', 'copy'),
        ('PowerMaxBuiltCapacity', 'meta'),
        ('EnergyCapitalCost', 'copy'),
        ('EnergyFixedOMCost', 'copy'),
        ('EnergyInitialCapacity', 'meta'),
        ('EnergyMaxBuiltCapacity', 'meta'),
        ('EnergyMaxInstalledCapacity', 'meta'),
        ('PowerMaxInstalledCapacity', 'meta'),
        ('StorageInitialEnergyLevel', 'copy'),
        ('StorageChargeEff', 'copy'),
        ('StorageDischargeEff', 'copy'),
        ('StoragePowToEnergy', 'copy'),
        ('StorageBleedEfficiency', 'copy'),
        ('Lifetime', 'copy'),
    ]

    print(f'Writing {args.output_file}...')
    with pd.ExcelWriter(args.output_file, engine='openpyxl') as writer:
        for name, kind in sheet_order:
            if kind == 'copy':
                raw_copies[name].to_excel(
                    writer, sheet_name=name,
                    index=False, header=False)
            else:
                meta[name].to_excel(
                    writer, sheet_name=name,
                    index=False, header=False)

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    print()
    print('=' * 80)
    print('STORAGE.XLSX GENERATION SUMMARY')
    print('=' * 80)
    print(f'Output: {args.output_file}')
    print()
    print('Sheets:')
    for name, kind in sheet_order:
        if kind == 'copy':
            print(f'  {name:35s} (copied from reference)')
        else:
            df_map = {
                'InitialPowerCapacity': ipc_df,
                'EnergyInitialCapacity': eic_df,
                'PowerMaxBuiltCapacity': pmbc_df,
                'EnergyMaxBuiltCapacity': embc_df,
                'PowerMaxInstalledCapacity': pmic_df,
                'EnergyMaxInstalledCapacity': emic_df,
            }
            nrows = len(df_map[name])
            print(f'  {name:35s} {nrows:>6} rows  (generated)')
    print()

    # Show the raw IAMC capacities used only to derive regional shares
    for prefix, nuts2_nodes in nodes_by_prefix.items():
        print(f'{prefix} NUTS2 Pumped Storage weights (IAMC, MW):')
        ps_nodes = {n: iamc_ps_mw.get(n, 0) for n in sorted(nuts2_nodes)}
        for node, val in ps_nodes.items():
            print(f'  {node}: {val:>10.1f} MW')
        print(f'  {"TOTAL":5s}: {sum(ps_nodes.values()):>10.1f} MW')
        print()

    country_list = ', '.join(NUTS2_COUNTRY_PREFIXES)
    print(f'NOTE: For Hydro Pump Storage in {country_list}, all national')
    print('  reference totals are conserved and IAMC is used only for regional')
    print('  shares. Countries without positive IAMC weights use an equal split.')
    print('  Li-Ion_BESS values remain evenly distributed across NUTS2 regions.')
    print()
    print('OK Storage.xlsx generation complete!')
    print('=' * 80)


if __name__ == '__main__':
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    main()
