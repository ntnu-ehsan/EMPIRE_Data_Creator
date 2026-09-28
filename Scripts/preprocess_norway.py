"""
Merge Norwegian market zones (NO1–NO5) into a single 'Norway' node.

The main pipeline (main.py) expects each country to be represented by exactly
one EMPIRE node.  Norway is an exception: the reference datasets use five
market/price zones (NO1–NO5).  This script preprocesses a reference EMPIRE
dataset by aggregating those five zones into a single 'Norway' node so that
the pipeline can then disaggregate it to Norwegian NUTS2 regions exactly as it
does for DE, DK, and NL.

Usage
-----
    python Scripts/preprocess_norway.py <input_folder> <output_folder>

    # Process all four reference datasets at once:
    python Scripts/preprocess_norway.py --all

Example
-------
    python Scripts/preprocess_norway.py \\
        data/EMPIRE/NECPEssentials \\
        data/EMPIRE/NECPEssentials_Norway

What it does
------------
  Generator.xlsx   Sum RefInitialCap, InitialCapacity, MaxBuiltCapacity,
                   MaxInstalledCapacity across NO1–NO5 → Norway.
  Node.xlsx        Sum ElectricAnnualDemand, HydroGenMaxAnnualProduction and
                   BiomassMaxAnnualActivity (when present);
                   copy NodeLostLoadCost value from NO1.
  Storage.xlsx     Sum all per-node capacity sheets across NO1–NO5 → Norway.
  Transmission.xlsx Remove internal NO corridors; sum capacity on external
                   corridors that share an endpoint; keep one value for
                   Lifetime / lineEfficiency (all Norwegian zones use the same).
  Sets.xlsx        Replace NO1–NO5 with Norway throughout all sheets.
  ScenarioData/    electricload.csv: sum NO1–NO5 columns → 'NO'.
                   Capacity-factor CSVs: capacity-weighted average → 'NO'.
                   windoffshore.csv: already has 'NO' column, kept as-is.
  General.xlsx     Copied unchanged.
"""

import argparse
import shutil
import sys
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import numpy as np
import pandas as pd
import openpyxl

ROOT = Path(__file__).parent.parent
# The reference datasets.  Each folder is named only after its scenario.
# These names are used by --all; a non-standard folder can still be explicit.
ALL_REFERENCE_FOLDERS = [
    'GoRES',
    'NECPEssentials',
    'REPowerEU++',
    'Trinity',
]

NO_ZONES = ['NO1', 'NO2', 'NO3', 'NO4', 'NO5']
NO_ZONES_SET = set(NO_ZONES)
NORWAY = 'Norway'

# ─────────────────────────────────────────────────────────────────────────────
# Generic Excel helpers
# ─────────────────────────────────────────────────────────────────────────────

def _read_raw(path, sheet):
    """Return list-of-lists from a sheet (all values, no interpretation)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()
    # Strip trailing all-None rows
    while rows and all(c is None for c in rows[-1]):
        rows.pop()
    return rows


def _read_empire_df(path, sheet, meta=2):
    """Read a sheet with `meta` metadata rows then a header row.

    Returns (meta_rows, df) where meta_rows is a list of raw rows and df uses
    the header row as column names. Trailing-None columns are dropped.
    """
    rows = _read_raw(path, sheet)
    if len(rows) <= meta:
        return rows, pd.DataFrame()
    meta_rows = rows[:meta]
    header = rows[meta]
    data = rows[meta + 1:]
    df = pd.DataFrame(data, columns=header)
    # Drop columns with None header (formatting artefacts)
    df = df.loc[:, df.columns.notna()]
    # Drop rows that are entirely None
    df = df.dropna(how='all').reset_index(drop=True)
    return meta_rows, df


def _write_empire_sheet(ws, meta_rows, header, df):
    """Write meta rows, then header row, then data rows into ws."""
    for row in meta_rows:
        ws.append(row)
    ws.append(list(header))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))


def _copy_sheet(src_path, dst_wb, sheet):
    """Copy a sheet verbatim from src_path into dst_wb."""
    rows = _read_raw(src_path, sheet)
    ws = dst_wb.create_sheet(sheet)
    for row in rows:
        ws.append(row)


# ─────────────────────────────────────────────────────────────────────────────
# Aggregation helpers
# ─────────────────────────────────────────────────────────────────────────────

def _aggregate_node_sheet(df, node_col, group_cols, value_col):
    """Remove NO zone rows and add a Norway row with summed values.

    group_cols: columns to group by (excluding node_col) when summing.
    value_col:  name of the numeric column to sum.
    """
    non_no = df[~df[node_col].isin(NO_ZONES_SET)].copy()
    no_df  = df[df[node_col].isin(NO_ZONES_SET)].copy()

    if no_df.empty:
        return non_no

    no_df[value_col] = pd.to_numeric(no_df[value_col], errors='coerce').fillna(0)
    agg = no_df.groupby(group_cols, dropna=False)[value_col].sum().reset_index()
    agg.insert(0, node_col, NORWAY)

    return pd.concat([non_no, agg], ignore_index=True)


def _aggregate_node_sheet_no_period(df, node_col, group_cols, value_col):
    """Like _aggregate_node_sheet but without a period dimension."""
    return _aggregate_node_sheet(df, node_col, group_cols, value_col)


# ─────────────────────────────────────────────────────────────────────────────
# Generator.xlsx
# ─────────────────────────────────────────────────────────────────────────────

# Sheets with (Node, GeneratorTechnology, Period, Value)
_GEN_NODE_PERIOD = ['InitialCapacity', 'MaxBuiltCapacity']
# Sheets with (Node, GeneratorTechnology, Value)
_GEN_NODE_ONLY   = ['RefInitialCap', 'MaxInstalledCapacity']
# All other sheets: copy verbatim
_GEN_COPY = [
    'CapitalCosts', 'FixedOMCosts', 'VariableOMCosts', 'FuelCosts',
    'ScaleFactorInitialCap', 'Efficiency', 'CO2Content', 'Lifetime',
    'GeneratorTypeAvailability', 'RampRate', 'CCSCostTSVariable',
]


def preprocess_generator(src, dst_wb):
    print('  Generator.xlsx')
    wb_sheets = openpyxl.load_workbook(src, read_only=True).sheetnames

    for sheet in wb_sheets:
        if sheet in _GEN_COPY:
            _copy_sheet(src, dst_wb, sheet)

        elif sheet in _GEN_NODE_PERIOD:
            meta, df = _read_empire_df(src, sheet)
            node_col = df.columns[0]
            gen_col  = df.columns[1]
            per_col  = df.columns[2]
            val_col  = df.columns[3]
            df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)
            df[per_col] = df[per_col].apply(lambda x: int(x) if pd.notna(x) else x)
            out = _aggregate_node_sheet(df, node_col, [gen_col, per_col], val_col)
            ws = dst_wb.create_sheet(sheet)
            _write_empire_sheet(ws, meta, df.columns.tolist(), out)
            print(f'    {sheet}: aggregated ({out[node_col].eq(NORWAY).sum()} Norway rows)')

        elif sheet in _GEN_NODE_ONLY:
            meta, df = _read_empire_df(src, sheet)
            node_col = df.columns[0]
            gen_col  = df.columns[1]
            val_col  = df.columns[2]
            df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)
            out = _aggregate_node_sheet(df, node_col, [gen_col], val_col)
            ws = dst_wb.create_sheet(sheet)
            _write_empire_sheet(ws, meta, df.columns.tolist(), out)
            print(f'    {sheet}: aggregated ({out[node_col].eq(NORWAY).sum()} Norway rows)')

        else:
            _copy_sheet(src, dst_wb, sheet)


# ─────────────────────────────────────────────────────────────────────────────
# Storage.xlsx
# ─────────────────────────────────────────────────────────────────────────────

_STO_NODE_PERIOD = [
    'InitialPowerCapacity', 'EnergyInitialCapacity',
    'EnergyMaxBuiltCapacity', 'PowerMaxBuiltCapacity',
]
_STO_NODE_ONLY = [
    'EnergyMaxInstalledCapacity', 'PowerMaxInstalledCapacity',
]
_STO_COPY = [
    'PowerCapitalCost', 'PowerFixedOMCost', 'EnergyCapitalCost',
    'EnergyFixedOMCost', 'StorageChargeEff', 'StorageDischargeEff',
    'Lifetime', 'StorageInitialEnergyLevel', 'StoragePowToEnergy',
    'StorageBleedEfficiency',
]


def preprocess_storage(src, dst_wb):
    print('  Storage.xlsx')
    wb_sheets = openpyxl.load_workbook(src, read_only=True).sheetnames

    for sheet in wb_sheets:
        if sheet in _STO_COPY:
            _copy_sheet(src, dst_wb, sheet)

        elif sheet in _STO_NODE_PERIOD:
            meta, df = _read_empire_df(src, sheet)
            node_col = df.columns[0]
            sto_col  = df.columns[1]
            per_col  = df.columns[2]
            val_col  = df.columns[3]
            df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)
            df[per_col] = df[per_col].apply(lambda x: int(x) if pd.notna(x) else x)
            out = _aggregate_node_sheet(df, node_col, [sto_col, per_col], val_col)
            ws = dst_wb.create_sheet(sheet)
            _write_empire_sheet(ws, meta, df.columns.tolist(), out)
            print(f'    {sheet}: aggregated')

        elif sheet in _STO_NODE_ONLY:
            meta, df = _read_empire_df(src, sheet)
            node_col = df.columns[0]
            sto_col  = df.columns[1]
            val_col  = df.columns[2]
            df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)
            out = _aggregate_node_sheet(df, node_col, [sto_col], val_col)
            ws = dst_wb.create_sheet(sheet)
            _write_empire_sheet(ws, meta, df.columns.tolist(), out)
            print(f'    {sheet}: aggregated')

        else:
            _copy_sheet(src, dst_wb, sheet)


# ─────────────────────────────────────────────────────────────────────────────
# Node.xlsx
# ─────────────────────────────────────────────────────────────────────────────

def preprocess_node(src, dst_wb):
    print('  Node.xlsx')

    # BiomassMaxAnnualActivity (optional): (Node, Period, Value) — sum.
    # The limit caps annual biomass production in a node, so it is extensive and
    # the five zones' budgets add up to Norway's.  Must be carried through:
    # generate_node.py reads the country row from here to build
    # BiomassMaxAnnualActivityCountry, and silently skips the whole feature when
    # the sheet is absent — leaving Norwegian biomass unconstrained.
    if 'BiomassMaxAnnualActivity' in openpyxl.load_workbook(
            src, read_only=True).sheetnames:
        meta, df = _read_empire_df(src, 'BiomassMaxAnnualActivity')
        nc, pc, vc = df.columns[0], df.columns[1], df.columns[2]
        df[vc] = pd.to_numeric(df[vc], errors='coerce').fillna(0)
        df[pc] = df[pc].apply(lambda x: int(x) if pd.notna(x) else x)
        out = _aggregate_node_sheet(df, nc, [pc], vc)
        ws = dst_wb.create_sheet('BiomassMaxAnnualActivity')
        _write_empire_sheet(ws, meta, df.columns.tolist(), out)
        print(f'    BiomassMaxAnnualActivity: Norway = '
              f'{out[out[nc]==NORWAY][vc].sum():.0f}')
    else:
        print('    WARNING: BiomassMaxAnnualActivity not in reference, skipped '
              '— downstream biomass limits will be unconstrained')

    # ElectricAnnualDemand: (Node, Period, Value) — sum
    meta, df = _read_empire_df(src, 'ElectricAnnualDemand')
    nc, pc, vc = df.columns[0], df.columns[1], df.columns[2]
    df[vc] = pd.to_numeric(df[vc], errors='coerce').fillna(0)
    df[pc] = df[pc].apply(lambda x: int(x) if pd.notna(x) else x)
    out = _aggregate_node_sheet(df, nc, [pc], vc)
    ws = dst_wb.create_sheet('ElectricAnnualDemand')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out)
    print(f'    ElectricAnnualDemand: Norway demand = {out[out[nc]==NORWAY][vc].sum():.0f}')

    # NodeLostLoadCost: (Node, Period, Value) — copy NO1 value for Norway
    meta, df = _read_empire_df(src, 'NodeLostLoadCost')
    nc, pc, vc = df.columns[0], df.columns[1], df.columns[2]
    no1_rows = df[df[nc] == 'NO1'].copy()
    no1_rows[nc] = NORWAY
    non_no = df[~df[nc].isin(NO_ZONES_SET)].copy()
    out = pd.concat([non_no, no1_rows], ignore_index=True)
    ws = dst_wb.create_sheet('NodeLostLoadCost')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out)
    print(f'    NodeLostLoadCost: Norway value copied from NO1')

    # HydroGenMaxAnnualProduction: (Node, Value) — sum
    meta, df = _read_empire_df(src, 'HydroGenMaxAnnualProduction')
    nc, vc = df.columns[0], df.columns[1]
    df[vc] = pd.to_numeric(df[vc], errors='coerce').fillna(0)
    no_sum = df[df[nc].isin(NO_ZONES_SET)][vc].sum()
    non_no = df[~df[nc].isin(NO_ZONES_SET)].copy()
    norway_row = pd.DataFrame({nc: [NORWAY], vc: [no_sum]})
    out = pd.concat([non_no, norway_row], ignore_index=True)
    ws = dst_wb.create_sheet('HydroGenMaxAnnualProduction')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out)
    print(f'    HydroGenMaxAnnualProduction: Norway = {no_sum:.0f}')


# ─────────────────────────────────────────────────────────────────────────────
# Transmission.xlsx
# ─────────────────────────────────────────────────────────────────────────────

def _is_internal(a, b):
    return a in NO_ZONES_SET and b in NO_ZONES_SET


def _remap_node(n):
    return NORWAY if n in NO_ZONES_SET else n


def preprocess_transmission(src, dst_wb):
    """Returns canonical_arcs: set of (From, To) tuples used in parameter sheets."""
    print('  Transmission.xlsx')

    # ── InitialCapacity: (From, To, Period, Value) ──────────────────────────
    meta, df = _read_empire_df(src, 'InitialCapacity')
    fc, tc, pc, vc = df.columns[0], df.columns[1], df.columns[2], df.columns[3]
    df[vc] = pd.to_numeric(df[vc], errors='coerce').fillna(0)
    df[pc] = df[pc].apply(lambda x: int(x) if pd.notna(x) else x)
    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    kept = kept.groupby([fc, tc, pc], dropna=False)[vc].sum().reset_index()
    ws = dst_wb.create_sheet('InitialCapacity')
    _write_empire_sheet(ws, meta, df.columns.tolist(), kept)
    print(f'    InitialCapacity: {mask_internal.sum()} internal rows removed')

    # ── MaxInstallCapacityRaw: (From, To, Period, MaxRaw, MaxWithTech, ind) ─
    meta, df = _read_empire_df(src, 'MaxInstallCapacityRaw')
    fc, tc, pc = df.columns[0], df.columns[1], df.columns[2]
    val_cols = list(df.columns[3:])
    for c in val_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    df[pc] = df[pc].apply(lambda x: int(x) if pd.notna(x) else x)
    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    # Sum the first two numeric cols; take max of the indicator (last col)
    if len(val_cols) >= 3:
        agg_dict = {val_cols[0]: 'sum', val_cols[1]: 'sum', val_cols[2]: 'max'}
    else:
        agg_dict = {c: 'sum' for c in val_cols}
    kept = kept.groupby([fc, tc, pc], dropna=False).agg(agg_dict).reset_index()
    ws = dst_wb.create_sheet('MaxInstallCapacityRaw')
    _write_empire_sheet(ws, meta, df.columns.tolist(), kept)

    # ── Length: (From, To, Length) — capacity-weighted average ──────────────
    meta, df = _read_empire_df(src, 'Length')
    fc, tc, lc = df.columns[0], df.columns[1], df.columns[2]
    df[lc] = pd.to_numeric(df[lc], errors='coerce')
    # Load InitialCapacity period 1 for weighting
    _, ic_df = _read_empire_df(src, 'InitialCapacity')
    ic_fc, ic_tc, ic_pc, ic_vc = ic_df.columns[0], ic_df.columns[1], ic_df.columns[2], ic_df.columns[3]
    ic_df[ic_vc] = pd.to_numeric(ic_df[ic_vc], errors='coerce').fillna(0)
    cap_p1 = ic_df[ic_df[ic_pc] == ic_df[ic_pc].min()]

    def weighted_length(group):
        # Match each row to its remapped corridor capacity
        total_w, total_wl = 0, 0
        for _, row in group.iterrows():
            orig_f, orig_t, length = row[fc], row[tc], row[lc]
            # Capacity for this specific corridor (could be 0)
            cap = cap_p1[
                ((cap_p1[ic_fc] == orig_f) & (cap_p1[ic_tc] == orig_t)) |
                ((cap_p1[ic_fc] == orig_t) & (cap_p1[ic_tc] == orig_f))
            ][ic_vc].sum()
            w = max(cap, 1)  # avoid zero weights
            total_w  += w
            total_wl += w * (length or 0)
        return total_wl / total_w if total_w else 0

    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    # Group pairs that now share the same From/To after remapping
    groups = kept.groupby([fc, tc], dropna=False)
    length_rows = []
    for (f, t), grp in groups:
        if len(grp) == 1:
            length_rows.append({fc: f, tc: t, lc: grp.iloc[0][lc]})
        else:
            length_rows.append({fc: f, tc: t, lc: round(weighted_length(grp))})
    out_len = pd.DataFrame(length_rows)
    ws = dst_wb.create_sheet('Length')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out_len)
    # Capture the canonical arc orientation (used to order DirectionalLines)
    canonical_arcs = set(zip(out_len[fc].astype(str), out_len[tc].astype(str)))

    # ── MaxBuiltCapacity: (From, To, Period, Value) — sum ───────────────────
    meta, df = _read_empire_df(src, 'MaxBuiltCapacity')
    fc, tc, pc, vc = df.columns[0], df.columns[1], df.columns[2], df.columns[3]
    df[vc] = pd.to_numeric(df[vc], errors='coerce').fillna(0)
    df[pc] = df[pc].apply(lambda x: int(x) if pd.notna(x) else x)
    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    kept = kept.groupby([fc, tc, pc], dropna=False)[vc].sum().reset_index()
    ws = dst_wb.create_sheet('MaxBuiltCapacity')
    _write_empire_sheet(ws, meta, df.columns.tolist(), kept)

    # ── lineEfficiency: (From, To, Efficiency) — first value (all same=0.97) ─
    meta, df = _read_empire_df(src, 'lineEfficiency')
    fc, tc, vc = df.columns[0], df.columns[1], df.columns[2]
    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    kept = kept.drop_duplicates(subset=[fc, tc])
    ws = dst_wb.create_sheet('lineEfficiency')
    _write_empire_sheet(ws, meta, df.columns.tolist(), kept)

    # ── Lifetime: (From, To, Lifetime) — first value (all same=40) ──────────
    meta, df = _read_empire_df(src, 'Lifetime')
    fc, tc = df.columns[0], df.columns[1]
    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    kept = kept.drop_duplicates(subset=[fc, tc])
    ws = dst_wb.create_sheet('Lifetime')
    _write_empire_sheet(ws, meta, df.columns.tolist(), kept)

    # ── TypeCapitalCost, TypeFixedOMCost: no node dimension ─────────────────
    for sheet in ['TypeCapitalCost', 'TypeFixedOMCost']:
        if sheet in openpyxl.load_workbook(src, read_only=True).sheetnames:
            _copy_sheet(src, dst_wb, sheet)

    print(f'    Done')
    return canonical_arcs


# ─────────────────────────────────────────────────────────────────────────────
# Sets.xlsx
# ─────────────────────────────────────────────────────────────────────────────

def preprocess_sets(src, dst_wb, canonical_arcs=None):
    print('  Sets.xlsx')
    wb_src = openpyxl.load_workbook(src, read_only=True)
    src_sheets = wb_src.sheetnames

    # ── Nodes ─────────────────────────────────────────────────────────────
    raw = _read_raw(src, 'Nodes')
    header = raw[0]  # ('Node',)
    nodes_in = [r[0] for r in raw[1:] if r[0] is not None]
    nodes_out = [n for n in nodes_in if n not in NO_ZONES_SET] + [NORWAY]
    nodes_out = sorted(set(nodes_out))
    ws = dst_wb.create_sheet('Nodes')
    ws.append(list(header))
    for n in nodes_out:
        ws.append([n])
    print(f'    Nodes: {len(nodes_in)} -> {len(nodes_out)} (replaced NO1-NO5 with Norway)')

    # ── OffshoreNodes, Horizon, Storage, Technology, Generators, LineType ──
    for sheet in ['OffshoreNodes', 'Horizon', 'Storage', 'Technology',
                  'Generators', 'LineType']:
        if sheet in src_sheets:
            _copy_sheet(src, dst_wb, sheet)

    # ── HourOfSeason, GeneratorsOfTechnology (raw copies) ──────────────────
    for sheet in ['HourOfSeason', 'GeneratorsOfTechnology']:
        if sheet in src_sheets:
            _copy_sheet(src, dst_wb, sheet)

    # ── StorageOfNodes: (Node, StorageType) — replace NO zones with Norway ─
    meta, df = _read_empire_df(src, 'StorageOfNodes')
    nc = df.columns[0]
    non_no = df[~df[nc].isin(NO_ZONES_SET)].copy()
    # Norway gets all storage types that any NO zone had
    no_types = df[df[nc].isin(NO_ZONES_SET)].iloc[:, 1].unique()
    norway_rows = pd.DataFrame({nc: NORWAY, df.columns[1]: no_types})
    out = pd.concat([non_no, norway_rows], ignore_index=True).drop_duplicates()
    ws = dst_wb.create_sheet('StorageOfNodes')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out)

    # ── DirectionalLines ────────────────────────────────────────────────────
    meta, df = _read_empire_df(src, 'DirectionalLines')
    fc, tc = df.columns[0], df.columns[1]
    # Remove internal NO corridors
    mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
    kept = df[~mask_internal].copy()
    kept[fc] = kept[fc].map(_remap_node)
    kept[tc] = kept[tc].map(_remap_node)
    kept = kept.drop_duplicates().reset_index(drop=True)
    # Build output with canonical orientation first, then its reverse.
    # The canonical orientation is whichever direction appears in the
    # transmission parameter sheets (passed in via canonical_arcs).
    # EMPIRE builds BidirectionalArc by keeping the first occurrence of each
    # undirected pair, so the parameter-sheet orientation must come first.
    seen_lines = set()
    ordered_rows = []
    for _, row in kept.iterrows():
        a, b = str(row[fc]), str(row[tc])
        line_key = frozenset([a, b])
        if line_key in seen_lines:
            continue
        seen_lines.add(line_key)
        if canonical_arcs and (b, a) in canonical_arcs:
            # Reverse is canonical — put (b,a) first
            ordered_rows.append({fc: b, tc: a})
            ordered_rows.append({fc: a, tc: b})
        else:
            # (a,b) is canonical (or no canonical info) — put (a,b) first
            ordered_rows.append({fc: a, tc: b})
            ordered_rows.append({fc: b, tc: a})
    out_dl = pd.DataFrame(ordered_rows)
    ws = dst_wb.create_sheet('DirectionalLines')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out_dl)
    n_removed = mask_internal.sum()
    print(f'    DirectionalLines: removed {n_removed} internal NO corridors, '
          f'{len(seen_lines)} bidirectional lines written')

    # ── LineTypeOfDirectionalLines ──────────────────────────────────────────
    if 'LineTypeOfDirectionalLines' in src_sheets:
        meta, df = _read_empire_df(src, 'LineTypeOfDirectionalLines')
        fc, tc, lc = df.columns[0], df.columns[1], df.columns[2]
        mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
        kept = df[~mask_internal].copy()
        kept[fc] = kept[fc].map(_remap_node)
        kept[tc] = kept[tc].map(_remap_node)
        kept = kept.drop_duplicates(subset=[fc, tc]).sort_values([fc, tc]).reset_index(drop=True)
        ws = dst_wb.create_sheet('LineTypeOfDirectionalLines')
        _write_empire_sheet(ws, meta, df.columns.tolist(), kept)

    # ── BidirectionalLines ──────────────────────────────────────────────────
    if 'BidirectionalLines' in src_sheets:
        meta, df = _read_empire_df(src, 'BidirectionalLines')
        fc, tc = df.columns[0], df.columns[1]
        mask_internal = df.apply(lambda r: _is_internal(r[fc], r[tc]), axis=1)
        kept = df[~mask_internal].copy()
        kept[fc] = kept[fc].map(_remap_node)
        kept[tc] = kept[tc].map(_remap_node)
        seen = set()
        rows = []
        for _, row in kept.iterrows():
            key = tuple(sorted([row[fc], row[tc]]))
            if key not in seen:
                seen.add(key)
                rows.append(row)
        out = pd.DataFrame(rows).sort_values([fc, tc]).reset_index(drop=True)
        ws = dst_wb.create_sheet('BidirectionalLines')
        _write_empire_sheet(ws, meta, df.columns.tolist(), out)

    # ── GeneratorsOfNode ────────────────────────────────────────────────────
    meta, df = _read_empire_df(src, 'GeneratorsOfNode')
    nc, gc = df.columns[0], df.columns[1]
    non_no = df[~df[nc].isin(NO_ZONES_SET)].copy()
    # Union of all generator types across NO zones
    no_gens = df[df[nc].isin(NO_ZONES_SET)][gc].unique()
    norway_rows = pd.DataFrame({nc: NORWAY, gc: no_gens})
    out = pd.concat([non_no, norway_rows], ignore_index=True).drop_duplicates()
    out = out.sort_values([nc, gc]).reset_index(drop=True)
    ws = dst_wb.create_sheet('GeneratorsOfNode')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out)

    # ── Coords ──────────────────────────────────────────────────────────────
    meta, df = _read_empire_df(src, 'Coords')
    loc_col, lat_col, lon_col = df.columns[0], df.columns[1], df.columns[2]
    non_no = df[~df[loc_col].isin(NO_ZONES_SET)].copy()
    no_df = df[df[loc_col].isin(NO_ZONES_SET)].copy()
    lat = pd.to_numeric(no_df[lat_col], errors='coerce').mean()
    lon = pd.to_numeric(no_df[lon_col], errors='coerce').mean()
    norway_coord = pd.DataFrame({loc_col: [NORWAY], lat_col: [round(lat, 6)], lon_col: [round(lon, 6)]})
    out = pd.concat([non_no, norway_coord], ignore_index=True).sort_values(loc_col).reset_index(drop=True)
    ws = dst_wb.create_sheet('Coords')
    _write_empire_sheet(ws, meta, df.columns.tolist(), out)
    print(f'    Coords: Norway centroid = ({lat:.4f}, {lon:.4f})')

    print('    Sets done')


# ─────────────────────────────────────────────────────────────────────────────
# ScenarioData
# ─────────────────────────────────────────────────────────────────────────────

def _get_ref_initial_cap(gen_file):
    """Return {zone: {generator: MW}} from Generator.xlsx RefInitialCap."""
    _, df = _read_empire_df(gen_file, 'RefInitialCap')
    nc, gc, vc = df.columns[0], df.columns[1], df.columns[2]
    df[vc] = pd.to_numeric(df[vc], errors='coerce').fillna(0)
    cap = {}
    for _, row in df[df[nc].isin(NO_ZONES_SET)].iterrows():
        zone = row[nc]
        gen  = row[gc]
        val  = row[vc]
        cap.setdefault(zone, {})[gen] = cap.get(zone, {}).get(gen, 0) + val
    return cap


def _zone_weights(cap, tech_keys):
    """Return normalised per-zone weights for the given generator types."""
    totals = {}
    for zone in NO_ZONES:
        totals[zone] = sum(cap.get(zone, {}).get(t, 0) for t in tech_keys)
    grand = sum(totals.values())
    if grand == 0:
        return {z: 1.0 / len(NO_ZONES) for z in NO_ZONES}
    return {z: v / grand for z, v in totals.items()}


# CSV file → (method, generator types for capacity weights)
_CSV_CONFIG = {
    'electricload.csv':       ('sum',      []),
    'hydroror.csv':           ('weighted', ['Hydro regulated', 'Hydro run-of-the-river']),
    # hydroseasonal holds ABSOLUTE seasonal inflow (EMPIRE's maxRegHydroGenRaw,
    # an energy budget), not a capacity factor -> extensive, so sum.
    'hydroseasonal.csv':      ('sum',      []),
    'solar.csv':              ('weighted', ['Solar']),
    'windonshore.csv':        ('weighted', ['Wind onshore']),
    'windoffshore.csv':          ('rename', []),   # already single NO column → rename to Norway
    'windoffshore_original.csv': ('rename', []),  # backup copy in noffshore datasets
    'windoffshore_cleaned.csv':  ('rename', []),
}


def preprocess_scenario_data(src_dir, dst_dir, gen_file):
    print('  ScenarioData/')
    src_dir = Path(src_dir)
    dst_dir = Path(dst_dir)
    dst_dir.mkdir(parents=True, exist_ok=True)

    cap = _get_ref_initial_cap(gen_file)

    for fname in sorted(src_dir.iterdir()):
        if not fname.suffix == '.csv':
            shutil.copy2(fname, dst_dir / fname.name)
            continue

        config = _CSV_CONFIG.get(fname.name)
        if config is None:
            # Unknown file — copy unchanged
            shutil.copy2(fname, dst_dir / fname.name)
            continue

        method, tech_keys = config
        df = pd.read_csv(fname)

        if method == 'rename':
            # windoffshore: already has a single 'NO' column → rename to 'Norway'
            # so EMPIRE's scenario_random.py does not split it into NO1–NO5
            if 'NO' in df.columns:
                df = df.rename(columns={'NO': NORWAY})
            df.to_csv(dst_dir / fname.name, index=False)
            print(f'    {fname.name}: renamed NO → Norway')

        elif method == 'sum':
            present = [z for z in NO_ZONES if z in df.columns]
            if present:
                df[NORWAY] = df[present].sum(axis=1)
                df = df.drop(columns=present)
            df.to_csv(dst_dir / fname.name, index=False)
            print(f'    {fname.name}: summed {present} → Norway')

        elif method == 'weighted':
            weights = _zone_weights(cap, tech_keys)
            present = [z for z in NO_ZONES if z in df.columns]
            if present:
                norway_cf = sum(df[z] * weights.get(z, 0) for z in present)
                df[NORWAY] = norway_cf
                df = df.drop(columns=present)
            df.to_csv(dst_dir / fname.name, index=False)
            w_str = ', '.join(f'{z}:{weights.get(z,0):.2f}' for z in NO_ZONES)
            print(f'    {fname.name}: weighted average [{w_str}] → Norway')


# ─────────────────────────────────────────────────────────────────────────────
# Consistency assertions
# ─────────────────────────────────────────────────────────────────────────────

def _assert_consistency(out_dir):
    """Assert the three data-consistency invariants required by EMPIRE.

    1. DirectionalLines orientation matches all transmission parameter sheets.
    2. ScenarioData CSVs have no NO/NO1-NO5 columns (must be named Norway).
    3. GeneratorsOfNode for Norway covers every generator in ScenarioData CSVs.
    """
    import re

    def strip_ws(s):
        return re.sub(r'\s+', '', str(s))

    violations = 0
    out_dir = Path(out_dir)

    # ── 1: DirectionalLines vs. transmission parameter sheets ────────────────
    _, dl_df = _read_empire_df(out_dir / 'Sets.xlsx', 'DirectionalLines')
    dl_fc, dl_tc = dl_df.columns[0], dl_df.columns[1]
    # Build BidirectionalArc exactly as EMPIRE does: first orientation seen wins
    arc_set = set()
    for _, row in dl_df.iterrows():
        a, b = strip_ws(row[dl_fc]), strip_ws(row[dl_tc])
        if a != b and (b, a) not in arc_set:
            arc_set.add((a, b))

    tx_sheets = ['Length', 'MaxBuiltCapacity', 'Lifetime',
                 'InitialCapacity', 'MaxInstallCapacityRaw']
    for sheet in tx_sheets:
        try:
            _, df = _read_empire_df(out_dir / 'Transmission.xlsx', sheet)
            sfc, stc = df.columns[0], df.columns[1]
            for _, row in df.iterrows():
                a, b = strip_ws(row[sfc]), strip_ws(row[stc])
                if (a, b) not in arc_set:
                    print(f'  [ASSERT FAIL] Transmission/{sheet}: '
                          f'({a}, {b}) not in BidirectionalArc')
                    violations += 1
        except Exception as e:
            print(f'  [WARN] Could not check Transmission/{sheet}: {e}')

    # ── 2: No NO/NO1-NO5 column names in ScenarioData CSVs ──────────────────
    scenario_dir = out_dir / 'ScenarioData'
    bad_cols = set(NO_ZONES) | {'NO'}
    if scenario_dir.exists():
        for csv_file in sorted(scenario_dir.glob('*.csv')):
            try:
                cols = pd.read_csv(csv_file, nrows=0).columns.tolist()
                found = [c for c in cols if c in bad_cols]
                if found:
                    print(f'  [ASSERT FAIL] ScenarioData/{csv_file.name}: '
                          f'column(s) {found} must be named Norway')
                    violations += 1
            except Exception as e:
                print(f'  [WARN] Could not check {csv_file.name}: {e}')

    # ── 3: GeneratorsOfNode(Norway) covers all ScenarioData generator columns ─
    _, gon_df = _read_empire_df(out_dir / 'Sets.xlsx', 'GeneratorsOfNode')
    gnc, ggc = gon_df.columns[0], gon_df.columns[1]
    norway_gens = set(
        strip_ws(g) for g in gon_df[gon_df[gnc] == NORWAY][ggc].tolist()
    )
    # Map CSV file names to acceptable generator name(s) (after whitespace strip).
    # windoffshore.csv: noffshore datasets use 'Wind offshore grounded+floating'
    # while yesoffshore datasets use separate 'Wind offshore grounded' /
    # 'Wind offshore floating' generators — accept any variant.
    csv_gen_map = {
        'solar.csv':         ['Solar'],
        'windonshore.csv':   ['Windonshore'],
        'windoffshore.csv':  ['Windoffshoregroundedfloating', 'Windoffshoregrounded',
                              'Windoffshorefloating'],
        'hydroror.csv':      ['Hydrorun-of-the-river'],
    }
    if scenario_dir.exists():
        for csv_name, gen_names in csv_gen_map.items():
            csv_path = scenario_dir / csv_name
            if not csv_path.exists():
                continue
            cols = pd.read_csv(csv_path, nrows=0).columns.tolist()
            if NORWAY in cols and not norway_gens.intersection(gen_names):
                _, gon_raw = _read_empire_df(out_dir / 'Sets.xlsx', 'GeneratorsOfNode')
                norway_gens_raw = set(
                    g for g in gon_raw[gon_raw[gnc] == NORWAY][ggc].tolist()
                )
                if not any(strip_ws(g) in gen_names for g in norway_gens_raw):
                    print(f'  [ASSERT FAIL] GeneratorsOfNode: Norway missing '
                          f'any of {gen_names!r} (required by {csv_name})')
                    violations += 1

    if violations == 0:
        print('\n[ASSERT OK] All consistency checks passed.')
    else:
        raise AssertionError(
            f'{violations} consistency violation(s) in output — see above.')


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def process_one(input_folder, output_folder):
    inp = Path(input_folder)
    out = Path(output_folder)

    if not inp.exists():
        print(f'ERROR: input folder not found: {inp}')
        sys.exit(1)

    out.mkdir(parents=True, exist_ok=True)
    print(f'\n{"="*70}')
    print(f'Input:  {inp}')
    print(f'Output: {out}')
    print(f'{"="*70}')

    # Generator.xlsx
    gen_src = inp / 'Generator.xlsx'
    gen_dst_wb = openpyxl.Workbook()
    gen_dst_wb.remove(gen_dst_wb.active)
    preprocess_generator(gen_src, gen_dst_wb)
    gen_dst_wb.save(out / 'Generator.xlsx')

    # Node.xlsx
    node_dst_wb = openpyxl.Workbook()
    node_dst_wb.remove(node_dst_wb.active)
    preprocess_node(inp / 'Node.xlsx', node_dst_wb)
    node_dst_wb.save(out / 'Node.xlsx')

    # Storage.xlsx
    sto_dst_wb = openpyxl.Workbook()
    sto_dst_wb.remove(sto_dst_wb.active)
    preprocess_storage(inp / 'Storage.xlsx', sto_dst_wb)
    sto_dst_wb.save(out / 'Storage.xlsx')

    # Transmission.xlsx — returns canonical arc set for DirectionalLines ordering
    tx_dst_wb = openpyxl.Workbook()
    tx_dst_wb.remove(tx_dst_wb.active)
    canonical_arcs = preprocess_transmission(inp / 'Transmission.xlsx', tx_dst_wb)
    tx_dst_wb.save(out / 'Transmission.xlsx')

    # Sets.xlsx
    sets_dst_wb = openpyxl.Workbook()
    sets_dst_wb.remove(sets_dst_wb.active)
    preprocess_sets(inp / 'Sets.xlsx', sets_dst_wb, canonical_arcs=canonical_arcs)
    sets_dst_wb.save(out / 'Sets.xlsx')

    # General.xlsx — copy unchanged
    shutil.copy2(inp / 'General.xlsx', out / 'General.xlsx')
    print('  General.xlsx: copied')

    # ScenarioData
    scenario_src = inp / 'ScenarioData'
    if scenario_src.exists():
        preprocess_scenario_data(scenario_src, out / 'ScenarioData', gen_src)
    else:
        print('  ScenarioData/: not found, skipping')

    print(f'\n[OK] Output written to {out}')
    _assert_consistency(out)


def _read_empire_folder_from_config():
    """Return the empire_folder value from config.yaml, stripping any _Norway suffix."""
    try:
        import yaml
        config_path = ROOT / 'config.yaml'
        with open(config_path, encoding='utf-8') as f:
            cfg = yaml.safe_load(f)
        folder = cfg.get('empire_folder', '').strip()
        # Strip the _Norway suffix so this script is idempotent
        if folder.endswith('_Norway'):
            folder = folder[: -len('_Norway')]
        return folder
    except Exception as e:
        print(f'  WARNING: could not read config.yaml ({e})')
        return ''


def main():
    parser = argparse.ArgumentParser(
        description='Merge NO1-NO5 into a single Norway node in EMPIRE reference datasets'
    )
    parser.add_argument('input_folder', nargs='?',
                        help='EMPIRE reference dataset folder to process '
                             '(default: empire_folder from config.yaml)')
    parser.add_argument('output_folder', nargs='?',
                        help='Output folder for the preprocessed dataset '
                             '(default: <input_folder>_Norway)')
    parser.add_argument('--all', action='store_true',
                        help='Process all four standard reference datasets')
    args = parser.parse_args()

    empire_dir = ROOT / 'data' / 'EMPIRE'

    if args.all:
        for folder in ALL_REFERENCE_FOLDERS:
            process_one(empire_dir / folder, empire_dir / (folder + '_Norway'))

    elif args.input_folder:
        inp = Path(args.input_folder)
        # If a bare folder name is given (no path separators) assume it lives
        # under data/EMPIRE/
        if not inp.is_absolute() and '/' not in args.input_folder and '\\' not in args.input_folder:
            inp = empire_dir / args.input_folder
        out = Path(args.output_folder) if args.output_folder else Path(str(inp) + '_Norway')
        process_one(inp, out)

    else:
        # Fall back to config.yaml
        folder = _read_empire_folder_from_config()
        if not folder:
            parser.print_help()
            print('\nAvailable reference folders:')
            for f in ALL_REFERENCE_FOLDERS:
                print(f'  {f}')
            sys.exit(1)
        inp = empire_dir / folder
        out = empire_dir / (folder + '_Norway')
        print(f'Using empire_folder from config.yaml: {folder}')
        process_one(inp, out)


if __name__ == '__main__':
    main()
