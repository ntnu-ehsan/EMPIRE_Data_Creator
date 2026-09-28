"""
Generate Generation.xlsx for EMPIRE with NUTS2 resolution for selected countries.

Replaces selected country-level nodes with their NUTS2 regions while keeping all
other European countries as single nodes.

Usage:
    python generate_generation_nuts2.py <sets_file> <iamc_file> <scenario> <output_file> [--reference <Generator.xlsx>]

Example:
    python generate_generation_nuts2.py Sets_NUTS2_DK_DE.xlsx ^
        data/IAMC/NorthSea/Denmark_Germany.csv "Transmission v1.0" ^
        Generator_NUTS2_DK_DE.xlsx
"""

import sys
import argparse
from pathlib import Path
import pandas as pd
import numpy as np

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

try:
    import pyam
except ImportError:
    print("ERROR: pyam-iamc not installed.  Run: pip install pyam-iamc")
    sys.exit(1)


# =============================================================================
# NUTS2 configuration — loaded from config.yaml in the project root.
# Edit config.yaml to change which countries are represented at NUTS2 level.
# =============================================================================
from config_loader import load_nuts2_config as _load_nuts2_config
_cfg = _load_nuts2_config()
NUTS2_COUNTRY_PREFIXES = _cfg['NUTS2_COUNTRY_PREFIXES']
NUTS2_COUNTRY_TO_EMPIRE_NODE = _cfg['NUTS2_COUNTRY_TO_EMPIRE_NODE']
NUTS2_REGIONS_BY_COUNTRY = _cfg['NUTS2_REGIONS_BY_COUNTRY']
ALL_NUTS2 = _cfg['ALL_NUTS2']
ALL_NUTS2_SET = _cfg['ALL_NUTS2_SET']
REPLACED_COUNTRY_NODES = _cfg['REPLACED_COUNTRY_NODES']


# =============================================================================
# Node name normalisation.  The reference Generator.xlsx labels some nodes with
# legacy no-space variants (e.g. "GreatBrit.", "CzechR", "DoggerBank") while the
# generated Sets.xlsx Nodes sheet uses the canonical spaced names.  Apply this
# mapping to every node read from the reference so that (a) replaced-country
# rows are correctly removed and (b) kept rows match the Sets node set.  Mirrors
# NODE_NAME_FIXES in generate_sets_nuts2.py.
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
    # Trinity's Nodes sheet spells these with a stray space (see
    # generate_sets_nuts2.py) — keep both forms on the same name.
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


# =============================================================================
# IAMC → EMPIRE generator mapping
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
    'Capacity|Electricity|Hydro|Pumped Storage': None,        # Storage
    'Capacity|Electricity|Nuclear': 'Nuclear',
    'Capacity|Electricity|Ocean': 'Wave',
    'Capacity|Electricity|Oil|CCGT|w/o CCS': 'Oil',
    'Capacity|Electricity|Oil|Steam Turbine|w/o CCS': 'Oil',
    'Capacity|Electricity|Oil|w/o CCS': 'Oil',
    'Capacity|Electricity|Other': None,                   # No mapping
    'Capacity|Electricity|Solar|CSP': 'Solar',
    'Capacity|Electricity|Solar|PV': 'Solar',
    'Capacity|Electricity|Waste|w/o CCS': 'Waste',
    'Capacity|Electricity|Wind': None,                    # legacy aggregate key
    'Capacity|Electricity|Wind|Offshore': 'Wind offshore grounded',
    'Capacity|Electricity|Wind|Onshore': 'Wind onshore',
}

# IAMC variable used as the fallback weight when a country has reference
# capacity for a technology that the IAMC file does not report for that country.
#   * Waste            -> distribute by the region's biomass capacity
#   * Hydro run-of-river -> distribute by the region's total hydro capacity
#   * anything else    -> distribute by the region's total installed capacity
WASTE_WEIGHT_VARIABLE = 'Capacity|Electricity|Biomass|w/o CCS'
HYDRO_WEIGHT_VARIABLES = (
    'Capacity|Electricity|Hydro',
    'Capacity|Electricity|Hydro|Reservoir',
    'Capacity|Electricity|Hydro|Run of River',
    'Capacity|Electricity|Hydro|Pumped Storage',
)

# Renewable technologies whose MaxInstalledCapacity (resource-potential) ceiling
# is *distributed* across a country's NUTS2 regions in proportion to the region's
# IAMC capacity for that technology, so the national ceiling is conserved
# (Σ regions == country reference) instead of replicated to every region. All
# other technologies are national-level limits (see MaxInstalledCapacityCountry
# in NUTS_national_limits_data_spec.md) and keep the replicated country value in
# the nodal sheet as a generous, non-binding headroom.
#
# Maps the tech name as it appears in the reference MaxInstalledCapacity sheet
# (EMPIRE short names, e.g. 'Wind_onshr') to the generator name used as the IAMC
# capacity weight (the spaced names produced by CAPACITY_TO_GENERATOR).
MAXINSTALLED_WEIGHTED_TECHS = {
    'Wind_onshr': 'Wind onshore',
    'Solar': 'Solar',
}

# Hydro MaxInstalledCapacity is likewise *distributed* instead of replicated, but
# with its own rule: the national ceiling is split across the regions that hold a
# hydro generator in proportion to their IAMC installed hydro capacity, and each
# region's share is scaled by HYDRO_MAXINSTALLED_HEADROOM to leave expansion room
# beyond today's fleet.  Replicating the national ceiling (the previous behaviour)
# let every single NUTS3 region expand to the whole country's hydro potential,
# which is physically impossible and inconsistent with the per-region energy cap
# in HydroGenMaxAnnualProduction (generate_node.py).
MAXINSTALLED_HYDRO_TECHS = {
    'Hydro_reg': 'Hydro regulated',
    'Hydro_ror': 'Hydro run-of-the-river',
}
HYDRO_MAXINSTALLED_HEADROOM = 1.5

# Offshore/wave MaxInstalledCapacity: a national-only ceiling would let ANY NUTS2/3
# region (coastal or not) build offshore wind/wave, which is physically wrong. So
# these stay distributed too, but restricted to the regions the pipeline has
# already wired an offshore/wave generator slot to (GeneratorsOfNode) — the
# pipeline's operational proxy for "coastal", since that's what
# generate_sets_nuts2.py's offshore-reconnection logic encodes. Weighted by IAMC
# installed capacity where available, else split equally over the eligible
# regions (no headroom factor). If no region is eligible the ceiling is dropped
# with a warning, same as the hydro no-eligible-region case.
# TODO: this is a proxy for real coastal eligibility, not a geographic
# coastal-potential dataset (cf. data/Renewable/Wind.csv for onshore wind /
# solar). Revisit if/when a proper offshore/wave resource-potential input for
# Spain's NUTS3 regions becomes available.
MAXINSTALLED_COASTAL_TECHS = {
    'Wave': 'Wave',
    'Wind_offshr_floating': 'Wind offshore floating',
    'Wind_offshr_grounded': 'Wind offshore grounded',
}

# =============================================================================
# Spain-specific resource-potential weights (data/Renewable/).
# These override the default IAMC-installed-capacity weights for ES regions when
# distributing the national MaxInstalledCapacity ceiling across NUTS3, giving a
# better spatial split of the resource potential.  National totals are still
# conserved — the data is used only as relative weights.
#   * Wind onshore : Wind.csv ref_MW per NUTS3 (direct, already NUTS3).
#   * Solar        : Solar.xlsx Capacity_GWe per NUTS2, split to its NUTS3
#                    children in proportion to IAMC installed solar.
# Only used when 'ES' is a modeled NUTS2/3 country; other configs are untouched.
# =============================================================================
RENEWABLE_DIR = Path(__file__).resolve().parent.parent / 'data' / 'Renewable'


def load_spain_potentials(renewable_dir=RENEWABLE_DIR):
    """Return (wind_ref_mw, solar_gwe_n2) from data/Renewable.

    wind_ref_mw  : {nuts3_code: onshore-wind potential in MW}  (Wind.csv)
    solar_gwe_n2 : {nuts2_code: solar potential in GW}         (Solar.xlsx, ES)

    Returns ({}, {}) for any file that is absent so callers can fall back to the
    default IAMC-capacity weights.
    """
    wind_ref_mw: dict[str, float] = {}
    solar_gwe_n2: dict[str, float] = {}

    wind_path = renewable_dir / 'Wind.csv'
    if wind_path.exists():
        w = pd.read_csv(wind_path)
        for _, r in w.iterrows():
            mw = pd.to_numeric(r.get('ref_MW'), errors='coerce')
            if pd.notna(mw) and mw > 0:
                wind_ref_mw[str(r['NUTS_ID'])] = float(mw)
    else:
        print(f"  [MaxInstalledCapacity] note: {wind_path} not found — "
              "ES wind keeps IAMC-capacity weights.")

    solar_path = renewable_dir / 'Solar.xlsx'
    if solar_path.exists():
        s = pd.read_excel(solar_path)
        for _, r in s.iterrows():
            if str(r.get('Code')) == 'ES':
                gwe = pd.to_numeric(r.get('Capacity_GWe'), errors='coerce')
                if pd.notna(gwe) and gwe > 0:
                    solar_gwe_n2[str(r['NUTS2 NEW'])] = float(gwe)
    else:
        print(f"  [MaxInstalledCapacity] note: {solar_path} not found — "
              "ES solar keeps IAMC-capacity weights.")

    return wind_ref_mw, solar_gwe_n2


def spain_solar_nuts3_weights(regions, solar_gwe_n2, cap_lookup):
    """Split each NUTS2 solar potential (Capacity_GWe) across its modeled NUTS3
    children in proportion to IAMC installed solar capacity.

    Returns {nuts3: weight} where weight = parent_potential * child_installed_share.
    A NUTS2 with potential but zero installed solar in its children is split
    equally; children of a NUTS2 with no potential get weight 0.
    """
    children_by_parent: dict[str, list[str]] = {}
    for rg in regions:
        children_by_parent.setdefault(str(rg)[:4], []).append(rg)

    weights: dict[str, float] = {}
    for parent, children in children_by_parent.items():
        potential = solar_gwe_n2.get(parent, 0.0)
        if potential <= 0:
            for rg in children:
                weights[rg] = 0.0
            continue
        installed = {rg: cap_lookup.get((rg, 'Solar'), 0.0) for rg in children}
        isum = sum(installed.values())
        if isum > 0:
            for rg in children:
                weights[rg] = potential * installed[rg] / isum
        else:
            equal = potential / len(children)
            for rg in children:
                weights[rg] = equal
    return weights


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
    data.columns = headers
    return data.reset_index(drop=True)


def read_sets_sheet_simple(sets_file, sheet_name):
    """Read a simple Sets.xlsx sheet (no metadata rows)."""
    return pd.read_excel(sets_file, sheet_name=sheet_name)


# =============================================================================
# Sheet builders
# =============================================================================

def make_ref_initial_cap(gen_of_node_df, iamc_file, scenario, ref_file):
    """RefInitialCap – installed capacity per generator per node.

    The EMPIRE reference national total is conserved.  For each modeled country
    (UK/NO/DK/NL) and generator, the reference country total is *redistributed*
    across that country's NUTS2 regions using the IAMC file only as spatial
    weights — so the national total per technology is identical before and after
    the NUTS2 split.

    Weight priority for a (country, generator):
      1. the region's IAMC capacity for that generator;
      2. fallback when the IAMC file has no capacity for that generator in that
         country (warning printed):
           * Waste              -> region biomass capacity
           * Hydro run-of-river -> region total hydro capacity
           * any other tech     -> region total installed generation capacity
         (if the chosen fallback weight is also zero everywhere it falls back to
          total installed capacity, then to an equal split).

    Non-NUTS2 nodes keep their reference Generator.xlsx values unchanged.
    """
    from iamc_utils import load_iamc

    node_to_prefix = {v: k for k, v in NUTS2_COUNTRY_TO_EMPIRE_NODE.items()}

    # ---- Reference rows ----
    ref_raw = pd.read_excel(ref_file, sheet_name='RefInitialCap', header=2)
    ref_df = ref_raw.iloc[:, :3].copy()
    ref_df.columns = [
        'Node', 'GeneratorTechnology',
        'generatoReferenceInitialCapacity in MW',
    ]
    # Normalise legacy no-space node names to match the Sets node set.
    ref_df['Node'] = ref_df['Node'].map(fix_node_name)
    ref_df['generatoReferenceInitialCapacity in MW'] = pd.to_numeric(
        ref_df['generatoReferenceInitialCapacity in MW'], errors='coerce').fillna(0)

    # Non-modeled nodes pass through unchanged.
    ref_non_de_dk = (
        ref_df[~ref_df['Node'].isin(REPLACED_COUNTRY_NODES)]
        .copy()
        .reset_index(drop=True)
    )

    # GeneratorsOfNode is the single source of truth for which (node, generator)
    # pairs exist. Drop pass-through reference rows whose pair is absent from it
    # (e.g. offshore-wind at a node that has no windoffshore.csv profile and was
    # therefore removed from GeneratorsOfNode). genRefInitCap is indexed by
    # GeneratorsOfNode, so a stray pair here aborts Pyomo model construction.
    valid_pairs = set(zip(gen_of_node_df['Node'], gen_of_node_df['Generator']))
    pair_mask = [
        (n, g) in valid_pairs
        for n, g in zip(ref_non_de_dk['Node'],
                        ref_non_de_dk['GeneratorTechnology'])
    ]
    n_dropped = len(pair_mask) - sum(pair_mask)
    if n_dropped:
        dropped_pairs = sorted({
            (n, g)
            for (n, g), keep in zip(
                zip(ref_non_de_dk['Node'], ref_non_de_dk['GeneratorTechnology']),
                pair_mask)
            if not keep
        })
        ref_non_de_dk = ref_non_de_dk[pair_mask].reset_index(drop=True)
        print(f"  [RefInitialCap] Dropped {n_dropped} reference row(s) whose "
              f"(node, generator) pair is not in GeneratorsOfNode: "
              f"{', '.join(f'{n}/{g}' for n, g in dropped_pairs)}")

    # Reference national total per (country prefix, generator) for modeled nodes.
    ref_country_total: dict[tuple[str, str], float] = {}
    for _, r in ref_df[ref_df['Node'].isin(REPLACED_COUNTRY_NODES)].iterrows():
        prefix = node_to_prefix.get(r['Node'])
        if prefix is None:
            continue
        key = (prefix, r['GeneratorTechnology'])
        ref_country_total[key] = (
            ref_country_total.get(key, 0.0)
            + r['generatoReferenceInitialCapacity in MW']
        )

    # ---- IAMC weight lookups for NUTS2 regions ----
    df_iamc = load_iamc(iamc_file)
    df_filt = df_iamc.filter(scenario=scenario)

    cap_lookup: dict[tuple[str, str], float] = {}   # (region, gen) -> MW
    waste_weight: dict[str, float] = {}             # region -> biomass MW
    hydro_weight: dict[str, float] = {}             # region -> total hydro MW
    total_weight: dict[str, float] = {}             # region -> total installed MW
    for _, r in df_filt.data.iterrows():
        region = r['region']
        if '>' in str(region) or region not in ALL_NUTS2_SET:
            continue
        var = r['variable']
        unit = r.get('unit', 'GW')
        val_mw = r['value'] * 1000 if unit == 'GW' else r['value']  # → MW
        gen = CAPACITY_TO_GENERATOR.get(var)
        if gen is not None:
            cap_lookup[(region, gen)] = cap_lookup.get((region, gen), 0.0) + val_mw
            total_weight[region] = total_weight.get(region, 0.0) + val_mw
        if var == WASTE_WEIGHT_VARIABLE:
            waste_weight[region] = waste_weight.get(region, 0.0) + val_mw
        if var in HYDRO_WEIGHT_VARIABLES:
            hydro_weight[region] = hydro_weight.get(region, 0.0) + val_mw

    # ---- Group NUTS2 generator slots (from GeneratorsOfNode) by country ----
    slots: dict[tuple[str, str], list[str]] = {}
    for _, row in gen_of_node_df.iterrows():
        node, gen = row['Node'], row['Generator']
        if node in ALL_NUTS2_SET:
            slots.setdefault((str(node)[:2], gen), []).append(node)

    # ---- Distribute each country/generator reference total across its NUTS2 ----
    nuts2_rows = []
    for (prefix, gen), regions in slots.items():
        country_node = NUTS2_COUNTRY_TO_EMPIRE_NODE.get(prefix, prefix)
        ref_total = ref_country_total.get((prefix, gen), 0.0)
        if ref_total <= 0:
            for region in regions:
                nuts2_rows.append({
                    'Node': region, 'GeneratorTechnology': gen,
                    'generatoReferenceInitialCapacity in MW': 0.0,
                })
            continue

        # Primary weights: IAMC capacity for this generator.
        weights = {region: cap_lookup.get((region, gen), 0.0) for region in regions}
        wsum = sum(weights.values())
        source = 'IAMC capacity'

        if wsum <= 0:
            # Fallback: this country has no IAMC capacity for this generator.
            if gen == 'Waste':
                weights = {rg: waste_weight.get(rg, 0.0) for rg in regions}
                source = 'biomass capacity'
            elif gen == 'Hydro run-of-the-river':
                weights = {rg: hydro_weight.get(rg, 0.0) for rg in regions}
                source = 'total hydro capacity'
            else:
                weights = {rg: total_weight.get(rg, 0.0) for rg in regions}
                source = 'total installed capacity'
            wsum = sum(weights.values())

            if wsum <= 0 and source != 'total installed capacity':
                # Chosen fallback weight is zero everywhere -> total capacity.
                weights = {rg: total_weight.get(rg, 0.0) for rg in regions}
                wsum = sum(weights.values())
                source = 'total installed capacity'
            if wsum <= 0:
                # Last resort: equal split.
                weights = {rg: 1.0 for rg in regions}
                wsum = float(len(regions))
                source = 'equal split'

            print(f"  [RefInitialCap][WARNING] {country_node} '{gen}': no IAMC "
                  f"capacity for this technology — distributing the {ref_total:,.1f} MW "
                  f"reference total across {len(regions)} NUTS2 region(s) by {source}.")

        for region in regions:
            cap = ref_total * weights[region] / wsum
            nuts2_rows.append({
                'Node': region, 'GeneratorTechnology': gen,
                'generatoReferenceInitialCapacity in MW': round(cap, 3),
            })

    # ---- Warn about any reference capacity that has no NUTS2 slot to land in ----
    for (prefix, gen), ref_total in ref_country_total.items():
        if ref_total > 0 and (prefix, gen) not in slots:
            country_node = NUTS2_COUNTRY_TO_EMPIRE_NODE.get(prefix, prefix)
            print(f"  [RefInitialCap][WARNING] {country_node} '{gen}': {ref_total:,.1f} MW "
                  f"of reference capacity could not be placed — no NUTS2 node has this "
                  f"generator in GeneratorsOfNode. National total will be short by this amount.")

    nuts2_df = pd.DataFrame(nuts2_rows)
    return pd.concat([ref_non_de_dk, nuts2_df], ignore_index=True)


def _read_ref_period_sheet(ref_file, sheet_name, value_col):
    """Read a Node x Tech x Period x MW sheet (MaxBuiltCapacity /
    MinBuiltCapacity layout: 2 metadata rows + header row, then data) from the
    reference workbook. Returns an empty frame (with a note) if the reference
    doesn't have that sheet at all — MinBuiltCapacity is absent from some
    older reference layouts.
    """
    try:
        ref_raw = pd.read_excel(ref_file, sheet_name=sheet_name, header=None)
    except ValueError:
        print(f"  [{sheet_name}] note: reference has no '{sheet_name}' sheet "
              "— treated as empty.")
        return pd.DataFrame(columns=['Node', 'GeneratorTechnology', 'Period', value_col])

    ref_data = ref_raw.iloc[3:].copy()
    ref_data.columns = ['Node', 'GeneratorTechnology', 'Period', value_col]
    ref_data[value_col] = pd.to_numeric(ref_data[value_col], errors='coerce').fillna(0)
    ref_data['Node'] = ref_data['Node'].map(fix_node_name)
    return ref_data.dropna(subset=['Node', 'GeneratorTechnology', 'Period'])


def make_yearly_availability(all_nodes, nodes_by_prefix, ref_file,
                             gen_of_node_df):
    '''Return a formatted optional YearlyAvailability sheet.

    Rows for nodes that remain at country level are preserved. A row belonging
    to a country replaced by NUTS regions is replicated to every replacement
    region because availability is a dimensionless multiplier, not an
    extensive quantity that should be divided between regions.

    Replication is restricted to regions that actually hold the generator in
    GeneratorsOfNode. EMPIRE declares genYearlyAvailability over
    ``Param(model.GeneratorsOfNode, model.Period)``, so a (node, generator) pair
    that is not in that set aborts model construction with a Pyomo KeyError —
    the same reason make_ref_initial_cap filters against it.
    '''
    raw = pd.read_excel(ref_file, sheet_name='YearlyAvailability', header=None)
    if raw.shape[0] < 3 or raw.shape[1] < 4:
        raise ValueError(
            'YearlyAvailability must contain two metadata rows, a header row, '
            'and four columns'
        )

    data = raw.iloc[3:, :4].copy()
    data.columns = ['Node', 'GeneratorTechnology', 'Period', 'Value']
    data['Node'] = data['Node'].map(fix_node_name)
    data = data.dropna(
        subset=['Node', 'GeneratorTechnology', 'Period', 'Value']
    )

    valid_nodes = set(all_nodes)
    valid_pairs = set(zip(gen_of_node_df['Node'], gen_of_node_df['Generator']))
    replaced_nodes = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())

    passthrough = data[
        data['Node'].isin(valid_nodes) & ~data['Node'].isin(replaced_nodes)
    ]
    output_rows = []
    dropped_passthrough = set()
    for _, row in passthrough.iterrows():
        pair = (row['Node'], row['GeneratorTechnology'])
        if pair in valid_pairs:
            output_rows.append(row.tolist())
        else:
            dropped_passthrough.add(pair)
    if dropped_passthrough:
        print(
            f'  [YearlyAvailability] Dropped {len(dropped_passthrough)} '
            'pass-through pair(s) not in GeneratorsOfNode: '
            + ', '.join(f'{n}/{g}' for n, g in sorted(dropped_passthrough))
        )

    for prefix, regions in nodes_by_prefix.items():
        country_node = NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]
        country_rows = data[data['Node'] == country_node]
        if country_rows.empty:
            continue
        if not regions:
            print(
                f'  [YearlyAvailability][WARNING] {country_node}: no '
                'replacement regions; dropping its availability rows.'
            )
            continue
        skipped: dict[str, list[str]] = {}
        for _, row in country_rows.iterrows():
            tech = row['GeneratorTechnology']
            eligible = [rg for rg in regions if (rg, tech) in valid_pairs]
            if not eligible:
                skipped.setdefault(tech, list(regions))
                continue
            missing = [rg for rg in regions if rg not in eligible]
            if missing:
                skipped.setdefault(tech, missing)
            for region in eligible:
                output_rows.append([
                    region,
                    tech,
                    row['Period'],
                    row['Value'],
                ])
        for tech, regs in sorted(skipped.items()):
            print(
                f'  [YearlyAvailability] {country_node}/{tech}: skipped '
                f'{len(regs)} region(s) without that generator in '
                f'GeneratorsOfNode: {", ".join(sorted(regs))}'
            )

    output_data = pd.DataFrame(output_rows, columns=range(4))
    metadata = raw.iloc[:3, :4].copy()
    metadata.columns = range(4)
    return pd.concat([metadata, output_data], ignore_index=True)


def make_max_built_capacity(nodes, ref_file):
    """MaxBuiltCapacity per node × tech × period.

    The reference sheet is a sparse *exception list*: every row it carries is a
    0, naming a (node, tech) pair that may not be built at all (Austria/Germany/
    Denmark Nuclear, Portugal/France Coal, ...).  Every combination it omits is
    unconstrained, and EMPIRE says so itself -- genMaxBuiltCap is declared with
    ``default=500000.0``, matching the sheet's own documented default.  So only
    reference rows are emitted; filling the node × tech × period cross product
    with 0 for the gaps would turn a handful of build bans into a blanket ban on
    all generation investment.

    - Non-NUTS2 nodes: reference rows passed through unchanged.
    - NUTS2 nodes: no rows are emitted — build-rate limits for NUTS-
      disaggregated countries are now national-level (see
      make_max_built_capacity_country).
    """
    value_col = 'generatorMaxBuildCapacity in MW'
    ref_data = _read_ref_period_sheet(ref_file, 'MaxBuiltCapacity', value_col)
    keep = set(nodes) - ALL_NUTS2_SET

    rows = [
        {
            'Node': r['Node'],
            'GeneratorTechnology': r['GeneratorTechnology'],
            'Period': r['Period'],
            value_col: round(r[value_col], 3),
        }
        for _, r in ref_data.iterrows()
        if r['Node'] in keep
    ]
    return pd.DataFrame(rows)


def make_max_built_capacity_country(ref_file):
    """MaxBuiltCapacityCountry — national build-rate ceiling per period.

    Build rate is inherently national (supply chain/permitting capacity), so
    every technology for every NUTS-disaggregated country's own reference row
    moves here unchanged (no more country_value / n_regions split). See
    NUTS_national_limits_data_spec.md.
    """
    value_col = 'generatorMaxBuildCapacity in MW'
    ref_data = _read_ref_period_sheet(ref_file, 'MaxBuiltCapacity', value_col)
    country_nodes = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())
    rows = ref_data[ref_data['Node'].isin(country_nodes)].rename(columns={'Node': 'Country'})
    return rows[['Country', 'GeneratorTechnology', 'Period', value_col]].reset_index(drop=True)


def _drop_min_above_max(min_data, ref_file, min_col):
    """Drop MinBuiltCapacity rows the reference's own MaxBuiltCapacity forbids.

    The two sheets are maintained independently and can contradict each other:
    NECPEssentials bans Dutch nuclear (MaxBuiltCapacity 0 in every period) while
    mandating 1000 MW of it in P3/P4.  EMPIRE turns such a pair into
    ``genInvCap <= 0`` and ``genInvCap >= 1000`` on one variable, and the model
    is infeasible with nothing to go on but an IIS naming two constraints.

    The build ban wins, because every value the reference MaxBuiltCapacity sheet
    carries is a deliberate 0 — a list of what may not be built at all.  Rows
    with no explicit Max are left alone; EMPIRE's 500 000 MW default covers
    them, which is exactly what make_max_built_capacity relies on.
    """
    max_col = 'generatorMaxBuildCapacity in MW'
    max_data = _read_ref_period_sheet(ref_file, 'MaxBuiltCapacity', max_col)
    max_lookup = {
        (r['Node'], r['GeneratorTechnology'], r['Period']): r[max_col]
        for _, r in max_data.iterrows()
    }

    keep = []
    for _, r in min_data.iterrows():
        cap = max_lookup.get((r['Node'], r['GeneratorTechnology'], r['Period']))
        if cap is not None and cap < r[min_col]:
            print(f"  [MinBuiltCapacity][WARNING] {r['Node']} "
                  f"'{r['GeneratorTechnology']}' P{r['Period']}: mandate of "
                  f"{r[min_col]:,.0f} MW dropped — the reference "
                  f"MaxBuiltCapacity for the same node/tech/period is "
                  f"{cap:,.0f} MW, so the two sheets contradict each other and "
                  "EMPIRE would be infeasible. The build ban wins; correct the "
                  "reference workbook if the mandate is the intended entry.")
            keep.append(False)
        else:
            keep.append(True)
    return min_data[keep]


def make_min_built_capacity(ref_file):
    """MinBuiltCapacity (nodal) — restored pass-through for non-disaggregated
    countries.

    Previously silently dropped by this pipeline (not in copy_sheets, never
    rebuilt), even though the reference workbook carries real commissioning
    mandates (e.g. France Nuclear P4/P5). Rows belonging to a NUTS-
    disaggregated country's own node are dropped here and routed to
    MinBuiltCapacityCountry instead (make_min_built_capacity_country).
    """
    value_col = 'generatorMinBuildCapacity in MW'
    ref_data = _read_ref_period_sheet(ref_file, 'MinBuiltCapacity', value_col)
    country_nodes = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())
    rows = ref_data[~ref_data['Node'].isin(country_nodes)]
    rows = _drop_min_above_max(rows, ref_file, value_col)
    return rows[['Node', 'GeneratorTechnology', 'Period', value_col]].reset_index(drop=True)


def make_min_built_capacity_country(ref_file):
    """MinBuiltCapacityCountry — national mandated build-out per period.

    Empty unless the reference workbook has an explicit MinBuiltCapacity row
    for a NUTS-disaggregated country's own node (none exist for Spain today —
    the sheet is still always emitted so a future mandate can be added
    without a pipeline change). See NUTS_national_limits_data_spec.md.
    """
    value_col = 'generatorMinBuildCapacity in MW'
    ref_data = _read_ref_period_sheet(ref_file, 'MinBuiltCapacity', value_col)
    country_nodes = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())
    rows = ref_data[ref_data['Node'].isin(country_nodes)]
    rows = _drop_min_above_max(rows, ref_file, value_col).rename(
        columns={'Node': 'Country'})
    return rows[['Country', 'GeneratorTechnology', 'Period', value_col]].reset_index(drop=True)


def _distribute_over_eligible_regions(alloc, regions_by_prefix, country_vals,
                                       techs_map, gen_of_node_slots, cap_lookup,
                                       headroom=1.0):
    """Distribute a national MaxInstalledCapacity ceiling across the NUTS2/3
    regions of a country that actually hold that generator (GeneratorsOfNode),
    weighted by IAMC installed capacity, falling back to an equal split over
    the eligible regions when the IAMC file has no capacity for that
    generator. The ceiling is dropped (with a warning) if no region of the
    country is eligible. Mutates ``alloc`` in place; every region of the
    country gets an explicit 0 entry so the caller never falls through to the
    replicated country value.
    """
    for prefix, regions in regions_by_prefix.items():
        vals = country_vals.get(prefix, {})
        country_node = NUTS2_COUNTRY_TO_EMPIRE_NODE.get(prefix, prefix)
        for tech, weight_gen in techs_map.items():
            for rg in regions:
                alloc[(rg, tech)] = 0.0

            country_val = vals.get(tech, 0)
            if not country_val or country_val <= 0:
                continue

            eligible = [rg for rg in regions if (rg, weight_gen) in gen_of_node_slots]
            if not eligible:
                print(f"  [MaxInstalledCapacity][WARNING] {country_node} '{tech}': "
                      f"{country_val:,.1f} MW ceiling dropped — no NUTS3 region "
                      f"holds a '{weight_gen}' generator.")
                continue

            weights = {rg: cap_lookup.get((rg, weight_gen), 0.0) for rg in eligible}
            wsum = sum(weights.values())
            if wsum > 0:
                source = f'IAMC installed {weight_gen}'
                if headroom != 1.0:
                    source += f' (x{headroom} headroom)'
            else:
                weights = {rg: 1.0 for rg in eligible}
                wsum = float(len(eligible))
                source = 'equal split (no IAMC capacity for this country)'
                if headroom != 1.0:
                    source += f' (x{headroom} headroom)'
                print(f"  [MaxInstalledCapacity][WARNING] {country_node} '{tech}': "
                      f"no IAMC capacity for '{weight_gen}' — splitting the "
                      f"{country_val:,.1f} MW ceiling equally over "
                      f"{len(eligible)} region(s).")

            for rg in eligible:
                alloc[(rg, tech)] = round(
                    country_val * weights[rg] / wsum * headroom, 3)

            print(f"  [MaxInstalledCapacity] {country_node} '{tech}': "
                  f"distributing the {country_val:,.1f} MW ceiling across "
                  f"{len(eligible)} NUTS3 region(s) by {source}.")


def make_max_installed_capacity(nodes, technologies, ref_file,
                                iamc_file, scenario, gen_of_node_df):
    """MaxInstalledCapacity – resource limits per node × technology.

    - Non-NUTS2 nodes: values from reference Generator.xlsx
    - NUTS2 nodes:
        * Technologies in ``MAXINSTALLED_WEIGHTED_TECHS`` (Wind onshore, Solar):
          the country reference ceiling is *distributed* across the country's
          NUTS2 regions in proportion to the region's IAMC capacity for that
          technology, so the national ceiling is conserved (Σ regions ==
          country reference) instead of replicated. When the IAMC file has no
          capacity for that technology in the country it falls back to the
          region's total installed IAMC capacity, then to an equal split.
        * Technologies in ``MAXINSTALLED_HYDRO_TECHS`` (Hydro regulated, Hydro
          run-of-the-river): the country reference ceiling is distributed over
          the regions that own that hydro generator in proportion to their IAMC
          installed hydro capacity, each share scaled by
          ``HYDRO_MAXINSTALLED_HEADROOM``. When the country has no IAMC capacity
          for that hydro technology the ceiling is split equally over those
          regions instead (no headroom factor — the equal split already spreads
          the full national ceiling).
        * Technologies in ``MAXINSTALLED_COASTAL_TECHS`` (Wave, Wind offshore
          floating, Wind offshore grounded): same eligible-region mechanism as
          hydro (see ``_distribute_over_eligible_regions``), no headroom factor.
        * All other technologies (Nuclear, Gas, Coal, CCS, Lignite, Bio, Oil,
          Geo, Waste, ...): these are national-level limits — see
          ``make_max_installed_capacity_country`` — and the country reference
          value is replicated to each NUTS2 region here only as a generous,
          non-binding nodal headroom (never the binding constraint).
    """
    ref_raw = pd.read_excel(ref_file, sheet_name='MaxInstalledCapacity', header=2)
    ref_df = ref_raw.iloc[:, :3].copy()
    ref_df.columns = [
        'Node', 'GeneratorTechnology',
        'generatorMaxInstallCapacity  in MW',
    ]
    ref_df = ref_df.dropna(subset=['Node', 'GeneratorTechnology'])
    # Normalise legacy no-space node names so replaced-country rows are removed
    # and the per-country lookup below matches the canonical empire_node names.
    ref_df['Node'] = ref_df['Node'].map(fix_node_name)
    val_col = 'generatorMaxInstallCapacity  in MW'

    # Filter rows whose technology is not in the Technology set
    valid_techs = set(technologies)
    invalid_mask = ~ref_df['GeneratorTechnology'].isin(valid_techs)
    if invalid_mask.any():
        dropped = sorted(ref_df.loc[invalid_mask, 'GeneratorTechnology'].unique())
        print(f'  [MaxInstalledCapacity] Dropping {invalid_mask.sum()} rows with '
              f'unknown tech names: {dropped}')
        ref_df = ref_df[~invalid_mask].copy()

    # Extract country-level reference values per technology
    country_vals: dict[str, dict[str, float]] = {
        prefix: {} for prefix in NUTS2_COUNTRY_PREFIXES
    }
    for _, row in ref_df.iterrows():
        node = row['Node']
        tech = row['GeneratorTechnology']
        val = row['generatorMaxInstallCapacity  in MW']
        for prefix in NUTS2_COUNTRY_PREFIXES:
            if node == NUTS2_COUNTRY_TO_EMPIRE_NODE[prefix]:
                country_vals[prefix][tech] = val

    # Keep non-NUTS2-country reference rows
    ref_non_de_dk = (
        ref_df[~ref_df['Node'].isin(REPLACED_COUNTRY_NODES)]
        .copy()
        .reset_index(drop=True)
    )

    # ---- IAMC capacity weights for NUTS2 regions (for the weighted techs) ----
    from iamc_utils import load_iamc
    df_iamc = load_iamc(iamc_file)
    df_filt = df_iamc.filter(scenario=scenario)

    cap_lookup: dict[tuple[str, str], float] = {}   # (region, gen) -> MW
    total_weight: dict[str, float] = {}             # region -> total installed MW
    for _, r in df_filt.data.iterrows():
        region = r['region']
        if '>' in str(region) or region not in ALL_NUTS2_SET:
            continue
        var = r['variable']
        unit = r.get('unit', 'GW')
        val_mw = r['value'] * 1000 if unit == 'GW' else r['value']  # → MW
        gen = CAPACITY_TO_GENERATOR.get(var)
        if gen is not None:
            cap_lookup[(region, gen)] = cap_lookup.get((region, gen), 0.0) + val_mw
            total_weight[region] = total_weight.get(region, 0.0) + val_mw

    # NUTS2 regions present in the Sets node list, grouped by country prefix.
    nuts2_nodes = sorted(set(nodes) & ALL_NUTS2_SET)
    regions_by_prefix: dict[str, list[str]] = {}
    for node in nuts2_nodes:
        regions_by_prefix.setdefault(str(node)[:2], []).append(node)

    # Spain resource-potential weights (override IAMC-capacity weights for ES).
    spain_wind_mw: dict[str, float] = {}
    spain_solar_gwe: dict[str, float] = {}
    if 'ES' in NUTS2_COUNTRY_PREFIXES:
        spain_wind_mw, spain_solar_gwe = load_spain_potentials()

    # Pre-compute the weighted allocation for the renewable techs so the national
    # ceiling is conserved across the country's regions.
    alloc: dict[tuple[str, str], float] = {}        # (node, tech) -> MW
    for prefix, regions in regions_by_prefix.items():
        vals = country_vals.get(prefix, {})
        for tech, weight_gen in MAXINSTALLED_WEIGHTED_TECHS.items():
            country_val = vals.get(tech, 0)
            if not country_val or country_val <= 0:
                continue  # nothing to distribute → falls through to value 0

            # Spain: prefer the resource-potential weights from data/Renewable.
            weights = None
            source = None
            if prefix == 'ES' and spain_wind_mw and tech == 'Wind_onshr':
                weights = {rg: spain_wind_mw.get(rg, 0.0) for rg in regions}
                source = 'Wind.csv ref_MW'
            elif prefix == 'ES' and spain_solar_gwe and tech == 'Solar':
                weights = spain_solar_nuts3_weights(
                    regions, spain_solar_gwe, cap_lookup)
                source = 'Solar.xlsx Capacity_GWe x IAMC installed-solar share'

            if weights is None or sum(weights.values()) <= 0:
                weights = {rg: cap_lookup.get((rg, weight_gen), 0.0) for rg in regions}
                source = 'IAMC capacity'
            wsum = sum(weights.values())
            if wsum <= 0:
                weights = {rg: total_weight.get(rg, 0.0) for rg in regions}
                wsum = sum(weights.values())
                source = 'total installed capacity'
            if wsum <= 0:
                weights = {rg: 1.0 for rg in regions}
                wsum = float(len(regions))
                source = 'equal split'
            country_node = NUTS2_COUNTRY_TO_EMPIRE_NODE.get(prefix, prefix)
            if source.startswith(('Wind.csv', 'Solar.xlsx')):
                print(f"  [MaxInstalledCapacity] {country_node} '{tech}': "
                      f"distributing the {country_val:,.1f} MW ceiling across "
                      f"{len(regions)} NUTS3 region(s) by {source}.")
            elif source not in ('IAMC capacity',):
                print(f"  [MaxInstalledCapacity][WARNING] {country_node} '{tech}': "
                      f"no IAMC capacity for this technology — distributing the "
                      f"{country_val:,.1f} MW ceiling across {len(regions)} NUTS2 "
                      f"region(s) by {source}.")

            for rg in regions:
                alloc[(rg, tech)] = round(country_val * weights[rg] / wsum, 3)

    # ---- Hydro / offshore wind / wave: distribute the national ceiling over
    # eligible regions, never replicate it. Only the regions that actually hold
    # that generator (GeneratorsOfNode) are eligible — for hydro that means the
    # ceiling isn't stranded in regions that cannot build hydro; for offshore
    # wind/wave it's the pipeline's proxy for "coastal" (see
    # MAXINSTALLED_COASTAL_TECHS). Every region of the country gets an explicit
    # entry (0 when it has no slot) so the lookup below can never fall through
    # to the replicated country value.
    gen_of_node_slots = set(zip(gen_of_node_df['Node'], gen_of_node_df['Generator']))
    _distribute_over_eligible_regions(
        alloc, regions_by_prefix, country_vals, MAXINSTALLED_HYDRO_TECHS,
        gen_of_node_slots, cap_lookup, headroom=HYDRO_MAXINSTALLED_HEADROOM)
    _distribute_over_eligible_regions(
        alloc, regions_by_prefix, country_vals, MAXINSTALLED_COASTAL_TECHS,
        gen_of_node_slots, cap_lookup, headroom=1.0)

    # Build NUTS2 rows (node-major, matching the reference layout).
    nuts2_rows = []
    for node in nuts2_nodes:
        vals = country_vals.get(str(node)[:2], {})
        for tech in technologies:
            val = alloc.get((node, tech), vals.get(tech, 0))
            nuts2_rows.append({
                'Node': node,
                'GeneratorTechnology': tech,
                'generatorMaxInstallCapacity  in MW': val,
            })
    nuts2_df = pd.DataFrame(nuts2_rows)

    return pd.concat([ref_non_de_dk, nuts2_df], ignore_index=True)


def make_max_installed_capacity_country(ref_file, technologies):
    """MaxInstalledCapacityCountry — national resource/policy ceilings.

    One row per (Country, Technology) for every NUTS-disaggregated country's
    "national-group" technologies — everything except the resource-based
    techs that stay distributed at NUTS2/3 level (MAXINSTALLED_WEIGHTED_TECHS,
    MAXINSTALLED_HYDRO_TECHS, MAXINSTALLED_COASTAL_TECHS). Values are taken
    unchanged from the reference Generator.xlsx (the same country-level row
    make_max_installed_capacity() reads to build the nodal headroom), so an
    explicit reference zero (e.g. Geo: 0) is kept as a real "not allowed"
    constraint rather than dropped. See NUTS_national_limits_data_spec.md.
    """
    # Hydro deliberately appears in both layers: regional limits retain 150%
    # siting headroom, while this national row keeps the unchanged reference
    # ceiling binding for the whole country.
    regional_techs = (set(MAXINSTALLED_WEIGHTED_TECHS) |
                      set(MAXINSTALLED_COASTAL_TECHS))
    national_techs = [t for t in technologies if t not in regional_techs]

    ref_raw = pd.read_excel(ref_file, sheet_name='MaxInstalledCapacity', header=2)
    ref_df = ref_raw.iloc[:, :3].copy()
    ref_df.columns = ['Node', 'GeneratorTechnology', 'generatorMaxInstallCapacity in MW']
    ref_df = ref_df.dropna(subset=['Node', 'GeneratorTechnology'])
    ref_df['Node'] = ref_df['Node'].map(fix_node_name)

    country_nodes = set(NUTS2_COUNTRY_TO_EMPIRE_NODE.values())
    rows = ref_df[
        ref_df['Node'].isin(country_nodes)
        & ref_df['GeneratorTechnology'].isin(national_techs)
    ].rename(columns={'Node': 'Country'})
    return rows.reset_index(drop=True)


# =============================================================================
# Main
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Generate Generation.xlsx for EMPIRE (NUTS2)')
    parser.add_argument('sets_file', help='Generated Sets_NUTS2_DK_DE.xlsx')
    parser.add_argument('iamc_file', help='IAMC data file (CSV / Excel)')
    parser.add_argument('scenario', help='IAMC scenario name')
    parser.add_argument('output_file', help='Output Generator.xlsx path')
    parser.add_argument('--reference',
                        default='data/EMPIRE/v_52/Generator.xlsx',
                        help='Reference Generator.xlsx '
                             '(default: data/EMPIRE/v_52/Generator.xlsx)')
    args = parser.parse_args()

    print('=' * 80)
    print(f'GENERATION.XLSX GENERATOR FOR EMPIRE  (NUTS2 {", ".join(NUTS2_COUNTRY_PREFIXES)})')
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
    gen_of_node = read_sets_sheet(args.sets_file, 'GeneratorsOfNode')
    nodes = read_sets_sheet_simple(args.sets_file, 'Nodes')['Node'].tolist()
    nodes_by_prefix = {
        prefix: [node for node in nodes if str(node).startswith(prefix)]
        for prefix in NUTS2_COUNTRY_PREFIXES
    }
    technologies = read_sets_sheet_simple(
        args.sets_file, 'Technology')['Technology'].tolist()
    periods = read_sets_sheet_simple(
        args.sets_file, 'Horizon')['Horizon'].tolist()

    print(f'Nodes:                  {len(nodes)}')
    print(f'Technologies:           {len(technologies)}')
    print(f'GeneratorsOfNode rows:  {len(gen_of_node)}')
    print(f'Investment periods:     {periods}')
    print()

    # ------------------------------------------------------------------
    # Sheets copied as-is from reference
    # ------------------------------------------------------------------
    copy_sheets = [
        'CapitalCosts', 'FixedOMCosts', 'VariableOMCosts', 'FuelCosts',
        'CCSCostTSVariable', 'Efficiency', 'ScaleFactorInitialCap',
        'InitialCapacity',
        'RampRate', 'GeneratorTypeAvailability', 'CO2Content', 'Lifetime',
    ]

    raw_copies: dict[str, pd.DataFrame] = {}
    print('Copying unchanged sheets from reference:')
    for name in copy_sheets:
        df = pd.read_excel(args.reference, sheet_name=name, header=None)
        print(f'  {name:30s}  {df.shape[0]:>4} rows')

        # Efficiency=0 for renewables (Wind, Solar, Hydro, Wave, Geo) causes
        # ZeroDivisionError in prepOperationalCostGen_rule: 3.6/genEfficiency[g,i].
        # Replace 0 with 1.0 — economics unchanged since their fuel/CO2 cost is 0.
        if name == 'Efficiency':
            eff_col = df.columns[2]   # column index 2 = efficiency value
            data_rows = df.index[3:]  # rows 0-2 are metadata/header
            zero_mask = (df.loc[data_rows, eff_col] == 0)
            n_fixed = zero_mask.sum()
            if n_fixed > 0:
                df.loc[data_rows[zero_mask], eff_col] = 1.0
                print(f'  {"":30s}       [replaced {n_fixed} zero efficiency values with 1.0]')

        raw_copies[name] = df
    print()

    reference_sheets = pd.ExcelFile(args.reference).sheet_names
    has_yearly_availability = 'YearlyAvailability' in reference_sheets
    if has_yearly_availability:
        print('Transforming optional YearlyAvailability sheet...')
        raw_copies['YearlyAvailability'] = make_yearly_availability(
            nodes, nodes_by_prefix, args.reference, gen_of_node
        )
        yearly_availability_rows = (
            raw_copies['YearlyAvailability'].shape[0] - 3
        )
        print(
            f'  YearlyAvailability: '
            f'{yearly_availability_rows} data rows'
        )
        print()
    else:
        # Non-fatal: some reference datasets genuinely lack the sheet.  Labelled
        # WARNING so it surfaces in a log scan — an upstream step silently
        # dropping the sheet looks identical to it never being there.
        print(
            '  WARNING: optional YearlyAvailability sheet not present in '
            'reference; skipping. Check this is expected for this dataset — a '
            'preprocessing step may have dropped it.'
        )
        print()

    # ------------------------------------------------------------------
    # Build new sheets
    # ------------------------------------------------------------------
    print('[1/6] RefInitialCap '
          '(reference for non-NUTS2, IAMC capacity for NUTS2)...')
    ric_df = make_ref_initial_cap(
        gen_of_node, args.iamc_file, args.scenario, args.reference)
    nonzero = (ric_df['generatoReferenceInitialCapacity in MW'] > 0).sum()
    print(f'  -> {len(ric_df)} rows  ({nonzero} with capacity > 0)')

    print('[2/6] MaxBuiltCapacity '
          '(reference for non-NUTS2; national-level for NUTS2 nodes, see '
          'MaxBuiltCapacityCountry)...')
    mbc_df = make_max_built_capacity(nodes, args.reference)
    print(f'  -> {len(mbc_df)} rows')

    print('[3/6] MaxInstalledCapacity '
          '(reference for non-NUTS2; Wind/Solar/Hydro/offshore/wave ceilings '
          'distributed across NUTS2, generous headroom for other techs)...')
    mic_df = make_max_installed_capacity(
        nodes, technologies, args.reference, args.iamc_file, args.scenario,
        gen_of_node)
    print(f'  -> {len(mic_df)} rows')

    print('[4/6] MinBuiltCapacity (nodal, restored for non-disaggregated countries)...')
    minbc_df = make_min_built_capacity(args.reference)
    print(f'  -> {len(minbc_df)} rows')

    print('[5/6] National-level limit sheets '
          '(MaxInstalledCapacityCountry, MaxBuiltCapacityCountry, '
          'MinBuiltCapacityCountry)...')
    mic_country_df = make_max_installed_capacity_country(args.reference, technologies)
    mbc_country_df = make_max_built_capacity_country(args.reference)
    minbc_country_df = make_min_built_capacity_country(args.reference)
    print(f'  -> MaxInstalledCapacityCountry: {len(mic_country_df)} rows')
    print(f'  -> MaxBuiltCapacityCountry:     {len(mbc_country_df)} rows')
    print(f'  -> MinBuiltCapacityCountry:     {len(minbc_country_df)} rows')
    print()

    # ------------------------------------------------------------------
    # Add metadata headers
    # ------------------------------------------------------------------
    meta: dict[str, pd.DataFrame] = {}
    meta['RefInitialCap'] = add_metadata_header(
        ric_df,
        source=args.iamc_file,
        description='Capacity in the reference investment period '
                    '(reference for non-DE/DK; IAMC GW->MW for NUTS2) '
                    '(default: 0)',
    )
    meta['MaxBuiltCapacity'] = add_metadata_header(
        mbc_df,
        source='Assumed',
        description='Maximum capacity expansion of generator types in one '
                    'node and investment period (default: 500 000)',
    )
    meta['MaxInstalledCapacity'] = add_metadata_header(
        mic_df,
        source='See reference Generator.xlsx',
        description='Maximum capacity that can exist (resource limit) '
                    'of a generator (default: 0)',
    )
    meta['MinBuiltCapacity'] = add_metadata_header(
        minbc_df,
        source='Assumed',
        description='Minimum capacity expansion of generator types in one '
                    'node and investment period (default: 0)',
    )
    meta['MaxInstalledCapacityCountry'] = add_metadata_header(
        mic_country_df,
        source='See reference Generator.xlsx',
        description='Maximum total installed capacity of a technology across '
                    'all nodes of a country, enforced in every period '
                    '(only listed pairs constrained)',
    )
    meta['MaxBuiltCapacityCountry'] = add_metadata_header(
        mbc_country_df,
        source='Assumed',
        description='Maximum capacity expansion of a technology across all '
                    'nodes of a country in one investment period '
                    '(only listed triples constrained)',
    )
    meta['MinBuiltCapacityCountry'] = add_metadata_header(
        minbc_country_df,
        source='Assumed',
        description='Minimum capacity expansion of a technology across all '
                    'nodes of a country in one investment period '
                    '(only listed triples constrained)',
    )

    # ------------------------------------------------------------------
    # Write output  (preserve sheet order from reference)
    # ------------------------------------------------------------------
    #  1  CapitalCosts                  (copy)
    #  2  FixedOMCosts                  (copy)
    #  3  VariableOMCosts               (copy)
    #  4  FuelCosts                     (copy)
    #  5  CCSCostTSVariable             (copy)
    #  6  Efficiency                    (copy)
    #  7  ScaleFactorInitialCap         (copy)
    #  8  RefInitialCap                 (new)
    #  9  InitialCapacity               (copy)
    # 10  MaxBuiltCapacity              (new)
    # 11  MaxBuiltCapacityCountry       (new)
    # 12  MaxInstalledCapacity          (new)
    # 13  MaxInstalledCapacityCountry   (new)
    # 14  MinBuiltCapacity              (new, restored)
    # 15  MinBuiltCapacityCountry       (new)
    # 16  RampRate                      (copy)
    # 17  GeneratorTypeAvailability     (copy)
    # 18  CO2Content                    (copy)
    # 19  Lifetime                      (copy)

    sheet_order = []
    if has_yearly_availability:
        sheet_order.append(('optional', 'YearlyAvailability'))
    sheet_order += [
        ('copy',  'CapitalCosts'),
        ('copy',  'FixedOMCosts'),
        ('copy',  'VariableOMCosts'),
        ('copy',  'FuelCosts'),
        ('copy',  'CCSCostTSVariable'),
        ('copy',  'Efficiency'),
        ('copy',  'ScaleFactorInitialCap'),
        ('meta',  'RefInitialCap'),
        ('copy',  'InitialCapacity'),
        ('meta',  'MaxBuiltCapacity'),
        ('meta',  'MaxBuiltCapacityCountry'),
        ('meta',  'MaxInstalledCapacity'),
        ('meta',  'MaxInstalledCapacityCountry'),
        ('meta',  'MinBuiltCapacity'),
        ('meta',  'MinBuiltCapacityCountry'),
        ('copy',  'RampRate'),
        ('copy',  'GeneratorTypeAvailability'),
        ('copy',  'CO2Content'),
        ('copy',  'Lifetime'),
    ]

    print(f'Writing {args.output_file} ...')
    with pd.ExcelWriter(args.output_file, engine='openpyxl') as writer:
        for kind, name in sheet_order:
            if kind in ('copy', 'optional'):
                raw_copies[name].to_excel(
                    writer, sheet_name=name, index=False, header=False)
            else:
                meta[name].to_excel(
                    writer, sheet_name=name, index=False, header=False)

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    print()
    print('=' * 80)
    print('GENERATION.XLSX  SUMMARY')
    print('=' * 80)
    print(f'Output: {args.output_file}')
    print()
    print(f'{"Sheet":<32s}  {"Rows":>6s}  Source')
    print('-' * 70)
    for kind, name in sheet_order:
        if kind == 'copy':
            n_rows = raw_copies[name].shape[0] - 3  # subtract metadata + header
            print(f'  {name:<30s}  {n_rows:>6}  copied from reference')
        elif kind == 'optional':
            n_rows = raw_copies[name].shape[0] - 3
            print(f'  {name:<30s}  {n_rows:>6}  reference + NUTS expansion')
        elif name == 'RefInitialCap':
            print(f'  {name:<30s}  {len(ric_df):>6}  '
                  f'ref + IAMC  ({nonzero} > 0)')
        elif name == 'MaxBuiltCapacity':
            print(f'  {name:<30s}  {len(mbc_df):>6}  '
                  f'ref (non-NUTS2); national for NUTS2 (see *Country)')
        elif name == 'MaxBuiltCapacityCountry':
            print(f'  {name:<30s}  {len(mbc_country_df):>6}  '
                  f'ref country-level rows (all techs)')
        elif name == 'MaxInstalledCapacity':
            print(f'  {name:<30s}  {len(mic_df):>6}  '
                  f'ref + NUTS2 distribution/headroom')
        elif name == 'MaxInstalledCapacityCountry':
            print(f'  {name:<30s}  {len(mic_country_df):>6}  '
                  f'ref country-level rows (national-group techs)')
        elif name == 'MinBuiltCapacity':
            print(f'  {name:<30s}  {len(minbc_df):>6}  '
                  f'ref (non-disaggregated countries only)')
        elif name == 'MinBuiltCapacityCountry':
            print(f'  {name:<30s}  {len(minbc_country_df):>6}  '
                  f'ref country-level rows (empty unless mandated)')
    print()

    # IAMC coverage
    gens_found = sorted(
        ric_df.loc[
            (ric_df['generatoReferenceInitialCapacity in MW'] > 0)
            & (ric_df['Node'].isin(ALL_NUTS2_SET)),
            'GeneratorTechnology',
        ].unique()
    )
    print(f'Generator types with IAMC capacity in NUTS2 nodes: {gens_found}')
    print()
    print('Done!')
    print('=' * 80)


if __name__ == '__main__':
    main()
