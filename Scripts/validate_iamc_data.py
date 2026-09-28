"""
IAMC data validation for the EMPIRE NUTS2 pipeline.

Checks that the IAMC input file contains the scenario, NUTS2 regions, and
variables required by the generate_*.py scripts before the pipeline runs.

Exit codes:
    0 — all critical checks passed (warnings are printed but do not block)
    1 — one or more critical checks failed

Standalone usage:
    python Scripts/validate_iamc_data.py <iamc_file> <scenario>

Called automatically by main.py before each run; bypass with --skip-validation.
"""

import sys
import argparse
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

try:
    import pyam
except ImportError:
    print("ERROR: pyam-iamc not installed. Run: pip install pyam-iamc")
    sys.exit(1)

sys.path.insert(0, str(Path(__file__).parent))
from config_loader import load_nuts2_config

# ---------------------------------------------------------------------------
# Variables used by the generate_*.py scripts
# ---------------------------------------------------------------------------

# Used by generate_generation_nuts2.py / generate_sets_nuts2.py (RefInitialCap,
# GeneratorsOfNode). Nodes with no data get zero capacity — not a crash, but
# a content problem worth flagging.
_CAPACITY_VARS = [
    'Capacity|Electricity|Wind|Onshore',
    'Capacity|Electricity|Wind|Offshore',
    'Capacity|Electricity|Solar|PV',
    'Capacity|Electricity|Solar|CSP',
    'Capacity|Electricity|Hydro',
    'Capacity|Electricity|Hydro|Pumped Storage',
    'Capacity|Electricity|Gas|CCGT|w/o CCS',
    'Capacity|Electricity|Gas|w/o CCS',
    'Capacity|Electricity|Coal|w/o CCS',
    'Capacity|Electricity|Biomass|w/o CCS',
    'Capacity|Electricity|Oil|w/o CCS',
]

# Used by generate_node.py to distribute country demand across NUTS2 regions.
# Falls back to equal weights when absent, so it is a warning not an error.
_DEMAND_SHARE_VAR = 'Final Energy|Electricity [Share]'

# Used by generate_transmission_nuts2.py for InitialCapacity and
# MaxInstallCapacityRaw. Both AC and DC variants are loaded; falls back to
# the reference file when absent.
_TRANSMISSION_VARS = [
    'Network|Electricity|Capacity',
    'Network|Electricity|DC|Capacity',
]


# ---------------------------------------------------------------------------
# Core validation logic
# ---------------------------------------------------------------------------

def validate(iamc_file: str, scenario: str, config_path=None):
    """Run pre-flight validation on the IAMC file for the NUTS2 pipeline.

    Returns (ok, errors, warnings):
        ok       — False when any critical check fails
        errors   — issues that will produce incorrect or missing output
        warnings — issues that fall back gracefully but deserve attention
    """
    errors: list[str] = []
    warnings: list[str] = []

    cfg = load_nuts2_config(config_path)
    all_nuts2 = cfg['ALL_NUTS2_SET']
    regions_by_country = cfg['NUTS2_REGIONS_BY_COUNTRY']
    prefixes = cfg['NUTS2_COUNTRY_PREFIXES']

    # ---------------------------------------------------------------
    # 1. Load file
    # ---------------------------------------------------------------
    print("[1/5] Loading IAMC file...")
    try:
        from iamc_utils import load_iamc
        df_full = load_iamc(iamc_file)
    except Exception as exc:
        errors.append(f"Cannot load IAMC file: {exc}")
        return False, errors, warnings
    print(f"      {len(df_full.data):,} rows  ({iamc_file})")

    # ---------------------------------------------------------------
    # 2. Scenario
    # ---------------------------------------------------------------
    print(f"[2/5] Checking scenario '{scenario}'...")
    available = sorted(set(df_full.scenario))
    if scenario not in available:
        errors.append(
            f"Scenario '{scenario}' not found.  Available: {available}"
        )
        return False, errors, warnings
    df = df_full.filter(scenario=scenario)
    print(f"      Found ({len(df.data):,} rows in this scenario)")

    # ---------------------------------------------------------------
    # 3. NUTS2 region coverage
    # ---------------------------------------------------------------
    print("[3/5] Checking NUTS2 region coverage...")
    iamc_nodes = {r for r in df.region if '>' not in r}

    for prefix in prefixes:
        configured = regions_by_country[prefix]
        missing = [r for r in configured if r not in iamc_nodes]
        n_found = len(configured) - len(missing)
        print(f"      {prefix}: {n_found}/{len(configured)} regions found", end="")
        if missing:
            short = ', '.join(missing[:5]) + ('...' if len(missing) > 5 else '')
            print(f"  — {len(missing)} missing: {short}")
            errors.append(
                f"{prefix}: {len(missing)} configured region(s) absent from IAMC: "
                + ', '.join(missing)
            )
        else:
            print()

    # ---------------------------------------------------------------
    # 4. Variables for NUTS2 nodes
    # ---------------------------------------------------------------
    print("[4/5] Checking variables for NUTS2 nodes...")
    nuts2_data = df.data[df.data['region'].isin(all_nuts2)]
    found_vars = set(nuts2_data['variable'].unique())

    if _DEMAND_SHARE_VAR in found_vars:
        print(f"      ✓ '{_DEMAND_SHARE_VAR}'")
    else:
        warnings.append(
            f"'{_DEMAND_SHARE_VAR}' absent — "
            "demand will be split equally across NUTS2 nodes per country"
        )
        print(f"      ⚠ '{_DEMAND_SHARE_VAR}' not found (demand falls back to equal weights)")

    found_cap = [v for v in _CAPACITY_VARS if v in found_vars]
    missing_cap = [v for v in _CAPACITY_VARS if v not in found_vars]
    print(f"      Capacity variables: {len(found_cap)}/{len(_CAPACITY_VARS)} found")
    for v in found_cap:
        print(f"        ✓ {v}")
    for v in missing_cap:
        print(f"        ⚠ {v}  (absent — affected generators get zero initial capacity)")
        warnings.append(
            f"'{v}' absent for NUTS2 nodes — affected generators will have zero initial capacity"
        )

    # ---------------------------------------------------------------
    # 5. Transmission link data
    # ---------------------------------------------------------------
    print("[5/5] Checking transmission link data...")
    nuts2_links = [
        r for r in df.region
        if '>' in r and any(p.strip() in all_nuts2 for p in r.split('>'))
    ]
    any_found = False
    for tvar in _TRANSMISSION_VARS:
        trans_rows = df.data[
            (df.data['region'].isin(nuts2_links)) &
            (df.data['variable'] == tvar)
        ]
        if not trans_rows.empty:
            print(f"      ✓ '{tvar}': {len(trans_rows)} rows")
            any_found = True
        else:
            print(f"      — '{tvar}': not found")
    if not any_found:
        warnings.append(
            f"Neither AC nor DC transmission capacity found for NUTS2 links — "
            "transmission capacities will use reference data"
        )
        print(f"      ⚠ No transmission capacity data for NUTS2 links "
              f"(will fall back to reference)")

    return len(errors) == 0, errors, warnings


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description='Validate IAMC input data for the EMPIRE NUTS2 pipeline'
    )
    parser.add_argument('iamc_file', help='Path to IAMC file (CSV or Excel)')
    parser.add_argument('scenario', help='Scenario name to validate')
    args = parser.parse_args()

    print('=' * 72)
    print('IAMC PRE-FLIGHT VALIDATION')
    print('=' * 72)
    print(f'File:     {args.iamc_file}')
    print(f'Scenario: {args.scenario}')
    print()

    ok, errors, warnings = validate(args.iamc_file, args.scenario)

    print()
    print('=' * 72)
    print('SUMMARY')
    print('=' * 72)
    if warnings:
        print(f'Warnings ({len(warnings)}):')
        for w in warnings:
            print(f'  ⚠ {w}')
        print()
    if errors:
        print(f'Errors ({len(errors)}):')
        for e in errors:
            print(f'  ✗ {e}')
        print()
    if ok:
        print('✓ Validation passed — pipeline can proceed.')
    else:
        print('✗ Validation FAILED — fix errors above before running main.py.')
    print('=' * 72)

    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
