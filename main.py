"""
EMPIRE NUTS2 input data generator — North Sea case.

Reads an existing EMPIRE reference dataset and generates a new version in which
the countries listed under ``nuts2_countries`` in config.yaml (NO, UK, DK, DE,
NL) are replaced by their NUTS2 regions, keeping all other countries at NUTS0.

Usage
-----
    python main.py [empire_folder] [--iamc <path>] [--scenario <name>]
                   [--output-folder <subfolder>] [--atlite-dir <dir>]
                   [--skip-validation]

Arguments
---------
    empire_folder   Name of the EMPIRE reference dataset folder under
                    data/EMPIRE/ (e.g. NECPEssentials_Norway).
                    Default: empire_folder in config.yaml.

Options
-------
    --iamc          Path to the IAMC file.   Default: iamc_file in config.yaml.
    --scenario      IAMC scenario name.      Default: read from the IAMC file.
    --output-folder Subfolder inside Result/. Default: output_folder in config.yaml.

Output
------
    Result/[<output_folder>/]<empire_folder>/
        Sets.xlsx         -- node list, technologies, corridors
        Generator.xlsx    -- generation capacities at NUTS2
        Transmission.xlsx -- transmission lines at NUTS2
        Node.xlsx         -- demand / hydro at NUTS2
        Storage.xlsx      -- storage at NUTS2
        General.xlsx      -- copied from the source folder
        ScenarioData/     -- hourly time-series at NUTS2
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
SCRIPTS_DIR = ROOT / 'Scripts'


def _config_defaults():
    """Read defaults from config.yaml."""
    try:
        import yaml
        with open(ROOT / 'config.yaml', encoding='utf-8') as f:
            cfg = yaml.safe_load(f)
        return (
            cfg.get('empire_folder', ''),
            cfg.get('iamc_file', ''),
            cfg.get('scenario', ''),
            cfg.get('output_folder', ''),
            cfg.get('atlite_dir', ''),
        )
    except Exception:
        return '', '', '', '', ''


def _load_flow_defaults():
    """Read the load_flow section from config.yaml (with defaults)."""
    try:
        import yaml
        with open(ROOT / 'config.yaml', encoding='utf-8') as f:
            cfg = yaml.safe_load(f) or {}
        lf = cfg.get('load_flow') or {}
        return {
            'enabled': bool(lf.get('enabled', False)),
            'reactance_per_km': float(lf.get('reactance_per_km', 0.246)),
            'block_capacity_mw': float(lf.get('block_capacity_mw', 1500)),
            'nominal_voltage_kv': float(lf.get('nominal_voltage_kv', 400)),
        }
    except Exception:
        return {'enabled': False, 'reactance_per_km': 0.246,
                'block_capacity_mw': 1500, 'nominal_voltage_kv': 400}


def _write_general_loadflow_sheets(general_path, lf):
    """Append the load-flow scalar sheets to a copied General.xlsx.

    Adds NominalVoltage plus the (optional) global fallback sheets. Each sheet
    follows the dataset convention: 2 description rows then the scalar value.
    """
    import openpyxl
    sheets = {
        'NominalVoltage': (
            'Default nominal voltage of the transmission grid. (kV)',
            lf['nominal_voltage_kv'],
        ),
        'LineBlockCapacityGlobal': (
            'Candidate lines capacity for transmission expansion planning. MW '
            '(global fallback; per-candidate LineBlockCapacity is authoritative)',
            lf['block_capacity_mw'],
        ),
        'LineBlockReactanceGlobal': (
            'Candidate lines reactance for transmission expansion planning. Ohm '
            '(global fallback = reactance_per_km x 100 km; per-candidate '
            'LineBlockReactance is authoritative)',
            round(lf['reactance_per_km'] * 100.0, 4),
        ),
    }
    # Re-saving with openpyxl drops the cached results of any formula cells (some
    # datasets, e.g. NECPEssentials CO2Cap, store values as '=123.4' formulas),
    # which would leave EMPIRE reading empty params. Flatten every formula to its
    # cached value first, using a data_only view of the same file.
    vals = openpyxl.load_workbook(general_path, data_only=True)
    wb = openpyxl.load_workbook(general_path)
    flattened = 0
    for ws in wb.worksheets:
        vws = vals[ws.title]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = vws[cell.coordinate].value
                    flattened += 1
    for name, (desc, value) in sheets.items():
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name)
        ws.append(['Source: -'])
        ws.append([f'Description: {desc}'])
        ws.append([value])
    wb.save(general_path)
    note = f" (flattened {flattened} formula cell(s))" if flattened else ""
    print(f"  Added load-flow scalar sheets to General.xlsx: {', '.join(sheets)}{note}")


def run_script(script_name: str, args: list):
    script = SCRIPTS_DIR / script_name
    cmd = [sys.executable, str(script)] + [str(a) for a in args]
    print(f"\n{'=' * 80}")
    print(f"Running: {script_name}")
    print(f"{'=' * 80}")
    result = subprocess.run(cmd, cwd=ROOT)
    if result.returncode != 0:
        print(f"\nERROR: {script_name} failed (exit code {result.returncode})")
        sys.exit(result.returncode)


def run_validation(iamc_file: str, scenario: str):
    """Run the IAMC pre-flight validation script.

    Exits the process if critical checks fail so the generate scripts
    are never invoked with bad input.
    """
    print(f"\n{'=' * 80}")
    print("Running: validate_iamc_data.py  (pre-flight check)")
    print(f"{'=' * 80}")
    script = SCRIPTS_DIR / 'validate_iamc_data.py'
    result = subprocess.run(
        [sys.executable, str(script), iamc_file, scenario],
        cwd=ROOT,
    )
    if result.returncode != 0:
        print("\nPipeline aborted: fix the validation errors above and re-run.")
        sys.exit(result.returncode)


def main():
    cfg_folder, cfg_iamc, cfg_scenario, cfg_output, cfg_atlite = _config_defaults()

    parser = argparse.ArgumentParser(
        description='Generate EMPIRE input data at NUTS2 resolution for the '
                    'countries listed in config.yaml'
    )
    parser.add_argument(
        'empire_folder',
        nargs='?',
        default=cfg_folder or None,
        help='EMPIRE reference dataset folder under data/EMPIRE/ '
             '(default: empire_folder in config.yaml)'
    )
    parser.add_argument(
        '--iamc',
        default=cfg_iamc or 'data/IAMC/NorthSea/North_Sea.xlsx',
        help='Path to IAMC file (default: value from config.yaml)'
    )
    parser.add_argument(
        '--scenario',
        default=cfg_scenario or '',
        help='IAMC scenario name (default: value from config.yaml; if that is '
             'unset too, the name is read from the IAMC file, which works as '
             'long as the file holds exactly one scenario)'
    )
    parser.add_argument(
        '--output-folder',
        default=cfg_output or '',
        help='Subfolder inside Result/ for this run. '
             'Output goes to Result/<output-folder>/<empire_folder>/ '
             '(default: value from config.yaml; if empty, Result/<empire_folder>/)'
    )
    parser.add_argument(
        '--atlite-dir',
        default=cfg_atlite or '',
        help='Directory with region-specific CF CSVs (solar.csv, windonshore.csv) '
             'from Scripts/preprocess_renewables_atlite.py. When set, the scenario '
             'step uses per-region Atlite series instead of broadcasting the national '
             'CF (default: value from config.yaml; if empty, broadcast).'
    )
    parser.add_argument(
        '--skip-validation',
        action='store_true',
        help='Skip IAMC pre-flight validation (not recommended)'
    )
    args = parser.parse_args()

    if not args.empire_folder:
        parser.error(
            "empire_folder is required — set it in config.yaml or pass it as an argument.\n"
            f"  Available folders: {', '.join(p.name for p in (ROOT / 'data' / 'EMPIRE').iterdir() if p.is_dir())}"
        )

    # The scenario name is a property of the IAMC file, not a choice: the
    # input file carries a single scenario.  Detect it unless the user pinned one.
    sys.path.insert(0, str(SCRIPTS_DIR))
    from iamc_utils import resolve_scenario
    try:
        detected = resolve_scenario(args.iamc, args.scenario)
    except (ValueError, FileNotFoundError, OSError) as exc:
        print(f'ERROR: {exc}')
        sys.exit(1)
    if not args.scenario:
        print(f'  Scenario detected from {args.iamc}: "{detected}"')
    args.scenario = detected

    ref_dir = ROOT / 'data' / 'EMPIRE' / args.empire_folder
    if args.output_folder:
        out_dir = ROOT / 'Result' / args.output_folder / args.empire_folder
    else:
        out_dir = ROOT / 'Result' / args.empire_folder

    if not ref_dir.exists():
        print(f"ERROR: EMPIRE folder not found: {ref_dir}")
        available = sorted(
            p.name for p in (ROOT / 'data' / 'EMPIRE').iterdir() if p.is_dir()
        )
        print(f"Available folders:\n  " + "\n  ".join(available))
        sys.exit(1)

    out_dir.mkdir(parents=True, exist_ok=True)

    sets_out         = out_dir / 'Sets.xlsx'
    generator_out    = out_dir / 'Generator.xlsx'
    transmission_out = out_dir / 'Transmission.xlsx'
    node_out         = out_dir / 'Node.xlsx'
    storage_out      = out_dir / 'Storage.xlsx'

    print(f"\nEMPIRE NUTS2 DATA GENERATOR")
    print(f"  Source : {ref_dir}")
    print(f"  Output : {out_dir}")
    print(f"  IAMC   : {args.iamc}")
    print(f"  Scenario: {args.scenario}")

    # 0. Validation
    if args.skip_validation:
        print("\nSkipping IAMC validation (--skip-validation)")
    else:
        run_validation(args.iamc, args.scenario)

    # 1. Sets
    run_script('generate_sets_nuts2.py', [
        args.iamc, args.scenario, sets_out,
        '--reference', ref_dir / 'Sets.xlsx',
    ])

    # 2. Generation
    run_script('generate_generation_nuts2.py', [
        sets_out, args.iamc, args.scenario, generator_out,
        '--reference', ref_dir / 'Generator.xlsx',
    ])

    # 3. Transmission
    run_script('generate_transmission_nuts2.py', [
        args.iamc, args.scenario, sets_out, transmission_out,
        '--reference', ref_dir / 'Transmission.xlsx',
    ])

    # 4. Node
    run_script('generate_node.py', [
        sets_out, args.iamc, args.scenario, node_out,
        '--reference', ref_dir / 'Node.xlsx',
    ])

    # 5. Storage
    run_script('generate_storage.py', [
        sets_out, args.iamc, args.scenario, storage_out,
        '--reference', ref_dir / 'Storage.xlsx',
    ])

    # 6. ScenarioData
    scenario_data_src = ref_dir / 'ScenarioData'
    if scenario_data_src.exists():
        scenario_args = [
            '--iamc', args.iamc,
            '--scenario', args.scenario,
            '--sets-file', sets_out,
            '--input-dir', scenario_data_src,
            '--output-dir', out_dir / 'ScenarioData',
        ]
        if args.atlite_dir:
            scenario_args += ['--atlite-dir', args.atlite_dir]
        run_script('generate_scenarios.py', scenario_args)
    else:
        print(f"\nWARNING: ScenarioData not found in {ref_dir}, skipping.")

    # 7. General.xlsx — no transformation needed, copy as-is
    general_src = ref_dir / 'General.xlsx'
    if general_src.exists():
        general_out = out_dir / 'General.xlsx'
        shutil.copy2(general_src, general_out)
        print(f"\nCopied General.xlsx from {ref_dir}")
        lf = _load_flow_defaults()
        if lf['enabled']:
            _write_general_loadflow_sheets(general_out, lf)
    else:
        print(f"\nWARNING: General.xlsx not found in {ref_dir}, skipping.")

    print(f"\n{'=' * 80}")
    print(f"Done. Output written to: {out_dir}")
    print(f"{'=' * 80}")


if __name__ == '__main__':
    main()
