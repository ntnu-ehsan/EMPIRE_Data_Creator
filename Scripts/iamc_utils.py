"""
Shared helpers for reading IAMC capacity/scenario files.

The North Sea NUTS2 generation file (generation_nuts2_IAMC_northsea.xlsx) lists
capacity at plant level, so the same (model, scenario, region, variable, unit,
year) key can appear on several rows that must be summed to the NUTS2 total.
``pyam.IamDataFrame`` rejects such a file outright ("Duplicate rows in `data`"),
which would break every script in the pipeline.

``load_iamc`` wraps ``pyam.IamDataFrame`` and, only when that duplicate-row
error is raised, falls back to reading the raw table with pandas, summing the
duplicate rows into one value per key, and rebuilding the IamDataFrame.  Files
without duplicates are passed straight through unchanged.
"""

from pathlib import Path

import pandas as pd

try:
    import pyam
except ImportError:  # pragma: no cover - handled by callers that import pyam too
    pyam = None

# IAMC long-format index columns (everything except the numeric value).
_IAMC_INDEX = ['model', 'scenario', 'region', 'variable', 'unit', 'year']

HYDRO_BARE_VARIABLE = 'Capacity|Electricity|Hydro'
HYDRO_RESERVOIR_VARIABLE = 'Capacity|Electricity|Hydro|Reservoir'
HYDRO_ROR_VARIABLE = 'Capacity|Electricity|Hydro|Run of River'
HYDRO_CAPACITY_VARIABLES = (
    HYDRO_BARE_VARIABLE,
    HYDRO_RESERVOIR_VARIABLE,
    HYDRO_ROR_VARIABLE,
)


def _read_raw_table(path):
    """Read an IAMC long-format file (xlsx or csv) into a DataFrame."""
    p = str(path)
    if p.lower().endswith('.csv'):
        return pd.read_csv(path)
    return pd.read_excel(path)


def _to_long_summed(raw):
    """Convert a raw IAMC table (long or wide) to a deduplicated long frame.

    Returns a DataFrame with columns model/scenario/region/variable/unit/year/
    value where duplicate (model, scenario, region, variable, unit, year) keys
    have been summed, or ``None`` if the layout cannot be interpreted.
    """
    raw = raw.rename(columns={c: str(c).strip() for c in raw.columns})
    raw = raw.rename(columns={c: c.lower() for c in raw.columns})
    id_cols = [c for c in ['model', 'scenario', 'region', 'variable', 'unit']
               if c in raw.columns]
    if 'region' not in id_cols or 'variable' not in id_cols:
        return None
    cols = list(raw.columns)

    if 'value' in cols:
        # Long format (explicit value column, optional year column).
        long = raw[id_cols + (['year'] if 'year' in cols else []) + ['value']].copy()
        if 'year' not in long.columns:
            long['year'] = 0
    else:
        # Wide format: one column per year, header parseable as a year number.
        year_cols = []
        for c in cols:
            if c in id_cols:
                continue
            try:
                y = int(float(str(c)))
            except (ValueError, TypeError):
                continue
            if 1900 <= y <= 2200:
                year_cols.append((c, y))
        if not year_cols:
            return None
        frames = []
        for col, year in year_cols:
            sub = raw[id_cols].copy()
            sub['year'] = year
            sub['value'] = pd.to_numeric(raw[col], errors='coerce')
            frames.append(sub)
        long = pd.concat(frames, ignore_index=True)

    long['value'] = pd.to_numeric(long['value'], errors='coerce')
    long = long.dropna(subset=['value'])
    if 'model' not in long.columns:
        long['model'] = 'model'
    idx = [c for c in _IAMC_INDEX if c in long.columns]
    agg = long.groupby(idx, as_index=False, dropna=False)['value'].sum()
    return long, agg


def load_iamc(source, **kwargs):
    """Return a ``pyam.IamDataFrame`` for *source*, tolerating duplicate rows.

    *source* may be a file path or anything ``pyam.IamDataFrame`` accepts.  When
    pyam raises a structural ValueError (duplicate plant-level rows that have not
    been aggregated to one row per region/variable, or stray empty cells), the
    table is re-read with pandas, the duplicate rows are summed, and a clean
    IamDataFrame is rebuilt.  Both long and single/multi-year wide layouts are
    handled.
    """
    if pyam is None:
        raise ImportError("pyam-iamc not installed. Run: pip install pyam-iamc")

    try:
        return pyam.IamDataFrame(source, **kwargs)
    except ValueError as exc:
        msg = str(exc)
        if 'Duplicate rows' not in msg and 'Empty cells' not in msg:
            raise
        if not isinstance(source, (str, Path)):
            raise

        result = _to_long_summed(_read_raw_table(source))
        if result is None:
            raise
        long, agg = result
        n_collapsed = len(long) - len(agg)
        print(f"  [iamc] Summed {n_collapsed} duplicate plant-level row(s) into "
              f"NUTS2 totals while reading {source}")
        return pyam.IamDataFrame(agg, **kwargs)


def _capacity_values_mw(data):
    '''Return IAMC capacity rows with a unit-normalised value_mw column.'''
    result = data.copy()
    units = result['unit'].fillna('MW').astype(str).str.strip()
    factors = units.map({'GW': 1000.0, 'MW': 1.0, 'kW': 0.001})
    unsupported = sorted(units[factors.isna()].unique())
    if unsupported:
        raise ValueError(
            f'Unsupported hydro-capacity unit(s): {unsupported}; expected GW, MW, or kW'
        )
    result['value_mw'] = pd.to_numeric(result['value'], errors='coerce') * factors
    return result.dropna(subset=['value_mw'])


def select_hydro_capacity_weights(data, regions_by_group, purpose):
    '''Return ({region: MW}, {group: source}) for a hydro constraint.

    annual_total sums all non-pumped generating hydro components.
    seasonal_reservoir uses Reservoir, with bare Hydro as a legacy fallback.
    '''
    if purpose not in {'annual_total', 'seasonal_reservoir'}:
        raise ValueError(f'Unknown hydro-weight purpose: {purpose!r}')
    required = {'region', 'variable', 'unit', 'value'}
    missing = required - set(data.columns)
    if missing:
        raise ValueError(f'IAMC hydro data missing columns: {sorted(missing)}')

    hydro = data[data['variable'].isin(HYDRO_CAPACITY_VARIABLES)].copy()
    hydro = hydro[~hydro['region'].astype(str).str.contains('>', na=False)]
    hydro = _capacity_values_mw(hydro)
    values = hydro.groupby(['region', 'variable'])['value_mw'].sum()
    weights, sources = {}, {}
    for group, regions in regions_by_group.items():
        regions = list(regions)

        def component(variable):
            return {
                r: float(values.get((r, variable), 0.0)) for r in regions
            }

        bare = component(HYDRO_BARE_VARIABLE)
        reservoir = component(HYDRO_RESERVOIR_VARIABLE)
        ror = component(HYDRO_ROR_VARIABLE)
        if purpose == 'annual_total':
            selected = {r: bare[r] + reservoir[r] + ror[r] for r in regions}
            sources[group] = (
                'total non-pumped hydro capacity '
                '(Reservoir + Run of River + bare Hydro residual)'
                if sum(selected.values()) > 0 else 'no hydro capacity data'
            )
        elif sum(reservoir.values()) > 0:
            selected, sources[group] = reservoir, 'reservoir capacity'
        elif sum(bare.values()) > 0:
            selected = bare
            sources[group] = 'bare Hydro capacity (legacy fallback)'
        else:
            selected = {r: 0.0 for r in regions}
            sources[group] = 'no reservoir or legacy bare-Hydro capacity data'
        weights.update(selected)
    return weights, sources


def load_hydro_capacity_weights(source, scenario, regions_by_group, purpose):
    '''Load an IAMC source and select unit-normalised hydro weights.'''
    idf = load_iamc(source)
    data = idf.filter(scenario=scenario,
                      variable=list(HYDRO_CAPACITY_VARIABLES)).data
    return select_hydro_capacity_weights(data, regions_by_group, purpose)


def list_scenarios(source):
    """Return the sorted scenario names present in an IAMC file.

    Reads the raw table rather than going through pyam: the caller only needs
    the scenario column, and this stays cheap on the large plant-level files.
    """
    raw = _read_raw_table(source)
    col = next((c for c in raw.columns if str(c).strip().lower() == 'scenario'),
               None)
    if col is None:
        return []
    return sorted(str(s) for s in raw[col].dropna().unique())


def resolve_scenario(source, requested=''):
    """Return the scenario to use for *source*.

    An IAMC input file describes one scenario, so every row carries the same
    scenario name and that name is a property of the file rather than something
    the caller has to repeat in config.yaml.  It is read from the file here.

    A *requested* name still wins, which is the escape hatch for a file whose
    scenario column is malformed: the superseded NorthSea/Denmark_Germany.csv
    labels its rows 'Transmission v1.0' through 'v1.109' although they belong
    to one scenario.  Several distinct names therefore means the file is
    broken, not that a choice is on offer, and this raises ValueError rather
    than picking one.
    """
    if requested:
        return requested

    scenarios = list_scenarios(source)
    if len(scenarios) == 1:
        return scenarios[0]
    if not scenarios:
        raise ValueError(
            f'No scenario column found in {source}; the file does not look '
            'like an IAMC table. Pass --scenario to override.'
        )
    raise ValueError(
        f'{source} labels its rows with {len(scenarios)} different scenario '
        f'names ({", ".join(scenarios[:3])}...), but an IAMC file describes '
        'one scenario and every row should carry the same name -- the file is '
        'malformed. Fix the scenario column, or pass --scenario / set '
        '"scenario:" in config.yaml to force one.'
    )
