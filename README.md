# EMPIRE North Sea — NUTS2 input data pipeline

Builds input datasets for the [EMPIRE](https://github.com/ntnuiotenergy/OpenEMPIRE) European power-system model in which the five North Sea countries — **Norway, the United Kingdom, Denmark, Germany and the Netherlands** — are resolved into their **102 NUTS2 regions**, while the rest of Europe stays at country level (NUTS0).

The repository holds all the data and code for a two-stage workflow, starting from the four country-level EMPIRE reference datasets of the [Energy Vision](https://energyvis.eu/) scenarios:

1. **Norway aggregation** (`Scripts/preprocess_norway.py`). The reference datasets model Norway as five price zones (NO1–NO5). This stage merges them into one `Norway` node, turning `data/EMPIRE/<scenario>/` into `data/EMPIRE/<scenario>_Norway/`.
2. **NUTS2 disaggregation** (`main.py`). This stage takes a `*_Norway` dataset and a regional IAMC-format file. It replaces each North Sea country node with its NUTS2 regions and distributes that country's capacities, demand, limits, transmission and hourly profiles across them.

**National totals from the reference dataset are conserved** in both stages. The regional data only decide *where* inside a country things go, not *how much* there is.

The outputs of stage 1 are included, so you can start directly at stage 2. Stage 1 can be rerun at any time and reproduces them.

| Country | EMPIRE node replaced | NUTS2 regions |
|---|---|---|
| Norway (`NO`) | `Norway` (NO1–NO5 merged) | 6 |
| United Kingdom (`UK`) | `Great Brit.` | 41 |
| Denmark (`DK`) | `Denmark` | 5 |
| Germany (`DE`) | `Germany` | 38 |
| Netherlands (`NL`) | `Netherlands` | 12 |
| **Total** | | **102** |

The dedicated offshore-wind hub nodes of the reference datasets (Dogger Bank, Hornsea, Moray Firth, Firth of Forth, Outer Dowsing, Norfolk, East Anglia, Borssele, Hollandsee Kust, Helgoländer Bucht, Nordsøen, Utsira Nord, Sørlige Nordsjø I and II) are kept. Their cables are reconnected to the nearest coastal NUTS2 region of the country they land in.

---

## Contents

- [Quick start](#quick-start)
- [Repository layout](#repository-layout)
- [Input data](#input-data)
- [Configuration](#configuration)
- [Running the pipeline](#running-the-pipeline)
- [Output](#output)
- [How the disaggregation works](#how-the-disaggregation-works)
- [Known data gaps and expected warnings](#known-data-gaps-and-expected-warnings)
- [Node-naming pitfalls](#node-naming-pitfalls)
- [Optional features](#optional-features)
- [Tests](#tests)
- [Troubleshooting](#troubleshooting)
- [Data sources](#data-sources)

---

## Quick start

```bash
# 1. Create an environment (Python 3.10+)
conda create -n empire-northsea python=3.11
conda activate empire-northsea
pip install -r requirements.txt

# 2. Stage 1 — merge NO1–NO5 into one Norway node in all four reference datasets
#    (optional: the *_Norway folders are already included)
python Scripts/preprocess_norway.py --all

# 3. Stage 2 — build one NUTS2 dataset (NECP Essentials, settings from config.yaml)
python main.py
```

The result is written to `Result/NorthSea/NO_UK_DK_DE_NL/NorthSea_NECPEssentials/NECPEssentials_Norway/`.

To build all four reference scenarios:

```bash
# bash
for ref in GoRES NECPEssentials REPowerEU++ Trinity; do
  python main.py "${ref}_Norway" --output-folder "NorthSea/NO_UK_DK_DE_NL/NorthSea_${ref}"
done
```

```powershell
# PowerShell
foreach ($ref in 'GoRES','NECPEssentials','REPowerEU++','Trinity') {
  python main.py "${ref}_Norway" --output-folder "NorthSea/NO_UK_DK_DE_NL/NorthSea_$ref"
}
```

Each run takes a few minutes. The ScenarioData step (hourly time series) is the slowest.

---

## Repository layout

```
.
├── main.py                            # Pipeline entry point
├── config.yaml                        # Case configuration (countries, inputs, options)
├── requirements.txt
├── Scripts/
│   ├── config_loader.py               # Reads config.yaml; derives NUTS2 regions, centroids, adjacency
│   ├── iamc_utils.py                  # IAMC loading, scenario detection, hydro weights
│   ├── validate_iamc_data.py          # Pre-flight checks on the IAMC file
│   ├── generate_sets_nuts2.py         # Sets.xlsx
│   ├── generate_generation_nuts2.py   # Generator.xlsx
│   ├── generate_transmission_nuts2.py # Transmission.xlsx
│   ├── generate_node.py               # Node.xlsx
│   ├── generate_storage.py            # Storage.xlsx
│   ├── generate_scenarios.py          # ScenarioData/*.csv (hourly profiles)
│   ├── preprocess_norway.py           # Merges NO1–NO5 into one Norway node (builds *_Norway)
│   ├── preprocess_renewables_atlite.py# Optional: per-region CF series from weather data
│   └── test_hydro_weights.py          # Unit tests for the hydro weighting
└── data/
    ├── EMPIRE/                        # Country-level EMPIRE reference datasets
    │   ├── GoRES/                     #   original, Norway as NO1–NO5 (stage 1 input)
    │   ├── NECPEssentials/
    │   ├── REPowerEU++/
    │   ├── Trinity/
    │   ├── GoRES_Norway/              #   Norway merged (stage 1 output, stage 2 input)
    │   ├── NECPEssentials_Norway/
    │   ├── REPowerEU++_Norway/
    │   └── Trinity_Norway/
    ├── IAMC/NorthSea/North_Sea.xlsx   # Regional capacities, network, demand shares
    ├── Renewable/NorthSea/            # Regional wind/solar resource atlases
    │   ├── onshore.csv
    │   └── Solar.xlsx
    └── shp/
        ├── NUTS_RG_20M_2024_3035.*    # NUTS 2024 boundaries (NO/DK/DE/NL and neighbours)
        └── NUTS_RG_60M_2021_3035.geojson # NUTS 2021 boundaries (fills the UK regions)
```

`Result/` (outputs) and `run_logs/` are created on demand and are git-ignored.

---

## Input data

### EMPIRE reference datasets — `data/EMPIRE/`

Each folder is a complete country-level EMPIRE dataset for one Energy Vision scenario, in two versions:

| Scenario | Original (NO1–NO5) | Norway merged |
|---|---|---|
| Go RES | `GoRES` | `GoRES_Norway` |
| NECP Essentials | `NECPEssentials` | `NECPEssentials_Norway` |
| REPowerEU++ | `REPowerEU++` | `REPowerEU++_Norway` |
| Trinity | `Trinity` | `Trinity_Norway` |

Each contains `General.xlsx`, `Generator.xlsx`, `Node.xlsx`, `Sets.xlsx`, `Storage.xlsx`, `Transmission.xlsx` and a `ScenarioData/` folder of hourly CSVs (`electricload`, `hydroror`, `hydroseasonal`, `solar`, `windonshore`, `windoffshore`, plus `sampling_key`). The original `REPowerEU++` folder has no `sampling_key.csv`.

### Stage 1 — Norway aggregation (`Scripts/preprocess_norway.py`)

Stage 2 needs exactly one node per disaggregated country, but the original datasets split Norway into five price zones. The script merges them into one `Norway` node:

| File | Treatment |
|---|---|
| `Generator.xlsx` | `RefInitialCap`, `InitialCapacity`, `MaxBuiltCapacity`, `MaxInstalledCapacity` summed over NO1–NO5 |
| `Node.xlsx` | `ElectricAnnualDemand`, `HydroGenMaxAnnualProduction`, `BiomassMaxAnnualActivity` summed; `NodeLostLoadCost` taken from NO1 |
| `Storage.xlsx` | All per-node capacity sheets summed |
| `Transmission.xlsx` | Internal NO–NO corridors removed; external corridors sharing an endpoint merged with capacities summed |
| `Sets.xlsx` | NO1–NO5 replaced by `Norway` in every sheet |
| `ScenarioData/` | `electricload.csv` summed; capacity-factor CSVs capacity-weighted; `windoffshore.csv` already has a Norway column and is kept |
| `General.xlsx` | Copied unchanged |

At the end, the script checks its own output. It confirms that no NO1–NO5 names remain, that the CSVs use the `Norway` column name, and that every generator profiled in ScenarioData exists on the Norway node.

```bash
# all four datasets: data/EMPIRE/<scenario> -> data/EMPIRE/<scenario>_Norway
python Scripts/preprocess_norway.py --all

# one dataset (a bare name is looked up under data/EMPIRE/)
python Scripts/preprocess_norway.py NECPEssentials

# explicit input and output folders
python Scripts/preprocess_norway.py data/EMPIRE/NECPEssentials path/to/NECPEssentials_Norway
```

Running it on the included originals reproduces the included `*_Norway` folders cell for cell. Rerunning it overwrites them in place. Git will then list the `.xlsx` files as modified: they are saved again as new files, but their contents are the same. The only file it does not produce is `REPowerEU++_Norway/ScenarioData/sampling_key.csv`, because the original `REPowerEU++` folder has none. The included copy is left untouched by a rerun.

### IAMC regional data — `data/IAMC/NorthSea/North_Sea.xlsx`

An IAMC-format table (model / scenario / region / variable / unit / years) covering the 102 NUTS2 regions. The pipeline uses:

| Variable family | Used for |
|---|---|
| `Capacity\|Electricity\|*` | Regional spatial weights for generation capacity and limits |
| `Capacity\|Electricity\|Hydro\|*` | Hydro weights (reservoir, run-of-river, pumped storage) |
| `Final Energy\|Electricity [Share]` | Regional share of each country's electricity demand |
| `Network\|Electricity\|*` | Transmission corridors between regions (AC and DC), capacity, reactance |

The file holds a single scenario (`Transmission v1.0`), and `main.py` reads the name from the file. If several rows describe the same `(region, variable)` (for example, one row per power plant), they are summed to one NUTS2 total when loaded.

### Renewable resource atlases — `data/Renewable/NorthSea/`

| File | Resolution | Columns used | Used for |
|---|---|---|---|
| `onshore.csv` | NUTS2 | `ref_MW`, `ref_GWh` | Per-region onshore-wind CF, `ref_GWh·1000 / (ref_MW·8760)` |
| `Solar.xlsx` | NUTS2 | `Capacity_Factor` | Per-region solar CF |

Both are pan-European *potential* atlases (optimally sited plants), so their levels sit above EMPIRE's fleet-average CFs. They are therefore used only for the spatial pattern (`mode: relative`, see [Configuration](#configuration)).

### NUTS boundaries — `data/shp/`

Two Eurostat GISCO boundary files (EPSG:3035) are needed, because no single NUTS version covers every region in the case:

| File | Version | Why it is needed |
|---|---|---|
| `NUTS_RG_20M_2024_3035.shp` (+ `.dbf`, `.shx`, `.prj`, `.cpg`) | NUTS 2024, 1:20 million | Has the current codes, including the revised Dutch regions (`NL35`, `NL36`) |
| `NUTS_RG_60M_2021_3035.geojson` | NUTS 2021, 1:60 million | Has the 41 UK regions, which NUTS 2024 dropped after Brexit |

The pipeline looks up each region in the 2024 file first and uses the 2021 file only for regions still missing. In practice, that means the UK regions. The boundaries give region centroids, corridor lengths, coastal distance for offshore reconnection, and land-border adjacency. Without the 2021 file, the UK regions have no geometry. The offshore hubs landing in Great Britain (Dogger Bank, Hornsea, East Anglia, …) then cannot be reconnected, and the transmission step stops with an error.

---

## Configuration

Everything is set in [`config.yaml`](config.yaml). Command-line flags override it.

| Key | Meaning |
|---|---|
| `iamc_file` | Path to the IAMC file (default `data/IAMC/NorthSea/North_Sea.xlsx`). |
| `scenario` | IAMC scenario name. **Leave unset**; it is read from the file. |
| `empire_folder` | Reference dataset under `data/EMPIRE/` (one of the `*_Norway` folders). |
| `output_folder` | Subfolder of `Result/`; the `empire_folder` name is appended to it. |
| `atlite_dir` | Optional folder of per-region weather-based CF series (see [Optional features](#optional-features)). Empty = off. |
| `renewable_cf` | Per-region rescaling of solar/onshore-wind CFs from the atlases (`dir`, `mode`, `countries`). |
| `load_flow` | Toggle and parameters for DC power-flow / transmission-expansion sheets. |
| `nuts2_countries` | The countries to disaggregate, each with the `empire_node` it replaces. |

### `renewable_cf` modes

| Mode | Regional multiplier | National CF / annual energy |
|---|---|---|
| `relative` *(used)* | `regional_CF / capacity-weighted mean regional_CF` | **Preserved** — only the spatial pattern changes |
| `absolute` | `regional_CF / national_CF` | Replaced by the atlas level |

The hourly shape of each series is kept; only its level changes by region. Values are clipped to `[0, 1]`.

### Adding or removing a country

Add an entry under `nuts2_countries` with the two-letter NUTS prefix and the node name it replaces, and make sure the IAMC file has data for that country's NUTS2 regions:

```yaml
nuts2_countries:
  BE:
    empire_node: Belgium
```

Note that `NO` must be quoted (`"NO"`), because YAML 1.1 reads a bare `NO` as `false`.

---

## Running the pipeline

The full chain from the original reference data to a NUTS2 dataset is:

```
data/EMPIRE/<scenario>/            original, Norway as NO1–NO5
        │  Stage 1: Scripts/preprocess_norway.py
        ▼
data/EMPIRE/<scenario>_Norway/     one Norway node
        │  Stage 2: main.py  (+ IAMC file, atlases, NUTS shapes)
        ▼
Result/.../<scenario>_Norway/      NO/UK/DK/DE/NL at NUTS2
```

Stage 1 is described [above](#stage-1--norway-aggregation-scriptspreprocess_norwaypy). Stage 2 is run with:

```
python main.py [empire_folder] [--iamc PATH] [--scenario NAME]
               [--output-folder SUBFOLDER] [--atlite-dir DIR] [--skip-validation]
```

`main.py` runs these steps in order, stopping at the first failure:

```
config.yaml
    │
    ▼
main.py
 0. validate_iamc_data.py          pre-flight checks on the IAMC file
 1. generate_sets_nuts2.py         → Sets.xlsx
 2. generate_generation_nuts2.py   → Generator.xlsx
 3. generate_transmission_nuts2.py → Transmission.xlsx
 4. generate_node.py               → Node.xlsx
 5. generate_storage.py            → Storage.xlsx
 6. generate_scenarios.py          → ScenarioData/*.csv
 7. copy General.xlsx              (+ load-flow sheets when enabled)
```

Every step after Sets reads the generated `Sets.xlsx`, so the node list stays the same across all files. To keep a log:

```bash
mkdir -p run_logs
python main.py NECPEssentials_Norway --output-folder "NorthSea/NO_UK_DK_DE_NL/NorthSea_NECPEssentials" \
  2>&1 | tee run_logs/northsea_NECPEssentials.log
```

---

## Output

```
Result/NorthSea/NO_UK_DK_DE_NL/NorthSea_<scenario>/<scenario>_Norway/
├── Sets.xlsx
├── Generator.xlsx
├── Transmission.xlsx
├── Node.xlsx
├── Storage.xlsx
├── General.xlsx
└── ScenarioData/
    ├── electricload.csv
    ├── hydroror.csv
    ├── hydroseasonal.csv
    ├── solar.csv
    ├── windonshore.csv
    ├── windoffshore.csv
    └── sampling_key.csv
```

The folder is a complete EMPIRE dataset. Point your EMPIRE run configuration at it the same way as at any other EMPIRE dataset.

In every file, the columns/rows of the five North Sea countries are replaced by NUTS2 codes (for example `DE11`, `UKI3`, `NO08`, `NL32`). All other countries keep their EMPIRE node names.

---

## How the disaggregation works

### Sets (`generate_sets_nuts2.py`)

- The five country nodes are replaced by their NUTS2 regions (codes detected from the IAMC file). Coordinates are the region centroids.
- **Corridors.** IAMC network links become EMPIRE corridors. AC and DC links are kept apart, so parallel HVAC and HVDC circuits between the same pair are both kept. Links from a NUTS2 region to a country outside the case are mapped to that country's EMPIRE node.
- **Offshore reconnection.** In the reference, offshore hubs connect to country nodes. When that country is split, the hub is reconnected to the **nearest coastal NUTS2 region of the same country**. "Nearest" means the distance from the platform to each region's polygon, which is effectively the distance to the coastline. These links are typed `HVDC_Cable`.
- **`GeneratorsOfNode`.** Regions get the generators that the IAMC file reports for them. In addition, every *dispatchable* technology the country has in the reference (Bio, Coal, Gas, Oil, Nuclear, Lignite, Waste, Geo, Wave, and the CCS variants) is made available in the country's regions. This way reference capacity and zero-capacity investment options are not lost when the IAMC file omits a technology.
- **Phantom-offshore filter.** EMPIRE reads offshore-wind availability from `windoffshore.csv` column by column. A node holding an offshore-wind generator with no matching column could never produce, so such generators are removed rather than given an invented profile.
- **`Countries` / `NodesOfCountry`.** These map each disaggregated country to its NUTS2 nodes, for the national-level limits below.

### Generation (`generate_generation_nuts2.py`)

- **`RefInitialCap` (existing capacity).** For each country and technology, the reference national total is **conserved** and split across the regions. The weight is the region's IAMC capacity for that technology. If the IAMC file has none for that technology in that country, a fallback is used (a warning is printed):
  - Waste → biomass capacity
  - run-of-river → total hydro capacity
  - anything else → total installed capacity
  - if all of those are zero, an equal split
- **`MaxInstalledCapacity` (resource ceilings).**
  - Solar and onshore wind: the national ceiling is split across regions by IAMC capacity. The sum over regions equals the national value.
  - Hydro: the ceiling is split across the regions that hold hydro, with ×1.5 headroom per region.
  - Offshore wind and wave: the ceiling is split across the regions that already hold the generator (a proxy for "coastal"). This stops inland regions from building offshore capacity.
- **National-level limits.** Build-rate limits and policy ceilings for non-resource technologies (nuclear, gas, coal, …) are national by nature. For the disaggregated countries they move to `MaxBuiltCapacityCountry`, `MaxInstalledCapacityCountry` and `MinBuiltCapacityCountry`, keyed by `(Country, Technology[, Period])`, with the reference values unchanged.
- Cost, efficiency, lifetime and other technology sheets are copied from the reference. Zero efficiencies for renewables are set to 1.0 to avoid division by zero.

### Transmission (`generate_transmission_nuts2.py`)

- One row per physical corridor in `InitialCapacity`, `MaxBuiltCapacity`, `MaxInstallCapacityRaw`, `Length` and `Lifetime`. `lineEfficiency` keeps both directions.
- If the IAMC file gives capacity in one direction only, that value is used. If both directions are given, their average is used.
- Offshore-hub landings inherit the reference capacity of the original hub–country line, so the reconnected cable keeps its calibrated capacity.
- Node names are matched ignoring spaces, so the spelling differences between reference datasets do not break corridors.

### Node (`generate_node.py`)

| Sheet | Countries outside the case | NUTS2 regions |
|---|---|---|
| `ElectricAnnualDemand` | Copied from the reference | Country total × region's `Final Energy\|Electricity [Share]` |
| `NodeLostLoadCost` | Copied | Country value on every region |
| `HydroGenMaxAnnualProduction` | Copied | Country total split by the region's non-pumped hydro capacity (equal split if none) |
| `BiomassMaxAnnualActivity` | Copied (when present) | Moved to `BiomassMaxAnnualActivityCountry` (national budget) |

### Storage (`generate_storage.py`)

- Pumped hydro power and energy capacity are split by each region's IAMC pumped-storage capacity. The power ceiling allows 2× expansion.
- Batteries and the remaining sheets are split evenly or copied from the reference.

### Hourly profiles (`generate_scenarios.py`)

| File | Method |
|---|---|
| `electricload.csv` | Country load × region's demand share (normalised per country) |
| `hydroseasonal.csv` | Inflow split by regional reservoir capacity |
| `hydroror.csv` | Country CF copied to each region holding run-of-river |
| `solar.csv` | Country CF × regional multiplier from `Solar.xlsx` (`renewable_cf`); country CF copied where the atlas has no data |
| `windonshore.csv` | Country CF × regional multiplier from `onshore.csv` (`renewable_cf`); country CF copied where the atlas has no data |
| `windoffshore.csv` | Country CF copied to the regions holding offshore wind (no regional offshore source exists) |

NUTS2 columns are inserted where the country column was, so the order of all other columns is kept. `sampling_key.csv` and any other file are copied unchanged.

---

## Known data gaps and expected warnings

These are known and do not stop the run:

- **Atlas coverage.** `onshore.csv` has **no UK rows** and `Solar.xlsx` has **no Norway rows**, so UK onshore wind and Norwegian solar use the national profile in every region. Coverage by country:

  | File | NO | UK | DK | DE | NL |
  |---|---|---|---|---|---|
  | `onshore.csv` | 6/6 | **0/41** | 5/5 | 38/38 | 12/12 |
  | `Solar.xlsx` | **0/6** | 35/41 | 5/5 | 38/38 | 10/12 |

- **Unplaceable capacity.** Germany's *Hydro regulated* (1,381 MW) and Great Britain's *Wave* (8 MW) cannot be placed, because no NUTS2 region holds those generators. Their national totals come out short by exactly those amounts, and a warning is printed. Totals for NO, DK and NL are conserved exactly.
- **GB offshore wind.** The reference sets `Great Brit.` offshore-wind `MaxInstalledCapacity` to 0. GB offshore expansion belongs to the hub nodes (Dogger Bank, Hornsea, …), not to the country, and the pipeline keeps it that way.
- **Regions with `ref_MW = 0`** in the atlas (for example Svalbard) produce no CF and fall back to the national profile.

Adding UK rows to `onshore.csv` and Norwegian rows to `Solar.xlsx` would close the only outstanding data gap for this case.

---

## Node-naming pitfalls

There are two separate traps with Great Britain. Don't mix them up.

**1. The node name depends on a space.** The reference workbooks disagree:

| Node | GoRES / NECPEssentials / REPowerEU++ | Trinity | What the pipeline emits |
|---|---|---|---|
| Great Britain | `GreatBrit.` | `Great Brit.` | **`Great Brit.`** |
| Czechia | `CzechR` | `Czech R` | **`Czech R`** |

The pipeline rewrites the closed-up forms to the spaced ones (`NODE_NAME_FIXES`), so all four datasets end up with the same names. As a result, **`empire_node` in `config.yaml` must be `Great Brit.`**. With `GreatBrit.` nothing matches, and Great Britain would stay as a country node *next to* its 41 NUTS2 regions, counting the whole country twice.

**2. Column code vs NUTS prefix.** Great Britain is `GB` in the ScenarioData CSV headers but `UK` in NUTS codes, and Norway is `Norway` in the CSVs but `NO` in NUTS. The pipeline always resolves a CSV column through its EMPIRE node name (`GB` → `Great Brit.` → `UK`) instead of comparing strings directly.

The offshore hub names have the same space issue (`DoggerBank` vs `Dogger Bank`) and are normalised the same way.

---

## Optional features

### Load-flow / transmission-expansion sheets

Set `load_flow.enabled: true` to add the sheets needed by EMPIRE's DC power-flow and binary transmission-expansion formulation:

| File | Sheet | Content |
|---|---|---|
| `Transmission.xlsx` | `lineReactance` | Reactance (Ω) of each AC corridor: IAMC `Network\|Electricity\|Reactance` if present, otherwise `length_km × reactance_per_km` |
| `Transmission.xlsx` | `CandidateTransmission` | Corridors that can be expanded |
| `Transmission.xlsx` | `LineBlockCapacity` | Block size per candidate (`block_capacity_mw`) |
| `Transmission.xlsx` | `LineBlockReactance` | Reactance of a new circuit per candidate |
| `Sets.xlsx` | `LineTypeOfCandidateLines` | HVAC / HVDC type of each candidate |
| `General.xlsx` | `NominalVoltage`, `LineBlockCapacityGlobal`, `LineBlockReactanceGlobal` | Scalars and global fallbacks |

Every existing corridor becomes a candidate. Every pair of land-adjacent regions with no corridor also gets a new zero-capacity HVAC candidate. HVDC corridors are left out of Ohm's law. Set `lopf_kwargs.dc_line_types: ["HVDC_Cable"]` in the EMPIRE run configuration.

### Weather-based regional CFs (Atlite)

`Scripts/preprocess_renewables_atlite.py` builds per-region solar and onshore-wind CF series from ERA5 weather data with [atlite](https://github.com/PyPSA/atlite) (`pip install atlite`; downloading ERA5 needs a Copernicus Climate Data Store account). Set `atlite_dir` to the folder with the resulting `solar.csv` / `windonshore.csv`. For any file Atlite provides, these series replace the national ones, and they take priority over `renewable_cf`.

---

## Tests

```bash
cd Scripts
python -m unittest test_hydro_weights -v
```

---

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `ModuleNotFoundError: pyam` | The environment lacks `pyam-iamc`. Run `pip install -r requirements.txt` in the environment you run `main.py` with. |
| `Filtered IamDataFrame is empty` / validation fails | The scenario name, region codes or units in the IAMC file do not match. Run `python Scripts/validate_iamc_data.py <iamc_file> "<scenario>"` for details. |
| A country appears both as a country node and as NUTS2 regions | `empire_node` is misspelled (see [Node-naming pitfalls](#node-naming-pitfalls)). |
| `NO` countries missing / YAML error | `NO` is not quoted in `config.yaml`. |
| `EMPIRE folder not found` | `empire_folder` does not match a folder under `data/EMPIRE/`. The error lists the ones available. |
| Pyomo fails on `genCapAvailStochRaw` when running EMPIRE | An offshore generator has no availability column. Check the phantom-offshore messages in the Sets step. |
| Excel `~$*.xlsx` files appear | Lock files from Excel. Close the workbook; they are git-ignored. |

---

## Data sources

- **EMPIRE model:** [OpenEMPIRE](https://github.com/ntnuiotenergy/OpenEMPIRE), NTNU.
- **Reference scenarios:** Energy Vision scenario datasets (GoRES, NECP Essentials, REPowerEU++, Trinity) — <https://energyvis.eu/>.
- **NUTS boundaries:** © EuroGeographics for the administrative boundaries; Eurostat GISCO, NUTS 2024 and NUTS 2021.
- **Renewable atlases:** pan-European NUTS2 onshore-wind and solar potential datasets (`data/Renewable/NorthSea/`).
