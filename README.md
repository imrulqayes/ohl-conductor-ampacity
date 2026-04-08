# Conductor Ampacity Calculator — IEEE 738

Calculates the thermal rating of overhead power line conductors per the
**IEEE 738** standard. Three calculation modes are supported:

| Mode | Description |
|---|---|
| **Steady-State Thermal Rating** | Maximum current a conductor can carry at a specified temperature limit |
| **Steady-State Temperature** | Conductor temperature when carrying a specified current |
| **Transient Temperature** | Temperature rise during a fault or overload event |

---

## Project Structure

```
conductor_ampacity/
├── conductor_ampacity.py     ← Main script (run this)
├── config.ini                ← All user inputs go here
├── run_ampacity.bat          ← Double-click to run on Windows
├── requirements.txt          ← Python dependencies
├── data/
│   └── conductor_data.xlsx                ← Conductor database
├── output/                   ← Excel results are saved here
├── src/
│   ├── ieee738.py            ← IEEE 738 calculation functions
│   ├── cigre.py              ← CIGRE placeholder (future)
│   ├── conductor_db.py       ← Excel database lookup
│   └── reporter.py           ← Excel report generator
└── previous_codes/           ← Original scripts (reference only)
```

---

## Quick Start

### Step 1 — Edit `config.ini`

Open `config.ini` in any text editor. The file is organised into sections:

#### `[conductor]`
Specify the conductor either by **name** (recommended) or by **manual data**.

```ini
[conductor]
name = Grosbeak     ; <-- change this to your conductor name
```

If the conductor is not in the database, comment out `name` and fill in the
manual fields instead:

```ini
; name =
diameter_mm              = 25.16
resistance_low_ohm_per_km  = 90.22
resistance_high_ohm_per_km = 107.60
heat_capacity_J_per_mK   = 1047
```

#### `[environment]`
Site and weather conditions:

```ini
elevation_m    = 0       ; Elevation above sea level (m)
ambient_temp_C = 40      ; Ambient temperature (°C)
wind_speed_m_s = 0.61    ; Wind speed (m/s) — 0.61 is IEEE 738 conservative default
wind_angle_deg = 90      ; Wind-to-conductor angle (90° = perpendicular)
atmosphere     = clear   ; clear or industrial
```

#### `[solar]`
Geographic and date/time information for solar heat gain:

```ini
latitude_deg     = 23.68       ; Degrees (+ = North, - = South)
date             = 10/06/2023  ; dd/mm/yyyy
hour             = 11          ; 24-h format (12 = solar noon)
line_azimuth_deg = 45          ; Degrees clockwise from true North
```

#### `[conductor_properties]`
Surface optical properties (affects heat exchange):

```ini
emissivity   = 0.7   ; Radiation heat loss (0–1; ~0.7 for aged conductors)
absorptivity = 0.9   ; Solar heat absorption (0.23–0.91)
```

#### `[calculations]`
Choose which calculations to run and set their inputs:

```ini
run_thermal_rating     = true   ; Enable/disable each calculation
run_steady_temperature = true
run_transient          = true

max_conductor_temp_C        = 150        ; Fallback max temp if not in database
steady_current_A            = 900        ; For steady-state temperature calculation
transient_pre_current_A     = auto       ; Pre-fault current (A), or 'auto' = 50% of rating
transient_fault_levels_pct  = 110, 120, 130, 150, 200  ; Fault levels as % of thermal rating
transient_duration_min      = 60         ; Calculation duration (minutes)
transient_interval_s        = 60         ; Time step (seconds)
```

---

### Step 2 — Run the Script

**Option A — Double-click `run_ampacity.bat`** (Windows, easiest)

**Option B — From the terminal:**

```bash
# Run with conductor from config.ini
python conductor_ampacity.py

# Override conductor name on the command line
python conductor_ampacity.py grosbeak
python conductor_ampacity.py Drake

# List all conductors in the database
python conductor_ampacity.py --list

# Console output only, no Excel file
python conductor_ampacity.py --no-output

# Use a different config file
python conductor_ampacity.py --config path/to/other.ini
```

---

### Step 3 — View Results

Results are printed to the console and saved as a formatted Excel file in the
`output/` folder. The file is named `{conductor}_{timestamp}.xlsx`.

The Excel file contains:
- **Summary sheet** — all inputs and key results in a formatted table
- **Transient Profile sheet** — time-series temperature data (if transient was run)

---

## Conductor Database

The Excel file `data/conductor_data.xlsx` is the conductor database.

To add a new conductor, add a new row with these columns:

| Column | Description | Example |
|---|---|---|
| Conductor Name | Name (used for lookup) | Grosbeak |
| Diameter (mm) | Outside diameter | 25.16 |
| R75 (ohm/m) | AC resistance at 75 °C | 0.107612 |
| R25 (ohm/m) | AC resistance at 25 °C | 0.090224 |
| Conductor heat capacity (J/m °C) | Thermal mass | 1047 |

> **Conductor name lookup is case-insensitive but requires exact spelling.**
> `grosbeak`, `Grosbeak`, and `GROSBEAK` all match — but `grosbeaks` does not.

---

## Requirements

- Python 3.10 or later (uses `|` union type hints)
- `pandas` — reading the conductor database Excel file
- `openpyxl` — creating the output Excel report

Install dependencies (if not already installed):

```bash
pip install -r requirements.txt
```

Or with conda (`data_analysis` environment):

```bash
conda activate data_analysis
pip install -r requirements.txt
```

---

## Future: CIGRE Standard

The file `src/cigre.py` is a placeholder for CIGRE thermal rating methods
(CIGRE TB 207 / CIGRE TB 601). Once implemented, switch the standard in
`config.ini`:

```ini
[general]
standard = cigre
```

No other changes are needed — the main script automatically loads the correct
calculation module.

---

## Notes

- All calculations follow **IEEE 738-2012** methodology.
- The IEEE 738 standard recommends using `t_rh >= t_max` (resistance reference
  temperature should be at or above the maximum conductor temperature) for
  conservative results.
- Transient calculation uses the Euler forward-difference method with the time
  step (`transient_interval_s`) you specify.
- A smaller time step improves accuracy but increases calculation time.
