# Conductor Ampacity Calculator — IEEE 738

Calculates the thermal rating of overhead power line conductors per the
**IEEE 738** standard. Three calculation modes are supported:

| Mode | Description |
|---|---|
| **Steady-State Thermal Rating** | Maximum current a conductor can carry at a specified temperature limit |
| **Steady-State Temperature** | Conductor temperature when carrying a specified current |
| **Transient Temperature** | Temperature rise during a fault or overload event |

---

## How to Use

For a step-by-step guide to installing Python, setting up the tool, and running
your first calculation, see **[HOW_TO_USE.md](HOW_TO_USE.md)**.

---

## Project Structure

```
conductor_ampacity/
├── conductor_ampacity.py     ← Main script
├── config.ini                ← All user inputs go here
├── run_ampacity.bat          ← Double-click to run on Windows
├── requirements.txt          ← Python dependencies
├── HOW_TO_USE.md             ← Step-by-step user guide
├── data/
│   └── conductor_data.xlsx   ← Conductor database
├── output/                   ← Excel results are saved here (auto-created)
└── src/
    ├── ieee738.py            ← IEEE 738 calculation functions
    ├── cigre.py              ← CIGRE placeholder (future)
    ├── conductor_db.py       ← Excel database lookup
    └── reporter.py           ← Excel report generator
```

---

## Configuration Reference (`config.ini`)

All calculation inputs are set in `config.ini`. Open it in any text editor.
**You only need to change the conductor name** — all other parameters have
working default values based on IEEE 738 standard conditions.

### `[conductor]`

Set the conductor name. The spelling must match an entry in `data/conductor_data.xlsx`
exactly, but capitalisation does not matter (`grosbeak`, `Grosbeak`, and `GROSBEAK`
all work).

```ini
[conductor]
name = Grosbeak     ; change this to your conductor name
```

### `[environment]`

Site and weather conditions:

```ini
elevation_m    = 0       ; Elevation above sea level (m)
ambient_temp_C = 40      ; Ambient temperature (°C)
wind_speed_m_s = 0.61    ; Wind speed (m/s) — 0.61 is IEEE 738 conservative default
wind_angle_deg = 90      ; Wind-to-conductor angle (90° = perpendicular)
atmosphere     = clear   ; clear or industrial
```

### `[solar]`

Geographic and date/time information for solar heat gain:

```ini
latitude_deg     = 23.68       ; Degrees (+ = North, - = South)
date             = 10/06/2023  ; dd/mm/yyyy
hour             = 11          ; 24-h format (12 = solar noon)
line_azimuth_deg = 45          ; Degrees clockwise from true North
```

### `[conductor_properties]`

Surface optical properties (affects heat exchange):

```ini
emissivity   = 0.7   ; Radiation heat loss (0–1; ~0.7 for aged conductors)
absorptivity = 0.9   ; Solar heat absorption (0.23–0.91)
```

### `[calculations]`

Choose which calculations to run and set their inputs:

```ini
run_thermal_rating     = true   ; Enable/disable each calculation
run_steady_temperature = true
run_transient          = true

steady_current_A            = 900        ; For steady-state temperature calculation
transient_pre_current_A     = auto       ; Pre-fault current (A), or 'auto' = 50% of rating
transient_fault_levels_pct  = 110, 120, 130, 150, 200  ; Fault levels as % of thermal rating
transient_duration_min      = 60         ; Calculation duration (minutes)
transient_interval_s        = 60         ; Time step (seconds)
```

---

## Output

Results are saved as a formatted Excel file in the `output/` folder, named
`{ConductorName}_{timestamp}.xlsx`. The file contains:

- **Summary sheet** — all inputs and key results in a formatted table
- **Transient Profile sheet** — time-series temperature data and chart (if transient was run)

---

## Conductor Database

The Excel file `data/conductor_data.xlsx` is the conductor database.

To add a new conductor, add a new row with these columns:

| Column | Description | Example |
|---|---|---|
| Conductor Name | Name used for lookup | Grosbeak |
| Diameter (mm) | Outside diameter | 25.16 |
| R75 (ohm/m) | AC resistance at 75 °C | 0.107612 |
| R25 (ohm/m) | AC resistance at 25 °C | 0.090224 |
| Conductor heat capacity (J/m °C) | Thermal mass | 1047 |

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
