"""
Conductor Ampacity Calculator — IEEE 738 / CIGRE
=================================================
Calculates overhead conductor thermal ratings per IEEE 738 standard.

Usage
-----
  python conductor_ampacity.py                   # uses config.ini
  python conductor_ampacity.py grosbeak          # override conductor name
  python conductor_ampacity.py --list            # list available conductors
  python conductor_ampacity.py --config my.ini   # use a custom config file
  python conductor_ampacity.py --no-output       # print results only, no Excel file

The configuration file (config.ini) controls all parameters.
Conductor data is looked up from the Excel database in the data/ folder.

Author       : Imrul Qayes
Email        : imrul27@gmail.com
Date Created : 03/2025
Last Modified: 04/2026
Version      : 1.0.0
AI Assistant : Developed with the assistance of Claude (Anthropic) — https://claude.ai
"""

# Licensed under the MIT License. See LICENSE file in the project root for details.

import argparse
import configparser
import sys
from pathlib import Path

# Ensure the project root is on the path so 'src' is importable
sys.path.insert(0, str(Path(__file__).parent))

from src.conductor_db import get_conductor_data, list_conductors
from src.reporter import create_report

CONFIG_FILE = Path(__file__).parent / "config.ini"


# ── Configuration helpers ─────────────────────────────────────────────────────

def load_config(config_path: Path) -> configparser.ConfigParser:
    if not config_path.exists():
        print(f"[ERROR] Configuration file not found: {config_path}")
        print("        Please ensure config.ini exists in the project root.")
        sys.exit(1)

    cfg = configparser.ConfigParser(inline_comment_prefixes=(";",))
    cfg.read(config_path, encoding="utf-8")
    return cfg


def _float(cfg, section, key, fallback=None):
    try:
        return cfg.getfloat(section, key)
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError):
        if fallback is not None:
            return fallback
        raise


def _bool(cfg, section, key, fallback=None):
    try:
        return cfg.getboolean(section, key)
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError):
        if fallback is not None:
            return fallback
        raise


def _str(cfg, section, key, fallback=None):
    try:
        return cfg.get(section, key).strip()
    except (configparser.NoSectionError, configparser.NoOptionError):
        if fallback is not None:
            return fallback
        raise


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    # ── CLI arguments ──────────────────────────────────────────────────────
    parser = argparse.ArgumentParser(
        description="IEEE 738 Conductor Ampacity Calculator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "conductor", nargs="?", default=None,
        help="Conductor name to look up (overrides config.ini). "
             "Must match the database spelling exactly (case-insensitive).",
    )
    parser.add_argument(
        "--config", default=str(CONFIG_FILE), metavar="FILE",
        help="Path to configuration file (default: config.ini)",
    )
    parser.add_argument(
        "--list", action="store_true",
        help="List all conductor names available in the database and exit.",
    )
    parser.add_argument(
        "--no-output", action="store_true",
        help="Print results to the console only; do not create an Excel file.",
    )
    args = parser.parse_args()

    # ── List conductors ────────────────────────────────────────────────────
    if args.list:
        try:
            names = list_conductors()
        except FileNotFoundError as e:
            print(f"[ERROR] {e}")
            sys.exit(1)
        print("\nAvailable conductors in database:")
        for n in names:
            print(f"  {n}")
        print()
        sys.exit(0)

    # ── Load configuration ─────────────────────────────────────────────────
    cfg = load_config(Path(args.config))

    # ── Load calculation standard ──────────────────────────────────────────
    standard = _str(cfg, "general", "standard", fallback="ieee738").lower()
    if standard == "ieee738":
        import src.ieee738 as calc
        standard_label = "IEEE 738"
    elif standard == "cigre":
        import src.cigre as calc
        standard_label = "CIGRE"
    else:
        print(f"[WARNING] Unknown standard '{standard}' in config.ini. Defaulting to ieee738.")
        import src.ieee738 as calc
        standard_label = "IEEE 738"

    print(f"\n{'='*60}")
    print(f"  Conductor Ampacity Calculator  ({standard_label})")
    print(f"{'='*60}\n")

    # ── Get conductor data ─────────────────────────────────────────────────
    conductor_name = args.conductor or _str(cfg, "conductor", "name", fallback="").strip()

    if conductor_name:
        print(f"Looking up conductor: '{conductor_name}' ...")
        try:
            conductor = get_conductor_data(conductor_name)
        except (FileNotFoundError, ValueError) as e:
            print(f"\n[ERROR] {e}")
            sys.exit(1)
        print(f"  Found: {conductor['name']}")
        print(f"  Diameter : {conductor['diameter_mm']} mm")
        print(f"  R (low)  : {conductor['r_low']:.6f} Ohm/m  @ {conductor['temp_low']:.0f} degC")
        print(f"  R (high) : {conductor['r_high']:.6f} Ohm/m  @ {conductor['temp_high']:.0f} degC")
        if conductor.get("heat_capacity"):
            print(f"  Heat Cap : {conductor['heat_capacity']} J/(m.degC)")
        if conductor.get("max_temp") is not None:
            print(f"  Max Temp : {conductor['max_temp']:.0f} degC  (from database)")
    else:
        # Manual conductor data from config.ini
        try:
            conductor = {
                "name":          "Manual (config.ini)",
                "diameter_mm":   _float(cfg, "conductor", "diameter_mm"),
                "temp_low":      _float(cfg, "conductor", "temp_low_C",  fallback=25.0),
                "temp_high":     _float(cfg, "conductor", "temp_high_C", fallback=75.0),
                "r_low":         _float(cfg, "conductor", "resistance_low_ohm_per_km")  / 1000.0,
                "r_high":        _float(cfg, "conductor", "resistance_high_ohm_per_km") / 1000.0,
                "heat_capacity": _float(cfg, "conductor", "heat_capacity_J_per_mK", fallback=None),
                "max_temp":      _float(cfg, "conductor", "max_allowable_temp_C",    fallback=None),
            }
        except (configparser.NoSectionError, configparser.NoOptionError) as e:
            print(f"\n[ERROR] No conductor name given and manual data is incomplete in config.ini.\n"
                  f"        Missing: {e}\n"
                  f"        Either set 'name = <conductor>' under [conductor], "
                  f"or fill in all manual fields.")
            sys.exit(1)
        print("Using manual conductor data from config.ini.")

    # ── Environmental parameters ───────────────────────────────────────────
    env = {
        "elevation":    _float(cfg, "environment", "elevation_m",      fallback=0.0),
        "ambient_temp": _float(cfg, "environment", "ambient_temp_C"),
        "wind_speed":   _float(cfg, "environment", "wind_speed_m_s",   fallback=0.61),
        "wind_angle":   _float(cfg, "environment", "wind_angle_deg",   fallback=90.0),
        "atmosphere":   _str  (cfg, "environment", "atmosphere",       fallback="clear"),
    }

    # ── Solar parameters ───────────────────────────────────────────────────
    solar = {
        "latitude":      _float(cfg, "solar", "latitude_deg"),
        "date":          _str  (cfg, "solar", "date"),
        "hour":          _float(cfg, "solar", "hour"),
        "line_azimuth":  _float(cfg, "solar", "line_azimuth_deg"),
    }

    # ── Optical properties ─────────────────────────────────────────────────
    optical = {
        "emissivity":    _float(cfg, "conductor_properties", "emissivity",    fallback=0.7),
        "absorptivity":  _float(cfg, "conductor_properties", "absorptivity",  fallback=0.9),
    }

    print(f"\nEnvironment : elevation={env['elevation']} m | "
          f"ambient={env['ambient_temp']} degC | "
          f"wind={env['wind_speed']} m/s @ {env['wind_angle']} deg | "
          f"atmosphere={env['atmosphere']}")
    print(f"Solar       : lat={solar['latitude']} deg | date={solar['date']} | "
          f"hour={solar['hour']}:00 | line azimuth={solar['line_azimuth']} deg")
    print(f"Surface     : emissivity={optical['emissivity']} | "
          f"absorptivity={optical['absorptivity']}\n")

    # ── Build shared kwargs for calculation functions ──────────────────────
    base_kwargs = dict(
        elev       = env["elevation"],
        t_a        = env["ambient_temp"],
        dia        = conductor["diameter_mm"] / 1000.0,
        lat        = solar["latitude"],
        dmy        = solar["date"],
        h24        = solar["hour"],
        z_l        = solar["line_azimuth"],
        t_rh       = conductor["temp_high"],
        t_rl       = conductor["temp_low"],
        r_h        = conductor["r_high"],
        r_l        = conductor["r_low"],
        phi_d      = env["wind_angle"],
        emsvty     = optical["emissivity"],
        alpha      = optical["absorptivity"],
        wnd_spd    = env["wind_speed"],
        atmosphere = env["atmosphere"],
    )

    # ── Which calculations to run ──────────────────────────────────────────
    run_rating  = _bool(cfg, "calculations", "run_thermal_rating",    fallback=True)
    run_temp    = _bool(cfg, "calculations", "run_steady_temperature", fallback=True)
    run_trans   = _bool(cfg, "calculations", "run_transient",         fallback=True)

    calc_settings = {}
    results = {}

    # ── Solar heat gain (for reporting) ───────────────────────────────────
    try:
        q_solar = calc.calculate_solar_heat_gain(
            alpha      = optical["absorptivity"],
            lat        = solar["latitude"],
            dmy        = solar["date"],
            h24        = solar["hour"],
            elev       = env["elevation"],
            z_l        = solar["line_azimuth"],
            dia        = conductor["diameter_mm"] / 1000.0,
            atmosphere = env["atmosphere"],
        )
        dia_m = conductor["diameter_mm"] / 1000.0
        results["solar_heat_gain"] = round(q_solar / dia_m, 2)
        print(f"Solar Heat Gain  : {results['solar_heat_gain']:.2f} W/m²")
    except Exception as e:
        results["solar_heat_gain"] = None
        print(f"  [WARNING] Could not compute solar heat gain: {e}")
    print()

    # ── 1. Steady-State Thermal Rating ─────────────────────────────────────
    # Priority: conductor database max_temp > config.ini max_conductor_temp_C
    db_max_temp = conductor.get("max_temp")
    if db_max_temp is not None:
        max_t = db_max_temp
        max_t_source = f"from conductor database ({conductor['name']})"
    else:
        max_t = _float(cfg, "calculations", "max_conductor_temp_C", fallback=80.0)
        max_t_source = "from config.ini"

    calc_settings["max_conductor_temp_C"] = max_t

    # Always compute the thermal rating — it may be needed as the auto pre-fault current baseline
    print(f"[1/3] Steady-State Thermal Rating  (max temp = {max_t} degC, {max_t_source}) ...")
    try:
        rating = calc.calculate_steady_thermal_rating(t_s=max_t, **base_kwargs)
        results["thermal_rating"] = rating
        if run_rating:
            print(f"      -> Maximum allowable current : {rating} A\n")
        else:
            print(f"      -> {rating} A  [used internally for auto pre-fault current]\n")
    except Exception as e:
        print(f"      [ERROR] {e}\n")
        results["thermal_rating"] = None
        rating = None

    if not run_rating:
        results["thermal_rating"] = None   # hide from report if disabled

    # ── 2. Steady-State Temperature ────────────────────────────────────────
    if run_temp:
        current = _float(cfg, "calculations", "steady_current_A", fallback=900.0)
        calc_settings["steady_current_A"] = current

        print(f"[2/3] Steady-State Conductor Temperature  (current = {current} A) ...")
        try:
            temperature = calc.calculate_steady_temperature(amps=current, **base_kwargs)
            results["steady_temperature"] = temperature
            print(f"      -> Conductor temperature : {temperature} degC\n")
        except Exception as e:
            print(f"      [ERROR] {e}\n")
            results["steady_temperature"] = None

    # ── 3. Transient Temperature ───────────────────────────────────────────
    if run_trans:
        # Pre-fault current: explicit value, or 'auto' = 50 % of thermal rating
        pre_current_raw = _str(cfg, "calculations", "transient_pre_current_A", fallback="auto")
        if pre_current_raw.strip().lower() == "auto":
            if rating is not None:
                ini_amps = round(rating * 0.5, 1)
                print(f"      [auto] Pre-fault current set to 50% of thermal rating: {ini_amps} A")
            else:
                print("      [WARNING] 'auto' pre-fault current requested but thermal rating "
                      "could not be computed. Falling back to 0 A (conductor at ambient).")
                ini_amps = 0.0
        else:
            try:
                ini_amps = float(pre_current_raw)
            except ValueError:
                print(f"      [WARNING] Invalid transient_pre_current_A value '{pre_current_raw}'. "
                      "Using 'auto' (50% of rating).")
                ini_amps = round(rating * 0.5, 1) if rating else 0.0

        duration  = _float(cfg, "calculations", "transient_duration_min",  fallback=60.0)
        interval  = _float(cfg, "calculations", "transient_interval_s",    fallback=60.0)

        # Parse fault current levels (percentages of thermal rating)
        levels_raw = _str(cfg, "calculations", "transient_fault_levels_pct",
                          fallback="110, 120, 130, 150, 200")
        try:
            fault_pcts = [float(p.strip()) for p in levels_raw.split(",") if p.strip()]
        except ValueError:
            print(f"      [WARNING] Invalid transient_fault_levels_pct '{levels_raw}'. "
                  "Using defaults: 110, 120, 130, 150, 200.")
            fault_pcts = [110, 120, 130, 150, 200]

        calc_settings["transient_pre_current_A"] = ini_amps
        calc_settings["transient_fault_pcts"]    = fault_pcts
        calc_settings["transient_rating_A"]      = rating  # base for % calculations

        hc = conductor.get("heat_capacity")
        if hc is None:
            print("[3/3] Transient -- SKIPPED (heat capacity not available for this conductor).\n")
            results["transient"] = None
        else:
            pct_str = ", ".join(f"{p:.0f}%" for p in fault_pcts)
            print(f"[3/3] Transient Temperature  "
                  f"(pre-fault: {ini_amps} A | fault levels: {pct_str} of {rating} A | "
                  f"{duration:.0f} min | dt={interval:.0f} s)")

            curves = []
            for pct in fault_pcts:
                tran_amps = round(rating * pct / 100.0, 1)
                try:
                    time_series, temp_series = calc.calculate_transient_temperature(
                        ini_amps   = ini_amps,
                        trans_amps = tran_amps,
                        mcp        = hc,
                        tran_dur   = int(duration),
                        del_t      = int(interval),
                        **base_kwargs,
                    )
                    peak_temp = max(temp_series)
                    peak_idx  = temp_series.index(peak_temp)
                    peak_min  = time_series[peak_idx] / 60.0
                    curves.append({
                        "pct":           pct,
                        "tran_amps":     tran_amps,
                        "time_s":        time_series,
                        "temp_c":        temp_series,
                        "peak_temp":     peak_temp,
                        "peak_time_min": peak_min,
                        "ini_temp":      temp_series[0],
                        "final_temp":    temp_series[-1],
                    })
                    print(f"      {pct:.0f}% ({tran_amps:.0f} A)  "
                          f"->  peak: {peak_temp} degC  at {peak_min:.1f} min")
                except Exception as e:
                    print(f"      {pct:.0f}% ({tran_amps:.0f} A)  ->  [ERROR] {e}")

            results["transient"] = curves if curves else None
            print()

    # ── Print summary ──────────────────────────────────────────────────────
    print("-" * 60)
    print("  Results Summary")
    print("-" * 60)
    if results.get("thermal_rating") is not None:
        print(f"  Thermal Rating        : {results['thermal_rating']} A  "
              f"(at {calc_settings.get('max_conductor_temp_C', '?')} degC)")
    if results.get("steady_temperature") is not None:
        print(f"  Steady Temperature    : {results['steady_temperature']} degC  "
              f"(at {calc_settings.get('steady_current_A', '?')} A)")
    if results.get("transient"):
        print(f"  Transient (pre-fault {calc_settings.get('transient_pre_current_A')} A):")
        for curve in results["transient"]:
            print(f"    {curve['pct']:.0f}%  ({curve['tran_amps']:.0f} A)"
                  f"  ->  peak {curve['peak_temp']} degC  at {curve['peak_time_min']:.1f} min")
    print("-" * 60)

    # ── Save Excel report ──────────────────────────────────────────────────
    if not args.no_output:
        output_folder = _str(cfg, "output", "output_folder", fallback="output")
        output_path   = Path(__file__).parent / output_folder

        try:
            report_path = create_report(
                output_folder  = output_path,
                conductor      = conductor,
                env            = env,
                solar          = solar,
                optical        = optical,
                calc_settings  = calc_settings,
                results        = results,
                standard       = standard_label,
            )
            print(f"\n  Excel report saved to:\n  {report_path}\n")
        except Exception as e:
            print(f"\n[ERROR] Could not save Excel report: {e}\n")
    else:
        print()


if __name__ == "__main__":
    main()
