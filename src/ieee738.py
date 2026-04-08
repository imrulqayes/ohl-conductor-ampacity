"""
IEEE 738 Standard — Overhead Conductor Thermal Rating Calculations.

Implements the IEEE 738-2012/2023 standard methods for calculating:
  - Steady-state thermal rating (maximum allowable current)
  - Steady-state conductor temperature (for a given current)
  - Transient conductor temperature (during fault/overload events)

All functions use SI units unless otherwise noted in individual docstrings.

Author       : Imrul Qayes
Email        : imrul27@gmail.com
Date Created : 03/2025
Last Modified: 04/2026
Version      : 1.0.0
AI Assistant : Developed with the assistance of Claude (Anthropic) — https://claude.ai
"""

# Licensed under the MIT License. See LICENSE file in the project root for details.

import math
from datetime import datetime


def calculate_air_density(elev, t_s, t_a):
    """
    Calculate air density based on elevation and film temperature.

    Film temperature is the average of surface and ambient temperatures,
    representing the boundary layer temperature.

    Parameters:
        elev  : Elevation above sea level (m)
        t_s   : Conductor surface temperature (°C)
        t_a   : Ambient air temperature (°C)

    Returns:
        Air density (kg/m³)
    """
    if t_s < t_a:
        raise ValueError("Conductor surface temperature must be >= ambient temperature.")

    t_film = (t_s + t_a) / 2
    numerator = 1.293 - 1.524e-4 * elev + 6.373e-9 * elev**2
    denominator = 1 + 3.67e-3 * t_film

    return numerator / denominator


def calculate_natural_convection(elev, t_s, t_a, dia):
    """
    Calculate natural convection heat loss per unit length (IEEE 738).

    Parameters:
        elev  : Elevation above sea level (m)
        t_s   : Conductor surface temperature (°C)
        t_a   : Ambient air temperature (°C)
        dia   : Outside diameter of conductor (m)

    Returns:
        Natural convection heat loss (W/m)
    """
    ro_f = calculate_air_density(elev, t_s, t_a)
    return 3.645 * ro_f**0.5 * dia**0.75 * (t_s - t_a) ** 1.25


def calculate_air_viscosity(t_s, t_a):
    """
    Calculate absolute (dynamic) viscosity of air at film temperature.

    Parameters:
        t_s : Conductor surface temperature (°C)
        t_a : Ambient air temperature (°C)

    Returns:
        Air viscosity (kg/m·s)
    """
    t_film = (t_s + t_a) / 2
    numerator = 1.458e-6 * (t_film + 273) ** 1.5
    denominator = t_film + 383.4

    return numerator / denominator


def calculate_reynolds_number(elev, t_s, t_a, dia, wnd_spd):
    """
    Calculate the dimensionless Reynolds number (IEEE 738).

    Parameters:
        elev    : Elevation above sea level (m)
        t_s     : Conductor surface temperature (°C)
        t_a     : Ambient air temperature (°C)
        dia     : Outside diameter of conductor (m)
        wnd_spd : Wind speed at conductor (m/s)

    Returns:
        Reynolds number (dimensionless)
    """
    ro_f = calculate_air_density(elev, t_s, t_a)
    miu_f = calculate_air_viscosity(t_s, t_a)

    return (dia * ro_f * wnd_spd) / miu_f


def calculate_forced_convection(elev, t_s, t_a, dia, wnd_spd, phi_d):
    """
    Calculate forced convection heat loss per unit length (IEEE 738).

    Parameters:
        elev    : Elevation above sea level (m)
        t_s     : Conductor surface temperature (°C)
        t_a     : Ambient air temperature (°C)
        dia     : Outside diameter of conductor (m)
        wnd_spd : Wind speed at conductor (m/s)
        phi_d   : Angle between wind direction and conductor axis (degrees)

    Returns:
        Forced convection heat loss (W/m)
    """
    n_re = calculate_reynolds_number(elev, t_s, t_a, dia, wnd_spd)

    phi_r = math.radians(phi_d)
    k_an = 1.194 - math.cos(phi_r) + 0.194 * math.cos(2 * phi_r) + 0.368 * math.sin(2 * phi_r)

    t_film = (t_s + t_a) / 2
    k_f = 2.424e-2 + 7.477e-5 * t_film - 4.407e-9 * t_film**2

    qc1 = k_an * (1.01 + 1.35 * n_re**0.52) * k_f * (t_s - t_a)
    qc2 = k_an * 0.754 * n_re**0.6 * k_f * (t_s - t_a)

    return max(qc1, qc2)


def calculate_radiation(t_s, t_a, dia, emsvty):
    """
    Calculate radiated heat loss per unit length (IEEE 738).

    Parameters:
        t_s    : Conductor surface temperature (°C)
        t_a    : Ambient air temperature (°C)
        dia    : Outside diameter of conductor (m)
        emsvty : Emissivity (0 to 1)

    Returns:
        Radiated heat loss (W/m)
    """
    if t_s < t_a:
        raise ValueError("Conductor surface temperature must be >= ambient temperature.")

    return 17.8e-8 * dia * emsvty * ((t_s + 273)**4 - (t_a + 273)**4)


def calculate_hour_angle(h24):
    """
    Calculate hour angle relative to solar noon (IEEE 738).

    Parameters:
        h24 : Hour of day in 24-hour format (e.g., 11 for 11 AM, 14 for 2 PM)

    Returns:
        Hour angle (degrees); negative before noon, positive after noon
    """
    return (h24 - 12) * 15


def calculate_solar_declination(dmy):
    """
    Calculate solar declination angle (-23.45° to +23.45°) (IEEE 738).

    Parameters:
        dmy : Date as string "dd/mm/yyyy" or integer day-of-year (1–365)

    Returns:
        Solar declination (degrees)
    """
    if isinstance(dmy, (int, float)):
        nd_y = int(dmy)
    elif isinstance(dmy, str):
        date_obj = datetime.strptime(dmy, "%d/%m/%Y")
        nd_y = date_obj.timetuple().tm_yday
    else:
        raise ValueError("Date must be a 'dd/mm/yyyy' string or integer day-of-year (1–365).")

    an_rad = math.radians((284 + nd_y) * 360 / 365)
    return 23.46 * math.sin(an_rad)


def calculate_sun_altitude(lat, dmy, h24):
    """
    Calculate sun altitude angle (0° to 90°) (IEEE 738).

    Parameters:
        lat : Latitude (degrees); positive = Northern, negative = Southern hemisphere
        dmy : Date as "dd/mm/yyyy" string or integer day-of-year
        h24 : Hour of day in 24-hour format

    Returns:
        Sun altitude angle (degrees)
    """
    delta_rad = math.radians(calculate_solar_declination(dmy))
    h_a_rad = math.radians(calculate_hour_angle(h24))
    lat_rad = math.radians(lat)

    h_c_rad = math.asin(
        math.sin(lat_rad) * math.sin(delta_rad)
        + math.cos(lat_rad) * math.cos(delta_rad) * math.cos(h_a_rad)
    )

    return math.degrees(h_c_rad)


def calculate_total_solarheat_radiation(lat, dmy, h24, elev, atmosphere):
    """
    Calculate total solar and sky radiated heat intensity corrected for elevation (IEEE 738).

    Parameters:
        lat        : Latitude (degrees)
        dmy        : Date as "dd/mm/yyyy" string or integer day-of-year
        h24        : Hour of day in 24-hour format
        elev       : Elevation above sea level (m)
        atmosphere : 'clear' or 'industrial'

    Returns:
        Elevation-corrected solar and sky radiated heat intensity (W/m²)
    """
    coeff_clear = [-42.2391, 63.8044, -1.9220, 3.46921e-2,
                   -3.61118e-4, 1.94318e-6, -4.07608e-9]
    coeff_industrial = [53.1821, 14.2110, 6.6138e-1, -3.1658e-2,
                        5.4654e-4, -4.3446e-6, 1.3236e-8]

    coeff = coeff_industrial if atmosphere.lower() == 'industrial' else coeff_clear

    h_c = calculate_sun_altitude(lat, dmy, h24)
    q_s = (coeff[0] + coeff[1] * h_c + coeff[2] * h_c**2 + coeff[3] * h_c**3
           + coeff[4] * h_c**4 + coeff[5] * h_c**5 + coeff[6] * h_c**6)

    k_solar = 1 + 1.148e-4 * elev + 1.108e-8 * elev**2

    return q_s * k_solar


def calculate_solar_azimuth_variable(lat, dmy, h24):
    """
    Calculate solar azimuth variable (IEEE 738).

    Parameters:
        lat : Latitude (degrees)
        dmy : Date as "dd/mm/yyyy" string or integer day-of-year
        h24 : Hour of day in 24-hour format

    Returns:
        Solar azimuth variable (dimensionless)
    """
    delta_rad = math.radians(calculate_solar_declination(dmy))
    h_a_rad = math.radians(calculate_hour_angle(h24))
    lat_rad = math.radians(lat)

    numerator = math.sin(h_a_rad)
    denominator = math.sin(lat_rad) * math.cos(h_a_rad) - math.cos(lat_rad) * math.tan(delta_rad)

    return numerator / denominator


def calculate_solar_azimuth(lat, dmy, h24):
    """
    Calculate solar azimuth angle (IEEE 738).

    Parameters:
        lat : Latitude (degrees)
        dmy : Date as "dd/mm/yyyy" string or integer day-of-year
        h24 : Hour of day in 24-hour format

    Returns:
        Solar azimuth (degrees)
    """
    gama = calculate_solar_azimuth_variable(lat, dmy, h24)
    h_a = calculate_hour_angle(h24)

    sol_azi_con = 0
    if h_a < 0 and h_a >= -180:
        if gama < 0:
            sol_azi_con = 180
    else:
        if gama < 0:
            sol_azi_con = 360
        else:
            sol_azi_con = 180

    return sol_azi_con + math.degrees(math.atan(gama))


def calculate_solar_heat_gain(alpha, lat, dmy, h24, elev, z_l, dia, atmosphere='clear'):
    """
    Calculate solar heat gain per unit length (IEEE 738).

    Parameters:
        alpha      : Solar absorptivity (0.23 to 0.91)
        lat        : Latitude (degrees)
        dmy        : Date as "dd/mm/yyyy" string or integer day-of-year
        h24        : Hour of day in 24-hour format
        elev       : Elevation above sea level (m)
        z_l        : Line azimuth (degrees, clockwise from true North)
        dia        : Outside diameter of conductor (m)
        atmosphere : 'clear' or 'industrial' (default: 'clear')

    Returns:
        Solar heat gain rate (W/m); 0 if result is negative (nighttime)
    """
    q_se = calculate_total_solarheat_radiation(lat, dmy, h24, elev, atmosphere)

    z_c = calculate_solar_azimuth(lat, dmy, h24)
    h_c = calculate_sun_altitude(lat, dmy, h24)
    h_c_rad = math.radians(h_c)

    theta_rad = math.acos(math.cos(h_c_rad) * math.cos(math.radians(z_c - z_l)))
    q_solar = alpha * q_se * math.sin(theta_rad) * dia

    return max(0.0, q_solar)


def calculate_resistance_at_temperature(t_rh, t_rl, r_h, r_l, t_s):
    """
    Calculate conductor AC resistance at a given temperature using linear interpolation (IEEE 738).

    Note: Linear interpolation is accurate between T_low and T_high. Beyond T_high,
    the calculated resistance may be slightly low (non-conservative for rating calculations).
    Per IEEE 738, T_high should be >= the maximum conductor temperature in the calculation.

    Parameters:
        t_rh : High reference temperature (°C)
        t_rl : Low reference temperature (°C)
        r_h  : AC resistance at high reference temperature (Ω/m)
        r_l  : AC resistance at low reference temperature (Ω/m)
        t_s  : Target temperature (°C) at which resistance is needed

    Returns:
        AC resistance at target temperature (Ω/m)
    """
    if t_rh == t_rl:
        raise ValueError("High and low reference temperatures must be different.")

    return r_l + (r_h - r_l) * (t_s - t_rl) / (t_rh - t_rl)


def calculate_steady_thermal_rating(
        elev, t_s, t_a, dia, lat, dmy, h24, z_l,
        t_rh, t_rl, r_h, r_l,
        phi_d=90, emsvty=0.7, alpha=0.8, wnd_spd=0.61, atmosphere='clear'):
    """
    Calculate steady-state thermal rating (maximum allowable current) (IEEE 738).

    Parameters:
        elev       : Elevation above sea level (m)
        t_s        : Maximum allowable conductor surface temperature (°C)
        t_a        : Ambient air temperature (°C)
        dia        : Outside diameter of conductor (m)
        lat        : Latitude (degrees; + = Northern, - = Southern hemisphere)
        dmy        : Date as "dd/mm/yyyy" string or integer day-of-year
        h24        : Hour of day in 24-hour format
        z_l        : Line azimuth (degrees, clockwise from true North)
        t_rh       : High reference temperature for resistance (°C)
        t_rl       : Low reference temperature for resistance (°C)
        r_h        : AC resistance at high reference temperature (Ω/m)
        r_l        : AC resistance at low reference temperature (Ω/m)
        phi_d      : Angle between wind and conductor axis (degrees; default 90)
        emsvty     : Emissivity (default 0.7)
        alpha      : Solar absorptivity (default 0.8)
        wnd_spd    : Wind speed (m/s; default 0.61 — conservative low-wind condition)
        atmosphere : 'clear' or 'industrial' (default 'clear')

    Returns:
        Steady-state thermal rating (A), rounded to nearest ampere
    """
    q_cn = calculate_natural_convection(elev, t_s, t_a, dia)
    q_f = calculate_forced_convection(elev, t_s, t_a, dia, wnd_spd, phi_d)
    q_c = max(q_cn, q_f)

    q_r = calculate_radiation(t_s, t_a, dia, emsvty)
    q_s = calculate_solar_heat_gain(alpha, lat, dmy, h24, elev, z_l, dia, atmosphere)

    r_d = calculate_resistance_at_temperature(t_rh, t_rl, r_h, r_l, t_s)

    if not isinstance(r_d, (int, float)):
        raise TypeError(f"Resistance must be numeric. Got: {type(r_d)}")
    if r_d <= 0:
        raise ValueError(f"Resistance must be positive. Got: {r_d}")

    thermal_power = q_c + q_r - q_s
    if thermal_power < 0:
        raise ValueError(
            f"Net heat loss (q_c + q_r - q_s = {thermal_power:.3f} W/m) is negative. "
            "Solar heating exceeds convection + radiation at the given conditions."
        )

    return round(math.sqrt(thermal_power / r_d))


def _muller_method(T0, T1, T2, f0, f1, f2):
    """
    Perform one Muller's method iteration to estimate the root of f(T) = 0.

    Parameters:
        T0, T1, T2 : Three successive temperature guesses (°C)
        f0, f1, f2 : Function values at T0, T1, T2

    Returns:
        T3 : Updated temperature guess (°C)
    """
    h1 = T1 - T0
    h2 = T2 - T1

    d1 = (f1 - f0) / h1
    d2 = (f2 - f1) / h2
    d = (d2 - d1) / (h2 + h1)

    radicand = d2**2 - 4 * d * f2
    if radicand < 0:
        raise ValueError("Muller's method: negative discriminant — complex root encountered.")

    sqrt_val = math.sqrt(radicand)
    denominator = (d2 + sqrt_val) if abs(d2 + sqrt_val) > abs(d2 - sqrt_val) else (d2 - sqrt_val)

    if denominator == 0:
        raise ValueError("Muller's method: zero denominator encountered.")

    return T2 + (-2 * f2 / denominator)


def calculate_steady_temperature(
        elev, t_a, dia, lat, dmy, h24, z_l,
        t_rh, t_rl, r_h, r_l, amps,
        phi_d=90, emsvty=0.7, alpha=0.8, wnd_spd=0.61, atmosphere='clear'):
    """
    Calculate steady-state conductor temperature for a given current (IEEE 738).

    Uses Muller's iterative method to solve the heat balance equation.

    Parameters:
        elev       : Elevation above sea level (m)
        t_a        : Ambient air temperature (°C)
        dia        : Outside diameter of conductor (m)
        lat        : Latitude (degrees)
        dmy        : Date as "dd/mm/yyyy" string or integer day-of-year
        h24        : Hour of day in 24-hour format
        z_l        : Line azimuth (degrees)
        t_rh       : High reference temperature for resistance (°C)
        t_rl       : Low reference temperature for resistance (°C)
        r_h        : AC resistance at high reference temperature (Ω/m)
        r_l        : AC resistance at low reference temperature (Ω/m)
        amps       : Steady-state current (A)
        phi_d      : Angle between wind and conductor axis (degrees; default 90)
        emsvty     : Emissivity (default 0.7)
        alpha      : Solar absorptivity (default 0.8)
        wnd_spd    : Wind speed (m/s; default 0.61)
        atmosphere : 'clear' or 'industrial' (default 'clear')

    Returns:
        Steady-state conductor temperature (°C), rounded to 1 decimal place
    """
    q_s = calculate_solar_heat_gain(alpha, lat, dmy, h24, elev, z_l, dia, atmosphere)

    guess = [t_a, t_a + 5, t_a + 10]
    tol = 0.0001

    for iteration in range(1, 101):
        errors = []
        for t in guess:
            r_d = calculate_resistance_at_temperature(t_rh, t_rl, r_h, r_l, t)
            q_cn = calculate_natural_convection(elev, t, t_a, dia)
            q_f = calculate_forced_convection(elev, t, t_a, dia, wnd_spd, phi_d)
            q_c = max(q_cn, q_f)
            q_r = calculate_radiation(t, t_a, dia, emsvty)
            errors.append((amps**2) * r_d + q_s - q_c - q_r)

        if abs(errors[-1]) < tol:
            break

        guess.append(_muller_method(guess[0], guess[1], guess[2], errors[0], errors[1], errors[2]))
        guess.pop(0)
    else:
        raise RuntimeError("calculate_steady_temperature: failed to converge after 100 iterations.")

    return round(guess[-1], 1)


def calculate_transient_temperature(
        elev, t_a, dia, lat, dmy, h24, z_l,
        t_rh, t_rl, r_h, r_l,
        ini_amps, trans_amps, mcp,
        phi_d=90, emsvty=0.7, alpha=0.8, wnd_spd=0.61, atmosphere='clear',
        tran_dur=60, del_t=10):
    """
    Calculate transient conductor temperature during/after a current step change (IEEE 738).

    The initial conductor temperature is computed from the pre-transient steady-state current.
    Temperature is then stepped forward in time using Euler integration.

    Parameters:
        elev       : Elevation above sea level (m)
        t_a        : Ambient air temperature (°C)
        dia        : Outside diameter of conductor (m)
        lat        : Latitude (degrees)
        dmy        : Date as "dd/mm/yyyy" string or integer day-of-year
        h24        : Hour of day in 24-hour format
        z_l        : Line azimuth (degrees)
        t_rh       : High reference temperature for resistance (°C)
        t_rl       : Low reference temperature for resistance (°C)
        r_h        : AC resistance at high reference temperature (Ω/m)
        r_l        : AC resistance at low reference temperature (Ω/m)
        ini_amps   : Pre-transient steady-state current (A)
        trans_amps : Transient (fault/overload) current (A)
        mcp        : Conductor heat capacity (J/m·°C)
        phi_d      : Angle between wind and conductor axis (degrees; default 90)
        emsvty     : Emissivity (default 0.7)
        alpha      : Solar absorptivity (default 0.8)
        wnd_spd    : Wind speed (m/s; default 0.61)
        atmosphere : 'clear' or 'industrial' (default 'clear')
        tran_dur   : Calculation duration (minutes; default 60)
        del_t      : Time step (seconds; default 10)

    Returns:
        time_series : List of time values in seconds [0, del_t, 2*del_t, ..., tran_dur*60]
        final_temp  : List of conductor temperatures (°C) at each time step
    """
    # Compute initial temperature from pre-transient current
    ini_temp = calculate_steady_temperature(
        elev, t_a, dia, lat, dmy, h24, z_l,
        t_rh, t_rl, r_h, r_l, ini_amps,
        phi_d=phi_d, emsvty=emsvty, alpha=alpha,
        wnd_spd=wnd_spd, atmosphere=atmosphere,
    )

    time_series = list(range(0, 60 * (tran_dur + 1), del_t))
    temp = ini_temp
    final_temp = [ini_temp]

    for i in range(1, len(time_series)):
        t_prev = final_temp[i - 1]

        q_cn = calculate_natural_convection(elev, t_prev, t_a, dia)
        q_f = calculate_forced_convection(elev, t_prev, t_a, dia, wnd_spd, phi_d)
        q_c = max(q_cn, q_f)
        q_r = calculate_radiation(t_prev, t_a, dia, emsvty)
        q_s = calculate_solar_heat_gain(alpha, lat, dmy, h24, elev, z_l, dia, atmosphere)
        r_d = calculate_resistance_at_temperature(t_rh, t_rl, r_h, r_l, t_prev)

        del_temp = (r_d * trans_amps**2 + q_s - q_c - q_r) * del_t / mcp
        temp = temp + del_temp
        final_temp.append(round(temp, 1))

    return time_series, final_temp
