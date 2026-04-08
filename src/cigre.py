"""
CIGRE Thermal Rating Calculations — Placeholder for Future Implementation.

When CIGRE methods are added, implement functions here with the same signatures
as in ieee738.py. The main script will then seamlessly switch between standards
based on the 'standard' setting in config.ini.

Functions to implement:
    calculate_steady_thermal_rating(...)  — same signature as ieee738 version
    calculate_steady_temperature(...)     — same signature as ieee738 version
    calculate_transient_temperature(...)  — same signature as ieee738 version

Reference: CIGRE TB 207 / CIGRE TB 601

Author       : Imrul Qayes
Email        : imrul27@gmail.com
Date Created : 03/2025
Last Modified: 04/2026
Version      : 1.0.0
AI Assistant : Developed with the assistance of Claude (Anthropic) — https://claude.ai
"""

# Licensed under the MIT License. See LICENSE file in the project root for details.


def calculate_steady_thermal_rating(*args, **kwargs):
    raise NotImplementedError(
        "CIGRE steady-state thermal rating is not yet implemented. "
        "Set 'standard = ieee738' in config.ini to use the IEEE 738 method."
    )


def calculate_steady_temperature(*args, **kwargs):
    raise NotImplementedError(
        "CIGRE steady-state temperature is not yet implemented. "
        "Set 'standard = ieee738' in config.ini to use the IEEE 738 method."
    )


def calculate_transient_temperature(*args, **kwargs):
    raise NotImplementedError(
        "CIGRE transient temperature is not yet implemented. "
        "Set 'standard = ieee738' in config.ini to use the IEEE 738 method."
    )
