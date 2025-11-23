"""
Settings module for the PCC Classlist Party Pack.
Holds default settings and loads user overrides from settings.json.
"""

import json
import os
import sys


# Default settings — safe for all users
DEFAULT_SETTINGS = {
    "department_prefix": None,
    "allowed_courses": {
        "170", "221", "223", "240", "242", "244",
        "246", "248", "252", "254", "260", "265",
        "266", "267", "270", "280A"
    },
    "email_domain": "@pcc.edu",
    "output_name_prefix": "GEO_Class_Lists",
    "output_subfolder": "Output Files",
}


def app_dir() -> str:
    """Return the directory where settings.json should live."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_settings() -> dict:
    """
    Load settings from settings.json if present.
    Merge them on top of DEFAULT_SETTINGS.
    """
    settings = DEFAULT_SETTINGS.copy()

    config_path = os.path.join(app_dir(), "settings.json")
    if not os.path.exists(config_path):
        return settings

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            user_settings = json.load(f)
            if not isinstance(user_settings, dict):
                raise ValueError("settings.json did not contain an object.")
            settings.update(user_settings)
    except Exception as e:
        print(f"Could not load settings.json, using defaults. Error: {e}")

    return settings
