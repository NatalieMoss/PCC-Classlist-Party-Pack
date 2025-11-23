"""
PCC Classlist Party Starter
---------------------------
A tool for parsing Banner SIS PDF class lists into clean spreadsheets.
Part of the PCC Classlist Party Pack.
"""

import os
import re
import sys
from collections import defaultdict
from tkinter import Tk, filedialog, messagebox
import json
import pandas as pd
import pdfplumber
from settings import load_settings

SETTINGS = load_settings()


# CONFIGURATION SETTINGS FOR THIS TOOL.
# Other departments or colleges can edit these values to customize behavior.
# 🔧 TO CUSTOMIZE FOR YOUR DEPARTMENT:
# - Change department_prefix to your subject code (e.g., "MTH", "WR").
# - Change allowed_courses to your course numbers or set to None.
# - Change email_domain if you're not at PCC.

DEFAULT_SETTINGS = {
    # If you want to restrict parsing to a single subject/department
    # (e.g., "GEO"), put the subject code here. Otherwise, use None.
    "department_prefix": None,  # e.g., "GEO" or None for all subjects

    # List of allowed course numbers as strings (e.g., ["170", "221"]).
    # Set to None to accept all course numbers.
    "allowed_courses": {"170", "221", "223", "240", "242", "244", "246",
                        "248", "252", "254", "260", "265", "266", "267",
                        "270", "280A"},

    # Email domain used for institutional emails in the PDF.
    # Other colleges can change this to their domain.
    "email_domain": "@pcc.edu",

    # Prefix for the output Excel filename (term will be appended if found).
    "output_name_prefix": "GEO_Class_Lists",

    # Subfolder where output files are written.
    "output_subfolder": "Output Files",
}

def load_settings() -> dict:
    """
    Load user settings from settings.json if it exists, and merge them
    over the DEFAULT_SETTINGS. This lets non-programmers customize the
    tool without editing the Python code.

    Returns:
        dict: The effective settings dictionary.
    """
    settings = DEFAULT_SETTINGS.copy()
    config_path = os.path.join(app_dir(), "settings.json")

    try:
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                user_settings = json.load(f)

            # Merge user settings onto defaults
            if isinstance(user_settings, dict):
                settings.update(user_settings)
            else:
                messagebox.showerror(
                    "Settings error",
                    "settings.json is not a JSON object. Using default settings instead."
                )
    except Exception as e:
        # If anything goes wrong, fall back to defaults but let the user know
        messagebox.showerror(
            "Settings error",
            f"Could not load settings.json.\nUsing default settings.\n\nDetails: {e}"
        )

    return settings

_TERM_TEXT_RE = re.compile(r'\b(Spring|Summer|Fall|Winter)\s+(20\d{2})\b', re.I)
_TERM_CODE_RE = re.compile(r'\b(20\d{2}0[1-4])\b')  # e.g. 202501..202504

def app_dir() -> str:
    """
    Get the folder where this tool should write its output.
    If the script is 'frozen' into an EXE (using something like PyInstaller),
    this returns the folder where the EXE lives. Otherwise, it returns the
    folder containing this .py file.
    Returns:
        str: Absolute path to the application's base directory.
    """
    if getattr(sys, "frozen", False):
        # Running as a bundled EXE
        return os.path.dirname(sys.executable)
    # Running as a normal .py file
    return os.path.dirname(os.path.abspath(__file__))

def _term_from_code(code: str) -> str:
    """
    Convert a Banner-style term code (e.g. '202503') into a human-readable label.
    Args:
        code (str): The 6-digit Banner term code. The last digit encodes the season:
                    1 = Winter, 2 = Spring, 3 = Summer, 4 = Fall.
    Returns:
        str: A string like 'Spring 2025' or an empty string if the code is not valid.
    """
    season_map = {1: "Winter", 2: "Spring", 3: "Summer", 4: "Fall"}
    # Example: '202503' -> last digit is 3 -> 'Summer', year is '2025'
    return f"{season_map.get(int(code[-1]), '')} {code[:4]}".strip()

def load_settings() -> dict:
    """
    Load user settings from settings.json if it exists, and merge them
    over the default SETTINGS from settings.py.

    This lets non-programmers tweak behavior without editing Python code.

    Rules:
      - If settings.json doesn't exist, we just use SETTINGS.
      - If it exists but is invalid, we show a warning and fall back to SETTINGS.
      - If it supplies allowed_courses as a list, we convert it to a set.
    """
    # Start with the defaults from settings.py
    settings = dict(SETTINGS)

    config_path = os.path.join(app_dir(), "settings.json")

    try:
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                user_settings = json.load(f)

            if isinstance(user_settings, dict):
                # If allowed_courses is supplied as a list in JSON,
                # convert it to a set to match our internal usage.
                if "allowed_courses" in user_settings:
                    ac = user_settings["allowed_courses"]
                    if ac is None:
                        # Explicitly disable course filtering
                        user_settings["allowed_courses"] = None
                    elif isinstance(ac, list):
                        user_settings["allowed_courses"] = set(ac)

                settings.update(user_settings)
            else:
                messagebox.showerror(
                    "Settings error",
                    "settings.json must contain a JSON object at the top level.\n"
                    "Using default settings instead."
                )
    except Exception as e:
        # If anything goes wrong, fall back to defaults but let the user know.
        messagebox.showerror(
            "Settings error",
            f"Could not load settings.json.\n"
            f"Using default settings instead.\n\nDetails: {e}"
        )

    return settings


def _detect_term_from_pdf(pdf) -> str:
    """
    Detect the academic term from the first page of a Banner class list PDF.
    This function checks for:
      1. Explicit term text like 'Fall 2025', or
      2. A numeric Banner term code like '202503', which it then converts.
    Args:
        pdf: An open pdfplumber PDF object.
    Returns:
        str: A human-readable term such as 'Fall 2025', or an empty string
             if no term information can be found.
    """
    try:
        header = (pdf.pages[0].extract_text() or "")
    except Exception:
        header = ""

    # Try a direct match like 'Fall 2025'
    m = _TERM_TEXT_RE.search(header)
    if m:
        return f"{m.group(1).title()} {m.group(2)}"

    # Otherwise, look for Banner codes like '202503'
    for m in _TERM_CODE_RE.finditer(header):
        guess = _term_from_code(m.group(0))
        if guess:
            return guess

    return ""  # okay if empty; we'll fall back to a generic filename


def _safe_filename(s: str) -> str:
    """
    Clean a string so it is safe to use as a filename on most operating systems.
    This:
      * Replaces None with an empty string.
      * Removes characters that are not letters, numbers, spaces, underscores, or hyphens.
      * Collapses repeated whitespace into a single space.
    Args:
        s (str): The original string (e.g. a course title or term label).
    Returns:
        str: A simplified, filesystem-safe version of the string.
    """
    s = (s or "").strip()
    # Keep only letters, digits, spaces, underscores, and hyphens
    s = re.sub(r'[^A-Za-z0-9 _\-]', '', s)
    # Collapse multiple spaces into one
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def main() -> None:
    """
    Main entry point for the Classlist Party Starter.

    Opens a file dialog to select a Banner class list PDF (landscape format),
    parses the contents into student records, and writes an Excel workbook
    with a Combined sheet and one sheet per class/CRN.
    """
    settings = load_settings()
    # Hide the root Tk window and show an "Open file" dialog instead
    root = Tk()
    root.withdraw()

    pdf_file = filedialog.askopenfilename(
        title="Select your class list PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_file:
        # User cancelled; just exit quietly
        return

    records = []
    current_class = {}

    with pdfplumber.open(pdf_file) as pdf:
        # Detect the term while the file is open
        term_text = _detect_term_from_pdf(pdf)   # e.g. "Fall 2025" or ""

        for page in pdf.pages:
            text = (page.extract_text() or "")
            lines = text.split("\n")

            # --- Class header parsing: capture CRN, subject, course number, section, name ---
            for line in lines:
                course_match = re.match(r"\s*(\d{5})\s+(\w+)\s+(\d+[A-Z]?)\s+(\d)\s+(.*)", line)
                if course_match:
                    course_number = course_match.group(3)

                    # Optional course filter from settings
                    allowed = settings.get("allowed_courses")
                    if allowed is not None and course_number not in allowed:
                        # Skip this class if it's not in the allowed list
                        current_class = {}
                    else:
                        current_class = {
                            "CRN": course_match.group(1),
                            "Subject": course_match.group(2),
                            "Course Number": course_number,
                            "Section": course_match.group(4),
                            "Course Name": course_match.group(5).strip(),
                        }

            # --- Student rows ---
            idx = 0
            while idx < len(lines):
                line = lines[idx]
                gnum_match = re.search(r"(G\d{8})", line)
                if gnum_match and current_class:
                    try:
                        # Extract "Last, First" before the G-number
                        name_part = line.split(gnum_match.group(1))[0]
                        last_first = name_part.split(None, 1)[1].split(',')
                        last_name = last_first[0].strip()
                        first_name = last_first[1].strip()
                    except Exception:
                        last_name, first_name = "", ""

                    g_number = gnum_match.group(1)

                    # Look for institutional email on the next line
                    email = ""
                    if idx + 1 < len(lines):
                        email_line = lines[idx + 1].strip()
                        if settings["email_domain"] in email_line:
                            email = email_line.split()[0]

                    records.append({
                        "First Name": first_name,
                        "Last Name": last_name,
                        "G Number": g_number,
                        "PCC email address": email,
                        "Non-PCC email": "",
                        "Class": f"{current_class.get('Subject')} {current_class.get('Course Number')}",
                        "CRN": current_class.get("CRN"),
                    })
                    # Skip to the line after the email line
                    idx += 2
                else:
                    idx += 1

    # --- Build dynamic output name (now that pdf is closed) ---
    prefix = settings["output_name_prefix"]
    stem = f"{prefix}_{term_text}" if term_text else prefix
    out_name = _safe_filename(stem) + ".xlsx"

    # Always write to: <tool folder>\<output_subfolder>
    output_dir = os.path.join(app_dir(), settings["output_subfolder"])
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(output_dir, out_name)

    # Group records by Class+CRN for per-class sheets
    grouped = defaultdict(list)
    for rec in records:
        grouped[f"{rec['Class']}_{rec['CRN']}"].append(rec)

    # Write the Excel file
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            pd.DataFrame(records).to_excel(writer, sheet_name="Combined", index=False)

            for key, recs in grouped.items():
                sheet_name = key[:31]  # Excel sheet name limit
                pd.DataFrame(recs).to_excel(writer, sheet_name=sheet_name, index=False)

        messagebox.showinfo("Done", f"Created:\n{output_path}")

    except PermissionError:
        messagebox.showerror(
            "Permission Error",
            "I couldn't write the Excel file because it appears to be open in another program.\n\n"
            "Please close the file in Excel (or any app that may be locking it) and run the tool again."
        )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        messagebox.showerror("Error", f"There was an unexpected error:\n{e}")
