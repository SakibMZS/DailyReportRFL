# =========================================================
# CONFIG.PY — CENTRAL INDUSTRIAL ENGINEERING CONFIGURATION
# =========================================================

# Machine tonnage size classes
EXCEL_SIZES = [
    "160", "90", "120", "250", "270", "280",
    "380", "330", "470", "530", "800", "428"
]

TOTAL_PLANT_MCS = 61
DAILY_AVAILABLE_HRS = TOTAL_PLANT_MCS * 24.0  # 1,464.0 Hours/Day

# Standard 61 Plastic-3 Injection Molding Machine Mappings
# Format: (Position / Local Name, Line / Shop Floor, ERP Machine SL)
MAPPING_DATA = [
    ("A1-160", "FF A-B", "IMM-160-6"),
    ("A2-120", "FF A-B", "IMM-120-20"),
    ("A3-120", "FF A-B", "IMM-120-28"),
    ("A4-120", "FF A-B", "IMM-120-29"),
    ("A5-160", "FF A-B", "IMM-160-7"),
    ("A6-160", "FF A-B", "IMM-160-12"),
    ("A7-160", "FF A-B", "IMM-160-48"),
    ("B1-120", "FF A-B", "IMM-120-11"),
    ("B2-120", "FF A-B", "IMM-120-15"),
    ("B3-120", "FF A-B", "IMM-120-14"),
    ("B4-120", "FF A-B", "IMM-120-75"),
    ("B5-90PC", "FF A-B", "IMM-90-8"),
    ("B6-90PC", "FF A-B", "IMM-90-9"),
    ("B7-120PC", "FF A-B", "IMM-120-32"),
    ("B8-120PC", "FF A-B", "IMM-120-27"),
    ("C1-120", "FF C-D", "IMM-120-4"),
    ("C2-160", "FF C-D", "IMM-160-17"),
    ("C3-120", "FF C-D", "IMM-120-22"),
    ("C4-120PC", "FF C-D", "IMM-120-46"),
    ("C5-90", "FF C-D", "IMM-90-4"),
    ("C6-120", "FF C-D", "IMM-120-47"),
    ("C7-160", "FF C-D", "IMM-160-51"),
    ("D1-160", "FF C-D", "IMM-160-39"),
    ("D2-160", "FF C-D", "IMM-160-79"),
    ("D3-160", "FF C-D", "IMM-160-80"),
    ("A1-280TC", "GF A-B", "IMM-280R-25"),
    ("A2-380", "GF A-B", "IMM-380-5"),
    ("A3-380 (PC)", "GF A-B", "IMM-380-81"),
    ("A4-380", "GF A-B", "IMM-380-80"),
    ("A5-HP-330", "GF A-B", "IMM-330-4"),
    ("B1-470", "GF A-B", "IMM-470-5"),
    ("B2-380", "GF A-B", "IMM-380-6"),
    ("B3-530", "GF A-B", "IMM-530-15"),
    ("B4-530", "GF A-B", "IMM-530-16"),
    ("B5-530", "GF A-B", "IMM-530-22"),
    ("B6-380", "GF A-B", "IMM-380-4"),
    ("C1-800-30", "GF C-D", "IMM-800-30"),
    ("C2-800-31", "GF C-D", "IMM-800-31"),
    ("C3-270-1", "GF C-D", "IMM-270-1"),
    ("C4-380-73", "GF C-D", "IMM-380-73"),
    ("C5-380-44", "GF C-D", "IMM-380-44"),
    ("C6-280TC", "GF C-D", "IMM-280R-3"),
    ("D1-280TC", "GF C-D", "IMM-280R-24"),
    ("D2-MA2-250", "GF C-D", "IMM-250-106"),
    ("D3-330-1", "GF C-D", "IMM-330-1"),
    ("D4-HP-330-5", "GF C-D", "IMM-330-5"),
    ("D5-428-1", "GF C-D", "IMM-428-1"),
    ("D6-HP-428-4", "GF C-D", "IMM-428-4"),
    ("D7-HP-330", "GF C-D", "IMM-330-8"),
    ("E1-380-90", "GF E-F", "IMM-380-90"),
    ("E2-380-94", "GF E-F", "IMM-380-94"),
    ("E3-380-88", "GF E-F", "IMM-380-88"),
    ("E4-380-76", "GF E-F", "IMM-380-76"),
    ("E5-380-62", "GF E-F", "IMM-380-62"),
    ("E6-380-75", "GF E-F", "IMM-380-75"),
    ("F1-380-92", "GF E-F", "IMM-380-92"),
    ("F2-380-93", "GF E-F", "IMM-380-93"),
    ("F3-380-98", "GF E-F", "IMM-380-98"),
    ("F4-380-99", "GF E-F", "IMM-380-99"),
    ("F5-380-101", "GF E-F", "IMM-380-101"),
    ("F6-380-100", "GF E-F", "IMM-380-100"),
]

# Fast Lookup Dictionaries
POS_MAP = {smart_manu: pos for pos, line, smart_manu in MAPPING_DATA}
LINE_MAP = {smart_manu: line for pos, line, smart_manu in MAPPING_DATA}

# Maintenance & Breakdown Defect Whitelist
MAINTENANCE_CAUSES = {
    "Machine Problem*",
    "Controller Problem*",
    "Robot Problem*",
    "Scheduled Maintenance*",
    "Power Breakdown (Unscheduled)*",
}
