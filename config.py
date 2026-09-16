# =========================================================
# CONFIG.PY — MULTI-FLOOR INDUSTRIAL ENGINEERING REGISTRY
# =========================================================

MAINTENANCE_CAUSES = {
    "Machine Problem*",
    "Controller Problem*",
    "Robot Problem*",
    "Scheduled Maintenance*",
    "Power Breakdown (Unscheduled)*",
}

# Master Section Registry parsed from machine_id.xlsx
# Key: Exact Section string as formatted in ERP Reports
SECTION_CONFIG = {
    "RIP> DPL> Plastic-3": {
        "total_mcs": 61,
        "mapping": {
            "IMM-160-6": "A1-160", "IMM-120-20": "A2-120", "IMM-120-28": "A3-120", "IMM-120-29": "A4-120",
            "IMM-160-7": "A5-160", "IMM-160-12": "A6-160", "IMM-160-48": "A7-160", "IMM-120-11": "B1-120",
            "IMM-120-15": "B2-120", "IMM-120-14": "B3-120", "IMM-120-75": "B4-120", "IMM-90-8": "B5-90PC",
            "IMM-90-9": "B6-90PC", "IMM-120-32": "B7-120PC", "IMM-120-27": "B8-120PC", "IMM-120-4": "C1-120",
            "IMM-160-17": "C2-160", "IMM-120-22": "C3-120", "IMM-120-46": "C4-120PC", "IMM-90-4": "C5-90",
            "IMM-120-47": "C6-120", "IMM-160-51": "C7-160", "IMM-160-39": "D1-160", "IMM-160-79": "D2-160",
            "IMM-160-80": "D3-160", "IMM-280R-25": "A1-280TC", "IMM-380-5": "A2-380", "IMM-380-81": "A3-380 (PC)",
            "IMM-380-80": "A4-380", "IMM-330-4": "A5-HP-330", "IMM-470-5": "B1-470", "IMM-380-6": "B2-380",
            "IMM-530-15": "B3-530", "IMM-530-16": "B4-530", "IMM-530-22": "B5-530", "IMM-380-4": "B6-380",
            "IMM-800-30": "C1-800-30", "IMM-800-31": "C2-800-31", "IMM-270-1": "C3-270-1", "IMM-380-73": "C4-380-73",
            "IMM-380-44": "C5-380-44", "IMM-280R-3": "C6-280TC", "IMM-280R-24": "D1-280TC", "IMM-250-106": "D2-MA2-250",
            "IMM-330-1": "D3-330-1", "IMM-330-5": "D4-HP-330-5", "IMM-428-1": "D5-428-1", "IMM-428-4": "D6-HP-428-4",
            "IMM-330-8": "D7-HP-330", "IMM-380-90": "E1-380-90", "IMM-380-94": "E2-380-94", "IMM-380-88": "E3-380-88",
            "IMM-380-76": "E4-380-76", "IMM-380-62": "E5-380-62", "IMM-380-75": "E6-380-75", "IMM-380-92": "F1-380-92",
            "IMM-380-93": "F2-380-93", "IMM-380-98": "F3-380-98", "IMM-380-99": "F4-380-99", "IMM-380-101": "F5-380-101",
            "IMM-380-100": "F6-380-100",
        },
    },
    "RIP> DPL> Plastic-6.1": {
        "total_mcs": 76,
        "mapping": {
            "IMM-380-86": "A1", "IMM-470-4": "A2", "IMM-380-43": "A3", "IMM-380-13": "A4", "IMM-380-29": "A5",
            "IMM-380-42": "A6", "IMM-380-97": "A7", "IMM-380-3": "B1", "IMM-380-28": "B2", "IMM-380-30": "B3",
            "IMM-380-31": "B4", "IMM-380-77": "B5", "IMM-380-78": "B6", "IMM-380-1": "C1", "IMM-380-2": "C2",
            "IMM-380-87": "C3", "IMM-380-41": "C4", "IMM-380-95": "C5", "IMM-380-79": "C6", "IMM-380-40": "D1",
            "IMM-380-39": "D2", "IMM-380-12": "D3", "IMM-380-38": "D4", "IMM-380-37": "D5", "IMM-380-96": "D6",
            "IMM-380-10": "E1", "IMM-380-8": "E2", "IMM-380-7": "E3", "IMM-380-9": "E4", "IMM-380-11": "E5",
            "IMM-470-3": "F1", "IMM-470-1": "F2", "IMM-470-2": "F3", "IMM-470-6": "F4", "IMM-530-18": "G1",
            "IMM-530-19": "G2", "IMM-530-20": "G3", "IMM-530-21": "G4", "IMM-500-1": "H1", "IMM-160-56": "H2",
            "IMM-160-57": "H3", "IMM-160-58": "H4", "IMM-160-59": "H5", "IMM-160-61": "H6", "IMM-160-62": "H7",
            "IMM-160-63": "I1", "IMM-160-64": "I2", "IMM-160-65": "I3", "IMM-160-66": "I4", "IMM-160-93": "I5",
            "IMM-160-94": "I6", "IMM-160-95": "I7", "IMM-160-96": "J1", "IMM-120-72": "J2", "IMM-90-5": "J3",
            "IMM-90-6": "J4", "IMM-260-1": "J5", "IMM-250-25": "K1", "IMM-250-26": "K2", "IMM-250-28": "K3",
            "IMM-250-29": "K4", "IMM-250-30": "K5", "IMM-250-31": "K6", "IMM-250-32": "K7", "IMM-250-33": "K8",
            "IMM-250-17": "L1", "IMM-250-18": "L2", "IMM-250-19": "L3", "IMM-250-20": "L4", "IMM-250-21": "L5",
            "IMM-250-22": "L6", "IMM-250-181": "L7", "IMM-250-182": "L8", "IMM-250-24": "M1", "IMM-250-183": "M2",
            "IMM-250-185": "M3",
        },
    },
    "RIP> DPL> Plastic-6.2": {
        "total_mcs": 65,
        "mapping": {
            "IMM-250-23": "A1", "IMM-250-100": "A2", "IMM-250-184": "A3", "IMM-160-101": "A4", "IMM-160-19": "A5",
            "IMM-250-35": "A6", "IMM-250-27": "A7", "IMM-250-193": "A8", "IMM-160-97": "B1", "IMM-160-98": "B2",
            "IMM-160-99": "B3", "IMM-160-100": "B4", "IMM-160-18": "B5", "IMM-250-41": "B6", "IMM-250-40": "B7",
            "IMM-250-39": "B8", "IMM-250-34": "C1", "IMM-250-37": "C2", "IMM-250-44": "C3", "IMM-250-45": "C4",
            "IMM-250-43": "C5", "IMM-250-42": "C6", "IMM-250-38": "C7", "IMM-250-36": "C8", "IMM-250-53": "D1",
            "IMM-250-52": "D2", "IMM-250-51": "D3", "IMM-250-50": "D4", "IMM-250-49": "D5", "IMM-250-48": "D6",
            "IMM-250-47": "D7", "IMM-250-46": "D8", "IMM-250-54": "E1", "IMM-250-55": "E2", "IMM-250-56": "E3",
            "IMM-250-57": "E4", "IMM-250-58": "E5", "IMM-250-59": "E6", "IMM-250-60": "E7", "IMM-250-61": "E8",
            "IMM-250-68": "F1", "IMM-250-67": "F2", "IMM-250-66": "F3", "IMM-250-65": "F4", "IMM-250-64": "F5",
            "IMM-250-63": "F6", "IMM-250-62": "F7", "IMM-120-62": "G1", "IMM-120-63": "G2", "IMM-120-64": "G3",
            "IMM-120-65": "G4", "IMM-120-66": "G5", "IMM-120-67": "G6", "IMM-120-68": "G7", "IMM-120-69": "G8",
            "IMM-120-70": "H1", "IMM-120-71": "H2", "IMM-120-73": "H3", "IMM-120-74": "H4", "IMM-120-77": "H5",
            "IMM-160-102": "H6", "IMM-160-20": "H7", "IMM-160-21": "H8", "IMM-160-22": "H9", "IMM-90-1": "I1",
            "IMM-90-2": "I2", "IMM-90-3": "I3", "IMM-90-7": "I4",
        },
    },
    "RIP> DPL> Plastic-7.1": {
        "total_mcs": 36,
        "mapping": {
            "IMM-530-17": "A1-530", "IMM-530-12": "A2-530", "IMM-330-6": "B1-330", "IMM-265-1": "B2-265",
            "IMM-268-1": "B3-268", "IMM-380-63": "B4-380", "IMM-270-4": "B5-270", "IMM-530-14": "B6-530",
            "IMM-270-3": "C1-270", "IMM-330-2": "C2-330", "IMM-330-3": "C3-330", "IMM-380-74": "C4-380",
            "IMM-380-7": "C5-380", "IMM-470-8": "C6-470", "IMM-470-7": "D1-470", "IMM-370-1": "D2-370",
            "IMM-380-8": "D3-380", "IMM-380-9": "D4-380", "IMM-470-9": "D5-470", "IMM-470-10": "D6-470",
            "IMM-470-11": "E1-470", "IMM-470-12": "E2-470", "IMM-470-13": "E3-470", "IMM-470-14": "E4-470",
            "IMM-470-15": "E5-470", "IMM-470-16": "E6-470", "IMM-530-13": "F1-530", "IMM-330-7": "F2-330",
            "IMM-428-2": "F3-428", "IMM-428-3": "F4-428", "IMM-428-5": "F5-428", "IMM-428-6": "F6-428",
            "IMM-428-7": "F7-428", "IMM-428-8": "F8-428", "IMM-270-2": "G1-270", "IMM-380-91": "G2-380",
        },
    },
    "RIP> DPL> Plastic-7.2": {
        "total_mcs": 76,
        "mapping": {
            "BMM-05L-11": "BMM-A1", "BMM-05L-04": "BMM-A2", "BMM-05L-03": "BMM-A3", "BMM-05L-06": "BMM-A4",
            "BMM-05L-05": "BMM-A5", "BMM-05L-16": "BMM-A6", "IMM-250-69": "A1", "IMM-250-176": "A2",
            "IMM-250-70": "A3", "IMM-250-177": "A4", "IMM-250-71": "A5", "IMM-250-178": "A6", "IMM-250-72": "A7",
            "IMM-250-179": "A8", "IMM-250-73": "A9", "IMM-250-180": "A10", "IMM-250-74": "B1", "IMM-250-75": "B2",
            "IMM-250-76": "B3", "IMM-250-77": "B4", "IMM-250-78": "B5", "IMM-250-79": "B6", "IMM-250-80": "B7",
            "IMM-250-81": "B8", "IMM-250-82": "B9", "IMM-250-83": "B10", "IMM-250-84": "C1", "IMM-250-85": "C2",
            "IMM-250-86": "C3", "IMM-250-87": "C4", "IMM-250-88": "C5", "IMM-250-89": "C6", "IMM-250-90": "C7",
            "IMM-250-91": "C8", "IMM-250-92": "C9", "IMM-250-93": "C10", "IMM-250-94": "D1", "IMM-250-95": "D2",
            "IMM-250-96": "D3", "IMM-250-97": "D4", "IMM-250-98": "D5", "IMM-250-99": "D6", "IMM-250-101": "D7",
            "IMM-250-102": "D8", "IMM-250-103": "D9", "IMM-250-104": "D10", "IMM-250-105": "D11", "IMM-160-68": "E1",
            "IMM-160-69": "E2", "IMM-160-70": "E3", "IMM-160-71": "E4", "IMM-160-72": "E5", "IMM-160-73": "E6",
            "IMM-160-74": "E7", "IMM-160-75": "E8", "IMM-160-76": "E9", "IMM-160-77": "E10", "IMM-160-78": "E11",
            "IMM-160-81": "F1", "IMM-160-82": "F2", "IMM-160-83": "F3", "IMM-160-84": "F4", "IMM-160-85": "F5",
            "IMM-160-86": "F6", "IMM-120-33": "G1", "IMM-120-34": "G2", "IMM-120-35": "G3", "IMM-120-36": "G4",
            "IMM-90-10": "G5", "IMM-90-11": "G6", "IMM-90-12": "G7", "IMM-260-55": "G8",
        },
    },
    "RIP> DPL> Plastic-7.3": {
        "total_mcs": 77,
        "mapping": {
            "IMM-260-49": "A1-260", "IMM-260-50": "A2-260", "IMM-160-31": "A3-160", "IMM-160-67": "A4-160",
            "IMM-260-51": "A5-260", "IMM-260-52": "A6-260", "IMM-260-53": "A7-260", "IMM-160-60": "A8-160",
            "IMM-160-32": "A9-160", "IMM-160-33": "A10-160", "IMM-260-31": "B1-260", "IMM-260-32": "B2-260",
            "IMM-260-33": "B3-260", "IMM-260-34": "B4-260", "IMM-260-35": "B5-260", "IMM-260-36": "B6-260",
            "IMM-260-37": "B7-260", "IMM-260-38": "B8-260", "IMM-260-39": "B9-260", "IMM-260-40": "B10-260",
            "IMM-260-41": "C1-260", "IMM-260-42": "C2-260", "IMM-260-43": "C3-260", "IMM-260-44": "C4-260",
            "IMM-260-45": "C5-260", "IMM-260-46": "C6-260", "IMM-260-47": "C7-260", "IMM-260-48": "C8-260",
            "IMM-260-54": "C9-260", "IMM-160-34": "C10-160", "IMM-250-107": "D1-250", "IMM-250-108": "D2-250",
            "IMM-250-109": "D3-250", "IMM-250-110": "D4-250", "IMM-250-111": "D5-250", "IMM-250-112": "D6-250",
            "IMM-250-113": "D7-250", "IMM-250-114": "D8-250", "IMM-250-115": "D9-250", "IMM-250-116": "D10-250",
            "IMM-250-117": "D11-250", "IMM-250-118": "E1-250", "IMM-250-119": "E2-250", "IMM-250-120": "E3-250",
            "IMM-250-121": "E4-250", "IMM-250-122": "E5-250", "IMM-250-123": "E6-250", "IMM-250-124": "E7-250",
            "IMM-250-125": "E8-250", "IMM-250-126": "E9-250", "IMM-250-127": "E10-250", "IMM-250-128": "E11-250",
            "IMM-250-129": "F1-250", "IMM-250-130": "F2-250", "IMM-250-131": "F3-250", "IMM-250-132": "F4-250",
            "IMM-250-133": "F5-250", "IMM-250-134": "F6-250", "IMM-160-35": "F7-160", "IMM-160-36": "F8-160",
            "IMM-160-37": "F9-160", "IMM-160-38": "F10-160", "IMM-160-49": "G1-160", "IMM-160-50": "G2-160",
            "IMM-160-52": "G3-160", "IMM-160-53": "G4-160", "IMM-160-54": "G5-160", "IMM-160-55": "G6-160",
            "IMM-120-21": "G7-120", "IMM-120-30": "G8-120", "IMM-260-26": "G9-260", "IMM-260-27": "G10-260",
            "IMM-260-28": "G11-260", "IMM-260-29": "G12-260", "IMM-260-30": "G13-260",
        },
    },
}

DEFAULT_SECTION = "RIP> DPL> Plastic-3"


def parse_machine_size(smart_manu_or_pos):
    """
    Extracts tonnage/capacity class dynamically from Smart Manu string or Position.
    E.g. IMM-380-86 -> '380', IMM-280R-25 -> '280', BMM-05L-11 -> '05L'
    """
    text = str(smart_manu_or_pos).strip().upper()
    parts = text.split("-")
    if len(parts) >= 2:
        raw_size = parts[1]
        clean_size = raw_size.replace("R", "").replace("TC", "").strip()
        return clean_size if clean_size else raw_size
    return "Other"


def resolve_section_config(df=None, fallback_name=DEFAULT_SECTION):
    """
    Auto-detects factory section from the DataFrame's 'Section' column.
    Falls back gracefully if section not directly registered.
    Returns: (section_name, total_mcs, daily_avail_hrs, pos_map, size_counts, unique_sizes)
    """
    section_name = fallback_name
    if df is not None and "Section" in df.columns:
        valid_sections = [s for s in df["Section"].dropna().unique() if str(s).strip()]
        for vs in valid_sections:
            for registered in SECTION_CONFIG.keys():
                if registered.lower() in str(vs).lower() or str(vs).lower() in registered.lower():
                    section_name = registered
                    break

    cfg = SECTION_CONFIG.get(section_name, SECTION_CONFIG[DEFAULT_SECTION])
    total_mcs = cfg["total_mcs"]
    daily_avail_hrs = total_mcs * 24.0
    pos_map = cfg["mapping"]

    # Compute machine size distribution dynamically for the resolved section
    sizes_series = [parse_machine_size(sm) for sm in pos_map.keys()]
    size_counts = {}
    for s in sizes_series:
        size_counts[s] = size_counts.get(s, 0) + 1

    unique_sizes = sorted(size_counts.keys(), key=lambda x: (len(x), x), reverse=True)
    return section_name, total_mcs, daily_avail_hrs, pos_map, size_counts, unique_sizes


# =========================================================
# BACKWARD COMPATIBILITY EXPORTS (PREVENTS SUBMODULE BREAKS)
# =========================================================
TOTAL_PLANT_MCS = SECTION_CONFIG[DEFAULT_SECTION]["total_mcs"]
DAILY_AVAILABLE_HRS = TOTAL_PLANT_MCS * 24.0
POS_MAP = SECTION_CONFIG[DEFAULT_SECTION]["mapping"]
LINE_MAP = {k: "-" for k in POS_MAP.keys()}
EXCEL_SIZES = ["160", "90", "120", "250", "270", "280", "380", "330", "470", "530", "800", "428"]
