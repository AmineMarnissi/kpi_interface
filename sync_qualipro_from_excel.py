import json
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path


ROOT = Path("/home/amine/site_amine/zouhour_interface/new_version")
XLSX_FILE = ROOT / "QHSE KPI 2026.xlsx"
QHSE_CLEAN_FILE = ROOT / "qhse_data_clean.json"
KPI_METRICS_FILE = ROOT / "data" / "kpi_metrics.json"
QUALIPRO_DETAILS_FILE = ROOT / "data" / "qualipro_usage_details.json"

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

PERIOD_BLOCKS = [
    ("2026-06", "P", "Q", "R", "Juin"),
    ("2026-05", "U", "V", "W", "Mai"),
    ("2026-04", "Z", "AA", "AB", "Avril"),
    ("2026-03", "AE", "AF", "AG", "Mars"),
    ("2026-02", "AJ", "AK", "AL", "Fevrier"),
    ("2026-01", "AO", "AP", "AQ", "Janvier"),
]


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values = []
    for si in root.findall(f"{NS}si"):
        text = "".join(node.text or "" for node in si.iter() if node.tag == f"{NS}t")
        values.append(text)
    return values


def resolve_sheet_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{PKG_REL_NS}Relationship")}
    for sheet in workbook.find(f"{NS}sheets"):
        if sheet.attrib.get("name") == sheet_name:
            rid = sheet.attrib[f"{REL_NS}id"]
            return f"xl/{rel_map[rid]}"
    raise KeyError(f"Sheet {sheet_name!r} not found")


def cell_value(cell: ET.Element, shared_strings: list[str]):
    value = cell.find(f"{NS}v")
    if value is None:
        return None
    raw = value.text
    if cell.attrib.get("t") == "s":
        return shared_strings[int(raw)]
    return raw


def parse_sheet(archive: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> dict[str, object]:
    sheet = ET.fromstring(archive.read(sheet_path))
    cells = {}
    for row in sheet.find(f"{NS}sheetData").findall(f"{NS}row"):
        for cell in row.findall(f"{NS}c"):
            cells[cell.attrib["r"]] = cell_value(cell, shared_strings)
    return cells


def to_float(value):
    if value in (None, "", "-", "–"):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "").replace(",", ".").replace("%", "")
    try:
        return float(text)
    except ValueError:
        return None


def to_percent(value):
    number = to_float(value)
    if number is None:
        return None
    return round(number * 100, 2) if abs(number) <= 1.5 else round(number, 2)


def parse_qualipro_details() -> dict:
    with zipfile.ZipFile(XLSX_FILE) as archive:
        shared_strings = load_shared_strings(archive)
        sheet_path = resolve_sheet_path(archive, "KPI Qualipro")
        cells = parse_sheet(archive, sheet_path, shared_strings)

    users = {}
    totals = {"counts": {}, "rates": {}}
    periods = []

    for period_key, name_col, count_col, rate_col, label in PERIOD_BLOCKS:
        periods.append({"key": period_key, "label": label})
        for row in range(14, 80):
            name = cells.get(f"{name_col}{row}")
            if not name:
                continue
            rate = to_percent(cells.get(f"{rate_col}{row}"))
            count = to_float(cells.get(f"{count_col}{row}"))
            if name == "Total général":
                totals["counts"][period_key] = int(count) if count is not None else None
                totals["rates"][period_key] = rate
                continue

            entry = users.setdefault(name, {"user": name, "counts": {}, "rates": {}})
            if count is not None:
                entry["counts"][period_key] = int(count)
            entry["rates"][period_key] = rate

    user_list = []
    for name, entry in sorted(users.items(), key=lambda item: item[0].lower()):
        rates = [value for value in entry["rates"].values() if value is not None]
        entry["average_rate"] = round(sum(rates) / len(rates), 2) if rates else None
        user_list.append(entry)

    return {
        "meta": {
            "source": XLSX_FILE.name,
            "sheet": "KPI Qualipro",
            "generated": date.today().isoformat(),
            "site": "MEL-T",
        },
        "periods": periods,
        "users": user_list,
        "totals": totals,
    }


def update_qhse_clean(details: dict):
    data = json.loads(QHSE_CLEAN_FILE.read_text(encoding="utf-8"))
    total_rates = details["totals"]["rates"]
    user_rates = {user["user"]: user["rates"] for user in details["users"]}

    for record in data.get("records", []):
        if record.get("indicateur") != "taux_exploitation_qualipro":
            continue
        period = record.get("periode")
        if record.get("utilisateur") == "Total général":
            if period in total_rates:
                record["valeur"] = total_rates[period]
            continue
        user = record.get("utilisateur")
        if user in user_rates and period in user_rates[user]:
            record["valeur"] = user_rates[user][period]

    QHSE_CLEAN_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def update_kpi_metrics(details: dict):
    metrics = json.loads(KPI_METRICS_FILE.read_text(encoding="utf-8"))
    metrics = [row for row in metrics if row.get("indicateur") != "taux_exploitation_qualipro"]
    for period, value in details["totals"]["rates"].items():
        metrics.append({
            "section": "s9_kpi_qualipro___taux_d_exploitation",
            "site": "melt",
            "site_label": "MEL-T",
            "indicateur": "taux_exploitation_qualipro",
            "unite": "%",
            "periode": period,
            "valeur": value,
        })

    def period_sort_key(value: str):
        match = re.match(r"^(\d{4})(?:-(\d{2}))?$", value)
        if not match:
            return (9999, 99)
        return (int(match.group(1)), int(match.group(2) or 0))

    metrics.sort(key=lambda row: (
        row.get("section", ""),
        row.get("site_label", ""),
        row.get("indicateur", ""),
        period_sort_key(row.get("periode", "")),
    ))
    KPI_METRICS_FILE.write_text(json.dumps(metrics, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_details(details: dict):
    QUALIPRO_DETAILS_FILE.write_text(json.dumps(details, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main():
    details = parse_qualipro_details()
    update_qhse_clean(details)
    update_kpi_metrics(details)
    write_details(details)
    print("Qualipro sync complete.")


if __name__ == "__main__":
    main()
