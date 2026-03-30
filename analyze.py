"""
Performance Review Analysis Agent
Reads employee data from input/, computes metrics, generates HTML dashboard.
"""

import csv
import os
import json
import statistics
from datetime import datetime, date
from collections import defaultdict

INPUT_DIR = os.path.join(os.path.dirname(__file__), "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")


def load_data():
    """Load CSV or Excel file from input folder."""
    all_files = os.listdir(INPUT_DIR)
    excel_files = [f for f in all_files if f.endswith((".xlsx", ".xls", ".xlsm"))]
    csv_files = [f for f in all_files if f.endswith(".csv")]

    if excel_files:
        filepath = os.path.join(INPUT_DIR, excel_files[0])
        print(f"Loading Excel file: {excel_files[0]}")
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        row_iter = ws.iter_rows(values_only=True)
        raw_headers = next(row_iter)
        headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(raw_headers)]
        print(f"Detected columns: {headers}")
        rows = []
        for row in row_iter:
            if all(c is None for c in row):
                continue
            d = {headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))}
            rows.append(d)
        wb.close()
        return rows, headers
    elif csv_files:
        filepath = os.path.join(INPUT_DIR, csv_files[0])
        print(f"Loading CSV file: {csv_files[0]}")
        rows = []
        with open(filepath, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames
            print(f"Detected columns: {headers}")
            for r in reader:
                rows.append(r)
        return rows, headers
    else:
        raise FileNotFoundError("No CSV or Excel file found in input/ folder. Supported: .csv, .xlsx, .xls, .xlsm")


def parse_number(val):
    """Parse a number from string, handling commas, currency symbols, etc."""
    if val is None:
        return None
    val = str(val).strip().replace(",", "").replace("$", "").replace("€", "").replace("£", "").replace(" ", "")
    if val == "" or val.lower() == "n/a" or val == "-" or val.startswith("#"):
        return None
    try:
        return float(val)
    except ValueError:
        return None


def parse_date(val):
    """Try multiple date formats."""
    if val is None:
        return None
    val = str(val).strip()
    if not val or val.lower() == "n/a":
        return None
    formats = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y",
               "%Y/%m/%d", "%d.%m.%Y", "%m.%d.%Y", "%B %d, %Y", "%b %d, %Y",
               "%d %B %Y", "%d %b %Y"]
    for fmt in formats:
        try:
            return datetime.strptime(val, fmt).date()
        except ValueError:
            continue
    return None


def normalize_columns(headers):
    """Map raw column names to canonical names using robust fuzzy matching."""
    import re as _re
    mapping = {}

    # Normalize: lowercase, replace separators with space, collapse whitespace
    def norm(s):
        return _re.sub(r'\s+', ' ', s.lower().replace('_', ' ').replace('-', ' ')
                        .replace('(', ' ').replace(')', ' ').replace('/', ' ')).strip()

    lh = {norm(h): h for h in headers}
    lh_keys = list(lh.keys())

    def has_all(k, *words):
        return all(w in k for w in words)

    def has_all_not(k, yes, no):
        return all(w in k for w in yes) and not any(w in k for w in no)

    def find(*tests):
        for t in tests:
            for k in lh_keys:
                if isinstance(t, str) and k == t:
                    return lh[k]
                if hasattr(t, 'match') and t.search(k):
                    return lh[k]
                if callable(t) and not hasattr(t, 'match') and t(k):
                    return lh[k]
        return None

    # Identity
    mapping["id"] = find("id", "employee id", "emp id", "person id",
                         lambda k: k == "id")
    mapping["name"] = find("name", "nome", "employee name", "full name",
                           lambda k: has_all(k, "first", "name"), lambda k: k == "name")
    mapping["team"] = find("team", "department", "dept", "productgroup", "product group",
                           lambda k: has_all(k, "product", "group"), _re.compile(r'group'), _re.compile(r'\borg\b'))
    mapping["ranking"] = find("ranking", "rank", "rating",
                              lambda k: has_all(k, "performance", "rating"),
                              lambda k: has_all(k, "current", "ranking"),
                              _re.compile(r'ranking'), _re.compile(r'^rank$'))

    # Dates
    mapping["hire_date"] = find("hire", "hire date", "hiredate", "start date", "date hired",
                                "joined", "join date", _re.compile(r'hire'), _re.compile(r'start.?date'))

    # Base salary: detect year dynamically
    mapping["base_2026_hr"] = find(
        _re.compile(r'\d{4}.*hr.*base'), _re.compile(r'hr.*base.*\d{4}'),
        lambda k: has_all_not(k, ["hr", "base"], ["suggested", "sug", "incr", "gap", "algo"]))
    mapping["base_2026_suggested"] = find(
        _re.compile(r'\d{4}.*suggest.*base'), _re.compile(r'suggest.*base.*\d{4}'),
        lambda k: has_all_not(k, ["suggested", "base"], ["hr", "bonus", "incr", "gap", "algo"]),
        lambda k: has_all_not(k, ["sug", "base"], ["hr", "bonus", "incr", "gap", "algo"]))

    # Base increase
    mapping["base_incr_pct_input"] = find(
        lambda k: has_all(k, "base", "incr", "%"),
        lambda k: has_all(k, "base", "increase", "%"),
        _re.compile(r'base.?incr.*%'), _re.compile(r'base.?incr.*pct'))
    mapping["base_incr_eur_input"] = find(
        lambda k: has_all(k, "base", "incr", "eur"),
        _re.compile(r'base.?incr.*eur'))

    # Algo
    mapping["italy_algo"] = find(
        lambda k: has_all(k, "enrico", "algo"), "algo",
        _re.compile(r'\balgo\b'), lambda k: has_all(k, "algorithm"))
    mapping["base_incr_vs_algo"] = find(
        lambda k: has_all(k, "incr", "vs", "algo"),
        lambda k: has_all(k, "base", "vs", "algo"),
        _re.compile(r'vs.?algo'))

    # Bonus: detect year dynamically
    mapping["bonus_2025_hr"] = find(
        _re.compile(r'\d{4}.*hr.*bonus'), _re.compile(r'hr.*bonus.*\d{4}'),
        lambda k: has_all_not(k, ["hr", "bonus"], ["suggested", "sug", "incr", "gap"]))
    mapping["bonus_2025_suggested"] = find(
        _re.compile(r'\d{4}.*suggest.*bonus'), _re.compile(r'suggest.*bonus.*\d{4}'),
        lambda k: has_all_not(k, ["suggested", "bonus"], ["hr", "incr", "gap"]),
        lambda k: has_all_not(k, ["sug", "bonus"], ["hr", "incr", "gap"]))

    # Bonus increase
    mapping["bonus_incr_pct_input"] = find(
        lambda k: has_all(k, "bonus", "incr", "%"),
        _re.compile(r'bonus.?incr.*%'), _re.compile(r'bonus.?incr.*pct'))
    mapping["bonus_incr_eur_input"] = find(
        lambda k: has_all(k, "bonus", "incr", "eur"),
        _re.compile(r'bonus.?incr.*eur'))

    # Historical bases: find by year pattern, sorted descending
    base_years = sorted(
        [k for k in lh_keys if _re.search(r'\d{4}', k) and 'base' in k
         and 'hr' not in k and 'suggest' not in k and 'sug' not in k
         and 'incr' not in k and 'gap' not in k and 'algo' not in k],
        key=lambda k: int(_re.search(r'(\d{4})', k).group(1)), reverse=True)
    if len(base_years) >= 1:
        mapping["base_2025"] = lh[base_years[0]]
    if len(base_years) >= 2:
        mapping["base_2024"] = lh[base_years[1]]
    if len(base_years) >= 3:
        mapping["base_2023"] = lh[base_years[2]]
    if "base_2025" not in mapping:
        mapping["base_2025"] = find(lambda k: has_all(k, "current", "base"))

    bonus_years = sorted(
        [k for k in lh_keys if _re.search(r'\d{4}', k) and 'bonus' in k
         and 'hr' not in k and 'suggest' not in k and 'sug' not in k
         and 'incr' not in k and 'gap' not in k],
        key=lambda k: int(_re.search(r'(\d{4})', k).group(1)), reverse=True)
    if len(bonus_years) >= 1:
        mapping["bonus_2024"] = lh[bonus_years[0]]

    # Criticality / Potential
    mapping["criticality"] = find(_re.compile(r'criticality'), _re.compile(r'\bcrit\b'))
    mapping["potential"] = find(_re.compile(r'potential'), _re.compile(r'\bpot\b'))

    # Comments
    mapping["comments"] = find("comments", "comment", "note", "notes", "remark", "remarks",
                               _re.compile(r'comment'), _re.compile(r'remark'), _re.compile(r'\bnote'))

    # Clean nulls
    mapping = {k: v for k, v in mapping.items() if v is not None}

    mapped = set(mapping.values())
    unmapped = [h for h in headers if h not in mapped]
    if unmapped:
        print(f"Unmapped columns: {unmapped}")
    print(f"Column mapping: {mapping}")
    return mapping


def compute_italy_algo(ranking, hr_base, criticality, potential):
    """Compute the Italy Base Algo suggested base from ranking, HR base, criticality and potential."""
    if hr_base is None or ranking is None:
        return None
    rank = str(ranking).strip()
    # Bottom 30% and Bottom 10% → keep HR base
    if rank in ("4 - BOTTOM 30% (Next 20%)", "5 - BOTTOM 10%"):
        return round(hr_base, 2)
    # Top 10%
    if rank == "1 - TOP 10%":
        if hr_base < 36000:
            return round(hr_base * 1.12, 2)
        elif hr_base <= 50000:
            return round(hr_base * 1.10, 2)
        elif hr_base <= 70000:
            return round(hr_base * 1.07, 2)
        else:
            return round(hr_base * 1.03, 2)
    # Top 30% (Next 20%)
    if rank == "2 - TOP 30% (Next 20%)":
        if hr_base < 36000:
            return round(hr_base * 1.06, 2)
        elif hr_base <= 50000:
            return round(hr_base * 1.05, 2)
        elif hr_base <= 70000:
            return round(hr_base * 1.03, 2)
        else:
            return round(hr_base * 1.02, 2)
    # Middle
    if rank == "3 - MIDDLE":
        if hr_base > 65000:
            return round(hr_base, 2)
        crit = criticality or 0
        pot = potential or 0
        score = crit * 0.6 + pot * 0.4
        if score >= 5:
            if hr_base < 36000:
                return round(hr_base * 1.04, 2)
            elif hr_base <= 50000:
                return round(hr_base * 1.03, 2)
            else:
                return round(hr_base * 1.02, 2)
        elif score >= 3.5:
            if hr_base < 36000:
                return round(hr_base * 1.03, 2)
            elif hr_base <= 50000:
                return round(hr_base * 1.02, 2)
            else:
                return round(hr_base * 1.01, 2)
        else:
            return round(hr_base, 2)
    # Fallback
    return round(hr_base, 2)


def compute_metrics(rows, col_map):
    """Compute per-person metrics."""
    people = []
    for r in rows:
        person = {}
        person["id"] = str(r.get(col_map.get("id", ""), "Unknown") or "Unknown")
        person["name"] = str(r.get(col_map.get("name", ""), "") or "")
        person["team"] = str(r.get(col_map.get("team", ""), "Unknown") or "Unknown")
        person["ranking"] = str(r.get(col_map.get("ranking", ""), "Unknown") or "Unknown")
        hire_raw = r.get(col_map.get("hire_date", ""), "")
        # openpyxl may return datetime objects directly
        if isinstance(hire_raw, (datetime, date)):
            person["hire_date_raw"] = hire_raw.strftime("%Y-%m-%d")
            person["hire_date"] = hire_raw.date() if isinstance(hire_raw, datetime) else hire_raw
        else:
            person["hire_date_raw"] = str(hire_raw or "")
            person["hire_date"] = parse_date(person["hire_date_raw"])

        person["base_2026_hr"] = parse_number(r.get(col_map.get("base_2026_hr", ""), None))
        person["base_2026_sug"] = parse_number(r.get(col_map.get("base_2026_suggested", ""), None))
        person["bonus_2025_hr"] = parse_number(r.get(col_map.get("bonus_2025_hr", ""), None))
        person["bonus_2025_sug"] = parse_number(r.get(col_map.get("bonus_2025_suggested", ""), None))
        person["base_2025"] = parse_number(r.get(col_map.get("base_2025", ""), None))
        person["bonus_2024"] = parse_number(r.get(col_map.get("bonus_2024", ""), None))
        person["base_2024"] = parse_number(r.get(col_map.get("base_2024", ""), None))
        person["base_2023"] = parse_number(r.get(col_map.get("base_2023", ""), None))
        person["comments"] = str(r.get(col_map.get("comments", ""), "") or "").strip()
        person["criticality"] = parse_number(r.get(col_map.get("criticality", ""), None))
        person["potential"] = parse_number(r.get(col_map.get("potential", ""), None))
        person["italy_algo_input"] = parse_number(r.get(col_map.get("italy_algo", ""), None))
        person["base_incr_vs_algo_input"] = parse_number(r.get(col_map.get("base_incr_vs_algo", ""), None))
        # Pre-calculated columns from spreadsheet (if provided)
        person["base_incr_pct_input"] = parse_number(r.get(col_map.get("base_incr_pct_input", ""), None))
        person["base_incr_eur_input"] = parse_number(r.get(col_map.get("base_incr_eur_input", ""), None))
        person["bonus_incr_pct_input"] = parse_number(r.get(col_map.get("bonus_incr_pct_input", ""), None))
        person["bonus_incr_eur_input"] = parse_number(r.get(col_map.get("bonus_incr_eur_input", ""), None))

        # Tenure in years
        if person["hire_date"]:
            delta = date.today() - person["hire_date"]
            person["tenure_years"] = round(delta.days / 365.25, 1)
        else:
            person["tenure_years"] = None

        # --- Italy Base Algo: compute if not provided from spreadsheet ---
        if person["italy_algo_input"] is not None:
            person["italy_algo"] = round(person["italy_algo_input"], 2)
        else:
            person["italy_algo"] = compute_italy_algo(
                person["ranking"], person["base_2026_hr"],
                person["criticality"], person["potential"])

        # Base Incr vs Algo (suggested - algo)
        if person["base_incr_vs_algo_input"] is not None:
            person["base_incr_vs_algo"] = round(person["base_incr_vs_algo_input"], 2)
        elif person["base_2026_sug"] is not None and person["italy_algo"] is not None:
            person["base_incr_vs_algo"] = round(person["base_2026_sug"] - person["italy_algo"], 2)
        else:
            person["base_incr_vs_algo"] = None

        # --- Base increase analysis (suggested vs 2025) ---
        # Use pre-calculated if available, otherwise compute
        if person["base_incr_pct_input"] is not None:
            person["base_incr_pct"] = round(person["base_incr_pct_input"] * 100, 2)
        elif person["base_2025"] and person["base_2026_sug"] and person["base_2025"] > 0:
            person["base_incr_pct"] = round((person["base_2026_sug"] - person["base_2025"]) / person["base_2025"] * 100, 2)
        else:
            person["base_incr_pct"] = None

        if person["base_incr_eur_input"] is not None:
            person["base_incr_abs"] = round(person["base_incr_eur_input"], 2)
        elif person["base_2025"] and person["base_2026_sug"]:
            person["base_incr_abs"] = round(person["base_2026_sug"] - person["base_2025"], 2)
        else:
            person["base_incr_abs"] = None

        # HR pool base increase vs 2025
        if person["base_2025"] and person["base_2026_hr"] and person["base_2025"] > 0:
            person["base_hr_incr_abs"] = round(person["base_2026_hr"] - person["base_2025"], 2)
            person["base_hr_incr_pct"] = round((person["base_2026_hr"] - person["base_2025"]) / person["base_2025"] * 100, 2)
        else:
            person["base_hr_incr_abs"] = None
            person["base_hr_incr_pct"] = None

        # --- Base gap: suggested vs HR pool ---
        if person["base_2026_sug"] is not None and person["base_2026_hr"] is not None:
            person["base_gap_abs"] = round(person["base_2026_sug"] - person["base_2026_hr"], 2)
            if person["base_2026_hr"] > 0:
                person["base_gap_pct"] = round((person["base_2026_sug"] - person["base_2026_hr"]) / person["base_2026_hr"] * 100, 2)
            else:
                person["base_gap_pct"] = None
        else:
            person["base_gap_abs"] = None
            person["base_gap_pct"] = None

        # --- Bonus gap: suggested vs HR pool ---
        if person["bonus_2025_sug"] is not None and person["bonus_2025_hr"] is not None:
            person["bonus_gap_abs"] = round(person["bonus_2025_sug"] - person["bonus_2025_hr"], 2)
            if person["bonus_2025_hr"] > 0:
                person["bonus_gap_pct"] = round((person["bonus_2025_sug"] - person["bonus_2025_hr"]) / person["bonus_2025_hr"] * 100, 2)
            elif person["bonus_2025_sug"] > 0:
                person["bonus_gap_pct"] = 100.0  # HR gave 0, you suggest positive
            else:
                person["bonus_gap_pct"] = 0.0
        else:
            person["bonus_gap_abs"] = None
            person["bonus_gap_pct"] = None

        # --- Bonus change vs prior year (suggested bonus vs 2024 bonus) ---
        if person["bonus_incr_pct_input"] is not None:
            person["bonus_change_pct"] = round(person["bonus_incr_pct_input"] * 100, 2)
        elif person["bonus_2024"] is not None and person["bonus_2025_sug"] is not None and person["bonus_2024"] > 0:
            person["bonus_change_pct"] = round((person["bonus_2025_sug"] - person["bonus_2024"]) / person["bonus_2024"] * 100, 2)
        else:
            person["bonus_change_pct"] = None

        if person["bonus_incr_eur_input"] is not None:
            person["bonus_change_abs"] = round(person["bonus_incr_eur_input"], 2)
        elif person["bonus_2024"] is not None and person["bonus_2025_sug"] is not None:
            person["bonus_change_abs"] = round(person["bonus_2025_sug"] - person["bonus_2024"], 2)
        else:
            person["bonus_change_abs"] = None

        # Base increase 2024->2025
        if person["base_2024"] and person["base_2025"] and person["base_2024"] > 0:
            person["base_incr_prev_abs"] = round(person["base_2025"] - person["base_2024"], 2)
            person["base_incr_prev_pct"] = round((person["base_2025"] - person["base_2024"]) / person["base_2024"] * 100, 2)
        else:
            person["base_incr_prev_abs"] = None
            person["base_incr_prev_pct"] = None

        # Base increase 2023->2024
        if person["base_2023"] and person["base_2024"] and person["base_2023"] > 0:
            person["base_incr_2324_abs"] = round(person["base_2024"] - person["base_2023"], 2)
            person["base_incr_2324_pct"] = round((person["base_2024"] - person["base_2023"]) / person["base_2023"] * 100, 2)
        else:
            person["base_incr_2324_abs"] = None
            person["base_incr_2324_pct"] = None

        # Total comp
        if person["base_2025"] is not None and person["bonus_2025_sug"] is not None:
            person["total_comp_2025"] = person["base_2025"] + person["bonus_2025_sug"]
        else:
            person["total_comp_2025"] = None

        # Bonus as % of base (suggested bonus / 2025 base)
        if person["base_2025"] and person["bonus_2025_sug"] is not None and person["base_2025"] > 0:
            person["bonus_pct_of_base_2025"] = round(person["bonus_2025_sug"] / person["base_2025"] * 100, 2)
        else:
            person["bonus_pct_of_base_2025"] = None

        people.append(person)
    return people


def compute_team_stats(people):
    """Aggregate stats by team."""
    teams = defaultdict(list)
    for p in people:
        teams[p["team"]].append(p)

    team_stats = {}
    for team, members in teams.items():
        incrs = [p["base_incr_pct"] for p in members if p["base_incr_pct"] is not None]
        hr_incrs = [p["base_hr_incr_pct"] for p in members if p["base_hr_incr_pct"] is not None]
        bonuses = [p["bonus_change_pct"] for p in members if p["bonus_change_pct"] is not None]
        base_gaps = [p["base_gap_abs"] for p in members if p["base_gap_abs"] is not None]
        base_gap_pcts = [p["base_gap_pct"] for p in members if p["base_gap_pct"] is not None]
        bonus_gaps = [p["bonus_gap_abs"] for p in members if p["bonus_gap_abs"] is not None]
        bonus_gap_pcts = [p["bonus_gap_pct"] for p in members if p["bonus_gap_pct"] is not None]
        tenures = [p["tenure_years"] for p in members if p["tenure_years"] is not None]

        team_stats[team] = {
            "count": len(members),
            "avg_base_incr_pct": round(statistics.mean(incrs), 2) if incrs else None,
            "median_base_incr_pct": round(statistics.median(incrs), 2) if incrs else None,
            "min_base_incr_pct": round(min(incrs), 2) if incrs else None,
            "max_base_incr_pct": round(max(incrs), 2) if incrs else None,
            "std_base_incr_pct": round(statistics.stdev(incrs), 2) if len(incrs) > 1 else None,
            "avg_hr_incr_pct": round(statistics.mean(hr_incrs), 2) if hr_incrs else None,
            "avg_bonus_change_pct": round(statistics.mean(bonuses), 2) if bonuses else None,
            "avg_tenure": round(statistics.mean(tenures), 1) if tenures else None,
            "total_base_gap": round(sum(base_gaps), 0) if base_gaps else None,
            "avg_base_gap_pct": round(statistics.mean(base_gap_pcts), 2) if base_gap_pcts else None,
            "total_bonus_gap": round(sum(bonus_gaps), 0) if bonus_gaps else None,
            "avg_bonus_gap_pct": round(statistics.mean(bonus_gap_pcts), 2) if bonus_gap_pcts else None,
            "over_pool_base": sum(1 for p in members if p["base_gap_abs"] is not None and p["base_gap_abs"] > 0),
            "under_pool_base": sum(1 for p in members if p["base_gap_abs"] is not None and p["base_gap_abs"] < 0),
            "at_pool_base": sum(1 for p in members if p["base_gap_abs"] is not None and p["base_gap_abs"] == 0),
        }
    return team_stats


def compute_ranking_stats(people):
    """Aggregate stats by ranking."""
    rankings = defaultdict(list)
    for p in people:
        rankings[p["ranking"]].append(p)

    ranking_stats = {}
    for rank, members in rankings.items():
        incrs = [p["base_incr_pct"] for p in members if p["base_incr_pct"] is not None]
        hr_incrs = [p["base_hr_incr_pct"] for p in members if p["base_hr_incr_pct"] is not None]
        bonuses = [p["bonus_change_pct"] for p in members if p["bonus_change_pct"] is not None]
        bonus_pcts = [p["bonus_pct_of_base_2025"] for p in members if p["bonus_pct_of_base_2025"] is not None]
        base_gap_pcts = [p["base_gap_pct"] for p in members if p["base_gap_pct"] is not None]
        bonus_gap_pcts = [p["bonus_gap_pct"] for p in members if p["bonus_gap_pct"] is not None]

        ranking_stats[rank] = {
            "count": len(members),
            "avg_base_incr_pct": round(statistics.mean(incrs), 2) if incrs else None,
            "median_base_incr_pct": round(statistics.median(incrs), 2) if incrs else None,
            "min_base_incr_pct": round(min(incrs), 2) if incrs else None,
            "max_base_incr_pct": round(max(incrs), 2) if incrs else None,
            "avg_hr_incr_pct": round(statistics.mean(hr_incrs), 2) if hr_incrs else None,
            "avg_bonus_change_pct": round(statistics.mean(bonuses), 2) if bonuses else None,
            "avg_bonus_pct_of_base": round(statistics.mean(bonus_pcts), 2) if bonus_pcts else None,
            "avg_base_gap_pct": round(statistics.mean(base_gap_pcts), 2) if base_gap_pcts else None,
            "avg_bonus_gap_pct": round(statistics.mean(bonus_gap_pcts), 2) if bonus_gap_pcts else None,
        }
    return ranking_stats


def detect_anomalies(people, team_stats, ranking_stats):
    """Detect misalignments and anomalies."""
    flags = []

    for p in people:
        pid = p["id"]
        team = p["team"]
        rank = p["ranking"]

        ts = team_stats.get(team, {})
        rs = ranking_stats.get(rank, {})

        # Flag 1: Base increase significantly above/below team average
        if p["base_incr_pct"] is not None and ts.get("avg_base_incr_pct") is not None and ts.get("std_base_incr_pct") is not None:
            std = ts["std_base_incr_pct"] if ts["std_base_incr_pct"] and ts["std_base_incr_pct"] > 0 else 1
            z = (p["base_incr_pct"] - ts["avg_base_incr_pct"]) / std
            if abs(z) > 1.5:
                direction = "above" if z > 0 else "below"
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": "Team Outlier - Base Increase",
                    "severity": "High" if abs(z) > 2 else "Medium",
                    "detail": f"Base increase {p['base_incr_pct']}% is {abs(round(z,1))} std devs {direction} team avg ({ts['avg_base_incr_pct']}%)"
                })

        # Flag 2: High ranking but low increase (or vice versa)
        if p["base_incr_pct"] is not None and rs.get("avg_base_incr_pct") is not None:
            diff = p["base_incr_pct"] - rs["avg_base_incr_pct"]
            if rs.get("avg_base_incr_pct") and abs(diff) > 3:
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": "Ranking Misalignment - Base Increase",
                    "severity": "High" if abs(diff) > 5 else "Medium",
                    "detail": f"Base increase {p['base_incr_pct']}% vs ranking-group avg {rs['avg_base_incr_pct']}% (diff: {round(diff,1)}pp)"
                })

        # Flag 3: Bonus decrease while base increase is positive
        if p["bonus_change_abs"] is not None and p["base_incr_pct"] is not None:
            if p["bonus_change_abs"] < 0 and p["base_incr_pct"] > 3:
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": "Bonus/Base Divergence",
                    "severity": "Medium",
                    "detail": f"Bonus decreased by {abs(p['bonus_change_abs'])} while base increased {p['base_incr_pct']}%"
                })

        # Flag 4: Zero or negative base increase
        if p["base_incr_pct"] is not None and p["base_incr_pct"] <= 0:
            flags.append({
                "person": pid, "team": team, "ranking": rank,
                "type": "Zero/Negative Base Increase",
                "severity": "High",
                "detail": f"Base increase is {p['base_incr_pct']}% (flat or decrease)"
            })

        # Flag 5: Large swing from previous year's increase pattern
        if p["base_incr_pct"] is not None and p["base_incr_prev_pct"] is not None:
            swing = abs(p["base_incr_pct"] - p["base_incr_prev_pct"])
            if swing > 5:
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": "YoY Increase Volatility",
                    "severity": "Medium",
                    "detail": f"Base increase changed from {p['base_incr_prev_pct']}% (24→25) to {p['base_incr_pct']}% (25→26), swing of {round(swing,1)}pp"
                })

        # Flag 6: Tenure vs compensation check - long tenure with below-avg increase
        if p["tenure_years"] is not None and p["base_incr_pct"] is not None and ts.get("avg_base_incr_pct") is not None:
            if p["tenure_years"] > 5 and p["base_incr_pct"] < ts["avg_base_incr_pct"] - 2:
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": "Tenure Risk - Below Avg Increase",
                    "severity": "Medium",
                    "detail": f"Tenure {p['tenure_years']}y with base increase {p['base_incr_pct']}% vs team avg {ts['avg_base_incr_pct']}% — retention risk?"
                })

        # Flag 7: Large base gap vs HR pool (>5% over or under)
        if p["base_gap_pct"] is not None and abs(p["base_gap_pct"]) > 5:
            direction = "OVER" if p["base_gap_pct"] > 0 else "UNDER"
            flags.append({
                "person": pid, "team": team, "ranking": rank,
                "type": f"Base {direction} Pool",
                "severity": "High" if abs(p["base_gap_pct"]) > 10 else "Medium",
                "detail": f"Suggested base {direction.lower()} HR pool by {abs(p['base_gap_abs'])} ({abs(p['base_gap_pct'])}%): HR={p['base_2026_hr']}, Suggested={p['base_2026_sug']}"
            })

        # Flag 8: Large bonus gap vs HR pool
        if p["bonus_gap_abs"] is not None and abs(p["bonus_gap_abs"]) > 0:
            direction = "OVER" if p["bonus_gap_abs"] > 0 else "UNDER"
            severity = "High" if abs(p["bonus_gap_pct"] or 0) > 30 else "Medium" if abs(p["bonus_gap_abs"]) > 500 else None
            if severity:
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": f"Bonus {direction} Pool",
                    "severity": severity,
                    "detail": f"Suggested bonus {direction.lower()} HR pool by {abs(p['bonus_gap_abs'])} ({abs(p['bonus_gap_pct'] or 0)}%): HR={p['bonus_2025_hr']}, Suggested={p['bonus_2025_sug']}"
                })

        # Flag 9: High performer but suggested below pool
        if p["base_gap_abs"] is not None and p["base_gap_abs"] < 0:
            if rank and ("top 10" in rank.lower() or rank.startswith("1")):
                flags.append({
                    "person": pid, "team": team, "ranking": rank,
                    "type": "Top Performer Under Pool",
                    "severity": "High",
                    "detail": f"Ranked '{rank}' but suggested base ({p['base_2026_sug']}) is BELOW HR pool ({p['base_2026_hr']}). Gap: {p['base_gap_abs']}"
                })

    return flags


def generate_commentary(people, team_stats, ranking_stats, flags):
    """Generate compact text commentary (3 key lines) for the dashboard."""
    comments = []

    all_incrs = [p["base_incr_pct"] for p in people if p["base_incr_pct"] is not None]
    all_hr_incrs = [p["base_hr_incr_pct"] for p in people if p["base_hr_incr_pct"] is not None]
    base_gaps = [p["base_gap_abs"] for p in people if p["base_gap_abs"] is not None]
    bonus_gaps = [p["bonus_gap_abs"] for p in people if p["bonus_gap_abs"] is not None]
    high_flags = [f for f in flags if f["severity"] == "High"]
    med_flags = [f for f in flags if f["severity"] == "Medium"]

    # Line 1: Population + base increase summary
    parts = [f"{len(people)} employees, {len(team_stats)} teams"]
    if all_incrs:
        parts.append(f"Sug base incr avg {round(statistics.mean(all_incrs),1)}% / med {round(statistics.median(all_incrs),1)}%")
    if all_hr_incrs:
        parts.append(f"HR pool avg {round(statistics.mean(all_hr_incrs),1)}%")
    comments.append(" | ".join(parts))

    # Line 2: Pool gap summary
    over_pool = sum(1 for g in base_gaps if g > 0)
    under_pool = sum(1 for g in base_gaps if g < 0)
    at_pool = sum(1 for g in base_gaps if g == 0)
    net_base = round(sum(base_gaps)) if base_gaps else 0
    net_bonus = round(sum(bonus_gaps)) if bonus_gaps else 0
    comments.append(f"Base vs pool: {over_pool} over / {at_pool} at / {under_pool} under (net {net_base:,}) | Bonus net gap: {net_bonus:,}")

    # Line 3: Flags + top/bottom teams
    parts2 = [f"{len(high_flags)} high + {len(med_flags)} medium flags"]
    if team_stats:
        best = max(team_stats.items(), key=lambda x: x[1]["avg_base_incr_pct"] or 0)
        worst = min(team_stats.items(), key=lambda x: x[1]["avg_base_incr_pct"] or 999)
        parts2.append(f"Top: {best[0]} ({best[1]['avg_base_incr_pct']}%)")
        parts2.append(f"Low: {worst[0]} ({worst[1]['avg_base_incr_pct']}%)")
    comments.append(" | ".join(parts2))

    return comments


def build_html():
    """Copy the self-contained HTML dashboard template (starts empty, user imports data)."""
    template_path = os.path.join(os.path.dirname(__file__), "template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        return f.read()



def main():
    print("Loading data from input/ ...")
    rows, headers = load_data()
    print(f"Loaded {len(rows)} rows")

    col_map = normalize_columns(headers)
    people = compute_metrics(rows, col_map)
    print(f"Computed metrics for {len(people)} people")

    team_stats = compute_team_stats(people)
    ranking_stats = compute_ranking_stats(people)
    flags = detect_anomalies(people, team_stats, ranking_stats)
    comments = generate_commentary(people, team_stats, ranking_stats, flags)

    print(f"Detected {len(flags)} flags")
    print("Generating dashboard...")

    html = build_html()

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "performance_review_dashboard.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Dashboard saved to: {output_path}")
    print("Open it in a browser to view.")


if __name__ == "__main__":
    main()
