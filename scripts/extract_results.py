"""
Извлечение и агрегирование результатов EnergyPlus (6 сценариев):
Base, DCV, Base_HP, DCV_HP, Base_Rec, DCV_Rec.

Выход:
  /Users/m/Documents/Openstudio/results/results.json
"""

import json
import os
import sqlite3
import re
from typing import Dict, List, Optional

ROOT_DIR = "/Users/m/Documents/Openstudio"
RESULTS_DIR = os.path.join(ROOT_DIR, "results")
RESULTS_PATH = os.path.join(RESULTS_DIR, "results.json")

SCENARIOS = {
    "Base":     os.path.join(ROOT_DIR, "sim_Base",     "run", "eplusout.sql"),
    "DCV":      os.path.join(ROOT_DIR, "sim_DCV",      "run", "eplusout.sql"),
    "Base_HP":  os.path.join(ROOT_DIR, "sim_Base_HP",  "run", "eplusout.sql"),
    "DCV_HP":   os.path.join(ROOT_DIR, "sim_DCV_HP",   "run", "eplusout.sql"),
    "Base_Rec": os.path.join(ROOT_DIR, "sim_Base_Rec", "run", "eplusout.sql"),
    "DCV_Rec":  os.path.join(ROOT_DIR, "sim_DCV_Rec",  "run", "eplusout.sql"),
}

ERR_PATHS = {
    "Base":     os.path.join(ROOT_DIR, "sim_Base",     "run", "eplusout.err"),
    "DCV":      os.path.join(ROOT_DIR, "sim_DCV",      "run", "eplusout.err"),
    "Base_HP":  os.path.join(ROOT_DIR, "sim_Base_HP",  "run", "eplusout.err"),
    "DCV_HP":   os.path.join(ROOT_DIR, "sim_DCV_HP",   "run", "eplusout.err"),
    "Base_Rec": os.path.join(ROOT_DIR, "sim_Base_Rec", "run", "eplusout.err"),
    "DCV_Rec":  os.path.join(ROOT_DIR, "sim_DCV_Rec",  "run", "eplusout.err"),
}


def query_one(db_path: str, sql: str, params=()) -> float:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(sql, params)
    row = cur.fetchone()
    conn.close()
    if not row or row[0] is None:
        return 0.0
    return float(row[0])


def query_all(db_path: str, sql: str, params=()) -> List[tuple]:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def sum_report_data(db_path: str, name: str, freq: str, key: Optional[str] = None) -> float:
    if key:
        return query_one(
            db_path,
            """
            SELECT SUM(rd.Value)
            FROM ReportData rd
            JOIN ReportDataDictionary d
              ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
            WHERE d.Name = ? AND d.ReportingFrequency = ? AND d.KeyValue = ?
            """,
            (name, freq, key),
        )
    return query_one(
        db_path,
        """
        SELECT SUM(rd.Value)
        FROM ReportData rd
        JOIN ReportDataDictionary d
          ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
        WHERE d.Name = ? AND d.ReportingFrequency = ?
        """,
        (name, freq),
    )


def monthly_kwh(db_path: str, name: str, freq: str = "Monthly", key: Optional[str] = None) -> List[float]:
    if key:
        rows = query_all(
            db_path,
            """
            SELECT t.Month, SUM(rd.Value)
            FROM ReportData rd
            JOIN ReportDataDictionary d
              ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
            JOIN Time t ON rd.TimeIndex = t.TimeIndex
            WHERE d.Name = ? AND d.ReportingFrequency = ? AND d.KeyValue = ?
              AND t.Month IS NOT NULL
            GROUP BY t.Month
            ORDER BY t.Month
            """,
            (name, freq, key),
        )
    else:
        rows = query_all(
            db_path,
            """
            SELECT t.Month, SUM(rd.Value)
            FROM ReportData rd
            JOIN ReportDataDictionary d
              ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
            JOIN Time t ON rd.TimeIndex = t.TimeIndex
            WHERE d.Name = ? AND d.ReportingFrequency = ?
              AND t.Month IS NOT NULL
            GROUP BY t.Month
            ORDER BY t.Month
            """,
            (name, freq),
        )
    by_month = {int(m): float(v) / 3_600_000.0 for m, v in rows}
    return [by_month.get(i, 0.0) for i in range(1, 13)]


def mean_hourly(db_path: str, var_name: str, key: str, min_value: Optional[float] = None) -> float:
    if min_value is None:
        sql = """
            SELECT AVG(rd.Value)
            FROM ReportData rd
            JOIN ReportDataDictionary d
              ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
            WHERE d.Name = ? AND d.KeyValue = ? AND d.ReportingFrequency = 'Hourly'
        """
        return query_one(db_path, sql, (var_name, key))

    sql = """
        SELECT AVG(rd.Value)
        FROM ReportData rd
        JOIN ReportDataDictionary d
          ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
        WHERE d.Name = ? AND d.KeyValue = ? AND d.ReportingFrequency = 'Hourly'
          AND rd.Value > ?
    """
    return query_one(db_path, sql, (var_name, key, min_value))


def pct(base: float, value: float) -> float:
    if abs(base) < 1e-9:
        return 0.0
    return (value - base) / base * 100.0


def calc_scenario(name: str, db_path: str) -> Dict:
    zone_key = f"Кинозал_135мест_{name.upper()}"
    coil_key = f"Калорифер_П1_{name.upper()}"

    elec_kwh = sum_report_data(db_path, "Electricity:Facility", "Monthly") / 3_600_000.0
    fan_kwh = sum_report_data(db_path, "Fans:Electricity", "Monthly") / 3_600_000.0
    coil_kwh = sum_report_data(db_path, "Heating Coil Heating Energy", "Monthly", key=coil_key) / 3_600_000.0

    vent_avg_m3s = mean_hourly(
        db_path,
        "Zone Mechanical Ventilation Current Density Volume Flow Rate",
        zone_key,
        min_value=0.001,
    )
    co2_avg_ppm = mean_hourly(
        db_path,
        "Zone Air CO2 Concentration",
        zone_key,
        min_value=400.0,
    )

    occ_hours = int(
        query_one(
            db_path,
            """
            SELECT COUNT(*)
            FROM ReportData rd
            JOIN ReportDataDictionary d
              ON rd.ReportDataDictionaryIndex = d.ReportDataDictionaryIndex
            WHERE d.Name='Zone People Occupant Count'
              AND d.KeyValue=?
              AND d.ReportingFrequency='Hourly'
              AND rd.Value > 10
            """,
            (zone_key,),
        )
    )

    return {
        "elec_kwh": elec_kwh,
        "fan_kwh": fan_kwh,
        "coil_kwh": coil_kwh,
        "vent_avg_m3h": vent_avg_m3s * 3600.0,
        "co2_avg_ppm": co2_avg_ppm,
        "occ_hours": occ_hours,
        "monthly": {
            "electricity_kwh": monthly_kwh(db_path, "Electricity:Facility"),
            "fans_kwh": monthly_kwh(db_path, "Fans:Electricity"),
            "coil_kwh": monthly_kwh(db_path, "Heating Coil Heating Energy", key=coil_key),
        },
    }


def parse_err_summary(err_path: str) -> Dict[str, Optional[int]]:
    if not os.path.exists(err_path):
        return {"warnings": None, "severe": None, "completed": False}

    with open(err_path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()

    completed = "EnergyPlus Completed Successfully" in text
    m = re.search(r"Completed Successfully--\s*(\d+) Warning;\s*(\d+) Severe Errors", text)
    if not m:
        return {"warnings": None, "severe": None, "completed": completed}
    return {"warnings": int(m.group(1)), "severe": int(m.group(2)), "completed": completed}


def main() -> None:
    print("=" * 72)
    print("  СБОР РЕЗУЛЬТАТОВ ENERGYPLUS")
    print("=" * 72)

    scenarios: Dict[str, Dict] = {}
    missing = []

    for name, db_path in SCENARIOS.items():
        if not os.path.exists(db_path):
            missing.append(name)
            print(f"[!] Нет SQL для {name}: {db_path}")
            continue

        scenarios[name] = calc_scenario(name, db_path)
        scenarios[name]["run_summary"] = parse_err_summary(ERR_PATHS.get(name, ""))
        print(
            f"[OK] {name:<7} | Elec={scenarios[name]['elec_kwh']:.1f} кВт·ч | "
            f"Fan={scenarios[name]['fan_kwh']:.1f} кВт·ч | "
            f"CO2={scenarios[name]['co2_avg_ppm']:.1f} ppm"
        )

    if "Base" not in scenarios:
        raise RuntimeError("Нужен как минимум сценарий Base для сравнений.")

    base = scenarios["Base"]
    comparisons = {}

    for name, row in scenarios.items():
        comparisons[name] = {
            "vs_base": {
                "elec_delta_kwh": row["elec_kwh"] - base["elec_kwh"],
                "elec_delta_pct": pct(base["elec_kwh"], row["elec_kwh"]),
                "fan_delta_kwh": row["fan_kwh"] - base["fan_kwh"],
                "fan_delta_pct": pct(base["fan_kwh"], row["fan_kwh"]),
                "coil_delta_kwh": row["coil_kwh"] - base["coil_kwh"],
                "coil_delta_pct": pct(base["coil_kwh"], row["coil_kwh"]),
                "co2_delta_ppm": row["co2_avg_ppm"] - base["co2_avg_ppm"],
                "vent_delta_pct": pct(base["vent_avg_m3h"], row["vent_avg_m3h"]),
            }
        }

    # Парные сравнения внутри пары (HP и не-HP)
    pairwise = {}
    if "DCV" in scenarios:
        dcv = scenarios["DCV"]
        pairwise["DCV_vs_Base"] = {
            "elec_save_kwh": base["elec_kwh"] - dcv["elec_kwh"],
            "elec_save_pct": -pct(base["elec_kwh"], dcv["elec_kwh"]),
            "fan_save_kwh": base["fan_kwh"] - dcv["fan_kwh"],
            "fan_save_pct": -pct(base["fan_kwh"], dcv["fan_kwh"]),
        }

    if "Base_HP" in scenarios and "DCV_HP" in scenarios:
        base_hp = scenarios["Base_HP"]
        dcv_hp = scenarios["DCV_HP"]
        pairwise["DCV_HP_vs_Base_HP"] = {
            "elec_save_kwh": base_hp["elec_kwh"] - dcv_hp["elec_kwh"],
            "elec_save_pct": -pct(base_hp["elec_kwh"], dcv_hp["elec_kwh"]),
            "fan_save_kwh": base_hp["fan_kwh"] - dcv_hp["fan_kwh"],
            "fan_save_pct": -pct(base_hp["fan_kwh"], dcv_hp["fan_kwh"]),
        }

    if "Base_Rec" in scenarios and "DCV_Rec" in scenarios:
        base_rec = scenarios["Base_Rec"]
        dcv_rec = scenarios["DCV_Rec"]
        pairwise["DCV_Rec_vs_Base_Rec"] = {
            "elec_save_kwh": base_rec["elec_kwh"] - dcv_rec["elec_kwh"],
            "elec_save_pct": -pct(base_rec["elec_kwh"], dcv_rec["elec_kwh"]),
            "fan_save_kwh": base_rec["fan_kwh"] - dcv_rec["fan_kwh"],
            "fan_save_pct": -pct(base_rec["fan_kwh"], dcv_rec["fan_kwh"]),
        }

    payload = {
        "project": {
            "name": "Кинозал Петербург-кино",
            "location": "Санкт-Петербург",
            "generated_at": __import__("datetime").datetime.now().isoformat(timespec="seconds"),
        },
        "scenarios": scenarios,
        "comparisons": comparisons,
        "pairwise": pairwise,
        "missing": missing,
    }

    os.makedirs(RESULTS_DIR, exist_ok=True)
    with open(RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print("-" * 72)
    print(f"[OK] Сохранено: {RESULTS_PATH}")
    print("=" * 72)


if __name__ == "__main__":
    main()
