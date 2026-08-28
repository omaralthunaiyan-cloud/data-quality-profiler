"""
generate_examples.py
---------------------
Generates two sample .xlsx files with realistic, intentionally-flawed data
so the profiler has something interesting to demo on out of the box:

  examples/employees.xlsx  - HR dataset (emails, dates, salaries, IDs)
  examples/projects.xlsx   - Project tracker (budgets, status, dates, %)

Neither file is required to run the app — a user can upload any spreadsheet
of their own — these just make the demo one click.
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(7)
OUT_DIR = Path(__file__).resolve().parent


def _rand_date(start_year=2022, end_year=2026):
    start = date(start_year, 1, 1)
    end = date(end_year, 6, 1)
    return start + timedelta(days=random.randint(0, (end - start).days))


def gen_employees(n=180):
    first_names = ["Omar", "Sara", "Faisal", "Noura", "Khalid", "Lama", "Yousef", "Rima",
                   "Abdullah", "Hind", "Salman", "Reem", "Turki", "Maha", "Bandar"]
    last_names = ["Al-Saud", "Al-Qahtani", "Al-Otaibi", "Al-Harbi", "Al-Zahrani",
                  "Al-Ghamdi", "Al-Dosari", "Al-Shahri", "Al-Amri", "Al-Malki"]
    departments = ["Human Resources", "HR", "IT", "Information Technology", "Finance",
                   "finance", "Internal Audit", "Data & Intelligence"]
    rows = []
    for i in range(n):
        fn, ln = random.choice(first_names), random.choice(last_names)
        email = f"{fn.lower()}.{ln.lower().replace('al-','')}@company.com"
        r = random.random()
        if r < 0.10:
            email = email.replace("@", "_")  # broken
        elif r < 0.16:
            email = ""  # missing

        hire = _rand_date()
        hire_str = hire.isoformat()
        r2 = random.random()
        if r2 < 0.12:
            hire_str = hire.strftime("%d/%m/%Y")
        elif r2 < 0.18:
            hire_str = ""

        salary = random.randint(6000, 32000)
        if random.random() < 0.05:
            salary = -salary
        salary_val = "" if random.random() < 0.05 else salary

        national_id = f"1{random.randint(100000000,999999999)}"
        if random.random() < 0.10:
            national_id = str(random.randint(10000, 99999))  # too short, invalid

        rows.append({
            "employee_id": 1000 + i,
            "full_name": f"{fn} {ln}",
            "email": email,
            "department": random.choice(departments),
            "hire_date": hire_str,
            "salary_sar": salary_val,
            "national_id": national_id,
            "performance_rating": random.choice([1, 2, 3, 4, 5, 5, 4, 3]),
        })

    # duplicate a few rows
    for _ in range(6):
        rows.append(random.choice(rows).copy())

    return pd.DataFrame(rows)


def gen_projects(n=90):
    statuses = ["Active", "active", "On Hold", "Completed", "Cancelled", "COMPLETED"]
    priorities = ["High", "Medium", "Low", "high", "med"]
    rows = []
    for i in range(n):
        start = _rand_date(2023, 2026)
        end = start + timedelta(days=random.randint(30, 400))
        budget = random.randint(50_000, 2_000_000)
        if random.random() < 0.06:
            budget = -budget
        completion = random.randint(0, 100)
        if random.random() < 0.05:
            completion = random.randint(101, 150)  # impossible

        rows.append({
            "project_id": f"PRJ-{2000+i}",
            "project_name": f"Initiative {i+1}",
            "start_date": start.isoformat() if random.random() > 0.1 else start.strftime("%m-%d-%y"),
            "end_date": end.isoformat() if random.random() > 0.15 else "",
            "budget_sar": budget if random.random() > 0.06 else "",
            "status": random.choice(statuses),
            "priority": random.choice(priorities),
            "completion_pct": f"{completion}%",
            "project_manager_email": f"pm{random.randint(1,20)}@company.com" if random.random() > 0.08 else "n/a",
        })
    return pd.DataFrame(rows)


def main():
    emp = gen_employees()
    proj = gen_projects()
    emp.to_excel(OUT_DIR / "employees.xlsx", index=False)
    proj.to_excel(OUT_DIR / "projects.xlsx", index=False)
    print(f"Wrote {len(emp)} rows -> employees.xlsx")
    print(f"Wrote {len(proj)} rows -> projects.xlsx")


if __name__ == "__main__":
    main()
