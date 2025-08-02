
"""
IAS 19 – Part B: Annual roll‑forward (Service Cost, Interest Cost, Actuarial G/L)
────────────────────────────────────────────────────────────────────────────
This script builds on the results of Part A (present value of the obligation
at 31‑12‑2024) and the opening balances to produce the movement schedule
required for the financial statements.

Required input files (all in the same folder as the script unless a full
path is supplied):
• data1.xlsx                – two sheets:
    ‑ "data"       : employee master file (see README below)
    ‑ "הנחות"      : discount‑rate yield curve (columns: Year, DiscountRate)
• open_Balance.xlsx         – opening balances (columns: employee_id, PV_open, Assets_open)
• partA_output.xlsx         – Part A results (columns: employee_id, PV_close)

Output:
• IAS19_partB_results.xlsx  – one sheet with the precise column layout you
  provided (obligation section on the left, assets section on the right).

README – mandatory columns in sheet "data" (header row is row 2 → header=1):
 0 "מספר עובד"          ⇒ employee_id (unique key)
 3 "מין"                 ⇒ gender  ("M"/"F") – determines retirement age 67/64
 4 "תאריך לידה"          ⇒ date_of_birth (dd/mm/yyyy)
 5 "תאריך תחילת עבודה"  ⇒ start_work_date
 6 "שכר "                ⇒ LastSalary (ILS)
 7 "תאריך  קבלת סעיף 14" ⇒ section14_start_date (not needed but useful)
 8 "אחוז סעיף 14"        ⇒ Section14Pct (e.g. 100, 72 …) /100 will be used
 9 "שווי נכס"            ⇒ Assets_close (fair value at year‑end, if known)
10 "הפקדות"              ⇒ deposits (employer and employee contributions 2024)
11 "תאריך עזיבה "        ⇒ leave_date (dd/mm/yyyy or "-"/blank if active)
12 "תשלום מהנכס"        ⇒ withdrawal_from_assets (benefits paid from plan assets)
13 "השלמה בצ'ק"          ⇒ completion_by_cheque (benefits paid directly by employer)

Feel free to rename columns in Excel – just adjust the mapping dictionaries
below accordingly.
"""

from __future__ import annotations
import pandas as pd
from datetime import datetime
from pathlib import Path

###########################################################################
# 0. CONFIGURATION ########################################################
###########################################################################
DATA_FOLDER = Path("./data")        # adapt if your files reside elsewhere
FILE_DATA1          = DATA_FOLDER / "data1.xlsx"
FILE_OPEN_BALANCES  = DATA_FOLDER / "open_Balance.xlsx"
FILE_PARTA_RESULTS  = DATA_FOLDER / "partA_output.xlsx"
FILE_OUTPUT         = "IAS19_part2_results.xlsx"
REPORT_DATE         = pd.Timestamp("2024-12-31")
RET_AGE_M, RET_AGE_F = 67, 64  # statutory retirement ages in Israel
###########################################################################

############################################################################
# 1. LOAD INPUTS ###########################################################
############################################################################

def load_employees(path: str | Path = FILE_DATA1) -> pd.DataFrame:
    """Load *sheet "data"* from **data1.xlsx** and standardise column names."""
    df = pd.read_excel(path, sheet_name="data", header=1)

    # Column mapping – adjust indices/names only if the Excel layout changes
    mapper = {
        "מספר עובד": "employee_id",
        "מין": "gender",
        "תאריך לידה": "date_of_birth",
        "תאריך תחילת עבודה ": "start_work_date",
        "שכר ": "LastSalary",
        "אחוז סעיף 14": "Section14Pct",
        "תאריך עזיבה ": "leave_date",
        "תשלום מהנכס": "withdrawal_from_assets",
        "השלמה בצ'ק": "completion_by_cheque",
        "הפקדות": "deposits",
        "שווי נכס": "Assets_close",  # value at 31‑12‑2024 if provided
    }
    df = df.rename(columns=mapper)

    # Basic cleaning / typing
    date_cols = ["date_of_birth", "start_work_date", "leave_date"]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce", format="%d/%m/%Y")

    num_cols = ["LastSalary", "Section14Pct", "withdrawal_from_assets",
                "completion_by_cheque", "deposits", "Assets_close"]
    df[num_cols] = df[num_cols].fillna(0)

    # Section 14: convert 100 ⇒ 1.00, 72 ⇒ 0.72 …
    df["Section14Pct"] = df["Section14Pct"] / 100.0

    # Choose the latest record per employee in case of duplicates
    df = (
        df.sort_values("start_work_date")
          .drop_duplicates(subset="employee_id", keep="last")
          .reset_index(drop=True)
    )

    # Derived fields
    df["Age_31_12_2024"] = ((REPORT_DATE - df["date_of_birth"]).dt.days / 365.25)
    df["Seniority"]      = ((REPORT_DATE - df["start_work_date"]).dt.days / 365.25)

    # Benefits paid by employer (liability side)
    df["BenefitsPaid"] = df["withdrawal_from_assets"] + df["completion_by_cheque"]

    return df


def load_opening_balances(path: str | Path = FILE_OPEN_BALANCES) -> pd.DataFrame:
    """Sheet 1 must contain columns: employee_id, PV_open, Assets_open."""
    df = pd.read_excel(path, sheet_name=0)
    mapper = {
        "מספר עובד": "employee_id",
        "ערך נוכחי התחייבות": "PV_open",
        "שווי הוגן": "Assets_open",
    }
    df = df.rename(columns=mapper)
    return df[["employee_id", "PV_open", "Assets_open"]]


def load_partA_results(path: str | Path = FILE_PARTA_RESULTS) -> pd.DataFrame:
    """Part A results – columns: employee_id, liability."""
    df = pd.read_excel(path)
    if "liability" not in df.columns:
        raise ValueError("partA_output.xlsx must contain a column named 'liability'.")
    return df.rename(columns={"liability": "PV_close"})[["employee_id", "PV_close"]]


def load_discount_curve(path: str | Path = FILE_DATA1) -> pd.DataFrame:
    """Sheet "היוון" – columns: Year, DiscountRate (as decimal, e.g. 0.0253)."""
    df = pd.read_excel(path, sheet_name="היוון")
    mapper = {
        df.columns[0]: "Year",
        df.columns[1]: "DiscountRate",
    }
    df = df.rename(columns=mapper)
    return df[["Year", "DiscountRate"]]

############################################################################
# 2. HELPER FUNCTIONS ######################################################
############################################################################

def years_of_future_service(row: pd.Series) -> float:
    """Expected future service from 31‑12‑2024 to statutory retirement age."""
    retirement_age = RET_AGE_F if row["gender"].strip().upper() == "F" else RET_AGE_M
    return max(retirement_age - row["Age_31_12_2024"], 0)


def lookup_discount_rate(years_left: float, curve: pd.DataFrame) -> float:
    """Nearest *lower or equal* tenor – if years_left<min(year) take the first."""
    eligible = curve[curve["Year"] <= years_left]
    if eligible.empty:
        return curve.iloc[0]["DiscountRate"]
    return eligible.sort_values("Year", ascending=False).iloc[0]["DiscountRate"]

############################################################################
# 3. CALCULATION STEPS #####################################################
############################################################################

def enrich_calculations(df: pd.DataFrame, curve: pd.DataFrame) -> pd.DataFrame:
    # 3.1 fraction of the year worked in 2024
    start_2024 = pd.Timestamp("2024-01-01")
    end_2024   = REPORT_DATE

    def fraction_2024(row):
        if pd.notna(row["leave_date"]) and row["leave_date"] < start_2024:
            return 0.0
        work_start = max(start_2024, row["start_work_date"])
        work_end   = min(end_2024, row["leave_date"]) if pd.notna(row["leave_date"]) else end_2024
        return (work_end - work_start).days / 365.25

    df["fraction_2024"] = df.apply(fraction_2024, axis=1)

    # 3.2 Actuarial factor
    divisor = df["LastSalary"] * df["Seniority"] * (1 - df["Section14Pct"])
    df["ActFactor"] = df["PV_close"] / divisor.replace({0: pd.NA}) # ----> לבדוק את השורה הזו !!!!!!!!!!!

    # 3.3 Service Cost (SC)
    df["SC"] = (
        df["LastSalary"] * df["fraction_2024"] * (1 - df["Section14Pct"]) * df["ActFactor"]
    )

    # 3.4 Discount rate & Interest Cost (IC)
    df["YearsLeft"] = df.apply(years_of_future_service, axis=1)
    df["DiscRate"]  = df["YearsLeft"].apply(lambda y: lookup_discount_rate(y, curve))

    # Formula per lecture: IC = [(PV_open * DiscRate) + ((SC – BenefitsPaid) × (DiscRate/2)) ]
    df["IC"] = ((df["PV_open"]* df["DiscRate"]) + ((df["SC"] - df["BenefitsPaid"])) * (df["DiscRate"]/2))

    # 3.5 Liability actuarial gain / loss
    df["LiabGainLoss"] = (
        df["PV_close"] - df["PV_open"] - df["SC"] - df["IC"] + df["BenefitsPaid"]
    )

    # 3.6 Expected return on assets (same rate as discount unless curve has separate column)
    df["ER"] = ((df["Assets_open"] * df["DiscRate"]) + ((df["deposits"] - df["withdrawal_from_assets"]) * (df["DiscRate"]/2)))

    # 3.7 Asset actuarial gain / loss
    df["AssetGainLoss"] = (
        df["Assets_close"] - df["Assets_open"] - df["ER"] - df["deposits"] + df["withdrawal_from_assets"]
    )

    return df

############################################################################
# 4. MAIN AND EXPORT ######################################################
############################################################################

def main():
    # ---------- load data ----------
    df_emp  = load_employees()
    df_open = load_opening_balances()
    df_A    = load_partA_results()
    curve   = load_discount_curve()

    # ---------- merge ----------
    df = (df_emp.merge(df_open, on="employee_id", how="left")
                  .merge(df_A,    on="employee_id", how="left"))

    # ---------- calculate ----------
    df = enrich_calculations(df, curve)

    # ---------- column order & export ----------
    obligation_cols = [
        "employee_id", "PV_open", "SC", "IC", "BenefitsPaid",
        "LiabGainLoss", "PV_close", "ActFactor",
    ]
    asset_cols = [
        "Assets_open", "ER", "deposits", "withdrawal_from_assets",
        "AssetGainLoss", "Assets_close",
    ]

    # --------- changing column to hebrew for better excel file -----
    rename_dict = {
        "employee_id": "מספר עובד",
        "PV_open": "יתרת פתיחה",
        "SC": "עלות שירות שוטף",
        "IC": "עלות היוון",
        "BenefitsPaid": "הטבות ששולמו",
        "LiabGainLoss": "הפסד אקטוארי",
        "PV_close": "יתרת סגירה",
        "ActFactor": "פקטור אקטוארי",
        "Assets_open": "יתרת פתיחה.1",
        "ER": "תשואה צפויה",
        "deposits": "הפקדות",
        "withdrawal_from_assets": "הטבות ששולמו מנכסים",
        "AssetGainLoss": "רווח אקטוארי",
        "Assets_close": "יתרת סגירה.1",
    }

    desired_order = obligation_cols + asset_cols

    df_out = df[desired_order].rename(columns=rename_dict).sort_values("מספר עובד")
    df_out.to_excel(FILE_OUTPUT, index=False)
    print(f"✓ Results exported → {FILE_OUTPUT}  (rows: {len(df_out)})")


if __name__ == "__main__":
    main()
    