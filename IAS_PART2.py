"""
IAS 19 – Part B: דוח שנתי אקטוארי (Service Cost, Interest Cost, Gain/Loss)
----------------------------------------
תהליך מלא:
1. טוען 'data1.xlsx' גליון 'data' – נתוני עובדים (start, leave, משיכות נכסים, השלמות צ’ק).
2. בוחר רק את הרשומה העדכנית (במקרה של כפילויות).
3. טוען 'open_Balance.xlsx' גליון 'Sheet1' – יתרות פתיחה (PV_open, Assets_open).
4. טוען 'results.xlsx' – פלט Part A (PV_close, LastSalary, Seniority, Section14Pct, Age_31_12_2024).
5. טוען 'data1.xlsx' גליון 'הנחות' – עקום שיעורי היוון לפי שנות שירות.
6. מאחד את כל ה-DataFrame’s לפי employee_id.
7. מחשב חלק השנה (fraction_2024) לפי start_work_date / leave_date.
8. מחשב Actuarial Factor.
9. מחשב Service Cost.
10. בוחר שיעור היוון ע"פ שנות שארית שירות.
11. מחשב Interest Cost.
12. מחשב רווח/הפסד אקטוארי – התחייבות.
13. מייצא תוצאה ל-'IAS19_partB_results.xlsx'.
"""

import pandas as pd
from datetime import datetime
import os

# ----------------------------
# 1. Load data1 (גליון data)
# ----------------------------
def load_data1(path="data/data1.xlsx"):
    """
    טוען גליון 'data' מ-data1.xlsx, header=1, ממפה עמודות:
      employee_id, start_work_date, leave_date,
      asset_withdrawal, check_completion
    בוחר את הרשומה האחרונה לכל עובד (במקרה של כפילויות).
    """
    df = pd.read_excel(path, sheet_name="data", header=1)
  
    # מיפוי שמות עמודות
    df = df.rename(columns={
        df.columns[0]:  "employee_id",
        df.columns[3]:  "gender",
        df.columns[4]:  "date_of_birth",
        df.columns[5]:  "start_work_date",
        df.columns[6]:  "LastSalary",
        df.columns[8]:  "Section14Pct",
        df.columns[11]: "leave_date",
        df.columns[12]: "asset_withdrawal",
        df.columns[13]: "check_completion",
    })
    # המרה לתאריכים
    df["start_work_date"] = pd.to_datetime(df["start_work_date"], dayfirst=True, errors="coerce")
    df["leave_date"]      = pd.to_datetime(df["leave_date"], dayfirst=True, errors="coerce")
    df["date_of_birth"]   = pd.to_datetime(df["date_of_birth"], dayfirst=True, errors="coerce")
   
    # חישוב ותק ל־31.12.2024
    ref_date = pd.Timestamp("2024-12-31")
    df["Seniority"] = ((ref_date - df["start_work_date"]).dt.days / 365).fillna(0).astype(int)

    # חישוב גיל ל־31.12.2024
    df["Age_31_12_2024"] = ((ref_date - df["date_of_birth"]).dt.days / 365).fillna(0).astype(int)

    # חישוב BenefitsPaid
    df["BenefitsPaid"] = df["asset_withdrawal"].fillna(0) + df["check_completion"].fillna(0)
   
    # בוחר הרשומה ה"עדכנית" ביותר לכל employee_id
    df = df.sort_values("start_work_date").drop_duplicates("employee_id", keep="last")
    
      # החזרת שדות רלוונטיים
    return df[[
        "employee_id",
        "gender",
        "start_work_date",
        "leave_date",
        "LastSalary",
        "Section14Pct",
        "Seniority",
        "Age_31_12_2024",
        "BenefitsPaid"
    ]]

# ---------------------------------------
# 2. Load opening balances (גליון Sheet1)
# ---------------------------------------
def load_open_balance(path="data/open_Balance.xlsx"):
    """
    טוען Sheet1 מ-open_Balance.xlsx:
      מספר עובד, ערך נוכחי התחייבות, שווי הוגן
    ממפה ל-employee_id, PV_open, Assets_open
    """
    df = pd.read_excel(path, sheet_name="Sheet1")
    df = df.rename(columns={
        "מספר עובד": "employee_id",
        "ערך נוכחי התחייבות": "PV_open",
        "שווי הוגן": "Assets_open",
    })
    return df[["employee_id","PV_open","Assets_open"]]

# ----------------------------------------
# 3. Load Part A results (results.xlsx)
# ----------------------------------------
def load_partA(path="data/partA_output.xlsx"):
    """
    טוען פלט Part A (partA_output.xlsx):
    - employee_id
    - liability (ממופה ל־PV_close)
    """
    df = pd.read_excel(path)

    # שינוי שם עמודה ראשונה אם צריך
    if "employee_id" not in df.columns:
        df = df.rename(columns={df.columns[0]: "employee_id"})

    # שינוי עמודת liability לשם אחיד
    df = df.rename(columns={"liability": "PV_close"})

    # מחזירים רק את השדות הדרושים
    return df[["employee_id", "PV_close"]]

# --------------------------------------------------
# 4. Load discount rate assumptions (גליון הנחות)
# --------------------------------------------------
def load_assumptions(path="data/data1.xlsx"):
    """
    טוען גליון 'הנחות' מ-data1.xlsx:
      עמודות: שנה, שיעור היוון
    ממפה ל-Year, DiscountRate
    """
    df = pd.read_excel(path, sheet_name="היוון")
    df = df.rename(columns={
    "year": "Year",
    "discountRate" : "DiscountRate"
    })

    return df[["Year","DiscountRate"]]

# ---------------------------
# 5. Merge all DataFrames
# ---------------------------
def merge_all(df1, dfA, dfO):
    """
    מאחד: df1 (data1), dfA (Part A), dfO (opening balances)
    לפי employee_id.
    """
    df = df1.merge(dfA, on="employee_id")\
            .merge(dfO, on="employee_id")
    return df

# ---------------------------------
# 6. Compute fraction of 2024
# ---------------------------------
def calc_fraction_2024(df):
    """
    מוסיף עמודה fraction_2024:
      0 אם left_date < 1/1/2024;
      1 אם still employed בסוף 2024;
      אחרת days/365.
    """
    start_2024 = datetime(2024,1,1)
    end_2024   = datetime(2024,12,31)

    def frac(row):
        ld = row["leave_date"]
        sd = row["start_work_date"]
        # left before 2024
        if pd.notna(ld) and ld < start_2024:
            return 0.0
        # start of work window
        ws = sd if sd > start_2024 else start_2024
        # end of work window
        we = ld if pd.notna(ld) and ld < end_2024 else end_2024
        days = (we - ws).days + 1
        return max(0, min(days, 365)) / 365.0

    df["fraction_2024"] = df.apply(frac, axis=1)
    return df

# ---------------------------
# 7. Actuarial Factor
# ---------------------------
def calc_actuarial_factor(df):
    """
    מוסיף עמודה ActFactor:
      PV_close / (LastSalary * Seniority * (1 - Section14Pct))
    """
    df["ActFactor"] = (
        df["PV_close"] /
        (df["LastSalary"] * df["Seniority"] * (1 - df["Section14Pct"]))
    )
    return df

# -------------------------------------------------
# 8. Service Cost (עלות שירות שוטף)
# -------------------------------------------------
def calc_service_cost(df):
    """
    מוסיף עמודה SC:
      LastSalary * fraction_2024 * (1 - Section14Pct) * ActFactor
    """
    df["SC"] = (
        df["LastSalary"] *
        df["fraction_2024"] *
        (1 - df["Section14Pct"]) *
        df["ActFactor"]
    )
    return df

# -------------------------------------------------
# 9. Lookup Discount Rate by years left
# -------------------------------------------------
def lookup_discount_rate(years_left, df_assump):
    """
    מחזיר DiscountRate מתאימה ל-years_left:
    העולה על השנה הקרובה ביותר אך לא יותר ממנה.
    """
    # בוחרים את השורות שה-Year <= years_left
    df = df_assump[df_assump["Year"] <= years_left]
    if df.empty:
        return None
    # הסדר יורד ובחירה בראשונה
    return df.sort_values("Year", ascending=False).iloc[0]["DiscountRate"]

# -------------------------------------------------
# 10. Interest Cost (עלות היוון)
# -------------------------------------------------
def calc_interest_cost(df, df_assump):
    """
    מוסיף עמודות DiscRate, IC:
      DiscRate = lookup_discount_rate(retirement_age - Age_31_12_2024)
      IC = (PV_open + SC - BenefitsPaid/2) * DiscRate
    """
    def get_rate(row):
        ret_age = 64 if row["Age_31_12_2024"] == "F" else 67
        years_left = ret_age - row["Age_31_12_2024"]
        return lookup_discount_rate(years_left, df_assump)

    df["DiscRate"] = df.apply(get_rate, axis=1)
    df["IC"] = (
        (df["PV_open"] + df["SC"] - df["BenefitsPaid"]/2) *
        df["DiscRate"]
    )
    return df

# -------------------------------------------------
# 11. Liability Actuarial Gain/Loss
# -------------------------------------------------
def calc_liability_gain_loss(df):
    """
    מוסיף עמודה LiabGainLoss:
      PV_close - PV_open - SC - IC + BenefitsPaid
    """
    df["LiabGainLoss"] = (
        df["PV_close"] - df["PV_open"]
        - df["SC"] - df["IC"] + df["BenefitsPaid"]
    )
    return df

# ---------------------------
# 12. Main
# ---------------------------
def main():
    # Load inputs
    df1    = load_data1("data/data1.xlsx")
    dfO    = load_open_balance("data/open_Balance.xlsx")
    dfA    = load_partA("data/partA_output.xlsx")
    dfAss  = load_assumptions("data/data1.xlsx")

    # Merge
    df = merge_all(df1, dfA, dfO)

    # Part B calculations
    print(df.dtypes)
    df = calc_fraction_2024(df)
    df = calc_actuarial_factor(df)
    df = calc_service_cost(df)
    df = calc_interest_cost(df, dfAss)
    df = calc_liability_gain_loss(df)

    # Export
    cols = [
        "employee_id", "PV_open", "PV_close",
        "ActFactor", "SC", "DiscRate", "IC",
        "BenefitsPaid", "LiabGainLoss"
    ]
    df[cols].to_excel("IAS19(partB)-results.xlsx", index=False)
    print("IAS19 Part B report exported to IAS19_partB_results.xlsx")


if __name__ == "__main__":
    main()

