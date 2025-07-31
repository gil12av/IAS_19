'''
IAS 19 Part B: Annual Actuarial Reporting

Steps:
1. Load Part A results (PV Close), Opening Balances, and Employee Data for Part B.
2. Merge datasets on EmployeeID.
3. Select PV Close (allow override via separate file).
4. Calculate days fraction worked in 2024.
5. Compute Actuarial Factor per employee.
6. Calculate Service Cost (current year service cost).
7. Determine Discount Rate based on remaining service years.
8. Compute Interest Cost.
9. Compute Liability Actuarial Gain/Loss.
10. Export detailed employee report.

Input files (same folder):
- results.xlsx                 : Part A output with PV_close, LastSalary, Seniority, Section14Pct, Age_31_12_2024
- יתרות פתיחה.xlsx              : Opening balances with PV_open, Assets_open
- קובץ נתונים עבור חלק 2.xlsx : Employee data including start_date, leave_date, BenefitsPaidAssets, BenefitsPaidCash
- override_PV_close.xlsx      : (optional) EmployeeID, PV_close_override
'''
import pandas as pd
from datetime import datetime

# 1. Load data

def load_partA_results(path="results.xlsx"):
    """
    Returns DataFrame with Part A results:
    columns: EmployeeID, PV_close, LastSalary, Seniority, Section14Pct, Age_31_12_2024
    """
    return pd.read_excel(path)


def load_opening_balances(path="יתרות פתיחה.xlsx"):
    """
    Returns DataFrame with opening balances:
    columns: EmployeeID, PV_open, Assets_open
    """
    return pd.read_excel(path)


def load_employee_data(path="data_part2.xlsx"):
    """
    Returns DataFrame with employee info for Part B:
    columns: EmployeeID, start_date, leave_date,
             BenefitsPaidAssets, BenefitsPaidCash
    """
    df = pd.read_excel(path)
    # ensure datetime
    df['start_date'] = pd.to_datetime(df['start_date'])
    df['leave_date'] = pd.to_datetime(df['leave_date'], errors='coerce')
    # total benefits paid = assets + cash
    df['BenefitsPaid'] = df['BenefitsPaidAssets'].fillna(0) + df['BenefitsPaidCash'].fillna(0)
    return df[['EmployeeID','start_date','leave_date','BenefitsPaid']]


def load_override(path="override_PV_close.xlsx"):
    """
    Returns DataFrame with optional PV_close overrides:
    columns: EmployeeID, PV_close_override
    If file missing, returns empty DataFrame.
    """
    try:
        return pd.read_excel(path)
    except FileNotFoundError:
        return pd.DataFrame(columns=['EmployeeID','PV_close_override'])

# 2. Merge all

def merge_datasets(dfA, df_open, df_emp, df_override):
    """
    Merge Part A, opening balances, employee data, and overrides on EmployeeID.
    """
    df = dfA.merge(df_open, on='EmployeeID') \
            .merge(df_emp, on='EmployeeID') \
            .merge(df_override, on='EmployeeID', how='left')
    return df

# 3. Select PV_close to use

def select_pv_close(df):
    """
    If override provided, use PV_close_override, else PV_close.
    Adds column PV_close_used.
    """
    df['PV_close_used'] = df['PV_close_override'].fillna(df['PV_close'])
    return df

# 4. Days fraction worked in 2024

def calc_days_fraction(df):
    """
    Computes fraction of year 2024 worked:
    - 0 if left before 2024
    - 1 if still employed end of 2024
    - fractional if left during 2024
    Adds columns days_worked_2024, fraction_2024
    """
    start_2024 = datetime(2024,1,1)
    end_2024 = datetime(2024,12,31)
    def fraction(row):
        ld = row['leave_date']
        sd = row['start_date']
        # left before 2024
        if pd.notna(ld) and ld < start_2024:
            return 0.0
        # start after 2024 start
        ws = sd if sd > start_2024 else start_2024
        # end date
        we = ld if pd.notna(ld) and ld < end_2024 else end_2024
        days = (we - ws).days + 1
        return max(0, min(days, 365)) / 365.0
    df['fraction_2024'] = df.apply(fraction, axis=1)
    return df

# 5. Actuarial Factor

def calc_actuarial_factor(df):
    """
    Computes actuarial factor for each employee:
      PV_close_used / (LastSalary * Seniority * (1 - Section14Pct))
    Adds column ActFactor
    """
    df['ActFactor'] = df['PV_close_used'] / (
        df['LastSalary'] * df['Seniority'] * (1 - df['Section14Pct'])
    )
    return df

# 6. Service Cost

def calc_service_cost(df):
    """
    Computes Service Cost (SC) for 2024:
      LastSalary * fraction_2024 * (1 - Section14Pct) * ActFactor
    Adds column SC
    """
    df['SC'] = (
        df['LastSalary'] * df['fraction_2024'] *
        (1 - df['Section14Pct']) * df['ActFactor']
    )
    return df

# 7. Discount Rate selection

def lookup_discount_rate(years_left, df_assumptions):
    """
    Returns discount rate matching or nearest below years_left
    df_assumptions: columns Year, DiscountRate
    """
    # find max Year <= years_left
    df = df_assumptions[df_assumptions['Year'] <= years_left]
    if df.empty: return None
    return df.sort_values('Year', ascending=False).iloc[0]['DiscountRate']

# 8. Interest Cost

def calc_interest_cost(df, df_assumptions):
    """
    Computes Interest Cost (IC):
      (PV_open + SC - BenefitsPaid/2) * DiscountRate
    Adds column IC
    """
    # compute years left for each row
    def rate(row):
        ret_age = 64 if row['Age_31_12_2024'] == 'F' else 67
        yrs = ret_age - row['Age_31_12_2024']
        return lookup_discount_rate(yrs, df_assumptions)
    df['DiscRate'] = df.apply(rate, axis=1)
    df['IC'] = (
        (df['PV_open'] + df['SC'] - df['BenefitsPaid']/2)
        * df['DiscRate']
    )
    return df

# 9. Liability Actuarial Gain/Loss

def calc_liability_gain_loss(df):
    """
    Computes liability gain/loss:
      PV_close_used - PV_open - SC - IC + BenefitsPaid
    Adds column LiabGainLoss
    """
    df['LiabGainLoss'] = (
        df['PV_close_used'] - df['PV_open'] -
        df['SC'] - df['IC'] + df['BenefitsPaid']
    )
    return df

# 10. Main pipeline

def main():
    # load inputs
    dfA  = load_partA_results()
    dfO  = load_opening_balances()
    dfE  = load_employee_data()
    dfOv = load_override()
    df   = merge_datasets(dfA, dfO, dfE, dfOv)

    df = select_pv_close(df)
    df = calc_days_fraction(df)
    df = calc_actuarial_factor(df)
    df = calc_service_cost(df)

    # load assumptions for discount rates
    df_assump = pd.read_excel("data1.xlsx", sheet_name="הנחות")[['year', 'rate']]  # adjust names
    df = calc_interest_cost(df, df_assump)
    df = calc_liability_gain_loss(df)

    # output
    cols = [
        'EmployeeID','PV_open','PV_close_used','ActFactor',
        'SC','IC','BenefitsPaid','LiabGainLoss'
    ]
    df[cols].to_excel('IAS19_partB_results.xlsx', index=False)

if __name__ == '__main__':
    main()
