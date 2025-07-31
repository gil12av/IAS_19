import pandas as pd
import sys
import io
from datetime import datetime


#----------------------קידוד להדפסות------------------------
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


#----------------------חלוקת ערכים ל3 מילונים-------------------------
employees_duplicates_dict ={}
employees_with_section_14 ={}
employees_without_section_14 ={}

#----------------------הפונקציה מחזירה את העובדים בצורת מילון---------employees[index][value]

def ReadExcelData():
    path = r'data/data1.xlsx'
    df = pd.read_excel(path, header=1)

    # שינוי שמות עמודות
    df.columns.values[0] = 'employee_id'
    df.columns.values[1] = 'first_name'
    df.columns.values[2] = 'last_name'
    df.columns.values[3] = 'gender'
    df.columns.values[4] = 'birth_date'
    df.columns.values[5] = 'start_work_date'
    df.columns.values[6] = 'last_salary'
    df.columns.values[7] = 'section_14_start_date'
    df.columns.values[8] = 'section_14_rate'
    df.columns.values[9] = 'assets_value'
    df.columns.values[10] = 'deposits'
    df.columns.values[11] = 'leave_date'
    df.columns.values[12] = 'asset_withdrawal'
    df.columns.values[13] = 'check_completion'
    df.columns.values[14] = 'leave_reason'

    # שמירה רק על 15 עמודות
    df = df.iloc[:, :15]

    # המרת תאריכים
    date_cols = ['birth_date', 'start_work_date', 'section_14_start_date', 'leave_date']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], format='%d/%m/%Y', errors='coerce')


    # יצירת מילונים
    duplicates_dict = {}
    with_section_14 = {}
    without_section_14 = {}

    # ספירת מופעים של כל employee_id
    id_counts = df['employee_id'].value_counts()

    for _, row in df.iterrows():
        emp_id = row['employee_id']
        emp_data = {
            'first_name': row['first_name'],
            'last_name': row['last_name'],
            'gender': row['gender'],
            'birth_date': row['birth_date'],
            'start_work_date': row['start_work_date'],
            'last_salary': row['last_salary'],
            'section_14_start_date': row['section_14_start_date'],
            'section_14_rate': row['section_14_rate'],
            'assets_value': row['assets_value'],
            'leave_date': row['leave_date'],
            'deposits': row['deposits'],
            'asset_withdrawal': row['asset_withdrawal'],
            'check_completion': row['check_completion'],
            'leave_reason': row['leave_reason']
        }

        if id_counts[emp_id] > 1:
            duplicates_dict[emp_id] = emp_data
        elif pd.notna(row['section_14_rate']) and row['section_14_rate'] > 0:
            with_section_14[emp_id] = emp_data
        else:
            without_section_14[emp_id] = emp_data

    return duplicates_dict, with_section_14, without_section_14

#------------------------יצירת מילונים מתוך לוח תמותה - עבור גברים-----------------------------
def read_male_mortality_table():
    df = pd.read_excel('data/לוח תמותה משרד האוצר.xlsx', header=1,sheet_name=0)

    df = df[['age', 'L(x)', 'P(x)', 'q(x)']].dropna()


    male_mortality_table_age_Lx = dict(zip(df['age'], df['L(x)']))
    male_mortality_table_age_Px = dict(zip(df['age'], df['P(x)']))
    male_mortality_table_age_Qx = dict(zip(df['age'], df['q(x)']))

    return male_mortality_table_age_Px, male_mortality_table_age_Lx, male_mortality_table_age_Qx

#------------------------יצירת מילונים מתוך לוח תמותה - עבור נשים-----------------------------
def read_Female_mortality_table():
    df = pd.read_excel('data/לוח תמותה משרד האוצר.xlsx', header=1,sheet_name=1)

    df = df[['age', 'L(x)', 'P(x)', 'q(x)']].dropna()

    Female_mortality_table_age_Lx = dict(zip(df['age'], df['L(x)']))
    Female_mortality_table_age_Px = dict(zip(df['age'], df['P(x)']))
    Female_mortality_table_age_Qx = dict(zip(df['age'], df['q(x)']))

    return Female_mortality_table_age_Px, Female_mortality_table_age_Lx, Female_mortality_table_age_Qx
#--------------------------------הסתברויות עזיבה כפי שנתון----------------------------------------
def leave_probabilities(age, leave_reason): 
    age_ranges = {
        '18-29': {'fired': 0.10, 'resigned': 0.15, 'total': 0.25},
        '30-39': {'fired': 0.06, 'resigned': 0.10, 'total': 0.16},
        '40-49': {'fired': 0.04, 'resigned': 0.09, 'total': 0.13},
        '50-59': {'fired': 0.04, 'resigned': 0.05, 'total': 0.09},
        '60-67': {'fired': 0.03, 'resigned': 0.03, 'total': 0.06}
    }
    if 18 <= age <= 29:
        return age_ranges['18-29'][leave_reason]
    elif 30 <= age <= 39:
        return age_ranges['30-39'][leave_reason]
    elif 40 <= age <= 49:
        return age_ranges['40-49'][leave_reason]
    elif 50 <= age <= 59:
        return age_ranges['50-59'][leave_reason]
    elif 60 <= age <= 67:
        return age_ranges['60-67'][leave_reason]
    
    return 0.0 


#----------------------------הנחות---שיעור היוון ---------------------------------------------------
def getDiscountRate(index):
    discount_rates = {
        1: 0.0181, 2: 0.0199, 3: 0.0211, 4: 0.0221, 5: 0.0230,
        6: 0.0239, 7: 0.0246, 8: 0.0253, 9: 0.0260, 10: 0.0267,
        11: 0.0274, 12: 0.0280, 13: 0.0286, 14: 0.0292, 15: 0.0299,
        16: 0.0305, 17: 0.0311, 18: 0.0317, 19: 0.0323, 20: 0.0329,
        21: 0.0335, 22: 0.0341, 23: 0.0348, 24: 0.0354, 25: 0.0360,
        26: 0.0366, 27: 0.0372, 28: 0.0378, 29: 0.0384, 30: 0.0391,
        31: 0.0397, 32: 0.0403, 33: 0.0409, 34: 0.0415, 35: 0.0421,
        36: 0.0427, 37: 0.0434, 38: 0.0440, 39: 0.0446, 40: 0.0452,
        41: 0.0458, 42: 0.0464, 43: 0.0470, 44: 0.0476, 45: 0.0483,
        46: 0.0489, 47: 0.0495
    }
    return discount_rates.get(index) 
#--------------------------------------עובדים-אתחול מילונים-----------------------------------------
employees_duplicates_dict , employees_with_section_14 , employees_without_section_14 =  ReadExcelData()

#--------------------------------------הסתברויות-אתחול מילונים-----------------------------------------
male_mortality_table_age_Px, male_mortality_table_age_Lx, male_mortality_table_age_Qx = read_male_mortality_table()
Female_mortality_table_age_Px, Female_mortality_table_age_Lx, Female_mortality_table_age_Qx =read_Female_mortality_table()




"""print("Duplicated:", len(employees_duplicates_dict))
for emp_id in employees_duplicates_dict:
    print(emp_id)

print("With סעיף 14:", len(employees_with_section_14))
print("Without סעיף 14:", len(employees_without_section_14))"""

def getgender(index,dict): 
    return dict[index]['gender']

def getLastSalary(index,dict): # משכורת אחרונה
    return dict[index]['last_salary']

def getSalaryGrowthRate(index): 
    if(index % 2 == 0):
        return 0.04
    else:
        return 0.02
    


employees =employees_without_section_14


#---------------------------# משכורת אחרונה - מידע ---------------------------------
def getLastSalary(index,dict): # משכורת אחרונה
    return dict[index]['last_salary']
#Seniority 

#-----------------------------אחוז סעיף 14 מידע-----------------------------------

#------------------------------נכסים - מידע -----------------------------------

def getAssetsValue(index,dict): #   נכסים
    return dict[index]['assets_value']


def getLeave_reason(index, dict):
    reason = dict[index]['leave_reason']
    if pd.notna(reason):
        return reason.strip()
    return None


#-------------------------------עליית שכר - הנחות -----------------------------------
def getSalaryGrowthRate(index): 
    if(index % 2 == 0):
        return 0.04
    else:
        return 0.02


#-------------------------------גבר לוח תמותה------------------------------
def getMale_Px(index):
    return male_mortality_table_age_Px[index]

def getMale_Lx(index):
    return male_mortality_table_age_Lx[index]

def getMale_qx(index):
    return male_mortality_table_age_Qx[index]
#-------------------------------אישה לוח תמותה------------------------------

def getFemale_Px(index):
    return Female_mortality_table_age_Px[index]
def getFemale_Lx(index):
    return Female_mortality_table_age_Lx[index]
def getFemale_qx(index):
    return Female_mortality_table_age_Qx[index]

# ---------------------------הסתברויות-------גבר--------------------------------

def Male_survive(age):
    """
    האםשרות שגבר בגיל מסויים ישרוד
     \n
    or -> return getMale_Lx(age +1) / getMale_Lx(age)
    """
    return getMale_Px(age) 

def Male_die(age):#גבר מוות
    return 1 - getMale_Px(age)

# ----------------------------הסתברויות ------אישה--------------------------------
def Female_survive(age):
    """
     האםשרות שאישה בגיל מסויים תשרוד
     \n
     or -> return getMale_Lx(age +1) / getMale_Lx(age) 
    """
    return getFemale_Lx(age) 

def Female_die(age):
    """
     האפשרות שאישה בגיל מסויים תמות
     \n
     1 - getFemale_Px(age) 
    """
    return 1 - getFemale_Px(age) 

def Male_survive_until(age,age_t):
   """
   ההסתברות שגבר בגיל \n
   X \n
   ישרוד עד גיל \n
   X+T
   """
   return getMale_Lx(age_t) / getMale_Lx(age)

def Female_survive_until(age,age_t):
   """
   ההסתברות שאישה בגיל \n
   X \n
   תשרוד עד גיל \n
   X+T
   """
   return getFemale_Lx(age_t) / getFemale_Lx(age)

def Female__die_in_t_plus_x(age,age_t):
    """
    ההסתברות שאישה בגיל \n
    X\n
    תמות בגיל 
    \nX + T
    """
    return ( getFemale_Lx(age) - getFemale_Lx(age_t) ) / getFemale_Lx(age)

def Male__die_in_t_plus_x(age,age_t):
    """
    ההסתברות שגבר בגיל \n
    X\n
    ימות בגיל 
    \nX + T
    """
    return ( getMale_Lx(age) - getMale_Lx(age_t) ) / getMale_Lx(age)

# ----------------------------הסתברויות --------------------------------------

def To_resign(age):# התפטרות
    return leave_probabilities(age, 'resigned')

def To_fired(age):# פיטורין
    return leave_probabilities(age, 'fired') 

def getseniority_without_section_14(index,dict):
    ref_date=datetime(2024, 12, 31)
    start_date = dict[index]['start_work_date']
    return round((ref_date.date() - start_date.date()).days / 365.25, 2)


def getAgeAtFixedDate(index,dict, ref_date=datetime(2024, 12, 31)):
    """
    הפונקציה מחזירה את הגיל בתאריך הבדיקה 
    """
    birth_date = dict[index]['birth_date']
    if pd.notna(birth_date):
        return (ref_date.date() - birth_date.date()).days // 365
    return None



def calc(index):
    """
    מחשב את ההתחייבות לעובד ללא סעיף 14 נכון ל-31.12.2024, כולל טיפול מותאם בהתפטרות, פיטורין ופרישה.
    """
    ref_date = datetime(2024, 12, 31)
    leave_date = employees[index]['leave_date']
    if pd.notna(leave_date) and ref_date >= leave_date:
        return 0

    age = getAgeAtFixedDate(index,employees)
    seniority = getseniority_without_section_14(index,employees)
    LastSalary = getLastSalary(index, employees)
    AssetsValue = getAssetsValue(index, employees)
    gender = getgender(index, employees)
    retirement_age = 64 if gender == "F" else 67
    retirement_years = retirement_age - age

    section14_rate = 0
    base = LastSalary * seniority * (1 - section14_rate)
    SalaryGrowthRate = getSalaryGrowthRate(index)
    salary_increase_start_date = datetime(2025, 6, 30)
    survival_prob = 1
    total_calc = 0

    reason = getLeave_reason(index, employees)
    leave_year = employees[index]['leave_date'].year if pd.notna(employees[index]['leave_date']) else None

    for t in range(retirement_years):
        current_year = 2024 + t + 1
        current_date = datetime(current_year, 12, 31)
        age_t = age + t + 1

        # עליית שכר כל שנתיים החל מ-2025
        increase_times = max(0, ((current_year - 2025) // 2) + 1) if current_date >= salary_increase_start_date else 0
        growth = (1 + SalaryGrowthRate) ** increase_times

        discount_rate = getDiscountRate(t + 1)
        if discount_rate is None:
            continue
        discount_factor = (1 + discount_rate) ** (t + 0.5)

        qR = To_resign(age_t)
        qF = To_fired(age_t)
        qD = getFemale_qx(age_t) if gender == 'F' else getMale_qx(age_t)

        not_left = survival_prob
        if not_left < 0:
            continue

        #  טיפול בעובד שעזב בפועל בשנה הנוכחית
        if leave_year == current_year:
            if reason == "התפטרות":
                part_resign = AssetsValue * growth
                total_calc += part_resign
                break
            elif reason == "פיטורין":
                part_dismissal = LastSalary * seniority
                part_resign = AssetsValue * growth
                total_calc += part_dismissal + part_resign
                break
            else:
                break  # סיבת עזיבה לא רלוונטית <- אין התחייבות

        #  אם זו השנה האחרונה ואין עזיבה בפועל → פרישה רגילה
        if t == retirement_years - 1 and reason is None:
            part_retirement = LastSalary * seniority * not_left
            total_calc += part_retirement
            break

        #  חישוב לפי הסתברויות
        part_dismissal = base * growth * qF * not_left / discount_factor
        part_death = base * growth * qD * not_left / discount_factor
        part_resign = AssetsValue * growth * qR * not_left / discount_factor # ask if to fix ?

        total_calc += part_dismissal + part_death + part_resign

        # עדכון הסתברות הישרדות
        survival_prob *= (1 - qF - qR - qD)

    return round(total_calc)







expected_results = {
    1: 0, 7: 0, 9: 0, 11: 440190, 13: 328249, 15: 61294, 17: 142884, 19: 50000,
    2: 16538, 4: 187359, 10: 11900, 12: 0, 16: 179905, 18: 98224
}

for id in expected_results:
    if id in employees:
        result = round(calc(id))
        status = "=" if result == expected_results[id] else f"!= (Expected {expected_results[id]})"
        print(f"{id} - {result} {status}")

for id in [4, 11]:
    print(f"--- עובד {id} ---")
    print("משכורת אחרונה:", getLastSalary(id, employees))
    print("נכסים:", getAssetsValue(id, employees))
    print("וותק:", getseniority_without_section_14(id,employees))
    print("גיל:", getAgeAtFixedDate(id,employees))
    print("תוצאה:", calc(id))
    print()

  

def get_non_section14_seniority(index, employee_dict, ref_date=datetime(2024, 12, 31)):
    """
    מחשבת את משך התקופה (בשנים) שהעובד עבד לפני כניסת סעיף 14 לתוקף.
    """
    start_date = employee_dict[index]['start_work_date']
    section14_date = employee_dict[index]['section_14_start_date']

    if pd.isna(start_date):
        return 0

    if pd.isna(section14_date) or section14_date <= start_date:
        return 0

    end = min(section14_date, ref_date)
    non_covered_days = (end.date() - start_date.date()).days
    return round(non_covered_days / 365.25)

def getsection14_rate(index,dict):
    """

    """
    section_14_rate = dict[index]['section_14_rate']
    if pd.notna(section_14_rate):
        return section_14_rate
    return None

def getsection_14_start_date(index,dict):
    """

    """
    section_14_start_date = dict[index]['section_14_start_date']
    if pd.notna(section_14_start_date):
        return section_14_start_date
    return None

def calc_with_section14(index):
    """
    מחשב את ההתחייבות לעובד עם סעיף 14 נכון ל-31.12.2024,
    כולל חישוב מותאם לפי מועד ואחוז סעיף 14.
    """
    ref_date = datetime(2024, 12, 31)
    employee = employees_with_section_14[index]

    leave_date = employee['leave_date']
    if pd.notna(leave_date) and ref_date >= leave_date:
        return 0

    section14_rate = getsection14_rate(index, employees_with_section_14)
    section_14_start_date = getsection_14_start_date(index, employees_with_section_14)
    start_work = employee['start_work_date']
    seniority = getseniority_without_section_14(index,employees_with_section_14)

    # אם סעיף 14 מלא (100%) מההתחלה - אין התחייבות
    if section14_rate == 100 and section_14_start_date == start_work:
        return 0

    # המרה לאחוז יחסי
    section14_rate /= 100
    #print(f"---------------------------------{section14_rate}")

    age = getAgeAtFixedDate(index, employees_with_section_14)
    uncovered_years = get_non_section14_seniority(index, employees_with_section_14)
    
    LastSalary = getLastSalary(index, employees_with_section_14)
    AssetsValue = getAssetsValue(index, employees_with_section_14)
    #print(f" uncovered_years {uncovered_years} AssetsValue {AssetsValue}")
    gender = getgender(index, employees_with_section_14)
    retirement_age = 64 if gender == "F" else 67
    retirement_years = retirement_age - age

    SalaryGrowthRate = getSalaryGrowthRate(index)
    salary_increase_start_date = datetime(2025, 6, 30)
    survival_prob = 1
    total_calc = 0

    reason = getLeave_reason(index, employees_with_section_14)
    leave_year = leave_date.year if pd.notna(leave_date) else None

    for t in range(retirement_years):
        current_year = 2024 + t + 1
        current_date = datetime(current_year, 12, 31)
        age_t = age + t + 1

        # עליית שכר
        increase_times = max(0, ((current_year - 2025) // 2) + 1) if current_date >= salary_increase_start_date else 0
        growth = (1 + SalaryGrowthRate) ** increase_times

        # עדכון אחוז סעיף 14 לפי השנה
        if pd.notna(section_14_start_date) and current_date >= section_14_start_date:
            current_rate = section14_rate
        else:
            current_rate = 0
        #print(f"\t\tsection14_rate -> {section14_rate}  in year {current_date.year} senerity {round(seniority,0)} current_rate14 {current_rate}")
        base = LastSalary * round(seniority,0) * (1 - current_rate)

        discount_rate = getDiscountRate(t + 1)
        if discount_rate is None:
            continue
        discount_factor = (1 + discount_rate) ** (t + 0.5)

        qR = To_resign(age_t)
        qF = To_fired(age_t)
        qD = getFemale_qx(age_t) if gender == 'F' else getMale_qx(age_t)
        not_left = survival_prob

        if not_left < 0:
            continue

        # טיפול במקרה של עזיבה בפועל
        if leave_year == current_year:
            if reason == "התפטרות":
                total_calc += AssetsValue * growth
                break
            elif reason == "פיטורין":
                part_dismissal = LastSalary * round(seniority,0) * (1 - current_rate)
                total_calc += part_dismissal + AssetsValue * growth
                break
            else:
                break

        # שנה אחרונה - פרישה
        if t == retirement_years - 1 and reason is None:
            part_retirement = LastSalary * round(seniority,0) * (1 - current_rate) * not_left
            total_calc += part_retirement + AssetsValue
            #print(f"\t\tThe employee has reached retirement age and is entitled to a retirement pension in the year {current_year}.")
            break

        # חישוב הסתברויות
        part_dismissal = base * growth * qF * not_left / discount_factor
        part_death = base * growth * qD * not_left / discount_factor
        part_resign = AssetsValue * growth * qR * not_left / discount_factor

        total_calc += part_dismissal + part_death + part_resign

        survival_prob *= (1 - qF - qR - qD)

    return round(total_calc)

# נבנה רשימת תוצאות
results = []

# עובדים ללא סעיף 14
for emp_id, emp_data in employees_without_section_14.items():
    liability = round(calc(emp_id))
    results.append({
        "employee_id": emp_id,
        "first_name": emp_data["first_name"],
        "last_name": emp_data["last_name"],
        "liability": liability
    })

# עובדים עם סעיף 14
for emp_id, emp_data in employees_with_section_14.items():
    liability = round(calc_with_section14(emp_id))
    results.append({
        "employee_id": emp_id,
        "first_name": emp_data["first_name"],
        "last_name": emp_data["last_name"],
        "liability": liability
    })

# שמירה לקובץ Excel
df = pd.DataFrame(results)
df.to_excel("partA_output.xlsx", index=False)

print(f"✓ נשמר בהצלחה! נמצאו {len(results)} עובדים בסך הכול.")