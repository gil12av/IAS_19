import pandas as pd

# קריאה עם כותרת משורה 2
df = pd.read_excel("data/data1.xlsx", sheet_name="data", header=1)

# --- שלב א: הסרת עמודות ריקות לגמרי (Unnamed וכו') ---
df = df.dropna(axis=1, how='all')

# --- שלב ב: ניקוי שמות עמודות מרווחים מיותרים ---
df.columns = df.columns.str.strip()  # מסיר רווחים מכל כיוון

# --- שלב ג: הדפסה לבדיקה ---
print("---- שמות עמודות לאחר ניקוי ----")
print(df.columns)

print("---- תצוגת נתונים ----")
print(df.head(50))

print(df['שווי נכס'])



