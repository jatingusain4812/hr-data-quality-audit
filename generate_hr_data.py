"""
HR Data Generator — Creates realistic HR dataset with intentional data quality issues
for audit and risk control demonstration.
"""

import pandas as pd
import random
import os
from datetime import date, timedelta

random.seed(42)

DEPARTMENTS = ["Engineering", "HR", "Finance", "Marketing", "Sales", "Operations", "Legal"]
ROLES = {
    "Engineering": ["Software Engineer", "Senior Engineer", "Tech Lead", "QA Analyst"],
    "HR": ["HR Executive", "HR Manager", "Recruiter", "HR Analyst"],
    "Finance": ["Financial Analyst", "Accountant", "Finance Manager", "Auditor"],
    "Marketing": ["Marketing Executive", "Content Writer", "SEO Analyst", "Brand Manager"],
    "Sales": ["Sales Executive", "Account Manager", "Sales Manager", "BDM"],
    "Operations": ["Operations Analyst", "Process Manager", "Ops Executive", "Coordinator"],
    "Legal": ["Legal Analyst", "Compliance Officer", "Legal Manager", "Paralegal"]
}
GENDERS = ["Male", "Female", "Other"]
STATUS = ["Active", "Inactive", "On Leave"]

def random_date(start_year=2015, end_year=2023):
    start = date(start_year, 1, 1)
    end = date(end_year, 12, 31)
    return start + timedelta(days=random.randint(0, (end - start).days))

def generate_employee(emp_id):
    dept = random.choice(DEPARTMENTS)
    role = random.choice(ROLES[dept])
    join_date = random_date(2015, 2023)
    age = random.randint(22, 58)
    salary = random.randint(25000, 150000)

    # Introduce data quality issues for ~20% of records
    issue = random.random()

    email = f"emp{emp_id}@company.com"
    phone = f"98{random.randint(10000000, 99999999)}"
    dob = date(2024 - age, random.randint(1,12), random.randint(1,28))
    status = random.choices(STATUS, weights=[75, 15, 10])[0]

    if issue < 0.05:
        email = ""                          # Missing email
    elif issue < 0.09:
        email = "not_an_email"              # Invalid email format
    if issue < 0.04:
        salary = None                       # Missing salary
    elif issue < 0.07:
        salary = -abs(salary)               # Negative salary (impossible)
    if issue < 0.03:
        phone = ""                          # Missing phone
    if issue < 0.06:
        dept = ""                           # Missing department
    if issue < 0.02:
        age = random.randint(10, 17)        # Underage employee (risk!)
    if issue < 0.03:
        dob = None                          # Missing DOB

    return {
        "Employee_ID": f"EMP{str(emp_id).zfill(4)}",
        "Full_Name": f"Employee_{emp_id}",
        "Age": age,
        "Date_of_Birth": dob,
        "Gender": random.choice(GENDERS),
        "Email": email,
        "Phone": phone,
        "Department": dept,
        "Job_Role": role,
        "Join_Date": join_date,
        "Salary": salary,
        "Employment_Status": status,
        "Manager_ID": f"EMP{str(random.randint(1, 50)).zfill(4)}" if emp_id > 50 else None,
        "PAN_Number": f"ABCDE{random.randint(1000,9999)}F" if random.random() > 0.1 else "",
        "Contract_Expiry": random_date(2022, 2025) if status == "Inactive" else None
    }

records = [generate_employee(i) for i in range(1, 501)]
df = pd.DataFrame(records)

os.makedirs("data", exist_ok=True)
df.to_csv("data/hr_raw_data.csv", index=False)
print(f"✅ Generated {len(df)} HR records → data/hr_raw_data.csv")
print(f"   Columns: {list(df.columns)}")
print(f"\n   Sample issues injected:")
print(f"   - Missing emails: {df['Email'].eq('').sum()}")
print(f"   - Invalid emails: {df[df['Email'].str.contains('@', na=True) == False & df['Email'].ne('')].shape[0]}")
print(f"   - Missing/negative salaries: {df['Salary'].isna().sum() + (df['Salary'].fillna(0) < 0).sum()}")
print(f"   - Missing departments: {df['Department'].eq('').sum()}")
print(f"   - Underage employees (age<18): {(df['Age'] < 18).sum()}")
