"""Generate a fake, deliberately messy HR master file for testing.
All names and numbers are invented."""
import random
from datetime import date, timedelta
import pandas as pd

random.seed(42)
N = 600
first = ["Aarav", "Fatima", "John", "Maria", "Wei", "Omar", "Priya", "Ahmed", "Sara", "Rahul",
         "Aisha", "David", "Li", "Noura", "Carlos", "Anita", "Yusuf", "Elena", "Ravi", "Hana"]
last = ["Sharma", "Khan", "Smith", "Garcia", "Zhang", "Ali", "Nair", "Hassan", "Lopez", "Patel",
        "Fernandes", "Chen", "Mansour", "Silva", "Iyer", "Rossi", "DSouza", "Kumar", "Lee", "Saeed"]
countries = {"UAE": ["Dubai", "Abu Dhabi"], "KSA": ["Riyadh", "Jeddah"], "Qatar": ["Doha"],
             "India": ["Mumbai", "Bengaluru"], "UK": ["London"], "USA": ["New York"]}
country_variants = {"UAE": ["UAE", "U.A.E", "United Arab Emirates", "uae "],
                    "KSA": ["KSA", "Saudi Arabia", "ksa"], "Qatar": ["Qatar", "QATAR"],
                    "India": ["India", "india"], "UK": ["UK", "United Kingdom"], "USA": ["USA", "US"]}
nationalities = ["Indian", "Pakistani", "Filipino", "British", "Egyptian", "Emirati", "American", "Lebanese"]
depts = ["Finance", "Operations", "IT", "HR", "Sales", "Marketing", "Logistics"]
visa_status = ["Valid", "Valid", "Valid", "Under Process", "Cancelled", "Expired"]
visa_variants = {"Valid": ["Valid", "valid", "VALID", "Active"], "Under Process": ["Under Process", "In progress", "processing"],
                 "Cancelled": ["Cancelled", "canceled"], "Expired": ["Expired", "expired"]}
today = date.today()


def fmt_date(d, messy):
    if d is None:
        return None
    if not messy:
        return d
    return random.choice([d.strftime("%d/%m/%Y"), d.strftime("%d-%m-%Y"), d.strftime("%Y-%m-%d"),
                          d.strftime("%d %b %Y"), d.strftime("%d.%m.%Y")])


rows = []
for i in range(1, N + 1):
    c = random.choice(list(countries))
    nat = "Emirati" if c == "UAE" and random.random() < 0.05 else random.choice(nationalities)
    issue = today - timedelta(days=random.randint(200, 3400))
    p_exp = issue + timedelta(days=3652)
    v_exp = today + timedelta(days=random.randint(-120, 900))
    vs = random.choice(visa_status)
    if vs == "Expired":
        v_exp = today - timedelta(days=random.randint(1, 200))
    messy = random.random() < 0.35
    rows.append({
        "Employee ID": f"EMP{i:05d}",
        "Full Name": f"{random.choice(first)} {random.choice(last)}",
        "Department": random.choice(depts),
        "Work Country": random.choice(country_variants[c]),
        "Work Location": random.choice(countries[c]),
        "Nationality": nat,
        "Email": f"emp{i:05d}@company.com",
        "Passport Number": f"{random.choice('ABCKLMNPZ')}{random.randint(1000000, 9999999)}",
        "Passport Issue Date": fmt_date(issue, messy),
        "Passport Expiry Date": fmt_date(p_exp, messy),
        "Visa Status": random.choice(visa_variants[vs]),
        "Visa Expiry Date": fmt_date(v_exp, messy),
        "Line Manager": f"{random.choice(first)} {random.choice(last)}",
    })

df = pd.DataFrame(rows)

# Inject realistic problems
idx = lambda k: random.sample(range(N), k)
for r in idx(15): df.at[r, "Passport Expiry Date"] = df.at[r, "Passport Issue Date"]  # expiry == issue
for r in idx(12): df.at[r, "Passport Expiry Date"] = "31/02/2027"                  # impossible date
for r in idx(10): df.at[r, "Passport Expiry Date"] = None                          # missing
for r in idx(10): df.at[r, "Visa Expiry Date"] = "TBC"                             # text in date
for r in idx(8):  df.at[r, "Passport Number"] = None
for r in idx(10): df.at[r, "Passport Number"] = " " + str(df.at[r, "Passport Number"]).lower()
for r in idx(10): df.at[r, "Full Name"] = "  " + df.at[r, "Full Name"].upper() + " "
for r in idx(6):  df.at[r, "Email"] = df.at[r, "Email"].replace("@", "")           # invalid email
for r in idx(8):  # past expiry date but status says valid
    df.at[r, "Visa Status"] = "Valid"
    df.at[r, "Visa Expiry Date"] = today - timedelta(days=random.randint(5, 90))
for r in idx(5):  # Emirati with visa (nationals don't need one)
    df.at[r, "Nationality"] = "Emirati"; df.at[r, "Work Country"] = "UAE"; df.at[r, "Visa Status"] = "Valid"
dups = df.sample(8, random_state=1)
df = pd.concat([df, dups], ignore_index=True)                                     # duplicate rows
dup_ids = df.sample(4, random_state=2).index
df.loc[dup_ids, "Employee ID"] = df.loc[df.sample(4, random_state=3).index, "Employee ID"].values  # same ID, different person

df.to_excel("hr_sample_messy.xlsx", index=False, sheet_name="Employees")
print(f"Wrote hr_sample_messy.xlsx with {len(df)} rows")
