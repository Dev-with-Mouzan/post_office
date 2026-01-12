import pandas as pd
import random

# ---------- CONFIG ----------
input_file = "art.xlsx"
output_file = "193_final.xlsx"
column_name = "Name"
# ----------------------------

PAKISTANI_MALE_NAMES = [
    "Muhammad Ali",
    "Ahmed Raza",
    "Hassan Ali",
    "Usman Khan",
    "Abdul Rehman",
    "Bilal Ahmed",
    "Ahsan Javed",
    "Fahad Iqbal",
    "Zeeshan Haider",
    "Imran Shah",
    "Kamran Aslam",
    "Noman Tariq"
]

def generate_name(exclude=None):
    """Generate a Pakistani male name different from exclude"""
    while True:
        name = random.choice(PAKISTANI_MALE_NAMES)
        if name != exclude:
            return name

# Load Excel
df = pd.read_excel(input_file)

# Ensure column exists
if column_name not in df.columns:
    raise ValueError(f"Column '{column_name}' not found")

# Process rows
for i in range(len(df)):
    current = df.at[i, column_name]

    if pd.isna(current) or str(current).strip() == "":
        upper_name = None
        if i > 0:
            upper_name = str(df.at[i - 1, column_name]).strip()

        df.at[i, column_name] = generate_name(exclude=upper_name)

# Save output
df.to_excel(output_file, index=False)

print("Done. Null names replaced successfully.")