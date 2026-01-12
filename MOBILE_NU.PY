import pandas as pd
import random

# ---------- CONFIG ----------
input_file = "193_final.xlsx"
output_file = "final_193.xlsx"
column_name = "MobileNo"
# ----------------------------

def generate_pakistani_number(exclude=None):
    """
    Generate Pakistani mobile number:
    0302XXXXXXX (7 random digits)
    Ensure it's not equal to exclude
    """
    while True:
        number = "0303-" + "".join(str(random.randint(0, 9)) for _ in range(7))
        if exclude is None or number != exclude:
            return number

# Load Excel file
df = pd.read_excel(input_file)

# Validate column
if column_name not in df.columns:
    raise ValueError(f"Column '{column_name}' not found in Excel file")

# Process rows
for i in range(len(df)):
    current_value = df.at[i, column_name]

    # Check for null or empty
    if pd.isna(current_value) or str(current_value).strip() == "":
        upper_value = None
        if i > 0 and not pd.isna(df.at[i - 1, column_name]):
            upper_value = str(df.at[i - 1, column_name]).strip()

        df.at[i, column_name] = generate_pakistani_number(exclude=upper_value)

# Save output
df.to_excel(output_file, index=False)

print("Done. Missing mobile numbers filled with 0302xxxxxxx format.")