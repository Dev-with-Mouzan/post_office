# import pandas as pd

# # ---- CONFIG ----
# input_file = "5770_art.xlsx"
# output_file = "output.xlsx"
# address_column = "Address"   # CHANGE if your column name is different
# target_city = "BUREWALA"
# # ----------------

# # Load Excel
# df = pd.read_excel(input_file)

# # Ensure column exists
# if address_column not in df.columns:
#     raise ValueError(f"Column '{address_column}' not found in Excel file")

# # Iterate from second row (index 1)
# for i in range(1, len(df)):
#     current_address = str(df.at[i, address_column]).strip()
#     previous_address = str(df.at[i - 1, address_column]).strip()

#     # If address does NOT end with BUREWALA
#     if not current_address.upper().endswith(target_city):
#         df.at[i, address_column] = previous_address

# # Save result
# df.to_excel(output_file, index=False)

# print("Processing completed. Output saved to:", output_file)


import pandas as pd

# ---- CONFIG ----
input_file = "final_193.xlsx"
output_file = "Artical_193.xlsx"
address_column = "Address"
target_city = "BUREWALA"
MAX_LENGTH = 30
# ----------------

df = pd.read_excel(input_file)

if address_column not in df.columns:
    raise ValueError(f"Column '{address_column}' not found in Excel file")

for i in range(1, len(df)):
    current_address = str(df.at[i, address_column]).strip()
    upper_address = str(df.at[i - 1, address_column]).strip()

    current_invalid = (
        not current_address.upper().endswith(target_city)
        or len(current_address) > MAX_LENGTH
    )

    upper_valid = len(upper_address) <= MAX_LENGTH

    if current_invalid and upper_valid:
        df.at[i, address_column] = upper_address

df.to_excel(output_file, index=False)

print("Processing completed. Output saved to:", output_file)