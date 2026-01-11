import pandas as pd

# Load main Excel file
df = pd.read_excel("final_merged.xlsx")

# Load IDs from txt (keep order
ids_df = pd.read_csv("id.txt", header=None, names=["UniqueId"])

# Ensure same datatype
ids_df["UniqueId"] = ids_df["UniqueId"].astype(df["UniqueId"].dtype)

# LEFT JOIN (keeps all IDs from txt)
result = ids_df.merge(df, on="UniqueId", how="left")

result["Name"] = (
    result["Name"]
    .str.replace(r"\s+(S/O|D/O|W/O)\s+.*", "", regex=True)
    .str.strip()
)

# Save output
result.to_excel("art.xlsx", index=False)
