import pandas as pd
import openpyxl as op

input_file = r"C:\Users\shant\Downloads\Mgi\return\MGI Group Return Dt 31-10-26 to 2-11-26.xlsx"
df = pd.read_excel(input_file)

print("Original rows:", len(df))
print("Original columns:", len(df.columns))

# 1. Clean column names
df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.lower()
    .str.replace(' ', '_')
)

# 2. Remove completely empty columns
df = df.dropna(axis=1, how='all')

# 3. Clean text cells
for column in df.select_dtypes(include=["object", "str"]).columns:
    df[column] = df[column].apply(
        lambda x: x.strip() if isinstance(x, str) else x
    )

df = df.replace(r"^\s*$", pd.NA, regex=True)

# 4. Find row containing the actual table headers
search_word = "Date"

matches = df.apply(
    lambda row: row.astype(str).str.contains(
        search_word, case=False, na=False
    ).any(),
    axis=1
)

if matches.any():
    first_row = matches.idxmax()
    print("Found header at Excel row:", first_row + 1)
    df = df.iloc[first_row:].reset_index(drop=True)
else:
    raise ValueError("Could not find the Date header row.")

# 5. Set the correct column headers
df.columns = df.iloc[0].astype(str).str.strip()
df = df.iloc[1:].reset_index(drop=True)

# Remove empty cells and clean column names
df = df.replace(r"^\s*$", pd.NA, regex=True)
df.columns = df.columns.astype(str).str.strip()

# 6. Make duplicate column names unique
new_columns = []
counts = {}

for col in df.columns:
    counts[col] = counts.get(col, 0) + 1

    if counts[col] == 1:
        new_columns.append(col)
    else:
        new_columns.append(f"{col}_{counts[col]}")

df.columns = new_columns

print("Cleaned columns:", df.columns.tolist())

# 7. Remove unwanted columns
columns_to_drop = [
    "Voucher Type",
    "Voucher Ref. No.",
    "Voucher No.",
    "Order No. & Date",
    "Narration",
    "Other References",
    "Terms of Delivery",
    "LOCAL GST SALE@18%",
    "OUTPUT CGST@9%",
    "OUTPUT SGST@9%",
    "IGST SALE@18%",
    "OUTPUT IGST@18%",
    "Terms of Payment",
    "Gross Total"
]

df = df.drop(columns=columns_to_drop, errors="ignore")

# 8. Identify the Particulars columns
particulars_cols = [
    col for col in df.columns
    if col == "Particulars" or col.startswith("Particulars_")
]

if not particulars_cols:
    raise ValueError("No Particulars column found.")

# 9. Identify order header rows
df["Date"] = df["Date"].replace(r"^\s*$", pd.NA, regex=True)
header_mask = df["Date"].notna()

# Assign an order ID to each header and its following rows
df["Order_ID"] = header_mask.cumsum()

# 10. Save order date and platform details
headers = df.loc[
    header_mask,
    ["Order_ID", "Date", particulars_cols[0]]
].copy()

headers = headers.rename(
    columns={particulars_cols[0]: "Platform"}
)

# 11. Get the SKU detail rows
sku_rows = df.loc[~header_mask].copy()

# Keep rows containing a SKU/particular
sku_rows = sku_rows[
    sku_rows[particulars_cols].notna().any(axis=1)
].copy()

# Extract the last non-empty Particulars value on each row
sku_rows["SKU"] = (
    sku_rows[particulars_cols]
    .ffill(axis=1)
    .iloc[:, -1]
)

# 12. Remove header-only fields from SKU rows
sku_rows = sku_rows.drop(
    columns=["Date"] + particulars_cols,
    errors="ignore"
)

# 13. Attach the correct date and platform to every SKU
result = sku_rows.merge(
    headers,
    on="Order_ID",
    how="left"
)

# 14. Remove temporary Order_ID
result = result.drop(columns=["Order_ID"])

# 15. Move Date, SKU and Platform to the beginning
first_columns = ["Date", "SKU", "Platform"]

other_columns = [
    col for col in result.columns
    if col not in first_columns
]

result = result[first_columns + other_columns]

# 16. Calculate net sale using each SKU's own value
if "Value" in result.columns:
    result["net sale"] = (
        pd.to_numeric(result["Value"], errors="coerce") * 1.18
    )
else:
    print("Warning: Value column not found. Net sale not calculated.")

# 17. Swap the positions of Value and net sale
if "Value" in result.columns and "net sale" in result.columns:
    cols = result.columns.tolist()
    i, j = cols.index("Value"), cols.index("net sale")
    cols[i], cols[j] = cols[j], cols[i]
    result = result[cols]

# 18. Reset index
result = result.reset_index(drop=True)

# 19. Save the cleaned Excel file
output_file = r"C:\Users\shant\Downloads\Mgi\MGI cleaned file.xlsx"

result.to_excel(output_file, index=False)

print(result.head(10))
print("Final rows:", len(result))
print("Saved to:", output_file)