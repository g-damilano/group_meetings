import pandas as pd
import unicodedata
from pathlib import Path

# Supplier locations (closest to London / UK-based)
SUPPLIER_LOCATIONS = {
    "Sigma-Aldrich": "Gillingham, UK",
    "Merck": "Gillingham, UK",
    "VWR": "Lutterworth, UK",
    "VWR Chemicals": "Lutterworth, UK",
    "Fisher Scientific": "Loughborough, UK",
    "Fisher Chemical": "Loughborough, UK",
    "Scientific Laboratory Supplies": "Nottingham, UK",
    "Fluorochem": "Hadfield, UK",
    "Honeywell Fluka": "Seelze, Germany",
    "Honeywell Chemicals": "Seelze, Germany",
    "TCI Chemicals": "Oxford, UK",
    "TCI America": "Portland, USA",
    "AnalaR": "Loughborough, UK",
    "Acros Organics": "Geel, Belgium",
    "Alfa Aesar": "Heysham, UK",
    "Thermo Scientific": "Loughborough, UK",
    "Thermo Scientific Chemicals": "Loughborough, UK",
    "BioServUK": "Sheffield, UK",
    "Bio-Serv": "Flemington, USA",
}

# ---------- Helper Functions ----------

def clean_text(value):
    """Normalize text and strip odd characters."""
    if pd.isna(value):
        return ""
    # Normalize Unicode and remove non-printable artifacts
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii")
    return text.strip()

def enrich_supplier(name):
    """Add UK-closest location if known."""
    if pd.isna(name):
        return ""
    name = clean_text(name)
    for k, loc in SUPPLIER_LOCATIONS.items():
        if k.lower() in name.lower():
            return f"{k} ({loc})"
    return name

def make_citation(row):
    """Generate clean citation string."""
    chem = clean_text(row["Product name"])
    cas = clean_text(row["CAS"])
    supplier = clean_text(row["Supplier (with location)"])
    supplier_clean = supplier if supplier else clean_text(row["Brand"])
    return f"{chem} ({cas}) was obtained from {supplier_clean}."

# ---------- Main Script ----------

# Input and output paths
input_file = Path("Chemicals_query_17_10_2025.xlsx")
output_file = Path("chemical_inventory_full.csv")

# Read Excel
df = pd.read_excel(input_file)
df.columns = [clean_text(c) for c in df.columns]

# Check required columns
required_cols = ["Product name", "CAS", "Brand", "Product number", "Quantity", "Quantity unit"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing columns in Excel: {missing}")

# Combine quantity and unit
df["Quantity"] = df["Quantity"].astype(str).str.strip() + " " + df["Quantity unit"].astype(str).str.strip()

# Clean and enrich supplier data
df["Brand"] = df["Brand"].apply(clean_text)
df["Supplier (with location)"] = df["Brand"].apply(enrich_supplier)

# Generate citation
df["Citation"] = df.apply(make_citation, axis=1)

# Clean remaining text fields
for col in ["Product name", "CAS", "Product number", "Quantity"]:
    df[col] = df[col].apply(clean_text)

# Prepare final DataFrame
output_df = df.rename(columns={
    "Product name": "Chemical name"
})[[
    "Chemical name",
    "CAS",
    "Supplier (with location)",
    "Product number",
    "Quantity",
    "Citation"
]]

# Export CSV (UTF-8 without BOM, clean printable chars)
output_df.to_csv(output_file, index=False, encoding="utf-8")

print(f"✅ Clean CSV created successfully: {output_file.resolve()}")
print(f"Total records exported: {len(output_df)}")
print(f"Suppliers enriched with location: {output_df['Supplier (with location)'].notna().sum()}")
