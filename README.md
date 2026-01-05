# -*- coding: utf-8 -*-
"""
Created on Tue Dec 30 10:23:35 2025

@author: cslamprineas
"""

import pandas as pd
import numpy as np
import os

# =====================
# Constants
# =====================
Ea = 83144        # J/mol (Activation Energy)
R = 8.314         # J/mol·K

# =====================
# File paths
# =====================
raw_path = input(
    "📂 Please enter the full path to the input Excel file:\n"
).strip()

# Remove surrounding quotes if pasted from Explorer
input_file = raw_path.strip('"').strip("'")

# Validate input file
if not os.path.isfile(input_file):
    raise FileNotFoundError(f"❌ Input file not found:\n{input_file}")

# Output in same directory as input
input_dir = os.path.dirname(input_file)

base_output = os.path.join(input_dir, "MKT_results.xlsx")
output_file = base_output
counter = 1

# Prevent overwrite / Excel lock issues
while os.path.exists(output_file):
    output_file = os.path.join(
        input_dir, f"MKT_results_{counter}.xlsx"
    )
    counter += 1

# =====================
# Load all sheets
# =====================
sheets = pd.read_excel(input_file, sheet_name=None)

results = []

for sheet_name, df in sheets.items():

    # ---------------------
    # Find temperature column automatically
    # ---------------------
    temp_col = None
    for col in df.columns:
        if "temp" in col.lower():
            temp_col = col
            break

    if temp_col is None:
        print(f"⚠️ No temperature column found in sheet: {sheet_name}")
        continue

    # ---------------------
    # Convert to numeric & drop missing
    # ---------------------
    temps_c = pd.to_numeric(df[temp_col], errors="coerce").dropna()

    if temps_c.empty:
        print(f"⚠️ No valid temperature data in sheet: {sheet_name}")
        continue

    # ---------------------
    # Convert °C → K
    # ---------------------
    temps_k = temps_c + 273.15

    # ---------------------
    # MKT calculation
    # ---------------------
    mkt_k = -Ea / (R * np.log(np.mean(np.exp(-Ea / (R * temps_k)))))
    mkt_c = mkt_k - 273.15

    # ---------------------
    # Store results
    # ---------------------
    results.append({
        "Sheet": sheet_name,
        "Temperature column": temp_col,
        "No. of points": len(temps_k),
        "MKT (K)": round(mkt_k, 2),
        "MKT (°C)": round(mkt_c, 2)
    })

# =====================
# Save results
# =====================
results_df = pd.DataFrame(results)
results_df.to_excel(output_file, index=False)

print("\n✅ MKT calculation completed successfully")
print(results_df)
print(f"\n📁 Results saved to:\n{output_file}")

