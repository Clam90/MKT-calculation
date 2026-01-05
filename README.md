📊 Mean Kinetic Temperature (MKT) Calculator

This repository contains a Python script for calculating Mean Kinetic Temperature (MKT) from temperature data stored in Excel files. It is designed for pharmaceutical, cold-chain, and stability-monitoring applications where accurate thermal exposure assessment is critical.

🔧 Features

Automatically reads all sheets from an Excel file

Detects temperature columns dynamically (based on column name)

Converts temperatures from °C to Kelvin

Calculates Mean Kinetic Temperature (MKT) using the Arrhenius equation

Outputs results to a new Excel file without overwriting existing files

Handles missing or invalid data gracefully

📁 Input

An Excel file (.xlsx) containing one or more sheets

Each sheet must include a column with temperature data (column name containing "temp")

📤 Output

A summary Excel file (MKT_results.xlsx) with:

Sheet name

Temperature column used

Number of data points

MKT in Kelvin and Celsius

🧪 Libraries Used

pandas

numpy

os

▶️ How to Use

Run the script

Paste the full path to your Excel file when prompted

Review the generated MKT_results.xlsx file in the same directory

📌 Typical Use Cases

Cold chain monitoring

Pharmaceutical stability studies

Temperature excursion analysis

Quality and compliance reporting
