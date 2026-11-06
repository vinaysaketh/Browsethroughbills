# Electricity Bill HTML Extractor

A Windows desktop utility that reads `.html` / `.htm` electricity bills in a selected folder and exports a consolidated Excel workbook.

## Fields extracted
- CMD (kVA)
- RMD (kVA)
- kWh
- kVAh
- Power Factor (PF)
- Total Electricity Bill (INR)

The workbook also includes the source filename and an extraction status column.

## Build the Windows EXE using GitHub Actions

1. Create a GitHub repository.
2. Upload these files, preserving the folder structure:
   - `electricity_bill_extractor.py`
   - `requirements.txt`
   - `.github/workflows/build-exe.yml`
3. Push to the `main` branch, or open **Actions → Build Windows EXE → Run workflow**.
4. When the workflow finishes, open its run page.
5. Under **Artifacts**, download `ElectricityBillExtractor-Windows`.
6. Extract the ZIP to get `ElectricityBillExtractor.exe`.

## Run the EXE

1. Double-click `ElectricityBillExtractor.exe`.
2. Select the folder containing the HTML bills.
3. Choose where to save the Excel workbook.
4. Review rows marked `Verify Missing Values`.

## Important validation note

Electricity bill HTML layouts and labels vary by DISCOM and bill-generation portal. The script uses generic label matching and is a starting point, not a guarantee of correct extraction. Validate it against a few representative bills before relying on the workbook for analysis. In particular, confirm that the selected total is the intended bill total (not arrears, taxes, or an amount after adjustments).

The EXE is built for Windows x64 by GitHub Actions.
