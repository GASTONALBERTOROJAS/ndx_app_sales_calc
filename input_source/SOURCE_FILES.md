# Input Source Files

## Location of Source Excel File

The Excel file used by this project is located in the Nordex SharePoint:

### Path
```
Documents > General > Transfer > Patxi > Sales Calc > 26.1
```

### SharePoint Link
```
https://nordex.sharepoint.com/:f:/r/sites/GlobalSourcingControlling/Shared%20Documents/General/Transfer/Patxi/Sales%20Calc/26.1?csf=1&web=1&e=NRDbaP
```

### File Name
```
20260327_Input Sheet_Sales Calculation Tower_v26.1_USD_for2026_2027_2028- Rev01.xlsx
```

---

## Download Instructions

1. Click the SharePoint link above
2. Locate the Excel file: `20260327_Input Sheet_Sales Calculation Tower_v26.1_USD_for2026_2027_2028- Rev01.xlsx`
3. Download the file
4. Update the path in `config.json` if you place it in a different location

---

## Configuration

Once you have the file, update `config.json` in the project root:

```json
{
  "input_file_path": "C:\\path\\to\\your\\file\\20260327_Input Sheet_Sales Calculation Tower_v26.1_USD_for2026_2027_2028- Rev01.xlsx",
  ...
}
```

Replace `C:\\path\\to\\your\\file\\` with the actual path where you saved the Excel file.

---

## Sheet Names Used

The project extracts data from **two sheets** in this Excel file:

1. **TS SC v26.1** — Steel tower components (400 towers)
2. **TCS_MB SC v26.1** — Concrete/hybrid tower components (108 towers)

Both sheets are automatically processed by the extraction script.

---

## Notes

- The file is regularly updated by the Sales Calc team
- Always use the latest version from SharePoint
- If you receive an updated file, simply replace it and run the extraction script again
- The extraction script will automatically read the new data

---

## Last Updated

April 22, 2026
