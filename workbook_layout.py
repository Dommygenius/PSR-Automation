"""Light PSR graphs workbook — fast to open (~1 MB). No .xlsb."""

TEMPLATE_GRAPHS = "PSR_Graphs_Template.xlsx"
SHEET_DATA = "Raw Data"
SHEET_DASHBOARD = "PSR_Snapshot"

# Same paste cells as your original scripts (hidden Raw Data sheet feeds charts).
RNC_ROW, BSC_ROW = 2, 45
LAC_ROW, LAC_COL = 45, 17
TECH_ROW, TECH_COL = 2, 20
SERVICE_ROW, SERVICE_COL = 2, 27

# PSR vs AVAIL yearly trend (column C = manual AVAIL, never written).
SHEET_PSR_AVAIL = "PSR&AVAIL"
PSR_AVAIL_ROW = 2
PSR_AVAIL_COL_DATE = 1
PSR_AVAIL_COL_PSR = 2
PSR_AVAIL_MAX_ROWS = 365

# Sheets hidden in output so boss sees graphs only.
HIDE_SHEETS = ("Raw Data", "Chart_Data", "LAC BSC_RNC", "LAC PERFROMANCE")
