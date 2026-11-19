import json, glob
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = r"C:\Users\pritam.b.patil\Desktop\User Story Agent"
countries = {}
for f in glob.glob(ROOT + r"\agent\cache\scope_items\[0-9A-Z]*.json"):
    for x in json.load(open(f, encoding="utf-8")).get("documents", []):
        if x.get("country_ID") and x["country_ID"] != "XX":
            countries[x["country_ID"]] = x["countryName"]

wb = Workbook()
H = PatternFill("solid", fgColor="1F4E78"); HF = Font(bold=True, color="FFFFFF")
def sheet(title, headers, rows=(), widths=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in ws[1]:
        c.fill = H; c.font = HF; c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows: ws.append(r)
    for i, w in enumerate(widths or [24] * len(headers), 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"
    return ws

rm = wb.active; rm.title = "Read Me"
for line in [
    "SAP S/4HANA Public Cloud - User Story Agent: input template",
    "",
    "Scope_Items: one row per SAP scope item (e.g. 1LQ). Required. 'Country override' is optional: comma-separated country codes that replace the Countries sheet for that item.",
    "Countries: one or several countries for this project. Required. Use the 2-letter code (DE, US...). The name is filled in for you.",
    "L4 plan: upload the L4 plan file separately (same format as L4_Plan_Rev1_Full).",
    "KDD: provisional layout - will be aligned with your real KDD register once you share a sample.",
    "Product_Backlog: provisional layout - will be aligned later.",
    "Owners are intentionally not captured: the agent leaves Owner blank.",
    "SAP release: always the latest published in Process Navigator.",
]: rm.append([line])
rm["A1"].font = Font(bold=True, size=14); rm.column_dimensions["A"].width = 140

sheet("Scope_Items", ["Scope Item ID", "Country override (optional, comma-separated)", "Notes"], [["1LQ", "", "example - delete"]], [16, 40, 40])
cs = sheet("Countries", ["Country Code", "Country Name (auto)", "Company code / entity (optional)"],
           [["DE", None, ""], ["US", None, ""]], [16, 30, 34])
for r in range(2, 60):
    cs.cell(r, 2).value = f'=IFERROR(VLOOKUP(A{r},Lookup_Countries!$A:$B,2,FALSE),"")' if r > 1 else None
sheet("KDD", ["KDD ID", "Title", "Scope Item ID", "Decision", "Rationale", "Status"], [["KDD-001", "example - delete", "1LQ", "", "", "Open"]], [12, 36, 14, 50, 40, 12])
sheet("Product_Backlog", ["Backlog ID", "Title", "Description", "Scope Item ID", "Priority"], [], [14, 36, 50, 14, 10])
lk = sheet("Lookup_Countries", ["Code", "Name"], sorted(countries.items()), [10, 30])
dv = DataValidation(type="list", formula1=f"=Lookup_Countries!$A$2:$A${len(countries)+1}", allow_blank=True)
cs.add_data_validation(dv); dv.add("A2:A59")
out = ROOT + r"\Agent_Input_Template.xlsx"
wb.save(out); print(out, len(countries), "countries")
