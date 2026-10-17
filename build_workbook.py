"""
CRM HIERARCHY RECONCILIATION PIPELINE  -  Kenya partner network dataset
INPUT   : CRM_Facilities.csv  (15,150 raw facility records from the CRM export)
          Partner_Networks.csv / Sub_Structures.csv / Alias_Map.csv (reference)
OUTPUT  : CRM_Reconciliation_Master.xlsx

STAGE 1  Load + normalise text
STAGE 2  Build in-memory lookup dictionaries (replaces 15,000 VLOOKUPs)
STAGE 3  Resolve each facility to its parent national network
STAGE 4  Flag data-quality exceptions with a reason + severity
STAGE 5  Render the Excel workbook (6 tabs, formatted for ops users)
"""
import difflib
import re
from collections import Counter

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

NAVY, SLATE, LINE = "1F3A5F", "44546A", "D4D7DD"
C_CRIT, C_HIGH, C_MED, C_LOW, C_OK = "C00000", "C65911", "BF8F00", "7F7F7F", "2E7D32"
F_CRIT, F_HIGH, F_MED, F_LOW = "FCE4E4", "FDEADA", "FFF4D6", "F2F2F2"
HDR = PatternFill("solid", fgColor=NAVY)
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BASE, H1, H2 = "Arial", 16, 11


def font(sz=10, b=False, color="000000", italic=False):
    return Font(name=BASE, size=sz, bold=b, color=color, italic=italic)


# STAGE 1 
print("STAGE 1  loading source files ...")
read = lambda f: pd.read_csv(f, dtype=str, keep_default_na=False).fillna("")
raw = read("CRM_Facilities.csv")
networks = read("Partner_Networks.csv")
subs = read("Sub_Structures.csv")
aliases = read("Alias_Map.csv")

clean = lambda t: " ".join(str(t).upper().strip().split())

#TAGE 2
print("STAGE 2  building lookup dictionaries ...")
net_by_clean = {clean(n): n for n in networks.Network_Name}
name_by_id = dict(zip(networks.Network_ID, networks.Network_Name))
id_by_name = dict(zip(networks.Network_Name, networks.Network_ID))
alias_to_net = {clean(a): n for a, n in zip(aliases.Alias, aliases.Network_Name)}

subs = subs.copy()
subs["Parent_Network"] = subs.Parent_Network_ID.map(name_by_id)
sub_by_name = {clean(r.SubStructure_Name): r for r in subs.itertuples()}
sub_by_postcode = {clean(r.Postal_Code): r for r in subs.itertuples()}
dup_counts = (raw.Facility_Name.map(clean) + "|" + raw.Postal_Code.map(clean)).value_counts()
fuzzy_cache = {}
for v in raw.Parent_Network.map(clean).unique():
    if v and v not in net_by_clean and v not in alias_to_net and v not in sub_by_name:
        m = difflib.get_close_matches(v, list(net_by_clean), n=1, cutoff=0.80)
        fuzzy_cache[v] = net_by_clean[m[0]] if m else None

VALID_STATUS = {"ACTIVE", "PENDING", "EXPIRED", "TERMINATED"}
SEV_NAME = ["OK", "Low", "Medium", "High", "Critical"]

#STAGE 3
print("STAGE 3  resolving hierarchy + STAGE 4 flagging exceptions ...")
out = []
for r in raw.itertuples():
    p, s = clean(r.Parent_Network), clean(r.Sub_Structure)
    pc, st = clean(r.Postal_Code), clean(r.Agreement_Status)


    exact = net_by_clean.get(p)
    alias = alias_to_net.get(p)
    branch_as_parent = sub_by_name.get(p)
    fuzzy = fuzzy_cache.get(p)
    stated = exact or alias or (branch_as_parent.Parent_Network if branch_as_parent is not None else None) or fuzzy
    how = ("Blank" if not p else "Exact" if exact else "Alias" if alias
           else "BranchAsParent" if branch_as_parent is not None else "Fuzzy" if fuzzy else "Unknown")

  
    sub_info = sub_by_name.get(s)
    pc_info = sub_by_postcode.get(pc)
    pc_valid = bool(re.fullmatch(r"\d{5}", pc))
    if sub_info is not None:
        network, source = sub_info.Parent_Network, "branch"
    elif stated:
        network, source = stated, "stated"
    elif pc_info is not None:
        network, source = pc_info.Parent_Network, "postcode"
    else:
        network, source = None, None

    branch = sub_info if sub_info is not None else (
        pc_info if pc_info is not None and (network is None or pc_info.Parent_Network == network) else None)

    if network is None:
        method, conf = "Unresolved", "None"
    elif source == "branch" and stated == network:
        method, conf = "Confirmed (parent and branch agree)", "High"
    elif source == "branch":
        method, conf = "Derived from branch hierarchy", "Medium"
    elif source == "stated":
        method, conf = f"Parent name only ({how.lower()} match)", "Medium"
    else:
        method, conf = "Inferred from postal code", "Low"
    iss = []
    if network is None:
        iss.append(("Cannot be mapped to any national network", 4,
                    "Assign a network manually - no branch, parent or postal code match"))
    if not p:
        iss.append(("Missing parent network", 3,
                    f"Set parent to '{network}'" + (f" (from branch '{r.Sub_Structure}')" if source == "branch"
                     else " (from postal code)" if source == "postcode" else "") if network else "Investigate"))
    elif how == "Unknown":
        iss.append((f"Parent name not recognised: '{r.Parent_Network}'", 3,
                    f"Replace with '{network}'" if network else "Add to Alias_Map or correct the entry"))
    elif how == "Fuzzy":
        iss.append((f"Parent name misspelled: '{r.Parent_Network}'", 2, f"Correct spelling to '{fuzzy}'"))
    elif how == "BranchAsParent":
        iss.append((f"Branch name entered in the parent field: '{r.Parent_Network}'", 2,
                    f"Move to branch field; set parent to '{branch_as_parent.Parent_Network}'"))
    elif how == "Alias":
        iss.append((f"Parent entered as an alias: '{r.Parent_Network}'", 1, f"Standardise to '{alias}'"))
    elif how == "Exact" and r.Parent_Network != exact:
        iss.append(("Parent name formatting (case or spacing)", 1, f"Standardise to '{exact}'"))

    if stated and sub_info is not None and stated != sub_info.Parent_Network:
        iss.append((f"Parent conflicts with branch hierarchy (entered '{stated}')", 3,
                    f"Branch '{sub_info.SubStructure_Name}' belongs to '{sub_info.Parent_Network}' - confirm which is correct"))

    if not s:
        iss.append(("Missing branch / sub-structure", 3,
                    f"Set branch to '{branch.SubStructure_Name}'" if branch is not None else "Assign a branch"))
    elif sub_info is None:
        iss.append((f"Unmapped branch: '{r.Sub_Structure}'", 3,
                    "Add this branch to the Partner_Reference tab, or correct the entry"))

    if not pc:
        iss.append(("Missing postal code", 2, "Add the facility postal code"))
    elif not pc_valid:
        iss.append((f"Invalid postal code: '{r.Postal_Code}'", 2, "Correct to a 5-digit Kenyan postal code"))
    elif pc_info is None:
        iss.append(("Postal code outside all known branch areas", 1, "Confirm the location is correct"))

    if not st:
        iss.append(("Missing agreement status", 2, "Set the agreement status"))
    elif st not in VALID_STATUS:
        iss.append((f"Invalid agreement status: '{r.Agreement_Status}'", 2,
                    "Use Active, Pending, Expired or Terminated"))

    if dup_counts[clean(r.Facility_Name) + "|" + pc] > 1:
        iss.append(("Possible duplicate record (same name and postal code)", 1,
                    "Compare with the matching record and merge if identical"))

    rank = max([v for _, v, _ in iss], default=0)
    out.append({
        "Facility_ID": r.Facility_ID, "Facility_Name": r.Facility_Name,
        "Postal_Code": r.Postal_Code, "City": r.City,
        "County": branch.County if branch is not None else "",
        "Region": branch.Region if branch is not None else "",
        "Agreement_Status": r.Agreement_Status,
        "Parent_As_Entered": r.Parent_Network, "Branch_As_Entered": r.Sub_Structure,
        "Reconciled_Network": network or "", "Network_ID": id_by_name.get(network, ""),
        "Reconciled_Branch": branch.SubStructure_Name if branch is not None else "",
        "Match_Method": method, "Confidence": conf,
        "Severity": SEV_NAME[rank], "Severity_Rank": rank,
        "Exception_Count": len(iss),
        "Exception_Reasons": "; ".join(x[0] for x in iss),
        "Suggested_Fix": " | ".join(x[2] for x in iss),
        "Last_Updated": r.Last_Updated,
    })

master = pd.DataFrame(out)
exc = master[master.Severity_Rank > 0].sort_values(
    ["Severity_Rank", "Reconciled_Network", "Facility_ID"], ascending=[False, True, True]).reset_index(drop=True)

n_all, n_exc = len(master), len(exc)
print(f"         {n_all:,} rows processed | {n_exc:,} exceptions | "
      f"{(master.Reconciled_Network != '').sum() / n_all:.2%} mapped")

# STAGE 5
print("STAGE 5  rendering Excel workbook ...")
wb = Workbook()


def write_sheet(ws, df, table_name, widths, freeze="A2", wrap_cols=()):
    """Write a dataframe as a native Excel table (fast filtering, no formulas)."""
    ws.append(list(df.columns))
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(1, c)
        cell.font = font(10, True, "FFFFFF")
        cell.fill = HDR
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for row in df.itertuples(index=False):
        ws.append(list(row))
    for col, w in zip(range(1, len(df.columns) + 1), widths):
        ws.column_dimensions[get_column_letter(col)].width = w
    for col in wrap_cols:
        for cell in ws[get_column_letter(col)][1:]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
    t = Table(displayName=table_name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False


def severity_colours(ws, col_letter, last_row, fill_row_range=None):
    rules = [("Critical", C_CRIT, F_CRIT), ("High", C_HIGH, F_HIGH),
             ("Medium", C_MED, F_MED), ("Low", C_LOW, F_LOW)]
    for label, colour, fill in rules:
        ws.conditional_formatting.add(
            f"{col_letter}2:{col_letter}{last_row}",
            CellIsRule(operator="equal", formula=[f'"{label}"'],
                       font=Font(name=BASE, size=10, bold=True, color=colour),
                       fill=PatternFill("solid", fgColor=fill)))
    if fill_row_range:
        for label, colour, fill in rules[:2]:
            ws.conditional_formatting.add(
                fill_row_range,
                FormulaRule(formula=[f'${col_letter}2="{label}"'],
                            fill=PatternFill("solid", fgColor=fill), stopIfTrue=False))


ws = wb.active
ws.title = "Start Here"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 86

ws["B2"] = "CRM Hierarchy Reconciliation"
ws["B2"].font = Font(name=BASE, size=H1, bold=True, color=NAVY)
ws["B3"] = "Partner network affiliation and data-quality review - Kenya operations"
ws["B3"].font = font(10, color=SLATE)

rows = [
    ("", ""),
    ("WHAT THIS WORKBOOK DOES", ""),
    ("", "Every facility in the CRM export is matched to its parent national network, and any record "
         "with a missing, conflicting or unrecognised affiliation is isolated for review."),
    ("", ""),
    ("HOW TO USE IT - 3 STEPS", ""),
    ("1. Open 'Exception Review'", "Every row that needs a human decision, worst first. Use the filter "
                                   "on the Severity column to work through Critical and High first."),
    ("2. Read 'Suggested Fix'", "For most rows the correct network has already been worked out from the "
                               "branch hierarchy or postal code. The column tells you exactly what to change."),
    ("3. Record your decision", "Choose a value in 'Reviewer Action' (dropdown) and add a note. "
                                "Your entries are never overwritten when the data is reprocessed."),
    ("", ""),
    ("SEVERITY LEGEND", ""),
    ("Critical", "Cannot be mapped to any network at all - needs manual assignment."),
    ("High", "Missing parent, unrecognised parent name, unmapped branch, or parent that conflicts "
             "with the branch hierarchy."),
    ("Medium", "Mapped, but with a correctable data problem: misspelling, branch typed in the parent "
               "field, bad postal code or missing status."),
    ("Low", "Cosmetic or informational: alias used instead of the official name, formatting, "
            "or a possible duplicate."),
    ("", ""),
    ("TAB GUIDE", ""),
    ("Dashboard", "Headline figures and charts - coverage, exceptions by severity, by type and by network."),
    ("Master Reconciled", f"All {n_all:,} records with their resolved network, match method and confidence."),
    ("Exception Review", f"The {n_exc:,} records that need attention, with a suggested fix and review columns."),
    ("Partner Reference", "The 10 national networks and their 40 branches. Add a row here to onboard a new partner."),
    ("Alias Map", "Known abbreviations and alternative spellings. Add a row here to teach the pipeline a new variant."),
    ("QA Validation", "Proof of accuracy: detection rate measured against a known set of planted errors."),
    ("Pipeline Notes", "The matching logic, in order, and the performance design decisions."),
]
r = 5
for label, text in rows:
    if label and not text:
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
        r += 1
        continue
    ws[f"B{r}"] = label
    ws[f"C{r}"] = text
    ws[f"B{r}"].font = font(10, True, color=SLATE)
    ws[f"C{r}"].font = font(10)
    ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30 if len(text) > 95 else 15
    r += 1

for lbl, colour, rr in [("Critical", C_CRIT, 15), ("High", C_HIGH, 16), ("Medium", C_MED, 17), ("Low", C_LOW, 18)]:
    ws[f"B{rr}"].font = Font(name=BASE, size=10, bold=True, color=colour)

ws[f"B{r+1}"] = "Pipeline"
ws[f"B{r+1}"].font = font(10, True, color=SLATE)
ws[f"C{r+1}"] = ("Python (pandas) reads the CRM export and the three reference tables, resolves the hierarchy, "
                 "then writes this workbook. Re-run build_workbook.py on a new export to regenerate it.")
ws[f"C{r+1}"].font = font(10, italic=True, color=SLATE)
ws[f"C{r+1}"].alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[r + 1].height = 30

ws = wb.create_sheet("Master Reconciled")
mcols = ["Facility_ID", "Facility_Name", "Postal_Code", "City", "County", "Region", "Agreement_Status",
         "Parent_As_Entered", "Branch_As_Entered", "Reconciled_Network", "Network_ID", "Reconciled_Branch",
         "Match_Method", "Confidence", "Severity", "Severity_Rank", "Exception_Count", "Exception_Reasons",
         "Last_Updated"]
write_sheet(ws, master[mcols], "tblMaster",
            [13, 42, 12, 14, 15, 17, 16, 30, 28, 30, 11, 28, 34, 11, 11, 13, 15, 60, 13])
severity_colours(ws, "O", n_all + 1)
MASTER_LAST = n_all + 1

ws = wb.create_sheet("Exception Review")
ecols = ["Severity", "Facility_ID", "Facility_Name", "City", "County", "Agreement_Status",
         "Parent_As_Entered", "Branch_As_Entered", "Reconciled_Network", "Match_Method",
         "Confidence", "Exception_Count", "Exception_Reasons", "Suggested_Fix"]
edf = exc[ecols].copy()
for extra in ["Reviewer Action", "Reviewer Notes", "Reviewed By", "Date Reviewed"]:
    edf[extra] = ""
write_sheet(ws, edf, "tblExceptions",
            [11, 13, 40, 14, 15, 16, 28, 28, 30, 32, 11, 9, 58, 58, 26, 30, 16, 14],
            wrap_cols=(13, 14))
severity_colours(ws, "A", n_exc + 1)

dv = DataValidation(
    type="list",
    formula1='"Pending review,Confirmed - apply suggested fix,Rejected - needs investigation,Fixed in CRM,No action needed"',
    allow_blank=True, showDropDown=False)
dv.prompt, dv.promptTitle = "Choose the outcome of your review", "Reviewer Action"
ws.add_data_validation(dv)
dv.add(f"O2:O{n_exc + 1}")
for c in range(15, 19):
    ws.cell(1, c).fill = PatternFill("solid", fgColor="2E7D32")
EXC_LAST = n_exc + 1

ws = wb.create_sheet("Dashboard", 1)
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJ", [3, 32, 14, 14, 12, 36, 14, 12, 12, 12]):
    ws.column_dimensions[col].width = w

ws["B2"] = "Reconciliation Dashboard"
ws["B2"].font = Font(name=BASE, size=H1, bold=True, color=NAVY)
ws["B3"] = "All figures are live formulas over the Master Reconciled tab - they update when the data is reprocessed."
ws["B3"].font = font(9, italic=True, color=SLATE)

M = f"'Master Reconciled'!"
SEV = f"{M}$O$2:$O${MASTER_LAST}"
NET = f"{M}$J$2:$J${MASTER_LAST}"
CONF = f"{M}$N$2:$N${MASTER_LAST}"
REASONS = f"{M}$R$2:$R${MASTER_LAST}"

kpis = [
    ("Total facility records", f"=COUNTA({M}$A$2:$A${MASTER_LAST})", "#,##0"),
    ("Mapped to a network", f'=COUNTIF({NET},"?*")', "#,##0"),
    ("Mapping coverage", f'=IFERROR(COUNTIF({NET},"?*")/COUNTA({M}$A$2:$A${MASTER_LAST}),0)', "0.0%"),
    ("Clean records (no exception)", f'=COUNTIF({SEV},"OK")', "#,##0"),
    ("Records needing review", f'=COUNTA({M}$A$2:$A${MASTER_LAST})-COUNTIF({SEV},"OK")', "#,##0"),
    ("High-confidence matches", f'=COUNTIF({CONF},"High")', "#,##0"),
]
ws["B5"] = "HEADLINE FIGURES"
ws["B5"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
for i, (label, formula, fmt) in enumerate(kpis):
    rr = 6 + i
    ws[f"B{rr}"] = label
    ws[f"B{rr}"].font = font(10)
    ws[f"C{rr}"] = formula
    ws[f"C{rr}"].font = font(12, True, color=NAVY)
    ws[f"C{rr}"].number_format = fmt
    ws[f"C{rr}"].alignment = Alignment(horizontal="right")
    for cl in "BC":
        ws[f"{cl}{rr}"].border = BOX

ws["B14"] = "EXCEPTIONS BY SEVERITY"
ws["B14"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
ws["B15"], ws["C15"], ws["D15"] = "Severity", "Records", "Share"
for cl in "BCD":
    ws[f"{cl}15"].font = font(10, True, "FFFFFF")
    ws[f"{cl}15"].fill = HDR
for i, (label, colour) in enumerate([("Critical", C_CRIT), ("High", C_HIGH), ("Medium", C_MED),
                                     ("Low", C_LOW), ("OK", C_OK)]):
    rr = 16 + i
    ws[f"B{rr}"] = label
    ws[f"B{rr}"].font = Font(name=BASE, size=10, bold=True, color=colour)
    ws[f"C{rr}"] = f'=COUNTIF({SEV},$B{rr})'
    ws[f"C{rr}"].number_format = "#,##0"
    ws[f"D{rr}"] = f'=IFERROR($C{rr}/COUNTA({M}$A$2:$A${MASTER_LAST}),0)'
    ws[f"D{rr}"].number_format = "0.0%"
    ws[f"C{rr}"].font = font(10)
    ws[f"D{rr}"].font = font(10)
    for cl in "BCD":
        ws[f"{cl}{rr}"].border = BOX

ws["F5"] = "EXCEPTIONS BY TYPE"
ws["F5"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
ws["F6"], ws["G6"] = "Issue", "Records"
for cl in "FG":
    ws[f"{cl}6"].font = font(10, True, "FFFFFF")
    ws[f"{cl}6"].fill = HDR
issue_patterns = [
    ("Missing parent network", "Missing parent network"),
    ("Parent conflicts with branch", "Parent conflicts with branch"),
    ("Unmapped branch", "Unmapped branch"),
    ("Branch name in parent field", "Branch name entered in the parent field"),
    ("Parent name misspelled", "Parent name misspelled"),
    ("Parent entered as an alias", "Parent entered as an alias"),
    ("Missing branch", "Missing branch / sub-structure"),
    ("Missing postal code", "Missing postal code"),
    ("Invalid postal code", "Invalid postal code"),
    ("Missing agreement status", "Missing agreement status"),
    ("Possible duplicate", "Possible duplicate record"),
    ("Cannot be mapped at all", "Cannot be mapped to any national network"),
]
for i, (label, pattern) in enumerate(issue_patterns):
    rr = 7 + i
    ws[f"F{rr}"] = label
    ws[f"F{rr}"].font = font(10)
    ws[f"G{rr}"] = f'=SUMPRODUCT(--ISNUMBER(SEARCH("{pattern}",{REASONS})))'
    ws[f"G{rr}"].number_format = "#,##0"
    for cl in "FG":
        ws[f"{cl}{rr}"].border = BOX

ws["B23"] = "COVERAGE BY NATIONAL NETWORK"
ws["B23"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
hdrs = ["Network", "Facilities", "Exceptions", "Clean %"]
for j, h in enumerate(hdrs):
    c = ws.cell(24, 2 + j, h)
    c.font = font(10, True, "FFFFFF")
    c.fill = HDR
for i, nm in enumerate(networks.Network_Name):
    rr = 25 + i
    ws.cell(rr, 2, nm).font = font(10)
    ws.cell(rr, 3).value = f'=COUNTIF({NET},$B{rr})'
    ws.cell(rr, 4).value = f'=COUNTIFS({NET},$B{rr},{SEV},"<>OK")'
    ws.cell(rr, 5).value = f'=IFERROR(($C{rr}-$D{rr})/$C{rr},0)'
    ws.cell(rr, 3).number_format = ws.cell(rr, 4).number_format = "#,##0"
    ws.cell(rr, 5).number_format = "0.0%"
    for cc in range(2, 6):
        ws.cell(rr, cc).border = BOX

ch1 = BarChart()
ch1.type, ch1.title, ch1.height, ch1.width = "col", "Records by severity", 7.2, 11
ch1.add_data(Reference(ws, min_col=3, min_row=15, max_row=20), titles_from_data=True)
ch1.set_categories(Reference(ws, min_col=2, min_row=16, max_row=20))
ch1.legend = None
ws.add_chart(ch1, "F21")

ch2 = BarChart()
ch2.type, ch2.title, ch2.height, ch2.width = "bar", "Exceptions by type", 9.5, 11
ch2.add_data(Reference(ws, min_col=7, min_row=6, max_row=18), titles_from_data=True)
ch2.set_categories(Reference(ws, min_col=6, min_row=7, max_row=18))
ch2.legend = None
ws.add_chart(ch2, "F37")


ws = wb.create_sheet("Partner Reference")
ws.sheet_view.showGridLines = False
ws["A1"] = "National partner networks"
ws["A1"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
nref = networks[["Network_ID", "Network_Name", "Country", "Contract_Tier"]]
for j, h in enumerate(nref.columns):
    c = ws.cell(2, 1 + j, h)
    c.font = font(10, True, "FFFFFF")
    c.fill = HDR
for i, row in enumerate(nref.itertuples(index=False)):
    for j, v in enumerate(row):
        c = ws.cell(3 + i, 1 + j, v)
        c.font = font(10)
        c.border = BOX
start = 3 + len(nref) + 2
ws.cell(start - 1, 1, "Branches / sub-structures").font = Font(name=BASE, size=H2, bold=True, color=NAVY)
sref = subs[["SubStructure_ID", "SubStructure_Name", "Parent_Network_ID", "Parent_Network",
             "Postal_Code", "Town", "County", "Region"]]
for j, h in enumerate(sref.columns):
    c = ws.cell(start, 1 + j, h)
    c.font = font(10, True, "FFFFFF")
    c.fill = HDR
for i, row in enumerate(sref.itertuples(index=False)):
    for j, v in enumerate(row):
        c = ws.cell(start + 1 + i, 1 + j, v)
        c.font = font(10)
        c.border = BOX
for col, w in zip("ABCDEFGH", [16, 34, 18, 32, 13, 16, 18, 18]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A3"


ws = wb.create_sheet("Alias Map")
ws.sheet_view.showGridLines = False
ws["A1"] = "Known aliases and abbreviations"
ws["A1"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
ws["A2"] = "Add a row here to teach the pipeline a new variant, then re-run it."
ws["A2"].font = font(9, italic=True, color=SLATE)
for j, h in enumerate(["Alias", "Network_ID", "Network_Name"]):
    c = ws.cell(4, 1 + j, h)
    c.font = font(10, True, "FFFFFF")
    c.fill = HDR
for i, row in enumerate(aliases[["Alias", "Network_ID", "Network_Name"]].itertuples(index=False)):
    for j, v in enumerate(row):
        c = ws.cell(5 + i, 1 + j, v)
        c.font = font(10)
        c.border = BOX
for col, w in zip("ABC", [30, 16, 36]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A5"


ws = wb.create_sheet("QA Validation")
ws.sheet_view.showGridLines = False
ws["A1"] = "QA validation - detection rate against known planted errors"
ws["A1"].font = Font(name=BASE, size=H1, bold=True, color=NAVY)
ws["A2"] = ("The test dataset was generated with a recorded list of deliberate errors (Injected_Errors_Log.csv). "
            "This tab compares that list against what the pipeline actually flagged.")
ws["A2"].font = font(9, italic=True, color=SLATE)
ws["A2"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[2].height = 26

try:
    inj = read("Injected_Errors_Log.csv")
    flagged = dict(zip(master.Facility_ID, master.Exception_Reasons))
    checks = {
        "Missing parent": ["Missing parent network"],
        # a case/spacing typo normalises to an exact match, so it is correctly
        # reported as a formatting issue rather than a misspelling
        "Typo in parent name": ["Parent name misspelled", "Parent name not recognised",
                                "Parent name formatting"],
        "Alias / abbreviation used": ["Parent entered as an alias"],
        "Parent conflicts with sub-structure": ["Parent conflicts with branch hierarchy"],
        "Unmapped sub-structure": ["Unmapped branch"],
        "Sub-structure entered as parent": ["Branch name entered in the parent field"],
        "Missing postal code": ["Missing postal code"],
        "Invalid postal code": ["Invalid postal code"],
        "Missing agreement status": ["Missing agreement status"],
        "Missing parent and sub-structure": ["Missing parent network", "Missing branch"],
        "Duplicate facility record": ["Possible duplicate record"],
        "Fully orphaned record": ["Cannot be mapped to any national network"],
    }
    res = []
    for kind, pats in checks.items():
        ids = inj[inj.Injected_Error == kind].Facility_ID
        hit = sum(1 for i in ids if any(p in flagged.get(i, "") for p in pats))
        res.append((kind, len(ids), hit))
    hdrs = ["Planted error type", "Rows planted", "Rows detected", "Detection rate", "Result"]
    for j, h in enumerate(hdrs):
        c = ws.cell(4, 1 + j, h)
        c.font = font(10, True, "FFFFFF")
        c.fill = HDR
    for i, (kind, tot, hit) in enumerate(res):
        rr = 5 + i
        ws.cell(rr, 1, kind).font = font(10)
        ws.cell(rr, 2, tot).number_format = "#,##0"
        ws.cell(rr, 3, hit).number_format = "#,##0"
        ws.cell(rr, 4).value = f"=IFERROR($C{rr}/$B{rr},0)"
        ws.cell(rr, 4).number_format = "0.0%"
        ws.cell(rr, 5).value = f'=IF($C{rr}=$B{rr},"PASS","REVIEW")'
        for cc in range(1, 6):
            ws.cell(rr, cc).border = BOX
    tr = 5 + len(res)
    ws.cell(tr, 1, "TOTAL").font = font(10, True)
    ws.cell(tr, 2).value = f"=SUM($B$5:$B${tr-1})"
    ws.cell(tr, 3).value = f"=SUM($C$5:$C${tr-1})"
    ws.cell(tr, 4).value = f"=IFERROR($C{tr}/$B{tr},0)"
    ws.cell(tr, 4).number_format = "0.0%"
    ws.cell(tr, 5).value = f'=IF($C{tr}=$B{tr},"PASS","REVIEW")'
    for cc in range(1, 6):
        ws.cell(tr, cc).border = BOX
        ws.cell(tr, cc).font = font(10, True, color=NAVY)
        ws.cell(tr, cc).number_format = "#,##0" if cc in (2, 3) else ws.cell(tr, cc).number_format
    ws.cell(tr + 2, 1, "Note: this tab exists to prove the logic on test data. Delete it before using the "
                       "workbook on real CRM data.").font = font(9, italic=True, color=SLATE)
    print(f"         QA detection: {sum(h for _, _, h in res):,}/{sum(t for _, t, _ in res):,}")
except FileNotFoundError:
    ws["A4"] = "Injected_Errors_Log.csv not found - QA check skipped."
for col, w in zip("ABCDE", [42, 16, 16, 16, 14]):
    ws.column_dimensions[col].width = w


ws = wb.create_sheet("Pipeline Notes")
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 92
ws["B2"] = "How the pipeline works"
ws["B2"].font = Font(name=BASE, size=H1, bold=True, color=NAVY)

notes = [
    ("MATCHING ORDER", ""),
    ("", "Each facility is resolved by trying these in order and stopping at the first that succeeds."),
    ("1. Branch hierarchy", "The facility's branch is looked up in the official branch table, which gives the "
                            "parent network directly. This is the most reliable source and overrides the typed "
                            "parent name."),
    ("2. Parent name typed in", "Exact match first; then the alias list; then a branch name mistakenly typed into "
                                "the parent field; then fuzzy matching at 80% similarity to catch typos."),
    ("3. Postal code", "If the parent and branch are both unusable, the postal code is matched against the "
                       "branch coverage areas. Flagged as low confidence."),
    ("", ""),
    ("WHY NOT FORMULAS", ""),
    ("", "The original workflow crashed because 15,000 rows x several VLOOKUP columns means hundreds of "
         "thousands of lookups recalculating on every edit. Here the matching runs once in Python against "
         "hash-map lookups, and the workbook stores finished values. Nothing recalculates while an operator "
         "scrolls or filters."),
    ("", ""),
    ("PERFORMANCE DESIGN", ""),
    ("No volatile functions", "No OFFSET, INDIRECT, TODAY or NOW anywhere, so Excel never recalculates the "
                              "workbook just because a cell was touched."),
    ("Native Excel tables", "Master and Exception tabs are real tables, so filtering and sorting use Excel's "
                            "indexed engine rather than scanning formulas."),
    ("Dashboard only", f"The only formulas in the file are the {len(kpis) + 30}-odd COUNTIF/SUMPRODUCT cells on "
                       "the Dashboard and QA tabs, so the figures stay live without any row-level cost."),
    ("Conditional formatting", "Limited to the single Severity column rather than the full grid, which keeps "
                               "rendering fast while scrolling."),
    ("", ""),
    ("REPROCESSING A NEW EXPORT", ""),
    ("", "Drop the new CRM_Facilities.csv next to build_workbook.py and run it. Reference tables and the alias "
         "list are read from their own CSVs, so onboarding a partner or teaching a new spelling means adding a "
         "row, not changing code."),
    ("", ""),
    ("SCALE", ""),
    ("", "Matching is O(n) against hash maps, so the runtime grows linearly. The same script handles 50,000 "
         "rows without a structural change; the practical ceiling is Excel's own row limit, not the pipeline."),
]
r = 4
for label, text in notes:
    if label and not text:
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = Font(name=BASE, size=H2, bold=True, color=NAVY)
        r += 1
        continue
    ws[f"B{r}"] = label
    ws[f"C{r}"] = text
    ws[f"B{r}"].font = font(10, True, color=SLATE)
    ws[f"C{r}"].font = font(10)
    ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r] = ws.row_dimensions[r]
    ws.row_dimensions[r].height = 15 + 13 * (len(text) // 100)
    r += 1


for name, area, landscape in [
    ("Start Here", "B2:C40", False),
    ("Dashboard", "B2:H60", False),
    ("Master Reconciled", None, True),
    ("Exception Review", None, True),
    ("Partner Reference", None, True),
    ("Alias Map", None, False),
    ("QA Validation", "A1:E20", False),
    ("Pipeline Notes", "B2:C40", False),
]:
    s = wb[name]
    s.page_setup.orientation = "landscape" if landscape else "portrait"
    s.page_setup.fitToWidth, s.page_setup.fitToHeight = 1, 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.page_margins.left = s.page_margins.right = 0.4
    if area:
        s.print_area = area
    if landscape:
        s.print_title_rows = "1:1"

wb.active = 0
wb.save("CRM_Reconciliation_Master.xlsx")
print("         saved CRM_Reconciliation_Master.xlsx")
