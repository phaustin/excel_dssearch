import context
from pathlib import Path
from openpyxl import load_workbook
import pandas as pd
import copy


excel_orig = context.dsci_search / Path("draft_evaluation_form.xlsx")
print(f"reading {str(excel_file)}")

outfile = Path() / Path("test.xlsx")
candidate_file = context.dsci_search / Path("EDS - applicant list.xlsx")
print(f'reading "{str(candidate_file)}"')
df_candidates = pd.read_excel(str(candidate_file), skiprows=[0])
columns = [
    "Applicant Name",
    "School (PhD)",
    "Year (PhD)",
    "Highest Education Level",
    "Pawlowicz",
    "Ameli",
    "Austin",
    "Haber",
    "Waterman",
]
df_candidates.fillna(0, inplace=True)
df_candidates = pd.DataFrame(df_candidates[columns], copy=True)
print(df_candidates.index.values)

names = ["Pawlowicz", "Ameli", "Austin", "Haber", "Waterman"]
initials = ["rp", "aa", "pa", "eh", "sw"]
initial_dict = dict(zip(names, initials))
reviewer_text = {1: "rev_1", 2: "rev_2"}


def assign_reviewer(row):
    reviewer_dict = {}
    for reviewer in initial_dict.keys():
        if row[reviewer] > 0:
            reviewer_val = int(row[reviewer])
            reviewer_head = reviewer_text[reviewer_val]
            reviewer_dict[reviewer_head] = initial_dict[reviewer]
    return reviewer_dict


def make_filename(row):
    the_name = row["Applicant Name"]
    the_name = the_name.replace(" ", "_")
    the_name = the_name.replace(",", "_")
    rev2_name = f'{row["rev_2"]}/{the_name}-2_{row["rev_2"]}.xlsx'
    rev1_name = f'{row["rev_1"]}/{the_name}-1_{row["rev_1"]}.xlsx'
    reviewer_dict = dict(rev1_file=rev1_name, rev2_file=rev2_name)
    return reviewer_dict


out = df_candidates.apply(assign_reviewer, axis=1)
df_revs = pd.DataFrame.from_records(out)
df_candidates[["rev_2", "rev_1"]] = df_revs[["rev_2", "rev_1"]]
out = df_candidates.apply(make_filename, axis=1)
df_revs = pd.DataFrame.from_records(out)
df_candidates[["rev2_file", "rev1_file"]] = df_revs[["rev2_file", "rev1_file"]]


def make_xlfile(filename):
    the_file = Path(filename)
    rev_dir = the_file.parent
    out_dir = base_dir / rev_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    outfile = out_dir / the_file.name
    return outfile


fields = ["id", "initials", "name", "school", "year", "level"]
colnames = ["Applicant Name", "School (PhD)", "Year (PhD)", "Highest Education Level"]
file_dict = dict(zip(fields, colnames))
rev_vals = {
    1: {"rev": "rev_1", "filename": "rev1_file"},
    2: {"rev": "rev_2", "filename": "rev2_file"},
}

item_dict = {}
wb_copy = load_workbook(filename=str(excel_orig))
for item in fields:
    the_range = wb_copy.defined_names[item]
    row_col = list(the_range.destinations)[0][1]
    item_dict[item] = row_col


def fill_blanks(sheet, row, reviewer_num):
    file_col = rev_vals[reviewer_num]["rev"]
    initials = row[file_col]
    cellnum = item_dict["initials"]
    print("start")
    sheet[cellnum] = initials
    return sheet


base_dir = Path() / "sheets"
all_rows = list(df_candidates.iterrows())
for item, row in all_rows[:10]:
    for key in [1]:
        wb_copy = load_workbook(filename=str(excel_orig))
        the_sheet = wb_copy["main"]
        filled_sheet = fill_blanks(the_sheet, row, key)
        name_col = rev_vals[key]["filename"]
        filename = row[name_col]
        outfile = make_xlfile(filename)
        wb_copy.save(outfile)
# df_candidates["reviewer"] = out
# df_candidates.head()
# groups = df_candidates.groupby("reviewer")
# print(groups)
