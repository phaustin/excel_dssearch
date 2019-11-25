import context
from pathlib import Path
from openpyxl import load_workbook
import pandas as pd

excel_file = context.dsci_search / Path("draft_evaluation_form.xlsx")
wb = load_workbook(filename=str(excel_file))
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
wb.save(filename=str(outfile))
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


out = df_candidates.apply(assign_reviewer, axis=1)
df_revs = pd.DataFrame.from_records(out)
df_candidates[['rev_2','rev_1']]=df_revs[['rev_2','rev_1']]
# df_candidates["reviewer"] = out
# df_candidates.head()
# groups = df_candidates.groupby("reviewer")
# print(groups)
