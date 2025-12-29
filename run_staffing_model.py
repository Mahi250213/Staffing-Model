import math
import pandas as pd
import argparse
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter


def run_staffing_model(input_excel, output_excel):

    def is_empty_day(df, d):
        required = ["total_rooms", "trainees", "crnas", "faculty"]
        for r in required:
            if r not in df.index:
                return True
            if pd.isna(df.loc[r, d]):
                return True
        return False

    def round_half_up(x):
        if x == "" or x is None:
            return ""
        return math.floor(x + 0.5)

    def round_1_decimal(x):
        if x == "" or x is None:
            return ""
        return round(x, 1)

    df = pd.read_excel(input_excel, index_col=0)
    df.index = df.index.str.strip().str.lower()
    dates = list(df.columns)

    yellow = PatternFill("solid", fgColor="FFF200")
    blue = PatternFill("solid", fgColor="CFEAF7")
    green = PatternFill("solid", fgColor="E3F4D7")
    bold = Font(bold=True)

    out_wb = load_workbook(input_excel)
    ws = out_wb.active

    ws.delete_rows(1, ws.max_row)
    ws.delete_cols(1, ws.max_column)

    col_positions = {}
    col = 3
    for i, d in enumerate(dates):
        col_positions[d] = col
        col += 1
        if (i + 1) % 5 == 0:
            col += 1

    for d, c in col_positions.items():
        ws.cell(row=1, column=c).value = d.strftime("%d-%b")
        ws.cell(row=1, column=c).font = bold

    ws["B2"] = "Main/NL/ASC"
    ws["B2"].font = bold

    for d, c in col_positions.items():
        ws.cell(row=2, column=c).value = d.strftime("%a")
        ws.cell(row=2, column=c).font = bold

    row = 3

    def write_row(label, values, fill=None, bold_row=False):
        nonlocal row
        header_cell = ws[f"B{row}"]
        header_cell.value = label
        if bold_row:
            header_cell.font = bold
        if fill:
            header_cell.fill = fill
        for d, val in values.items():
            cell = ws.cell(row=row, column=col_positions[d])
            cell.value = val
            if fill:
                cell.fill = fill
        row += 1

    def blank():
        nonlocal row
        row += 1

    computed = {}

    for d in dates:

        if is_empty_day(df, d):
            computed[d] = {k: "" for k in [
                "nfp","no_flex","trainee","solo","crna","diff",
                "1:1","1:2","1:3.5","supervisory",
                "addition_supervisory_faculty","faculty_needed",
                "final_faculty","nl","mor","pct_solo",
                "faculty_sched","overage","crna_sched",
                "crna_demand","crna_needed"
            ]}
            continue

        total_rooms = int(df.loc["total_rooms", d])
        trainees = int(df.loc["trainees", d])
        crnas = int(df.loc["crnas", d])
        faculty = int(df.loc["faculty", d])

        fixed_crnas = 6
        cardiac_faculty = 3
        nl_fac = 5

        scheduled_flex_crnas = max(crnas - fixed_crnas, 0)

        available_rooms = total_rooms - fixed_crnas - cardiac_faculty
        available_faculty = faculty - cardiac_faculty

        trainee_rooms = min(trainees, available_rooms)
        faculty_for_trainees = round_1_decimal(min((trainee_rooms / 2) - 1.5, available_faculty))

        available_rooms -= trainee_rooms
        available_faculty -= faculty_for_trainees

        crna_rooms = min(scheduled_flex_crnas, available_rooms)
        faculty_for_crnas = round_1_decimal(min(crna_rooms / 3.5, available_faculty))

        available_rooms -= crna_rooms
        available_faculty -= faculty_for_crnas

        solo_faculty = min(available_faculty, available_rooms)

        supervisory = cardiac_faculty + faculty_for_trainees + faculty_for_crnas
        addition_faculty_supervisory = supervisory + solo_faculty
        faculty_needed = round_half_up(addition_faculty_supervisory)
        final_faculty_required = faculty_needed + 2
        overage_faculty = faculty - final_faculty_required

        mor_asc_sat = final_faculty_required - nl_fac
        percent_solo = round((solo_faculty / mor_asc_sat) * 100) if mor_asc_sat > 0 else 0

        crna_demand = total_rooms - fixed_crnas - trainee_rooms - solo_faculty
        crna_needed = crna_demand - scheduled_flex_crnas

        computed[d] = {
            "nfp": total_rooms,
            "no_flex": total_rooms - fixed_crnas,
            "trainee": trainees,
            "solo": solo_faculty,
            "crna": crna_rooms,
            "diff": (total_rooms - fixed_crnas) - trainees - crna_rooms,
            "1:1": cardiac_faculty,
            "1:2": faculty_for_trainees,
            "1:3.5": faculty_for_crnas,
            "supervisory": supervisory,
            "addition_supervisory_faculty": addition_faculty_supervisory,
            "faculty_needed": faculty_needed,
            "final_faculty": final_faculty_required,
            "nl": nl_fac,
            "mor": mor_asc_sat,
            "pct_solo": f"{percent_solo}%",
            "faculty_sched": faculty,
            "overage": overage_faculty,
            "crna_sched": scheduled_flex_crnas,
            "crna_demand": crna_demand,
            "crna_needed": crna_needed,
        }

    write_row("NFP demand", {d: computed[d]["nfp"] for d in dates})
    write_row("Demand (no Flex CRNAs)", {d: computed[d]["no_flex"] for d in dates})
    write_row("Main/NL/ASC Trainee", {d: computed[d]["trainee"] for d in dates})
    write_row("Solo Faculty", {d: computed[d]["solo"] for d in dates}, fill=yellow)

    out_wb.save(output_excel)
    print(f"Staffing report written to: {output_excel}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Input Excel file")
    args = parser.parse_args()

    output_file = args.input.replace(".xlsx", "_output.xlsx")
    run_staffing_model(args.input, output_file)
