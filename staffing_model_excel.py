import math
import argparse
import sys
from dataclasses import dataclass, asdict
import pandas as pd
import os


# ---------------------------------------------------------
# Parameters
# ---------------------------------------------------------

@dataclass
class Parameters:
    trainee_supervision_ratio: int = 2
    crna_supervision_ratio: int = 4

    crnas_location_buffer: int = 6
    faculty_location_buffer: int = 1
    board_runner_faculty: int = 1


# ---------------------------------------------------------
# Scenario input
# ---------------------------------------------------------

@dataclass
class ScenarioInput:
    name: str
    total_rooms: int
    trainees_available: int
    crnas_available: int
    faculty_available: int


# ---------------------------------------------------------
# Scenario output
# ---------------------------------------------------------

@dataclass
class ScenarioResult:
    name: str
    total_rooms: int

    crnas_total: int
    crnas_fixed_buffer: int
    crnas_available_for_rooms: int

    faculty_total: int
    faculty_main_or_buffer: int
    faculty_board_runner: int
    faculty_available_for_coverage: int

    rooms_covered_by_trainees: int
    rooms_covered_by_crnas: int

    faculty_used_for_trainee_supervision: int
    faculty_used_for_crna_supervision: int

    rooms_covered_by_solo_faculty: int

    faculty_buffer: int
    max_rooms_coverable: int
    rooms_left_to_cover: int


# ---------------------------------------------------------
# Core calculation (UNCHANGED)
# ---------------------------------------------------------

def compute_staffing(scenario: ScenarioInput, params: Parameters):

    total_rooms = scenario.total_rooms

    crnas_total = scenario.crnas_available
    faculty_total = scenario.faculty_available

    crnas_fixed_buffer = params.crnas_location_buffer
    faculty_main_or_buffer = params.faculty_location_buffer
    faculty_board_runner = params.board_runner_faculty

    crnas_available_for_rooms = max(
        crnas_total - crnas_fixed_buffer, 0
    )

    faculty_available_for_coverage = max(
        faculty_total - faculty_main_or_buffer - faculty_board_runner, 0
    )

    trainee_rooms = min(scenario.trainees_available, total_rooms)
    rooms_left = total_rooms - trainee_rooms

    crna_rooms = min(crnas_available_for_rooms, rooms_left)
    rooms_left -= crna_rooms

    faculty_for_trainees = (
        math.ceil(trainee_rooms / params.trainee_supervision_ratio)
        if trainee_rooms > 0 else 0
    )

    faculty_for_crnas = (
        math.ceil(crna_rooms / params.crna_supervision_ratio)
        if crna_rooms > 0 else 0
    )

    faculty_left = faculty_available_for_coverage - (
        faculty_for_trainees + faculty_for_crnas
    )
    faculty_left = max(faculty_left, 0)

    solo_rooms = min(faculty_left, rooms_left)
    faculty_left -= solo_rooms
    rooms_left -= solo_rooms

    faculty_buffer = faculty_left
    max_rooms_coverable = trainee_rooms + crna_rooms + solo_rooms
    rooms_left_to_cover = total_rooms - max_rooms_coverable

    return ScenarioResult(
        name=scenario.name,
        total_rooms=total_rooms,

        crnas_total=crnas_total,
        crnas_fixed_buffer=crnas_fixed_buffer,
        crnas_available_for_rooms=crnas_available_for_rooms,

        faculty_total=faculty_total,
        faculty_main_or_buffer=faculty_main_or_buffer,
        faculty_board_runner=faculty_board_runner,
        faculty_available_for_coverage=faculty_available_for_coverage,

        rooms_covered_by_trainees=trainee_rooms,
        rooms_covered_by_crnas=crna_rooms,

        faculty_used_for_trainee_supervision=faculty_for_trainees,
        faculty_used_for_crna_supervision=faculty_for_crnas,

        rooms_covered_by_solo_faculty=solo_rooms,

        faculty_buffer=faculty_buffer,
        max_rooms_coverable=max_rooms_coverable,
        rooms_left_to_cover=rooms_left_to_cover,
    )


def run_scenarios(scenarios, params):
    return pd.DataFrame(
        [asdict(compute_staffing(sc, params)) for sc in scenarios]
    )


# ---------------------------------------------------------
# Main (manual Excel input)
# ---------------------------------------------------------

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Run staffing model using Excel input"
    )
    parser.add_argument(
        "--input_file",
        required=True,
        help="Excel input file name"
    )

    args = parser.parse_args()
    input_file = args.input_file

    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        sys.exit(1)

    df_input = pd.read_excel(input_file)

    required_cols = {
        "scenario_name",
        "total_rooms",
        "trainees",
        "crnas",
        "faculty",
    }

    if not required_cols.issubset(df_input.columns):
        print(f"ERROR: Input file must contain columns {required_cols}")
        sys.exit(1)

    scenarios = [
        ScenarioInput(
            name=row["scenario_name"],
            total_rooms=int(row["total_rooms"]),
            trainees_available=int(row["trainees"]),
            crnas_available=int(row["crnas"]),
            faculty_available=int(row["faculty"]),
        )
        for _, row in df_input.iterrows()
    ]

    params = Parameters()
    df_output = run_scenarios(scenarios, params)

    output_file = input_file.replace(
        "_input.xlsx", "_output.xlsx"
    )

    df_output.to_excel(output_file, index=False)

    print(f"Results saved to {output_file}")


