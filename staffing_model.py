import math
import argparse
import sys
from dataclasses import dataclass, asdict
import pandas as pd


# ---------------------------------------------------------
# Parameters
# ---------------------------------------------------------

@dataclass
class Parameters:
    trainee_supervision_ratio: int = 2   # 1 faculty per 2 trainee rooms
    crna_supervision_ratio: int = 4      # 1 faculty per 4 CRNA rooms

    crnas_location_buffer: int = 6       # Fixed CRNA buffers
    faculty_location_buffer: int = 1     # Main OR buffer faculty
    board_runner_faculty: int = 1        # Board runner


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
    rooms_covered_by_faculty: int

    faculty_used_for_trainee_supervision: int
    faculty_used_for_crna_supervision: int

    faculty_buffer: int
    max_rooms_coverable: int
    rooms_left_to_cover: int


# ---------------------------------------------------------
# Core calculation
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

    # Step 1: Trainee rooms
    trainee_rooms = min(scenario.trainees_available, total_rooms)
    rooms_left = total_rooms - trainee_rooms

    # Step 2: CRNA rooms
    crna_rooms = min(crnas_available_for_rooms, rooms_left)
    rooms_left -= crna_rooms

    # Step 3: Supervision
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

    # Step 4: Solo faculty rooms
    solo_rooms = min(faculty_left, rooms_left)
    faculty_left -= solo_rooms
    rooms_left -= solo_rooms

    faculty_buffer = faculty_left

    max_rooms_coverable = (
        trainee_rooms + crna_rooms + solo_rooms
    )

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
        rooms_covered_by_faculty=solo_rooms,

        faculty_used_for_trainee_supervision=faculty_for_trainees,
        faculty_used_for_crna_supervision=faculty_for_crnas,

        faculty_buffer=faculty_buffer,
        max_rooms_coverable=max_rooms_coverable,
        rooms_left_to_cover=rooms_left_to_cover,
    )


# ---------------------------------------------------------
# Run scenarios
# ---------------------------------------------------------

def run_scenarios(scenarios, params):
    rows = []
    for sc in scenarios:
        res = compute_staffing(sc, params)
        rows.append(asdict(res))
    return pd.DataFrame(rows)


# ---------------------------------------------------------
# CLI utilities
# ---------------------------------------------------------

def parse_csv_ints(value, name):
    try:
        return [int(x.strip()) for x in value.split(",")]
    except ValueError:
        raise ValueError(f"Invalid integer in {name}")


def build_scenarios(total_rooms, trainees, crnas, faculty):
    if not (
        len(total_rooms)
        == len(trainees)
        == len(crnas)
        == len(faculty)
    ):
        raise ValueError("All input lists must be the same length")

    scenarios = []
    for i in range(len(total_rooms)):
        scenarios.append(
            ScenarioInput(
                name=f"Scenario {i+1}",
                total_rooms=total_rooms[i],
                trainees_available=trainees[i],
                crnas_available=crnas[i],
                faculty_available=faculty[i],
            )
        )
    return scenarios


# ---------------------------------------------------------
# Main entry
# ---------------------------------------------------------

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Anesthesia staffing model"
    )

    parser.add_argument("--total_rooms", required=True)
    parser.add_argument("--trainees", required=True)
    parser.add_argument("--crnas", required=True)
    parser.add_argument("--faculty", required=True)

    args = parser.parse_args()

    try:
        total_rooms = parse_csv_ints(args.total_rooms, "total_rooms")
        trainees = parse_csv_ints(args.trainees, "trainees")
        crnas = parse_csv_ints(args.crnas, "crnas")
        faculty = parse_csv_ints(args.faculty, "faculty")

        scenarios = build_scenarios(
            total_rooms, trainees, crnas, faculty
        )

    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    params = Parameters()
    df = run_scenarios(scenarios, params)

    print(df)

    output_file = "staffing_model_output.xlsx"
    df.to_excel(output_file, index=False)

    print(f"\nResults saved to {output_file}")
