import math
import argparse
import sys
from dataclasses import dataclass, asdict
import pandas as pd
import os


# ---------------------------------------------------------
# Parameters (aligned with Archit's file)
# ---------------------------------------------------------

@dataclass
class Parameters:
    trainee_supervision_ratio: float = 2.0      # 1 faculty per 2 trainees
    crna_supervision_ratio: float = 3.5         # 1 faculty per 3.5 CRNAs

    fixed_crnas: int = 6                         # Location CRNA buffer
    fixed_faculty_cardiac: int = 3              # Cardiac faculty
    fixed_faculty_main_or: int = 1               # Main OR
    fixed_faculty_nl_or: int = 1                 # NL OR


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

    trainees_used: int
    crnas_total: int
    crnas_fixed: int
    crnas_available_for_rooms: int

    faculty_total: int
    faculty_fixed_total: int
    faculty_available_for_coverage: int

    faculty_used_for_trainee_supervision: int
    faculty_used_for_crna_supervision: int

    solo_faculty_rooms: int

    crna_demand: int
    crnas_shortage: int

    faculty_buffer: int
    max_rooms_coverable: int
    rooms_left_to_cover: int


# ---------------------------------------------------------
# Core calculation
# ---------------------------------------------------------

def compute_staffing(scenario: ScenarioInput, params: Parameters):

    # -------------------------
    # Fixed staff
    # -------------------------
    fixed_faculty_total = (
        params.fixed_faculty_cardiac
        + params.fixed_faculty_main_or
        + params.fixed_faculty_nl_or
    )

    faculty_available_for_coverage = max(
        scenario.faculty_available - fixed_faculty_total, 0
    )

    crnas_available_for_rooms = max(
        scenario.crnas_available - params.fixed_crnas, 0
    )

    total_rooms = scenario.total_rooms

    # -------------------------
    # Assign trainees
    # -------------------------
    trainee_rooms = min(scenario.trainees_available, total_rooms)
    rooms_left = total_rooms - trainee_rooms

    # -------------------------
    # Assign CRNAs
    # -------------------------
    crna_rooms = min(crnas_available_for_rooms, rooms_left)
    rooms_left -= crna_rooms

    # -------------------------
    # Supervision
    # -------------------------
    faculty_for_trainees = math.ceil(
        trainee_rooms / params.trainee_supervision_ratio
    ) if trainee_rooms > 0 else 0

    faculty_for_crnas = math.ceil(
        crna_rooms / params.crna_supervision_ratio
    ) if crna_rooms > 0 else 0

    faculty_left = faculty_available_for_coverage - (
        faculty_for_trainees + faculty_for_crnas
    )
    faculty_left = max(faculty_left, 0)

    # -------------------------
    # Solo faculty rooms
    # -------------------------
    solo_faculty_rooms = min(faculty_left, rooms_left)
    faculty_left -= solo_faculty_rooms
    rooms_left -= solo_faculty_rooms

    # -------------------------
    # CRNA demand (Archit logic)
    # -------------------------
    crna_demand = max(
        total_rooms
        - trainee_rooms
        - solo_faculty_rooms
        - params.fixed_crnas,
        0
    )

    crnas_shortage = max(crna_demand - scenario.crnas_available, 0)

    # -------------------------
    # Final
    # -------------------------
    max_rooms_coverable = (
        trainee_rooms + crna_rooms + solo_faculty_rooms
    )

    return ScenarioResult(
        name=scenario.name,
        total_rooms=total_rooms,

        trainees_used=trainee_rooms,
        crnas_total=scenario.crnas_available,
        crnas_fixed=params.fixed_crnas,
        crnas_available_for_rooms=crnas_available_for_rooms,

        faculty_total=scenario.faculty_available,
        faculty_fixed_total=fixed_faculty_total,
        faculty_available_for_coverage=faculty_available_for_coverage,

        faculty_used_for_trainee_supervision=faculty_for_trainees,
        faculty_used_for_crna_supervision=faculty_for_crnas,

        solo_faculty_rooms=solo_faculty_rooms,

        crna_demand=crna_demand,
        crnas_shortage=crnas_shortage,

        faculty_buffer=faculty_left,
        max_rooms_coverable=max_rooms_coverable,
        rooms_left_to_cover=rooms_left,
    )


def run_scenarios(scenarios, params):
    return pd.DataFrame(
        [asdict(compute_staffing(sc, params)) for sc in scenarios]
    )


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Run staffing model using Excel input"
    )
    parser.add_argument("--input_file", required=True)

    args = parser.parse_args()
    input_file = args.input_file

    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        sys.exit(1)

    df_input = pd.read_excel(input_file)

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

    output_file = input_file.replace("_input.xlsx", "_output.xlsx")
    df_output.to_excel(output_file, index=False)

    print(f"Results saved to {output_file}")
