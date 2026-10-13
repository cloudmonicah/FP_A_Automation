from pathlib import Path
from datetime import datetime
import argparse
import json
import traceback

from automations.actuals.opex_validation_service import (
    run_live_opex_validation,
)
from automations.actuals.individual_actuals_validation_service import (
    run_individual_opex_pilot,
)
from automations.actuals.actuals_report_layout import report_status
from automations.actuals.mapping import OPEX_MAPPINGS


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--aliv-fs", required=True)
    parser.add_argument("--cbl-fs", required=True)
    parser.add_argument("--tableau", required=True)
    parser.add_argument("--period", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument(
        "--company",
        choices=["ALIV", "CBL", "Both"],
        default="ALIV",
    )
    args = parser.parse_args()

    closed_period = datetime.strptime(
        args.period, "%Y-%m-%d"
    ).date()

    for name, value in (
        ("ALIV statement", args.aliv_fs),
        ("CBL statement", args.cbl_fs),
        ("Consolidated Tableau", args.tableau),
    ):
        if not Path(value).is_file():
            raise FileNotFoundError(
                f"{name} not accessible: {value}"
            )

    enabled = [
        route for route in OPEX_MAPPINGS
        if route.get("individual_actuals_enabled") is True
        and (
            args.company == "Both"
            or route["company"].upper() == args.company
        )
    ]

    if not enabled:
        raise RuntimeError(
            "No individual OPEX mappings enabled "
            "for selected company."
        )

    output = Path(args.output)

    output_paths = {
        "opex": str(output / "OPEX"),
        "tableau_raw_copies": str(
            output / "Tableau_Raw_Copies"
        ),
        "logs": str(output / "Logs"),
    }

    for folder in output_paths.values():
        Path(folder).mkdir(
            parents=True, exist_ok=True
        )

    print("\nSTEP 1: Consolidated OPEX validation")

    consolidated = run_live_opex_validation(
        tableau_workbook_path=Path(args.tableau),
        aliv_fs_path=Path(args.aliv_fs),
        cbl_fs_path=Path(args.cbl_fs),
        output_paths=output_paths,
        closed_period=closed_period,
        open_tableau_first=True,
        run_individual_validation=False,
    )

    report_path = Path(
        consolidated["validation_workbook"]
    )

    print("Consolidated report:", report_path)

    print("\nSTEP 2: Individual OPEX validation")

    individual = run_individual_opex_pilot(
        report_path=report_path,
        aliv_fs_path=Path(args.aliv_fs),
        cbl_fs_path=Path(args.cbl_fs),
        output_paths=output_paths,
        closed_period=closed_period,
        tableau_settings={"company": args.company},
    )

    print("\nSTEP 3: Final reconciliation status")

    final_status = report_status(report_path)

    results = {
        "company": args.company,
        "period": args.period,
        "consolidated_status": consolidated.get(
            "overall_status"
        ),
        "individual_status": individual.get("status"),
        "individual_comparisons": individual.get("statuses"),
        "final_status": final_status,
        "report": str(report_path),
        "audit": individual.get("audit_workbook"),
    }

    print(json.dumps(results, indent=2, default=str))

    if (
        individual.get("status") != "PASS"
        or final_status != "PASS"
    ):
        raise RuntimeError(
            "Validation requires review. "
            f"Open the report: {report_path}"
        )

    print("\nSUCCESS: OPEX validation passed.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise SystemExit(1)
