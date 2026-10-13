from __future__ import annotations

# ============================================================
# VALIDATION STATUS CONSTANTS
# ============================================================

STATUS_PASS = "PASS"
STATUS_FAIL = "FAIL"
STATUS_PENDING = "PENDING"
STATUS_ACTION_REQUIRED = "ACTION REQUIRED"


# ============================================================
# AGGREGATION TYPES
# ============================================================

AGG_DIRECT = "DIRECT"
AGG_SUM = "SUM"
AGG_TOTAL = "TOTAL"
AGG_PENDING = "PENDING"
AGG_REVIEW = "REVIEW"


# ============================================================
# TABLEAU SOURCES
# ============================================================

OPEX_TABLEAU_WORKBOOK = None  # Source file is selected by the user.

OPEX_TABLEAU_WORKSHEET = "FY26 OPEX Summary YTD by Dept GL (Excl Media)"

DIRECT_COST_TABLEAU_WORKBOOK = None  # Source file is selected by the user.

DIRECT_COST_TABLEAU_WORKSHEET = "Group Direct Costs Year to Date Summary"

REVENUE_TABLEAU_WORKBOOK = None  # Source file is selected by the user.


# ============================================================
# FINANCIAL STATEMENT SOURCES
# ============================================================

ALIV_FINANCIAL_STATEMENT = None  # Source file is selected by the user.

CBL_FINANCIAL_STATEMENT = None  # Source file is selected by the user.


# ============================================================
# OPEX VALIDATION MAPPINGS
# ============================================================

from .mapping import OPEX_MAPPINGS

# ============================================================
# DIRECT COST / COST OF GOODS SOLD TABLEAU SOURCE
# ============================================================
#
# Direct Cost uses the DETAIL worksheet so that all underlying
# GL rows can be aggregated before comparison to the Financial
# Statements.
# ============================================================

DIRECT_COST_TABLEAU_WORKBOOK = None  # Source file is selected by the user.

DIRECT_COST_TABLEAU_WORKSHEET = "Group Direct Costs Details"


# ============================================================
# DIRECT COST VALIDATION MAPPINGS
# ============================================================
#
# IMPORTANT:
#
# Tableau contains GL-level detail.
#
# We first aggregate every row belonging to the same
# GL Account Mapping.
#
# Then we map those Tableau category totals to the Financial
# Statement line.
#
# Example:
#
# Financial Statement:
#     Roaming Service Costs
#
# Tableau:
#     ROAMING COSTS - SERVICE COSTS
#     ROAMING COSTS - DIGICEL COMMISSIONS
#     ROAMING COSTS - CONSULTING COSTS
#     ROAMING COSTS - TECHNICAL COSTS
#
# All rows fall under:
#     GL Account Mapping = ROAMING COSTS
#
# Therefore the entire ROAMING COSTS category is summed.
# ============================================================

from .mapping import DIRECT_COST_MAPPINGS, INDIVIDUAL_DIRECT_COST_MAPPINGS

# ============================================================
# Do not maintain individual workbook routes here.
# All paths, worksheets and enabled flags come from mapping.py.

_INDIVIDUAL_OPEX_DEFAULTS = {
    "actual_measure": "Actual YTD",
    "monthly_scope": "CLOSED_MONTH",
    "ytd_scope": "FISCAL_START_THROUGH_CLOSED_MONTH",
    "fiscal_start_month": 7,
    "statement_sheet": "P & L_",
    "statement_label_column": 2,
    "statement_monthly_column": 3,
    "statement_ytd_column": 7,
    "statement_period_cell": "D1",
    "statement_monthly_header_cell": "C5",
    "statement_ytd_header_cell": "G5",
    "statement_ytd_section_cell": "H3",
}

INDIVIDUAL_OPEX_ACTUALS_MAPPINGS = [
    {
        **_INDIVIDUAL_OPEX_DEFAULTS,
        **route,
        "mapping_id": (
            route["company"] + "_" + route["fs_line"].upper().replace(" ", "_")
        ),
        "output_company_sheet": (route["company"] + " OPEX"),
    }
    for route in OPEX_MAPPINGS
    if route.get("individual_tableau_workbook")
    and route.get("individual_tableau_worksheet")
]
