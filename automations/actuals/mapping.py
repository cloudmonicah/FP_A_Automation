"""
Actuals validation mapping configuration.

Single source of truth for individual OPEX Tableau routes.
Direct Cost and Revenue can be migrated here later.
"""

from shared.settings import workbook_path

AGG_DIRECT = "DIRECT"
AGG_SUM = "SUM"
AGG_TOTAL = "TOTAL"
AGG_PENDING = "PENDING"
AGG_REVIEW = "REVIEW"

OPEX_MAPPINGS = [
    # =======================================================
    # ALIV
    # =======================================================
    {
        "company": "ALIV",
        "fs_line": "General and Administration",
        "tableau_category": "GENERAL AND ADMINISTRATION",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "ALIV Gen & Admin_Finance.twb"
        ),
        "individual_tableau_worksheet": "Monthly Actual v Budget by Sub account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Human Resources",
        "tableau_category": "HUMAN RESOURCES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Human Resources.twb"),
        "individual_tableau_worksheet": "ALIV Group Monthly _Actual v Budget by Sub Account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Financial Services",
        "tableau_category": "FINANCIAL SERVICES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Finance.twb"),
        "individual_tableau_worksheet": "ALIV Monthly Trend YTD by Sub Account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Legal",
        "tableau_category": "LEGAL",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Legal.twb"),
        "individual_tableau_worksheet": "ALIV Group Monthly YTD Actual v Budget by Sub account"
        "",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Facilities",
        "tableau_category": "FACILITIES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Engineering.twb"),
        "individual_tableau_worksheet": "Facilities (ALIV) - By Sub Account - Monthly YTD Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Information Technology",
        "tableau_category": "INFORMATION TECHNOLOGY",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Information Technology.twb"
        ),
        "individual_tableau_worksheet": "ALIV Monthly Subaccount Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Technical Operations",
        "tableau_category": "TECHNICAL OPERATIONS",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Engineering.twb"),
        "individual_tableau_worksheet": "Technical Operations (ALIV) - By Sub Account - Monthly YTD Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Commercial and Product Development",
        "tableau_category": "COMMERCIAL AND PRODUCT DEVELOPMENT",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Consumer Marketing & Product Development.twb"
        ),
        "individual_tableau_worksheet": "PRODUCT DEVELOPMENT (CBL & ALIV) - By Subaccount - Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Marketing",
        "tableau_category": "MARKETING",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Consumer Marketing & Product Development.twb"
        ),
        "individual_tableau_worksheet": "MARKETING (ALIV) - By Subaccount - Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Consumer Sales",
        "tableau_category": "CONSUMER SALES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Consumer Sales.twb"),
        "individual_tableau_worksheet": "ALIV Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Contact Center",
        "tableau_category": "CUSTOMER EXPERIENCE & CONTACT CENTER",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Customer Experience & Contact Center.twb"
        ),
        "individual_tableau_worksheet": "ALIV Monthly Trend by Sub Account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Roaming",
        "tableau_category": "ROAMING",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Consumer Marketing & Product Development.twb"
        ),
        "individual_tableau_worksheet": "ROAMING - By Subaccount - Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Corporate Sales",
        "tableau_category": "CORPORATE SALES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "ALIV Business.twb"),
        "individual_tableau_worksheet": "ALIV BUSINESS - Group Monthly YTD Actual v Budget by Sub account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Operations",
        "tableau_category": "OPERATIONS",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Operations.twb"),
        "individual_tableau_worksheet": "ALIV Monthly Subaccount Actual v Budget",
        "individual_actuals_enabled": True,
    },
    # {
    #     "company": "ALIV",
    #     "fs_line": "Service Delivery",
    #     "tableau_category": "SERVICE DELIVERY",
    #     "aggregation": AGG_DIRECT,
    #     "individual_tableau_workbook": workbook_path("OPEX", "Service Delivery.twb"),
    #     "individual_tableau_worksheet": "Monthly_Actual v Budget",
    #     "individual_actuals_enabled": True,
    # },
    {
        "company": "ALIV",
        "fs_line": "Operating Expenses",
        "tableau_category": "ALIV Total",
        "aggregation": AGG_TOTAL,
    },
    # =======================================================
    # CBL
    # =======================================================
    {
        "company": "CBL",
        "fs_line": "General & Administration",
        "tableau_category": "GENERAL AND ADMINISTRATION",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "CBL Gen & Admin_Finance.twb"
        ),
        "individual_tableau_worksheet": "Monthly YTD Actual v Budget by Sub account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Human Resources",
        "tableau_category": "HUMAN RESOURCES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Human Resources.twb"),
        "individual_tableau_worksheet": "CBL Group Monthly _Actual v Budget by Sub Account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Financial Services",
        "tableau_category": "FINANCIAL SERVICES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Finance.twb"),
        "individual_tableau_worksheet": "CBL Monthly Trend YTD by Sub Account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Legal",
        "tableau_category": "LEGAL",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Legal.twb"),
        "individual_tableau_worksheet": "CBL Group Monthly YTD Actual v Budget by Sub account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Facilities",
        "tableau_category": "FACILITIES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Engineering.twb"),
        "individual_tableau_worksheet": "Facilities (CBL) - By Sub Account - Monthly YTD Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Information Technology",
        "tableau_category": "INFORMATION TECHNOLOGY",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Information Technology.twb"
        ),
        "individual_tableau_worksheet": "CBL Monthly Subaccount Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Technical Operations",
        "tableau_category": "TECHNICAL OPERATIONS",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Engineering.twb"),
        "individual_tableau_worksheet": "Technical Operations (CBL) - By Sub Account - Monthly YTD Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Commercial & Product Development",
        "tableau_category": "COMMERCIAL AND PRODUCT DEVELOPMENT",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Consumer Marketing & Product Development.twb"
        ),
        "individual_tableau_worksheet": "PRODUCT DEVELOPMENT (CBL & ALIV) - By Subaccount - Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Marketing",
        "tableau_category": "MARKETING",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Consumer Marketing & Product Development.twb"
        ),
        "individual_tableau_worksheet": "MARKETING (CBL) - By Subaccount - Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Consumer Sales",
        "tableau_category": "CONSUMER SALES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Consumer Sales.twb"),
        "individual_tableau_worksheet": "CBL Monthly Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Service Delivery",
        "tableau_category": "SERVICE DELIVERY",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Service Delivery.twb"),
        "individual_tableau_worksheet": "Monthly_Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Customer Experience & Contact Center",
        "tableau_category": "CUSTOMER EXPERIENCE & CONTACT CENTER",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path(
            "OPEX", "Customer Experience & Contact Center.twb"
        ),
        "individual_tableau_worksheet": "CBL Monthly Trend by Sub Account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Corporate Sales",
        "tableau_category": "CORPORATE SALES",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "ALIV Business.twb"),
        "individual_tableau_worksheet": "ALIV BUSINESS - Group Monthly YTD Actual v Budget by Sub account",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Outside Plant",
        "tableau_category": "OUTSIDE PLANT",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Outside Plant.twb"),
        "individual_tableau_worksheet": "Monthly Subaccount Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Operations",
        "tableau_category": "OPERATIONS",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Operations.twb"),
        "individual_tableau_worksheet": "CBL Monthly Subaccount Actual v Budget",
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Media",
        "tableau_category": "MEDIA",
        "aggregation": AGG_DIRECT,
        "individual_tableau_workbook": workbook_path("OPEX", "Media.twb"),
        "individual_tableau_worksheet": "Monthly Subaccount Actual v Budget",
        "individual_actuals_enabled": True,
    },
    # {
    #     "company": "CBL",
    #     "fs_line": "Total Opex",
    #     "tableau_category": "CBL Total",
    #     "aggregation": AGG_TOTAL,
    # },
]


# ============================================================
# INDIVIDUAL DIRECT COST ACTUALS MAPPINGS
# Financial Statement -> Individual Tableau Workbook
# ============================================================

_DC_FOLDER = workbook_path("Direct Cost")

INDIVIDUAL_DIRECT_COST_MAPPINGS = [
    # ---------------- ALIV ----------------
    {
        "company": "ALIV",
        "fs_line": "Interconnect",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Group Direct Costs_Consumer Product Development.twb",
        "individual_tableau_worksheet": "ALIV Group Direct Costs Details YTD - Actual v Budget",
        "tableau_categories": ["INTERCONNECT COST"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Reseller Commissions",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Engineering.twb",
        "individual_tableau_worksheet": "ALIV Group Direct Costs Details",
        "tableau_categories": [
            "COMMISSIONS - PUSH TO TALK",
            "COMMISSIONS - RECAHRAGE DEALERS",
            "COMMISSIONS - KIOSKS",
            "COMISSIONS - BTS RECHARGE",
            "COMMISSIONS - OTHER SERVICE COMMISSIONS & BANK CHA",
        ],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Roaming Service Costs",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Group Direct Costs_Consumer Product Development.twb",
        "individual_tableau_worksheet": "ALIV Group Direct Costs Details YTD - Actual v Budget",
        "tableau_categories": ["ROAMING COSTS"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Service Fees",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Finance.twb",
        "individual_tableau_worksheet": "Group Direct Costs Details",
        "tableau_categories": ["SERVICE FEES"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Direct Cost",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Engineering.twb",
        "individual_tableau_worksheet": "ALIV Group Direct Costs Details",
        "tableau_categories": [],
        "aggregation": AGG_TOTAL,
        "individual_actuals_enabled": True,
        "note": "Total/control mapping; do not double count as a component.",
    },
    # ---------------- CBL ----------------
    {
        "company": "CBL",
        "fs_line": "Programming fees",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\CBL Content Fees.twb",
        "individual_tableau_worksheet": "CBL Programming Fees - Monthly Budget v Actual",
        "tableau_categories": ["CONTENT FEES"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Hardware/PBX",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Group_Direct Costs_Aliv Business.twb",
        "individual_tableau_worksheet": "CBL Group Direct Costs Details - Actual v Budget",
        "tableau_categories": ["PBX"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Voice termination costs",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Engineering.twb",
        "individual_tableau_worksheet": "CBL Group Direct Costs Details",
        "tableau_categories": ["INTERCONNECT COST"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
        "note": "Worksheet not specified; sum all applicable interconnect cost rows.",
    },
    {
        "company": "CBL",
        "fs_line": "Internet Bandwidth",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Engineering.twb",
        "individual_tableau_worksheet": "CBL Group Direct Costs Details",
        "tableau_categories": ["INTERCONNECT CAPACITY"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
        "note": "Sum all applicable interconnect bandwidth rows.",
    },
    {
        "company": "CBL",
        "fs_line": "Cloud services COS",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Group_Direct Costs_Aliv Business.twb",
        "individual_tableau_worksheet": "CBL Group Direct Costs Details - Actual v Budget",
        "tableau_categories": ["SERVICE FEES"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "DTH cost of sales",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Engineering.twb",
        "individual_tableau_worksheet": "CBL Group Direct Costs Details",
        "tableau_categories": ["DTH"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Smart Home Products COS",
        "individual_tableau_workbook": rf"{_DC_FOLDER}\Cable Bahamas Group Direct Costs_Engineering.twb",
        "individual_tableau_worksheet": "CBL Group Direct Costs Details",
        "tableau_categories": ["SMART HOME PRODUCTS COST"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
]

# DIRECT COST TEST MAPPINGS (managed block)
DIRECT_COST_MAPPINGS = [
    dict(
        company="ALIV",
        fs_line="Interconnect",
        tableau_categories=["INTERCONNECT COST"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="ALIV",
        fs_line="Reseller Commissions",
        tableau_categories=["PUSH TO TALK COMMISSION", "TOP UP & PLAN COMMISSION"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="ALIV",
        fs_line="Roaming Service Costs",
        tableau_categories=["ROAMING COSTS"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="ALIV",
        fs_line="Service Fees",
        tableau_categories=["SERVICE FEES"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="ALIV",
        fs_line="Direct Cost",
        tableau_categories=[],
        aggregation=AGG_TOTAL,
    ),
    dict(
        company="CBL",
        fs_line="Programming fees",
        tableau_categories=["CONTENT FEES"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="Hardware/PBX",
        tableau_categories=["PBX"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="Voice termination costs",
        tableau_categories=["INTERCONNECT COST"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="Internet Bandwidth",
        tableau_categories=["INTERCONNECT CAPACITY"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="Cloud services COS",
        tableau_categories=["SERVICE FEES"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="DTH cost of sales",
        tableau_categories=["DTH"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="Smart Home Products COS",
        tableau_categories=["SMART HOME PRODUCTS COST"],
        aggregation=AGG_SUM,
    ),
    dict(
        company="CBL",
        fs_line="Ocean Wifi costs",
        tableau_categories=[],
        aggregation="PENDING",
        note="Tableau category unconfirmed; leave blank",
    ),
    dict(
        company="CBL",
        fs_line="Cost of Goods Sold",
        tableau_categories=[],
        aggregation=AGG_TOTAL,
    ),
]

# END DIRECT COST TEST MAPPINGS

# ============================================================
# REVENUE ACTUALS VALIDATION MAPPINGS
# Financial Statements -> Tableau Revenue Workbooks
# ============================================================

_REVENUE_FOLDER = workbook_path("Revenue")

REVENUE_GROUP_WORKBOOK = None

REVENUE_GROUP_WORKSHEET = "CBL FS Rollup View Group Monthly Revenue YTD_SUMMARY"

REVENUE_CONSUMER_WORKBOOK = (
    rf"{_REVENUE_FOLDER}\Cable Bahamas Group_CONSUMER REVENUE.twb"
)

REVENUE_CORPORATE_WORKBOOK = (
    rf"{_REVENUE_FOLDER}\Cable Bahamas Group_CORPORATE REVENUE.twb"
)
REVENUE_MEDIA_WORKBOOK = rf"{_REVENUE_FOLDER}\Cable Bahamas Group_MEDIA REVENUE.twb"


# ============================================================
# CONSOLIDATED REVENUE MAPPINGS
# ============================================================

REVENUE_MAPPINGS = [
    # ---------------- ALIV ----------------
    {
        "company": "ALIV",
        "fs_line": "Consumer Prepaid",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "ALIV",
        "fs_line": "Consumer Postpaid",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "ALIV",
        "fs_line": "Corporate",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "ALIV",
        "fs_line": "Visitor Roaming",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "ALIV",
        "fs_line": "Interconnect",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "ALIV",
        "fs_line": "Other",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "ALIV",
        "fs_line": "Service Revenues",
        "tableau_categories": [],
        "aggregation": AGG_TOTAL,
        "note": "Financial statement total/control; do not double count.",
    },
    # ---------------- CBL ----------------
    {
        "company": "CBL",
        "fs_line": "Residential",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "Corporate - Recurring",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "Corporate - One off",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "Media",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "Other",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "MSA-ALIV",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "Government",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
    },
    {
        "company": "CBL",
        "fs_line": "Gross Revenue",
        "tableau_categories": [],
        "aggregation": AGG_TOTAL,
        "note": "Financial statement total/control; do not double count.",
    },
]


# ============================================================
# INDIVIDUAL REVENUE WORKBOOK MAPPINGS
# ============================================================

INDIVIDUAL_REVENUE_MAPPINGS = [
    # ---------------- ALIV ----------------
    {
        "company": "ALIV",
        "fs_line": "Consumer Prepaid",
        "individual_tableau_workbook": REVENUE_CONSUMER_WORKBOOK,
        "individual_tableau_worksheet": (
            "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": ["PAY IN ADVANCE"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Consumer Postpaid",
        "individual_tableau_workbook": REVENUE_CONSUMER_WORKBOOK,
        "individual_tableau_worksheet": (
            "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": ["POSTPAID"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Visitor Roaming",
        "individual_tableau_workbook": REVENUE_CONSUMER_WORKBOOK,
        "individual_tableau_worksheet": (
            "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": ["VISITOR ROAMING"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Interconnect",
        "individual_tableau_workbook": REVENUE_CONSUMER_WORKBOOK,
        "individual_tableau_worksheet": (
            "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": ["INTERCONNECT REVENUE"],
        "aggregation": AGG_SUM,
        "individual_actuals_enabled": True,
    },
    {
        "company": "ALIV",
        "fs_line": "Other",
        "individual_tableau_workbook": REVENUE_CONSUMER_WORKBOOK,
        "individual_tableau_worksheet": (
            "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": [],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
        "note": "Confirm all applicable OTHER CONSUMER GL FS Type values.",
    },
    {
        "company": "ALIV",
        "fs_line": "Corporate",
        "individual_tableau_workbook": REVENUE_CORPORATE_WORKBOOK,
        "individual_tableau_worksheet": "",
        "mapping_column": "GL FS Type",
        "mapping_values": [],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
        "note": "Confirm ALIV Corporate worksheet and categories.",
    },
    {
        "company": "ALIV",
        "fs_line": "Service Revenues",
        "aggregation": AGG_TOTAL,
        "individual_actuals_enabled": True,
        "note": "Control total; not an individual component.",
    },
    # ---------------- CBL ----------------
    {
        "company": "CBL",
        "fs_line": "Residential",
        "individual_tableau_workbook": REVENUE_CONSUMER_WORKBOOK,
        "individual_tableau_worksheet": (
            "CBL RESIDENTIAL REVENUE SUMMARY - Monthly Actuals"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": [],
        "aggregation": AGG_TOTAL,
        "individual_actuals_enabled": True,
        "note": "Use Residential Grand Total; verify the crosstab total.",
    },
    {
        "company": "CBL",
        "fs_line": "Corporate - Recurring",
        "individual_tableau_workbook": REVENUE_CORPORATE_WORKBOOK,
        "individual_tableau_worksheet": (
            "CBL COMMERCIAL REVENUE DETAILS - Monthly Actuals"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": [],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Corporate - One off",
        "individual_tableau_workbook": REVENUE_CORPORATE_WORKBOOK,
        "individual_tableau_worksheet": (
            "CBL COMMERCIAL REVENUE DETAILS - Monthly Actuals"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": ["CORPORATE POST PAID - ONE OFF"],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Government",
        "individual_tableau_workbook": REVENUE_CORPORATE_WORKBOOK,
        "individual_tableau_worksheet": (
            "CBL COMMERCIAL REVENUE DETAILS - Monthly Actuals"
        ),
        "mapping_column": "GL FS Type",
        "mapping_values": ["CORPORATE GOVERNMENT"],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Media",
        "individual_tableau_workbook": REVENUE_MEDIA_WORKBOOK,
        "individual_tableau_worksheet": (" Monthly YTD Actual v Budget by Sub account"),
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Other",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "MSA-ALIV",
        "tableau_categories": [],
        "aggregation": AGG_PENDING,
        "individual_actuals_enabled": True,
    },
    {
        "company": "CBL",
        "fs_line": "Gross Revenue",
        "aggregation": AGG_TOTAL,
        "individual_actuals_enabled": True,
        "note": "Control total; do not double count.",
    },
]


# ============================================================
# END REVENUE ACTUALS VALIDATION MAPPINGS
# ============================================================
