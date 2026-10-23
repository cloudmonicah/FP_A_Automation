"""Step 3 individual Tableau selections, applied on top of mapping.FORECAST_ROUTES.

For each forecast sheet: the individual workbook/tab to read and which of its
lines make up the sheet - the same selection its consolidated tab makes
(Company, GL FS Category, GL Account Mapping, GL Sub Account Description),
limited to fields the individual tab shows in its rows. Derived from
FY27 FORECAST WORKBOOKS.twb and the OPEX / Direct Cost workbook folders
(October 2026). Lines not listed keep the mapping as it is.
"""

SELECTIONS = {
    ('FY27 Legal BOARD EXPENSES', 'ALIV'): {
        "individual_tableau_workbook": 'OPEX _Board Expenses.twb',
        "individual_tableau_worksheet": 'CBL & ALIV Board Expenses - Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Sub Account Description': [
                'GEN ADMIN - BOARD MEETING COSTS',
            ],
        },
        # Not checkable in the individual tab's rows: GL Account Mapping (include 1), GL FS Category (include 2)
    },
    ('FY27 Legal BOARD EXPENSES', 'CBL'): {
        "individual_tableau_workbook": 'OPEX _Board Expenses.twb',
        "individual_tableau_worksheet": 'CBL & ALIV Board Expenses - Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Sub Account Description': [
                'GEN ADMIN - BOARD MEETING COSTS',
            ],
        },
        # Not checkable in the individual tab's rows: GL Account Mapping (include 1), GL FS Category (include 2)
    },
    ('FY27 OPEX Information Technology', 'ALIV'): {
        "individual_tableau_workbook": 'Information Technology.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Information Technology', 'CBL'): {
        "individual_tableau_workbook": 'Information Technology.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Engineering & Technology', 'Tech Ops - ALIV'): {
        "individual_tableau_workbook": 'Engineering.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast by Sup Department',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL FS Category': [
                'TECHNICAL OPERATIONS',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4)
    },
    ('FY27 OPEX Engineering & Technology', 'Tech Ops - CBL'): {
        "individual_tableau_workbook": 'Engineering.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast by Sup Department',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL FS Category': [
                'TECHNICAL OPERATIONS',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 4), Type (include 4)
    },
    ('FY27 OPEX Engineering & Technology', 'Facilities - ALIV'): {
        "individual_tableau_workbook": 'Group Premises & Facilities_ALL Departments.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL FS Category': [
                'FACILITIES',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4)
    },
    ('FY27 OPEX Engineering & Technology', 'Facilities- CBL'): {
        "individual_tableau_workbook": 'Group Premises & Facilities_ALL Departments.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL FS Category': [
                'FACILITIES',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 5), Type (include 4)
    },
    ('FY27 OPEX Customer Experience & Contact Center', 'ALIV'): {
        "individual_tableau_workbook": 'Customer Experience & Contact Center.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Customer Experience & Contact Center', 'CBL'): {
        "individual_tableau_workbook": 'Customer Experience & Contact Center.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Consumer Sales', 'ALIV'): {
        "individual_tableau_workbook": 'Consumer Sales.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4)
    },
    ('FY27 OPEX Consumer Sales', 'CBL'): {
        "individual_tableau_workbook": 'Consumer Sales.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4)
    },
    ('FY27 OPEX Consumer Marketing & Product Development', 'COMM & PROD - CBL'): {
        "individual_tableau_workbook": 'Consumer Marketing & Product Development.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast by Subdepartment',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL FS Category': [
                'COMMERCIAL AND PRODUCT DEVELOPMENT',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4)
    },
    ('FY27 OPEX Consumer Marketing & Product Development', 'COMM & PROD - ALIV'): {
        "individual_tableau_workbook": 'Consumer Marketing & Product Development.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast Summary by Subdepartment',  # 'by Subdepartment' is saved filtered to CBL
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL FS Category': [
                'COMMERCIAL AND PRODUCT DEVELOPMENT',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4)
    },
    ('FY27 OPEX Consumer Marketing & Product Development', 'Marketing - CBL'): {
        "individual_tableau_workbook": 'Consumer Marketing & Product Development.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast by Subdepartment',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL FS Category': [
                'MARKETING',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 4), Type (include 4)
    },
    ('FY27 OPEX Consumer Marketing & Product Development', 'Marketing - ALIV'): {
        "individual_tableau_workbook": 'Consumer Marketing & Product Development.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast Summary by Subdepartment',  # 'by Subdepartment' is saved filtered to CBL
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL FS Category': [
                'MARKETING',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 5), Type (include 4)
    },
    ('FY27 OPEX Consumer Marketing & Product Development', 'ROAMING'): {
        "individual_tableau_workbook": 'Consumer Marketing & Product Development.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast Summary by Subdepartment',  # 'by Subdepartment' is saved filtered to CBL
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL FS Category': [
                'ROAMING',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4)
    },
    ('FY27 OPEX Consumer Marketing & Product Development', 'Marketing - B2B (ALIV)'): {
        "individual_tableau_workbook": 'Group Marketing & Sponsorships_ALL Departments.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL FS Category': [
                'CORPORATE SALES',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4)
    },
    ('FY27 OPEX ALIV Head Office', 'Aliv Head Office'): {
        "individual_tableau_workbook": 'ALIV Head Office_John Gomez.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'GL Account Mapping': [
                'GEN ADMIN',
                'OTHER STAFF COSTS',
                'PAYROLL',
                'PROFESSIONAL FEES',
                'SPONSORSHIP & DONATIONS',
                'TRAVEL',
            ],
            'GL Sub Account Description': [
                'GEN ADMIN - EXTERNAL MEETING EXPENSES',
                'GEN ADMIN - GROUP RECHARGE OTHER',
                'GEN ADMIN - OTHER',
                'OTHER STAFF COSTS - STAFF MEETINGS',
                'OTHER STAFF COSTS-OTHER MISC. COSTS',
                'OTHER STAFF COSTS-PROFESSIONAL SUBSCRIPTIONS',
                'OTHER STAFF COSTS-TEAM INCENTIVES',
                'OTHER STAFF COSTS-TRAINING & DEVELOPMENT',
                'OTHER STAFF COSTS-UNIFORMS',
                'PAYROLL-BONUS EXPENSE',
                'PAYROLL-GROUP MEDICAL',
                'PAYROLL-GROUP RECHARGE PAYROLL COSTS',
                'PAYROLL-NATIONAL INSURANCE',
                'PAYROLL-OVERTIME / CALL OUT',
                'PAYROLL-PENSION',
                'PAYROLL-SALARIES',
                'PAYROLL-VEHICLE ALLOWANCE',
                'SPONSORSHIP-ART, MUSIC & FESTIVALS',
                'TRAVEL  - LOCAL FLIGHTS',
                'TRAVEL  - LOCAL VEHICLE RENTAL',
                'TRAVEL  - OVERSEAS ACCOMODATIONS',
                'TRAVEL  - OVERSEAS FLIGHTS',
                'TRAVEL  - OVERSEAS MEALS & MISC',
                'TRAVEL  - OVERSEAS PER DIEM',
                'TRAVEL -LOCAL ACCOMMODATION',
                'TRAVEL -LOCAL PER DIEM',
            ],
        },
        # Not checkable in the individual tab's rows: Company (include 1), GL Business Code (include 2), Type (include 4)
    },
    ('FY27 OPEX ALIV Business', 'CBL'): {
        "individual_tableau_workbook": 'ALIV Business.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
        },
        # Not checkable in the individual tab's rows: GL Bus Code Name (include 6), GL FS Category (include 1), Type (include 4)
    },
    ('FY27 OPEX ALIV Business', 'ALIV'): {
        "individual_tableau_workbook": 'ALIV Business.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4)
    },
    ('FY27 CBL OPEX Finance', 'CBL'): {
        "individual_tableau_workbook": 'Finance.twb',
        "individual_tableau_worksheet": 'Monthly Forecast Details',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
        },
        "individual_exclude": {
            'GL Sub Account Description': [
                'INTERNAL AUDIT - LOCAL FLIGHTS',
                'INTERNAL AUDIT - LOCAL VEHICLE RENTAL',
                'INTERNAL AUDIT - OVERSEAS ACCOMODATIONS',
                'INTERNAL AUDIT - OVERSEAS FLIGHTS',
                'INTERNAL AUDIT - OVERSEAS PER DIEM',
                'INTERNAL AUDIT - PROFESSIONAL FEES  - OTHER',
                'INTERNAL AUDIT - PROFESSIONAL FEES - AUDIT',
                'INTERNAL AUDIT -LOCAL ACCOMMODATION',
                'INTERNAL AUDIT -LOCAL PER DIEM',
                'INTERNAL AUDIT-OTHER MISC. COSTS',
                'INTERNAL AUDIT-PROFESSIONAL SUBSCRIPTIONS',
                'INTERNAL AUDIT-TEAM INCENTIVES',
                'INTERNAL AUDIT-TRAINING & DEVELOPMENT',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4)
    },
    ('FY27 CBL Gen & Admin Finance Opex', 'CBL'): {
        "individual_tableau_workbook": 'CBL Gen & Admin_Finance.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Account Mapping': [
                '',
                'ASSET IMPAIRMENT',
                'BAD DEBT EXPENSE',
                'GEN ADMIN',
                'INSURANCE',
                'NETWORK EXPENSES',
                'PAYABLES AND ACCRUALS',
                'PROFESSIONAL FEES',
                'REGULATORY',
                'RELATED PARTY PAYABLES',
            ],
            'GL Sub Account Description': [
                '',
                'ACCOUNTS PAYABLE-TRADE',
                'BAD DEBT EXPENSE - CONSUMER BAD DEBT PROVISION',
                'BAD DEBT EXPENSE - CORPORATE BAD DEBT PROVISION',
                'BAD DEBT EXPENSE - WRITE -OFFS OTHER',
                'GEN ADMIN -  SUBSCRIPTIONS',
                'GEN ADMIN - BOARD MEETING COSTS',
                'GEN ADMIN - DISASTER RECOVERY COSTS',
                'GEN ADMIN - OTHER',
                'GEN ADMIN - SHIPPING & COURIER',
                'INSURANCE - CYBER SECURITY',
                "INSURANCE - DIRECTORS' LIABILITY",
                'INSURANCE - GENERAL',
                'INSURANCE - MOTOR VEHICLE',
                'LEGAL SETTLEMENT PROVISIONS',
                'OTHER TAXES',
                'PROFESSIONAL FEES  - DIRECTORS FEES',
                'PROFESSIONAL FEES  - OTHER',
                'REGULATORY - BUSINESS LICENSE',
                'REGULATORY - COMMUNICATIONS FEE',
                'REGULATORY - FCC FEES',
                'REGULATORY - PROPERTY TAX',
                'REGULATORY - ROYALTY FEES',
                'REGULATORY - TRIBUNAL FEES',
                'REGULATORY - URCA FEES',
                'REGULATORY - WITHOLDING TAX',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 3)
    },
    ('FY27 Aliv Gen & Admin Finance Opex', 'ALIV Gen & Admin_Finance'): {
        "individual_tableau_workbook": 'ALIV Gen & Admin_Finance.twb',
        "individual_tableau_worksheet": 'Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Account Mapping': [
                '',
                'ASSET IMPAIRMENT',
                'BAD DEBT EXPENSE',
                'INSURANCE',
                'NETWORK EXPENSES',
                'PAYABLES AND ACCRUALS',
                'PROFESSIONAL FEES',
                'REGULATORY',
                'RELATED PARTY PAYABLES',
                'SPONSORSHIP & DONATIONS',
            ],
            'GL Sub Account Description': [
                '',
                'ACCOUNTS PAYABLE-TRADE',
                'BAD DEBT EXPENSE - CONSUMER BAD DEBT PROVISION',
                'BAD DEBT EXPENSE - CORPORATE BAD DEBT PROVISION',
                'BAD DEBT EXPENSE - OTHER',
                'BAD DEBT EXPENSE - WRITE -OFFS OTHER',
                'DUE TO CABLE BAHAMAS',
                'GEN ADMIN - GROUP RECHARGE MANAGED SERVICES',
                'INSURANCE - CYBER SECURITY',
                "INSURANCE - DIRECTORS' LIABILITY",
                'INSURANCE - GENERAL',
                'INSURANCE - MOTOR VEHICLE',
                'INVENTORY IMPAIRMENT',
                'NETWORK EXPENSES - GROUP RECHARGE FREEPORT CALL CE',
                'NETWORK EXPENSES - GROUP RECHARGE IFRS NETWORK ADJ',
                'PROFESSIONAL FEES  - DIRECTORS FEES',
                'PROFESSIONAL FEES  - OTHER',
                'REGULATORY - BUSINESS LICENSE',
                'REGULATORY - COMMUNICATIONS FEE',
                'REGULATORY - URCA FEES',
                'SPONSORSHIP-ART, MUSIC & FESTIVALS',
                'TRAVEL  - LOCAL FLIGHTS',
                'TRAVEL  - LOCAL VEHICLE RENTAL',
                'TRAVEL  - OVERSEAS ACCOMODATIONS',
                'TRAVEL  - OVERSEAS FLIGHTS',
                'TRAVEL  - OVERSEAS MEALS & MISC',
                'TRAVEL  - OVERSEAS PER DIEM',
                'TRAVEL -LOCAL ACCOMMODATION',
                'TRAVEL -LOCAL PER DIEM',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 2), Type (include 4)
    },
    ('FY27 OPEX INTERNAL AUDIT', 'ALIV'): {
        "individual_tableau_workbook": 'Finance.twb',
        "individual_tableau_worksheet": 'Monthly Forecast Details',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Sub Account Description': [
                'INTERNAL AUDIT - LOCAL FLIGHTS',
                'INTERNAL AUDIT - LOCAL VEHICLE RENTAL',
                'INTERNAL AUDIT - OVERSEAS ACCOMODATIONS',
                'INTERNAL AUDIT - OVERSEAS FLIGHTS',
                'INTERNAL AUDIT - OVERSEAS PER DIEM',
                'INTERNAL AUDIT - PROFESSIONAL FEES  - OTHER',
                'INTERNAL AUDIT - PROFESSIONAL FEES - AUDIT',
                'INTERNAL AUDIT -LOCAL ACCOMMODATION',
                'INTERNAL AUDIT -LOCAL PER DIEM',
                'INTERNAL AUDIT-OTHER MISC. COSTS',
                'INTERNAL AUDIT-PROFESSIONAL SUBSCRIPTIONS',
                'INTERNAL AUDIT-TEAM INCENTIVES',
                'INTERNAL AUDIT-TRAINING & DEVELOPMENT',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4)
    },
    ('FY27 OPEX INTERNAL AUDIT', 'CBL'): {
        "individual_tableau_workbook": 'Finance.twb',
        "individual_tableau_worksheet": 'Monthly Forecast Details',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Sub Account Description': [
                'INTERNAL AUDIT - LOCAL FLIGHTS',
                'INTERNAL AUDIT - LOCAL VEHICLE RENTAL',
                'INTERNAL AUDIT - OVERSEAS ACCOMODATIONS',
                'INTERNAL AUDIT - OVERSEAS FLIGHTS',
                'INTERNAL AUDIT - OVERSEAS PER DIEM',
                'INTERNAL AUDIT - PROFESSIONAL FEES  - OTHER',
                'INTERNAL AUDIT - PROFESSIONAL FEES - AUDIT',
                'INTERNAL AUDIT -LOCAL ACCOMMODATION',
                'INTERNAL AUDIT -LOCAL PER DIEM',
                'INTERNAL AUDIT-OTHER MISC. COSTS',
                'INTERNAL AUDIT-PROFESSIONAL SUBSCRIPTIONS',
                'INTERNAL AUDIT-TEAM INCENTIVES',
                'INTERNAL AUDIT-TRAINING & DEVELOPMENT',
            ],
        },
        # Not checkable in the individual tab's rows: GL FS Category (include 1), Type (include 4)
    },
    ('FY27 OPEX Media (ALIV BUSINESS)', 'FY27 MEDIA OPEX (ALIV BUSINESS)'): {
        "individual_tableau_workbook": 'Media.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL FS Category': [
                'CORPORATE SALES',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Media (MARKETING)', 'Media (MARKETING)'): {
        "individual_tableau_workbook": 'Media.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL FS Category': [
                'MARKETING',
            ],
        },
        # Not checkable in the individual tab's rows: Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Operations', 'ALIV'): {
        "individual_tableau_workbook": 'Operations.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 6), Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 OPEX Operations', 'CBL'): {
        "individual_tableau_workbook": 'Operations.twb',
        "individual_tableau_worksheet": 'GROUP Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 5), Type (include 4), Zero Filter (copy)_1422293069252046848 (include 1)
    },
    ('FY27 ALIV DIRECT COSTS COMMISSIONS SERVICE FEES', 'ALIV'): {
        "individual_tableau_workbook": 'Cable Bahamas Group Direct Costs_Finance.twb',
        "individual_tableau_worksheet": 'Direct Cost Forecast ',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Account Mapping': [
                'SERVICE FEES',
                'TOP UP & PLAN COMMISSION',
            ],
            'GL Sub Account Description': [
                'BANK & CREDIT CARD FEES',
                'COMMISSIONS - OTHER SERVICE COMMISSIONS & BANK CHA',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 ALIV DIRECT COSTS (CONSUMER SALES) COMMISSIONS', 'ALIV'): {
        "individual_tableau_workbook": 'Cable Bahamas Group Direct Costs_Consumer Sales.twb',
        "individual_tableau_worksheet": 'Direct Cost Forecast ',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Account Mapping': [
                'TOP UP & PLAN COMMISSION',
            ],
            'GL Sub Account Description': [
                'BANK & CREDIT CARD FEES',
                'COMMISSIONS - BTS RECHARGE',
                'COMMISSIONS - KIOSKS',
                'COMMISSIONS - RECHARGE DEALERS',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 ALIV DIRECT COSTS (FINANCE) COMMISSIONS SERVICE FEES', 'ALIV'): {
        "individual_tableau_workbook": 'Cable Bahamas Group Direct Costs_Finance.twb',
        "individual_tableau_worksheet": 'Direct Cost Forecast ',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Account Mapping': [
                'SERVICE FEES',
                'TOP UP & PLAN COMMISSION',
            ],
            'GL Sub Account Description': [
                'BANK & CREDIT CARD FEES',
                'COMMISSIONS - OTHER SERVICE COMMISSIONS & BANK CHA',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 CBL DTH COSTS', 'CBL'): {
        "individual_tableau_workbook": 'Group Direct Costs_Consumer Product Development.twb',
        "individual_tableau_worksheet": 'CBL Group Direct Cost Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Account Mapping': [
                '',
                'DTH',
                'INTERCONNECT COST',
                'PAYABLES AND ACCRUALS',
                'SMART HOME PRODUCTS COST',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 6), GL FS Category (include 24)
    },
    ('FY27 CBL DIRECT COSTS PBX SERVICE FEES', 'CBL'): {
        "individual_tableau_workbook": 'Group_Direct Costs_Aliv Business.twb',
        "individual_tableau_worksheet": 'Group Direct Cost Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Account Mapping': [
                'PBX',
                'SERVICE FEES',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 CBL DIRECT COSTS INTERNET BANDWIDTH (ENGINEERING)', 'CBL BANDWIDTH'): {
        "individual_tableau_workbook": 'Cable Bahamas Group Direct Costs_Engineering.twb',
        "individual_tableau_worksheet": 'Direct Costs - Interconnect Capacity Monthly Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Account Mapping': [
                'INTERCONNECT CAPACITY',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 7), GL FS Category (include 24)
    },
    ('FY27 CBL DIRECT COSTS INTERNET BANDWIDTH (ALIV BUSINESS)', 'CBL BANDWIDTH'): {
        "individual_tableau_workbook": 'Direct Costs_Internet Bandwidth_Aliv Business.twb',
        "individual_tableau_worksheet": 'Group Direct Cost Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Account Mapping': [
                'INTERCONNECT CAPACITY',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 1), GL FS Category (include 24)
    },
    ('FY27 CBL DIRECT COSTS CONTENT FEES', 'CBL'): {
        "individual_tableau_workbook": 'CBL Content Fees_Consumer Product Development.twb',
        "individual_tableau_worksheet": 'CBL Content Fees - Forecast (Excl Media)',
        "validation_status": "confirmed",
        "individual_include": {
            'GL Account Mapping': [
                'CONTENT FEES',
            ],
        },
        # Not checkable in the individual tab's rows: Company (include 1), GL FS Category (include 24)
    },
    ('FY27 CBL & ALIV DIRECT COSTS ROAMING INTERCONNECT OTHER', 'ALIV'): {
        "individual_tableau_workbook": 'Group Direct Costs_Consumer Product Development.twb',
        "individual_tableau_worksheet": 'ALIV Group Direct Cost Forecast ',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Account Mapping': [
                'DTH',
                'INTERCONNECT COST',
                'ROAMING COSTS',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 CBL & ALIV DIRECT COSTS ROAMING INTERCONNECT OTHER', 'CBL'): {
        "individual_tableau_workbook": 'Group Direct Costs_Consumer Product Development.twb',
        "individual_tableau_worksheet": 'CBL Group Direct Cost Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'CBL',
            ],
            'GL Account Mapping': [
                'DTH',
                'INTERCONNECT COST',
                'ROAMING COSTS',
                'SMART HOME PRODUCTS COST',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 ALIV DIRECT COSTS PUSH TO TALK', 'ALIV'): {
        "individual_tableau_workbook": 'Group_Direct Costs_Aliv Business.twb',
        "individual_tableau_worksheet": 'Group Direct Cost Forecast',
        "validation_status": "confirmed",
        "individual_include": {
            'Company': [
                'ALIV',
            ],
            'GL Account Mapping': [
                'PUSH TO TALK COMMISSION',
            ],
        },
        # Not checkable in the individual tab's rows: GL Business Code (include 8), GL FS Category (include 24)
    },
    ('FY27 INTERCOMPANY REVENUE', 'CBL'): {
        # Paired with the MSA Revenue workbook (October 2026). That workbook
        # shows Intercompany (IC) revenue for MSA REVENUE only, so other IC
        # lines on the sheet are not in its total.
        "individual_tableau_workbook": 'Cable Bahamas Group_MSA REVENUE.twb',
        "individual_tableau_worksheet": 'CBL COMMERCIAL REVENUE DETAILS - Monthly Forecast',
        "validation_status": "confirmed",
    },
}


def apply(routes) -> None:
    """Update the routes in place from SELECTIONS (by forecast workbook + sheet)."""
    for route in routes:
        # Individual tab names were tidied in Tableau (no leading, trailing or
        # double spaces); the app matches them ignoring spaces either way.
        if route.get("individual_tableau_worksheet"):
            route["individual_tableau_worksheet"] = " ".join(route["individual_tableau_worksheet"].split())
        stem = route["workbook"].rsplit("_", 1)[0]
        change = SELECTIONS.get((stem, route["excel_sheet"]))
        if change:
            route.update(change)
            route.pop("validation_filter_note", None)
