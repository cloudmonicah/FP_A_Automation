"""Direct Cost mapping data. Import from validation_config.py; do not replace OPEX_MAPPINGS."""
AGG_SUM = 'SUM'
AGG_TOTAL = 'TOTAL'

DIRECT_COST_MAPPINGS = [
    dict(company='ALIV', fs_line='Interconnect', tableau_categories=['INTERCONNECT COST'], aggregation=AGG_SUM),
    dict(company='ALIV', fs_line='Reseller Commissions', tableau_categories=['PUSH TO TALK COMMISSION','TOP UP & PLAN COMMISSION'], aggregation=AGG_SUM),
    dict(company='ALIV', fs_line='Roaming Service Costs', tableau_categories=['ROAMING COSTS'], aggregation=AGG_SUM),
    dict(company='ALIV', fs_line='Service Fees', tableau_categories=['SERVICE FEES'], aggregation=AGG_SUM),
    dict(company='ALIV', fs_line='Direct Cost', tableau_categories=[], aggregation=AGG_TOTAL),
    dict(company='CBL', fs_line='Programming fees', tableau_categories=['CONTENT FEES'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='Hardware/PBX', tableau_categories=['PBX'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='Voice termination costs', tableau_categories=['INTERCONNECT COST'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='Internet Bandwidth', tableau_categories=['INTERCONNECT CAPACITY'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='Cloud services COS', tableau_categories=['SERVICE FEES'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='DTH cost of sales', tableau_categories=['DTH'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='Smart Home Products COS', tableau_categories=['SMART HOME PRODUCTS COST'], aggregation=AGG_SUM),
    dict(company='CBL', fs_line='Ocean Wifi costs', tableau_categories=[], aggregation='PENDING', note='Tableau category unconfirmed; leave blank'),
    dict(company='CBL', fs_line='Cost of Goods Sold', tableau_categories=[], aggregation=AGG_TOTAL),
]

# Individual workbook routes are CANDIDATES until GL-filter and company reconciliation.
# None of these routes should be treated as validated automatically.
INDIVIDUAL_DIRECT_COST_MAPPINGS = [
    dict(company='CBL', fs_line='Programming fees', individual_tableau_workbook=r'E:\Workbooks\Direct Cost\CBL Content Fees.twb', individual_tableau_worksheet='CBL Content Fees Summary', individual_actuals_enabled=False, note='Content Fees = Programming Fees; confirm full category and month/YTD measures'),
    dict(company='ALIV', fs_line='Interconnect', individual_tableau_workbook=r'E:\Workbooks\Direct Cost\Cable Bahamas Group Direct Costs_Engineering.twb', individual_tableau_worksheet='Group Direct Costs Year to Date Summary', individual_actuals_enabled=False, note='Candidate only: Engineering view may include multiple FS lines'),
    dict(company='ALIV', fs_line='Reseller Commissions', individual_tableau_workbook=r'E:\Workbooks\Direct Cost\Group Direct Costs_Consumer Product Development.twb', individual_tableau_worksheet='Group Direct Costs Year to Date Summary', individual_actuals_enabled=False, note='Candidate only: must confirm GL and ALIV filters'),
    dict(company='ALIV', fs_line='Roaming Service Costs', individual_tableau_workbook=r'E:\Workbooks\Direct Cost\Group_Direct Costs_Aliv Business.twb', individual_tableau_worksheet='Group Direct Costs Year to Date Summary', individual_actuals_enabled=False, note='Candidate only: must confirm GL and ALIV filters'),
    dict(company='ALIV', fs_line='Service Fees', individual_tableau_workbook=r'E:\Workbooks\Direct Cost\Cable Bahamas Group Direct Costs_Finance.twb', individual_tableau_worksheet='Group Direct Costs Year to Date Summary', individual_actuals_enabled=False, note='Candidate only: must confirm GL and ALIV filters'),
]
