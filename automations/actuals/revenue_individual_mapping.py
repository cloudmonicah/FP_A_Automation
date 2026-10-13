"""Initial confirmed ALIV Consumer individual Revenue routes.
Only enable a row after its GL FS Type is confirmed in a captured crosstab.
"""
CONSUMER_TWB = r"E:\Workbooks\Revenue\Cable Bahamas Group_CONSUMER REVENUE.twb"
CONSUMER_SHEET = "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD"

ALIV_CONSUMER_ROUTES = [
    {"company":"ALIV", "fs_line":"Consumer Prepaid", "gl_fs_type":"PAY IN ADVANCE"},
    {"company":"ALIV", "fs_line":"Consumer Postpaid", "gl_fs_type":"POSTPAID"},
    {"company":"ALIV", "fs_line":"Visitor Roaming", "gl_fs_type":"VISITOR ROAMING"},
    {"company":"ALIV", "fs_line":"Interconnect", "gl_fs_type":"INTERCONNECT REVENUE"},
]
# ALIV Other, Corporate and Service Revenues are intentionally excluded until
# their detail filters/control totals are confirmed.
