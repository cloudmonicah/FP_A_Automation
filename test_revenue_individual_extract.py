from datetime import date
from decimal import Decimal
from automations.actuals.revenue_individual_extract import extract_aliv_consumer_tsv
text="\t\tJuly 2026\tAugust 2026\tGrand Total\nCompany\tGL FS Type\tActual\tActual\tActual\nALIV\tPAY IN ADVANCE\t100\t200\t300\nALIV\tPAY IN ADVANCE\t50\t60\t110\nALIV\tPOSTPAID\t10\t20\t30\nALIV\tVISITOR ROAMING\t1\t2\t3\nALIV\tINTERCONNECT REVENUE\t4\t5\t9\n"
r=extract_aliv_consumer_tsv(text,date(2026,8,31))
assert r["Consumer Prepaid"]["monthly"]==Decimal(260)
assert r["Consumer Prepaid"]["ytd"]==Decimal(410)
assert r["Consumer Prepaid"]["rows"]==2
print("PASS: ALIV Consumer GL FS Type summation and period selection")
