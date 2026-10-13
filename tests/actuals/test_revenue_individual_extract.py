import unittest
from datetime import date
from decimal import Decimal

from automations.actuals.revenue_individual_extract import (
    amount,
    extract_aliv_consumer_tsv,
)


class TestRevenueIndividualExtract(unittest.TestCase):

    def setUp(self):
        self.closed_period = date(2026, 8, 31)

        self.tsv = (
            "Company\tGL FS Type\tJuly 2026\t"
            "August 2026\tGrand Total\n"
            "\t\tActual\tActual\tActual\n"
            "ALIV\tPAY IN ADVANCE\t100\t200\t300\n"
            "ALIV\tPOSTPAID\t400\t500\t900\n"
            "ALIV\tVISITOR ROAMING\t50\t75\t125\n"
            "ALIV\tINTERCONNECT REVENUE\t20\t30\t50\n"
        )

    def test_amount_parsing(self):
        self.assertEqual(amount("1,250.50"), Decimal("1250.50"))
        self.assertEqual(amount("(250.00)"), Decimal("-250.00"))
        self.assertIsNone(amount(""))

    def test_monthly_and_ytd(self):
        result = extract_aliv_consumer_tsv(
            self.tsv,
            self.closed_period,
        )

        expected = {
            "Consumer Prepaid": ("200", "300"),
            "Consumer Postpaid": ("500", "900"),
            "Visitor Roaming": ("75", "125"),
            "Interconnect": ("30", "50"),
        }

        for category, (monthly, ytd) in expected.items():
            with self.subTest(category=category):
                actual = result[category]

                self.assertEqual(
                    actual["monthly"],
                    Decimal(monthly),
                )
                self.assertEqual(
                    actual["ytd"],
                    Decimal(ytd),
                )
                self.assertEqual(
                    actual["status"],
                    "EXTRACTED - NOT VALIDATED",
                )

    def test_missing_revenue_category(self):
        tsv = (
            "Company\tGL FS Type\tAugust 2026 ACTUAL\t"
            "ACTUAL YTD\n"
            "ALIV\tPAY IN ADVANCE\t200\t300\n"
        )

        result = extract_aliv_consumer_tsv(
            tsv,
            self.closed_period,
        )

        self.assertEqual(
            result["Consumer Postpaid"]["status"],
            "REVIEW REQUIRED",
        )

        self.assertIsNone(
            result["Consumer Postpaid"]["monthly"],
        )

    def test_two_row_header_without_detail(self):
        tsv = "Company\tGL FS Type\tAugust 2026\tGrand Total\n" "\t\tActual\tActual\n"

        with self.assertRaises(ValueError):
            extract_aliv_consumer_tsv(
                tsv,
                self.closed_period,
            )


if __name__ == "__main__":
    unittest.main()
