"""Calculate simple Excel formulas on one sheet without Excel.

openpyxl saves formulas but not their results, so a workbook the app just
wrote has no value in a SUM cell until Excel opens it. This evaluates the
formulas forecast sheets use for totals: numbers, cell references, ranges,
SUM(), + - * / and brackets, on the same sheet. Anything else raises
UnsupportedFormula so the caller can say so.
"""

from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string, range_boundaries


class UnsupportedFormula(ValueError):
    """A formula outside the small set this module understands."""


_TOKEN = re.compile(
    r"\s*(?:(?P<range>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)"
    r"|(?P<cell>\$?[A-Z]{1,3}\$?\d+)"
    r"|(?P<number>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)"
    r"|(?P<func>[A-Z][A-Z0-9.]*)\("
    r"|(?P<op>[-+*/(),]))"
)


def _tokens(formula: str) -> list[tuple[str, str]]:
    text = formula.lstrip("=").upper()
    position, out = 0, []
    while position < len(text):
        if text[position].isspace():
            position += 1
            continue
        match = _TOKEN.match(text, position)
        if not match:
            raise UnsupportedFormula(formula)
        kind = match.lastgroup
        out.append((kind, match.group(kind)))
        position = match.end()
    return out


class SheetCalculator:
    """Values of cells on one worksheet, calculating formulas as needed."""

    def __init__(self, worksheet):
        self.ws = worksheet
        self._cache: dict[tuple[int, int], float] = {}
        self._busy: set[tuple[int, int]] = set()

    def value(self, row: int, col: int) -> float:
        key = (row, col)
        if key in self._cache:
            return self._cache[key]
        if key in self._busy:
            raise UnsupportedFormula(f"circular reference at {self.ws.cell(row, col).coordinate}")
        raw = self.ws.cell(row, col).value
        self._busy.add(key)
        try:
            if isinstance(raw, str) and raw.startswith("="):
                result = self.formula(raw)
            elif isinstance(raw, bool) or raw is None or isinstance(raw, str):
                result = 0.0   # Excel's SUM ignores text and blanks
            elif hasattr(raw, "year"):
                # A number formatted as a date (e.g. 67,500 showing as "Oct-84"):
                # Excel adds the number underneath, so use that.
                from openpyxl.utils.datetime import to_excel

                result = float(to_excel(raw))
            else:
                result = float(raw)
        finally:
            self._busy.discard(key)
        self._cache[key] = result
        return result

    def formula(self, text: str) -> float:
        return _Parser(self, text).run()

    def cell_ref(self, text: str) -> float:
        match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", text)
        return self.value(int(match.group(2)), column_index_from_string(match.group(1)))


class _Parser:
    """Recursive-descent parser for one formula (its own position)."""

    def __init__(self, calculator: SheetCalculator, text: str):
        self.calc = calculator
        self.text = text
        self.tokens = _tokens(text)
        self.at = 0

    def run(self) -> float:
        result = self._expression()
        if self.at != len(self.tokens):
            raise UnsupportedFormula(self.text)
        return result

    def _peek(self):
        return self.tokens[self.at] if self.at < len(self.tokens) else (None, None)

    def _take(self):
        token = self._peek()
        self.at += 1
        return token

    def _expression(self) -> float:
        total = self._term()
        while self._peek() in (("op", "+"), ("op", "-")):
            _, op = self._take()
            total = total + self._term() if op == "+" else total - self._term()
        return total

    def _term(self) -> float:
        total = self._factor()
        while self._peek() in (("op", "*"), ("op", "/")):
            _, op = self._take()
            right = self._factor()
            total = total * right if op == "*" else total / right
        return total

    def _factor(self) -> float:
        kind, text = self._take()
        if (kind, text) == ("op", "-"):
            return -self._factor()
        if (kind, text) == ("op", "+"):
            return self._factor()
        if kind == "number":
            return float(text)
        if kind == "cell":
            return self.calc.cell_ref(text)
        if (kind, text) == ("op", "("):
            value = self._expression()
            if self._take() != ("op", ")"):
                raise UnsupportedFormula("unbalanced brackets")
            return value
        if kind == "func" and text == "SUM":
            return self._sum()
        raise UnsupportedFormula(text or "end of formula")

    def _sum(self) -> float:
        total = 0.0
        while True:
            kind, text = self._peek()
            if kind == "range":
                self._take()
                min_col, min_row, max_col, max_row = range_boundaries(text.replace("$", ""))
                total += sum(self.calc.value(r, c) for r in range(min_row, max_row + 1)
                             for c in range(min_col, max_col + 1))
            else:
                total += self._expression()
            kind, text = self._take()
            if (kind, text) == ("op", ")"):
                return total
            if (kind, text) != ("op", ","):
                raise UnsupportedFormula("SUM arguments")
