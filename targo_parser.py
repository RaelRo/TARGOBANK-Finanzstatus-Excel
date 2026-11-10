"""Convert TARGOBANK "Finanzstatus" PDF statements into one Excel workbook.

Usage:
    python targobank_parser.py [FOLDER] [-o OUTPUT.xlsx]
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator

import pandas as pd
import pdfplumber

logger = logging.getLogger("targobank")

DEFAULT_OUTPUT = "Finanzstatus01_11.xlsx"
COLUMNS = ["Datum", "Buchungstext", "Guthaben", "Ausgaben", "Einnahmen"]

HEADER_MARKER = "F I N A N Z S T A T U S vom"
FOOTER_MARKER = "TARGOBANK AG I Vorstand:"
OPENING_BALANCE = "ANFANGSSALDO"
CLOSING_BALANCE = "ENDSALDO"

# A posting starts with DD.MM or DD.MM. followed by a space and more text.
DATE_LINE = re.compile(r"^(0[1-9]|[12][0-9]|3[01])\.(0[1-9]|1[0-2])\.? .+$")
YEAR_AT_END = re.compile(r"(\d{4})$")
GERMAN_NUMBER = re.compile(r"-?\d[\d.]*(?:,\d+)?")

# Informational lines that merely look like postings (they start with a date).
# Matched as substrings: against the amount column, or the first word of the text.
IGNORED_AMOUNT_MARKERS = (
    "+++",
    "SOLLZINSSÄTZE",
    "DISPOSITIONSKREDITE",
    "FürdieEinhaltunggesetzlicher",
    "KARTENLIMIT",
    "Änderung",
)
IGNORED_TEXT_MARKERS = (
    "RESERV",
    "VORAUSDATIERTE",
    "ÄNDERUNGDAUERAUFTRAG",
    "SOLLZINSSÄTZE:",
    "Änderung",
)


class StatementParseError(Exception):
    """The PDF does not look like a statement we can parse."""


@dataclass(frozen=True)
class Transaction:
    date: str
    text: str
    balance: float
    expense: float | None = None
    income: float | None = None

    def as_row(self) -> dict[str, object]:
        return {
            "Datum": self.date,
            "Buchungstext": self.text,
            "Guthaben": self.balance,
            "Ausgaben": self.expense,
            "Einnahmen": self.income,
        }


@dataclass
class _Pending:
    """A posting whose (possibly multi-line) text is not complete yet."""

    date: str
    balance: float
    amount: float | None  # None for the opening balance
    text: list[str]


def parse_german_number(token: str) -> float:
    """'1.234,56' -> 1234.56 (thousands separator '.', decimal comma)."""
    if not GERMAN_NUMBER.fullmatch(token):
        raise ValueError(f"not a German-formatted number: {token!r}")
    return float(token.replace(".", "").replace(",", "."))


def _clean_underscores(text: str) -> str:
    return re.sub(r"_{2,}", " ", text).replace("_", "")


def _is_informational(amount_token: str, text: list[str]) -> bool:
    if any(marker in amount_token for marker in IGNORED_AMOUNT_MARKERS):
        return True
    return bool(text) and any(marker in text[0] for marker in IGNORED_TEXT_MARKERS)


class _StatementParser:
    """Line-by-line state machine.

    A posting stays *pending* until the next posting, the page footer or the
    closing balance shows that its text (which can span several lines) is
    complete.
    """

    def __init__(self) -> None:
        self.transactions: list[Transaction] = []
        self.finished = False
        self._year: str | None = None
        self._pending: _Pending | None = None

    def feed(self, raw_line: str) -> None:
        line = raw_line.strip()
        if not line:
            return
        if HEADER_MARKER in line:
            self._year = self._read_year(line)
            return

        line = _clean_underscores(line)
        if FOOTER_MARKER in line:
            self._flush()  # end of page: close the last posting
        elif DATE_LINE.match(line):
            self._start_posting(line)
        elif self._pending is not None:
            self._pending.text.append(line)  # continuation line

    # -- helpers -----------------------------------------------------------

    def _start_posting(self, line: str) -> None:
        self._flush()

        tokens = line.split()
        # Layout: <date> [<value date>] <text...> <amount> <balance> ["-"]
        negative = tokens[-1] == "-"
        trailing = 3 if negative else 2
        if len(tokens) <= trailing:
            return

        date = tokens[0]
        amount_token = tokens[-trailing]
        text = tokens[2:-trailing]
        is_balance_line = amount_token in (OPENING_BALANCE, CLOSING_BALANCE)

        if not is_balance_line and _is_informational(amount_token, text):
            return

        balance_token = tokens[-2] if negative else tokens[-1]
        amount: float | None = None
        try:
            balance = parse_german_number(("-" if negative else "") + balance_token)
            if not is_balance_line:
                amount = parse_german_number(amount_token)
        except ValueError:
            logger.warning("Skipping unparseable line: %r", line)
            return

        if amount_token == OPENING_BALANCE:
            self._pending = _Pending(date, balance, None, [OPENING_BALANCE])
        elif amount_token == CLOSING_BALANCE:
            self.transactions.append(
                Transaction(self._format_date(date), CLOSING_BALANCE, balance)
            )
            self.finished = True
        else:
            self._pending = _Pending(date, balance, amount, text)

    def _flush(self) -> None:
        pending, self._pending = self._pending, None
        if pending is not None:
            self.transactions.append(self._build(pending))

    def _build(self, pending: _Pending) -> Transaction:
        date = self._format_date(pending.date)
        text = " ".join(pending.text)
        previous = self.transactions[-1] if self.transactions else None

        if previous is None or pending.amount is None:
            return Transaction(date, text, pending.balance)

        # The statement prints unsigned amounts, so the direction is inferred
        # from the balance change (a zero-amount posting fills both columns).
        expense = pending.amount if previous.balance >= pending.balance else None
        income = pending.amount if previous.balance <= pending.balance else None
        return Transaction(date, text, pending.balance, expense, income)

    def _format_date(self, day_month: str) -> str:
        if self._year is None:
            raise StatementParseError("posting found before the FINANZSTATUS header")
        return f"{day_month.rstrip('.')}.{self._year}"

    @staticmethod
    def _read_year(header: str) -> str:
        match = YEAR_AT_END.search(header)
        if match is None:
            raise StatementParseError(f"cannot read year from header: {header!r}")
        return match.group(1)


def parse_lines(lines: Iterable[str]) -> list[Transaction]:
    """Parse the text lines of one statement (stops at the closing balance)."""
    parser = _StatementParser()
    for line in lines:
        parser.feed(line)
        if parser.finished:
            return parser.transactions
    raise StatementParseError(f"no {CLOSING_BALANCE} line found")


def _pdf_lines(pdf: pdfplumber.PDF) -> Iterator[str]:
    for page in pdf.pages:
        yield from (page.extract_text() or "").split("\n")


def parse_statement(pdf_path: Path) -> pd.DataFrame:
    with pdfplumber.open(pdf_path) as pdf:
        transactions = parse_lines(_pdf_lines(pdf))
    return pd.DataFrame([t.as_row() for t in transactions], columns=COLUMNS)


def find_pdfs(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"
    )


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert TARGOBANK statement PDFs into one Excel file."
    )
    parser.add_argument(
        "folder",
        nargs="?",
        type=Path,
        help="folder containing the PDF statements (prompted for if omitted)",
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=Path(DEFAULT_OUTPUT), help="Excel output file"
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    args = _parse_args(argv)

    folder = Path(args.folder or input("File path to folder where to read the PDF files: "))
    folder = folder.expanduser().resolve()
    if not folder.is_dir():
        logger.error("Not a directory: %s", folder)
        return 1

    frames: list[pd.DataFrame] = []
    failed: list[str] = []
    for pdf_path in find_pdfs(folder):
        logger.info("Parsing %s", pdf_path.name)
        try:
            frames.append(parse_statement(pdf_path))
        except StatementParseError as exc:
            logger.error("%s: %s (file skipped)", pdf_path.name, exc)
            failed.append(pdf_path.name)

    if not frames:
        logger.error("No statements could be parsed in %s", folder)
        return 1

    pd.concat(frames, ignore_index=True).to_excel(args.output, index=False)
    logger.info("Wrote %d statement(s) to %s", len(frames), args.output)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())