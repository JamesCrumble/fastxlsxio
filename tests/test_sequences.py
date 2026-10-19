import re
import tempfile
import zipfile
from datetime import UTC, date, datetime
from pathlib import Path

import pytest
from fastxlsxio import XIOWOptions, XIOWWorkbook


def written_strings(value) -> str:
    path = Path(tempfile.mkdtemp()) / "t.xlsx"
    workbook = XIOWWorkbook(str(path), XIOWOptions(constant_memory=True))
    workbook.add_worksheet("s").write_rows(0, 0, [(value,)])
    workbook.save()
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    match = re.search(r"<t[^>]*>(.*?)</t>", xml)
    return match.group(1) if match else ""


@pytest.mark.parametrize(
    ("value", "expected"),
    (
        (["a", "b", "c"], "a, b, c"),
        (("a", 1, 2.5), "a, 1, 2.5"),
        ([1, None, 3], "1, , 3"),
        ([True, False], "Да, Нет"),
        (["only"], "only"),
    ),
)
def test_sequence_is_joined_by_comma(value, expected: str):
    assert written_strings(value) == expected


@pytest.mark.parametrize(
    "value",
    (
        [datetime(2026, 1, 2, 3, 4, 5), datetime(2025, 12, 31, 23, 59, 59)],
        (datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC), datetime(2025, 12, 31, 23, 59, 59, tzinfo=UTC)),
    ),
)
def test_datetimes_are_joined_by_semicolon_in_russian_format(value):
    assert written_strings(value) == "02.01.2026, 03:04:05; 31.12.2025, 23:59:59"


def test_datetime_sequence_with_other_items_is_joined_by_comma():
    assert written_strings([datetime(2026, 1, 2, 3, 4, 5), "x"]) == "2026-01-02 03:04:05, x"
    assert written_strings([datetime(2026, 1, 2, 3, 4, 5), None]) == "2026-01-02 03:04:05, "
    assert written_strings([date(2026, 1, 2), date(2026, 1, 3)]) == "2026-01-02, 2026-01-03"
