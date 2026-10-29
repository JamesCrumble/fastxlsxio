import asyncio
import sys
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import UTC, date, datetime
from datetime import time as dttime
from itertools import batched

from fastxlsxio import ColTypeHint, XIOWOptions, XIOWWorkbook
from xlsxwriter.workbook import Workbook

NUM_ROWS = 300_000
STATIC_COLS = 7
DYNAMIC_COLS = 10
NUM_COLS = STATIC_COLS + DYNAMIC_COLS

SHEETS = ("Sheet 3", "Sheet 1", "Sheet 2")
TPOOL = ThreadPoolExecutor(max_workers=len(SHEETS))


def data_():
    for _ in range(NUM_ROWS):
        ret = {
            "uuid": uuid.uuid4(),
            "dt": datetime.now(),
            "d": datetime.now().date(),
            "t": datetime.now().time(),
            "dtz": datetime.now(tz=UTC),
            "dz": datetime.now(tz=UTC).date(),
            "tz": datetime.now(tz=UTC).time(),
        }
        ret.update({f"key: {v}": f"{v * 1000}" if v % 2 == 0 else v / 1 for v in range(DYNAMIC_COLS)})
        yield ret


def data_by_batch():
    return (batch for batch in batched(data_(), 100))


data = list(data_by_batch())


def fastxlsxio_build(xlsx: str, multithreaded: bool = False):
    options = XIOWOptions(constant_memory=True, cache_typehints_write_optimization=True)
    excel = XIOWWorkbook(xlsx, options=options)
    excel.add_format({"num_format": "###.0"}, bind_to_datatype=ColTypeHint.Int)
    excel.add_format({"num_format": "###.123"}, bind_to_datatype=ColTypeHint.Float)
    excel.add_format({"num_format": "dd.mm.yyyy"}, bind_to_datatype=ColTypeHint.Date)
    excel.add_format({"num_format": "dd.mm.yyyy hh:mm:ss"}, bind_to_datatype=ColTypeHint.DateTime)
    excel.add_format({"num_format": "hh:mm:ss"}, bind_to_datatype=ColTypeHint.Time)

    tasks = list()  # noqa
    rows_written: int = 0
    for sheet_name in SHEETS:
        _ = excel.add_worksheet(sheet_name, options)

        def _write_sheet(sheet_name: str) -> int:
            st = time.monotonic()
            acc = 0
            sheet = excel.get_by_name(sheet_name)

            for i, batch in enumerate(data):
                sheet.write_rows(i * len(batch), 0, batch)
                acc += len(batch)

            print(f"fastxlsxio: {sheet.name} sheet {NUM_ROWS}x{NUM_COLS} rows written for {time.monotonic() - st:.2f}s")
            return acc

        if not multithreaded:
            rows_written += _write_sheet(sheet_name)
        else:
            tasks.append((_write_sheet, sheet_name))

    if tasks:
        for res in as_completed(TPOOL.submit(*task) for task in tasks):
            rows_written += res.result()

    print(f"fastxlsxio: {rows_written}x{NUM_COLS} rows written")

    sst = time.monotonic()
    excel.save()
    print(f"fastxlsxio: xlsx saved for {time.monotonic() - sst:.2f}s. Note: fully detached from gil save")


def xlsxwriter_build(xlsx: str, multithreaded: bool = False):
    excel = Workbook(xlsx, options={"constant_memory": True, "remove_timezone": True})
    intfmt = excel.add_format({"num_format": "###.0"})
    floatfmt = excel.add_format({"num_format": "###.123"})
    datefmt = excel.add_format({"num_format": "dd.mm.yyyy"})
    datetimefmt = excel.add_format({"num_format": "dd.mm.yyyy hh:mm:ss"})
    timefmt = excel.add_format({"num_format": "hh:mm:ss"})

    tasks = list()  # noqa
    rows_written: int = 0
    for sheet_name in SHEETS:
        _ = excel.add_worksheet(sheet_name)

        def _write_sheet(sheet_name: str) -> int:
            st = time.monotonic()
            sheet = excel.get_worksheet_by_name(sheet_name)

            row_: int = 0
            for batch in data:
                for row in batch:
                    for j, value in enumerate(row.values()):
                        if isinstance(value, uuid.UUID):
                            value = str(value)

                        if type(value) is int:
                            sheet.write_number(row_, j, value, intfmt)
                        elif type(value) is float:
                            sheet.write_number(row_, j, value, floatfmt)
                        elif type(value) is date:
                            sheet.write_datetime(row_, j, value, datefmt)
                        elif type(value) is datetime:
                            sheet.write_datetime(row_, j, value, datetimefmt)
                        elif type(value) is dttime:
                            sheet.write_datetime(row_, j, value, timefmt)
                        else:
                            sheet.write(row_, j, value)

                row_ += 1

            print(f"xlsxwriter: {sheet.name} sheet {NUM_ROWS}x{NUM_COLS} rows written for {time.monotonic() - st:.2f}s")
            return row_

        if not multithreaded:
            rows_written += _write_sheet(sheet_name)
        else:
            tasks.append((_write_sheet, sheet_name))

    if tasks:
        for res in as_completed(TPOOL.submit(*task) for task in tasks):
            rows_written += res.result()

    print(f"xlsxwriter: {rows_written}x{NUM_COLS} rows written")

    sst = time.monotonic()
    excel.close()
    print(f"xlsxwriter: xlsx saved for {time.monotonic() - sst:.2f}s")


async def main():
    print("Start fastxlsxio")
    st = time.monotonic()
    fastxlsxio_build("fastxlsxio_build.xlsx", True)
    print(f"write for {time.monotonic() - st:.2f} fastxlsxio")

    print("Start xlsxwriter")
    st = time.monotonic()
    xlsxwriter_build("xlsxwriter_build.xlsx", True)
    print(f"write for {time.monotonic() - st:.2f} xlsxwriter")


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
