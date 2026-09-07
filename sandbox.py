import asyncio
import sys
import time
from itertools import batched

from fastxlsxio import XIOWOptions, XIOWWorkbook
from xlsxwriter.workbook import Workbook

NUM_ROWS = 300_000
NUM_COLS = 20


def data_():
    for _ in range(NUM_ROWS):
        yield {v: v * 100 for v in range(NUM_COLS)}


def data_by_batch():
    return (batch for batch in batched(data_(), 100))


data = list(data_by_batch())


def fastxlsxio_build(xlsx: str):
    sheets = ("Sheet 3", "Sheet 1", "Sheet 2")

    excel = XIOWWorkbook()
    print(excel.options)
    format_ = excel.add_format({"num_format": "0"})
    rows_written: int = 0
    for sheet in sheets:
        st = time.monotonic()
        sheet = excel.add_worksheet(sheet)

        acc: int = 0
        for batch in data:
            for row in batch:
                # sheet.write_row(acc, 0, row)
                for j, value in enumerate(row.values()):
                    sheet.write_cell(acc, j, value, format_)
                acc += 1

        # for i, batch in enumerate(data):
        #     sheet.write_rows(i * len(batch), 0, batch)
        #     rows_written += len(batch)

        rows_written += acc

        print(f"fastxlsxio {sheet.name} sheet {NUM_ROWS}x{NUM_COLS} rows written for {time.monotonic() - st:.2f}")
    print(f"fastxlsxio {rows_written}x{NUM_COLS} rows written")

    excel.save(xlsx)


def xlsxwriter_build(xlsx: str):
    sheets = ("Sheet 3", "Sheet 1", "Sheet 2")

    excel = Workbook(xlsx, options={"constant_memory": True})
    rows_written: int = 0
    for sheet in sheets:
        st = time.monotonic()
        sheet = excel.add_worksheet(sheet)

        cursor: int = 0
        for batch in data:
            for row in batch:
                for j, value in enumerate(row.values()):
                    sheet.write(cursor, j, value)
                cursor += 1

        rows_written += cursor
        print(f"xlsxwriter {sheet} sheet {NUM_ROWS}x{NUM_COLS} rows written for {time.monotonic() - st:.2f}")
    print(f"xlsxwriter {rows_written}x{NUM_COLS} rows written")

    excel.close()


async def main():
    print("Start fastxlsxio")
    st = time.monotonic()
    fastxlsxio_build("fastxlsxio_build.xlsx")
    print(f"write for {time.monotonic() - st:.2f} fastxlsxio")

    print("Start xlsxwriter")
    st = time.monotonic()
    xlsxwriter_build("xlsxwriter_build.xlsx")
    print(f"write for {time.monotonic() - st:.2f} xlsxwriter")


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
