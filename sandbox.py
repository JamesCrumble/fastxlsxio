import asyncio
import sys
import time
from itertools import batched

from fastxlsxio import XIOWorkbook
from xlsxwriter.workbook import Workbook


def data_():
    for i in range(1000000):
        yield {"k": i, "v": i, "q": i}


def data_by_batch():
    return (batch for batch in batched(data_(), 100))


def fastxlsxio_build(xlsx: str):
    sheets = ("Sheet 3", "Sheet 1", "Sheet 2")

    excel = XIOWorkbook()
    rows_written: int = 0
    for sheet in sheets:
        st = time.monotonic()
        data = data_by_batch()
        sheet = excel.add_worksheet(sheet, constant_memory=True)

        # cursor: int = 0
        # for batch in data:
        #     for row in batch:
        #         sheet.write_row(cursor, 0, row)
        #         cursor += 1
        #         # for j, value in enumerate(row.values()):
        #         #     sheet.write_cell(i, j, value)

        # rows_written += cursor

        for i, batch in enumerate(data):
            sheet.write_rows(i * len(batch), 0, batch)
            rows_written += len(batch)
        print(f"fastxlsxio {sheet.name} sheet {rows_written} rows written for {time.monotonic() - st:.2f}")
    print(f"fastxlsxio {rows_written} rows written")

    excel.save(xlsx)


def xlsxwriter_build(xlsx: str):
    sheets = ("Sheet 3", "Sheet 1", "Sheet 2")

    excel = Workbook(xlsx, options={"constant_memory": True})
    rows_written: int = 0
    for sheet in sheets:
        st = time.monotonic()
        data = data_by_batch()
        sheet = excel.add_worksheet(sheet)

        cursor: int = 0
        for batch in data:
            for row in batch:
                for j, value in enumerate(row.values()):
                    sheet.write(cursor, j, value)
                cursor += 1

        rows_written += cursor
        print(f"xlsxwriter {sheet} sheet {rows_written} rows written for {time.monotonic() - st:.2f}")
    print(f"xlsxwriter {rows_written} rows written")

    excel.close()


async def main():
    # service = Q2XLSXService()

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
