import asyncio
import sys
import time
from itertools import batched

from fastxlsxio import ColTypeHint, XIOWOptions, XIOWWorkbook
from xlsxwriter.workbook import Workbook

NUM_ROWS = 300_000
NUM_COLS = 2


def data_():
    for _ in range(NUM_ROWS):
        yield {f"key: {v}": f"{v * 1000}" if v % 2 == 0 else v / 1 for v in range(NUM_COLS)}


def data_by_batch():
    return (batch for batch in batched(data_(), 100))


data = list(data_by_batch())


def fastxlsxio_build(xlsx: str):
    sheets = ("Sheet 3", "Sheet 1", "Sheet 2")

    excel = XIOWWorkbook(xlsx)
    excel.add_format({"num_format": "###.0"}, bind_to_datatype=ColTypeHint.Int)
    excel.add_format({"num_format": "###.123"}, bind_to_datatype=ColTypeHint.Float)
    rows_written: int = 0
    for sheet in sheets:
        st = time.monotonic()
        sheet = excel.add_worksheet(sheet, XIOWOptions(constant_memory=True, cache_typehints_write_optimization=True))

        acc: int = 0

        # for batch in data:
        #     for row in batch:
        #         sheet.write_row(acc, 0, row)
        #         # for j, value in enumerate(row.values()):
        #         #     sheet.write_cell(acc, j, value, format_)
        #         acc += 1

        for i, batch in enumerate(data):
            sheet.write_rows(i * len(batch), 0, batch)
            rows_written += len(batch)

        rows_written += acc

        print(f"fastxlsxio {sheet.name} sheet {NUM_ROWS}x{NUM_COLS} rows written for {time.monotonic() - st:.2f}")
    print(f"fastxlsxio {rows_written}x{NUM_COLS} rows written")

    excel.save()


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


# async def main():
#     wb = XIOWWorkbook(options=XIOWOptions(cache_col_formats=True))
#     ws = wb.add_worksheet("test")
#     firstf = wb.add_format({"num_format": "###00.0"})
#     secondf = wb.add_format({"num_format": "###00"})
#     ws.write_cell(0, 0, 123, firstf)
#     ws.write_cell(1, 0, 123, secondf)

#     wb.save("test_row_different_col_format_with_cache.xlsx")


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
