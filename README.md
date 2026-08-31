# fastxlsxio

[Originally written from](https://github.com/shuangluoxss/fastxlsx)

**A lightweight, high-performance Python library for blazing-fast XLSX I/O operations.**  

### ✅ Supported Capabilities

- **Data Types**: Native support for `bool`, `int`, `float`, `date`, `datetime`, and `str`.
- **Data Operations**: Scalars, rows, columns, matrices, and batch processing.
- **Coordinate Systems**: Dual support for **A1** (e.g., `B2`) and **R1C1** (e.g., `(2, 3)`) notation.
- **Parallel Processing**: Multi-threaded read/write operations for massive datasets.
- **Type Safety**: Full type hints and IDE-friendly documentation.
- **Blasting Performance**: 5-10x faster compared to `openpyxl`.

### 🚫 Current Limitations

- **File Formats**: Only XLSX (no XLS/XLSB support).
- **Formulas & Styling**: Cell formulas, merged cells, and formatting not supported.
- **Modifications**: Append/update operations on existing files unavailable.
- **Advanced Features**: Charts, images, and other advanced features not supported.

## 🛠️ Installation

<!-- ### PyPI Install

```bash
pip install fastxlsxio
``` -->

### Source Build (Requires Rust Toolchain)

```bash
git clone ...
cd fastxlsxio
pip install .
```