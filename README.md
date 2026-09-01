# fastxlsxio

[Originally written from](https://github.com/shuangluoxss/fastxlsx)

**A lightweight, high-performance Python library for fast XLSX I/O operations.**  
*The project aims to create a write operation in xlsx that has the same functionality as Python’s xlsxwriter but faster, while keeping the reading implementation as it is from fastxlsx project.*

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
git clone https://github.com/JamesCrumble/fastxlsxio.git
cd fastxlsxio
pip install .
```