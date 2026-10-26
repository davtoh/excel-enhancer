# Excel Enhancer 🚀

[![Python Version](https://img.shields.io/badge/python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12-blue.svg)](https://www.python.org/)
[![xlwings](https://img.shields.io/badge/xlwings-%3E%3D0.30-orange.svg)](https://www.xlwings.org/)
[![pandas](https://img.shields.io/badge/pandas-%3E%3D2.0-150458.svg)](https://pandas.pydata.org/)
[![License](https://img.shields.io/badge/license-BSD--3--Clause-green.svg)](LICENSE.txt)
[![Tests](https://img.shields.io/badge/tests-47%20passed-brightgreen.svg)]()

**Excel Enhancer** bridges the gap between traditional spreadsheet modeling and high-performance Python data manipulation. Powered by [xlwings](https://www.xlwings.org/) and [pandas](https://pandas.pydata.org/), it exposes expressive User Defined Functions (UDFs) directly inside Excel cells, enabling complex table transforms, fuzzy matching, vector algebra, string templating, and formula step debugging without writing cumbersome VBA macros.

---

## 🌟 Key Capabilities

- **DataFrame Operations Inside Cells**: Group, aggregate, filter, sort, concatenate, and slice tabular data directly with spreadsheet formulas.
- **Smart Column Auto-Fitting**: Use whole-column references like `A:A` in formulas—Excel Enhancer dynamically detects active data boundaries to avoid processing millions of empty cells.
- **Intelligent Formula Caching**: Automatically hashes and caches intermediate formula blocks to accelerate complex, interdependent calculations.
- **Vector & String Manipulation**: Accent removal (`xv_normalize_accents`), fuzzy string searching with Levenshtein metrics (`xl_find_similar`, `xv_find_similar`), and multi-field formatters (`xlv_format_by`).
- **Dynamic Expression Evaluation**: Evaluate Python expressions dynamically across ranges (`xlv_apply_expr`, `xlv_map_by`).
- **Interactive Formula Debugger & Watchers**: Set breakpoints and data watchers directly on Excel cells using `xf_set_watcher` to inspect intermediate dataframes during formula evaluation.
- **Modular Extensions**: Easily extend capabilities with drop-in modules for geometry, fake data generation, JSON manipulation, and web scraping.

---

## ⚡ Quick Start

### 1. Start the UDF Server via CLI
```bash
excel-enhancer
```

### 2. Start via Python
```python
from excel_enhancer import serve

if __name__ == "__main__":
    serve()
```

---

## 📦 Installation & Excel Configuration

### Prerequisites
- Windows 10/11
- Microsoft Excel (2016, 2019, 2021, or Microsoft 365)
- Python 3.9+ (Python 3.12 recommended)

### Step 1: Install Python Dependencies
```bash
git clone https://github.com/davtoh/excel-enhancer.git
cd excel-enhancer
python -m venv .venv
.\.venv\Scripts\activate
pip install -e .
```

### Step 2: Install the xlwings Excel Add-in
Run in PowerShell / Command Prompt:
```bash
xlwings addin install
```
*(Note: If prompted for a password on add-in sheets, the default xlwings password is `xlwings`.)*

### Step 3: Enable Trust Access to the VBA Project
In Excel:
1. Navigate to **File > Options > Trust Center > Trust Center Settings...**.
2. Select **Macro Settings**.
3. Check the box **"Trust access to the VBA project object model"**.
4. Click **OK**.

### Step 4: Reference xlwings in VBA (Workbook Setup)
1. Open your macro-enabled workbook (`.xlsm`) or use [excel_enhancer.xlsm](file:///c:/dev/projects/products/excel_enhancer/excel_enhancer.xlsm).
2. Press `Alt + F11` to open the VBA Editor.
3. Go to **Tools > References...**.
4. Check **xlwings** in the list and click **OK**.
5. Save the workbook.

---

## 📖 UDF Function Catalog

All functions can be called directly inside Excel cells like standard formulas (e.g. `=xl_group_by(A1:C50, "Department", "Salary", "mean")`).

### 1. Table & DataFrame Operations (`xl_*`)
| Function | Description |
| :--- | :--- |
| `xl_group_by(table, by, target, operation, ...)` | Groups table by one or more columns and applies aggregations (`sum`, `mean`, `count`, `first`, `last`, `join`). |
| `xl_filter_by(table, column, value, mode, ...)` | Filters table rows matching a condition (`==`, `!=`, `<`, `>`, `contains`, `regex`). |
| `xl_sort_by(table, by, ascending, ...)` | Sorts table by one or more columns with ascending/descending flags. |
| `xl_select_by(table, columns, ...)` | Selects or reorders specific columns from a dataset. |
| `xl_concat(table1, table2, ...)` | Concatenates multiple tables along rows (`axis=0`) or columns (`axis=1`). |
| `xl_rename(table, old_names, new_names)` | Renames table column headers. |
| `xl_combinations(table, n)` | Computes all pairwise or n-ary combinations of table rows. |
| `xl_find_similar(table, column, query, ...)` | Finds closest matching rows based on string similarity. |
| `xl_expand_down(table, n)` / `xl_expand_right` | Expands or tiles table ranges vertically or horizontally. |
| `xl_expand_table(table, ...)` | Dynamically expands a 2D range to match target dimensions. |

### 2. Vector & Array Operations (`xv_*`)
| Function | Description |
| :--- | :--- |
| `xv_normalize_accents(vector)` | Strips Spanish and international diacritics/accents (e.g., `á`, `é`, `í`, `ó`, `ú`, `ñ` -> `a`, `e`, `i`, `o`, `u`, `n`). |
| `xv_index(vector, indices)` | Slices a vector by a list of 1-based row/column indices. |
| `xv_find_similar(vector, query, mode, ...)` | Returns similarity scores or closest matches for a 1D vector against target strings. |

### 3. Functional & Expressions (`xlv_*`)
| Function | Description |
| :--- | :--- |
| `xlv_apply_expr(range, expression, ...)` | Evaluates a Python expression row-by-row or element-wise (e.g. `x * 1.15` or `f"{row[0]}_{row[1]}"`). |
| `xlv_map_by(range, mapping_dict_or_range)` | Maps values across a dictionary or 2-column key-value lookup range. |
| `xlv_format_by(template, *ranges)` | Python format string template applied element-wise across input ranges (e.g. `xlv_format_by("ID-{:04d}", A1:A10)`). |
| `xlv_apply_str(range, method_name, *args)` | Invokes standard Python `str` methods (`upper`, `lower`, `strip`, `replace`, etc.). |

### 4. Utilities & Formula Watchers (`xf_*`)
| Function | Description |
| :--- | :--- |
| `xf_set_watcher(range, watcher_name)` | Attaches a named interactive watcher to track intermediate data states during formula computation. |
| `xf_del_watcher(watcher_name)` | Removes an active formula watcher. |
| `xf_fit_table(range)` | Trims empty peripheral cells and bounds from whole-column selections (`A:A`). |
| `xf_recalculate_udfs()` | Forces recalculation and cache invalidation for all active Excel Enhancer UDFs. |
| `xf_get_caller_address()` | Returns the calling cell address (e.g. `$B$4`). |
| `xf_load_library(module_name)` / `xf_delete_library` | Dynamically imports or reloads external Python modules at runtime. |
| `xf_evaluate(expression)` | Evaluates arbitrary Python expressions directly in the spreadsheet. |
| `xf_test_data(generator_type, rows, cols)` | Generates dummy numeric, categorical, or random test data. |

---

## 🛠️ Built-in Extensions

Excel Enhancer includes extensible modules located in the `extensions/` directory:
- **`udf_geometry.py`**: Computes point distances, polygon areas, bounding boxes, and geometric centroids.
- **`udf_json_support.py`**: Parses nested JSON strings, extracts keys via JSONPath, and serializes ranges to JSON.
- **`udf_fake_data.py`**: Generates synthetic names, addresses, emails, dates, and realistic mock business records.
- **`udf_web_scraper.py`**: Fetches and parses live tables or metadata from HTTP endpoints directly into Excel.

---

## 🐞 Interactive Debugging & Watchers

When debugging complex nested formulas, you can intercept intermediate calculation states without print statements or logging files:

1. Wrap the intermediate expression with `xf_set_watcher`:
   ```excel
   =xl_sort_by(xf_set_watcher(xl_filter_by(A1:D100, "Status", "Active"), "filtered_active"), "Amount", FALSE)
   ```
2. The intermediate DataFrame is saved in memory under `"filtered_active"`.
3. In another cell or in the Python REPL, inspect the cached watcher state:
   ```excel
   =xf_evaluate("watchers['filtered_active']")
   ```
4. Once verified, simply remove the watcher with `=xf_del_watcher("filtered_active")`.

---

## 📁 Workbook Examples

The repository provides ready-to-run workbooks demonstrating functions in action:
- **`examples.xlsm`**: Master workbook featuring dedicated sheets with sample datasets for `map_by`, `group_by`, `select_by`, `fit_table`, `filter_by`, `combinations`, and `debug`.
- **`examples/xl_sort_by.xlsm`**: Multi-column sorting demonstration.
- **`examples/xlv_apply_expr.xlsm`**: Row-wise Python expression evaluation.
- **`examples/xlv_format_by.xlsm`**: Multi-column string formatting.
- **`examples/xlv_map_by.xlsm`**: Dynamic key-value dictionary lookup.
- **`examples/xv_normalize_accents.xlsm`**: International character normalization.
- **`examples/excel_enhancer.xlsm`**: Blank template preconfigured with VBA code and xlwings references.

---

## 🔧 Troubleshooting & FAQ

#### Q: Excel displays `#VALUE!` or `Run-time error '91': Object variable or With block variable not set` / `Object Required`
- **Cause**: VBA project object model access is disabled or xlwings add-in is not referenced.
- **Fix**:
  1. Ensure **Trust access to the VBA project object model** is enabled in Excel Options (see Step 3).
  2. In the VBA editor (`Alt + F11`), verify **Tools > References > xlwings** is checked.
  3. See [StackOverflow discussion #49384532](https://stackoverflow.com/a/49384532/5288758).

#### Q: "No command specified in the configuration, cannot autostart server"
- **Cause**: The xlwings ribbon cannot locate the Python interpreter path.
- **Fix**:
  1. On the **xlwings** ribbon tab in Excel, verify `Interpreter` points to your virtual environment (e.g., `c:\dev\projects\products\excel_enhancer\.venv\Scripts\python.exe`).
  2. Click **Restart UDF Server**, then click **Import Functions**.

#### Q: How do I run the test suite?
Run `pytest` from the root of the repository:
```bash
pytest -v
```
All 47 unit tests cover DataFrame transformations, converters, error handlers, and extension modules.

---

## 📄 License

This project is licensed under the BSD 3-Clause License. See [LICENSE.txt](file:///c:/dev/projects/products/excel_enhancer/LICENSE.txt) for details.
