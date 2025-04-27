# Excel Report Generator

This project **automatically generates Excel reports** based on a **template** and **input JSON data**.
It preserves all the template's **styles** (e.g., colors, fonts, borders) and **dynamically replaces placeholders** in the template with real computed or extracted values.

---

## How It Works

- Reads a **template Excel file** (`.xlsx`) that contains:
  - Pre-defined headers
  - Styles (colors, fonts, etc.)
  - Special **placeholders** like `{o_t_1}`, `{o_t_2}`, `{o_t_3}`, `{o_t_4}`
- Fills in the placeholders with:
  - Random generated values (for `{o_t_1}` and `{o_t_2}`)
  - Values extracted from a **JSON report file** (for `{o_t_3}`, `{o_t_4}`)
- Saves the final customized Excel into a specified output path.

---

## Project Structure

```
project/
│
├── source/
│   ├── template_input/
│   │   └── template.xlsx          # Your Excel template
│   ├── report_input/
│   │   └── report.json             # Your CMM JSON output
│   └── report_output/
│       └── output.xlsx             # Final generated report
│
├── .env                            # Environment variables
├── requirements.txt                # Python dependencies
├── main.py                         # Entry script (example)
└── README.md                       # You're here
```

---

## .env Settings

You need to define these variables in a `.env` file in the root of your project:

```dotenv
TEMPLATE_INPUT_PATH=source/template_input/template.xlsx
REPORT_OUTPUT_PATH=source/report_output/output.xlsx
REPORT_INPUT_PATH=source/report_input/report.json
MIN_ROW=1
MAX_ROW=100
MIN_COL=1
MAX_COL=40
```

| Variable              | Meaning                               |
| --------------------- | ------------------------------------- |
| `TEMPLATE_INPUT_PATH` | Path to your Excel template           |
| `REPORT_OUTPUT_PATH`  | Path to save the final output Excel   |
| `REPORT_INPUT_PATH`   | Path to the input JSON data           |
| `MIN_ROW`, `MAX_ROW`  | Rows in the Excel template to copy    |
| `MIN_COL`, `MAX_COL`  | Columns in the Excel template to copy |

Use `python-dotenv` to load these automatically.

---

## Install Requirements

Install dependencies:

```bash
pip install -r requirements.txt
```

Contents of `requirements.txt`:

```
openpyxl
python-dotenv
```

---

## How To Run

Example:

```python
from your_module import ExcelReportOutput

# Read from environment variables
import os
from dotenv import load_dotenv
load_dotenv()

template_path = os.getenv('TEMPLATE_INPUT_PATH')
output_path = os.getenv('REPORT_OUTPUT_PATH')
input_path = os.getenv('REPORT_INPUT_PATH')

excel_report = ExcelReportOutput(template_path, output_path, input_path)
excel_report.output()
```

---

## Code Explanation

| Part                       | What it does                                                                                |
| -------------------------- | ------------------------------------------------------------------------------------------- |
| `read_template()`          | Reads the cell values from the template (only between MIN_ROW–MAX_ROW, MIN_COL–MAX_COL)     |
| `copy_template_format(ws)` | Copies all formatting (colors, fonts, fills, borders) from the template into a new workbook |
| `copy_cell()`              | Copies individual cell values and styles                                                    |
| `switch(value)`            | Matches placeholders and decides whether to compute a random value or fetch from JSON       |
| `computeResult()`          | Generates a random integer between 200–4000                                                 |
| `extractFromJson(value)`   | Pulls specific fields from the input JSON based on the placeholder                          |

---

## Example

If your template has a cell with:

```
{o_t_3}
```

It will be replaced by the `duplication` value from the JSON inside:

```json
"deduplicated": {
    "duplication": 6418756
}
```

Similarly:

- `{o_t_1}` → random number
- `{o_t_2}` → random number
- `{o_t_4}` → duplicationPercent from the JSON

---

## Future Improvements

- Add validation for missing `.env` variables.
- Automatically expand dynamic rows if multiple campaigns exist.
- Add better error handling (e.g., file not found, bad JSON).
- Optional: generate summary charts inside Excel.

---
