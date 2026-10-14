# Students' Statement of Marks Generator — Version 1 starter

## Current milestone
This modular starter application uses CustomTkinter to select an input workbook and output folder, then inspect the workbook's student/module structure and validate basic data. It does not generate statement files yet; we will add that after confirming extraction.

## Project structure
- `main.py`: application entry point
- `ui/`: CustomTkinter interface
- `readers/`: Excel workbook loading
- `extractors/`: student and module extraction
- `processors/`: data validation
- `models/`: Student and ModuleResult data classes
- `generators/`: next milestone — statement generation
- `templates/statement_template.xlsx`: copy of the supplied statement format
- `input/BBA.xlsx`: supplied test workbook
- `output/`: generated files will be written here in a later milestone

## Setup (Windows PowerShell)
1. Extract this ZIP.
2. Open PowerShell in the `students_marks_system` folder.
3. Create a virtual environment (recommended):
   `python -m venv .venv`
4. Activate it:
   `.venv\Scripts\Activate.ps1`
5. Install packages:
   `pip install -r requirements.txt`
6. Run:
   `python main.py`
7. Click Browse and select `input/BBA.xlsx`, then click **Inspect Workbook**.

## Important
- The initial extractor is based on the shared layout visible in the supplied BBA workbook: student records begin at row 13, module codes are on row 8, and each module has CA / SE / TOT / Grade columns.
- Module blocks are discovered from the workbook headers rather than hard-coding the BBA module count.
- Student statement filenames will use the registration number, e.g. `25311552080.xlsx`.
- The statement generator is not implemented in this milestone. First verify the detected students and modules before generating official-looking statements.
- The supplied statement template contains external workbook formulas. The eventual generator should write values into a copy of the template and remove/replace external formulas rather than relying on the linked workbook.
