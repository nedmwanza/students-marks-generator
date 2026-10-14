from models.student import Student, ModuleResult

# Common workbook structure used by the current results-sheet sample.
# Module blocks are discovered from the module-code row, not hard-coded by programme.
STUDENT_HEADER_ROW = 12
FIRST_STUDENT_ROW = 13
MODULE_CODE_ROW = 8
MODULE_HEADER_ROW = 9


def _text(value):
    return str(value).strip() if value is not None else ""


def _discover_modules(ws):
    modules = []
    # The source sheet labels module codes in row 8 and each module block has
    # CA, SE, TOT, Grade in row 9. Scan the row instead of assuming six modules.
    for col in range(1, ws.max_column - 3):
        code = ws.cell(MODULE_CODE_ROW, col).value
        ca_label = _text(ws.cell(MODULE_HEADER_ROW, col).value).upper()
        se_label = _text(ws.cell(MODULE_HEADER_ROW, col + 1).value).upper()
        total_label = _text(ws.cell(MODULE_HEADER_ROW, col + 2).value).upper()
        grade_label = _text(ws.cell(MODULE_HEADER_ROW, col + 3).value).upper()

        if (
            isinstance(code, str)
            and code.strip()
            and ca_label == "CA"
            and se_label == "SE"
            and total_label in {"TOT", "TOTAL"}
            and grade_label == "GRADE"
        ):
            credit = ws.cell(7, col).value
            modules.append({
                "code": code.strip(),
                "credit": credit,
                "start_col": col,
            })
    return modules


def _discover_module_names(ws, modules):
    """Find module names in the list below the student records."""
    module_names = {
        module["code"]: ""
        for module in modules
    }

    # Normalise codes so accidental spaces do not prevent matching.
    code_lookup = {
        "".join(module["code"].upper().split()): module["code"]
        for module in modules
    }

    # Search below the student-data starting row because the module
    # catalogue is located after the student records.
    for row in range(FIRST_STUDENT_ROW, ws.max_row + 1):
        for col in range(1, ws.max_column):
            value = ws.cell(row, col).value

            if not isinstance(value, str):
                continue

            normalised_code = "".join(value.upper().split())

            if normalised_code not in code_lookup:
                continue

            code = code_lookup[normalised_code]
            possible_name = ws.cell(row, col + 1).value

            if isinstance(possible_name, str):
                name = possible_name.strip()

                if name and name.upper() not in {
                    "CA", "SE", "TOT", "TOTAL", "GRADE"
                }:
                    module_names[code] = name

    return module_names

def extract_students(ws):
    modules = _discover_modules(ws)
    module_names = _discover_module_names(ws, modules)
    if not modules:
        raise ValueError(
            "No module columns were detected. This workbook may use a different layout."
        )

    students = []
    for row in range(FIRST_STUDENT_ROW, ws.max_row + 1):
        serial = ws.cell(row, 1).value
        name = _text(ws.cell(row, 2).value)
        reg_no = _text(ws.cell(row, 3).value)


        # Skip rows without all three required student identifiers.
        if (
            serial is None
            or not str(serial).strip()
            or not name
            or not reg_no
        ):
            continue

        module_results = []
        for module in modules:
            col = module["start_col"]
            module_results.append(ModuleResult(
                code=module["code"],
                name=module_names.get(module["code"], ""),
                credits=module["credit"],
                ca=ws.cell(row, col).value,
                exam=ws.cell(row, col + 1).value,
                total=ws.cell(row, col + 2).value,
                grade=ws.cell(row, col + 3).value,
            ))

        
        # Summary fields start immediately after the last module's
        # four columns: CA, SE, Total, and Grade.
        summary_start_col = max(
            module["start_col"] + 4 for module in modules
        )

        students.append(Student(
            serial_number=serial,
            name=name,
            registration_number=reg_no,
            sex=ws.cell(row, 4).value,
            date_of_birth=ws.cell(row, 5).value,
            date_of_entry=ws.cell(row, 6).value,
            modules=module_results,
            semester_gpa=ws.cell(row, summary_start_col).value,
            cgpa=ws.cell(row, summary_start_col + 1).value,
            course_attendance=ws.cell(row, summary_start_col + 2).value,
            remarks=ws.cell(row, summary_start_col + 3).value,
            modules_failed=ws.cell(row, summary_start_col + 4).value,

            
        ))

    return students, modules
