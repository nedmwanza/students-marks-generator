def validate_students(students):
    errors = []
    seen_registration_numbers = set()

    for student in students:
        reg_no = student.registration_number.strip()
        if not reg_no:
            errors.append("A student is missing a registration number.")
            continue

        if reg_no in seen_registration_numbers:
            errors.append(f"Duplicate registration number: {reg_no}")
        seen_registration_numbers.add(reg_no)

        for module in student.modules:
            for label, value in (("CA", module.ca), ("SE", module.exam), ("total", module.total)):
                if isinstance(value, (int, float)) and not 0 <= value <= 100:
                    errors.append(
                        f"{student.registration_number}: {module.code} has invalid {label} mark {value}."
                    )

    return errors
