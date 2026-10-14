from dataclasses import dataclass, field
from typing import Any


@dataclass
class ModuleResult:
    code: str
    name: str = ""
    credits: Any = None
    ca: Any = None
    exam: Any = None
    total: Any = None
    grade: Any = None


@dataclass
class Student:
    serial_number: Any
    name: str
    registration_number: str
    sex: Any = None
    date_of_birth: Any = None
    date_of_entry: Any = None
    modules: list[ModuleResult] = field(default_factory=list)
    semester_gpa: Any = None
    cgpa: Any = None
    course_attendance: Any = None
    remarks: Any = None
    modules_failed: Any = None
