from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def apply_rtl(sheet) -> None:
    sheet.sheet_view.rightToLeft = True
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A5"


def style_title_cell(cell) -> None:
    cell.font = Font(name="Arial", size=20, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color="111111"),
        right=Side(style="thin", color="111111"),
        top=Side(style="thin", color="111111"),
        bottom=Side(style="thin", color="111111"),
    )


def add_instruction_card(sheet, title: str, instructions: list[str]) -> None:
    sheet.merge_cells("A1:C2")
    title_cell = sheet["A1"]
    title_cell.value = title
    style_title_cell(title_cell)

    start_row = 4
    sheet.merge_cells(f"A{start_row}:C{start_row + 4}")
    header = sheet[f"A{start_row}"]
    header.value = "כרטיס הוראות"
    header.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header.fill = PatternFill("solid", fgColor="5B7DB1")
    header.alignment = Alignment(horizontal="center", vertical="center")
    header.border = Border(
        left=Side(style="thin", color="333333"),
        right=Side(style="thin", color="333333"),
        top=Side(style="thin", color="333333"),
        bottom=Side(style="thin", color="333333"),
    )

    for idx, text in enumerate(instructions, start=1):
        row = start_row + idx
        sheet.merge_cells(f"A{row}:C{row}")
        cell = sheet[f"A{row}"]
        cell.value = f"{idx}. {text}"
        cell.font = Font(name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor="EAF3FF")
        cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color="C0C0C0"),
            right=Side(style="thin", color="C0C0C0"),
            top=Side(style="thin", color="C0C0C0"),
            bottom=Side(style="thin", color="C0C0C0"),
        )


def add_questions_table(sheet, questions: list[dict]) -> None:
    start_row = 20
    headers = ["מספר", "שאלה", "תשובה"]
    for col_idx, header in enumerate(headers, start=1):
        cell = sheet.cell(row=start_row, column=col_idx, value=header)
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4F81BD")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin", color="303030"),
            right=Side(style="thin", color="303030"),
            top=Side(style="thin", color="303030"),
            bottom=Side(style="thin", color="303030"),
        )

    for idx, question in enumerate(questions, start=1):
        row = start_row + idx
        sheet.cell(row=row, column=1, value=idx).alignment = Alignment(horizontal="center")
        q_cell = sheet.cell(row=row, column=2, value=question["prompt"])
        q_cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        q_cell.fill = PatternFill("solid", fgColor="F9FBFD")

        answer_cell = sheet.cell(row=row, column=3, value="")
        answer_cell.alignment = Alignment(horizontal="center", vertical="center")
        answer_cell.fill = PatternFill("solid", fgColor="FFF7D6")

        for col in range(1, 4):
            cell = sheet.cell(row=row, column=col)
            cell.border = Border(
                left=Side(style="thin", color="B7B7B7"),
                right=Side(style="thin", color="B7B7B7"),
                top=Side(style="thin", color="B7B7B7"),
                bottom=Side(style="thin", color="B7B7B7"),
            )

    ws_cols = {"A": 12, "B": 48, "C": 18}
    for col_letter, width in ws_cols.items():
        sheet.column_dimensions[col_letter].width = width


def add_feedback_and_errors(sheet, feedback: list[str], errors: list[str]) -> None:
    feedback_row = 15
    sheet.merge_cells(f"A{feedback_row}:C{feedback_row}")
    feedback_header = sheet[f"A{feedback_row}"]
    feedback_header.value = "משוב"
    feedback_header.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    feedback_header.fill = PatternFill("solid", fgColor="7AAE5E")
    feedback_header.alignment = Alignment(horizontal="center", vertical="center")
    feedback_header.border = Border(
        left=Side(style="thin", color="333333"),
        right=Side(style="thin", color="333333"),
        top=Side(style="thin", color="333333"),
        bottom=Side(style="thin", color="333333"),
    )

    for idx, text in enumerate(feedback, start=1):
        row = feedback_row + idx
        sheet.merge_cells(f"A{row}:C{row}")
        cell = sheet[f"A{row}"]
        cell.value = text
        cell.font = Font(name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor="EAFBEA")
        cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color="B7B7B7"),
            right=Side(style="thin", color="B7B7B7"),
            top=Side(style="thin", color="B7B7B7"),
            bottom=Side(style="thin", color="B7B7B7"),
        )

    errors_row = feedback_row + len(feedback) + 2
    sheet.merge_cells(f"A{errors_row}:C{errors_row}")
    errors_header = sheet[f"A{errors_row}"]
    errors_header.value = "הודעות שגיאה"
    errors_header.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    errors_header.fill = PatternFill("solid", fgColor="B94A48")
    errors_header.alignment = Alignment(horizontal="center", vertical="center")
    errors_header.border = Border(
        left=Side(style="thin", color="333333"),
        right=Side(style="thin", color="333333"),
        top=Side(style="thin", color="333333"),
        bottom=Side(style="thin", color="333333"),
    )

    for idx, text in enumerate(errors, start=1):
        row = errors_row + idx
        sheet.merge_cells(f"A{row}:C{row}")
        cell = sheet[f"A{row}"]
        cell.value = text
        cell.font = Font(name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor="FCE9E9")
        cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color="B7B7B7"),
            right=Side(style="thin", color="B7B7B7"),
            top=Side(style="thin", color="B7B7B7"),
            bottom=Side(style="thin", color="B7B7B7"),
        )


def build_practice_workbook(
    title: str,
    subject: str,
    grade: str,
    instructions: list[str],
    feedback: list[str],
    errors: list[str],
    questions: list[dict],
) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "תרגילים"
    apply_rtl(ws)

    ws["A3"] = f"{subject} | {grade}"
    ws["A3"].font = Font(name="Arial", size=11, italic=True, color="444444")
    ws["A3"].alignment = Alignment(horizontal="right", vertical="center")

    add_instruction_card(ws, title, instructions)
    add_questions_table(ws, questions)
    add_feedback_and_errors(ws, feedback, errors)

    return wb


def generate_sample_workbook(output_path: str | Path | None = None) -> str:
    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    if output_path is None:
        output_path = output_dir / "תרגול_אחוזים.xlsx"
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

    questions = [
        {"prompt": "חשב/י 25% מתוך 200."},
        {"prompt": "מוצר עולה 150 ש\"ח. לאחר הנחה של 20%, מהו המחיר החדש?"},
        {"prompt": "30 מתוך 120 תלמידים השתתפו. מה אחוז ההשתתפות?"},
        {"prompt": "השלם/י: 40% מתוך 300 הוא ___ ."},
        {"prompt": "מהו 10% מתוך 450?"},
    ]

    workbook = build_practice_workbook(
        title="תרגול אחוזים",
        subject="מתמטיקה",
        grade="כיתה ז'",
        instructions=[
            "קראו את השאלה בעיון.",
            "פתרו את התרגיל במחברת.",
            "רשמו את התשובה בגיליון.",
            "בדקו כל שורה לפני סיום העבודה.",
        ],
        feedback=[
            "כל הכבוד!",
            "נסו שוב.",
            "בדקו את החישוב.",
        ],
        errors=[
            "הזינו מספר תקין.",
            "בדקו שהאחוז מתאים למספר.",
            "הכפילו את האחוז במספר נכון.",
        ],
        questions=questions,
    )

    workbook.save(output_path)
    return str(output_path)


if __name__ == "__main__":
    result = generate_sample_workbook()
    print(f"קובץ נוצר: {result}")
