from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_DIR = Path(__file__).resolve().parent / "output"
THIN = Side(style="thin", color="B7B7B7")
DARK = Side(style="thin", color="333333")


def border(side: Side = THIN) -> Border:
    return Border(left=side, right=side, top=side, bottom=side)


def apply_rtl(sheet) -> None:
    sheet.sheet_view.rightToLeft = True
    sheet.sheet_view.showGridLines = False


def style_title(sheet, title: str, subtitle: str = "") -> None:
    sheet.merge_cells("A1:C2")
    cell = sheet["A1"]
    cell.value = title
    cell.font = Font(name="Arial", size=18, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border(DARK)
    sheet.row_dimensions[1].height = 28
    sheet.row_dimensions[2].height = 28
    if subtitle:
        sheet["A3"] = subtitle
        sheet["A3"].font = Font(name="Arial", size=11, italic=True, color="555555")
        sheet["A3"].alignment = Alignment(horizontal="right")


def add_instruction_card(sheet, instructions: Iterable[str], start_row: int = 5) -> int:
    instructions = list(instructions)
    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
    header = sheet.cell(start_row, 1, "כרטיס הוראות")
    header.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header.fill = PatternFill("solid", fgColor="5B7DB1")
    header.alignment = Alignment(horizontal="center", vertical="center")
    header.border = border(DARK)

    for index, text in enumerate(instructions, start=1):
        row = start_row + index
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell = sheet.cell(row, 1, f"{index}. {text}")
        cell.font = Font(name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor="EAF3FF")
        cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        cell.border = border()
        sheet.row_dimensions[row].height = 25
    return start_row + len(instructions) + 1


def add_feedback_card(sheet, feedback: list[str], start_row: int) -> int:
    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
    header = sheet.cell(start_row, 1, "משוב לתלמיד")
    header.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header.fill = PatternFill("solid", fgColor="70AD47")
    header.alignment = Alignment(horizontal="center")
    header.border = border(DARK)
    for index, text in enumerate(feedback, start=1):
        row = start_row + index
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell = sheet.cell(row, 1, text)
        cell.fill = PatternFill("solid", fgColor="E2F0D9")
        cell.alignment = Alignment(horizontal="right", wrap_text=True)
        cell.border = border()
    return start_row + len(feedback) + 1


def add_error_card(sheet, errors: list[str], start_row: int) -> int:
    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
    header = sheet.cell(start_row, 1, "שגיאות נפוצות")
    header.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header.fill = PatternFill("solid", fgColor="C0504D")
    header.alignment = Alignment(horizontal="center")
    header.border = border(DARK)
    for index, text in enumerate(errors, start=1):
        row = start_row + index
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell = sheet.cell(row, 1, text)
        cell.fill = PatternFill("solid", fgColor="FCE4D6")
        cell.alignment = Alignment(horizontal="right", wrap_text=True)
        cell.border = border()
    return start_row + len(errors) + 1


def style_table_header(sheet, row: int, headers: list[str]) -> None:
    for column, value in enumerate(headers, start=1):
        cell = sheet.cell(row, column, value)
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4F81BD")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border(DARK)


def style_data_range(sheet, first_row: int, last_row: int, first_col: int, last_col: int) -> None:
    for row in sheet.iter_rows(min_row=first_row, max_row=last_row, min_col=first_col, max_col=last_col):
        for cell in row:
            cell.border = border()
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)


def set_common_widths(sheet) -> None:
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 48
    sheet.column_dimensions["C"].width = 20


def build_percentages_sheet(workbook: Workbook) -> None:
    sheet = workbook.create_sheet("תרגול אחוזים")
    apply_rtl(sheet)
    style_title(sheet, "תרגול אחוזים", "מתמטיקה | כיתה ז׳")
    instructions_end = add_instruction_card(sheet, [
        "קראו כל שאלה בעיון.",
        "פתרו את התרגיל במחברת.",
        "רשמו את התשובה בעמודה הצהובה.",
        "בדקו את החישוב לפני הגשה.",
    ])
    table_row = instructions_end + 2
    style_table_header(sheet, table_row, ["מספר", "שאלה", "תשובת התלמיד"])
    questions = [
        "חשב/י 25% מתוך 200.",
        "מוצר עולה 150 ש״ח. לאחר הנחה של 20%, מהו המחיר החדש?",
        "30 מתוך 120 תלמידים השתתפו. מהו אחוז ההשתתפות?",
        "השלם/י: 40% מתוך 300 הוא ___.",
        "מהו 10% מתוך 450?",
    ]
    for index, question in enumerate(questions, start=1):
        row = table_row + index
        sheet.cell(row, 1, index)
        sheet.cell(row, 2, question)
        sheet.cell(row, 3, "").fill = PatternFill("solid", fgColor="FFF2CC")
    style_data_range(sheet, table_row + 1, table_row + len(questions), 1, 3)
    feedback_row = table_row + len(questions) + 3
    feedback_end = add_feedback_card(sheet, ["כל הכבוד!", "נסו שוב ובדקו את החישוב.", "אפשר להשתמש ברמז: אחוז הוא חלק מתוך 100."], feedback_row)
    add_error_card(sheet, ["הזנתם טקסט במקום מספר.", "שכחתם לחלק ב־100.", "בדקו אם מדובר בהנחה או בתוספת."], feedback_end + 2)
    set_common_widths(sheet)
    sheet.freeze_panes = f"A{table_row + 1}"


def build_conditional_formatting_sheet(workbook: Workbook) -> None:
    sheet = workbook.create_sheet("עיצוב מותנה")
    apply_rtl(sheet)
    style_title(sheet, "תרגול עיצוב מותנה", "מיומנויות Excel | כיתה ז׳ ומעלה")
    end = add_instruction_card(sheet, [
        "סמנו את טווח הציונים C9:C18.",
        "צבעו באדום ציונים מתחת ל־60.",
        "צבעו בירוק ציונים מעל או שווים ל־90.",
        "הוסיפו סרגל צבעים לטווח הציונים.",
    ])
    table_row = end + 2
    style_table_header(sheet, table_row, ["מספר", "שם תלמיד/ה", "ציון"])
    names_scores = [("דנה", 85), ("יוסי", 45), ("רון", 95), ("מיה", 72), ("אורי", 58), ("נועה", 91), ("תמר", 66), ("גל", 38), ("עידו", 88), ("שחר", 100)]
    for index, (name, score) in enumerate(names_scores, start=1):
        row = table_row + index
        sheet.cell(row, 1, index)
        sheet.cell(row, 2, name)
        sheet.cell(row, 3, score)
    last_row = table_row + len(names_scores)
    style_data_range(sheet, table_row + 1, last_row, 1, 3)
    score_range = f"C{table_row + 1}:C{last_row}"
    sheet.conditional_formatting.add(score_range, CellIsRule(operator="lessThan", formula=["60"], fill=PatternFill("solid", fgColor="FFC7CE")))
    sheet.conditional_formatting.add(score_range, CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=PatternFill("solid", fgColor="C6EFCE")))
    sheet.conditional_formatting.add(score_range, ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    add_feedback_card(sheet, ["הצבעים מאפשרים לזהות במהירות הישגים ותחומים לשיפור."], last_row + 3)
    add_error_card(sheet, ["אל תצבעו ידנית במקום להשתמש בעיצוב מותנה.", "בדקו שהטווח כולל את כל הציונים."], last_row + 6)
    set_common_widths(sheet)


def build_sort_filter_sheet(workbook: Workbook) -> None:
    sheet = workbook.create_sheet("מיון וסינון")
    apply_rtl(sheet)
    style_title(sheet, "תרגול מיון וסינון", "מיומנויות Excel | טבלת נתונים")
    end = add_instruction_card(sheet, [
        "הפכו את הנתונים לטבלת Excel.",
        "מיינו את המכירות מהגבוהה לנמוכה.",
        "סננו והציגו רק מוצרים מהקטגוריה 'ציוד'.",
        "חשבו את סך המכירות בעזרת SUM.",
    ])
    table_row = end + 2
    headers = ["מוצר", "קטגוריה", "מחיר", "כמות", "סך מכירות"]
    style_table_header(sheet, table_row, headers)
    rows = [("מחברת", "כתיבה", 12, 15), ("עטים", "כתיבה", 8, 30), ("תיק", "ציוד", 120, 4), ("קלמר", "ציוד", 45, 8), ("מחשבון", "ציוד", 75, 6)]
    for index, (product, category, price, quantity) in enumerate(rows, start=1):
        row = table_row + index
        sheet.cell(row, 1, product)
        sheet.cell(row, 2, category)
        sheet.cell(row, 3, price)
        sheet.cell(row, 4, quantity)
        sheet.cell(row, 5, f"=C{row}*D{row}")
    last_row = table_row + len(rows)
    style_data_range(sheet, table_row + 1, last_row, 1, 5)
    table = Table(displayName="SalesPracticeTable", ref=f"A{table_row}:E{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    sheet.add_table(table)
    for col, width in {"A": 18, "B": 18, "C": 14, "D": 14, "E": 18}.items():
        sheet.column_dimensions[col].width = width
    add_feedback_card(sheet, ["טבלה מאפשרת מיון וסינון מהירים בלי לשנות את הנתונים."], last_row + 3)
    add_error_card(sheet, ["ודאו שכל העמודות כלולות בטווח הטבלה.", "אל תמחקו את שורת הכותרות."], last_row + 6)


def build_chart_sheet(workbook: Workbook) -> None:
    sheet = workbook.create_sheet("יצירת גרף")
    apply_rtl(sheet)
    style_title(sheet, "תרגול יצירת גרף", "מיומנויות Excel | הצגת נתונים")
    end = add_instruction_card(sheet, [
        "בחרו את טווח הנתונים A9:B13.",
        "הוסיפו תרשים עמודות.",
        "תנו לתרשים את הכותרת 'מכירות לפי חודש'.",
        "הוסיפו שמות לציר האופקי ולציר האנכי.",
    ])
    table_row = end + 2
    style_table_header(sheet, table_row, ["חודש", "מכירות"])
    values = [("ינואר", 120), ("פברואר", 180), ("מרץ", 150), ("אפריל", 230), ("מאי", 200)]
    for index, (month, sales) in enumerate(values, start=1):
        row = table_row + index
        sheet.cell(row, 1, month)
        sheet.cell(row, 2, sales)
    last_row = table_row + len(values)
    style_data_range(sheet, table_row + 1, last_row, 1, 2)
    chart = BarChart()
    chart.title = "מכירות לפי חודש"
    chart.y_axis.title = "מכירות"
    chart.x_axis.title = "חודש"
    chart.add_data(Reference(sheet, min_col=2, min_row=table_row, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=table_row + 1, max_row=last_row))
    chart.height = 7
    chart.width = 13
    sheet.add_chart(chart, "E8")
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 18
    add_feedback_card(sheet, ["גרף טוב מציג את הנתונים בצורה ברורה ומאפשר להשוות בין חודשים."], last_row + 3)
    add_error_card(sheet, ["אל תכללו את שורת ההוראות בטווח הגרף.", "בדקו שהכותרות נמצאות בשורה הראשונה של הטבלה."], last_row + 6)


def build_answer_sheet(workbook: Workbook) -> None:
    sheet = workbook.create_sheet("תשובות למורה")
    apply_rtl(sheet)
    style_title(sheet, "תשובות למורה", "גיליון זה מיועד לבדיקה")
    style_table_header(sheet, 5, ["תרגיל", "תשובה / בדיקה", "הערה למורה"])
    answers = [
        ("תרגול אחוזים", "50; 120; 25; 120; 45", "יש לקבל גם תשובה עשרונית שקולה."),
        ("עיצוב מותנה", "מתחת ל־60 אדום; 90 ומעלה ירוק", "בדקו שימוש בכלל ולא צביעה ידנית."),
        ("מיון וסינון", "סך המכירות מחושב בעמודה E", "בדקו שהטבלה כוללת מסננים."),
        ("יצירת גרף", "תרשים עמודות לפי חודשים", "בדקו כותרות וצירי התרשים."),
    ]
    for index, row_data in enumerate(answers, start=6):
        for col, value in enumerate(row_data, start=1):
            sheet.cell(index, col, value)
    style_data_range(sheet, 6, 5 + len(answers), 1, 3)
    for col, width in {"A": 25, "B": 42, "C": 38}.items():
        sheet.column_dimensions[col].width = width
    sheet.freeze_panes = "A6"


def generate_sample_workbook(output_path: str | Path | None = None) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = Path(output_path) if output_path else OUTPUT_DIR / "דוגמאות_תרגול.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    build_percentages_sheet(workbook)
    build_conditional_formatting_sheet(workbook)
    build_sort_filter_sheet(workbook)
    build_chart_sheet(workbook)
    build_answer_sheet(workbook)
    workbook.save(output_path)
    return str(output_path)


if __name__ == "__main__":
    print(f"קובץ נוצר: {generate_sample_workbook()}")
