"""
Excel Export module for Expense Tracker Telegram Bot.
Generates beautifully formatted Excel (.xlsx) reports using openpyxl.
"""

import os
from datetime import datetime
from typing import List, Dict, Any
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from config import REPORTS_DIR


def generate_excel_report(expenses: List[Dict[str, Any]], period_name: str, start_date_str: str, end_date_str: str, user_id: int) -> str:
    """
    Generates a professionally styled Excel file from expenses list.
    Returns the absolute path to the generated .xlsx file.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Харажатлар"
    ws.views.sheetView[0].showGridLines = True

    # Palette
    NAVY_HEADER = "1F4E79"
    ACCENT_COLOR = "2E75B6"
    LIGHT_BG = "F2F2F2"
    BORDER_COLOR = "D9D9D9"
    ZEBRA_BG = "F9FAFB"
    TOTAL_BG = "D9E1F2"

    font_title = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11, color="000000")
    font_total = Font(name="Calibri", size=12, bold=True, color="1F4E79")

    fill_header = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_cat_header = PatternFill(start_color=ACCENT_COLOR, end_color=ACCENT_COLOR, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_BG, end_color=ZEBRA_BG, fill_type="solid")
    fill_total = PatternFill(start_color=TOTAL_BG, end_color=TOTAL_BG, fill_type="solid")

    thin_border_side = Side(border_style="thin", color=BORDER_COLOR)
    double_bottom = Side(border_style="double", color="1F4E79")
    border_data = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(top=thin_border_side, bottom=double_bottom)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # 1. Title Banner
    ws.merge_cells("A1:E1")
    ws["A1"] = "📊 ХАРАЖАТЛАР БЎЙИЧА ҲИСОБОТ"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_left
    ws.row_dimensions[1].height = 25

    # Subtitles
    ws["A2"] = f"Давр: {period_name} ({start_date_str[:10]} — {end_date_str[:10]})"
    ws["A2"].font = font_subtitle
    ws["A3"] = f"Яратилган вақт: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A3"].font = font_subtitle
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 18

    # 2. Main Expenses Table Header
    start_row = 5
    headers = ["Т/р", "Сана ва вақт", "Категория", "Сумма (сўм)", "Изоҳ"]
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center if col_idx in [1, 2] else (align_right if col_idx == 4 else align_left)
    ws.row_dimensions[start_row].height = 24

    # 3. Data Rows
    current_row = start_row + 1
    total_amount = 0.0
    category_totals: Dict[str, float] = {}

    for i, exp in enumerate(expenses, 1):
        ws.row_dimensions[current_row].height = 20
        c1 = ws.cell(row=current_row, column=1, value=i)
        c2 = ws.cell(row=current_row, column=2, value=exp.get("created_at", ""))
        c3 = ws.cell(row=current_row, column=3, value=exp.get("category", ""))
        c4 = ws.cell(row=current_row, column=4, value=float(exp.get("amount", 0)))
        c5 = ws.cell(row=current_row, column=5, value=exp.get("description", "") or "-")

        c1.alignment = align_center
        c2.alignment = align_center
        c3.alignment = align_left
        c4.alignment = align_right
        c5.alignment = align_left

        c4.number_format = '#,##0'

        for cell in [c1, c2, c3, c4, c5]:
            cell.font = font_data
            cell.border = border_data
            if i % 2 == 0:
                cell.fill = fill_zebra

        amt = float(exp.get("amount", 0))
        total_amount += amt
        cat = exp.get("category", "Бошқа")
        category_totals[cat] = category_totals.get(cat, 0.0) + amt

        current_row += 1

    # If no expenses
    if not expenses:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        empty_cell = ws.cell(row=current_row, column=1, value="Ушбу давр учун харажатлар топилмади.")
        empty_cell.font = font_subtitle
        empty_cell.alignment = align_center
        current_row += 1

    # 4. Total Row
    total_row = current_row
    ws.row_dimensions[total_row].height = 24
    ws.cell(row=total_row, column=1, value="")
    ws.cell(row=total_row, column=2, value="")
    ws.cell(row=total_row, column=3, value="ЖАМИ ХАРАЖАТ:")
    t_sum_cell = ws.cell(row=total_row, column=4)

    if expenses:
        t_sum_cell.value = f"=SUM(D{start_row + 1}:D{total_row - 1})"
    else:
        t_sum_cell.value = 0

    t_sum_cell.number_format = '#,##0 "сўм"'
    ws.cell(row=total_row, column=5, value="")

    for c in range(1, 6):
        cell = ws.cell(row=total_row, column=c)
        cell.font = font_total
        cell.fill = fill_total
        cell.border = border_total
        if c == 3:
            cell.alignment = align_right
        elif c == 4:
            cell.alignment = align_right

    # 5. Category Breakdown Table (Side Summary)
    cat_start_col = 7  # Column G
    ws.cell(row=start_row - 1, column=cat_start_col, value="📊 Категориялар бўйича таҳлил").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
    
    cat_headers = ["Категория", "Сумма (сўм)", "Улуши (%)"]
    for idx, ch in enumerate(cat_headers):
        cell = ws.cell(row=start_row, column=cat_start_col + idx, value=ch)
        cell.font = font_header
        cell.fill = fill_cat_header
        cell.alignment = align_left if idx == 0 else align_right

    sorted_cats = sorted(category_totals.items(), key=lambda x: x[1], reverse=True)
    cat_row = start_row + 1

    for cat_name, cat_sum in sorted_cats:
        percent = (cat_sum / total_amount * 100) if total_amount > 0 else 0
        r1 = ws.cell(row=cat_row, column=cat_start_col, value=cat_name)
        r2 = ws.cell(row=cat_row, column=cat_start_col + 1, value=cat_sum)
        r3 = ws.cell(row=cat_row, column=cat_start_col + 2, value=f"{percent:.1f}%")

        r1.alignment = align_left
        r2.alignment = align_right
        r3.alignment = align_right

        r2.number_format = '#,##0'

        for cell in [r1, r2, r3]:
            cell.font = font_data
            cell.border = border_data

        cat_row += 1

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or "")
            if "\n" in val:
                val = max(val.split("\n"), key=len)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Specific fixed minimums for good look
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 26
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 14

    # File path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"harajatlar_{user_id}_{timestamp}.xlsx"
    file_path = os.path.join(REPORTS_DIR, filename)
    wb.save(file_path)
    return file_path
