"""
Exporters for:
- Headerless Final Excel (.xlsx) with Top & Bottom sheets
- Tab-delimited CRLF .AOI files (Top.AOI, Bottom.AOI)
- Professional Colorized BOM Validation Report Excel (.xlsx)
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Optional
from .models import MatchedRow, ValidationSummary, ValidationDifference


def export_final_excel(
    top_rows: List[MatchedRow],
    bottom_rows: List[MatchedRow],
    output_path: str,
    include_barcode: bool = True,
    include_description: bool = True
) -> str:
    """
    Exports the final headerless Excel with 'Top' and 'Bottom' sheets.
    Strictly starts at row 1 without headers.
    Preserves exact dot-separated coordinate strings without rounding.
    """
    wb = openpyxl.Workbook()
    # Setup Top sheet
    ws_top = wb.active
    ws_top.title = "Top"

    for r_idx, row in enumerate(top_rows, start=1):
        row_vals = row.to_row_values(include_barcode=include_barcode, include_description=include_description)
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_top.cell(row=r_idx, column=c_idx)
            cell.value = str(val)

    # Setup Bottom sheet
    ws_bot = wb.create_sheet(title="Bottom")
    for r_idx, row in enumerate(bottom_rows, start=1):
        row_vals = row.to_row_values(include_barcode=include_barcode, include_description=include_description)
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_bot.cell(row=r_idx, column=c_idx)
            cell.value = str(val)

    wb.save(output_path)
    return output_path


def export_aoi_file(
    rows: List[MatchedRow],
    output_path: str,
    encoding: str = "cp1254",
    include_barcode: bool = True,
    include_description: bool = True
) -> str:
    """
    Exports tab-delimited CRLF .AOI file matching the machine format.
    No header row; starts at line 1.
    """
    lines = []
    for row in rows:
        row_vals = row.to_row_values(include_barcode=include_barcode, include_description=include_description)
        line_str = "\t".join(row_vals)
        lines.append(line_str)
    
    content = "\r\n".join(lines) + "\r\n" if lines else ""
    try:
        with open(output_path, "w", encoding=encoding, errors="replace") as f:
            f.write(content)
    except Exception:
        with open(output_path, "w", encoding="utf-8-sig") as f:
            f.write(content)

    return output_path


def export_validation_report_excel(
    summary: ValidationSummary,
    output_path: str,
    project_name: str = "Proje"
) -> str:
    """
    Generates a colorized, professional BOM Validation Report.
    """
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Sheet 1: Summary Dashboard
    # -------------------------------------------------------------
    ws_summary = wb.active
    ws_summary.title = "Doğrulama Özeti"
    ws_summary.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_summary.merge_cells("A1:F2")
    title_cell = ws_summary["A1"]
    title_cell.value = f"CANİAS BOM & P&P EXCEL DOĞRULAMA VE DEĞİŞİKLİK RAPORU"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # KPI Box Styling
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    kpis = [
        ("Toplam Denetlenen RefDes", summary.total_checked_refdes, "D9E1F2"),
        ("Birebir Uyumlu (Doğru)", summary.match_count, "E2EFDA"),
        ("Kritik Hata (Değişiklik)", summary.error_count, "FCE4D6"),
        ("Uyarı (Biçim / Semantik)", summary.warning_count, "FFF2CC"),
        ("Excel'de Eksik RefDes", summary.missing_in_excel_count, "F8CBAD"),
        ("Genel Doğruluk Oranı", f"%{summary.accuracy_percent}", "D9E1F2")
    ]

    for idx, (label, val, col_hex) in enumerate(kpis, start=4):
        ws_summary.cell(row=idx, column=2, value=label).font = Font(name="Calibri", size=11, bold=True)
        val_cell = ws_summary.cell(row=idx, column=4, value=val)
        val_cell.font = Font(name="Calibri", size=12, bold=True)
        val_cell.alignment = Alignment(horizontal="center")
        val_cell.fill = PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid")
        val_cell.border = border_thin

    ws_summary.column_dimensions["B"].width = 30
    ws_summary.column_dimensions["C"].width = 5
    ws_summary.column_dimensions["D"].width = 20

    # -------------------------------------------------------------
    # Sheet 2: Detailed Differences
    # -------------------------------------------------------------
    ws_details = wb.create_sheet(title="Fark ve Hata Detayları")
    ws_details.views.sheetView[0].showGridLines = True

    headers = [
        "Durum", "Katman / Sayfa", "Yerleşim Yeri", "Kategori",
        "Hata / Fark Açıklaması", "Excel'deki Değer", "Güncel Canias BOM Değeri", "Excel Satırı"
    ]
    
    ws_details.append(headers)
    
    # Header format
    header_fill = PatternFill(start_color="173B32", end_color="173B32", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_num in range(1, len(headers) + 1):
        cell = ws_details.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Status formatting
    fill_error = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
    font_error = Font(color="C5221F", bold=True)
    fill_warning = PatternFill(start_color="FEF7E0", end_color="FEF7E0", fill_type="solid")
    font_warning = Font(color="B06000", bold=True)
    fill_ok = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    font_ok = Font(color="137333", bold=True)

    for diff in summary.differences:
        r_vals = [
            diff.status,
            diff.side,
            diff.refdes,
            diff.category,
            diff.detail,
            diff.excel_value,
            diff.bom_value,
            diff.excel_row or "-"
        ]
        ws_details.append(r_vals)
        cur_row = ws_details.max_row
        
        status_cell = ws_details.cell(row=cur_row, column=1)
        if diff.status == "HATA":
            status_cell.fill = fill_error
            status_cell.font = font_error
        elif diff.status == "UYARI":
            status_cell.fill = fill_warning
            status_cell.font = font_warning
        else:
            status_cell.fill = fill_ok
            status_cell.font = font_ok

    # Auto-fit column widths
    for col in ws_details.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_details.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    ws_details.auto_filter.ref = ws_details.dimensions
    ws_details.freeze_panes = "A2"

    wb.save(output_path)
    return output_path
