"""
P&P Verifier (Module 3).
Compares Customer Raw Pick & Place file(s) (.txt, .csv, .xlsx) against Our Prepared .AOI / Final Excel file(s).
Verifies:
- All customer components were transferred without omission (RefDes completeness).
- Exact X coordinate precision (after mil->mm and .000 normalization).
- Exact Y coordinate precision (after mil->mm and .000 normalization).
- Exact Rotation angle (Allegro 1000x correction).
- Exact Layer / Side placement (Top vs Bottom).
- Identifies legitimately excluded non-SMD components (TH, Krem Lehim, Dummy) if BOM is provided.
"""
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Set, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .models import BomItem
from .bom_parser import load_canias_bom, extract_erp_prefixes, resolve_pnp_file_prefix, lookup_refdes_in_bom
from .pnp_parser import load_pnp_file, format_coord_3dec, clean_angle
from .validator_engine import read_prepared_files
from .matcher_pipeline import natural_sort_key


@dataclass
class PnpVerificationDifference:
    """Represents a discrepancy between Customer P&P and Prepared .AOI/Excel."""
    refdes: str
    side: str                          # 'Top', 'Bottom', or 'Genel'
    status: str                        # 'UYGUN', 'HATA', 'BILGI'
    category: str                      # 'X_KOORDINAT', 'Y_KOORDINAT', 'ROTATION', 'KATMAN', 'EKSIK_REFDES', 'FAZLA_REFDES', 'SMD_DISI'
    detail: str                        # Explanation
    customer_value: str
    prepared_value: str
    customer_file: str = ""
    prepared_file: str = ""


@dataclass
class PnpVerificationSummary:
    """Summary KPI statistics for P&P vs Prepared file verification."""
    total_customer_components: int = 0
    total_prepared_components: int = 0
    match_count: int = 0
    coord_error_count: int = 0
    angle_error_count: int = 0
    layer_error_count: int = 0
    missing_in_prepared_count: int = 0
    extra_in_prepared_count: int = 0
    skipped_non_smd_count: int = 0
    accuracy_percent: float = 0.0
    differences: List[PnpVerificationDifference] = field(default_factory=list)


def verify_pnp_against_prepared(
    customer_pnp_paths: List[str],
    prepared_paths: List[str],
    bom_path: Optional[str] = None
) -> PnpVerificationSummary:
    """
    Executes deep verification between Customer P&P files and Prepared .AOI / Excel files.
    """
    # 1. Read Prepared Excel / AOI rows first so their prefixes are known even without a BOM
    prepared_rows = read_prepared_files(prepared_paths)

    non_smd_ref_map: Dict[str, str] = {}
    smd_bom_refs: Set[str] = set()
    refdes_to_codes: Dict[str, List[str]] = {}
    item_by_code: Dict[str, BomItem] = {}
    erp_prefixes: Set[str] = set()
    erp_refs_by_pfx: Dict[str, Set[str]] = defaultdict(set)
    has_bom = False

    # Extract prefixes present in prepared_rows
    for pr in prepared_rows:
        ref_u = pr["refdes"].strip().upper()
        if "_" in ref_u:
            pfx, base = ref_u.split("_", 1)
            if 2 <= len(pfx) <= 10 and re.match(r"^[A-Z]+\d+", base):
                erp_refs_by_pfx[pfx].add(base)
    for pfx, bases in list(erp_refs_by_pfx.items()):
        if len(bases) >= 2:
            erp_prefixes.add(pfx)

    # 2. Optionally parse BOM to identify SMD vs non-SMD / DNP components and board prefixes
    if bom_path:
        try:
            bom_items, refdes_to_codes, item_by_code = load_canias_bom(bom_path)
            has_bom = True
            bom_refs_by_pfx = extract_erp_prefixes(bom_items)
            for pfx, bases in bom_refs_by_pfx.items():
                erp_prefixes.add(pfx)
                erp_refs_by_pfx[pfx].update(bases)
            for item in bom_items:
                if item.is_smd:
                    for r in item.designator_list:
                        r_u = r.upper()
                        smd_bom_refs.add(r_u)
                        if "_" in r_u:
                            p_u, b_u = r_u.split("_", 1)
                            if p_u.endswith(("L", "R")) and len(p_u) >= 3:
                                smd_bom_refs.add(f"{p_u[:-1]}_{b_u}")
                else:
                    for r in item.designator_list:
                        non_smd_ref_map[r.upper()] = item.skip_reason or "SMD Değil"
        except Exception:
            pass

    # 3. Parse Customer P&P rows (resolving board prefix with raw designators)
    customer_rows = []
    cust_filenames = [p for p in customer_pnp_paths]

    for p in customer_pnp_paths:
        raw_p_rows = load_pnp_file(p, prefix="")
        pnp_refs = {r.raw_designator for r in raw_p_rows}
        pfx = resolve_pnp_file_prefix(p, pnp_refs, erp_prefixes, erp_refs_by_pfx, all_filenames=cust_filenames)
        if pfx:
            clean_pfx = pfx.strip().rstrip("_")
            for r in raw_p_rows:
                r.prefix = pfx
                if not r.raw_designator.upper().startswith(f"{clean_pfx.upper()}_"):
                    r.designator = f"{clean_pfx}_{r.raw_designator}"
        customer_rows.extend(raw_p_rows)

    summary = PnpVerificationSummary(
        total_customer_components=len(customer_rows),
        total_prepared_components=len(prepared_rows)
    )

    # 4. Build multi-instance lookup map for prepared components (supporting L/R duplicate boards)
    prep_by_ref: Dict[str, List[Tuple[int, Dict[str, Any]]]] = defaultdict(list)
    prep_unprefixed: Dict[str, List[Tuple[int, Dict[str, Any]]]] = defaultdict(list)
    for idx, pr in enumerate(prepared_rows):
        ref_u = pr["refdes"].strip().upper()
        prep_by_ref[ref_u].append((idx, pr))
        if "_" in ref_u:
            pfx_u, base_r = ref_u.split("_", 1)
            if pfx_u.endswith(("L", "R")) and len(pfx_u) >= 3:
                prep_by_ref[f"{pfx_u[:-1]}_{base_r}"].append((idx, pr))
        else:
            prep_unprefixed[ref_u].append((idx, pr))

    used_prep_indices: Set[int] = set()
    matched_pairs: List[Tuple[Any, int, Dict[str, Any]]] = []

    for crow in customer_rows:
        c_ref = crow.designator.strip().upper()
        c_raw_ref = crow.raw_designator.strip().upper()

        # Check if non-SMD or absent (DNP) in BOM
        is_known_non_smd = (c_ref in non_smd_ref_map) or (c_raw_ref in non_smd_ref_map)
        skip_reason = non_smd_ref_map.get(c_ref) or non_smd_ref_map.get(c_raw_ref)

        if not is_known_non_smd and has_bom:
            codes, _ = lookup_refdes_in_bom(c_ref, refdes_to_codes, item_by_code)
            if not codes:
                is_known_non_smd = True
                skip_reason = "Canias BOM'da Yok / DNP (Dizgisiz)"
            else:
                b_item = item_by_code.get(codes[0])
                if b_item and not b_item.is_smd:
                    is_known_non_smd = True
                    skip_reason = b_item.skip_reason or "SMD Değil"

        # Gather candidate prepared rows for this customer component (strictly same board)
        candidates = list(prep_by_ref.get(c_ref, []))
        if not candidates and "_" in c_ref:
            pfx_u, base_u = c_ref.split("_", 1)
            if pfx_u.endswith(("L", "R")) and len(pfx_u) >= 3:
                candidates = list(prep_by_ref.get(f"{pfx_u[:-1]}_{base_u}", []))
            else:
                candidates = list(prep_by_ref.get(f"{pfx_u}L_{base_u}", [])) + list(prep_by_ref.get(f"{pfx_u}R_{base_u}", []))
        if not candidates:
            candidates = list(prep_unprefixed.get(c_raw_ref, []))

        matched_prep = None
        matched_idx = None
        if candidates:
            unused = [(i, pr) for i, pr in candidates if i not in used_prep_indices]
            if unused:
                try:
                    cx_f = float(format_coord_3dec(crow.center_x))
                    cy_f = float(format_coord_3dec(crow.center_y))
                except ValueError:
                    cx_f, cy_f = 0.0, 0.0

                def _cand_sort_key(item: Tuple[int, Dict[str, Any]]):
                    _, pr = item
                    side_pen = 0 if pr["side"] == crow.layer else 1
                    try:
                        px_f = float(format_coord_3dec(pr["x"]))
                        py_f = float(format_coord_3dec(pr["y"]))
                        dist = (px_f - cx_f) ** 2 + (py_f - cy_f) ** 2
                    except ValueError:
                        dist = 1e9
                    return (side_pen, dist)

                unused.sort(key=_cand_sort_key)
                matched_idx, matched_prep = unused[0]

        if not matched_prep:
            if is_known_non_smd:
                summary.skipped_non_smd_count += 1
                summary.differences.append(PnpVerificationDifference(
                    refdes=crow.designator,
                    side=crow.layer,
                    status="BILGI",
                    category="SMD_DISI",
                    detail=f"Müşteri P&P'de mevcut ancak BOM'da SMD dışı ({skip_reason}) olduğu için dosyaya alınmamış (Doğru İşlem).",
                    customer_value="Müşteri P&P'de Var",
                    prepared_value="Hazırlanan Dosyada Yok (SMD Dışı)",
                    customer_file=crow.source_file,
                    prepared_file=""
                ))
            else:
                summary.missing_in_prepared_count += 1
                summary.differences.append(PnpVerificationDifference(
                    refdes=crow.designator,
                    side=crow.layer,
                    status="HATA",
                    category="EKSIK_REFDES",
                    detail="Müşteri P&P dosyasında mevcut, ancak hazırlanan .AOI / Excel dosyasında UNUTULMUŞ / EKSİK!",
                    customer_value=f"X={crow.center_x}, Y={crow.center_y}, Açı={crow.rotation}°",
                    prepared_value="Hazırlanan Dosyada Yok!",
                    customer_file=crow.source_file,
                    prepared_file=""
                ))
            continue

        used_prep_indices.add(matched_idx)
        matched_pairs.append((crow, matched_idx, matched_prep))

    # 5. Detect board/side constant Gerber panel offsets (when >= 5 parts and >= 80% share exact same non-zero delta)
    group_dx: Dict[Tuple[str, str], List[float]] = defaultdict(list)
    group_dy: Dict[Tuple[str, str], List[float]] = defaultdict(list)
    for crow, _, matched_prep in matched_pairs:
        grp_key = (crow.prefix or "", crow.layer)
        try:
            dx = round(float(format_coord_3dec(matched_prep["x"])) - float(format_coord_3dec(crow.center_x)), 3)
            dy = round(float(format_coord_3dec(matched_prep["y"])) - float(format_coord_3dec(crow.center_y)), 3)
            group_dx[grp_key].append(dx)
            group_dy[grp_key].append(dy)
        except ValueError:
            pass

    panel_offsets: Dict[Tuple[str, str], Tuple[float, float]] = {}
    for grp_key, dxs in group_dx.items():
        dys = group_dy[grp_key]
        if len(dxs) >= 5:
            top_dx, count_dx = Counter(dxs).most_common(1)[0]
            top_dy, count_dy = Counter(dys).most_common(1)[0]
            eff_dx = top_dx if (count_dx / len(dxs)) >= 0.80 else 0.0
            eff_dy = top_dy if (count_dy / len(dys)) >= 0.80 else 0.0
            if eff_dx != 0.0 or eff_dy != 0.0:
                panel_offsets[grp_key] = (eff_dx, eff_dy)
                board_lbl = grp_key[0] or "Ana Kart"
                summary.differences.append(PnpVerificationDifference(
                    refdes=f"{board_lbl} ({grp_key[1]})",
                    side=grp_key[1],
                    status="BILGI",
                    category="PANEL_OFFSET",
                    detail=f"Sabit Gerber/Panel koordinat ofseti tespit edildi: dX={eff_dx:+.3f} mm, dY={eff_dy:+.3f} mm ({len(dxs)} komponent).",
                    customer_value="Orijinal CAD Orjini",
                    prepared_value=f"dX={eff_dx:+.3f}, dY={eff_dy:+.3f} mm",
                    customer_file="",
                    prepared_file=""
                ))

    # 6. Compare Coordinates, Angle, and Layer for each matched pair
    for crow, _, matched_prep in matched_pairs:
        has_coord_mismatch = False
        has_angle_mismatch = False
        has_layer_mismatch = False

        cust_x = format_coord_3dec(crow.center_x)
        cust_y = format_coord_3dec(crow.center_y)
        prep_x = format_coord_3dec(matched_prep["x"])
        prep_y = format_coord_3dec(matched_prep["y"])

        cust_rot = clean_angle(crow.rotation)
        prep_rot = clean_angle(matched_prep["rotation"])

        cust_side = crow.layer
        prep_side = matched_prep["side"]

        grp_key = (crow.prefix or "", cust_side)
        off_x, off_y = panel_offsets.get(grp_key, (0.0, 0.0))

        expected_x = format_coord_3dec(str(float(cust_x) + off_x)) if off_x != 0.0 else cust_x
        expected_y = format_coord_3dec(str(float(cust_y) + off_y)) if off_y != 0.0 else cust_y

        # Check X
        if expected_x != prep_x:
            has_coord_mismatch = True
            summary.coord_error_count += 1
            summary.differences.append(PnpVerificationDifference(
                refdes=crow.designator,
                side=cust_side,
                status="HATA",
                category="X_KOORDINAT",
                detail=f"Center-X koordinatı uyuşmuyor: Müşteri={cust_x} mm vs Hazırlanan={prep_x} mm",
                customer_value=cust_x,
                prepared_value=prep_x,
                customer_file=crow.source_file,
                prepared_file=matched_prep.get("source_file", "")
            ))

        # Check Y
        if expected_y != prep_y:
            has_coord_mismatch = True
            summary.coord_error_count += 1
            summary.differences.append(PnpVerificationDifference(
                refdes=crow.designator,
                side=cust_side,
                status="HATA",
                category="Y_KOORDINAT",
                detail=f"Center-Y koordinatı uyuşmuyor: Müşteri={cust_y} mm vs Hazırlanan={prep_y} mm",
                customer_value=cust_y,
                prepared_value=prep_y,
                customer_file=crow.source_file,
                prepared_file=matched_prep.get("source_file", "")
            ))

        # Check Rotation
        if cust_rot != prep_rot:
            has_angle_mismatch = True
            summary.angle_error_count += 1
            summary.differences.append(PnpVerificationDifference(
                refdes=crow.designator,
                side=cust_side,
                status="HATA",
                category="ROTATION",
                detail=f"Dönme açısı (Rotation) uyuşmuyor: Müşteri={cust_rot}° vs Hazırlanan={prep_rot}°",
                customer_value=f"{cust_rot}°",
                prepared_value=f"{prep_rot}°",
                customer_file=crow.source_file,
                prepared_file=matched_prep.get("source_file", "")
            ))

        # Check Layer / Side
        if cust_side != prep_side:
            has_layer_mismatch = True
            summary.layer_error_count += 1
            summary.differences.append(PnpVerificationDifference(
                refdes=crow.designator,
                side=cust_side,
                status="HATA",
                category="KATMAN",
                detail=f"Yüzey/Katman uyuşmuyor: Müşteri P&P'de '{cust_side}' katmanında, ancak hazırlanan dosyada '{prep_side}' sayfasına/dosyasına konulmuş!",
                customer_value=cust_side,
                prepared_value=prep_side,
                customer_file=crow.source_file,
                prepared_file=matched_prep.get("source_file", "")
            ))

        if not has_coord_mismatch and not has_angle_mismatch and not has_layer_mismatch:
            summary.match_count += 1

    # 7. Check for components present in prepared file but NOT in customer P&P
    for idx, pr in enumerate(prepared_rows):
        if idx not in used_prep_indices:
            summary.extra_in_prepared_count += 1
            summary.differences.append(PnpVerificationDifference(
                refdes=pr["refdes"],
                side=pr["side"],
                status="HATA",
                category="FAZLA_REFDES",
                detail="Hazırlanan .AOI / Excel dosyasında yer alıyor, ancak Müşteri P&P dosyasında bulunamadı!",
                customer_value="Müşteri P&P'de Yok",
                prepared_value=f"X={pr['x']}, Y={pr['y']}, Canias={pr['code']}",
                customer_file="",
                prepared_file=pr.get("source_file", "")
            ))

    # Calculate accuracy
    total_active_checks = summary.total_customer_components - summary.skipped_non_smd_count + summary.extra_in_prepared_count
    if total_active_checks > 0:
        summary.accuracy_percent = round((summary.match_count / total_active_checks) * 100, 2)
    else:
        summary.accuracy_percent = 100.0

    return summary


def export_pnp_verification_report(
    summary: PnpVerificationSummary,
    output_path: str,
    project_name: str = "Proje"
) -> str:
    """Exports colorized Excel report comparing Customer P&P with Prepared Excel/AOI."""
    wb = openpyxl.Workbook()

    # Sheet 1: Dashboard
    ws_dash = wb.active
    ws_dash.title = "P&P Doğrulama Özeti"
    ws_dash.views.sheetView[0].showGridLines = True

    ws_dash.merge_cells("A1:F2")
    title_cell = ws_dash["A1"]
    title_cell.value = "MÜŞTERİ PICK & PLACE <-> HAZIRLANAN DOSYA KARŞILAŞTIRMA RAPORU"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    kpis = [
        ("Müşteri P&P Toplam Komponent", summary.total_customer_components, "D9E1F2"),
        ("Hazırlanan Dosyadaki Komponent", summary.total_prepared_components, "D9E1F2"),
        ("Birebir Doğru ve Eksiksiz Aktarılan", summary.match_count, "E2EFDA"),
        ("Koordinat (X/Y) Hatası", summary.coord_error_count, "FCE4D6" if summary.coord_error_count else "E2EFDA"),
        ("Açı (Rotation) Hatası", summary.angle_error_count, "FCE4D6" if summary.angle_error_count else "E2EFDA"),
        ("Katman / Yüzey (Top/Bot) Hatası", summary.layer_error_count, "FCE4D6" if summary.layer_error_count else "E2EFDA"),
        ("Hazırlanan Dosyada Unutulan (Eksik)", summary.missing_in_prepared_count, "F8CBAD" if summary.missing_in_prepared_count else "E2EFDA"),
        ("Müşteri P&P'de Olmayan (Fazla)", summary.extra_in_prepared_count, "F8CBAD" if summary.extra_in_prepared_count else "E2EFDA"),
        ("SMD Dışı (Doğru Elenen)", summary.skipped_non_smd_count, "FFF2CC"),
        ("Genel Aktarım Doğruluk Skoru", f"%{summary.accuracy_percent}", "D9E1F2"),
    ]

    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for idx, (label, val, col_hex) in enumerate(kpis, start=4):
        ws_dash.cell(row=idx, column=2, value=label).font = Font(name="Calibri", size=11, bold=True)
        val_cell = ws_dash.cell(row=idx, column=4, value=val)
        val_cell.font = Font(name="Calibri", size=12, bold=True)
        val_cell.alignment = Alignment(horizontal="center")
        val_cell.fill = PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid")
        val_cell.border = border_thin

    ws_dash.column_dimensions["B"].width = 36
    ws_dash.column_dimensions["C"].width = 5
    ws_dash.column_dimensions["D"].width = 24

    # Sheet 2: Details
    ws_det = wb.create_sheet(title="Hata ve Fark Detayları")
    ws_det.views.sheetView[0].showGridLines = True

    headers = [
        "Durum", "Katman", "Yerleşim Yeri", "Kategori",
        "Hata / Fark Nedeni", "Müşteri P&P Değeri", "Hazırlanan Dosya Değeri", "Müşteri Dosyası"
    ]
    ws_det.append(headers)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_num in range(1, len(headers) + 1):
        c = ws_det.cell(row=1, column=col_num)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    fill_error = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
    font_error = Font(color="C5221F", bold=True)
    fill_info = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")
    font_info = Font(color="1A73E8", bold=True)

    for diff in summary.differences:
        r_vals = [
            diff.status,
            diff.side,
            diff.refdes,
            diff.category,
            diff.detail,
            diff.customer_value,
            diff.prepared_value,
            diff.customer_file or "-"
        ]
        ws_det.append(r_vals)
        cur_r = ws_det.max_row
        scell = ws_det.cell(row=cur_r, column=1)
        if diff.status == "HATA":
            scell.fill = fill_error
            scell.font = font_error
        else:
            scell.fill = fill_info
            scell.font = font_info

    for col in ws_det.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_det.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    ws_det.auto_filter.ref = ws_det.dimensions
    ws_det.freeze_panes = "A2"

    wb.save(output_path)
    return output_path
