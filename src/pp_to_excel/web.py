"""
FastAPI web application for:
- Module 1: SMT Pick & Place to Excel/AOI Automation
- Module 2: Prepared Excel / .AOI vs Canias BOM Verification
- Module 3: Customer Raw P&P vs Prepared .AOI / Excel Verification
"""
from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from threading import Lock
from typing import List, Optional, Dict, Any
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from .bom_parser import load_canias_bom, extract_erp_prefixes, resolve_pnp_file_prefix, suggest_prefix_for_filename
from .pnp_parser import load_pnp_file
from .matcher_pipeline import run_matching_pipeline
from .validator_engine import validate_prepared_excel_against_bom
from .pnp_verifier import verify_pnp_against_prepared, export_pnp_verification_report
from .exporter import (
    export_final_excel,
    export_aoi_file,
    export_validation_report_excel
)

_STATIC_DIR = Path(__file__).parent / "static"
_jobs: dict[str, dict] = {}
_jobs_lock = Lock()
_RESULT_TTL = timedelta(hours=3)

app = FastAPI(title="SMT Pick & Place to Excel/AOI and Canias BOM Verification")

if _STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")


def _cleanup_expired_jobs():
    now = datetime.now(timezone.utc)
    with _jobs_lock:
        expired = [jid for jid, info in _jobs.items() if now - info["created_at"] > _RESULT_TTL]
        for jid in expired:
            td = _jobs[jid].get("temp_dir")
            if td and os.path.exists(td):
                try:
                    shutil.rmtree(td, ignore_errors=True)
                except Exception:
                    pass
            del _jobs[jid]


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_file = _STATIC_DIR / "index.html"
    if not index_file.exists():
        return HTMLResponse("<h3>Arayüz dosyası hazırlanıyor, lütfen bekleyin...</h3>", status_code=200)
    with open(index_file, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.post("/api/analyze-bom")
async def analyze_bom(bom_file: UploadFile = File(...)):
    """
    Parses uploaded Canias BOM and returns component statistics and detected prefixes.
    """
    temp_dir = tempfile.mkdtemp()
    bom_path = os.path.join(temp_dir, bom_file.filename or "bom.xlsx")
    try:
        with open(bom_path, "wb") as f:
            f.write(await bom_file.read())

        items, ref_map, item_by_code = load_canias_bom(bom_path)
        refs_by_pfx = extract_erp_prefixes(items)
        valid_prefixes = set(refs_by_pfx.keys())

        smd_items = [i for i in items if i.is_smd]
        non_smd_items = [i for i in items if not i.is_smd]

        total_smd_refs = sum(len(i.designator_list) for i in smd_items)

        pfx_dict = {p: list(refs_by_pfx.get(p, set()))[:5] for p in valid_prefixes}

        return {
            "success": True,
            "filename": bom_file.filename,
            "total_items": len(items),
            "smd_item_count": len(smd_items),
            "smd_ref_count": total_smd_refs,
            "non_smd_count": len(non_smd_items),
            "detected_prefixes": pfx_dict,
            "non_smd_samples": [{"canias_code": i.canias_code, "reason": i.skip_reason, "desc": i.description} for i in non_smd_items[:10]]
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Canias BOM analiz hatası: {str(e)}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.post("/api/generate")
async def generate_pipeline(
    bom_file: UploadFile = File(...),
    pnp_files: List[UploadFile] = File(...),
    prefixes_json: str = Form("{}"),
    project_name: str = Form("Proje"),
    include_barcode: bool = Form(True),
    include_description: bool = Form(True),
):
    """
    Module 1: Merges Canias BOM and P&P files into Final Excel and .AOI files.
    """
    _cleanup_expired_jobs()
    temp_dir = tempfile.mkdtemp()
    job_id = str(uuid4())

    try:
        # Save BOM
        bom_safe_name = os.path.basename(bom_file.filename or "bom.xlsx")
        bom_path = os.path.join(temp_dir, bom_safe_name)
        with open(bom_path, "wb") as f:
            f.write(await bom_file.read())

        bom_items, ref_map, item_by_code = load_canias_bom(bom_path)
        refs_by_pfx = extract_erp_prefixes(bom_items)
        valid_prefixes = set(refs_by_pfx.keys())

        prefix_mapping = {}
        try:
            prefix_mapping = json.loads(prefixes_json)
        except Exception:
            pass

        all_orig_names = [pf.filename or f"pnp_{idx}.txt" for idx, pf in enumerate(pnp_files)]

        # Parse all P&P files
        all_pnp_rows = []
        parsed_files_info = []

        for idx, pf in enumerate(pnp_files):
            orig_name = pf.filename or f"pnp_{idx}.txt"
            pf_safe_name = f"{idx}_{os.path.basename(orig_name)}"
            pf_path = os.path.join(temp_dir, pf_safe_name)
            with open(pf_path, "wb") as f:
                f.write(await pf.read())

            if orig_name in prefix_mapping and prefix_mapping[orig_name] is not None:
                pfx = prefix_mapping[orig_name].strip()
            else:
                pfx = resolve_pnp_file_prefix(orig_name, set(), valid_prefixes, refs_by_pfx, all_filenames=all_orig_names)

            rows = load_pnp_file(pf_path, prefix=pfx)
            all_pnp_rows.extend(rows)
            parsed_files_info.append({
                "filename": orig_name,
                "prefix": pfx,
                "row_count": len(rows)
            })

        # Run matching
        top_rows, bottom_rows, report = run_matching_pipeline(
            all_pnp_rows, bom_items, ref_map, item_by_code
        )

        # Generate output files
        safe_proj = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in project_name) or "Proje"

        final_excel_path = os.path.join(temp_dir, f"{safe_proj}_Final.xlsx")
        top_aoi_path = os.path.join(temp_dir, f"{safe_proj}_Top.AOI")
        bot_aoi_path = os.path.join(temp_dir, f"{safe_proj}_Bottom.AOI")
        zip_path = os.path.join(temp_dir, f"{safe_proj}_Ciktilar.zip")

        export_final_excel(
            top_rows, bottom_rows, final_excel_path,
            include_barcode=include_barcode, include_description=include_description
        )
        export_aoi_file(
            top_rows, top_aoi_path,
            include_barcode=include_barcode, include_description=include_description
        )
        export_aoi_file(
            bottom_rows, bot_aoi_path,
            include_barcode=include_barcode, include_description=include_description
        )

        # Create ZIP bundle
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(final_excel_path, arcname=os.path.basename(final_excel_path))
            zf.write(top_aoi_path, arcname=os.path.basename(top_aoi_path))
            zf.write(bot_aoi_path, arcname=os.path.basename(bot_aoi_path))

        with _jobs_lock:
            _jobs[job_id] = {
                "created_at": datetime.now(timezone.utc),
                "temp_dir": temp_dir,
                "files": {
                    "final_excel": final_excel_path,
                    "top_aoi": top_aoi_path,
                    "bottom_aoi": bot_aoi_path,
                    "zip_all": zip_path
                }
            }

        return {
            "success": True,
            "job_id": job_id,
            "project_name": safe_proj,
            "files_parsed": parsed_files_info,
            "report": {
                "total_pnp": report.total_pnp_rows,
                "top_count": report.top_placed_count,
                "bottom_count": report.bottom_placed_count,
                "unmatched_count": report.unmatched_pnp_count,
                "skipped_non_smd_count": report.skipped_non_smd_count,
                "bom_missing_count": report.bom_missing_in_pnp_count,
                "skipped_samples": [{"ref": p.designator, "reason": r} for p, r in report.skipped_non_smd_items[:25]],
                "unmatched_samples": [p.designator for p in report.unmatched_pnp_items[:25]],
                "bom_missing_samples": report.bom_missing_in_pnp_items[:25],
            },
            "top_preview": [r.to_row_values(include_barcode, include_description) for r in top_rows[:10]],
            "bottom_preview": [r.to_row_values(include_barcode, include_description) for r in bottom_rows[:10]]
        }
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=f"İşlem hatası: {str(e)}")


@app.post("/api/validate")
async def validate_excel(
    bom_file: UploadFile = File(...),
    excel_files: List[UploadFile] = File(...),
    project_name: str = Form("Proje")
):
    """
    Module 2: Validates prepared Excel / .AOI file(s) against the current Canias ERP BOM.
    """
    _cleanup_expired_jobs()
    temp_dir = tempfile.mkdtemp()
    job_id = str(uuid4())

    try:
        # Save BOM
        bom_path = os.path.join(temp_dir, bom_file.filename or "bom.xlsx")
        with open(bom_path, "wb") as f:
            f.write(await bom_file.read())

        bom_items, ref_map, item_by_code = load_canias_bom(bom_path)

        # Save prepared files (.xlsx, .AOI, .txt, .csv)
        saved_prep_paths = []
        for idx, ef in enumerate(excel_files):
            ef_safe_name = f"{idx}_{os.path.basename(ef.filename or 'prep.xlsx')}"
            ef_path = os.path.join(temp_dir, ef_safe_name)
            with open(ef_path, "wb") as f:
                f.write(await ef.read())
            saved_prep_paths.append(ef_path)

        # Run validation
        summary, repaired_top, repaired_bottom = validate_prepared_excel_against_bom(
            saved_prep_paths, bom_items, ref_map, item_by_code
        )

        safe_proj = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in project_name) or "Proje"

        # Generate report and repaired files
        report_path = os.path.join(temp_dir, f"{safe_proj}_BOM_Karsilastirma_Raporu.xlsx")
        repaired_excel_path = os.path.join(temp_dir, f"{safe_proj}_Guncellenmis_Final.xlsx")
        repaired_top_aoi = os.path.join(temp_dir, f"{safe_proj}_Guncellenmis_Top.AOI")
        repaired_bot_aoi = os.path.join(temp_dir, f"{safe_proj}_Guncellenmis_Bottom.AOI")
        repaired_zip_path = os.path.join(temp_dir, f"{safe_proj}_Guncellenmis_Ciktilar.zip")

        export_validation_report_excel(summary, report_path, safe_proj)
        export_final_excel(repaired_top, repaired_bottom, repaired_excel_path)
        export_aoi_file(repaired_top, repaired_top_aoi)
        export_aoi_file(repaired_bottom, repaired_bot_aoi)

        with zipfile.ZipFile(repaired_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(repaired_excel_path, arcname=os.path.basename(repaired_excel_path))
            zf.write(repaired_top_aoi, arcname=os.path.basename(repaired_top_aoi))
            zf.write(repaired_bot_aoi, arcname=os.path.basename(repaired_bot_aoi))

        with _jobs_lock:
            _jobs[job_id] = {
                "created_at": datetime.now(timezone.utc),
                "temp_dir": temp_dir,
                "files": {
                    "validation_report": report_path,
                    "repaired_excel": repaired_excel_path,
                    "repaired_zip": repaired_zip_path
                }
            }

        differences_data = [
            {
                "status": d.status,
                "side": d.side,
                "refdes": d.refdes,
                "category": d.category,
                "detail": d.detail,
                "excel_val": d.excel_value,
                "bom_val": d.bom_value,
                "row": d.excel_row or "-"
            }
            for d in summary.differences
        ]

        return {
            "success": True,
            "job_id": job_id,
            "project_name": safe_proj,
            "summary": {
                "total_checked": summary.total_checked_refdes,
                "match_count": summary.match_count,
                "error_count": summary.error_count,
                "warning_count": summary.warning_count,
                "missing_in_excel": summary.missing_in_excel_count,
                "extra_in_excel": summary.extra_in_excel_count,
                "accuracy_percent": summary.accuracy_percent,
            },
            "differences": differences_data
        }
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=f"Karşılaştırma hatası: {str(e)}")


@app.post("/api/verify-pnp")
async def verify_pnp_endpoint(
    customer_pnp_files: List[UploadFile] = File(...),
    prepared_files: List[UploadFile] = File(...),
    bom_file: Optional[UploadFile] = File(None),
    project_name: str = Form("PnP_Dogrulama")
):
    """
    Module 3: Compares Customer Raw P&P file(s) against Our Prepared .AOI / Final Excel file(s).
    """
    _cleanup_expired_jobs()
    temp_dir = tempfile.mkdtemp()
    job_id = str(uuid4())

    try:
        # Save BOM if provided
        bom_path = None
        if bom_file and bom_file.filename:
            bom_path = os.path.join(temp_dir, os.path.basename(bom_file.filename))
            with open(bom_path, "wb") as f:
                f.write(await bom_file.read())

        # Save customer P&P files
        saved_cust_paths = []
        for idx, cf in enumerate(customer_pnp_files):
            cf_safe = f"cust_{idx}_{os.path.basename(cf.filename or 'pnp.txt')}"
            cf_path = os.path.join(temp_dir, cf_safe)
            with open(cf_path, "wb") as f:
                f.write(await cf.read())
            saved_cust_paths.append(cf_path)

        # Save prepared files
        saved_prep_paths = []
        for idx, pf in enumerate(prepared_files):
            pf_safe = f"prep_{idx}_{os.path.basename(pf.filename or 'prep.xlsx')}"
            pf_path = os.path.join(temp_dir, pf_safe)
            with open(pf_path, "wb") as f:
                f.write(await pf.read())
            saved_prep_paths.append(pf_path)

        # Execute verification
        summary = verify_pnp_against_prepared(
            customer_pnp_paths=saved_cust_paths,
            prepared_paths=saved_prep_paths,
            bom_path=bom_path
        )

        safe_proj = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in project_name) or "PnP_Dogrulama"
        report_path = os.path.join(temp_dir, f"{safe_proj}_Musteri_PnP_Karsilastirma_Raporu.xlsx")
        export_pnp_verification_report(summary, report_path, safe_proj)

        with _jobs_lock:
            _jobs[job_id] = {
                "created_at": datetime.now(timezone.utc),
                "temp_dir": temp_dir,
                "files": {
                    "pnp_verification_report": report_path
                }
            }

        diffs_data = [
            {
                "status": d.status,
                "side": d.side,
                "refdes": d.refdes,
                "category": d.category,
                "detail": d.detail,
                "cust_val": d.customer_value,
                "prep_val": d.prepared_value,
                "file": os.path.basename(d.customer_file) if d.customer_file else ""
            }
            for d in summary.differences
        ]

        return {
            "success": True,
            "job_id": job_id,
            "project_name": safe_proj,
            "summary": {
                "total_customer": summary.total_customer_components,
                "total_prepared": summary.total_prepared_components,
                "match_count": summary.match_count,
                "coord_error_count": summary.coord_error_count,
                "angle_error_count": summary.angle_error_count,
                "layer_error_count": summary.layer_error_count,
                "missing_in_prepared": summary.missing_in_prepared_count,
                "extra_in_prepared": summary.extra_in_prepared_count,
                "skipped_non_smd": summary.skipped_non_smd_count,
                "accuracy_percent": summary.accuracy_percent
            },
            "differences": diffs_data
        }
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=f"P&P Doğrulama hatası: {str(e)}")


@app.get("/api/download/{job_id}/{file_type}")
async def download_file(job_id: str, file_type: str):
    """
    Downloads generated files.
    """
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Oturum süresi dolmuş veya dosya bulunamadı.")

    file_path = job.get("files", {}).get(file_type)
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"İstenen dosya ({file_type}) bulunamadı.")

    filename = os.path.basename(file_path)
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/octet-stream"
    )
