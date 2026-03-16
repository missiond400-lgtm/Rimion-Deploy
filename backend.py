"""
RIMION Backend API
Antigravity Full Stack Deployment
Serves the HTML frontend and provides API endpoints.
"""

import os
import io
import logging
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="RIMION API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Serve frontend ────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def root():
    try:
        with open("RIMION_FINAL_PROPER.html", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return HTMLResponse("<h2>RIMION API running. Upload RIMION_FINAL_PROPER.html to serve the UI.</h2>")

# ── Health check ──────────────────────────────────────────────────────────────
@app.get("/health")
async def health():
    return {"status": "ok", "service": "RIMION API"}

# ── Audit endpoint ────────────────────────────────────────────────────────────
@app.post("/api/audit")
async def audit_file(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=False)
        ws = wb.active

        issues = []
        total_rows = ws.max_row
        total_cols = ws.max_column

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    issues.append({
                        "cell": cell.coordinate,
                        "issue": "Empty cell",
                        "severity": "low"
                    })
                elif isinstance(cell.value, str):
                    try:
                        float(cell.value.replace(',', '').replace('৳', '').strip())
                        issues.append({
                            "cell": cell.coordinate,
                            "issue": "Number stored as text",
                            "severity": "medium"
                        })
                    except ValueError:
                        pass

        return JSONResponse({
            "status": "success",
            "stats": {
                "rows": total_rows,
                "columns": total_cols,
                "issues_found": len(issues),
            },
            "issues": issues[:50]  # cap at 50 for response size
        })
    except Exception as e:
        logger.error(f"Audit error: {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)

# ── Report generation endpoint ────────────────────────────────────────────────
@app.post("/api/generate/{report_type}")
async def generate_report(report_type: str):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = report_type.upper()

        thin = Side(style='thin', color='CCCCCC')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws['A1'] = f'RIMION — {report_type.upper()} REPORT'
        ws['A1'].font = Font(bold=True, size=14, color='FFD700', name='Calibri')
        ws['A1'].fill = PatternFill('solid', fgColor='1A1A2E')
        ws['A1'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A1:E1')

        ws['A3'] = 'Use the Telegram bot commands for fully formatted reports:'
        ws['A3'].font = Font(italic=True, size=10, color='64748B', name='Calibri')
        ws['A5'] = '/pl — P&L Statement'
        ws['A6'] = '/salary — Salary Sheet'
        ws['A7'] = '/emi — EMI Calculator'
        ws['A8'] = '/dashboard — KPI Dashboard'
        ws['A9'] = '/cashflow — Cash Flow Statement'
        ws['A10'] = '/kpi — KPI Tracker'

        for r in range(5, 11):
            ws[f'A{r}'].font = Font(bold=True, size=10, name='Calibri')

        ws.column_dimensions['A'].width = 40

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="RIMION_{report_type}.xlsx"'}
        )
    except Exception as e:
        logger.error(f"Generate error: {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    logger.info(f"Starting RIMION on port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
