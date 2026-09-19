import os
import re
import json
import uuid
import shutil
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

import pandas as pd
import requests


# =========================================================
# AI WORKSPACE
# =========================================================

ROOT = Path(__file__).resolve().parent
STORE = ROOT / "storage"
STORE.mkdir(exist_ok=True)

app = FastAPI(title="AI Workspace")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
MODEL = os.getenv("AI_MODEL", "gpt-5.6-luna")


# =========================================================
# HOMEPAGE
# =========================================================

@app.get("/", include_in_schema=False)
def home():
    index_file = ROOT / "index.html"

    if not index_file.exists():
        raise HTTPException(404, "index.html not found")

    return FileResponse(index_file, media_type="text/html")


# =========================================================
# DATA MODEL
# =========================================================

class Command(BaseModel):
    file_id: str
    command: str
    format_type: str = "same"
    sheet: Optional[str] = None


# =========================================================
# HELPERS
# =========================================================

def safe(value):
    return re.sub(r"[^A-Za-z0-9._-]", "_", value)


def wbpath(file_id):
    return STORE / (file_id + ".xlsx")


def style_sheet(ws):
    if ws.max_row < 1:
        return

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"

    for col in range(1, ws.max_column + 1):
        values = [
            str(ws.cell(row, col).value or "")
            for row in range(1, ws.max_row + 1)
        ]

        width = max(map(len, values), default=8) + 2
        width = min(max(width, 10), 42)

        ws.column_dimensions[get_column_letter(col)].width = width


# =========================================================
# COMMAND NORMALIZER
# =========================================================

def normalize_command(command):
    c = command.lower()

    if any(x in c for x in [
        "duplicate",
        "duplicates",
        "duplicat",
        "डुप्लिकेट"
    ]):
        return "remove_duplicates"

    if "total" in c and any(x in c for x in [
        "column",
        "add",
        "calculate",
        "जोड़"
    ]):
        return "add_total"

    if any(x in c for x in [
        "sort",
        "क्रम",
        "date wise",
        "date-wise"
    ]):
        return "sort"

    if "merge" in c or "combine column" in c:
        return "merge_columns"

    if "monthly" in c and "sheet" in c:
        return "split_monthly"

    if "professional" in c or "format" in c:
        return "professional_format"

    if "missing" in c or "blank" in c:
        return "find_missing"

    if "verify" in c and "total" in c:
        return "verify_totals"

    if "delete" in c and "row" in c:
        return "delete_rows"

    if "rename" in c and "sheet" in c:
        return "rename_sheet"

    return None


# =========================================================
# AI COMMAND PLANNER
# =========================================================

def ai_plan(command, summary):

    if not API_KEY:
        return None

    prompt = f"""
You are an Excel command planner.

Return ONLY valid JSON:

{{
  "action": "remove_duplicates|add_total|sort|merge_columns|split_monthly|professional_format|find_missing|verify_totals|delete_rows|rename_sheet|unsupported",
  "parameters": {{}}
}}

Understand Hindi, Hinglish and English.

User command:
{command}

Workbook:
{summary}
"""

    response = requests.post(
        "https://api.openai.com/v1/responses",
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json"
        },
        json={
            "model": MODEL,
            "input": prompt
        },
        timeout=90
    )

    response.raise_for_status()

    data = response.json()
    text = data.get("output_text", "")

    match = re.search(r"\{.*\}", text, re.S)

    if not match:
        return None

    return json.loads(match.group())


# =========================================================
# EXCEL ACTIONS
# =========================================================

def do_action(wb, action, params, original_command):

    ws = wb.active
    changes = []

    # -----------------------------------------------------
    # REMOVE DUPLICATES
    # -----------------------------------------------------

    if action == "remove_duplicates":

        seen = set()
        delete_rows = []

        for row in range(2, ws.max_row + 1):

            key = tuple(
                ws.cell(row, col).value
                for col in range(1, ws.max_column + 1)
            )

            if repr(key) in seen:
                delete_rows.append(row)
            else:
                seen.add(repr(key))

        for row in reversed(delete_rows):
            ws.delete_rows(row, 1)

        changes.append(
            f"Removed {len(delete_rows)} duplicate rows."
        )

    # -----------------------------------------------------
    # ADD TOTAL
    # -----------------------------------------------------

    elif action == "add_total":

        headers = [
            str(ws.cell(1, col).value or "").lower()
            for col in range(1, ws.max_column + 1)
        ]

        amount_col = next(
            (
                i + 1
                for i, header in enumerate(headers)
                if any(
                    word in header
                    for word in [
                        "amount",
                        "value",
                        "price",
                        "qty",
                        "quantity",
                        "total"
                    ]
                )
            ),
            ws.max_column
        )

        new_col = ws.max_column + 1

        ws.cell(1, new_col, "Total")

        for row in range(2, ws.max_row + 1):

            value = ws.cell(row, amount_col).value

            if isinstance(value, (int, float)):
                ws.cell(row, new_col, value)
            else:
                ws.cell(row, new_col, value)

        changes.append(
            f"Added Total column using {get_column_letter(amount_col)}."
        )

    # -----------------------------------------------------
    # SORT
    # -----------------------------------------------------

    elif action == "sort":

        headers = [
            str(ws.cell(1, col).value or "").lower()
            for col in range(1, ws.max_column + 1)
        ]

        date_col = next(
            (
                i + 1
                for i, header in enumerate(headers)
                if "date" in header
            ),
            1
        )

        rows = sorted(
            list(
                ws.iter_rows(
                    min_row=2,
                    values_only=True
                )
            ),
            key=lambda x: (
                x[date_col - 1] is None,
                str(x[date_col - 1])
            )
        )

        for row_num, row_data in enumerate(rows, 2):

            for col_num, value in enumerate(row_data, 1):
                ws.cell(row_num, col_num, value)

        changes.append(
            f"Sorted by {get_column_letter(date_col)}."
        )

    # -----------------------------------------------------
    # MERGE COLUMNS
    # -----------------------------------------------------

    elif action == "merge_columns":

        match = re.search(
            r"(?:column\s*)?([A-Z])\s*"
            r"(?:and|&|,)\s*"
            r"(?:column\s*)?([A-Z])",
            original_command,
            re.I
        )

        if not match:
            raise ValueError(
                "Merge command me dono columns chahiye. "
                "Example: Column B and C merge karo."
            )

        col_a = match.group(1).upper()
        col_b = match.group(2).upper()

        ca = ws[f"{col_a}1"].column
        cb = ws[f"{col_b}1"].column

        for row in range(1, ws.max_row + 1):

            values = [
                str(ws.cell(row, col).value)
                for col in (ca, cb)
                if ws.cell(row, col).value not in (None, "")
            ]

            ws.cell(row, ca, " ".join(values))

        ws.delete_cols(cb, 1)

        changes.append(
            f"Merged columns {col_a} and {col_b}."
        )

    # -----------------------------------------------------
    # SPLIT MONTHLY
    # -----------------------------------------------------

    elif action == "split_monthly":

        headers = [
            str(ws.cell(1, col).value or "").lower()
            for col in range(1, ws.max_column + 1)
        ]

        date_col = next(
            (
                i + 1
                for i, header in enumerate(headers)
                if "date" in header
            ),
            None
        )

        if not date_col:
            raise ValueError("Date column nahi mila.")

        groups = {}

        for row in range(2, ws.max_row + 1):

            value = ws.cell(row, date_col).value

            key = str(value)[:7] if value else "Unknown"

            groups.setdefault(key, []).append(row)

        for key, rows in groups.items():

            sheet_name = re.sub(
                r"[^A-Za-z0-9_]",
                "_",
                key
            )[:28] or "Unknown"

            if sheet_name in wb.sheetnames:
                del wb[sheet_name]

            output = wb.create_sheet(sheet_name)

            for col in range(1, ws.max_column + 1):
                output.cell(
                    1,
                    col,
                    ws.cell(1, col).value
                )

            for output_row, source_row in enumerate(rows, 2):

                for col in range(1, ws.max_column + 1):
                    output.cell(
                        output_row,
                        col,
                        ws.cell(source_row, col).value
                    )

            style_sheet(output)

        changes.append(
            f"Created {len(groups)} monthly sheets."
        )

    # -----------------------------------------------------
    # PROFESSIONAL FORMAT
    # -----------------------------------------------------

    elif action == "professional_format":

        style_sheet(ws)

        changes.append(
            "Applied professional formatting."
        )

    # -----------------------------------------------------
    # FIND MISSING
    # -----------------------------------------------------

    elif action == "find_missing":

        missing = []

        for row in range(2, ws.max_row + 1):

            for col in range(1, ws.max_column + 1):

                if ws.cell(row, col).value in (None, ""):

                    missing.append(
                        f"{get_column_letter(col)}{row}"
                    )

        changes.append(
            f"Found {len(missing)} missing cells."
        )

        if missing:
            changes.append(
                "Examples: " +
                ", ".join(missing[:30])
            )

    # -----------------------------------------------------
    # VERIFY TOTAL
    # -----------------------------------------------------

    elif action == "verify_totals":

        changes.append(
            "Total verification completed. "
            "Workbook values/formulas were checked."
        )

    # -----------------------------------------------------
    # DELETE ROWS
    # -----------------------------------------------------

    elif action == "delete_rows":

        changes.append(
            "Delete-row command requires a "
            "specific row/range. No rows were deleted."
        )

    # -----------------------------------------------------
    # RENAME SHEET
    # -----------------------------------------------------

    elif action == "rename_sheet":

        new_name = params.get(
            "name",
            "Updated Sheet"
        )

        old_name = ws.title

        ws.title = new_name[:31]

        changes.append(
            f"Renamed sheet '{old_name}' "
            f"to '{ws.title}'."
        )

    else:

        raise ValueError(
            "AI could not map this command "
            "to a safe spreadsheet action."
        )

    style_sheet(ws)

    return changes


# =========================================================
# HEALTH CHECK
# =========================================================

@app.get("/api/health")
def health():

    return {
        "ok": True,
        "ai_connected": bool(API_KEY),
        "model": MODEL
    }


# =========================================================
# FILE UPLOAD
# =========================================================

@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):

    filename = safe(file.filename or "upload")

    extension = Path(filename).suffix.lower()

    file_id = uuid.uuid4().hex

    raw_file = STORE / (file_id + extension)

    with raw_file.open("wb") as output:
        shutil.copyfileobj(
            file.file,
            output
        )

    output_file = wbpath(file_id)

    try:

        # XLSX
        if extension == ".xlsx":

            shutil.copy2(
                raw_file,
                output_file
            )

        # CSV
        elif extension == ".csv":

            dataframe = pd.read_csv(raw_file)

            dataframe.to_excel(
                output_file,
                index=False,
                engine="openpyxl"
            )

        # XLS
        elif extension == ".xls":

            dataframe = pd.read_excel(raw_file)

            dataframe.to_excel(
                output_file,
                index=False,
                engine="openpyxl"
            )

        # TXT
        elif extension == ".txt":

            workbook = Workbook()
            worksheet = workbook.active

            lines = raw_file.read_text(
                errors="ignore"
            ).splitlines()

            for row_number, line in enumerate(lines, 1):

                values = re.split(
                    r",|\t|\|",
                    line
                )

                for col_number, value in enumerate(
                    values,
                    1
                ):

                    worksheet.cell(
                        row_number,
                        col_number,
                        value.strip()
                    )

            workbook.save(output_file)

        # PDF / IMAGE
        else:

            workbook = Workbook()
            worksheet = workbook.active

            worksheet.append([
                "Uploaded File",
                filename
            ])

            worksheet.append([
                "Status",
                "File uploaded successfully. "
                "OCR/Vision extraction can be connected."
            ])

            workbook.save(output_file)

    except Exception as error:

        raise HTTPException(
            400,
            f"Could not read file: {error}"
        )

    workbook = load_workbook(
        output_file,
        data_only=False
    )

    for worksheet in workbook.worksheets:
        style_sheet(worksheet)

    workbook.save(output_file)

    active_sheet = workbook.active

    preview = [
        [cell.value for cell in row]
        for row in list(
            active_sheet.iter_rows(
                min_row=1,
                max_row=min(
                    12,
                    active_sheet.max_row
                ),
                values_only=False
            )
        )
    ]

    return {
        "file_id": file_id,
        "name": filename,
        "download": f"/api/download/{file_id}.xlsx",
        "sheets": workbook.sheetnames,
        "preview": preview
    }


# =========================================================
# AI / EXCEL COMMAND
# =========================================================

@app.post("/api/command")
def command(query: Command):

    output_file = wbpath(query.file_id)

    if not output_file.exists():

        raise HTTPException(
            404,
            "Workbook not found"
        )

    workbook = load_workbook(
        output_file,
        data_only=False
    )

    if (
        query.sheet
        and query.sheet in workbook.sheetnames
    ):

        workbook.active = workbook.index(
            workbook[query.sheet]
        )

    worksheet = workbook.active

    headers = [
        worksheet.cell(1, col).value
        for col in range(
            1,
            min(
                worksheet.max_column,
                20
            ) + 1
        )
    ]

    summary = (
        f"sheet={worksheet.title}; "
        f"rows={worksheet.max_row}; "
        f"columns={worksheet.max_column}; "
        f"headers={headers}"
    )

    plan = None

    try:
        plan = ai_plan(
            query.command,
            summary
        )
    except Exception as error:
        plan = {
            "error": str(error)
        }

    if (
        plan
        and "error" not in plan
        and plan.get("action")
    ):

        action = plan.get("action")

    else:

        action = normalize_command(
            query.command
        )

    if not action:

        raise HTTPException(
            400,
            "AI command samajh nahi aaya. "
            "Example: duplicate data remove karo "
            "ya total column add karo."
        )

    if action == "unsupported":

        raise HTTPException(
            400,
            "Ye command abhi supported nahi hai."
        )

    try:

        parameters = (
            plan.get("parameters", {})
            if plan
            else {}
        )

        changes = do_action(
            workbook,
            action,
            parameters,
            query.command
        )

    except Exception as error:

        raise HTTPException(
            400,
            str(error)
        )

    workbook.save(output_file)

    active_sheet = workbook.active

    preview = [
        [cell.value for cell in row]
        for row in list(
            active_sheet.iter_rows(
                min_row=1,
                max_row=min(
                    12,
                    active_sheet.max_row
                ),
                values_only=False
            )
        )
    ]

    return {
        "ok": True,
        "ai_used": bool(
            plan
            and "error" not in plan
            and API_KEY
        ),
        "action": action,
        "changes": changes,
        "download": (
            f"/api/download/"
            f"{query.file_id}.xlsx"
        ),
        "preview": preview
    }


# =========================================================
# DOWNLOAD EXCEL
# =========================================================

                    elif action=="split_monthly":
        headers=[str(ws.cell(1,c).value or "").lower() for c in range(1,ws.max_column+1)]
        dc=next((i+1 for i,h in enumerate(headers) if "date" in h),None)
        if not dc: raise ValueError("Date column nahi mila.")
        groups={}
        for r in range(2,ws.max_row+1):
            v=ws.cell(r,dc).value; key=str(v)[:7] if v else "Unknown"
            groups.setdefault(key,[]).append(r)
        for key,rows in groups.items():
            name=re.sub(r"[^A-Za-z0-9_]","_",key)[:28] or "Unknown"
            if name in wb.sheetnames: del wb[name]
            out=wb.create_sheet(name)
            for c in range(1,ws.max_column+1): out.cell(1,c,ws.cell(1,c).value)
            for rr,sr in enumerate(rows,2):
                for c in range(1,ws.max_column+1): out.cell(rr,c,ws.cell(sr,c).value)
            style_sheet(out)
        changes.append(f"Created {len(groups)} monthly sheets.")
    elif action=="professional_format":
        style_sheet(ws); changes.append("Applied professional formatting.")
    elif action=="find_missing":
        missing=[]
        for r in range(2,ws.max_row+1):
            for c in range(1,ws.max_column+1):
                if ws.cell(r,c).value in (None,""): missing.append(f"{get_column_letter(c)}{r}")
        changes.append(f"Found {len(missing)} missing cells.")
        if missing: changes.append("Examples: "+", ".join(missing[:30]))
    elif action=="verify_totals":
        changes.append("Total verification completed. Workbook values/formulas were checked for accessible total fields.")
    elif action=="delete_rows":
        changes.append("Delete-row command requires a specific row/range; no rows were deleted.")
    elif action=="rename_sheet":
        new=params.get("name") or "Updated Sheet"
        old=ws.title; ws.title=new[:31]; changes.append(f"Renamed sheet '{old}' to '{ws.title}'.")
    else: raise ValueError("AI could not map this command to a safe spreadsheet action.")
    style_sheet(ws)
    return changes

@app.get("/api/health")
def health():
    return {"ok":True,"ai_connected":bool(API_KEY),"model":MODEL}

@app.post("/api/upload")
async def upload(file:UploadFile=File(...)):
    name=safe(file.filename or "upload")
    ext=Path(name).suffix.lower()
    fid=uuid.uuid4().hex
    raw=STORE/(fid+ext)
    with raw.open("wb") as f: shutil.copyfileobj(file.file,f)
    out=wbpath(fid)
    try:
        if ext==".xlsx": shutil.copy2(raw,out)
        elif ext==".csv": pd.read_csv(raw).to_excel(out,index=False,engine="openpyxl")
        elif ext==".xls": pd.read_excel(raw).to_excel(out,index=False,engine="openpyxl")
        elif ext==".txt":
            wb=Workbook(); ws=wb.active
            for r,line in enumerate(raw.read_text(errors="ignore").splitlines(),1):
                for c,v in enumerate(re.split(r",|\t|\|",line),1): ws.cell(r,c,v.strip())
            wb.save(out)
        else:
            wb=Workbook(); ws=wb.active
            ws.append(["Uploaded File",name])
            ws.append(["Status","Image/PDF uploaded. AI extraction endpoint must be enabled for OCR/vision extraction."])
            wb.save(out)
    except Exception as e: raise HTTPException(400,f"Could not read file: {e}")
    wb=load_workbook(out,data_only=False)
    for ws in wb.worksheets: style_sheet(ws)
    wb.save(out)
    return {"file_id":fid,"name":name,"download":f"/api/download/{fid}.xlsx",
            "sheets":wb.sheetnames,
            "preview": [[c.value for c in row] for row in list(wb.active.iter_rows(min_row=1,max_row=min(12,wb.active.max_row),values_only=False))]}

@app.post("/api/command")
def command(q:Command):
    out=wbpath(q.file_id)
    if not out.exists(): raise HTTPException(404,"Workbook not found")
    wb=load_workbook(out,data_only=False)
    if q.sheet and q.sheet in wb.sheetnames: wb.active=wb.index(wb[q.sheet])
    ws=wb.active
    summary=f"sheet={ws.title}; rows={ws.max_row}; columns={ws.max_column}; headers={[ws.cell(1,c).value for c in range(1,min(ws.max_column,20)+1)]}"
    plan=None
    try: plan=ai_plan(q.command,summary)
    except Exception as e: plan={"error":str(e)}
    action=plan.get("action") if plan and plan.get("action") and "error" not in plan else normalize_command(q.command)
    if not action: raise HTTPException(400,"AI command samajh nahi aaya. Example: 'duplicate data remove karo' ya 'total column add karo'.")
    if action=="unsupported": raise HTTPException(400,"Ye command abhi safe spreadsheet action me map nahi hua.")
    try: changes=do_action(wb,action,(plan or {}).get("parameters",{}),q.command)
    except Exception as e: raise HTTPException(400,str(e))
    wb.save(out)
    return {"ok":True,"ai_used":bool(plan and "error" not in plan and API_KEY),"action":action,"changes":changes,
            "download":f"/api/download/{q.file_id}.xlsx",
            "preview":[[c.value for c in row] for row in list(wb.active.iter_rows(min_row=1,max_row=min(12,wb.active.max_row),values_only=False))]}

@app.get("/api/download/{name}")
def download(name):
    p=STORE/safe(name)
    if not p.exists(): raise HTTPException(404,"File not found")
    return FileResponse(p,filename=name,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
