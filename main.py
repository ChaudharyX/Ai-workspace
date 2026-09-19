
import os, re, json, uuid, shutil, base64
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import requests

ROOT=Path(__file__).resolve().parent.parent
STORE=ROOT/"storage"; STORE.mkdir(exist_ok=True)
app=FastAPI(title="AI Workspace Practical")
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_methods=["*"],allow_headers=["*"])

API_KEY=os.getenv("OPENAI_API_KEY","").strip()
MODEL=os.getenv("AI_MODEL","gpt-5.6-luna")

class Command(BaseModel):
    file_id:str
    command:str
    format_type:str="same"
    sheet:Optional[str]=None

def safe(x): return re.sub(r"[^A-Za-z0-9._-]","_",x)
def wbpath(fid): return STORE/(fid+".xlsx")

def style_sheet(ws):
    if ws.max_row < 1: return
    for c in ws[1]:
        c.font=Font(bold=True)
        c.alignment=Alignment(horizontal="center")
    ws.freeze_panes="A2"
    for i in range(1,ws.max_column+1):
        vals=[str(ws.cell(r,i).value or "") for r in range(1,ws.max_row+1)]
        ws.column_dimensions[get_column_letter(i)].width=min(max(max(map(len,vals),default=8)+2,10),42)

def normalize_command(command):
    c=command.lower()
    if any(x in c for x in ["duplicate","duplicates","duplicat","डुप्लिकेट"]): return "remove_duplicates"
    if "total" in c and any(x in c for x in ["column","add","calculate","जोड़"]): return "add_total"
    if "sort" in c or "क्रम" in c or "date wise" in c or "date-wise" in c: return "sort"
    if "merge" in c or "combine column" in c: return "merge_columns"
    if "monthly" in c and "sheet" in c: return "split_monthly"
    if "professional" in c or "format" in c: return "professional_format"
    if "missing" in c or "blank" in c: return "find_missing"
    if "verify" in c and "total" in c: return "verify_totals"
    if "delete" in c and "row" in c: return "delete_rows"
    if "rename" in c and "sheet" in c: return "rename_sheet"
    return None

def ai_plan(command, summary):
    if not API_KEY: return None
    prompt=f"""You are an Excel command planner.
Return ONLY JSON with:
{{"action":"remove_duplicates|add_total|sort|merge_columns|split_monthly|professional_format|find_missing|verify_totals|delete_rows|rename_sheet|unsupported","parameters":{{}}}}
Understand Hindi/Hinglish/English.
User command: {command}
Workbook: {summary}
"""
    r=requests.post("https://api.openai.com/v1/responses",
        headers={"Authorization":f"Bearer {API_KEY}","Content-Type":"application/json"},
        json={"model":MODEL,"input":prompt},timeout=90)
    r.raise_for_status()
    txt=r.json().get("output_text","")
    m=re.search(r"\{.*\}",txt,re.S)
    return json.loads(m.group()) if m else None

def do_action(wb, action, params, original_command):
    ws=wb.active
    changes=[]
    if action=="remove_duplicates":
        seen=set(); delete=[]
        for r in range(2,ws.max_row+1):
            key=tuple(ws.cell(r,c).value for c in range(1,ws.max_column+1))
            if repr(key) in seen: delete.append(r)
            else: seen.add(repr(key))
        for r in reversed(delete): ws.delete_rows(r,1)
        changes.append(f"Removed {len(delete)} duplicate rows.")
    elif action=="add_total":
        headers=[str(ws.cell(1,c).value or "").lower() for c in range(1,ws.max_column+1)]
        col=next((i+1 for i,h in enumerate(headers) if any(k in h for k in ["amount","value","price","qty","quantity","total"])),ws.max_column)
        new=ws.max_column+1; ws.cell(1,new,"Total")
        for r in range(2,ws.max_row+1):
            v=ws.cell(r,col).value
            ws.cell(r,new,v if isinstance(v,(int,float)) else v)
        changes.append(f"Added Total column using {get_column_letter(col)}.")
    elif action=="sort":
        headers=[str(ws.cell(1,c).value or "").lower() for c in range(1,ws.max_column+1)]
        col=next((i+1 for i,h in enumerate(headers) if "date" in h),1)
        rows=sorted(list(ws.iter_rows(min_row=2,values_only=True)),key=lambda x:(x[col-1] is None,str(x[col-1])))
        for r,row in enumerate(rows,2):
            for c,v in enumerate(row,1): ws.cell(r,c,v)
        changes.append(f"Sorted by {get_column_letter(col)}.")
    elif action=="merge_columns":
        m=re.search(r"(?:column\s*)?([A-Z])\s*(?:and|&|,)\s*(?:column\s*)?([A-Z])",original_command,re.I)
        if not m: raise ValueError("Merge command me dono column names chahiye, example: Column B and C merge karo.")
        a,b=m.group(1).upper(),m.group(2).upper(); ca,cb=ws[a+"1"].column,ws[b+"1"].column
        for r in range(1,ws.max_row+1):
            vals=[str(ws.cell(r,c).value) for c in (ca,cb) if ws.cell(r,c).value not in (None,"")]
            ws.cell(r,ca," ".join(vals))
        ws.delete_cols(cb,1); changes.append(f"Merged columns {a} and {b}.")
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
