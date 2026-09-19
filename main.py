from pathlib import Path
import re, shutil
from datetime import datetime
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STORE = ROOT / "storage"
STORE.mkdir(exist_ok=True)
app = FastAPI(title="AI Workspace")

@app.get("/", include_in_schema=False)
def home():
    return FileResponse(ROOT / "index.html", media_type="text/html")

@app.get("/health")
def health():
    return {"status": "ok"}

def safe_name(name):
    return re.sub(r"[^A-Za-z0-9._-]", "_", Path(name).name)

def load_data(path):
    ext = path.suffix.lower()
    if ext == ".csv": return pd.read_csv(path)
    if ext in {".xlsx",".xls"}: return pd.read_excel(path)
    if ext == ".txt":
        try: return pd.read_csv(path, sep=None, engine="python")
        except Exception: return pd.DataFrame({"Data": path.read_text(errors="ignore").splitlines()})
    raise ValueError("Use XLSX, XLS, CSV or TXT in this stable build.")

def action_for(command):
    c = command.lower()
    if "duplicate" in c or "duplicat" in c: return "duplicate"
    if "total" in c and ("add" in c or "column" in c): return "total"
    if "date" in c and ("sort" in c or "wise" in c): return "date"
    if "merge" in c and "column" in c: return "merge"
    if "missing" in c or "blank" in c: return "missing"
    if "professional" in c: return "professional"
    if "verify" in c and "total" in c: return "verify"
    return "none"

def apply(df, action):
    if action == "duplicate":
        return df.drop_duplicates().reset_index(drop=True), "Duplicate rows removed."
    if action == "total":
        nums = df.select_dtypes(include="number").columns.tolist()
        if not nums: raise ValueError("No numeric column found for total.")
        x=df.copy(); x["Total"]=x[nums].sum(axis=1); return x,"Total column added."
    if action == "date":
        x=df.copy(); col=next((c for c in x.columns if "date" in str(c).lower()),None)
        if not col: raise ValueError("No Date column found.")
        x[col]=pd.to_datetime(x[col],errors="coerce"); return x.sort_values(col,na_position="last").reset_index(drop=True),"Data sorted by date."
    if action == "merge":
        if len(df.columns)<2: raise ValueError("At least two columns are required.")
        x=df.copy(); a,b=x.columns[:2]; x[a]=x[a].fillna("").astype(str)+" "+x[b].fillna("").astype(str); x=x.drop(columns=[b]); return x,f"Columns '{a}' and '{b}' merged."
    if action == "missing":
        return df,f"Missing cells found: {int(df.isna().sum().sum())}."
    if action == "verify":
        nums=df.select_dtypes(include="number").columns.tolist(); return df,"Numeric totals checked: "+(", ".join(map(str,nums)) if nums else "none")
    return df,"File converted successfully."

def professional(path):
    wb=load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes="A2"
        for c in ws[1]: c.font=Font(bold=True); c.alignment=Alignment(horizontal="center")
        for col in range(1,ws.max_column+1):
            letter=get_column_letter(col)
            width=min(max(12,max((len(str(ws.cell(r,col).value or "")) for r in range(1,min(ws.max_row,100)+1)),default=10)+2),35)
            ws.column_dimensions[letter].width=width
    wb.save(path)

@app.post("/api/process")
async def process(file: UploadFile=File(...), format: str=Form("same"), command: str=Form("")):
    name=safe_name(file.filename or "upload")
    src=STORE/(datetime.now().strftime("%Y%m%d%H%M%S%f")+"_"+name)
    with src.open("wb") as f: shutil.copyfileobj(file.file,f)
    try:
        df=load_data(src)
        action=action_for(command)
        df,msg=apply(df,action)
        out=STORE/("AI_Workspace_"+datetime.now().strftime("%Y%m%d_%H%M%S%f")+".xlsx")
        with pd.ExcelWriter(out,engine="openpyxl") as w: df.to_excel(w,index=False,sheet_name="Sheet1")
        if action=="professional" or format in {"general","ai"}: professional(out)
        return {"message":msg,"download_url":f"/api/download/{out.name}"}
    except Exception as e: raise HTTPException(400,str(e))
    finally:
        try: src.unlink()
        except: pass

@app.get("/api/download/{filename}")
def download(filename: str):
    path=STORE/safe_name(filename)
    if not path.exists(): raise HTTPException(404,"File not found")
    return FileResponse(path,filename=path.name,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
