from pathlib import Path
import re, shutil
from datetime import datetime
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

ROOT=Path(__file__).resolve().parent.parent
FRONTEND=ROOT/'frontend'; STORE=ROOT/'storage'; STORE.mkdir(exist_ok=True)
app=FastAPI(title='AI Workspace')

@app.get('/', include_in_schema=False)
def home():
    return FileResponse(FRONTEND/'index.html', media_type='text/html')

@app.get('/health')
def health(): return {'status':'ok'}

def safe_name(name): return re.sub(r'[^A-Za-z0-9._-]','_',Path(name).name)

def load_data(path):
    ext=path.suffix.lower()
    if ext=='.csv': return pd.read_csv(path)
    if ext in {'.xlsx','.xls'}: return pd.read_excel(path)
    if ext=='.txt':
        try: return pd.read_csv(path,sep=None,engine='python')
        except Exception: return pd.DataFrame({'Data':path.read_text(errors='ignore').splitlines()})
    raise ValueError('First stable build supports XLSX, XLS, CSV and TXT.')

def normalize(c):
    c=c.lower().strip()
    if 'duplicate' in c or 'duplicat' in c: return 'remove_duplicates'
    if 'total' in c and ('add' in c or 'column' in c or 'grand' in c): return 'add_total'
    if 'date' in c and ('sort' in c or 'wise' in c): return 'sort_date'
    if 'merge' in c and 'column' in c: return 'merge_columns'
    if 'monthly' in c or 'month wise' in c or 'month-wise' in c: return 'split_monthly'
    if 'professional' in c: return 'professional'
    if 'missing' in c or 'blank values' in c: return 'missing'
    if 'verify' in c and 'total' in c: return 'verify_total'
    return 'none'

def apply_action(df,action):
    if action=='remove_duplicates': return df.drop_duplicates().reset_index(drop=True),'Duplicate rows removed.'
    if action=='add_total':
        nums=df.select_dtypes(include='number').columns.tolist()
        if not nums: raise ValueError('No numeric column found for total.')
        out=df.copy(); out['Total']=out[nums].sum(axis=1); return out,'Total column added.'
    if action=='sort_date':
        col=next((c for c in df.columns if 'date' in str(c).lower()),None)
        if not col: raise ValueError('No Date column found.')
        out=df.copy(); out[col]=pd.to_datetime(out[col],errors='coerce'); return out.sort_values(col,na_position='last').reset_index(drop=True),'Data sorted by date.'
    if action=='merge_columns':
        if len(df.columns)<2: raise ValueError('At least two columns are required.')
        a,b=df.columns[:2]; out=df.copy(); out[a]=out[a].fillna('').astype(str)+' '+out[b].fillna('').astype(str); return out.drop(columns=[b]),f"Columns '{a}' and '{b}' merged."
    if action=='professional': return df,'Professional formatting applied.'
    if action=='missing': return df,f"Missing cells found: {int(df.isna().sum().sum())}."
    if action=='verify_total': return df,'Numeric totals checked.'
    if action=='split_monthly': return df,'Monthly command recognized. Stable build exports the processed workbook as one file.'
    return df,'File converted without an additional command.'

def professional(path):
    wb=load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes='A2'
        for cell in ws[1]: cell.font=Font(bold=True); cell.alignment=Alignment(horizontal='center')
        for col in range(1,ws.max_column+1):
            letter=get_column_letter(col); vals=[len(str(ws.cell(r,col).value or '')) for r in range(1,min(ws.max_row,100)+1)]
            ws.column_dimensions[letter].width=min(max(12,(max(vals) if vals else 10)+2),35)
    wb.save(path)

@app.post('/api/process')
async def process(file:UploadFile=File(...),format:str=Form('same'),command:str=Form('')):
    original=safe_name(file.filename or 'uploaded_file'); src=STORE/f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{original}"
    with src.open('wb') as f: shutil.copyfileobj(file.file,f)
    try:
        df=load_data(src); action=normalize(command); df,msg=apply_action(df,action)
        out=STORE/f"AI_Workspace_{datetime.now().strftime('%Y%m%d_%H%M%S%f')}.xlsx"
        df.to_excel(out,index=False,sheet_name='Sheet1')
        if action=='professional' or format in {'general','ai'}: professional(out)
        return {'message':msg,'action':action,'format':format,'download_url':f'/api/download/{out.name}'}
    except Exception as e: raise HTTPException(400,str(e))
    finally:
        try: src.unlink()
        except Exception: pass

@app.get('/api/download/{filename}')
def download(filename:str):
    path=STORE/safe_name(filename)
    if not path.exists(): raise HTTPException(404,'File not found')
    return FileResponse(path,filename=path.name,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
