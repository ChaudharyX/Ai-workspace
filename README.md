# AI Workspace - Fresh Stable Build

Flow: Upload -> Choose Format -> AI Command -> Actual Excel -> Download

Supported first stable build: XLSX, XLS, CSV, TXT.

Commands include duplicate removal, total column, date sorting, column merge, professional formatting, missing-value scan and total verification.

## Render
Build command: `pip install -r backend/requirements.txt`
Start command: `uvicorn backend.main:app --host 0.0.0.0 --port $PORT`
