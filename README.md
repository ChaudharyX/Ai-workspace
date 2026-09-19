# AI Workspace Practical

This keeps the first index design and makes the AI Sheet workflow runnable.

## Flow
Upload -> Choose Format -> AI Command -> Actual Excel -> Download

## Real workbook operations
- Duplicate removal
- Total column
- Date-wise sorting
- Column merge
- Monthly sheets
- Professional formatting
- Missing-value scan
- Total verification
- Sheet rename
- CSV/XLS/XLSX import

## Start
Windows:
1. Run `run.bat`
2. Set `OPENAI_API_KEY` in the same terminal before starting uvicorn.
3. Open http://127.0.0.1:8000/frontend/index.html if served separately, or serve frontend with a local server.

For easiest local use:
`python -m http.server 5500 --directory frontend`
Then open http://127.0.0.1:5500 and keep the API at port 8000.

Do NOT put the API key inside index.html.
