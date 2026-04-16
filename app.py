from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import openpyxl

app = FastAPI()

# Mount the static directory for CSS and JS
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
def read_root():
    with open("static/index.html", "r") as f:
        return f.read()

@app.get("/api/progress")
def get_progress():
    try:
        df = pd.read_excel("data/tasks.xlsx")
        
        # Replace NaN with None so it becomes null in JSON
        df = df.where(pd.notnull(df), None)
        
        # Convert to dictionary records
        records = df.to_dict(orient="records")
        return {"status": "success", "data": records}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
