import os
import io
import traceback
from datetime import datetime
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from job_cost_report import compute_report, write_grouped_excel, update_running_master, check_if_week_exists

app = FastAPI(title="BAEL Job Cost Report API")

# Setup Persistent Save Directory
# Render maps disks to a folder, e.g. /var/data
SAVE_DIR = os.environ.get("RENDER_DISK_PATH", "saved_reports")
os.makedirs(SAVE_DIR, exist_ok=True)

# Allow CORS for local development (if frontend is served separately)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/process")
async def process_file(file: UploadFile = File(...)):
    try:
        # Read the uploaded Excel file
        contents = await file.read()
        xl = pd.ExcelFile(io.BytesIO(contents))
        
        # Use the first sheet name as the week identifier (e.g., the date)
        week_name = str(xl.sheet_names[0]).strip()
        
        # We no longer block duplicate uploads. Let job_cost_report.py overwrite it.
        master_filepath = os.path.join(SAVE_DIR, "Master_Running_Tracker.xlsx")
        # if check_if_week_exists(week_name, master_filepath):
        #     raise HTTPException(status_code=400, detail=f"A report for week '{week_name}' has already been processed and saved.")
            
        # Save the original raw file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_original_filename = file.filename.replace(" ", "_").replace("..", "")
        raw_filename = f"raw_{timestamp}_{safe_original_filename}"
        raw_filepath = os.path.join(SAVE_DIR, raw_filename)
        with open(raw_filepath, "wb") as f:
            f.write(contents)
        
        df = xl.parse(0)
        
        # Validate required columns
        required = {"All Jobs", "Employee Name", "Reg", "OT", "Reg.1"}
        missing = required - set(df.columns)
        if missing:
            raise ValueError(f"Missing required columns in input: {sorted(missing)}")
            
        # Run the core business logic
        agg, job_avg, job_expenses = compute_report(df)
        
        # Generate the formatted weekly Excel workbook in memory
        output_buffer = write_grouped_excel(agg, job_avg, job_expenses)
        
        # Save a persistent copy locally or to the Render Disk
        safe_filename = f"{week_name.replace(' ', '_')}_{file.filename.replace(' ', '_')}"
        saved_filename = f"processed_{timestamp}_{safe_filename}"
        saved_filepath = os.path.join(SAVE_DIR, saved_filename)
        
        with open(saved_filepath, "wb") as f:
            f.write(output_buffer.getvalue())
            
        # Update the Running Master Table
        update_running_master(agg, job_expenses, week_name, master_filepath)
        
        # Return the file as a StreamingResponse so the browser downloads it
        headers = {
            'Content-Disposition': f'attachment; filename="Processed_{file.filename}"'
        }
        return StreamingResponse(
            iter([output_buffer.getvalue()]), 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
        
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"An error occurred while processing the file: {str(e)}")

@app.get("/api/history")
def get_history():
    files = []
    if os.path.exists(SAVE_DIR):
        for f in os.listdir(SAVE_DIR):
            if f.endswith('.xlsx') or f.endswith('.xls'):
                if f == "Master_Running_Tracker.xlsx":
                    continue
                    
                filepath = os.path.join(SAVE_DIR, f)
                stat = os.stat(filepath)
                
                file_type = "processed"
                if f.startswith("raw_"):
                    file_type = "raw"
                elif f.startswith("processed_"):
                    file_type = "processed"
                    
                files.append({
                    "name": f,
                    "date": datetime.fromtimestamp(stat.st_mtime).isoformat() + "Z",
                    "size": stat.st_size,
                    "type": file_type
                })
    # Sort by date descending
    files.sort(key=lambda x: x["date"], reverse=True)
    return {"files": files}

@app.get("/api/history/{filename}")
def download_history_file(filename: str):
    # Security check to prevent path traversal
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Invalid filename")
        
    filepath = os.path.join(SAVE_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found")
        
    return FileResponse(
        filepath, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )

@app.delete("/api/history/{filename}")
def delete_history_file(filename: str):
    # Security check to prevent path traversal
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Invalid filename")
        
    filepath = os.path.join(SAVE_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found")
        
    try:
        os.remove(filepath)
        return {"status": "success", "message": f"{filename} deleted"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to delete file: {str(e)}")

# Mount the static frontend directory
import os
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)

# Try to find exactly where the user uploaded index.html
possible_dirs = [
    os.path.join(parent_dir, "frontend"),
    os.path.join(current_dir, "frontend"),
    current_dir, # If uploaded directly alongside main.py
    parent_dir   # If main.py is in backend/ but index.html is in root/
]

for d in possible_dirs:
    if os.path.exists(os.path.join(d, "index.html")):
        app.mount("/", StaticFiles(directory=d, html=True), name="frontend")
        break
