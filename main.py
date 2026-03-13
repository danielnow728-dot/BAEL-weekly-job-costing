import io
import traceback
import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from job_cost_report import compute_report, write_grouped_excel

app = FastAPI(title="BAEL Job Cost Report API")

# Allow CORS for local development (if frontend is served separately)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/process")
async def process_report(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
         raise HTTPException(status_code=400, detail="Invalid file type. Please upload an Excel (.xlsx or .xls) file.")
    
    try:
        # Read the uploaded Excel file into pandas
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents), sheet_name=0)
        
        # Validate required columns
        required = {"All Jobs", "Employee Name", "Reg", "OT", "Reg.1"}
        missing = required - set(df.columns)
        if missing:
            raise ValueError(f"Missing required columns in input: {sorted(missing)}")
            
        # Run the business logic
        agg, job_avg = compute_report(df)
        
        # Generate the formatted Excel workbook in memory
        output_buffer = write_grouped_excel(agg, job_avg)
        
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

# Mount the static frontend directory
import os
frontend_dir = os.path.join(os.path.dirname(__file__), "frontend")
if os.path.exists(frontend_dir):
    app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")
