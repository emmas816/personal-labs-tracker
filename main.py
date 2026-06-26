import os
import platform
import subprocess
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
from dotenv import load_dotenv

import data_parser
import excel_manager

# Load environment variables
load_dotenv()

app = FastAPI(title="Labs Tracker")

WORKSPACE_DIR = r"c:\Users\admin\Desktop\emma-ai-capstone"

def get_absolute_excel_path(excel_path: str) -> str:
    """
    Resolves the provided path to an absolute path.
    If relative, resolves it relative to the workspace directory.
    """
    excel_path = excel_path.strip()
    if not excel_path:
        raise HTTPException(status_code=400, detail="Excel file path cannot be empty.")
    
    if not os.path.isabs(excel_path):
        return os.path.abspath(os.path.join(WORKSPACE_DIR, excel_path))
    return os.path.abspath(excel_path)

def open_file_locally(file_path: str) -> None:
    """
    Opens the file at file_path using the local system's default launcher.
    """
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found at {abs_path}")
        
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(abs_path)
            # Poll and force the Excel window to the foreground as soon as it initializes (up to 5 seconds)
            import time
            cmd = "((New-Object -ComObject wscript.shell).AppActivate('blood_tests') -or (New-Object -ComObject wscript.shell).AppActivate('Excel'))"
            for _ in range(10):
                time.sleep(0.5)
                res = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
                if "True" in res.stdout:
                    break
        elif system == "Darwin":  # macOS
            subprocess.run(["open", abs_path], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", abs_path], check=True)
    except Exception as e:
        print(f"Failed to open file locally: {e}")
        raise RuntimeError(f"Could not open spreadsheet automatically: {e}")

@app.post("/process-labs")
async def process_labs(
    excel_path: str = Form(...),
    gender: str = Form(...),
    model_name: str = Form("gemini-3.5-flash"),
    file: UploadFile = File(...)
):
    try:
        # 1. Resolve path
        resolved_excel_path = get_absolute_excel_path(excel_path)
        
        # 2. Read uploaded file contents
        content_bytes = await file.read()
        if not content_bytes:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")
        
        # 3. Parse content using Gemini
        try:
            parsed_data = data_parser.parse_blood_test(content_bytes, file.filename, model_name)
        except Exception as parse_err:
            raise HTTPException(
                status_code=500,
                detail=f"Gemini Parser Error: {str(parse_err)}"
            )
            
        # 4. Check if we need to create or append
        is_new_file = not os.path.exists(resolved_excel_path)
        
        try:
            # Append data (this will also create the template first if it doesn't exist)
            excel_manager.append_data_column(resolved_excel_path, parsed_data, gender)
        except Exception as excel_err:
            raise HTTPException(
                status_code=500,
                detail=f"Excel Manager Error: {str(excel_err)}"
            )
            
        # 5. Open spreadsheet locally
        try:
            open_file_locally(resolved_excel_path)
            opened_successfully = True
        except Exception as open_err:
            opened_successfully = False
            # We won't crash the request if opening fails, but we'll report it in the response
            print(f"Warning: {open_err}")
            
        return {
            "status": "success",
            "message": f"Successfully updated Excel sheet!",
            "excel_path": resolved_excel_path,
            "created_new_file": is_new_file,
            "parsed_test_date": parsed_data.get("test_date"),
            "extracted_count": len(parsed_data.get("results", [])),
            "opened_locally": opened_successfully,
            "data": parsed_data
        }
        
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")

@app.post("/open-excel")
async def open_excel(excel_path: str = Form(...)):
    """
    Endpoint to manually trigger opening of an existing Excel file.
    """
    try:
        resolved_excel_path = get_absolute_excel_path(excel_path)
        if not os.path.exists(resolved_excel_path):
            raise HTTPException(status_code=404, detail="Excel file does not exist.")
            
        open_file_locally(resolved_excel_path)
        return {"status": "success", "message": "Spreadsheet opened successfully."}
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(WORKSPACE_DIR, "templates", "index.html")
    if not os.path.exists(index_path):
        return HTMLResponse(
            content="<h3>index.html not found! Please create it in the templates directory.</h3>",
            status_code=404
        )
    with open(index_path, "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
