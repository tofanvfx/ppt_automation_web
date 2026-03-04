import os
import sys
import uuid
import shutil
import zipfile
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException, Depends, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from jose import JWTError, jwt

# Add the project directory to path so we can import modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_to_ppt import generate_ppt
import auth

# ── App Setup ────────────────────────────────────────────────────────────────
app = FastAPI(title="PPT Automation")
security = HTTPBearer(auto_error=False)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Path to the template file
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.pptx"

# Temp directory for processing
TEMP_DIR = BASE_DIR / "temp_uploads"
TEMP_DIR.mkdir(exist_ok=True)


# ── Auth Dependencies ────────────────────────────────────────────────────────
def verify_token(credentials: HTTPAuthorizationCredentials = Depends(security)) -> dict:
    """Validate JWT token from Authorization header."""
    if not credentials:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    token = credentials.credentials
    try:
        payload = jwt.decode(token, auth.SECRET_KEY, algorithms=[auth.ALGORITHM])
        return payload
    except JWTError:
        raise HTTPException(status_code=401, detail="Invalid or expired token")

def require_role(role_name: str, allow_admin: bool = True):
    """Dependency factory: ensures the user has a specific role (or admin)."""
    def _check_role(token_payload: dict = Depends(verify_token)):
        user_role = token_payload.get("role")
        if user_role == role_name:
            return token_payload
        if allow_admin and user_role == "admin":
            return token_payload
            
        raise HTTPException(status_code=403, detail=f"Access denied. Required role: {role_name}")
    return _check_role


# ── Auth Endpoints ───────────────────────────────────────────────────────────
@app.post("/auth/login", response_model=auth.Token)
def login(username: str = Form(...), password: str = Form(...), db=Depends(auth.get_db)):
    """Authenticate user and return a JWT token."""
    user = auth.get_user_by_username(db, username)
    if not user or not auth.verify_password(password, user["password_hash"]):
        raise HTTPException(status_code=401, detail="Incorrect username or password")
    
    # Create token payload
    access_token = auth.create_access_token(
        data={"sub": user["username"], "role": user["role"], "id": user["id"]}
    )
    return {"access_token": access_token, "token_type": "bearer"}


@app.get("/auth/me")
def get_current_user(token_payload: dict = Depends(verify_token), db=Depends(auth.get_db)):
    """Return info about the currently authenticated user."""
    user = auth.get_user_by_username(db, token_payload["sub"])
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
        
    return {
        "id": user["id"],
        "username": user["username"],
        "role": user["role"],
        "created_at": user["created_at"]
    }


# ── Core Routes ──────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    """Serve the single-page HTML frontend."""
    html_path = BASE_DIR / "index.html"
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))


@app.post("/upload")
async def upload_and_convert(
    files: list[UploadFile] = File(...),
    token_payload: dict = Depends(require_role("ppt_generator"))
):
    """Accept multiple DOCX uploads, convert to PPTX, and return the file(s). Requires ppt_generator or admin role."""
    # Filter files to only those ending in .docx
    valid_files = [f for f in files if f.filename.lower().endswith(".docx")]
    if not valid_files:
        raise HTTPException(status_code=400, detail="No .docx files were provided.")

    username = token_payload.get("sub", "unknown")
    job_id = str(uuid.uuid4())[:8]
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    try:
        generated_files = []
        for file in valid_files:
            docx_path = job_dir / file.filename
            output_filename = file.filename.rsplit(".", 1)[0] + "_Presentation.pptx"
            output_path = job_dir / output_filename

            content = await file.read()
            docx_path.write_bytes(content)

            print(f"[{username}] Generating PPT from {file.filename}...")
            generate_ppt(
                docx_path=str(docx_path),
                template_path=str(TEMPLATE_PATH),
                output_path=str(output_path),
            )

            if output_path.exists():
                generated_files.append((output_filename, output_path))

        if not generated_files:
            raise HTTPException(status_code=500, detail="Failed to generate any presentations.")

        # If exactly one file was uploaded, return just that PPTX
        if len(generated_files) == 1:
            return FileResponse(
                path=str(generated_files[0][1]),
                filename=generated_files[0][0],
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            
        # If multiple files, zip them together
        zip_filename = f"Generated_Presentations_{job_id}.zip"
        zip_path = job_dir / zip_filename
        
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for out_name, out_path in generated_files:
                zipf.write(out_path, arcname=out_name)
                
        return FileResponse(
            path=str(zip_path),
            filename=zip_filename,
            media_type="application/zip",
        )

    except HTTPException:
        raise
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Error generating PPT: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    print("Starting PPT Automation server...")
    print("Built-in SQLite auth initialized.")
    print("Open http://localhost:8000 in your browser")
    uvicorn.run(app, host="0.0.0.0", port=8000)
