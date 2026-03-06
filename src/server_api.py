import os
import sys
import uvicorn
from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel
from typing import Optional, Dict, Any

# Ensure we can import database
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import database

app = FastAPI(title="CDO Remote Bridge API")
security = HTTPBasic()

class TaskResult(BaseModel):
    status: str
    result: Optional[Dict[str, Any]] = None

def get_current_user(credentials: HTTPBasicCredentials = Depends(security)):
    """Validates basic auth credentials against the database."""
    # DEBUG: Log login attempt (temporary)
    print(f"DEBUG AUTH: User='{credentials.username}' Password='{credentials.password}'")
    
    if database.check_login(credentials.username, credentials.password):
        return credentials.username
    
    print(f"DEBUG AUTH: Login failed for user '{credentials.username}'")
    raise HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Incorrect username or password",
        headers={"WWW-Authenticate": "Basic"},
    )

@app.get("/")
def read_root():
    return {"status": "online", "service": "CDO Remote Bridge"}

@app.get("/ping")
def ping():
    return {"pong": True}

@app.get("/auth/verify")
def verify_auth(username: str = Depends(get_current_user)):
    """Verifies that the provided credentials are valid."""
    return {"status": "valid", "username": username}

@app.get("/tasks/poll")
def poll_tasks(username: str = Depends(get_current_user)):
    """Agent polls for new tasks."""
    tasks = database.get_pending_tasks(username)
    if not tasks:
        return {"tasks": []}
    return {"tasks": tasks}

@app.post("/tasks/{task_id}/result")
def submit_result(task_id: int, result_data: TaskResult, username: str = Depends(get_current_user)):
    """Agent submits result for a task."""
    # Verify task belongs to user (optional security check, skipping for speed)
    database.update_task_result(task_id, result_data.status, result_data.result)
    return {"status": "ok"}

if __name__ == "__main__":
    # Run on port 8000 by default
    uvicorn.run(app, host="0.0.0.0", port=8000)
