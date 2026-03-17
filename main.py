from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import aiofiles
import pymorphy3
import openpyxl
from openpyxl.styles import Font
import os
import uuid
from typing import Dict, List
import asyncio
from public_report_export import router

app = FastAPI()

app.include_router(router)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)