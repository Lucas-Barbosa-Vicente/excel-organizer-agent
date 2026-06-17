import os
import asyncio
import time
import uuid
import aiofiles

UPLOAD_DIR = "storage/uploads"
OUTPUT_DIR = "storage/outputs"
FILE_TTL_SECONDS = 3600  # 1 hour


async def save_upload(data: bytes, filename: str) -> str:
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    ext = os.path.splitext(filename)[1]
    dest = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}{ext}")
    async with aiofiles.open(dest, "wb") as f:
        await f.write(data)
    return dest


def generate_output_path(ext: str = ".xlsx") -> tuple[str, str]:
    token = str(uuid.uuid4())
    path = os.path.join(OUTPUT_DIR, f"{token}{ext}")
    return token, path


async def cleanup_old_files():
    now = time.time()
    for directory in (UPLOAD_DIR, OUTPUT_DIR):
        if not os.path.exists(directory):
            continue
        for fname in os.listdir(directory):
            fpath = os.path.join(directory, fname)
            if os.path.isfile(fpath) and now - os.path.getmtime(fpath) > FILE_TTL_SECONDS:
                await asyncio.to_thread(os.remove, fpath)
