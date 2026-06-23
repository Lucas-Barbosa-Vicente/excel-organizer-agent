from typing import List, Optional
from pydantic import BaseModel


class RenameResponse(BaseModel):
    status: str
    download_token: Optional[str] = None
    matched: int
    unmatched: int
    unmatched_files: List[str]
    total: int
