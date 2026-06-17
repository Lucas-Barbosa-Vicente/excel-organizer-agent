from typing import Optional
from datetime import datetime
from pydantic import BaseModel
from app.schemas.organize import OrganizeParameters


class ProfileCreate(BaseModel):
    name: str
    description: Optional[str] = None
    parameters: OrganizeParameters


class ProfileUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    parameters: Optional[OrganizeParameters] = None


class ProfileResponse(BaseModel):
    id: int
    name: str
    description: Optional[str]
    parameters: OrganizeParameters
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True
