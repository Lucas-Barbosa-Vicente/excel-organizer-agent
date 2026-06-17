from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.database import get_db
from app.models.profile import Profile
from app.schemas.profile import ProfileCreate, ProfileUpdate, ProfileResponse

router = APIRouter()


@router.get("/profiles", response_model=list[ProfileResponse])
def list_profiles(db: Session = Depends(get_db)):
    return db.query(Profile).all()


@router.post("/profiles", response_model=ProfileResponse, status_code=201)
def create_profile(body: ProfileCreate, db: Session = Depends(get_db)):
    if db.query(Profile).filter(Profile.name == body.name).first():
        raise HTTPException(status_code=400, detail="Já existe um perfil com esse nome.")
    profile = Profile(
        name=body.name,
        description=body.description,
        parameters=body.parameters.model_dump(),
    )
    db.add(profile)
    db.commit()
    db.refresh(profile)
    return profile


@router.get("/profiles/{profile_id}", response_model=ProfileResponse)
def get_profile(profile_id: int, db: Session = Depends(get_db)):
    profile = db.get(Profile, profile_id)
    if not profile:
        raise HTTPException(status_code=404, detail="Perfil não encontrado.")
    return profile


@router.put("/profiles/{profile_id}", response_model=ProfileResponse)
def update_profile(
    profile_id: int, body: ProfileUpdate, db: Session = Depends(get_db)
):
    profile = db.get(Profile, profile_id)
    if not profile:
        raise HTTPException(status_code=404, detail="Perfil não encontrado.")
    if body.name is not None:
        profile.name = body.name
    if body.description is not None:
        profile.description = body.description
    if body.parameters is not None:
        profile.parameters = body.parameters.model_dump()
    db.commit()
    db.refresh(profile)
    return profile


@router.delete("/profiles/{profile_id}", status_code=204)
def delete_profile(profile_id: int, db: Session = Depends(get_db)):
    profile = db.get(Profile, profile_id)
    if not profile:
        raise HTTPException(status_code=404, detail="Perfil não encontrado.")
    db.delete(profile)
    db.commit()
