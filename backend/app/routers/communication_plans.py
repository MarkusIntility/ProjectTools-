from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import CommunicationPlan, CommunicationEntry, Project
from ..schemas import CommunicationPlanCreate, CommunicationPlanResponse, CommunicationEntryCreate, CommunicationEntryResponse

router = APIRouter(prefix="/projects/{project_id}/communication-plans", tags=["communication-plans"])


@router.get("/", response_model=list[CommunicationPlanResponse])
def list_plans(project_id: str, db: Session = Depends(get_db)):
    return db.query(CommunicationPlan).filter(CommunicationPlan.project_id == project_id).all()


@router.post("/", response_model=CommunicationPlanResponse, status_code=201)
def create_plan(project_id: str, data: CommunicationPlanCreate, db: Session = Depends(get_db)):
    if not db.query(Project).filter(Project.id == project_id).first():
        raise HTTPException(status_code=404, detail="Project not found")
    plan = CommunicationPlan(project_id=project_id, **data.model_dump())
    db.add(plan)
    db.commit()
    db.refresh(plan)
    return plan


@router.get("/{plan_id}", response_model=CommunicationPlanResponse)
def get_plan(project_id: str, plan_id: str, db: Session = Depends(get_db)):
    plan = db.query(CommunicationPlan).filter(CommunicationPlan.id == plan_id, CommunicationPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Communication plan not found")
    return plan


@router.put("/{plan_id}", response_model=CommunicationPlanResponse)
def update_plan(project_id: str, plan_id: str, data: CommunicationPlanCreate, db: Session = Depends(get_db)):
    plan = db.query(CommunicationPlan).filter(CommunicationPlan.id == plan_id, CommunicationPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Communication plan not found")
    plan.title = data.title
    db.commit()
    db.refresh(plan)
    return plan


@router.delete("/{plan_id}", status_code=204)
def delete_plan(project_id: str, plan_id: str, db: Session = Depends(get_db)):
    plan = db.query(CommunicationPlan).filter(CommunicationPlan.id == plan_id, CommunicationPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Communication plan not found")
    db.delete(plan)
    db.commit()


@router.post("/{plan_id}/entries", response_model=CommunicationEntryResponse, status_code=201)
def add_entry(project_id: str, plan_id: str, data: CommunicationEntryCreate, db: Session = Depends(get_db)):
    plan = db.query(CommunicationPlan).filter(CommunicationPlan.id == plan_id, CommunicationPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Communication plan not found")
    entry = CommunicationEntry(plan_id=plan_id, **data.model_dump())
    db.add(entry)
    db.commit()
    db.refresh(entry)
    return entry


@router.put("/{plan_id}/entries/{entry_id}", response_model=CommunicationEntryResponse)
def update_entry(project_id: str, plan_id: str, entry_id: str, data: CommunicationEntryCreate, db: Session = Depends(get_db)):
    entry = db.query(CommunicationEntry).filter(CommunicationEntry.id == entry_id, CommunicationEntry.plan_id == plan_id).first()
    if not entry:
        raise HTTPException(status_code=404, detail="Entry not found")
    for key, value in data.model_dump().items():
        setattr(entry, key, value)
    db.commit()
    db.refresh(entry)
    return entry


@router.delete("/{plan_id}/entries/{entry_id}", status_code=204)
def delete_entry(project_id: str, plan_id: str, entry_id: str, db: Session = Depends(get_db)):
    entry = db.query(CommunicationEntry).filter(CommunicationEntry.id == entry_id, CommunicationEntry.plan_id == plan_id).first()
    if not entry:
        raise HTTPException(status_code=404, detail="Entry not found")
    db.delete(entry)
    db.commit()
