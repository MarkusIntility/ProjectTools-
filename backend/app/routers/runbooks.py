from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import Runbook, RunbookActivity, Project
from ..schemas import RunbookCreate, RunbookUpdate, RunbookResponse, ActivityCreate, ActivityResponse

router = APIRouter(prefix="/projects/{project_id}/runbooks", tags=["runbooks"])


@router.get("/", response_model=list[RunbookResponse])
def list_runbooks(project_id: str, db: Session = Depends(get_db)):
    return db.query(Runbook).filter(Runbook.project_id == project_id).all()


@router.post("/", response_model=RunbookResponse, status_code=201)
def create_runbook(project_id: str, data: RunbookCreate, db: Session = Depends(get_db)):
    if not db.query(Project).filter(Project.id == project_id).first():
        raise HTTPException(status_code=404, detail="Project not found")
    runbook = Runbook(project_id=project_id, **data.model_dump())
    db.add(runbook)
    db.commit()
    db.refresh(runbook)
    return runbook


@router.get("/{runbook_id}", response_model=RunbookResponse)
def get_runbook(project_id: str, runbook_id: str, db: Session = Depends(get_db)):
    runbook = db.query(Runbook).filter(Runbook.id == runbook_id, Runbook.project_id == project_id).first()
    if not runbook:
        raise HTTPException(status_code=404, detail="Runbook not found")
    return runbook


@router.put("/{runbook_id}", response_model=RunbookResponse)
def update_runbook(project_id: str, runbook_id: str, data: RunbookUpdate, db: Session = Depends(get_db)):
    runbook = db.query(Runbook).filter(Runbook.id == runbook_id, Runbook.project_id == project_id).first()
    if not runbook:
        raise HTTPException(status_code=404, detail="Runbook not found")
    runbook.title = data.title
    runbook.external_url = data.external_url
    db.commit()
    db.refresh(runbook)
    return runbook


@router.delete("/{runbook_id}", status_code=204)
def delete_runbook(project_id: str, runbook_id: str, db: Session = Depends(get_db)):
    runbook = db.query(Runbook).filter(Runbook.id == runbook_id, Runbook.project_id == project_id).first()
    if not runbook:
        raise HTTPException(status_code=404, detail="Runbook not found")
    db.delete(runbook)
    db.commit()


@router.post("/{runbook_id}/activities", response_model=ActivityResponse, status_code=201)
def add_activity(project_id: str, runbook_id: str, data: ActivityCreate, db: Session = Depends(get_db)):
    runbook = db.query(Runbook).filter(Runbook.id == runbook_id, Runbook.project_id == project_id).first()
    if not runbook:
        raise HTTPException(status_code=404, detail="Runbook not found")
    max_order = max((a.sort_order for a in runbook.activities), default=-1)
    activity = RunbookActivity(runbook_id=runbook_id, sort_order=max_order + 1, **{
        k: v for k, v in data.model_dump().items() if k != "sort_order"
    })
    db.add(activity)
    db.commit()
    db.refresh(activity)
    return activity


@router.put("/{runbook_id}/activities/{activity_id}", response_model=ActivityResponse)
def update_activity(project_id: str, runbook_id: str, activity_id: str, data: ActivityCreate, db: Session = Depends(get_db)):
    activity = db.query(RunbookActivity).filter(
        RunbookActivity.id == activity_id, RunbookActivity.runbook_id == runbook_id
    ).first()
    if not activity:
        raise HTTPException(status_code=404, detail="Activity not found")
    for key, value in data.model_dump().items():
        setattr(activity, key, value)
    db.commit()
    db.refresh(activity)
    return activity


@router.delete("/{runbook_id}/activities/{activity_id}", status_code=204)
def delete_activity(project_id: str, runbook_id: str, activity_id: str, db: Session = Depends(get_db)):
    activity = db.query(RunbookActivity).filter(
        RunbookActivity.id == activity_id, RunbookActivity.runbook_id == runbook_id
    ).first()
    if not activity:
        raise HTTPException(status_code=404, detail="Activity not found")
    db.delete(activity)
    db.commit()
