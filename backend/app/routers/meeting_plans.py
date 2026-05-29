from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import MeetingPlan, Meeting, Project
from ..schemas import MeetingPlanCreate, MeetingPlanResponse, MeetingCreate, MeetingResponse

router = APIRouter(prefix="/projects/{project_id}/meeting-plans", tags=["meeting-plans"])


@router.get("/", response_model=list[MeetingPlanResponse])
def list_plans(project_id: str, db: Session = Depends(get_db)):
    return db.query(MeetingPlan).filter(MeetingPlan.project_id == project_id).all()


@router.post("/", response_model=MeetingPlanResponse, status_code=201)
def create_plan(project_id: str, data: MeetingPlanCreate, db: Session = Depends(get_db)):
    if not db.query(Project).filter(Project.id == project_id).first():
        raise HTTPException(status_code=404, detail="Project not found")
    plan = MeetingPlan(project_id=project_id, **data.model_dump())
    db.add(plan)
    db.commit()
    db.refresh(plan)
    return plan


@router.get("/{plan_id}", response_model=MeetingPlanResponse)
def get_plan(project_id: str, plan_id: str, db: Session = Depends(get_db)):
    plan = db.query(MeetingPlan).filter(MeetingPlan.id == plan_id, MeetingPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Meeting plan not found")
    return plan


@router.put("/{plan_id}", response_model=MeetingPlanResponse)
def update_plan(project_id: str, plan_id: str, data: MeetingPlanCreate, db: Session = Depends(get_db)):
    plan = db.query(MeetingPlan).filter(MeetingPlan.id == plan_id, MeetingPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Meeting plan not found")
    plan.title = data.title
    db.commit()
    db.refresh(plan)
    return plan


@router.delete("/{plan_id}", status_code=204)
def delete_plan(project_id: str, plan_id: str, db: Session = Depends(get_db)):
    plan = db.query(MeetingPlan).filter(MeetingPlan.id == plan_id, MeetingPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Meeting plan not found")
    db.delete(plan)
    db.commit()


@router.post("/{plan_id}/meetings", response_model=MeetingResponse, status_code=201)
def add_meeting(project_id: str, plan_id: str, data: MeetingCreate, db: Session = Depends(get_db)):
    plan = db.query(MeetingPlan).filter(MeetingPlan.id == plan_id, MeetingPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Meeting plan not found")
    meeting = Meeting(plan_id=plan_id, **data.model_dump())
    db.add(meeting)
    db.commit()
    db.refresh(meeting)
    return meeting


@router.put("/{plan_id}/meetings/{meeting_id}", response_model=MeetingResponse)
def update_meeting(project_id: str, plan_id: str, meeting_id: str, data: MeetingCreate, db: Session = Depends(get_db)):
    meeting = db.query(Meeting).filter(Meeting.id == meeting_id, Meeting.plan_id == plan_id).first()
    if not meeting:
        raise HTTPException(status_code=404, detail="Meeting not found")
    for key, value in data.model_dump().items():
        setattr(meeting, key, value)
    db.commit()
    db.refresh(meeting)
    return meeting


@router.delete("/{plan_id}/meetings/{meeting_id}", status_code=204)
def delete_meeting(project_id: str, plan_id: str, meeting_id: str, db: Session = Depends(get_db)):
    meeting = db.query(Meeting).filter(Meeting.id == meeting_id, Meeting.plan_id == plan_id).first()
    if not meeting:
        raise HTTPException(status_code=404, detail="Meeting not found")
    db.delete(meeting)
    db.commit()
