from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import ProjectPlan, ProjectPlanTask, Project
from ..schemas import ProjectPlanCreate, ProjectPlanUpdate, ProjectPlanResponse, ProjectPlanTaskCreate, ProjectPlanTaskResponse

router = APIRouter(prefix="/projects/{project_id}/project-plans", tags=["project-plans"])


@router.get("/", response_model=list[ProjectPlanResponse])
def list_project_plans(project_id: str, db: Session = Depends(get_db)):
    return db.query(ProjectPlan).filter(ProjectPlan.project_id == project_id).all()


@router.post("/", response_model=ProjectPlanResponse, status_code=201)
def create_project_plan(project_id: str, data: ProjectPlanCreate, db: Session = Depends(get_db)):
    if not db.query(Project).filter(Project.id == project_id).first():
        raise HTTPException(status_code=404, detail="Project not found")
    plan = ProjectPlan(project_id=project_id, **data.model_dump())
    db.add(plan)
    db.commit()
    db.refresh(plan)
    return plan


@router.get("/{plan_id}", response_model=ProjectPlanResponse)
def get_project_plan(project_id: str, plan_id: str, db: Session = Depends(get_db)):
    plan = db.query(ProjectPlan).filter(ProjectPlan.id == plan_id, ProjectPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Project plan not found")
    return plan


@router.put("/{plan_id}", response_model=ProjectPlanResponse)
def update_project_plan(project_id: str, plan_id: str, data: ProjectPlanUpdate, db: Session = Depends(get_db)):
    plan = db.query(ProjectPlan).filter(ProjectPlan.id == plan_id, ProjectPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Project plan not found")
    plan.title = data.title
    plan.external_url = data.external_url
    db.commit()
    db.refresh(plan)
    return plan


@router.delete("/{plan_id}", status_code=204)
def delete_project_plan(project_id: str, plan_id: str, db: Session = Depends(get_db)):
    plan = db.query(ProjectPlan).filter(ProjectPlan.id == plan_id, ProjectPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Project plan not found")
    db.delete(plan)
    db.commit()


@router.post("/{plan_id}/tasks", response_model=ProjectPlanTaskResponse, status_code=201)
def add_task(project_id: str, plan_id: str, data: ProjectPlanTaskCreate, db: Session = Depends(get_db)):
    plan = db.query(ProjectPlan).filter(ProjectPlan.id == plan_id, ProjectPlan.project_id == project_id).first()
    if not plan:
        raise HTTPException(status_code=404, detail="Project plan not found")
    max_order = max((t.sort_order for t in plan.tasks), default=-1)
    task = ProjectPlanTask(plan_id=plan_id, sort_order=max_order + 1, **{
        k: v for k, v in data.model_dump().items() if k != "sort_order"
    })
    db.add(task)
    db.commit()
    db.refresh(task)
    return task


@router.put("/{plan_id}/tasks/{task_id}", response_model=ProjectPlanTaskResponse)
def update_task(project_id: str, plan_id: str, task_id: str, data: ProjectPlanTaskCreate, db: Session = Depends(get_db)):
    task = db.query(ProjectPlanTask).filter(
        ProjectPlanTask.id == task_id, ProjectPlanTask.plan_id == plan_id
    ).first()
    if not task:
        raise HTTPException(status_code=404, detail="Task not found")
    for key, value in data.model_dump().items():
        setattr(task, key, value)
    db.commit()
    db.refresh(task)
    return task


@router.delete("/{plan_id}/tasks/{task_id}", status_code=204)
def delete_task(project_id: str, plan_id: str, task_id: str, db: Session = Depends(get_db)):
    task = db.query(ProjectPlanTask).filter(
        ProjectPlanTask.id == task_id, ProjectPlanTask.plan_id == plan_id
    ).first()
    if not task:
        raise HTTPException(status_code=404, detail="Task not found")
    db.delete(task)
    db.commit()
