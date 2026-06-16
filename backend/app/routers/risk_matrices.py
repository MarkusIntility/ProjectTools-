from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import RiskMatrix, RiskItem, Project
from ..schemas import RiskMatrixCreate, RiskMatrixResponse, RiskItemCreate, RiskItemResponse

router = APIRouter(prefix="/projects/{project_id}/risk-matrices", tags=["risk-matrices"])


@router.get("/", response_model=list[RiskMatrixResponse])
def list_matrices(project_id: str, db: Session = Depends(get_db)):
    return db.query(RiskMatrix).filter(RiskMatrix.project_id == project_id).all()


@router.post("/", response_model=RiskMatrixResponse, status_code=201)
def create_matrix(project_id: str, data: RiskMatrixCreate, db: Session = Depends(get_db)):
    if not db.query(Project).filter(Project.id == project_id).first():
        raise HTTPException(status_code=404, detail="Project not found")
    matrix = RiskMatrix(project_id=project_id, **data.model_dump())
    db.add(matrix)
    db.commit()
    db.refresh(matrix)
    return matrix


@router.get("/{matrix_id}", response_model=RiskMatrixResponse)
def get_matrix(project_id: str, matrix_id: str, db: Session = Depends(get_db)):
    matrix = db.query(RiskMatrix).filter(RiskMatrix.id == matrix_id, RiskMatrix.project_id == project_id).first()
    if not matrix:
        raise HTTPException(status_code=404, detail="Risk matrix not found")
    return matrix


@router.put("/{matrix_id}", response_model=RiskMatrixResponse)
def update_matrix(project_id: str, matrix_id: str, data: RiskMatrixCreate, db: Session = Depends(get_db)):
    matrix = db.query(RiskMatrix).filter(RiskMatrix.id == matrix_id, RiskMatrix.project_id == project_id).first()
    if not matrix:
        raise HTTPException(status_code=404, detail="Risk matrix not found")
    matrix.title = data.title
    db.commit()
    db.refresh(matrix)
    return matrix


@router.delete("/{matrix_id}", status_code=204)
def delete_matrix(project_id: str, matrix_id: str, db: Session = Depends(get_db)):
    matrix = db.query(RiskMatrix).filter(RiskMatrix.id == matrix_id, RiskMatrix.project_id == project_id).first()
    if not matrix:
        raise HTTPException(status_code=404, detail="Risk matrix not found")
    db.delete(matrix)
    db.commit()


@router.post("/{matrix_id}/risks", response_model=RiskItemResponse, status_code=201)
def add_risk(project_id: str, matrix_id: str, data: RiskItemCreate, db: Session = Depends(get_db)):
    matrix = db.query(RiskMatrix).filter(RiskMatrix.id == matrix_id, RiskMatrix.project_id == project_id).first()
    if not matrix:
        raise HTTPException(status_code=404, detail="Risk matrix not found")
    risk = RiskItem(matrix_id=matrix_id, **data.model_dump())
    db.add(risk)
    db.commit()
    db.refresh(risk)
    return risk


@router.put("/{matrix_id}/risks/{risk_id}", response_model=RiskItemResponse)
def update_risk(project_id: str, matrix_id: str, risk_id: str, data: RiskItemCreate, db: Session = Depends(get_db)):
    risk = db.query(RiskItem).filter(RiskItem.id == risk_id, RiskItem.matrix_id == matrix_id).first()
    if not risk:
        raise HTTPException(status_code=404, detail="Risk not found")
    for key, value in data.model_dump().items():
        setattr(risk, key, value)
    db.commit()
    db.refresh(risk)
    return risk


@router.delete("/{matrix_id}/risks/{risk_id}", status_code=204)
def delete_risk(project_id: str, matrix_id: str, risk_id: str, db: Session = Depends(get_db)):
    risk = db.query(RiskItem).filter(RiskItem.id == risk_id, RiskItem.matrix_id == matrix_id).first()
    if not risk:
        raise HTTPException(status_code=404, detail="Risk not found")
    db.delete(risk)
    db.commit()


@router.put("/{matrix_id}/set-primary", response_model=RiskMatrixResponse)
def set_primary(project_id: str, matrix_id: str, db: Session = Depends(get_db)):
    matrix = db.query(RiskMatrix).filter(RiskMatrix.id == matrix_id, RiskMatrix.project_id == project_id).first()
    if not matrix:
        raise HTTPException(status_code=404, detail="Risk matrix not found")
    db.query(RiskMatrix).filter(RiskMatrix.project_id == project_id).update({"is_primary": False})
    matrix.is_primary = True
    db.commit()
    db.refresh(matrix)
    return matrix
