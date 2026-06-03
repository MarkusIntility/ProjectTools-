from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import Template
from ..schemas import TemplateCreate, TemplateUpdate, TemplateResponse

router = APIRouter(prefix="/templates", tags=["templates"])


@router.get("/", response_model=list[TemplateResponse])
def list_templates(type: str | None = None, db: Session = Depends(get_db)):
    q = db.query(Template)
    if type:
        q = q.filter(Template.type == type)
    return q.order_by(Template.name).all()


@router.post("/", response_model=TemplateResponse, status_code=201)
def create_template(body: TemplateCreate, db: Session = Depends(get_db)):
    t = Template(name=body.name, type=body.type, data=body.data)
    db.add(t)
    db.commit()
    db.refresh(t)
    return t


@router.get("/{template_id}", response_model=TemplateResponse)
def get_template(template_id: str, db: Session = Depends(get_db)):
    t = db.get(Template, template_id)
    if not t:
        raise HTTPException(404, "Mal ikke funnet")
    return t


@router.put("/{template_id}", response_model=TemplateResponse)
def update_template(template_id: str, body: TemplateUpdate, db: Session = Depends(get_db)):
    t = db.get(Template, template_id)
    if not t:
        raise HTTPException(404, "Mal ikke funnet")
    t.name = body.name
    t.data = body.data
    db.commit()
    db.refresh(t)
    return t


@router.delete("/{template_id}", status_code=204)
def delete_template(template_id: str, db: Session = Depends(get_db)):
    t = db.get(Template, template_id)
    if not t:
        raise HTTPException(404, "Mal ikke funnet")
    db.delete(t)
    db.commit()
