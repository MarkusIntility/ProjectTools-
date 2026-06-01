from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from ..database import get_db
from ..models import OppgaveListe, Oppgave, Project
from ..schemas import OppgaveListeCreate, OppgaveListeUpdate, OppgaveListeResponse, OppgaveCreate, OppgaveResponse

router = APIRouter(prefix="/projects/{project_id}/oppgave-lister", tags=["oppgaver"])


@router.get("/", response_model=list[OppgaveListeResponse])
def list_oppgave_lister(project_id: str, db: Session = Depends(get_db)):
    return db.query(OppgaveListe).filter(OppgaveListe.project_id == project_id).all()


@router.post("/", response_model=OppgaveListeResponse, status_code=201)
def create_oppgave_liste(project_id: str, data: OppgaveListeCreate, db: Session = Depends(get_db)):
    if not db.query(Project).filter(Project.id == project_id).first():
        raise HTTPException(status_code=404, detail="Project not found")
    liste = OppgaveListe(project_id=project_id, **data.model_dump())
    db.add(liste)
    db.commit()
    db.refresh(liste)
    return liste


@router.get("/{liste_id}", response_model=OppgaveListeResponse)
def get_oppgave_liste(project_id: str, liste_id: str, db: Session = Depends(get_db)):
    liste = db.query(OppgaveListe).filter(OppgaveListe.id == liste_id, OppgaveListe.project_id == project_id).first()
    if not liste:
        raise HTTPException(status_code=404, detail="Oppgaveliste not found")
    return liste


@router.put("/{liste_id}", response_model=OppgaveListeResponse)
def update_oppgave_liste(project_id: str, liste_id: str, data: OppgaveListeUpdate, db: Session = Depends(get_db)):
    liste = db.query(OppgaveListe).filter(OppgaveListe.id == liste_id, OppgaveListe.project_id == project_id).first()
    if not liste:
        raise HTTPException(status_code=404, detail="Oppgaveliste not found")
    liste.title = data.title
    liste.external_url = data.external_url
    db.commit()
    db.refresh(liste)
    return liste


@router.delete("/{liste_id}", status_code=204)
def delete_oppgave_liste(project_id: str, liste_id: str, db: Session = Depends(get_db)):
    liste = db.query(OppgaveListe).filter(OppgaveListe.id == liste_id, OppgaveListe.project_id == project_id).first()
    if not liste:
        raise HTTPException(status_code=404, detail="Oppgaveliste not found")
    db.delete(liste)
    db.commit()


@router.post("/{liste_id}/oppgaver", response_model=OppgaveResponse, status_code=201)
def add_oppgave(project_id: str, liste_id: str, data: OppgaveCreate, db: Session = Depends(get_db)):
    liste = db.query(OppgaveListe).filter(OppgaveListe.id == liste_id, OppgaveListe.project_id == project_id).first()
    if not liste:
        raise HTTPException(status_code=404, detail="Oppgaveliste not found")
    max_order = max((o.sort_order for o in liste.oppgaver), default=-1)
    oppgave = Oppgave(liste_id=liste_id, sort_order=max_order + 1, **{
        k: v for k, v in data.model_dump().items() if k != "sort_order"
    })
    db.add(oppgave)
    db.commit()
    db.refresh(oppgave)
    return oppgave


@router.put("/{liste_id}/oppgaver/{oppgave_id}", response_model=OppgaveResponse)
def update_oppgave(project_id: str, liste_id: str, oppgave_id: str, data: OppgaveCreate, db: Session = Depends(get_db)):
    oppgave = db.query(Oppgave).filter(Oppgave.id == oppgave_id, Oppgave.liste_id == liste_id).first()
    if not oppgave:
        raise HTTPException(status_code=404, detail="Oppgave not found")
    for key, value in data.model_dump().items():
        setattr(oppgave, key, value)
    db.commit()
    db.refresh(oppgave)
    return oppgave


@router.delete("/{liste_id}/oppgaver/{oppgave_id}", status_code=204)
def delete_oppgave(project_id: str, liste_id: str, oppgave_id: str, db: Session = Depends(get_db)):
    oppgave = db.query(Oppgave).filter(Oppgave.id == oppgave_id, Oppgave.liste_id == liste_id).first()
    if not oppgave:
        raise HTTPException(status_code=404, detail="Oppgave not found")
    db.delete(oppgave)
    db.commit()
