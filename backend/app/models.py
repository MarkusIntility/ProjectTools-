import uuid
from datetime import datetime
from sqlalchemy import String, Text, Integer, Float, DateTime, ForeignKey, Enum, Boolean
from sqlalchemy.orm import Mapped, mapped_column, relationship
from .database import Base
import enum


def new_uuid() -> str:
    return str(uuid.uuid4())


class Project(Base):
    __tablename__ = "projects"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    name: Mapped[str] = mapped_column(String(200))
    description: Mapped[str | None] = mapped_column(Text)
    project_manager: Mapped[str | None] = mapped_column(String(200))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    risk_matrices: Mapped[list["RiskMatrix"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    communication_plans: Mapped[list["CommunicationPlan"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    meeting_plans: Mapped[list["MeetingPlan"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    runbooks: Mapped[list["Runbook"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    project_plans: Mapped[list["ProjectPlan"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    oppgave_lister: Mapped[list["OppgaveListe"]] = relationship(back_populates="project", cascade="all, delete-orphan")


class RiskMatrix(Base):
    __tablename__ = "risk_matrices"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="risk_matrices")
    risks: Mapped[list["RiskItem"]] = relationship(back_populates="matrix", cascade="all, delete-orphan")


class RiskStatus(str, enum.Enum):
    open = "open"
    mitigated = "mitigated"
    closed = "closed"


class RiskItem(Base):
    __tablename__ = "risk_items"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    matrix_id: Mapped[str] = mapped_column(String(36), ForeignKey("risk_matrices.id"))
    description: Mapped[str] = mapped_column(Text)
    probability: Mapped[int] = mapped_column(Integer)  # 1-5
    consequence: Mapped[int] = mapped_column(Integer)  # 1-5
    mitigation: Mapped[str | None] = mapped_column(Text)
    owner: Mapped[str | None] = mapped_column(String(200))
    status: Mapped[RiskStatus] = mapped_column(String(20), default=RiskStatus.open)
    fagomrade: Mapped[str | None] = mapped_column(String(200))
    risk_owner: Mapped[str | None] = mapped_column(String(200))
    residual_probability: Mapped[int | None] = mapped_column(Integer)
    residual_consequence: Mapped[int | None] = mapped_column(Integer)
    fase: Mapped[str | None] = mapped_column(String(50))

    matrix: Mapped["RiskMatrix"] = relationship(back_populates="risks")

    @property
    def risk_score(self) -> int:
        return self.probability * self.consequence

    @property
    def residual_score(self) -> int | None:
        if self.residual_probability and self.residual_consequence:
            return self.residual_probability * self.residual_consequence
        return None


class CommunicationPlan(Base):
    __tablename__ = "communication_plans"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="communication_plans")
    entries: Mapped[list["CommunicationEntry"]] = relationship(back_populates="plan", cascade="all, delete-orphan")


class CommunicationEntry(Base):
    __tablename__ = "communication_entries"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    plan_id: Mapped[str] = mapped_column(String(36), ForeignKey("communication_plans.id"))
    stakeholder: Mapped[str] = mapped_column(String(200))
    message: Mapped[str] = mapped_column(Text)
    channel: Mapped[str] = mapped_column(String(100))  # Teams, e-post, møte, etc.
    frequency: Mapped[str] = mapped_column(String(100))  # Ukentlig, månedlig, etc.
    responsible: Mapped[str] = mapped_column(String(200))

    plan: Mapped["CommunicationPlan"] = relationship(back_populates="entries")


class MeetingPlan(Base):
    __tablename__ = "meeting_plans"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="meeting_plans")
    meetings: Mapped[list["Meeting"]] = relationship(back_populates="plan", cascade="all, delete-orphan")


class Meeting(Base):
    __tablename__ = "meetings"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    plan_id: Mapped[str] = mapped_column(String(36), ForeignKey("meeting_plans.id"))
    title: Mapped[str] = mapped_column(String(200))
    date: Mapped[datetime] = mapped_column(DateTime)
    purpose: Mapped[str | None] = mapped_column(Text)
    outlook_id: Mapped[str | None] = mapped_column(String(200), unique=False)

    plan: Mapped["MeetingPlan"] = relationship(back_populates="meetings")


class RunbookSource(str, enum.Enum):
    own = "own"
    planner = "planner"
    smartsheet = "smartsheet"


class ActivityStatus(str, enum.Enum):
    not_started = "not_started"
    in_progress = "in_progress"
    done = "done"
    cancelled = "cancelled"


class Runbook(Base):
    __tablename__ = "runbooks"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
    source: Mapped[str] = mapped_column(String(20), default="own")
    external_url: Mapped[str | None] = mapped_column(String(500))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="runbooks")
    activities: Mapped[list["RunbookActivity"]] = relationship(
        back_populates="runbook", cascade="all, delete-orphan", order_by="RunbookActivity.sort_order"
    )


class ProjectPlan(Base):
    __tablename__ = "project_plans"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
    source: Mapped[str] = mapped_column(String(20), default="own")
    external_url: Mapped[str | None] = mapped_column(String(500))
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="project_plans")
    tasks: Mapped[list["ProjectPlanTask"]] = relationship(
        back_populates="plan", cascade="all, delete-orphan", order_by="ProjectPlanTask.sort_order"
    )


class ProjectPlanTask(Base):
    __tablename__ = "project_plan_tasks"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    plan_id: Mapped[str] = mapped_column(String(36), ForeignKey("project_plans.id"))
    name: Mapped[str] = mapped_column(String(300))
    bucket: Mapped[str | None] = mapped_column(String(200))
    percent_complete: Mapped[int] = mapped_column(Integer, default=0)
    start_date: Mapped[datetime | None] = mapped_column(DateTime)
    end_date: Mapped[datetime | None] = mapped_column(DateTime)
    responsible: Mapped[str | None] = mapped_column(String(200))
    description: Mapped[str | None] = mapped_column(Text)
    sort_order: Mapped[int] = mapped_column(Integer, default=0)

    plan: Mapped["ProjectPlan"] = relationship(back_populates="tasks")


class RunbookActivity(Base):
    __tablename__ = "runbook_activities"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    runbook_id: Mapped[str] = mapped_column(String(36), ForeignKey("runbooks.id"))
    name: Mapped[str] = mapped_column(String(300))
    phase: Mapped[str | None] = mapped_column(String(200))
    status: Mapped[str] = mapped_column(String(20), default="not_started")
    start_date: Mapped[datetime | None] = mapped_column(DateTime)
    end_date: Mapped[datetime | None] = mapped_column(DateTime)
    responsible: Mapped[str | None] = mapped_column(String(200))
    description: Mapped[str | None] = mapped_column(Text)
    sort_order: Mapped[int] = mapped_column(Integer, default=0)

    runbook: Mapped["Runbook"] = relationship(back_populates="activities")


class OppgaveListe(Base):
    __tablename__ = "oppgave_lister"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
    source: Mapped[str] = mapped_column(String(20), default="own")
    external_url: Mapped[str | None] = mapped_column(String(500))
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="oppgave_lister")
    oppgaver: Mapped[list["Oppgave"]] = relationship(
        back_populates="liste", cascade="all, delete-orphan", order_by="Oppgave.sort_order"
    )


class Oppgave(Base):
    __tablename__ = "oppgaver"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    liste_id: Mapped[str] = mapped_column(String(36), ForeignKey("oppgave_lister.id"))
    name: Mapped[str] = mapped_column(String(300))
    responsible: Mapped[str | None] = mapped_column(String(200))
    due_date: Mapped[datetime | None] = mapped_column(DateTime)
    status: Mapped[str] = mapped_column(String(20), default="not_started")
    description: Mapped[str | None] = mapped_column(Text)
    sort_order: Mapped[int] = mapped_column(Integer, default=0)

    liste: Mapped["OppgaveListe"] = relationship(back_populates="oppgaver")


class Template(Base):
    __tablename__ = "templates"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    name: Mapped[str] = mapped_column(String(200))
    type: Mapped[str] = mapped_column(String(50))  # risk_matrix | communication_plan | meeting_plan | project_plan | oppgave_liste | runbook
    data: Mapped[str] = mapped_column(Text)  # JSON blob
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
