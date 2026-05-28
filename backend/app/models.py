import uuid
from datetime import datetime
from sqlalchemy import String, Text, Integer, Float, DateTime, ForeignKey, Enum
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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    risk_matrices: Mapped[list["RiskMatrix"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    communication_plans: Mapped[list["CommunicationPlan"]] = relationship(back_populates="project", cascade="all, delete-orphan")
    meeting_plans: Mapped[list["MeetingPlan"]] = relationship(back_populates="project", cascade="all, delete-orphan")


class RiskMatrix(Base):
    __tablename__ = "risk_matrices"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    project_id: Mapped[str] = mapped_column(String(36), ForeignKey("projects.id"))
    title: Mapped[str] = mapped_column(String(200))
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

    matrix: Mapped["RiskMatrix"] = relationship(back_populates="risks")

    @property
    def risk_score(self) -> int:
        return self.probability * self.consequence


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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project: Mapped["Project"] = relationship(back_populates="meeting_plans")
    meetings: Mapped[list["Meeting"]] = relationship(back_populates="plan", cascade="all, delete-orphan")


class Meeting(Base):
    __tablename__ = "meetings"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=new_uuid)
    plan_id: Mapped[str] = mapped_column(String(36), ForeignKey("meeting_plans.id"))
    title: Mapped[str] = mapped_column(String(200))
    date: Mapped[datetime] = mapped_column(DateTime)
    location: Mapped[str | None] = mapped_column(String(300))
    agenda: Mapped[str | None] = mapped_column(Text)
    participants: Mapped[str | None] = mapped_column(Text)
    minutes: Mapped[str | None] = mapped_column(Text)

    plan: Mapped["MeetingPlan"] = relationship(back_populates="meetings")
