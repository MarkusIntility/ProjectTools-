from datetime import datetime
from pydantic import BaseModel
from .models import RiskStatus


# Project
class ProjectBase(BaseModel):
    name: str
    description: str | None = None


class ProjectCreate(ProjectBase):
    pass


class ProjectResponse(ProjectBase):
    id: str
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


# Risk Matrix
class RiskItemBase(BaseModel):
    description: str
    probability: int
    consequence: int
    mitigation: str | None = None
    owner: str | None = None
    status: RiskStatus = RiskStatus.open


class RiskItemCreate(RiskItemBase):
    pass


class RiskItemResponse(RiskItemBase):
    id: str
    matrix_id: str
    risk_score: int

    model_config = {"from_attributes": True}


class RiskMatrixBase(BaseModel):
    title: str


class RiskMatrixCreate(RiskMatrixBase):
    pass


class RiskMatrixResponse(RiskMatrixBase):
    id: str
    project_id: str
    created_at: datetime
    risks: list[RiskItemResponse] = []

    model_config = {"from_attributes": True}


# Communication Plan
class CommunicationEntryBase(BaseModel):
    stakeholder: str
    message: str
    channel: str
    frequency: str
    responsible: str


class CommunicationEntryCreate(CommunicationEntryBase):
    pass


class CommunicationEntryResponse(CommunicationEntryBase):
    id: str
    plan_id: str

    model_config = {"from_attributes": True}


class CommunicationPlanBase(BaseModel):
    title: str


class CommunicationPlanCreate(CommunicationPlanBase):
    pass


class CommunicationPlanResponse(CommunicationPlanBase):
    id: str
    project_id: str
    created_at: datetime
    entries: list[CommunicationEntryResponse] = []

    model_config = {"from_attributes": True}


# Meeting Plan
class MeetingBase(BaseModel):
    title: str
    date: datetime
    location: str | None = None
    agenda: str | None = None
    participants: str | None = None
    minutes: str | None = None


class MeetingCreate(MeetingBase):
    pass


class MeetingResponse(MeetingBase):
    id: str
    plan_id: str

    model_config = {"from_attributes": True}


class MeetingPlanBase(BaseModel):
    title: str


class MeetingPlanCreate(MeetingPlanBase):
    pass


class MeetingPlanResponse(MeetingPlanBase):
    id: str
    project_id: str
    created_at: datetime
    meetings: list[MeetingResponse] = []

    model_config = {"from_attributes": True}


# Runbook
class RunbookCreate(BaseModel):
    title: str
    source: str = "own"
    external_url: str | None = None


class RunbookUpdate(BaseModel):
    title: str
    external_url: str | None = None


class ActivityCreate(BaseModel):
    name: str
    phase: str | None = None
    status: str = "not_started"
    start_date: datetime | None = None
    end_date: datetime | None = None
    responsible: str | None = None
    description: str | None = None
    sort_order: int = 0


class ActivityResponse(ActivityCreate):
    id: str
    runbook_id: str

    model_config = {"from_attributes": True}


class RunbookResponse(BaseModel):
    id: str
    project_id: str
    title: str
    source: str
    external_url: str | None = None
    created_at: datetime
    activities: list[ActivityResponse] = []

    model_config = {"from_attributes": True}
