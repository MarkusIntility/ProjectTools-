from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import inspect, text
from .database import engine, Base
from .routers import projects, risk_matrices, communication_plans, meeting_plans, runbooks, project_plans, oppgaver, templates

Base.metadata.create_all(bind=engine)


def _migrate_risk_items():
    """Add new columns to risk_items if they don't exist (no Alembic — manual migration)."""
    inspector = inspect(engine)
    if "risk_items" not in inspector.get_table_names():
        return
    existing = {col["name"] for col in inspector.get_columns("risk_items")}
    additions = [
        ("fagomrade", "NVARCHAR(200)"),
        ("risk_owner", "NVARCHAR(200)"),
        ("residual_probability", "INT"),
        ("residual_consequence", "INT"),
        ("fase", "NVARCHAR(50)"),
    ]
    with engine.connect() as conn:
        for col, typ in additions:
            if col not in existing:
                conn.execute(text(f"ALTER TABLE risk_items ADD [{col}] {typ} NULL"))
        conn.commit()


_migrate_risk_items()


def _migrate_meetings():
    """Add purpose and outlook_id columns to meetings if they don't exist."""
    inspector = inspect(engine)
    if "meetings" not in inspector.get_table_names():
        return
    existing = {col["name"] for col in inspector.get_columns("meetings")}
    additions = [
        ("purpose", "NVARCHAR(MAX)"),
        ("outlook_id", "NVARCHAR(200)"),
    ]
    with engine.connect() as conn:
        for col, typ in additions:
            if col not in existing:
                conn.execute(text(f"ALTER TABLE meetings ADD [{col}] {typ} NULL"))
        conn.commit()


_migrate_meetings()


def _migrate_is_primary():
    """Add is_primary column to resource tables if not present."""
    inspector = inspect(engine)
    tables = {
        "risk_matrices": "BIT",
        "meeting_plans": "BIT",
        "project_plans": "BIT",
        "oppgave_lister": "BIT",
    }
    with engine.connect() as conn:
        for table, typ in tables.items():
            if table not in inspector.get_table_names():
                continue
            existing = {col["name"] for col in inspector.get_columns(table)}
            if "is_primary" not in existing:
                conn.execute(text(f"ALTER TABLE [{table}] ADD [is_primary] {typ} NOT NULL DEFAULT 0"))
        conn.commit()


_migrate_is_primary()


def _migrate_project_manager():
    """Add project_manager column to projects if not present."""
    inspector = inspect(engine)
    if "projects" not in inspector.get_table_names():
        return
    existing = {col["name"] for col in inspector.get_columns("projects")}
    if "project_manager" not in existing:
        with engine.connect() as conn:
            conn.execute(text("ALTER TABLE [projects] ADD [project_manager] NVARCHAR(200) NULL"))
            conn.commit()


_migrate_project_manager()


def _migrate_project_status():
    """Add status column to projects if not present."""
    inspector = inspect(engine)
    if "projects" not in inspector.get_table_names():
        return
    existing = {col["name"] for col in inspector.get_columns("projects")}
    if "status" not in existing:
        with engine.connect() as conn:
            conn.execute(text("ALTER TABLE [projects] ADD [status] NVARCHAR(20) NOT NULL DEFAULT 'active'"))
            conn.commit()


_migrate_project_status()

app = FastAPI(title="ProjectTools API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(projects.router)
app.include_router(risk_matrices.router)
app.include_router(communication_plans.router)
app.include_router(meeting_plans.router)
app.include_router(runbooks.router)
app.include_router(project_plans.router)
app.include_router(oppgaver.router)
app.include_router(templates.router)


@app.get("/health")
def health():
    return {"status": "ok"}
