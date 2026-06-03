from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from .database import engine, Base
from .routers import projects, risk_matrices, communication_plans, meeting_plans, runbooks, project_plans, oppgaver, templates

Base.metadata.create_all(bind=engine)

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
