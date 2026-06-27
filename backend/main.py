import os
import logging
from pathlib import Path
from typing import Annotated, Any

import httpx
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from google.auth.transport import requests as google_requests
from google.oauth2 import id_token
from pydantic import BaseModel, Field

try:
    import psycopg
except ImportError:
    psycopg = None

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

try:
    from .datalex_auditoria import PROMPT_INFORME, fila_a_columnas, guardar_exportacion_web, guardar_intercambio_web
except ImportError:
    
    PROMPT_INFORME = None
    fila_a_columnas = None
    guardar_exportacion_web = None
    guardar_intercambio_web = None
    DATALEX_LOGGER_IMPORT_ERROR = exc
else:
    DATALEX_LOGGER_IMPORT_ERROR = None



GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "DataLex1581_v2")
AUDIT_LOG_DIR_CONFIG = Path(os.getenv("AUDIT_LOG_DIR", "auditorias"))
AUDIT_LOG_DIR = str(AUDIT_LOG_DIR_CONFIG if AUDIT_LOG_DIR_CONFIG.is_absolute() else BASE_DIR / AUDIT_LOG_DIR_CONFIG)
FRONTEND_ORIGINS = [
    origin.strip()
    for origin in os.getenv(
        "FRONTEND_ORIGINS",
        "http://localhost:5173,http://localhost:3000,http://127.0.0.1:8000,null",
    ).split(",")
    if origin.strip()
]
FRONTEND_FILE = BASE_DIR.parent / "frontend" / "index.html"

logger = logging.getLogger(__name__)

app = FastAPI(title="Hackathon Backend", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=FRONTEND_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

bearer_scheme = HTTPBearer(auto_error=False)


class GoogleUser(BaseModel):
    sub: str
    email: str | None = None
    name: str | None = None
    picture: str | None = None


class ChatMessage(BaseModel):
    role: str = Field(pattern="^(system|user|assistant)$")
    content: str = Field(min_length=1)


class ChatRequest(BaseModel):
    message: str = Field(min_length=1)
    history: list[ChatMessage] = Field(default_factory=list)


class ChatResponse(BaseModel):
    answer: str
    model: str


class MeResponse(BaseModel):
    user: GoogleUser


class AuditProfile(BaseModel):
    empresa: str = Field(min_length=1)
    nit: str = ""
    usuario: str = Field(min_length=1)
    tamano_empresa: str = Field(min_length=1)
    sector: str = Field(min_length=1)


class AuditStartRequest(BaseModel):
    profile: AuditProfile


class AuditChatRequest(BaseModel):
    message: str = Field(min_length=1)
    history: list[ChatMessage] = Field(default_factory=list)
    profile: AuditProfile


class AuditChatResponse(BaseModel):
    answer: str
    model: str
    excel_path: str | None = None
    postgres_saved: bool = False


class AuditExportRequest(BaseModel):
    history: list[ChatMessage] = Field(default_factory=list)
    profile: AuditProfile


class AuditExportResponse(BaseModel):
    answer: str
    model: str
    excel_path: str | None = None
    matrix_rows: int
    raw_saved: bool
    json_valid: bool
    postgres_saved: bool


def db_config() -> dict[str, str]:
    return {
        "host": os.getenv("DB_HOST", os.getenv("HOST", "localhost")),
        "port": os.getenv("DB_PORT", os.getenv("PORT", "5432")),
        "dbname": os.getenv("DB_NAME", "postgres"),
        "user": os.getenv("DB_USER", "postgres"),
        "password": os.getenv("DB_PASSWORD", ""),
        "connect_timeout": os.getenv("DB_CONNECT_TIMEOUT", "5"),
        "options": "-c client_encoding=UTF8",
    }


def db_connect():
    if psycopg is None:
        raise RuntimeError("psycopg is not installed")
    return psycopg.connect(**db_config())


def ensure_database() -> None:
    with db_connect() as conn:
        with conn.cursor() as cur:
            cur.execute("CREATE SCHEMA IF NOT EXISTS hackaton")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS hackaton.conversacion (
                    fecha_hora TIMESTAMP NOT NULL,
                    usuario TEXT NOT NULL,
                    rol TEXT NOT NULL,
                    turno INTEGER NOT NULL,
                    contenido TEXT NOT NULL
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS hackaton.matriz (
                    usuario TEXT NOT NULL,
                    dominio TEXT,
                    criterio TEXT,
                    evidencia TEXT,
                    documentacion TEXT,
                    control TEXT,
                    seguimiento TEXT,
                    obsevaciones TEXT,
                    confianza TEXT
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS hackaton.crudo (
                    usuario TEXT NOT NULL,
                    fecha_hora TIMESTAMP NOT NULL,
                    respuesta_cruda_modelo TEXT NOT NULL
                )
                """
            )


def guardar_conversacion_postgres(usuario: str, turno: int, rol: str, contenido: str) -> bool:
    try:
        ensure_database()
        with db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO hackaton.conversacion (fecha_hora, usuario, rol, turno, contenido)
                    VALUES (CURRENT_TIMESTAMP, %s, %s, %s, %s)
                    """,
                    (usuario, rol, turno, contenido),
                )
        return True
    except Exception as exc:
        logger.warning("Could not save conversation row to PostgreSQL: %s", exc)
        return False


def guardar_crudo_postgres(usuario: str, respuesta_cruda: str) -> bool:
    try:
        ensure_database()
        with db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO hackaton.crudo (usuario, fecha_hora, respuesta_cruda_modelo)
                    VALUES (%s, CURRENT_TIMESTAMP, %s)
                    """,
                    (usuario, respuesta_cruda),
                )
        return True
    except Exception as exc:
        logger.warning("Could not save raw model response to PostgreSQL: %s", exc)
        return False


def guardar_matriz_postgres(usuario: str, filas: list[Any]) -> bool:
    if not filas:
        return False
    if fila_a_columnas is None:
        logger.warning("DATALEX matrix parser is not available")
        return False

    try:
        ensure_database()
        with db_connect() as conn:
            with conn.cursor() as cur:
                for fila in filas:
                    if not isinstance(fila, dict):
                        continue
                    valores = fila_a_columnas(fila)
                    cur.execute(
                        """
                        INSERT INTO hackaton.matriz (
                            usuario, dominio, criterio, evidencia, documentacion,
                            control, seguimiento, obsevaciones, confianza
                        )
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            usuario,
                            valores[0],
                            valores[1],
                            valores[2],
                            valores[3],
                            valores[4],
                            valores[5],
                            valores[6],
                            valores[7],
                        ),
                    )
        return True
    except Exception as exc:
        logger.warning("Could not save matrix rows to PostgreSQL: %s", exc)
        return False


def perfil_a_contexto(profile: AuditProfile) -> str:
    return (
        "TAREA: REDACTAR PREGUNTAS. "
        "Redacta la primera pregunta de una entrevista de auditoria sobre cumplimiento de la Ley 1581 de 2012. "
        "Responde UNICAMENTE la pregunta, sin JSON, sin explicaciones y sin listas. "
        "La pregunta debe estar adaptada a este perfil: "
        f"Empresa: {profile.empresa}. "
        f"NIT: {profile.nit or 'No informado'}. "
        f"Tamaño de la empresa: {profile.tamano_empresa}. "
        f"Sector: {profile.sector}. "
        "Empieza por verificar autorizacion o consentimiento para el tratamiento de datos personales."
    )


def prompt_siguiente_pregunta(profile: AuditProfile, respuesta_usuario: str) -> str:
    return (
        "TAREA: REDACTAR PREGUNTAS. "
        "El auditado acaba de responder lo siguiente a la pregunta anterior: "
        f"{respuesta_usuario}. "
        "Con base en el historial de la entrevista y en esa respuesta, redacta UNICAMENTE la siguiente pregunta "
        "necesaria para continuar la auditoria de cumplimiento de la Ley 1581 de 2012. "
        "No entregues diagnostico, porcentajes, recomendaciones ni JSON. "
        "No repitas una pregunta ya contestada. "
        "Adapta la pregunta a este perfil: "
        f"Empresa: {profile.empresa}; tamaño: {profile.tamano_empresa}; sector: {profile.sector}."
    )


async def preguntar_ollama(messages: list[dict[str, str]], timeout: float = 90.0) -> str:
    try:
        async with httpx.AsyncClient(timeout=timeout) as client:
            ollama_response = await client.post(
                f"{OLLAMA_BASE_URL.rstrip('/')}/api/chat",
                json={
                    "model": OLLAMA_MODEL,
                    "messages": messages,
                    "stream": False,
                },
            )
            ollama_response.raise_for_status()
    except httpx.ConnectError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Could not connect to local Ollama. Make sure it is running.",
        ) from exc
    except httpx.HTTPStatusError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Ollama returned an error: {exc.response.text}",
        ) from exc
    except httpx.HTTPError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Unexpected error while asking Ollama",
        ) from exc

    data: dict[str, Any] = ollama_response.json()
    answer = data.get("message", {}).get("content")
    if not answer:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Ollama response did not include an answer",
        )
    return answer


def guardar_auditoria_chat(user: GoogleUser, mensaje: str, respuesta: str) -> None:
    if guardar_intercambio_web is None:
        logger.warning("DATALEX logger is not available: %s", DATALEX_LOGGER_IMPORT_ERROR)
        return

    usuario = user.email or user.sub
    try:
        ruta = guardar_intercambio_web(usuario, mensaje, respuesta, AUDIT_LOG_DIR)
    except Exception:
        logger.exception("Could not save DATALEX audit log")
        return

    logger.info("DATALEX audit log saved to %s", ruta)


async def verify_google_user(
    credentials: Annotated[HTTPAuthorizationCredentials | None, Depends(bearer_scheme)],
) -> GoogleUser:
    if credentials is None or credentials.scheme.lower() != "bearer":
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token",
        )

    if not GOOGLE_CLIENT_ID:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="GOOGLE_CLIENT_ID is not configured",
        )

    try:
        token_info = id_token.verify_oauth2_token(
            credentials.credentials,
            google_requests.Request(),
            GOOGLE_CLIENT_ID,
        )
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid Google token",
        ) from exc

    return GoogleUser(
        sub=token_info["sub"],
        email=token_info.get("email"),
        name=token_info.get("name"),
        picture=token_info.get("picture"),
    )


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/")
async def frontend() -> FileResponse:
    if not FRONTEND_FILE.exists():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="frontend/index.html not found")
    return FileResponse(FRONTEND_FILE)


@app.get("/auth/me", response_model=MeResponse)
async def me(user: Annotated[GoogleUser, Depends(verify_google_user)]) -> MeResponse:
    return MeResponse(user=user)


@app.post("/audit/start", response_model=AuditChatResponse)
async def audit_start(payload: AuditStartRequest) -> AuditChatResponse:
    usuario = payload.profile.usuario
    prompt = perfil_a_contexto(payload.profile)
    answer = await preguntar_ollama([{"role": "user", "content": prompt}])

    excel_path = None
    if guardar_intercambio_web is not None:
        try:
            excel_path = guardar_intercambio_web(usuario, prompt, answer, AUDIT_LOG_DIR)
        except Exception:
            logger.exception("Could not save DATALEX audit start to Excel")

    turno_ok = guardar_conversacion_postgres(usuario, 1, "usuario", prompt)
    answer_ok = guardar_conversacion_postgres(usuario, 2, "datalex", answer)

    return AuditChatResponse(
        answer=answer,
        model=OLLAMA_MODEL,
        excel_path=excel_path,
        postgres_saved=turno_ok and answer_ok,
    )


@app.post("/audit/chat", response_model=AuditChatResponse)
async def audit_chat(payload: AuditChatRequest) -> AuditChatResponse:
    usuario = payload.profile.usuario
    messages = [message.model_dump() for message in payload.history]
    messages.append({"role": "user", "content": prompt_siguiente_pregunta(payload.profile, payload.message)})
    answer = await preguntar_ollama(messages)

    excel_path = None
    if guardar_intercambio_web is not None:
        try:
            excel_path = guardar_intercambio_web(usuario, payload.message, answer, AUDIT_LOG_DIR)
        except Exception:
            logger.exception("Could not save DATALEX audit chat to Excel")

    turno = len(payload.history) + 1
    turno_ok = guardar_conversacion_postgres(usuario, turno, "usuario", payload.message)
    answer_ok = guardar_conversacion_postgres(usuario, turno + 1, "datalex", answer)

    return AuditChatResponse(
        answer=answer,
        model=OLLAMA_MODEL,
        excel_path=excel_path,
        postgres_saved=turno_ok and answer_ok,
    )


@app.post("/audit/export", response_model=AuditExportResponse)
async def audit_export(payload: AuditExportRequest) -> AuditExportResponse:
    if PROMPT_INFORME is None or guardar_exportacion_web is None:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="DATALEX logger is not available",
        )

    usuario = payload.profile.usuario
    messages = [message.model_dump() for message in payload.history]
    messages.append({"role": "user", "content": PROMPT_INFORME})
    answer = await preguntar_ollama(messages, timeout=300.0)

    exportacion = guardar_exportacion_web(usuario, answer, AUDIT_LOG_DIR)
    raw_ok = guardar_crudo_postgres(usuario, answer)
    matrix_ok = guardar_matriz_postgres(usuario, exportacion["filas"])
    turno = len(payload.history) + 1
    request_ok = guardar_conversacion_postgres(usuario, turno, "usuario", "[Solicitud de informe estructurado]")
    answer_ok = guardar_conversacion_postgres(usuario, turno + 1, "datalex", answer)

    return AuditExportResponse(
        answer=answer,
        model=OLLAMA_MODEL,
        excel_path=exportacion["ruta"],
        matrix_rows=exportacion["filas_escritas"],
        raw_saved=raw_ok,
        json_valid=exportacion["json_valido"],
        postgres_saved=raw_ok and matrix_ok and request_ok and answer_ok,
    )


@app.post("/chat", response_model=ChatResponse)
async def chat(
    payload: ChatRequest,
    user: Annotated[GoogleUser, Depends(verify_google_user)],
) -> ChatResponse:
    messages = [message.model_dump() for message in payload.history]
    messages.append(
        {
            "role": "user",
            "content": payload.message,
        }
    )

    try:
        async with httpx.AsyncClient(timeout=90.0) as client:
            ollama_response = await client.post(
                f"{OLLAMA_BASE_URL.rstrip('/')}/api/chat",
                json={
                    "model": OLLAMA_MODEL,
                    "messages": messages,
                    "stream": False,
                    "options": {
                        "temperature": 0.7,
                    },
                },
            )
            ollama_response.raise_for_status()
    except httpx.ConnectError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Could not connect to local Ollama. Make sure it is running.",
        ) from exc
    except httpx.HTTPStatusError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Ollama returned an error: {exc.response.text}",
        ) from exc
    except httpx.HTTPError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Unexpected error while asking Ollama",
        ) from exc

    data: dict[str, Any] = ollama_response.json()
    answer = data.get("message", {}).get("content")
    if not answer:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Ollama response did not include an answer",
        )

    guardar_auditoria_chat(user, payload.message, answer)

    return ChatResponse(answer=answer, model=OLLAMA_MODEL)
