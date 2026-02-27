from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse, Response
import shutil
import os
import uuid
import json
import traceback

try:
    from .processor import processar_planilha
except ImportError:
    from processor import processar_planilha

app = FastAPI(title="SICAP Uploader", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Diretórios
UPLOAD_DIR = "/tmp/sicap_uploads" if os.name != 'nt' else os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(os.path.dirname(CURRENT_DIR), "frontend")
if not os.path.exists(FRONTEND_DIR):
    FRONTEND_DIR = os.path.join(os.getcwd(), "frontend")

# ============================================================
# FRONTEND — servido com FileResponse explícito.
# SEM StaticFiles mount: elimina qualquer conflito de rota.
# ============================================================

@app.get("/")
@app.get("/index.html")
async def serve_index():
    return FileResponse(os.path.join(FRONTEND_DIR, "index.html"), media_type="text/html")

@app.get("/style.css")
async def serve_css():
    return FileResponse(os.path.join(FRONTEND_DIR, "style.css"), media_type="text/css")

@app.get("/app.js")
async def serve_js():
    return FileResponse(os.path.join(FRONTEND_DIR, "app.js"), media_type="application/javascript")

# ============================================================
# API
# ============================================================

@app.get("/health")
async def health():
    return {
        "status": "ok",
        "frontend_dir": FRONTEND_DIR,
        "frontend_exists": os.path.exists(FRONTEND_DIR),
        "upload_dir": UPLOAD_DIR,
    }


@app.post("/api/test")
async def test_post():
    return Response(
        content='{"status":"ok","message":"POST funcionando"}',
        status_code=200,
        media_type="application/json"
    )


@app.post("/api/processar")
async def processar_arquivo(
    file: UploadFile = File(...),
    usuario: str = Form(...),
    senha: str = Form(...),
    mes: str = Form(None),
    ano: str = Form(None),
    prestacao_id: str = Form(None)
):
    if not file.filename.endswith(('.xlsx', '.xls')):
        return Response(
            content='{"status":"erro","mensagem":"Formato invalido. Use .xlsx ou .xls"}',
            status_code=400,
            media_type="application/json"
        )

    filename = f"{uuid.uuid4()}_{file.filename}"
    file_path = os.path.join(UPLOAD_DIR, filename)

    try:
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        resultado = processar_planilha(file_path, usuario, senha, mes, ano, prestacao_id)

        status_code = 422 if resultado.get("status") == "erro" else 200
        return Response(
            content=json.dumps(resultado, ensure_ascii=False, default=str),
            status_code=status_code,
            media_type="application/json"
        )

    except Exception as e:
        erro = {
            "status": "erro",
            "mensagem": f"Erro interno: {str(e)}",
            "detalhes": {"traceback": traceback.format_exc()[-800:]}
        }
        return Response(
            content=json.dumps(erro, ensure_ascii=False),
            status_code=500,
            media_type="application/json"
        )
    finally:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
