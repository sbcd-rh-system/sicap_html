from fastapi import FastAPI, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
import shutil
import os
import uuid

try:
    from .processor import processar_planilha
except ImportError:
    from processor import processar_planilha

app = FastAPI(title="SICAP Uploader", version="1.0.0")

# Configurar CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Diretórios - No Render, usamos /tmp para garantir permissão de escrita
UPLOAD_DIR = "/tmp/sicap_uploads" if os.name != 'nt' else os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Diretório do Frontend
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(os.path.dirname(CURRENT_DIR), "frontend")
if not os.path.exists(FRONTEND_DIR):
    FRONTEND_DIR = os.path.join(os.getcwd(), "frontend")

# ============================================================
# ROTAS DA API  (registradas ANTES do mount estático)
# ============================================================

@app.get("/health")
async def health():
    """Endpoint de diagnóstico para confirmar que a API está viva."""
    return {"status": "ok", "frontend_dir": FRONTEND_DIR, "upload_dir": UPLOAD_DIR}


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
        return JSONResponse(
            status_code=400,
            content={"status": "erro", "mensagem": "Formato de arquivo inválido. Use .xlsx ou .xls"}
        )

    filename = f"{uuid.uuid4()}_{file.filename}"
    file_path = os.path.join(UPLOAD_DIR, filename)

    try:
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        resultado = processar_planilha(file_path, usuario, senha, mes, ano, prestacao_id)

        if resultado["status"] == "erro":
            return JSONResponse(status_code=422, content=resultado)

        return JSONResponse(status_code=200, content=resultado)

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "erro",
                "mensagem": f"Erro interno do servidor: {str(e)}",
                "detalhes": {}
            }
        )
    finally:
        # Limpa o arquivo enviado depois de processar
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass


# ============================================================
# ARQUIVOS ESTÁTICOS (montado POR ÚLTIMO)
# Serve CSS, JS, imagens do frontend
# ============================================================
if os.path.exists(FRONTEND_DIR):
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")

    # Serve o index.html na raiz e em qualquer rota não encontrada
    @app.get("/")
    async def serve_root():
        return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))

    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str):
        file_candidate = os.path.join(FRONTEND_DIR, full_path)
        if os.path.isfile(file_candidate):
            return FileResponse(file_candidate)
        return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))

else:
    print(f"Erro Crítico: Diretório frontend não encontrado em {FRONTEND_DIR}")


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
