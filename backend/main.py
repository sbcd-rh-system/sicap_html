from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import shutil
import os
import uuid
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

# Diretórios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Endpoint de Upload e Processamento
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

    # Gerar nome único
    filename = f"{uuid.uuid4()}_{file.filename}"
    file_path = os.path.join(UPLOAD_DIR, filename)

    try:
        # Salvar arquivo
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Chamar processador com novos parâmetros
        resultado = processar_planilha(file_path, usuario, senha, mes, ano, prestacao_id)
        
        # Opcional: remover arquivo
        # os.remove(file_path)
        
        if resultado["status"] == "erro":
            return JSONResponse(status_code=500, content=resultado)
            
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

# Servir Frontend
# Tenta localizar o diretório frontend em relação à raiz do projeto
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(os.path.dirname(CURRENT_DIR), "frontend")

if not os.path.exists(FRONTEND_DIR):
    # Fallback caso a estrutura mude no deploy do Render
    FRONTEND_DIR = os.path.join(os.getcwd(), "frontend")

if os.path.exists(FRONTEND_DIR):
    app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
else:
    print(f"Erro Crítico: Diretório frontend não encontrado em {FRONTEND_DIR}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
