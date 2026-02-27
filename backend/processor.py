import pandas as pd
import logging
import time
import os
import json
import requests
import unicodedata
import re
import math
import shutil
from datetime import datetime
from pathlib import Path

# Configuração de logs
log_dir = "/tmp/sicap_logs" if os.name != 'nt' else os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"sicap_log_{datetime.now().strftime('%Y%m%d')}.log")

try:
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
except Exception:
    # Fallback para console caso não tenha permissão de escrita (Render/Produção)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logging.info("Logging configurado para Console (StreamHandler)")

# API
API_BASE_URL = "https://sicap.prefeitura.sp.gov.br/v1"
LOGIN_ENDPOINT = f"{API_BASE_URL}/Autenticacao/Login"
FOLHA_PJ_ENDPOINT = f"{API_BASE_URL}/FolhaPagamentoPessoaJuridica"

# Caminho para Mapeamentos
# Caminho para Mapeamentos
# Se rodar via 'run.py' na raiz, o BASE_DIR deve ser a própria raiz.
# Se rodar via 'backend/main.py', sobe um nível.
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if not os.path.exists(os.path.join(BASE_DIR, "Utils")):
    BASE_DIR = os.path.dirname(os.path.abspath(__file__)) # Fallback local
ARQUIVO_JSON_MAPEAMENTOS = os.path.join(BASE_DIR, "Utils", "mapeamentos.json")

# ==================================================================================
# HELPER FUNCTIONS
# ==================================================================================

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto

def _normalize_col_name(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def find_column(df, example):
    cols_norm = { _normalize_col_name(c): c for c in df.columns }
    target_norm = _normalize_col_name(example)
    if target_norm in cols_norm:
        return cols_norm[target_norm]
    words = re.findall(r'\w+', target_norm)
    for norm, orig in cols_norm.items():
        if all(w in norm for w in words):
            return orig
    for norm, orig in cols_norm.items():
        if 'cns' in norm:
            return orig
    return None

def sanitize_filename(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"[\\/]+", "-", s)
    s = re.sub(r'[:\*\?"<>|]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def parse_money(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor = valor.replace("R$", "").strip()
        if "," in valor and "." in valor:
            if valor.rindex(",") > valor.rindex("."):
                valor = valor.replace(".", "").replace(",", ".")
            else:
                valor = valor.replace(",", "")
        elif "," in valor:
            valor = valor.replace(",", ".")
        return float(valor)
    return float(valor)

def mapear(valor, categoria, mapas):
    val = normalizar_texto(valor)
    mapa_categoria = mapas.get(categoria, {})

    if categoria == "LinhaServicoId":
        extra = mapas.get("LinhasDeServico", {})
        if extra:
            merged = dict(mapa_categoria)
            merged.update(extra)
            mapa_categoria = merged
    try:
        if categoria in ("LinhaServicoId", "Unidade"):
            if isinstance(valor, (int, float)) and not pd.isna(valor):
                return int(valor)
            if isinstance(valor, str) and valor.strip().isdigit():
                return int(valor.strip())
    except Exception:
        pass

    mapa_normalizado = {normalizar_texto(k): v for k, v in mapa_categoria.items()}
    if val in mapa_normalizado:
        return mapa_normalizado[val]

    if categoria == "CargoId":
        for chave_orig, id_cargo in mapa_categoria.items():
            chave_norm = normalizar_texto(chave_orig)
            if chave_norm in val or val in chave_norm:
                return id_cargo

    if categoria == "Unidade":
        for chave_orig, id_un in mapa_categoria.items():
            chave_norm = normalizar_texto(chave_orig)
            if chave_norm and (chave_norm in val or val in chave_norm):
                return id_un
        words_val = set(re.findall(r'\w+', val))
        for chave_orig, id_un in mapa_categoria.items():
            chave_norm = normalizar_texto(chave_orig)
            words_ch = set(re.findall(r'\w+', chave_norm))
            if words_ch and words_ch.issubset(words_val):
                return id_un

    return None

def _only_digits(s: str) -> str:
    return re.sub(r'\D', '', str(s or ''))

def is_valid_cpf(cpf: str) -> bool:
    s = _only_digits(cpf)
    if len(s) != 11:
        return False
    if s == s[0] * 11:
        return False
    def _calc(digs):
        ssum = sum(int(a) * b for a, b in zip(digs, range(len(digs)+1, 1, -1)))
        rem = (ssum * 10) % 11
        return rem if rem < 10 else 0
    try:
        d1 = _calc(s[:9])
        d2 = _calc(s[:9] + str(d1))
        return int(s[9]) == d1 and int(s[10]) == d2
    except Exception:
        return False

# ==================================================================================
# API & LOGIC
# ==================================================================================

def fazer_login(usuario, senha):
    logging.info(f"Fazendo login na API SICAP para usuario: {usuario}")
    payload = {"login": usuario, "senha": senha}
    headers = {"Content-Type": "application/json"}
    
    try:
        r = requests.post(LOGIN_ENDPOINT, json=payload, headers=headers, timeout=30)
        r.raise_for_status()
        data = r.json()
        token = data.get("token") or data.get("Token") or data.get("access_token") or data.get("accessToken")
        if not token:
            raise ValueError("Token não encontrado na resposta da API")
        logging.info("Login realizado com sucesso!")
        return token
    except Exception as e:
        logging.error(f"Erro no login: {e}")
        if 'r' in locals() and r:
             logging.error(f"Response login: {r.text}")
        raise ValueError(f"Falha na autenticação: {str(e)}")

def enviar_folha_pj(token, payload):
    logging.info("Enviando folha de pagamento para SICAP...")
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {token}"}
    
    try:
        r = requests.post(FOLHA_PJ_ENDPOINT, json=payload, headers=headers, timeout=120)
        return r
    except Exception as e:
        logging.error(f"Erro ao enviar folha: {e}")
        raise ConnectionError(f"Erro na conexão com SICAP: {str(e)}")

def processar_planilha(caminho_arquivo: str, usuario: str, senha: str, mes: str = None, ano: str = None, prestacao_id: any = None) -> dict:
    start_time = time.time()
    try:
        logging.info(f"Iniciando processamento do arquivo: {caminho_arquivo}")
        logging.info(f"Parâmetros recebidos: Mes={mes}, Ano={ano}")

        if not os.path.exists(ARQUIVO_JSON_MAPEAMENTOS):
             logging.error(f"Arquivo de mapeamentos não encontrado em: {ARQUIVO_JSON_MAPEAMENTOS}")
             return {
                 "status": "erro",
                 "mensagem": "Arquivo Utils/mapeamentos.json não encontrado no servidor. Contate o suporte.",
                 "detalhes": {"caminho_esperado": ARQUIVO_JSON_MAPEAMENTOS}
             }

        with open(ARQUIVO_JSON_MAPEAMENTOS, "r", encoding="utf-8") as f:
            MAPAS = json.load(f)
            
        # Determinação do mês de referência
        mes_ref = None
        if mes:
            mes_ref = mes.lower() # garantir lowercase (jan, fev...)
        else:
            # Fallback para detecção automática (legado)
            m = re.search(r"\(([A-Za-z]{3})[\.)]", os.path.basename(caminho_arquivo))
            if m:
                mes_ref = m.group(1).lower()
                logging.info(f"Mês detectado via nome do arquivo: {mes_ref}")
        
        if not prestacao_id:
             return {
                 "status": "erro",
                 "mensagem": "O ID da Prestação de Contas é obrigatório.",
                 "detalhes": {"acao": "Informe o ID da competência obtido no portal SICAP."}
             }
        
        logging.info(f"Usando PrestacaoContaId: {prestacao_id}")

        # Abas Fixas
        ABA_EMPRESA = "600"
        ABA_PRESTADORES = "610"

        try:
            df_emp = pd.read_excel(caminho_arquivo, sheet_name=ABA_EMPRESA)
            df = pd.read_excel(caminho_arquivo, sheet_name=ABA_PRESTADORES)
        except Exception as e:
            return {
                 "status": "erro",
                 "mensagem": f"Erro ao ler abas da planilha ({ABA_EMPRESA}, {ABA_PRESTADORES}). Verifique o formato.",
                 "detalhes": {"erro_tecnico": str(e)}
            }
        
        # Montar Empresa
        try:
            empresa = {
                "Id": 4623,
                "ParceriaId": 31,
                "PrestacaoContaId": prestacao_id,
                "RazaoSocialEmpresa": df_emp.loc[0, "Razao Social Empresa"],
                "CnpjEmpresa": df_emp.loc[0, "CNPJ Empresa"],
                "ValorBrutoNf": parse_money(df_emp.loc[0, "Valor Bruto NF"]),
                "NumNotaFiscal": str(int(float(df_emp.loc[0, "Nº Nota Fiscal"]))),
                "ValorLiquido": parse_money(df_emp.loc[0, "Valor Liquido"])
            }
        except Exception as e:
            return {
                "status": "erro",
                "mensagem": "Erro ao ler dados da aba Empresa (600). Verifique colunas e valores.",
                "detalhes": {"erro": str(e)}
            }
            
        # Mapeamento e Validação (Mantém lógica anterior)
        expected = {
            "Nome": "Nome Completo",
            "NomeSocial": "Nome Social",
            "CPF": "CPF Funcionário",
            "DataNascimento": "Data Nascimento",
            "AutoDeclaracaoGenero": "Autodeclaração de Gênero",
            "AutoDeclaracaoRacial": "Autodeclaração Racial",
            "CargoId": "Categoria Profissional",
            "NumConselhoClasse": "Nº Conselho de Classe",
            "CnsDoProfissional": "Cns Do Profissional",
            "CargaHorariaSemanalId": "Carga Horária Semanal/Plantão",
            "TurnoTrabalho": "Turno de Trabalho",
            "Unidade": "Unidade",
            "LinhaServicoId": "Linha de Serviço",
            "ValorPorProfissional": "Valor por Profissional",
            "TipoCoordenadoria": "Tipo de Coordenadoria",
            "TipoAtividade": "Tipo de Atividade"
        }
        
        cols = {}
        missing_cols = []
        for key, example in expected.items():
            col = find_column(df, example)
            if not col:
                missing_cols.append(f"{key} (ex: {example})")
            cols[key] = col
        
        if missing_cols:
             return {
                 "status": "erro",
                 "mensagem": "Colunas obrigatórias não encontradas na aba 610.",
                 "detalhes": {"colunas_faltantes": missing_cols}
             }

        const_tipo_coordenadoria = mapear("SEMPRE", "TipoCoordenadoria", MAPAS) or 0
        const_tipo_atividade = mapear("SEMPRE", "TipoAtividade", MAPAS) or 0

        saida = pd.DataFrame({
            "Id": 0,
            "Nome": df[cols["Nome"]].astype(str).str.strip(),
            "NomeSocial": df[cols["NomeSocial"]].astype(str).str.strip(),
            "CPF": df[cols["CPF"]].astype(str).str.replace(r'\D', '', regex=True).str.zfill(11),
            "DataNascimento": pd.to_datetime(df[cols["DataNascimento"]], errors="coerce", dayfirst=True).apply(lambda x: x.strftime("%Y-%m-%dT00:00:00") if pd.notna(x) else "1900-01-01T00:00:00"),
            "AutoDeclaracaoGenero": df[cols["AutoDeclaracaoGenero"]].apply(lambda x: mapear(x, "AutoDeclaracaoGenero", MAPAS)),
            "AutoDeclaracaoRacial": df[cols["AutoDeclaracaoRacial"]].apply(lambda x: mapear(x, "AutoDeclaracaoRacial", MAPAS)),
            "CargoId": df[cols["CargoId"]].apply(lambda x: mapear(x, "CargoId", MAPAS)),
            "NumConselhoClasse": df[cols["NumConselhoClasse"]].astype(str).str.strip(),
            "CnsDoProfissional": df[cols["CnsDoProfissional"]].astype(str).str.strip(),
            "CargaHorariaSemanalId": df[cols["CargaHorariaSemanalId"]].apply(lambda x: mapear(x, "CargaHorariaSemanalId", MAPAS)),
            "TurnoTrabalho": df[cols["TurnoTrabalho"]].apply(lambda x: mapear(x, "TurnoTrabalho", MAPAS)),
            "UnidadeId": df[cols["Unidade"]].apply(lambda x: mapear(x, "Unidade", MAPAS)),
            "LinhaServicoId": df[cols["LinhaServicoId"]].apply(lambda x: mapear(x, "LinhaServicoId", MAPAS)),
            "ValorPorProfissional": df[cols["ValorPorProfissional"]].apply(parse_money),
            "TipoCoordenadoria": [const_tipo_coordenadoria] * len(df),
            "TipoAtividade": [const_tipo_atividade] * len(df),
            "Especificacao": ""
        })

        colunas_num = [
            "AutoDeclaracaoGenero", "AutoDeclaracaoRacial", "CargoId",
            "CargaHorariaSemanalId", "TurnoTrabalho", "UnidadeId",
            "LinhaServicoId", "TipoCoordenadoria", "TipoAtividade"
        ]
        for col in colunas_num:
            saida[col] = pd.to_numeric(saida[col], errors="coerce").fillna(0).astype(int)
            
        orig_unidades = df[cols["Unidade"]].astype(str).fillna("").str.strip()
        unmapped = sorted(set(orig_unidades[saida["UnidadeId"] == 0].unique()))
        unmapped = [u for u in unmapped if u != "" and u.upper() != "NAN"]
        
        if unmapped:
            return {
                "status": "erro",
                "mensagem": "Existem Unidades sem mapeamento (UnidadeId = 0).",
                "detalhes": {"unidades_sem_mapa": unmapped[:50]}
            }
            
        problemas = []
        
        mask_cargo0 = saida['CargoId'] == 0
        if mask_cargo0.any():
            problemas.append("Existem COLABORADORES com Cargo não mapeado (CargoId=0).")
            
        mask_linha0 = saida['LinhaServicoId'] == 0
        if mask_linha0.any():
            problemas.append("Existem COLABORADORES com Linha de Serviço não mapeada (LinhaServicoId=0).")
            
        mask_cpf_inv = []
        cpf_series = saida['CPF'].astype(str).fillna("")
        for i, raw in cpf_series.items():
            if not is_valid_cpf(raw):
                mask_cpf_inv.append(f"Linha {i+2}")
        
        if mask_cpf_inv:
            problemas.append(f"CPFs inválidos detectados: {', '.join(mask_cpf_inv[:20])}...")
            
        if problemas:
            return {
                "status": "erro",
                "mensagem": "Erros de validação pré-envio detectados.",
                "detalhes": {"problemas": problemas}
            }
            
        prestadores_lista = saida.to_dict(orient="records")
        for prestador in prestadores_lista:
            for key, value in list(prestador.items()):
                if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                    if key in ["ValorPorProfissional"]:
                        prestador[key] = 0.0
                    elif key in colunas_num:
                        prestador[key] = 0
                    else:
                        prestador[key] = ""

        payload = {**empresa, "Prestadores": prestadores_lista}
        for key, value in list(payload.items()):
            if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                if key in ["ValorBrutoNf", "ValorLiquido"]:
                    payload[key] = 0.0
                else:
                    payload[key] = ""
                    
        payload["SourceArquivo"] = os.path.basename(caminho_arquivo)
        
        token = fazer_login(usuario, senha)
        r = enviar_folha_pj(token, payload)
        
        elapsed_time = time.time() - start_time
        
        result_json = None
        try:
            result_json = r.json()
        except:
             pass
             
        if r.status_code >= 400:
            logging.error(f"Erro API {r.status_code}: {r.text}")
            return {
                "status": "erro",
                "mensagem": f"Erro retornado pela API SICAP (Status {r.status_code})",
                "detalhes": {
                    "resposta_api": result_json if result_json else r.text,
                    "nota_fiscal": payload.get("NumNotaFiscal")
                }
            }
            
        logging.info(f"Sucesso! NF: {payload.get('NumNotaFiscal')}")
        return {
            "status": "sucesso",
            "mensagem": f"Folha enviada com sucesso! NF: {payload.get('NumNotaFiscal')}",
            "detalhes": {
                "prestadores_enviados": len(prestadores_lista),
                "resposta_sucesso": result_json,
                "tempo": f"{elapsed_time:.2f}s"
            }
        }

    except Exception as e:
        logging.error(f"Exceção não tratada: {str(e)}")
        return {
            "status": "erro",
            "mensagem": f"Erro interno: {str(e)}",
            "detalhes": {
                "tipo_erro": type(e).__name__,
                "log": log_file
            }
        }
