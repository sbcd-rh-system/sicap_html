import argparse
import json
import unicodedata
import re
import sys
from pathlib import Path
from datetime import datetime
import shutil
import math

import pandas as pd
import requests

# === CONFIGURAÇÕES (mesmas dos scripts originais) ===
ARQUIVO_EXCEL_DEFAULT = "PersonalMed - PSM Santana (out.25).xlsx"
ARQUIVO_JSON_MAPEAMENTOS = "Utils\\mapeamentos.json"

# API
API_BASE_URL = "https://sicap.prefeitura.sp.gov.br/v1"
LOGIN_ENDPOINT = f"{API_BASE_URL}/Autenticacao/Login"
FOLHA_PJ_ENDPOINT = f"{API_BASE_URL}/FolhaPagamentoPessoaJuridica"

# Credenciais (reaproveitadas do repo)
USUARIO = "amanda.kawauchi"
SENHA = "Am280309#"

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

def fazer_login():
    print(" Fazendo login na API SICAP...")
    payload = {"login": USUARIO, "senha": SENHA}
    headers = {"Content-Type": "application/json"}
    r = requests.post(LOGIN_ENDPOINT, json=payload, headers=headers)
    r.raise_for_status()
    data = r.json()
    token = data.get("token") or data.get("Token") or data.get("access_token") or data.get("accessToken")
    if not token:
        raise ValueError("Token não encontrado na resposta da API")
    print(" Login realizado com sucesso!")
    return token

def enviar_folha_pj(token, payload):
    print("\n Enviando folha de pagamento para SICAP...")
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {token}"}
    # envia e captura resposta mesmo em caso de erro (status >=400)
    r = requests.post(FOLHA_PJ_ENDPOINT, json=payload, headers=headers)
    api_num = None
    resultado = None
    try:
        resultado = r.json()
    except ValueError:
        resultado = None

    # imprime resposta bruta para debugging (JSON ou texto)
    if resultado is not None:
        try:
            print(f"Resposta da API: {json.dumps(resultado, indent=2, ensure_ascii=False)}")
        except Exception:
            print(f"Resposta da API (json parse ok, but print failed): {resultado}")
        if isinstance(resultado, dict) and resultado.get("NumNotaFiscal"):
            api_num = resultado.get("NumNotaFiscal")
    else:
        txt = r.text.strip()
        if txt:
            print(f"Resposta da API (texto): {txt}")

    # se houve erro HTTP, exiba também o código para ajudar no diagnóstico
    sent_num = payload.get("NumNotaFiscal")
    total_prest = len(payload.get("Prestadores", []))
    if r.status_code >= 400:
        print(f"ERRO HTTP {r.status_code} ao enviar (ver resposta acima).\nTotal de prestadores no payload: {total_prest}\nNumNotaFiscal (payload): {sent_num}")
    else:
        if api_num is not None:
            print(f"NumNotaFiscal retornado pela API: {api_num}")
        else:
            print(f"NumNotaFiscal (do payload): {sent_num}")
        print(f"Total de prestadores enviados: {total_prest}")

    return r

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", help="Caminho para a planilha Excel (opcional)")
    parser.add_argument("--id", help="ID da Prestação de Contas manual (ignora mapeamento)")
    parser.add_argument("--dry-run", action="store_true", help="Gerar JSON apenas, sem enviar")
    parser.add_argument("--save-on-success", action="store_true", help="Salvar JSON e mover planilha para Enviados apenas se o envio for bem-sucedido")
    args = parser.parse_args()

    excel_path = args.excel or ARQUIVO_EXCEL_DEFAULT
    excel_path = str(excel_path)
    if not Path(excel_path).exists():
        print(f"Arquivo Excel não encontrado: {excel_path}")
        sys.exit(1)

    with open(ARQUIVO_JSON_MAPEAMENTOS, "r", encoding="utf-8") as f:
        MAPAS = json.load(f)

    m = re.search(r"\(([A-Za-z]{3})[\.)]", excel_path)
    if m:
        mes_ref = m.group(1).lower()
        print(f"Mês detectado automaticamente a partir do nome do arquivo: '{mes_ref}'")
    else:
        print("Não foi possível detectar mês no nome do arquivo. Passe --excel com padrão (xxx.yy) ou ajuste o nome.")
        sys.exit(1)

    prestacao_id = args.id
    if prestacao_id:
        print(f"Usando ID de Prestação manual: {prestacao_id}")
    else:
        prestacao_id = MAPAS["PrestacaoContaId"].get(mes_ref)
        if prestacao_id is None:
            raise ValueError(f"Mês '{mes_ref}' não encontrado no JSON de mapeamentos e nenhum --id foi informado.")
        print(f"ID detectado automaticamente: {prestacao_id}")

    ABA_EMPRESA = "600"
    ABA_PRESTADORES = "610"
    df_emp = pd.read_excel(excel_path, sheet_name=ABA_EMPRESA)
    df = pd.read_excel(excel_path, sheet_name=ABA_PRESTADORES)

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
    for key, example in expected.items():
        col = find_column(df, example)
        if not col:
            raise KeyError(f"Coluna esperada '{example}' (para '{key}') não encontrada. Colunas disponíveis: {', '.join(map(str, df.columns))}")
        cols[key] = col

    const_tipo_coordenadoria = mapear("SEMPRE", "TipoCoordenadoria", MAPAS) or 0
    const_tipo_atividade = mapear("SEMPRE", "TipoAtividade", MAPAS) or 0

    primeira_unidade = df[cols["Unidade"]].iloc[0]
    nome_unidade = primeira_unidade.replace("PSM ", "").replace(" - LAURO RIBAS BRAGA", "").strip()
    nome_unidade_safe = sanitize_filename(nome_unidade) or mes_ref
    saida_json_name = f"sicap_enviar_{nome_unidade_safe}_{mes_ref}.json"

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
        print("\n🚨 ERRO: Há unidades sem mapeamento (UnidadeId = 0). Atualize Utils/mapeamentos.json")
        for u in unmapped[:50]:
            print(f"   - '{u}'")
        sys.exit(1)

    # validações extras antes do envio: CargoId, LinhaServicoId, CPF
    problemas = []

    # cargos sem id
    mask_cargo0 = saida['CargoId'] == 0
    if mask_cargo0.any():
        indices = list(saida[mask_cargo0].index)
        linhas = []
        for i in indices:
            excel_row = i + 2
            nome = df.at[i, cols['Nome']]
            orig_cargo = df.at[i, cols['CargoId']]
            linhas.append((excel_row, nome, orig_cargo))
        problemas.append(('CargoId ausente', linhas))

    # linha de serviço sem id
    mask_linha0 = saida['LinhaServicoId'] == 0
    if mask_linha0.any():
        indices = list(saida[mask_linha0].index)
        linhas = []
        for i in indices:
            excel_row = i + 2
            nome = df.at[i, cols['Nome']]
            orig_val = df.at[i, cols['LinhaServicoId']]
            linhas.append((excel_row, nome, orig_val))
        problemas.append(('LinhaServicoId ausente', linhas))

    # CPF inválido — validar usando CPF já normalizado/zero-filled em `saida`
    mask_cpf_inv = []
    cpf_series = saida['CPF'].astype(str).fillna("")
    for i, raw in cpf_series.items():
        if not is_valid_cpf(raw):
            mask_cpf_inv.append(i)
    if mask_cpf_inv:
        linhas = []
        for i in mask_cpf_inv:
            excel_row = i + 2
            nome = df.at[i, cols['Nome']]
            raw_orig = df.at[i, cols['CPF']]
            cleaned = saida.at[i, 'CPF']
            linhas.append((excel_row, nome, f"{raw_orig} -> {cleaned}"))
        problemas.append(('CPF inválido', linhas))

    if problemas:
        header = f"ERROS DE VALIDAÇÃO ANTES DO ENVIO - Arquivo: {excel_path}"
        print('\n' + '='*len(header))
        print(header)
        print('='*len(header))
        for tipo, linhas in problemas:
            print(f"\n- {tipo}: {len(linhas)} ocorrência(s)")
            for r, nome, val in linhas[:50]:
                print(f"   linha {r}: {nome} | valor: '{val}'")
        print('\nCopie e cole o bloco acima para o responsável corrigir a planilha.')
        sys.exit(1)

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

    payload["SourceArquivo"] = excel_path

    out_path = Path(saida_json_name)

    print(f"Total de prestadores: {len(prestadores_lista)}")

    # Opções de salvamento: por padrão envia em memória e NÃO salva o JSON
    # para compatibilidade com solicitações que não queiram arquivos.
    # --dry-run: gera o JSON e não envia (útil para inspeção)
    # --save-on-success: salva o JSON e move a planilha para Enviados apenas se o envio for bem-sucedido
    if args.dry_run:
        # grava o JSON para inspeção quando em dry-run
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"--dry-run ativo: JSON gerado em {out_path}, sem envio.")
        return

    token = fazer_login()
    resp = enviar_folha_pj(token, payload)
    if resp is not None and resp.ok:
        enviados_dir = Path("Enviados")
        enviados_dir.mkdir(parents=True, exist_ok=True)

        def _unique_target(dest: Path) -> Path:
            if not dest.exists():
                return dest
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            return dest.with_name(f"{dest.stem}_{ts}{dest.suffix}")

        # mover sempre a planilha após envio bem-sucedido
        src_x = Path(payload.get("SourceArquivo"))
        if src_x.exists():
            tgt_x = enviados_dir / src_x.name
            tgt_x = _unique_target(tgt_x)
            shutil.move(str(src_x), str(tgt_x))
            print(f"Planilha movida para: {tgt_x}")
        else:
            print(f"Planilha não encontrada para mover: {payload.get('SourceArquivo')}")

        # somente salvar o JSON se solicitado
        if args.save_on_success:
            # grava JSON apenas após sucesso
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            tgt_json = enviados_dir / out_path.name
            tgt_json = _unique_target(tgt_json)
            shutil.move(str(out_path), str(tgt_json))
            print(f"Arquivo JSON movido para: {tgt_json}")
    else:
        # resumo amigável e copiable para encaminhar ao responsável
        resp_text = None
        resp_json = None
        try:
            resp_json = resp.json()
        except Exception:
            resp_text = resp.text

        header = f"ERRO NO ENVIO - Arquivo: {excel_path} | NumNotaFiscal: {payload.get('NumNotaFiscal')}"
        print('\n' + '='*len(header))
        print(header)
        print('='*len(header))

        print(f"HTTP status: {resp.status_code}")
        print(f"Total de prestadores no payload: {len(payload.get('Prestadores', []))}")

        if resp_json is not None:
            # tenta extrair mensagens/erros do JSON
            def _collect_messages(o):
                msgs = []
                if isinstance(o, dict):
                    for k, v in o.items():
                        if isinstance(v, (str, int, float)):
                            msgs.append(f"{k}: {v}")
                        else:
                            msgs.extend(_collect_messages(v))
                elif isinstance(o, list):
                    for it in o:
                        msgs.extend(_collect_messages(it))
                else:
                    msgs.append(str(o))
                return msgs

            messages = _collect_messages(resp_json)
            if messages:
                print("\nMensagens de erro retornadas pela API:")
                for m in messages:
                    print(" - ", m)
            else:
                print("Resposta JSON sem mensagens legíveis; veja a resposta completa abaixo:")
                print(json.dumps(resp_json, ensure_ascii=False, indent=2))
        else:
            print("Resposta da API (texto):")
            print(resp_text)

        print('\nCopie e cole o bloco acima para o responsável corrigir a planilha ou diagnosticar o erro.')

if __name__ == '__main__':
    main()
