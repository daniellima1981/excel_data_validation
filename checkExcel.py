import os
import json
import pandas as pd
from datetime import datetime
import warnings
from io import BytesIO

warnings.filterwarnings("ignore", category=UserWarning)

# ==================================================
# CONFIGURAÇÕES
# ==================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Origem do arquivo: 'local', 'sharepoint', 'azure_blob'
ORIGEM_ARQUIVO = 'local'

# -------- LOCAL / SHAREPOINT (via sync OneDrive) executar localmente--------
CAMINHO_ARQUIVO = r"C:\...\..."
NOME_ARQUIVO = "file_name.xlsx"

# -------- SHAREPOINT  --------
SHAREPOINT_SITE_URL = ""  # Ex: https://empresa.sharepoint.com/sites/meusite
SHAREPOINT_FILE_PATH = ""  # Ex: /Shared Documents/pasta/arquivo.xlsx
SHAREPOINT_CLIENT_ID = ""
SHAREPOINT_CLIENT_SECRET = ""
SHAREPOINT_TENANT_ID = ""

# -------- AZURE BLOB --------
AZURE_CONNECTION_STRING = ""
AZURE_CONTAINER_NAME = ""
AZURE_BLOB_NAME = ""  # Nome do arquivo no blob

# -------- OUTPUT --------
OUTPUT_LAYOUT_FILE = os.path.join(BASE_DIR, "layout_baseline.json")
LOG_FILE = os.path.join(BASE_DIR, "log_divergencias.txt")

# -------- EXECUÇÃO --------
MODO_BASELINE = "no"

# -------- REGRAS --------
LOWER_OUTLIER_MODE = 'iqr'

TOLERANCIA_MEDIA = 20
TOLERANCIA_STD = 30
TOLERANCIA_DISTINCT = 50

# ==================================================
# FUNÇÕES AUXILIARES
# ==================================================

def detectar_tipo_serie(serie):
    tipos = set()
    for val in serie.dropna().head(100):
        tipos.add(type(val).__name__)
    return sorted(list(tipos)) if tipos else "vazio"


def normalizar_serie_numerica(serie):
    serie_str = serie.astype(str).str.strip()
    serie_str = serie_str.str.replace(",", ".", regex=False)
    return pd.to_numeric(serie_str, errors='coerce')


def variacao_percentual(valor_antigo, valor_novo):
    if not valor_antigo:
        return 0
    return abs((valor_novo - valor_antigo) / valor_antigo) * 100

# ==================================================
# CARREGAMENTO
# ==================================================

def carregar_arquivo():
    if ORIGEM_ARQUIVO == 'local':
        caminho_completo = os.path.join(CAMINHO_ARQUIVO, NOME_ARQUIVO)
        return pd.ExcelFile(caminho_completo)

    elif ORIGEM_ARQUIVO == 'sharepoint':
        # Modo atual: via pasta sincronizada (OneDrive)
        caminho_completo = os.path.join(CAMINHO_ARQUIVO, NOME_ARQUIVO)
        return pd.ExcelFile(caminho_completo)

    elif ORIGEM_ARQUIVO == 'azure_blob':
        from azure.storage.blob import BlobServiceClient

        blob_service = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)
        blob_client = blob_service.get_blob_client(container=AZURE_CONTAINER_NAME, blob=AZURE_BLOB_NAME)

        blob_data = blob_client.download_blob().readall()
        return pd.ExcelFile(BytesIO(blob_data))

    else:
        raise Exception("Origem de arquivo inválida")

# ==================================================
# ANÁLISE
# ==================================================

def analisar_coluna_numerica(serie):
    resultado = {}

    serie_limpa = normalizar_serie_numerica(serie)

    total = len(serie)
    nulls = serie_limpa.isna().sum()
    resultado["percentual_null"] = round((nulls / total) * 100, 2) if total > 0 else 0

    serie_valida = serie_limpa.dropna()

    if len(serie_valida) > 0:
        q1 = serie_valida.quantile(0.25)
        q3 = serie_valida.quantile(0.75)
        iqr = q3 - q1

        limite_superior = q3 + 1.5 * iqr

        if LOWER_OUTLIER_MODE == 'non_negative':
            limite_inferior = 0
            outliers_inferior = serie_valida[serie_valida < 0]
        else:
            limite_inferior = q1 - 1.5 * iqr
            outliers_inferior = serie_valida[serie_valida < limite_inferior]

        outliers_superior = serie_valida[serie_valida > limite_superior]

        total_validos = len(serie_valida)

        resultado.update({
            "percentual_outliers_total": round(((len(outliers_inferior) + len(outliers_superior)) / total_validos) * 100, 2),
            "media": float(serie_valida.mean()),
            "std": float(serie_valida.std()) if pd.notna(serie_valida.std()) else 0.0,
            "min": float(serie_valida.min()),
            "max": float(serie_valida.max()),
            "distinct": int(serie_valida.nunique())
        })

    return resultado

# ==================================================
# IDENTIFICAR LAYOUT
# ==================================================

def extrair_layout(xls):
    layout = {}

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)

        estrutura = {}

        for col in df.columns:
            tipo = detectar_tipo_serie(df[col])
            info_coluna = {"tipo": tipo}

            if "int" in tipo or "float" in tipo:
                info_coluna["analise_numerica"] = analisar_coluna_numerica(df[col])

            estrutura[str(col)] = info_coluna

        layout[sheet] = {
            "linhas": len(df),
            "ordem_colunas": [str(c) for c in df.columns],
            "colunas": estrutura
        }

    return layout

# ==================================================
# BASELINE
# ==================================================

def salvar_layout(layout):
    with open(OUTPUT_LAYOUT_FILE, "w", encoding="utf-8") as f:
        json.dump(layout, f, indent=4, ensure_ascii=False)


def carregar_layout_salvo():
    if not os.path.exists(OUTPUT_LAYOUT_FILE):
        return None
    try:
        with open(OUTPUT_LAYOUT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return None

# ==================================================
# COMPARAÇÃO
# ==================================================

def comparar_layouts(layout_antigo, layout_novo):
    divergencias = []
    resumo = {"CRÍTICO": 0, "ALERTA": 0, "INFO": 0}

    abas_antigas = set(layout_antigo.keys())
    abas_novas = set(layout_novo.keys())

    for aba in abas_antigas - abas_novas:
        msg = f"Aba removida: '{aba}'"
        divergencias.append(f"[CRÍTICO] {msg}")
        resumo["CRÍTICO"] += 1

    for aba in abas_novas - abas_antigas:
        msg = f"Aba nova: '{aba}'"
        divergencias.append(f"[CRÍTICO] {msg}")
        resumo["CRÍTICO"] += 1

    for aba in abas_antigas & abas_novas:
        antigo = layout_antigo[aba]
        novo = layout_novo[aba]

        if antigo["ordem_colunas"] != novo["ordem_colunas"]:
            msg = f"Ordem de colunas alterada na aba '{aba}'"
            divergencias.append(f"[CRÍTICO] {msg}")
            resumo["CRÍTICO"] += 1

        cols_antigas = set(antigo["colunas"].keys())
        cols_novas = set(novo["colunas"].keys())

        for col in cols_antigas - cols_novas:
            msg = f"Coluna removida na aba '{aba}': '{col}'"
            divergencias.append(f"[CRÍTICO] {msg}")
            resumo["CRÍTICO"] += 1

        for col in cols_novas - cols_antigas:
            msg = f"Coluna nova na aba '{aba}': '{col}'"
            divergencias.append(f"[CRÍTICO] {msg}")
            resumo["CRÍTICO"] += 1

        for col in cols_antigas & cols_novas:
            a_old = antigo["colunas"][col].get("analise_numerica")
            a_new = novo["colunas"][col].get("analise_numerica")

            if a_old and a_new:
                if variacao_percentual(a_old.get("media"), a_new.get("media")) > TOLERANCIA_MEDIA:
                    divergencias.append(f"[CRÍTICO] Mudança de MÉDIA na aba '{aba}', coluna '{col}'")
                    resumo["CRÍTICO"] += 1

                if variacao_percentual(a_old.get("std"), a_new.get("std")) > TOLERANCIA_STD:
                    divergencias.append(f"[ALERTA] Mudança de STD na aba '{aba}', coluna '{col}'")
                    resumo["ALERTA"] += 1

                if variacao_percentual(a_old.get("distinct"), a_new.get("distinct")) > TOLERANCIA_DISTINCT:
                    divergencias.append(f"[INFO] Mudança de DISTINCT na aba '{aba}', coluna '{col}'")
                    resumo["INFO"] += 1

                if a_new["percentual_null"] > a_old["percentual_null"]:
                    divergencias.append(f"[ALERTA] Aumento de NULL na aba '{aba}', coluna '{col}'")
                    resumo["ALERTA"] += 1

                if a_new["percentual_outliers_total"] > a_old["percentual_outliers_total"]:
                    divergencias.append(f"[ALERTA] Aumento de OUTLIERS na aba '{aba}', coluna '{col}'")
                    resumo["ALERTA"] += 1

    return divergencias, resumo

# ==================================================
# SCORE
# ==================================================

def calcular_score(divergencias):
    penalidade = 0
    for d in divergencias:
        if "CRÍTICO" in d:
            penalidade += 20
        elif "ALERTA" in d:
            penalidade += 10
        else:
            penalidade += 5

    return max(0, 100 - penalidade)

# ==================================================
# LOG
# ==================================================

def registrar_log(divergencias, resumo, score):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write("\n" + "=" * 50 + "\n")
        f.write(f"Execução: {datetime.now()}\n")

        f.write("\nResumo:\n")
        for k, v in resumo.items():
            f.write(f"{k}: {v}\n")

        f.write(f"\nScore de Qualidade: {score}\n")

        f.write("\nDetalhes:\n")
        for d in divergencias:
            f.write(d + "\n")

# ==================================================
# MAIN
# ==================================================

def main():
    print("Carregando arquivo...")
    xls = carregar_arquivo()

    print("Extraindo layout...")
    layout_atual = extrair_layout(xls)

    if MODO_BASELINE.lower() == "yes":
        salvar_layout(layout_atual)
        print("Baseline criado.")
        return

    layout_salvo = carregar_layout_salvo()

    if layout_salvo is None:
        print("Execute primeiro com MODO_BASELINE = 'yes'")
        return

    print("Comparando...")
    divergencias, resumo = comparar_layouts(layout_salvo, layout_atual)

    score = calcular_score(divergencias)

    registrar_log(divergencias, resumo, score)

    print(f"Score de Qualidade: {score}")

    if divergencias:
        print("Divergências encontradas!")
    else:
        print("Tudo OK!")


if __name__ == "__main__":
    main()
