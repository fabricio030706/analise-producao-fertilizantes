import pandas as pd
import os
import glob
import re
import unicodedata
import math
from pathlib import Path

# =========================================================
# CONFIGURACAO
# =========================================================

PASTA_RELATORIOS = [
    r"S:\Relatorio_Dosagem\Mistura",
    r"S:\Relatorio_Dosagem\Backup\Mistura",
]

# Template definido abaixo na versao separada
ARQUIVO_DASHBOARD = "dashboard_fertipar_consolidado.html"
ARQUIVO_CONSOLIDADO = "fertipar_consolidado.xlsx"

# Pasta onde esta este arquivo ProducaoMistura.py
# Isso evita erro quando voce roda o comando a partir de outra pasta.
BASE_DIR = Path(__file__).resolve().parent

# =========================================================
# FUNCOES AUXILIARES
# =========================================================


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower()
    texto = texto.replace("\n", " ").replace("\r", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpar_numero(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return None
        return float(valor)

    texto = str(valor).strip()
    if texto == "":
        return None

    texto = texto.replace("ton", "").replace("TON", "").replace(" ", "")
    match = re.search(r"-?\d+[\d.,]*", texto)
    if not match:
        return None

    numero = match.group(0)

    # Formato brasileiro: 1.234,56 -> 1234.56
    if "," in numero and "." in numero:
        numero = numero.replace(".", "").replace(",", ".")
    elif "," in numero:
        numero = numero.replace(",", ".")

    try:
        return float(numero)
    except Exception:
        return None


def converter_tempo_min(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return None
        # Se vier como serial de tempo do Excel: 0.5 = 12 horas = 720 min.
        # Se vier como numero grande, assume que ja esta em minutos.
        if 0 <= float(valor) < 1:
            return float(valor) * 24 * 60
        return float(valor)

    texto = str(valor).strip()
    if texto == "":
        return None

    convertido = pd.to_timedelta(texto, errors="coerce")
    if not pd.isna(convertido):
        return convertido.total_seconds() / 60

    numero = limpar_numero(texto)
    return numero


def extrair_data_nome_arquivo(nome_arquivo):
    match = re.search(r"(\d{2}-\d{2}-\d{4})", nome_arquivo)
    if not match:
        return pd.NaT
    return pd.to_datetime(match.group(1), format="%d-%m-%Y", errors="coerce")


def abrir_arquivo(caminho):
    try:
        os.startfile(str(caminho))
    except Exception:
        print(f"Arquivo gerado em: {caminho}")


def localizar_coluna(raw, palavras, evitar=None):
    if evitar is None:
        evitar = []

    palavras_norm = [normalizar_texto(p) for p in palavras]
    evitar_norm = [normalizar_texto(e) for e in evitar]

    for i in range(raw.shape[0]):
        for j in range(raw.shape[1]):
            texto = normalizar_texto(raw.iat[i, j])
            if texto == "":
                continue
            if any(e in texto for e in evitar_norm):
                continue
            if any(p in texto for p in palavras_norm):
                return i, j, raw.iat[i, j]

    return None, None, None


def localizar_peso_total(raw):
    """
    Procura PESO TOTAL no relatorio.
    Se encontrar apenas 1.0, ignora porque geralmente e pagina/unidade
    em relatorio sem producao.
    """
    for i in range(raw.shape[0]):
        for j in range(raw.shape[1]):
            texto = normalizar_texto(raw.iat[i, j])
            if "peso total" not in texto:
                continue

            # Procura valor na mesma linha, a direita.
            for c in range(j + 1, raw.shape[1]):
                valor = limpar_numero(raw.iat[i, c])
                if valor is not None and valor > 1:
                    return valor

            # Procura valor nas proximas linhas.
            for r in range(i + 1, min(i + 4, raw.shape[0])):
                for c in range(raw.shape[1]):
                    valor = limpar_numero(raw.iat[r, c])
                    if valor is not None and valor > 1:
                        return valor

    return None


def criar_linha_zero(data_arquivo, nome_arquivo):
    return pd.DataFrame(
        [
            {
                "data_arquivo": data_arquivo,
                "arquivo_origem": nome_arquivo,
                "ordem": "",
                "cliente": "SEM PRODUCAO",
                "formula": "",
                "placa": "",
                "usuario": "",
                "peso": 0,
                "tempo_min": 0,
                "ton_hora": 0,
                "tipo_registro": "SEM_PRODUCAO",
                "registro_valido": 0,
            }
        ]
    )


# =========================================================
# PROCESSAR CADA ARQUIVO
# =========================================================


def processar_arquivo_fertipar(arquivo):
    nome_arquivo = os.path.basename(arquivo)
    data_arquivo = extrair_data_nome_arquivo(nome_arquivo)

    resultado_log = {
        "arquivo": nome_arquivo,
        "data": data_arquivo,
        "status": "",
        "linhas": 0,
        "peso_total_lido": 0,
    }

    try:
        # Relatorios antigos .XLS precisam de xlrd.
        engine = "xlrd" if arquivo.lower().endswith(".xls") else "openpyxl"
        raw = pd.read_excel(arquivo, engine=engine, header=None)
        raw = raw.dropna(how="all")
        raw = raw.dropna(axis=1, how="all")

        if pd.isna(data_arquivo):
            resultado_log["status"] = "Data nao encontrada no nome"
            return pd.DataFrame(), resultado_log

        if raw.empty:
            resultado_log["status"] = "Arquivo vazio / sem producao"
            return criar_linha_zero(data_arquivo, nome_arquivo), resultado_log

        linha_peso, col_peso, _ = localizar_coluna(raw, ["peso"], evitar=["peso total"])
        linha_tempo, col_tempo, _ = localizar_coluna(raw, ["tempo"])
        linha_ton, col_ton, _ = localizar_coluna(raw, ["ton/hora", "ton hora", "ton"])
        linha_ordem, col_ordem, _ = localizar_coluna(raw, ["ordem"])
        linha_cliente, col_cliente, _ = localizar_coluna(raw, ["cliente"])
        linha_formula, col_formula, _ = localizar_coluna(raw, ["formula", "fórmula"])
        linha_placa, col_placa, _ = localizar_coluna(raw, ["placa"])
        linha_usuario, col_usuario, _ = localizar_coluna(raw, ["usuario", "usuário"])

        registros = []

        if col_peso is not None:
            linha_inicio_dados = linha_peso + 1

            for i in range(linha_inicio_dados, raw.shape[0]):
                linha_completa = " ".join(
                    normalizar_texto(raw.iat[i, j]) for j in range(raw.shape[1])
                )

                if "peso total" in linha_completa:
                    continue
                if "pagina" in linha_completa or "página" in linha_completa:
                    continue
                if "relatorio" in linha_completa or "relatório" in linha_completa:
                    continue
                if "filtros" in linha_completa:
                    continue

                peso = limpar_numero(raw.iat[i, col_peso])
                if peso is None or peso <= 0:
                    continue

                tempo_min = (
                    converter_tempo_min(raw.iat[i, col_tempo])
                    if col_tempo is not None
                    else None
                )
                ton_hora = (
                    limpar_numero(raw.iat[i, col_ton]) if col_ton is not None else None
                )

                ordem = raw.iat[i, col_ordem] if col_ordem is not None else ""
                cliente = raw.iat[i, col_cliente] if col_cliente is not None else ""
                formula = raw.iat[i, col_formula] if col_formula is not None else ""
                placa = raw.iat[i, col_placa] if col_placa is not None else ""
                usuario = raw.iat[i, col_usuario] if col_usuario is not None else ""

                ordem_txt = normalizar_texto(ordem)
                cliente_txt = normalizar_texto(cliente)
                formula_txt = normalizar_texto(formula)
                placa_txt = normalizar_texto(placa)
                usuario_txt = normalizar_texto(usuario)

                tempo_zero = tempo_min is None or tempo_min == 0
                ton_zero = ton_hora is None or ton_hora == 0

                # Remove falso 1.0 de dias sem producao.
                if (
                    peso <= 1
                    and tempo_zero
                    and ton_zero
                    and ordem_txt == ""
                    and cliente_txt == ""
                    and formula_txt == ""
                    and placa_txt == ""
                    and usuario_txt == ""
                ):
                    continue

                registros.append(
                    {
                        "data_arquivo": data_arquivo,
                        "arquivo_origem": nome_arquivo,
                        "ordem": ordem,
                        "cliente": cliente,
                        "formula": formula,
                        "placa": placa,
                        "usuario": usuario,
                        "peso": peso,
                        "tempo_min": tempo_min if tempo_min is not None else 0,
                        "ton_hora": ton_hora if ton_hora is not None else 0,
                        "tipo_registro": "DETALHE",
                        "registro_valido": 1,
                    }
                )

        df = pd.DataFrame(registros)

        if df.empty:
            peso_total = localizar_peso_total(raw)
            if peso_total is not None and peso_total > 1:
                df = pd.DataFrame(
                    [
                        {
                            "data_arquivo": data_arquivo,
                            "arquivo_origem": nome_arquivo,
                            "ordem": "",
                            "cliente": "TOTAL DO ARQUIVO",
                            "formula": "",
                            "placa": "",
                            "usuario": "",
                            "peso": peso_total,
                            "tempo_min": 0,
                            "ton_hora": 0,
                            "tipo_registro": "TOTAL_ARQUIVO",
                            "registro_valido": 1,
                        }
                    ]
                )
                resultado_log["status"] = "Usou PESO TOTAL do arquivo"
                resultado_log["linhas"] = 1
                resultado_log["peso_total_lido"] = peso_total
                return df, resultado_log

            resultado_log["status"] = "Sem producao no dia"
            resultado_log["linhas"] = 0
            resultado_log["peso_total_lido"] = 0
            return criar_linha_zero(data_arquivo, nome_arquivo), resultado_log

        resultado_log["status"] = "Detalhes lidos com sucesso"
        resultado_log["linhas"] = len(df)
        resultado_log["peso_total_lido"] = df["peso"].sum()
        return df, resultado_log

    except Exception as e:
        resultado_log["status"] = f"Erro: {e}"
        resultado_log["linhas"] = 0
        resultado_log["peso_total_lido"] = 0

        if not pd.isna(data_arquivo):
            return criar_linha_zero(data_arquivo, nome_arquivo), resultado_log

        return pd.DataFrame(), resultado_log


# =========================================================
# DASHBOARD PRINCIPAL - ARQUIVOS SEPARADOS
# =========================================================

import json

# Nesta versao, o Python fica responsavel apenas por:
# 1) ler os relatorios;
# 2) consolidar a base;
# 3) gerar o Excel;
# 4) injetar os dados JSON no HTML.
# Toda a parte visual/graficos/filtros fica no arquivo dashboard_fertipar.js.

ARQUIVO_TEMPLATE = "dashboard_fertipar.html"
ARQUIVO_JS = "dashboard_fertipar.js"
ARQUIVO_JS_SAIDA = Path(ARQUIVO_DASHBOARD).with_suffix(".js").name


def _texto_seguro(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    return "" if texto.lower() == "nan" else texto


def _formatar_data_br(valor):
    data = pd.to_datetime(valor, errors="coerce")
    return "" if pd.isna(data) else data.strftime("%d/%m/%Y")


def _formatar_data_iso(valor):
    data = pd.to_datetime(valor, errors="coerce")
    return "" if pd.isna(data) else data.strftime("%Y-%m-%d")


def _num(valor, casas=3):
    try:
        if pd.isna(valor):
            return 0
        return round(float(valor), casas)
    except Exception:
        return 0


def preparar_dados_dashboard(df_final):
    df = df_final.copy()
    df["data_arquivo"] = pd.to_datetime(df["data_arquivo"], errors="coerce")

    for col in ["peso", "tempo_min", "ton_hora", "registro_valido"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    registros = []
    for _, row in df.iterrows():
        registros.append(
            {
                "data_br": _formatar_data_br(row.get("data_arquivo")),
                "data_iso": _formatar_data_iso(row.get("data_arquivo")),
                "arquivo_origem": _texto_seguro(row.get("arquivo_origem")),
                "ordem": _texto_seguro(row.get("ordem")),
                "cliente": _texto_seguro(row.get("cliente")),
                "formula": _texto_seguro(row.get("formula")),
                "placa": _texto_seguro(row.get("placa")),
                "usuario": _texto_seguro(row.get("usuario")),
                "peso": _num(row.get("peso")),
                "tempo_min": _num(row.get("tempo_min")),
                "ton_hora": _num(row.get("ton_hora")),
                "tipo_registro": _texto_seguro(row.get("tipo_registro")),
                "registro_valido": int(_num(row.get("registro_valido"), 0)),
            }
        )
    return registros


def preparar_log_dashboard(logs):
    saida = []
    for item in logs:
        saida.append(
            {
                "arquivo": _texto_seguro(item.get("arquivo")),
                "data": _formatar_data_br(item.get("data")),
                "status": _texto_seguro(item.get("status")),
                "linhas": int(_num(item.get("linhas"), 0)),
                "peso_total_lido": _num(item.get("peso_total_lido")),
            }
        )
    return saida


def executar_dashboard_fertipar():
    arquivos = []
    for pasta in PASTA_RELATORIOS:
        for ext in ("*.xls", "*.XLS", "*.xlsx", "*.XLSX"):
            arquivos.extend(glob.glob(os.path.join(pasta, ext)))

    arquivos = sorted(set(arquivos))

    if not arquivos:
        print("Nenhum arquivo encontrado nas pastas informadas.")
        return

    lista_df = []
    logs = []

    for arquivo in arquivos:
        df, log = processar_arquivo_fertipar(arquivo)
        logs.append(log)
        if not df.empty:
            lista_df.append(df)

        data_log = log.get("data")
        data_txt = (
            data_log.strftime("%d/%m/%Y") if not pd.isna(data_log) else "sem data"
        )
        print(
            f"{data_txt} | {log['status']} | Registros: {log['linhas']} | Peso: {log['peso_total_lido']:.2f}"
        )

    if not lista_df:
        print("Nenhum dado valido foi encontrado.")
        return

    df_final = pd.concat(lista_df, ignore_index=True)
    df_final["data_arquivo"] = pd.to_datetime(df_final["data_arquivo"], errors="coerce")
    df_final["peso"] = pd.to_numeric(df_final["peso"], errors="coerce").fillna(0)
    df_final["tempo_min"] = pd.to_numeric(
        df_final["tempo_min"], errors="coerce"
    ).fillna(0)
    df_final["ton_hora"] = pd.to_numeric(df_final["ton_hora"], errors="coerce").fillna(
        0
    )
    df_final["registro_valido"] = (
        pd.to_numeric(df_final["registro_valido"], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    caminho_excel = BASE_DIR / ARQUIVO_CONSOLIDADO
    log_df = pd.DataFrame(logs)
    por_dia = (
        df_final.groupby("data_arquivo", dropna=False)
        .agg(
            peso=("peso", "sum"),
            registros=("registro_valido", "sum"),
            tempo_medio_min=("tempo_min", "mean"),
            produtividade_media=("ton_hora", "mean"),
        )
        .sort_index()
        .fillna(0)
    )

    with pd.ExcelWriter(caminho_excel, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Base Consolidada", index=False)
        por_dia.to_excel(writer, sheet_name="Producao por Dia")
        log_df.to_excel(writer, sheet_name="Log Arquivos", index=False)

    caminho_template = BASE_DIR / ARQUIVO_TEMPLATE
    caminho_js = BASE_DIR / ARQUIVO_JS
    caminho_saida = BASE_DIR / ARQUIVO_DASHBOARD
    caminho_js_saida = BASE_DIR / ARQUIVO_JS_SAIDA

    if not caminho_template.exists():
        raise FileNotFoundError(f"Template nao encontrado: {caminho_template}")
    if not caminho_js.exists():
        raise FileNotFoundError(f"JavaScript nao encontrado: {caminho_js}")

    html = caminho_template.read_text(encoding="utf-8")
    html = html.replace("%%PASTA_RELATORIOS%%", str(PASTA_RELATORIOS))
    html = html.replace(
        "%%DATA_ATUALIZACAO%%", pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    )
    html = html.replace(
        "%%DADOS_JSON%%",
        json.dumps(
            preparar_dados_dashboard(df_final),
            ensure_ascii=False,
            separators=(",", ":"),
        ),
    )
    html = html.replace(
        "%%LOG_JSON%%",
        json.dumps(
            preparar_log_dashboard(logs), ensure_ascii=False, separators=(",", ":")
        ),
    )
    html = html.replace("%%ARQUIVO_JS%%", ARQUIVO_JS_SAIDA)

    # Gera tambem o arquivo JavaScript de saida. Assim, ao rodar somente este Python,
    # ele atualiza o dashboard HTML e o JS correspondente juntos.
    caminho_js_saida.write_text(
        caminho_js.read_text(encoding="utf-8"), encoding="utf-8"
    )
    caminho_saida.write_text(html, encoding="utf-8")

    print("\nDashboard consolidado gerado com sucesso!")
    print(f"HTML: {caminho_saida}")
    print(f"Excel conferencia: {caminho_excel}")
    print(f"JavaScript gerado: {caminho_js_saida}")

    abrir_arquivo(caminho_saida)


if __name__ == "__main__":
    executar_dashboard_fertipar()
