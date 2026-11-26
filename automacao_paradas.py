import pandas as pd
from dashboard_powerbi import gerar_dashboard

def executar_automacao():
    print("📥 Lendo dados...")
    
    # 1. Ler dados
    df = pd.read_excel("Maquinas_Paradas.xlsx", engine="openpyxl")

    # 2. Validar colunas
    colunas_necessarias = ["inicio", "fim", "maquina"]
    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória não encontrada: {col}")

    # 3. Converter datas
    df["inicio"] = pd.to_datetime(df["inicio"], errors="coerce")
    df["fim"] = pd.to_datetime(df["fim"], errors="coerce")

    # 4. Remover inválidos
    df = df.dropna(subset=["inicio", "fim"])

    # 5. Calcular duração
    df["duracao_min"] = (df["fim"] - df["inicio"]).dt.total_seconds() / 60
    df["duracao_min"] = df["duracao_min"].clip(lower=0)

    # 6. Criar colunas úteis
    df["data"] = df["inicio"].dt.date
    df["mes"] = df["inicio"].dt.month
    df["ano"] = df["inicio"].dt.year
    df["hora_inicio"] = df["inicio"].dt.hour

    # 7. Agrupar
    resumo_maquina = df.groupby("maquina")["duracao_min"].sum().reset_index()

    # 8. Salvar arquivos tratados
    df.to_excel("paradas_tratadas.xlsx", index=False, engine="openpyxl")
    resumo_maquina.to_excel("resumo_maquina.xlsx", index=False, engine="openpyxl")

    print(f"✅ Dados tratados! Linhas processadas: {len(df)}")

    # 9. Gerar dashboard automático
    gerar_dashboard()