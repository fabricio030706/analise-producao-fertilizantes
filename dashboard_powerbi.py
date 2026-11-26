import pandas as pd
import json
import os

def gerar_dashboard():
    # =========================
    # 📥 CARREGAR DADOS
    # =========================
    df = pd.read_excel("paradas_tratadas.xlsx", engine="openpyxl")
    resumo = pd.read_excel("resumo_maquina.xlsx", engine="openpyxl")

    # limpar nomes
    df.columns = df.columns.str.strip()
    resumo.columns = resumo.columns.str.strip()

    # lista máquinas
    maquinas = df["maquina"].unique()

    # converter para json
    df_json = df.to_json(orient="records")
    resumo_json = resumo.to_json(orient="records")

    # =========================
    # 📄 HTML
    # =========================
    html = f"""
<html>
<head>
    <title>Dashboard de Paradas</title>

    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>

    <style>
        body {{
            font-family: 'Segoe UI';
            background: linear-gradient(135deg, #eef2f7, #f8fafc);
            margin: 0;
            padding: 30px;
        }}

        h1 {{
            text-align: center;
            margin-bottom: 20px;
        }}

        .filtro {{
            text-align: center;
            margin-bottom: 30px;
        }}

        select {{
            padding: 10px;
            font-size: 16px;
            border-radius: 8px;
        }}

        .grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
            max-width: 1300px;
            margin: auto;
        }}

        .card {{
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 6px 15px rgba(0,0,0,0.1);
        }}

        .full {{
            grid-column: span 2;
        }}
    </style>
</head>

<body>

<h1>📊 Dashboard de Paradas</h1>

<div class="filtro">
    <label><b>Filtrar Máquina:</b></label>
    <select id="filtroMaquina" onchange="atualizar()">
        <option value="Todos">Todos</option>
        {''.join([f'<option value="{m}">{m}</option>' for m in maquinas])}
    </select>
</div>

<div class="grid">
    <div class="card"><div id="grafico1"></div></div>
    <div class="card"><div id="grafico2"></div></div>
    <div class="card full"><div id="grafico3"></div></div>
</div>

<script>
var df = {df_json};
var resumo = {resumo_json};

function filtrarDados(maquina) {{
    if (maquina === "Todos") return df;
    return df.filter(d => d.maquina === maquina);
}}

function filtrarResumo(maquina) {{
    if (maquina === "Todos") return resumo;
    return resumo.filter(d => d.maquina === maquina);
}}

function atualizar() {{
    var maquina = document.getElementById("filtroMaquina").value;

    var dfFiltrado = filtrarDados(maquina);
    var resumoFiltrado = filtrarResumo(maquina);

    // 📊 GRAFICO 1
    var maquinas = resumoFiltrado.map(d => d.maquina);
    var valores = resumoFiltrado.map(d => d.duracao_min);

    Plotly.newPlot("grafico1", [{{
        x: maquinas,
        y: valores,
        type: 'bar',
        marker: {{
            color: valores,
            colorscale: 'Inferno'
        }}
    }}], {{title: "Paradas por Máquina"}});

    // 📊 GRAFICO 2
    let horas = {{}};
    dfFiltrado.forEach(d => {{
        if (!horas[d.hora_inicio]) horas[d.hora_inicio] = 0;
        horas[d.hora_inicio] += d.duracao_min;
    }});

    Plotly.newPlot("grafico2", [{{
        x: Object.keys(horas),
        y: Object.values(horas),
        type: 'bar',
        marker: {{
            color: Object.values(horas),
            colorscale: 'Inferno'
        }}
    }}], {{title: "Paradas por Hora"}});

    // 📊 GRAFICO 3 (ranking)
    Plotly.newPlot("grafico3", [{{
        x: valores,
        y: maquinas,
        type: 'bar',
        orientation: 'h',
        marker: {{
            color: valores,
            colorscale: 'Inferno'
        }}
    }}], {{title: "Ranking de Máquinas"}});
}}

atualizar();
</script>

</body>
</html>
"""

    with open("dashboard_paradas.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ Dashboard gerado sem erros!")

    os.system("start dashboard_paradas.html")


if __name__ == "__main__":
    gerar_dashboard()
