import os
import socket
import math
from datetime import datetime

import pandas as pd
import dash
from dash import dcc, html, Input, Output, State
import plotly.express as px
import plotly.graph_objects as go

# ================= CONFIG =================

ARQUIVO_VENDAS = r"C:\Users\wgdam\OneDrive\Desktop\Meus projetos\dados\BASE DE VENDAS.xlsx"
PORTA = 8050

AZUL = "#D31D1D"
LARANJA = "#B85426"
VERDE = "#2ecc71"  # para recomposição
FUNDO = "#FFFFFF"

# ================= VALIDAÇÃO =================

if not os.path.exists(ARQUIVO_VENDAS):
    raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_VENDAS}")

df = pd.read_excel(ARQUIVO_VENDAS)

# Garantir que as colunas numéricas sejam tratadas corretamente
df["Margem PDV"] = pd.to_numeric(df["Margem PDV"], errors="coerce").fillna(0)
df["R$ Real Venda"] = pd.to_numeric(df["R$ Real Venda"], errors="coerce").fillna(0)

# --- TRATAMENTO DAS COLUNAS SELL IN E SELL OUT ---
# Verifica se as colunas existem; se não, cria com zero
if "Sell In" in df.columns:
    df["Sell In"] = pd.to_numeric(df["Sell In"], errors="coerce").fillna(0)
else:
    df["Sell In"] = 0.0

if "Sell Out" in df.columns:
    df["Sell Out"] = pd.to_numeric(df["Sell Out"], errors="coerce").fillna(0)
else:
    df["Sell Out"] = 0.0

# Cria a coluna Recomposição como soma de Sell In e Sell Out
df["Recomposição"] = df["Sell In"] + df["Sell Out"]

# ================= FUNÇÕES =================

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_percentual(valor):
    return f"{valor:.1f}%"

def get_local_ip(porta):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return f"http://{ip}:{porta}"
    except:
        return f"http://localhost:{porta}"

def calcular_kpis(base):
    neg = base[base["Margem PDV"] < 0]

    impacto = neg["Margem PDV"].sum()
    venda_neg = neg["R$ Real Venda"].sum()
    percentual_negativo = (impacto / venda_neg * 100) if venda_neg != 0 else 0

    comprador_ofensor = (
        neg.groupby("Comprador")["Margem PDV"]
        .sum()
        .sort_values()
        .head(1)
    )

    nome_ofensor = comprador_ofensor.index[0] if len(comprador_ofensor) else "N/D"
    valor_ofensor = comprador_ofensor.iloc[0] if len(comprador_ofensor) else 0

    # Calcular recomposição total (opcional para um card)
    recomposicao_total = base["Recomposição"].sum()

    return impacto, venda_neg, percentual_negativo, nome_ofensor, valor_ofensor, recomposicao_total

# ================= APP =================

app = dash.Dash(
    __name__,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}]
)

app.title = "Margem Negativa"

data_execucao = datetime.now().strftime("%d/%m/%Y %H:%M")

card_style = {
    "backgroundColor": "white",
    "padding": "20px",
    "borderRadius": "12px",
    "boxShadow": "0 4px 12px rgba(0,0,0,0.08)",
    "flex": "1 1 220px",
    "minWidth": "200px",
    "textAlign": "center"
}

app.layout = html.Div([

    html.H1("Margem Negativa", style={"color": AZUL}),
    html.Div(f"Atualizado em {data_execucao}", style={"marginBottom": "20px"}),

    # Linha com KPIs e botão de limpar filtros
    html.Div([
        html.Div(id="kpis", style={"flex": "1"}),
        html.Button("Limpar Filtros", id="botao_limpar", n_clicks=0,
                    style={"backgroundColor": AZUL, "color": "white", "border": "none",
                           "padding": "10px 20px", "borderRadius": "8px", "cursor": "pointer",
                           "marginLeft": "20px", "alignSelf": "center"})
    ], style={"display": "flex", "alignItems": "center", "marginBottom": "20px"}),

    html.Hr(),

    html.H3("Impacto por Loja (ordenado)"),
    dcc.Graph(id="grafico_lojas", config={"responsive": True}),

    html.Hr(),

    html.H3("Ranking de Compradores Negativos"),
    dcc.Graph(id="grafico_compradores", config={"responsive": True}),

    html.Hr(),

    # Título dinâmico para a tabela de produtos
    html.Div(id="titulo_produtos", style={"fontSize": "18px", "fontWeight": "bold", "marginBottom": "10px"}),
    html.Div(id="tabela_produtos")

], style={
    "padding": "30px",
    "fontFamily": "Arial",
    "backgroundColor": FUNDO
})

# ================= CALLBACK PRINCIPAL (KPIs e Gráficos) =================

@app.callback(
    Output("kpis", "children"),
    Output("grafico_lojas", "figure"),
    Output("grafico_compradores", "figure"),
    Input("grafico_lojas", "clickData"),
    Input("grafico_compradores", "clickData"),
    Input("botao_limpar", "n_clicks")
)
def atualizar_dashboard(click_loja, click_comp, n_clicks):
    # Determina se deve limpar os filtros (baseado no botão)
    ctx = dash.callback_context
    trigger_id = ctx.triggered[0]["prop_id"].split(".")[0] if ctx.triggered else None

    base = df.copy()

    # Aplica filtros apenas se o botão não foi clicado
    if trigger_id != "botao_limpar":
        if click_loja:
            loja = click_loja["points"][0]["label"] if "label" in click_loja["points"][0] else click_loja["points"][0]["x"]
            base = base[base["Loja"] == loja]

        if click_comp:
            comprador = click_comp["points"][0]["y"]
            base = base[base["Comprador"] == comprador]

    impacto, venda_neg, perc_neg, nome_ofensor, valor_ofensor, recomposicao_total = calcular_kpis(base)

    # ===== KPIs =====
    # Agora com 5 cards (incluindo recomposição total)
    kpi_layout = html.Div([
        html.Div([
            html.Div("Impacto Total"),
            html.H2(formatar_moeda(impacto))
        ], style=card_style),

        html.Div([
            html.Div("Venda Negativa"),
            html.H2(formatar_moeda(venda_neg))
        ], style=card_style),

        html.Div([
            html.Div("% Margem Negativa"),
            html.H2(formatar_percentual(perc_neg))
        ], style=card_style),

        html.Div([
            html.Div("Maior Ofensor"),
            html.H3(nome_ofensor),
            html.Div(formatar_moeda(valor_ofensor))
        ], style=card_style),

        html.Div([
            html.Div("Recomposição Total"),
            html.H2(formatar_moeda(recomposicao_total))
        ], style={**card_style, "backgroundColor": "#e8f5e9"}),  # fundo verde claro

    ], style={
        "display": "flex",
        "flexWrap": "wrap",
        "gap": "20px"
    })

    # ===== GRÁFICO DE LOJAS =====
    resumo_lojas = (
        base.groupby("Loja")
        .agg({
            "Margem PDV": "sum",
            "R$ Real Venda": "sum",
            "Recomposição": "sum"  # adicionado para possível uso no hover
        })
        .reset_index()
        .sort_values("Margem PDV")
    )

    ofensor_por_loja = (
        base[base["Margem PDV"] < 0]
        .groupby(["Loja", "Comprador"])["Margem PDV"]
        .sum()
        .reset_index()
        .sort_values(["Loja", "Margem PDV"])
        .groupby("Loja")
        .first()
        .reset_index()[["Loja", "Comprador"]]
        .rename(columns={"Comprador": "Maior Ofensor"})
    )

    resumo_lojas = resumo_lojas.merge(ofensor_por_loja, on="Loja", how="left")
    resumo_lojas["Maior Ofensor"] = resumo_lojas["Maior Ofensor"].fillna("N/D")

    fig_lojas = px.bar(
        resumo_lojas,
        x="Margem PDV",
        y="Loja",
        orientation='h',
        text="Margem PDV",
        hover_data={
            "R$ Real Venda": True,
            "Maior Ofensor": True,
            "Recomposição": True
        },
        color="Margem PDV",
        color_continuous_scale="Reds",
        title="Impacto Negativo por Loja"
    )

    fig_lojas.update_traces(
        texttemplate='%{text:.2s}',
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Impacto: %{x:$,.0f}<br>Venda Negativa: %{customdata[0]:$,.0f}<br>Maior Ofensor: %{customdata[1]}<br>Recomposição: %{customdata[2]:$,.0f}<extra></extra>'
    )

    fig_lojas.update_layout(
        plot_bgcolor="white",
        paper_bgcolor=FUNDO,
        xaxis_title="Margem PDV (R$)",
        yaxis_title="",
        height=500,
        coloraxis_showscale=False
    )

    # ===== GRÁFICO DE COMPRADORES =====
    neg = base[base["Margem PDV"] < 0]
    compradores_df = (
        neg.groupby("Comprador")
        .agg({
            "Margem PDV": "sum",
            "R$ Real Venda": "sum",
            "Recomposição": "sum"  # adicionado para hover
        })
        .reset_index()
        .sort_values("Margem PDV")
        .head(10)
    )

    fig_compradores = px.bar(
        compradores_df,
        x="Margem PDV",
        y="Comprador",
        orientation='h',
        text="Margem PDV",
        hover_data={
            "R$ Real Venda": True,
            "Recomposição": True
        },
        color_discrete_sequence=[LARANJA],
        labels={"Margem PDV": "Impacto (R$)", "Comprador": "Comprador"},
        title="Top 10 Compradores com Maior Impacto Negativo"
    )

    fig_compradores.update_traces(
        texttemplate='%{text:.2s}',
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Impacto: %{x:$,.0f}<br>Venda Negativa: %{customdata[0]:$,.0f}<br>Recomposição: %{customdata[1]:$,.0f}<extra></extra>'
    )

    fig_compradores.update_layout(
        plot_bgcolor="white",
        paper_bgcolor=FUNDO,
        xaxis_title="Impacto (R$)",
        yaxis_title="",
        height=500
    )

    return kpi_layout, fig_lojas, fig_compradores

# ================= DRILL DOWN (Tabela de Produtos) =================

@app.callback(
    Output("titulo_produtos", "children"),
    Output("tabela_produtos", "children"),
    Input("grafico_compradores", "clickData"),
    State("grafico_lojas", "clickData"),
    State("botao_limpar", "n_clicks")
)
def detalhar_comprador(click_comp, click_loja, n_clicks):
    # Se não houver clique em comprador, limpa o título e a tabela
    if not click_comp:
        return "", "Clique em um comprador para expandir os produtos."

    comprador = click_comp["points"][0]["y"]
    base = df.copy()

    # Aplica filtro de loja se houver
    if click_loja:
        loja = click_loja["points"][0]["label"] if "label" in click_loja["points"][0] else click_loja["points"][0]["x"]
        base = base[base["Loja"] == loja]

    # Filtrar apenas produtos do comprador com margem negativa
    base = base[(base["Comprador"] == comprador) & (base["Margem PDV"] < 0)]

    if base.empty:
        return f"Produtos do comprador: {comprador}", "Nenhum produto com margem negativa para este comprador."

    # Agrupar por produto (e opcionalmente por loja se quiser maior granularidade)
    # Para manter o foco em produto, agrupamos apenas por produto.
    # Se quiser incluir loja, descomente a linha abaixo e ajuste o groupby.
    resumo = (
        base.groupby("Produto")  # poderia ser ["Loja", "Produto"] se quiser separar por loja
        .agg({
            "Margem PDV": "sum",
            "R$ Real Venda": "sum",
            "Recomposição": "sum"
        })
        .sort_values("Margem PDV")
        .reset_index()
    )

    # Formatar valores
    resumo["Margem PDV"] = resumo["Margem PDV"].apply(formatar_moeda)
    resumo["R$ Real Venda"] = resumo["R$ Real Venda"].apply(formatar_moeda)
    resumo["Recomposição"] = resumo["Recomposição"].apply(formatar_moeda)

    # Criar tabela HTML
    tabela = html.Table(
        [html.Tr([html.Th(col) for col in resumo.columns])] +
        [html.Tr([html.Td(resumo.iloc[i][col]) for col in resumo.columns])
         for i in range(len(resumo))],
        style={"width": "100%", "borderCollapse": "collapse", "marginTop": "10px",
               "border": "1px solid #ddd"}
    )

    titulo = f"Produtos do comprador: {comprador}"

    return titulo, tabela

# ================= EXECUÇÃO =================

if __name__ == "__main__":
    print("Dashboard iniciado")
    print("Local:", f"http://localhost:{PORTA}")
    print("Rede Wi-Fi:", get_local_ip(PORTA))
    app.run(host="0.0.0.0", port=PORTA)