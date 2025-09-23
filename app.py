import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from datetime import timedelta
import numpy as np

# ===================== CARREGAMENTO E PREPARO =====================
ARQUIVO = "Linha do tempo.xlsx"
SHEET   = "Plan1"

df = pd.read_excel(ARQUIVO, sheet_name=SHEET)

# Equipamento formatado (só visual; NÃO usamos para filtrar)
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]

# Tipo macro
def classifica_tipo(row):
    desc = str(row["Descrição da Operação"]).strip().upper()
    grupo = str(row["Descrição do Grupo da Operação"]).strip().upper()
    if desc == "DESLOCAMENTO":
        return "Deslocamento"
    if desc == "MANOBRA":
        return "Manobra"
    if grupo == "PRODUTIVA":
        return "Produtiva"
    if grupo == "IMPRODUTIVA":
        return "Improdutiva"
    return "Outro"

df["Tipo"] = df.apply(classifica_tipo, axis=1)

# Parsing
df["Hora Inicial"]   = pd.to_datetime(df["Hora Inicial"], format="%H:%M:%S", errors="coerce").dt.time
df["Hora Final"]     = pd.to_datetime(df["Hora Final"],   format="%H:%M:%S", errors="coerce").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Nome", "Hora Inicial", "Hora Final", "Data Hora Local"]).copy()

# Instantes absolutos
df["Inicio"] = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Inicial']}"), axis=1)
df["Fim"]    = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Final']}"), axis=1)

# Cruza meia-noite? soma 1 dia no Fim
mask_cross = df["Fim"] < df["Inicio"]
df.loc[mask_cross, "Fim"] = df.loc[mask_cross, "Fim"] + pd.Timedelta(days=1)

# Mapeamento de Parada (negócio)
def classifica_tipo_parada(row):
    grupo = str(row["Descrição do Grupo da Operação"]).strip().upper()
    desc  = str(row["Descrição da Operação"]).strip().upper()
    gerenciaveis = {"AGUARDANDO COMBUSTIVEL","AGUARDANDO ORDENS","AGUARDANDO MOVIMENTACAO PIVO","FALTA DE INSUMOS"}
    essenciais   = {"REFEICAO","BANHEIRO"}
    mecanicas    = {"AGUARDANDO MECANICO","BORRACHARIA","EXCESSO DE TEMPERATURA DO MOTOR","IMPLEMENTO QUEBRADO",
                    "MANUTENCAO ELETRICA","MANUTENCAO MECANICA","TRATOR QUEBRADO","SEM SINAL GPS"}
    if grupo == "PRODUTIVA":  return "Efetivo"
    if grupo == "IMPRODUTIVA":
        if desc in gerenciaveis: return "Parada Gerenciável"
        if desc in mecanicas:    return "Parada Mecânica"
        if desc in essenciais:   return "Parada Essencial"
        if desc == "OUTROS":     return "Outros"
        return "Parada Improdutiva"
    if desc == "DESLOCAMENTO":  return "Deslocamento"
    if desc == "MANOBRA":       return "Manobra"
    return "Outro"

df["Tipo Parada"] = df.apply(classifica_tipo_parada, axis=1)

# Agrupar blocos contíguos da MESMA operação e MESMO equipamento (gap <= 2 min)
def agrupar_paradas(df_filtrado):
    if df_filtrado.empty:
        return pd.DataFrame(columns=["Nome","Inicio","Fim","Descrição da Operação","Duracao Min","Tipo Parada","Equipamento"])
    d = df_filtrado.sort_values("Inicio").reset_index(drop=True)
    out = []; i = 0
    while i < len(d):
        a = d.loc[i]
        ini, fim = a["Inicio"], a["Fim"]
        op, nome, tipo, equip = a["Descrição da Operação"], a["Nome"], a["Tipo Parada"], a["Equipamento"]
        j = i + 1
        while j < len(d):
            p = d.loc[j]
            gap = (p["Inicio"] - fim).total_seconds()/60.0
            if (p["Descrição da Operação"] == op) and (p["Equipamento"] == equip) and gap <= 2:
                fim = max(fim, p["Fim"]); j += 1
            else:
                break
        out.append({
            "Nome": nome, "Inicio": ini, "Fim": fim,
            "Descrição da Operação": op,
            "Duracao Min": max(0.0,(fim-ini).total_seconds()/60.0),
            "Tipo Parada": tipo,
            "Equipamento": equip
        })
        i = j
    return pd.DataFrame(out)

# Colapsar por equipamento (faixas translúcidas no fundo)
def blocos_equipamento(dff_raw):
    d = dff_raw.sort_values("Inicio").reset_index(drop=True)
    out = []; i = 0
    while i < len(d):
        a = d.loc[i]
        equip, ini, fim = a["Equipamento"], a["Inicio"], a["Fim"]
        j = i + 1
        while j < len(d):
            p = d.loc[j]
            gap = (p["Inicio"] - fim).total_seconds()/60.0
            if (p["Equipamento"] == equip) and gap <= 2:
                fim = max(fim, p["Fim"]); j += 1
            else:
                break
        out.append({"Equipamento": equip, "Inicio": ini, "Fim": fim})
        i = j
    return pd.DataFrame(out)

# Divisores de dia (linhas 00:00 + rótulo "Dia dd/mm")
def add_divisores_de_dia(fig, tmin, tmax):
    if pd.isna(tmin) or pd.isna(tmax):
        return
    start_day = pd.to_datetime(tmin.date())
    end_day   = pd.to_datetime(tmax.date()) + pd.Timedelta(days=1)
    cur = start_day
    while cur <= end_day:
        fig.add_vline(x=cur, line_width=1, line_dash="dot", line_color="#9aa0a6")
        mid = cur + pd.Timedelta(hours=12)
        fig.add_annotation(x=mid, y=1.08, yref="paper",
                           text=cur.strftime("Dia %d/%m"),
                           showarrow=False, font=dict(size=11, color="#e0e0e0"))
        cur += pd.Timedelta(days=1)

# Calcula janela visível (data selecionada OU pan/zoom do usuário)
def janela_visivel(data_str, relayoutData, df_base):
    # fallback = dia escolhido (00:00 → +24h) ou primeiro dia do dado
    if data_str:
        base_day = pd.to_datetime(data_str).normalize()
    else:
        base_day = pd.to_datetime(df_base["Inicio"].min()).normalize()
    x0 = base_day
    x1 = x0 + pd.Timedelta(days=1)

    rd = relayoutData or {}
    try:
        r0 = rd.get("xaxis.range[0]"); r1 = rd.get("xaxis.range[1]")
        if r0 and r1:
            x0 = pd.to_datetime(r0); x1 = pd.to_datetime(r1)
    except Exception:
        pass
    return x0, x1

# ===================== APP =====================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

# Valores iniciais
primeiro_nome = sorted(df["Nome"].dropna().unique())[0]
primeiras_datas = sorted(df[df["Nome"] == primeiro_nome]["Data Hora Local"].dt.date.unique())
data_padrao = str(primeiras_datas[-2]) if len(primeiras_datas) >= 2 else str(primeiras_datas[-1])

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),

        # Filtros (só Operador e Data inicial)
        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(
                    id="operador-dropdown",
                    options=[{"label": n, "value": n} for n in sorted(df["Nome"].dropna().unique())],
                    value=primeiro_nome,
                    placeholder="Selecione um Operador"
                ), md=6),
                dbc.Col(dcc.Dropdown(
                    id="data-dropdown",
                    options=[{"label": str(d), "value": str(d)} for d in primeiras_datas],
                    value=data_padrao,
                    placeholder="Data inicial (00:00 → +24h)"
                ), md=6),
            ], align="center")
        ]), className="mb-4"),

        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-3"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "600px"}))),
    ], fluid=False)
])

# Atualiza opções de data quando troca o operador
@app.callback(
    Output("data-dropdown", "options"),
    Output("data-dropdown", "value"),
    Input("operador-dropdown", "value"),
)
def atualizar_datas(operador):
    datas = sorted(df[df["Nome"] == operador]["Data Hora Local"].dt.date.unique())
    opts = [{"label": str(d), "value": str(d)} for d in datas]
    val = str(datas[-2]) if len(datas) >= 2 else (str(datas[-1]) if len(datas) else None)
    return opts, val

# ===================== CALLBACK PRINCIPAL =====================
@app.callback(
    Output("grafico-linha-tempo", "figure"),
    Output("stats-div", "children"),
    Input("operador-dropdown", "value"),
    Input("data-dropdown", "value"),
    Input("grafico-linha-tempo", "relayoutData"),
)
def atualizar_grafico(operador, data_str, relayoutData):
    if not operador:
        return {}, html.Div("Selecione um operador.", className="text-center text-muted p-4")

    base = df[df["Nome"] == operador].copy()
    if base.empty:
        return {}, html.Div("Nenhum dado encontrado.", className="text-center text-muted p-4")

    # Agrupa e remove operações de fim de expediente
    dff = agrupar_paradas(base)
    exp_set = {"FINAL DE EXPEDIENTE", "FIM DE EXPEDIENTE"}
    dff = dff[~dff["Descrição da Operação"].str.upper().str.strip().isin(exp_set)]
    if dff.empty:
        return {}, html.Div("Nenhum dado útil para visualizar.", className="text-center text-muted p-4")

    # Janela visível = data escolhida OU pan/zoom atual
    x0, x1 = janela_visivel(data_str, relayoutData, dff)

    # Stats: recorte pelo que está VISÍVEL (batelada)
    dff["Inicio_clip"] = dff["Inicio"].clip(lower=x0)
    dff["Fim_clip"]    = dff["Fim"].clip(upper=x1)
    dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds() / 60.0
    win = dff[dff["Duracao Min Clip"] > 0].copy()

    if win.empty:
        stats_html = html.Div("Sem atividade nessa janela.", className="text-center text-muted p-3")
    else:
        ini_batelada = win["Inicio_clip"].min().strftime("%d/%m %H:%M")
        fim_batelada = win["Fim_clip"].max().strftime("%d/%m %H:%M")
        total_horas  = win["Duracao Min Clip"].sum() / 60.0

        def soma_h(tipo):  # horas por tipo na janela
            return win.loc[win["Tipo Parada"] == tipo, "Duracao Min Clip"].sum() / 60.0

        def card(t, v, c):
            return dbc.Col(dbc.Card(dbc.CardBody([
                html.H4(v, style={"color": c, "fontWeight": "bold"}),
                html.P(t, className="text-muted")
            ]), className="text-center shadow-sm"), md=2, className="mb-2")

        stats_html = dbc.Row([
            card("Início da batelada", ini_batelada, "#6c757d"),
            card("Fim da batelada",    fim_batelada, "#6c757d"),
            card("Total (janela)",     f"{total_horas:.2f}h", "#343a40"),
            card("Efetivo (janela)",   f"{soma_h('Efetivo'):.2f}h", "#046414"),
            card("Gerenciável",        f"{soma_h('Parada Gerenciável'):.2f}h", "#B26B00"),
            card("Mecânica",           f"{soma_h('Parada Mecânica'):.2f}h", "#A52657"),
        ], justify="center")

    # Gráfico com TODA a linha do tempo (pan infinito), mas inicia enquadrado na janela atual
    dff["Resumo"] = dff.apply(lambda r: (
        f"Operador: {r['Nome']}<br>"
        f"Equipamento: {r['Equipamento']}<br>"
        f"Tipo: {r['Tipo Parada']}<br>"
        f"Operação: {r['Descrição da Operação']}<br>"
        f"Início: {r['Inicio'].strftime('%d/%m %H:%M')}<br>"
        f"Fim: {r['Fim'].strftime('%d/%m %H:%M')}<br>"
        f"Duração: {round(r['Duracao Min'], 1)} min"
    ), axis=1)

    fig = px.timeline(
        dff, x_start="Inicio", x_end="Fim", y="Nome", color="Tipo Parada",
        hover_name="Resumo",
        color_discrete_map={
            "Efetivo": "#046414", "Parada Gerenciável": "#FF9393", "Parada Mecânica": "#A52657",
            "Parada Improdutiva": "#FF0000", "Parada Essencial": "#0026FF",
            "Deslocamento": "#ffee00", "Manobra": "#93c9f7", "Outros": "#8C8C8C", "Outro": "#222"
        }
    )
    fig.update_layout(
        title=f"<b>Atividades de {operador}</b> — pan/zoom atualiza a batelada",
        plot_bgcolor="#181818", paper_bgcolor="#181818",
        font=dict(color="#e9e9e9"), xaxis_title="Horário", yaxis_title="",
        margin=dict(l=40, r=40, t=80, b=60), height=600,
        legend=dict(orientation="v", x=1.02, y=1),
        dragmode="pan"
    )
    fig.update_traces(marker=dict(line=dict(width=1, color="white")))
    fig.update_yaxes(autorange="reversed")

    # Enquadra na janela apenas quando vier dos filtros (não sobrescreve seu pan)
    if ctx.triggered_id in ("operador-dropdown", "data-dropdown"):
        fig.update_xaxes(range=[x0, x1])

    # ===== Faixas por Equipamento (sobreposição translúcida) =====
    blocos_eq = blocos_equipamento(base)
    palette = px.colors.qualitative.Pastel + px.colors.qualitative.Set3 + px.colors.qualitative.Light24
    eqs = blocos_eq["Equipamento"].fillna("Sem equipamento").unique().tolist()
    cmap = {e: palette[i % len(palette)] for i, e in enumerate(eqs)}

    for _, r in blocos_eq.iterrows():
        fig.add_vrect(x0=r["Inicio"], x1=r["Fim"], fillcolor=cmap.get(r["Equipamento"], "#3a2e5f"),
                      opacity=0.12, layer="below", line_width=0)
        xm = r["Inicio"] + (r["Fim"] - r["Inicio"]) / 2
        fig.add_annotation(x=xm, y=1.02, yref="paper", text=str(r["Equipamento"]),
                           showarrow=False, font=dict(size=10, color="#aab2bd"), align="center")

    # Divisores e rótulos de dia (00:00 / "Dia dd/mm")
    add_divisores_de_dia(fig, dff["Inicio"].min(), dff["Fim"].max())

    # Range slider + spikes (navegação confortável)
    fig.update_xaxes(rangeslider_visible=True, showspikes=True,
                     spikemode="across", spikecolor="#bbb", spikedash="dot")

    return fig, stats_html

# ===================== RUN =====================
if __name__ == "__main__":
    app.run_server(debug=True)
