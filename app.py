import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx, dash_table
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from datetime import datetime
import numpy as np

# ================== CONFIGS ==================
START_OF_DAY_HOUR = 7        # Âncora do "dia operacional" (ex.: 07:00 -> 07:00)
JORNADA_END_HOUR = 16        # Referência jornada (16:48)
JORNADA_END_MIN  = 48
NIGHT_START_HOUR = 19        # Noite: 19:00 → 07:00
NIGHT_END_HOUR   = 7
DESTAQUE_PADRAO_MIN = 15     # limiar para destaque (min)

# ================== DADOS ==================
arquivo = "Linha do tempo.xlsx"
df = pd.read_excel(arquivo, sheet_name="Plan1")

# Equipamento
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]

# Macro tipo
def classifica_tipo(row):
    desc = str(row["Descrição da Operação"]).strip().upper()
    grupo = str(row["Descrição do Grupo da Operação"]).strip().upper()
    if desc == "DESLOCAMENTO":
        return "Deslocamento"
    elif desc == "MANOBRA":
        return "Manobra"
    elif grupo == "PRODUTIVA":
        return "Produtiva"
    elif grupo == "IMPRODUTIVA":
        return "Improdutiva"
    else:
        return "Outro"

df["Tipo"] = df.apply(classifica_tipo, axis=1)

# Tempos
df["Hora Inicial"] = pd.to_datetime(df["Hora Inicial"], format="%H:%M:%S", errors="coerce").dt.time
df["Hora Final"]   = pd.to_datetime(df["Hora Final"],   format="%H:%M:%S", errors="coerce").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Hora Inicial", "Hora Final", "Data Hora Local"]).copy()

# Concatena data + hora
df["Inicio"] = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Inicial']}"), axis=1)
df["Fim"]    = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Final']}"), axis=1)

# CORREÇÃO DE VIRADA DE DIA: se Fim < Inicio, adiciona 1 dia ao Fim (NÃO trocar!)
mask_cross = df["Fim"] < df["Inicio"]
df.loc[mask_cross, "Fim"] = df.loc[mask_cross, "Fim"] + pd.Timedelta(days=1)   # <<<

# Tipo Parada (negócio)
def classifica_tipo_parada(row):
    grupo = str(row["Descrição do Grupo da Operação"]).strip().upper()
    desc = str(row["Descrição da Operação"]).strip().upper()
    gerenciaveis = {"AGUARDANDO COMBUSTIVEL","AGUARDANDO ORDENS","AGUARDANDO MOVIMENTACAO PIVO","FALTA DE INSUMOS"}
    essenciais   = {"REFEICAO","BANHEIRO"}
    mecanicas    = {"AGUARDANDO MECANICO","BORRACHARIA","EXCESSO DE TEMPERATURA DO MOTOR","IMPLEMENTO QUEBRADO",
                    "MANUTENCAO ELETRICA","MANUTENCAO MECANICA","TRATOR QUEBRADO","SEM SINAL GPS"}
    if grupo == "PRODUTIVA":
        return "Efetivo"
    elif grupo == "IMPRODUTIVA":
        if desc in gerenciaveis: return "Parada Gerenciável"
        if desc in mecanicas:    return "Parada Mecânica"
        if desc in essenciais:   return "Parada Essencial"
        if desc == "OUTROS":     return "Outros"
        return "Parada Improdutiva"
    elif desc == "DESLOCAMENTO":
        return "Deslocamento"
    elif desc == "MANOBRA":
        return "Manobra"
    return "Outro"

df["Tipo Parada"] = df.apply(classifica_tipo_parada, axis=1)

# Agrupa blocos contíguos (gap <= 2 min) da MESMA operação
def agrupar_paradas(df_filtrado):
    if df_filtrado.empty:
        return pd.DataFrame(columns=["Nome","Inicio","Fim","Descrição da Operação","Duracao Min","Tipo Parada"])
    d = df_filtrado.sort_values("Inicio").reset_index(drop=True).copy()
    out = []
    i = 0
    while i < len(d):
        atual = d.loc[i]
        ini, fim = atual["Inicio"], atual["Fim"]
        op, nome, tipo = atual["Descrição da Operação"], atual["Nome"], atual["Tipo Parada"]
        j = i + 1
        while j < len(d):
            prox = d.loc[j]
            gap_min = (prox["Inicio"] - fim).total_seconds() / 60.0
            if prox["Descrição da Operação"] == op and gap_min <= 2:
                fim = max(fim, prox["Fim"])
                j += 1
            else:
                break
        dur_min = max(0.0, (fim - ini).total_seconds() / 60.0)
        out.append({"Nome": nome, "Inicio": ini, "Fim": fim,
                    "Descrição da Operação": op, "Duracao Min": dur_min, "Tipo Parada": tipo})
        i = j
    return pd.DataFrame(out)

# Janela 07→07
def janela_operacional(data_str: str, hour_anchor: int = START_OF_DAY_HOUR):
    d = pd.to_datetime(data_str).date()
    win_start = pd.to_datetime(f"{d} {hour_anchor:02d}:00")
    win_end   = win_start + pd.Timedelta(days=1)
    return win_start, win_end

def overlap_min(a0, a1, b0, b1):
    start = max(a0, b0)
    end   = min(a1, b1)
    return max(0.0, (end - start).total_seconds() / 60.0)

def night_intervals(win_start, win_end):
    """Noite: 19:00→00:00 e 00:00→07:00 dentro da janela."""
    intervals = []
    # 1) D 19:00 → D+1 00:00
    eve_start = win_start.replace(hour=NIGHT_START_HOUR, minute=0, second=0, microsecond=0)
    if eve_start < win_start:
        eve_start = eve_start + pd.Timedelta(days=1)
    midnight = eve_start.normalize() + pd.Timedelta(days=1)
    if eve_start < win_end:
        intervals.append((max(win_start, eve_start), min(win_end, midnight)))
    # 2) D+1 00:00 → D+1 07:00
    next_midnight = win_start.normalize() + pd.Timedelta(days=1)
    morning_end = next_midnight.replace(hour=NIGHT_END_HOUR, minute=0, second=0, microsecond=0)
    intervals.append((max(win_start, next_midnight), min(win_end, morning_end)))
    return [(s, e) for s, e in intervals if e > s]

def fmt_hhmm(ts):
    return "-" if ts is pd.NaT or pd.isna(ts) else pd.to_datetime(ts).strftime("%H:%M")

def fmt_horas(mins: float) -> str:
    mins = max(0.0, float(mins))
    h = int(mins // 60); m = int(round(mins % 60))
    return f"{h:02d}:{m:02d}"

# ================== APP ==================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

# Iniciais
primeiro_nome = sorted(df["Nome"].dropna().unique())[0]
primeiro_eq = sorted(df[df["Nome"] == primeiro_nome]["Equipamento"].dropna().unique())[0]
primeiras_datas = sorted(df[(df["Nome"] == primeiro_nome) & (df["Equipamento"] == primeiro_eq)]["Data Hora Local"].dt.date.unique())
data_padrao = str(primeiras_datas[-2]) if len(primeiras_datas) >= 2 else str(primeiras_datas[-1])

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),

        # Filtros (linha do tempo)
        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(id="operador-dropdown",
                                     options=[{"label": n, "value": n} for n in sorted(df["Nome"].dropna().unique())],
                                     value=primeiro_nome, placeholder="Selecione um Operador"), md=3),
                dbc.Col(dcc.Dropdown(id="equipamento-dropdown", value=primeiro_eq, placeholder="Selecione um Equipamento"), md=4),
                dbc.Col(dcc.Dropdown(id="data-dropdown", value=data_padrao, placeholder="Selecione uma Data"), md=3),
                dbc.Col(dbc.Button([html.I(className="fa fa-arrow-left me-2"), "Retroceder 1 dia"],
                                   id="retroceder-dia", n_clicks=0, color="dark", outline=True, className="w-100"), md=2),
            ], align="center")
        ]), className="mb-4"),

        # Cards + gráfico
        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-4"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "550px"}))),

        # Tabela minimalista
        dbc.Card(dbc.CardBody([
            html.H4("Resumo diário por operador (dia operacional)", className="mb-1"),
            html.P(f"Janela: {START_OF_DAY_HOUR:02d}:00 → {START_OF_DAY_HOUR:02d}:00 (D+1)  •  Noite: {NIGHT_START_HOUR:02d}:00 → {NIGHT_END_HOUR:02d}:00",
                   className="text-muted", style={"marginTop": "-4px"}),
            dbc.Row([
                dbc.Col([
                    html.Label("Limiar para destaque (min)"),
                    dcc.Slider(id="delta-min-slider", min=0, max=120, step=5, value=DESTAQUE_PADRAO_MIN,
                               marks={0:"0",30:"30",60:"60",90:"90",120:"120"})
                ], md=6),
                dbc.Col([
                    html.Label("Filtro"),
                    dcc.Checklist(
                        id="somente-destaques",
                        options=[{"label":" Mostrar somente destaques","value":"on"}],
                        value=[],
                        inputStyle={"marginRight":"6px","marginLeft":"4px"}
                    )
                ], md=6),
            ], className="mb-2"),
            html.Div(id="tabela-resumo-dia")
        ]), className="mt-4"),
    ], fluid=False)
])

# ---------- Equipamentos por operador ----------
@app.callback(
    Output("equipamento-dropdown", "options"),
    Output("equipamento-dropdown", "value"),
    Input("operador-dropdown", "value"),
)
def atualizar_equipamento(operador):
    if not operador:
        return [], None
    equipamentos = sorted(df[df["Nome"] == operador]["Equipamento"].dropna().unique().tolist())
    if not equipamentos:
        return [], None
    return [{"label": eq, "value": eq} for eq in equipamentos], equipamentos[0]

# ---------- Datas por operador + equipamento + retroceder ----------
@app.callback(
    Output("data-dropdown", "options"),
    Output("data-dropdown", "value"),
    Input("operador-dropdown", "value"),
    Input("equipamento-dropdown", "value"),
    Input("retroceder-dia", "n_clicks"),
    State("data-dropdown", "value"),
)
def atualizar_datas(operador, equipamento, n_clicks, data_atual):
    if not operador or not equipamento:
        return [], None
    datas = sorted(df[(df["Nome"] == operador) & (df["Equipamento"] == equipamento)]["Data Hora Local"].dt.date.unique().tolist())
    if not datas:
        return [], None
    opcoes_datas = [{"label": str(d), "value": str(d)} for d in datas]

    trigger = ctx.triggered_id
    if trigger == "retroceder-dia" and data_atual:
        str_datas = [str(d) for d in datas]
        try: idx = str_datas.index(data_atual)
        except ValueError: idx = len(str_datas) - 1
        novo_idx = max(0, idx - 1)
        valor = str_datas[novo_idx]
    else:
        valor = str(datas[-2]) if len(datas) >= 2 else str(datas[-1])

    return opcoes_datas, valor

# ---------- Gráfico + stats com noite destacada ----------
@app.callback(
    Output("grafico-linha-tempo", "figure"),
    Output("stats-div", "children"),
    Input("operador-dropdown", "value"),
    Input("equipamento-dropdown", "value"),
    Input("data-dropdown", "value")
)
def atualizar_grafico(operador, equipamento, data_str):
    if not operador or not equipamento or not data_str:
        return {}, html.Div("Ajuste os filtros.", className="text-center text-muted p-4")

    win_start, win_end = janela_operacional(data_str, START_OF_DAY_HOUR)
    dff_raw = df[(df["Nome"] == operador) &
                 (df["Equipamento"] == equipamento) &
                 (df["Fim"] > win_start) & (df["Inicio"] < win_end)].copy()
    if dff_raw.empty:
        return {}, html.Div("Nenhum dado encontrado.", className="text-center text-muted p-4")

    dff = agrupar_paradas(dff_raw)
    exp_set = {"FINAL DE EXPEDIENTE", "FIM DE EXPEDIENTE"}
    dff = dff[~dff["Descrição da Operação"].str.upper().str.strip().isin(exp_set)]
    if dff.empty:
        return {}, html.Div("Nenhum dado útil para visualizar.", className="text-center text-muted p-4")

    # Clip na janela
    dff["Inicio_clip"] = dff["Inicio"].clip(lower=win_start)
    dff["Fim_clip"]    = dff["Fim"].clip(upper=win_end)
    dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds() / 60.0
    dff = dff[dff["Duracao Min Clip"] > 0]

    hora_inicio = dff["Inicio_clip"].min().strftime("%H:%M")
    hora_fim    = dff["Fim_clip"].max().strftime("%H:%M")

    # Dia x Noite no Efetivo
    night_ints = night_intervals(win_start, win_end)
    def minutos_noite(row):
        return sum(overlap_min(row["Inicio_clip"], row["Fim_clip"], s, e) for s, e in night_ints)
    dff["Min_Noite"] = dff.apply(minutos_noite, axis=1)
    dff["Min_Dia"]   = dff["Duracao Min Clip"] - dff["Min_Noite"]

    efet = dff[dff["Tipo Parada"] == "Efetivo"]
    efet_noite_h = efet["Min_Noite"].sum() / 60.0
    efet_dia_h   = efet["Min_Dia"].sum() / 60.0
    total_h      = dff["Duracao Min Clip"].sum() / 60.0

    def create_stat_card(title, value, color):
        return dbc.Col(dbc.Card(dbc.CardBody([
            html.H4(value, style={"color": color, "fontWeight": "bold"}),
            html.P(title, className="text-muted")
        ]), className="text-center shadow-sm"), md=2, className="mb-2")

    stats_html = dbc.Row([
        create_stat_card("Início (janela)", hora_inicio, "#6c757d"),
        create_stat_card("Fim (janela)",    hora_fim,    "#6c757d"),
        create_stat_card("Total Horas", f"{total_h:.2f}h", "#343a40"),
        create_stat_card("Efetivo Dia",   f"{efet_dia_h:.2f}h", "#046414"),
        create_stat_card("Efetivo Noite", f"{efet_noite_h:.2f}h", "#6f42c1"),
    ], justify="center")

    dff["Resumo"] = dff.apply(lambda r: (
        f"Operador: {r['Nome']}<br>Tipo: {r['Tipo Parada']}<br>"
        f"Operação: {r['Descrição da Operação']}<br>"
        f"Início: {r['Inicio_clip'].strftime('%H:%M')}<br>"
        f"Fim: {r['Fim_clip'].strftime('%H:%M')}<br>"
        f"Duração: {round(r['Duracao Min Clip'], 2)} min<br>"
        f"Noite: {round(r['Min_Noite'], 1)} min"
    ), axis=1)

    fig = px.timeline(
        dff, x_start="Inicio_clip", x_end="Fim_clip", y="Nome", color="Tipo Parada",
        hover_name="Resumo",
        color_discrete_map={
            "Efetivo": "#046414", "Parada Gerenciável": "#FF9393", "Parada Mecânica": "#A52657",
            "Parada Improdutiva": "#FF0000", "Parada Essencial": "#0026FF",
            "Deslocamento": "#ffee00", "Manobra": "#93c9f7", "Outros": "#8C8C8C", "Outro": "#222"
        }
    )
    fig.update_layout(
        title=f"<b>Atividades de {operador}</b> — janela {START_OF_DAY_HOUR:02d}:00 → {START_OF_DAY_HOUR:02d}:00 (D+1)",
        plot_bgcolor='#181818', paper_bgcolor='#181818',
        font=dict(color="#e9e9e9"), xaxis_title="Horário", yaxis_title="",
        margin=dict(l=40, r=40, t=80, b=60), height=550,
        legend=dict(orientation="v", x=1.02, y=1)
    )
    fig.update_traces(marker=dict(line=dict(width=1, color='white')))
    fig.update_yaxes(autorange="reversed")

    # Faixas de noite
    for s, e in night_intervals(win_start, win_end):
        fig.add_vrect(x0=s, x1=e, fillcolor="#3a2e5f", opacity=0.25, layer="below", line_width=0)

    # Linhas de referência
    fig.add_vline(x=win_start, line_width=1, line_dash="dot", line_color="#9aa0a6")
    jornada_end = win_start.replace(hour=JORNADA_END_HOUR, minute=JORNADA_END_MIN)
    if win_start <= jornada_end <= win_end:
        fig.add_vline(x=jornada_end, line_width=1, line_dash="dash", line_color="#9aa0a6")

    return fig, stats_html

# ---------- Tabela diária minimalista ----------
@app.callback(
    Output("tabela-resumo-dia", "children"),
    Input("data-dropdown", "value"),
    Input("delta-min-slider", "value"),
    Input("somente-destaques", "value"),
)
def atualizar_resumo_dia(data_str, limiar_min, somente_destaques):
    if not data_str:
        return html.Div("Selecione uma data.", className="text-center text-muted p-2")

    win_start, win_end = janela_operacional(data_str, START_OF_DAY_HOUR)
    df_win = df[(df["Fim"] > win_start) & (df["Inicio"] < win_end)].copy()
    if df_win.empty:
        return html.Div("Nenhum dado encontrado para a janela operacional.", className="text-center text-muted p-2")

    linhas = []
    exp_set = {"FINAL DE EXPEDIENTE", "FIM DE EXPEDIENTE"}

    for nome, grupo_raw in df_win.groupby("Nome"):
        dff = agrupar_paradas(grupo_raw)
        if dff.empty:
            continue
        dff = dff[~dff["Descrição da Operação"].str.upper().str.strip().isin(exp_set)]
        if dff.empty:
            continue

        # clip na janela
        dff["Inicio_clip"] = dff["Inicio"].clip(lower=win_start)
        dff["Fim_clip"]    = dff["Fim"].clip(upper=win_end)
        dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds() / 60.0
        dff = dff[dff["Duracao Min Clip"] > 0]
        if dff.empty:
            continue

        # início/fim gerais
        h_inicio = dff["Inicio_clip"].min()
        h_fim    = dff["Fim_clip"].max()

        # primeiro/último efetivo
        efet = dff[dff["Tipo Parada"] == "Efetivo"]
        h_inicio_ef = efet["Inicio_clip"].min() if not efet.empty else pd.NaT
        h_fim_ef    = efet["Fim_clip"].max() if not efet.empty else pd.NaT

        # deltas (min)
        delta_ini = (h_inicio_ef - h_inicio).total_seconds()/60.0 if pd.notna(h_inicio_ef) else np.nan
        delta_fim = (h_fim - h_fim_ef).total_seconds()/60.0 if pd.notna(h_fim_ef) else np.nan

        linhas.append({
            "Nome": nome,
            "Hora início": fmt_hhmm(h_inicio),
            "Hora fim": fmt_hhmm(h_fim),
            "Hora início efetivo": fmt_hhmm(h_inicio_ef),
            "Hora fim efetivo": fmt_hhmm(h_fim_ef),
            "Δ início (min)": None if pd.isna(delta_ini) else round(delta_ini, 1),
            "Δ fim (min)":    None if pd.isna(delta_fim) else round(delta_fim, 1),
        })

    if not linhas:
        return html.Div("Sem registros consolidados para a janela operacional.", className="text-center text-muted p-2")

    df_resumo = pd.DataFrame(linhas).sort_values(["Nome"]).reset_index(drop=True)

    # somente destaques
    if "on" in (somente_destaques or []):
        mask = (df_resumo["Δ início (min)"].abs().fillna(0) >= limiar_min) | (df_resumo["Δ fim (min)"].abs().fillna(0) >= limiar_min)
        df_resumo = df_resumo[mask]

    style_cond = [
        {
            "if": {"filter_query": f"abs({{{{Δ início (min)}}}}) >= {limiar_min}", "column_id": "Δ início (min)"},
            "backgroundColor": "#fff3cd", "color": "#5c4400", "fontWeight": "600"
        },
        {
            "if": {"filter_query": f"abs({{{{Δ fim (min)}}}}) >= {limiar_min}", "column_id": "Δ fim (min)"},
            "backgroundColor": "#ffd6d6", "color": "#7a0000", "fontWeight": "600"
        },
    ]

    return dash_table.DataTable(
        id="datatable-resumo",
        data=df_resumo.to_dict("records"),
        columns=[{"name": c, "id": c} for c in df_resumo.columns],
        page_size=20,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"padding": "6px", "fontFamily": "Inter, system-ui, sans-serif", "fontSize": "14px"},
        style_header={"backgroundColor": "#f1f3f5", "fontWeight": "700"},
        style_data_conditional=style_cond
    )

# ================== RUN ==================
if __name__ == "__main__":
    app.run_server(debug=True)
