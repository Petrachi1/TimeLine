import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx, dash_table
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from datetime import datetime
import numpy as np
import re, unicodedata

# ================== CONFIGS ==================
START_OF_DAY_HOUR = 7        # âncora do "dia operacional" (07:00 → 07:00)
JORNADA_END_HOUR = 16        # linha de referência no gráfico
JORNADA_END_MIN  = 48
NIGHT_START_HOUR = 19        # noite: 19:00→07:00
NIGHT_END_HOUR   = 7
DESTAQUE_PADRAO_MIN = 15     # limiar p/ destacar deltas na tabela
VAL_ALL_EQUIPS = "__ALL__"   # valor do dropdown p/ "Todos os equipamentos"

# ================== DADOS ==================
arquivo = "Linha do tempo.xlsx"
df = pd.read_excel(arquivo, sheet_name="Plan1")

# -------- mapeamento robusto de colunas --------
def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("ASCII")
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s

def find_col(df, aliases):
    cols_norm = {norm(c): c for c in df.columns}
    for a in aliases:
        na = norm(a)
        if na in cols_norm:
            return cols_norm[na]
    # tentativa por "contém" (fuzzy leve)
    for na, c in cols_norm.items():
        for a in aliases:
            qa = norm(a)
            if qa in na:
                return c
    return None

col_nome       = find_col(df, ["nome","nome do operador","operador"])
col_cod_equip  = find_col(df, ["codigo equipamento","codigo do equipamento","cod equipamento","id equipamento","equipamento id"])
col_desc_equip = find_col(df, ["descricao do equipamento","descrição do equipamento","descricao equipamento","desc equipamento","equipamento"])
col_datahora   = find_col(df, ["data hora local","data hora","data/hora local","data/hora","timestamp","data e hora"])
col_hini       = find_col(df, ["hora inicial","hora inicio","hora início","inicio","início"])
col_hfim       = find_col(df, ["hora final","hora fim","fim","termino","término"])
col_desc_oper  = find_col(df, ["descricao da operacao","descrição da operação","operacao","operação","desc operacao"])
col_grupo      = find_col(df, ["descricao do grupo da operacao","descrição do grupo da operação","grupo da operacao","grupo operacao"])

# checagem mínima
required = [col_nome, col_datahora, col_hini, col_hfim, col_desc_oper, col_grupo]
if any(c is None for c in required):
    print("ATENÇÃO: não reconheci algumas colunas. Veja mapeamento:")
    print("nome:", col_nome, "| datahora:", col_datahora, "| hini:", col_hini, "| hfim:", col_hfim,
          "| desc_oper:", col_desc_oper, "| grupo:", col_grupo)
    # o app continua, mas dropdowns podem vir vazios

# renomeia para nomes canônicos usados no restante do código
rename_map = {}
if col_nome:      rename_map[col_nome] = "Nome"
if col_datahora:  rename_map[col_datahora] = "Data Hora Local"
if col_hini:      rename_map[col_hini] = "Hora Inicial"
if col_hfim:      rename_map[col_hfim] = "Hora Final"
if col_desc_oper: rename_map[col_desc_oper] = "Descrição da Operação"
if col_grupo:     rename_map[col_grupo] = "Descrição do Grupo da Operação"
df = df.rename(columns=rename_map)

# constrói "Equipamento"
if col_cod_equip and col_desc_equip:
    df["Equipamento"] = df[col_cod_equip].astype(str) + " - " + df[col_desc_equip].astype(str)
elif col_desc_equip:
    df["Equipamento"] = df[col_desc_equip].astype(str)
elif col_cod_equip:
    df["Equipamento"] = df[col_cod_equip].astype(str)
else:
    df["Equipamento"] = np.nan  # sem info de equipamento → usaremos "Todos os equipamentos"

# -------- parsing de datas/horas --------
df["Hora Inicial"] = pd.to_datetime(df["Hora Inicial"], format="%H:%M:%S", errors="coerce").dt.time
df["Hora Final"]   = pd.to_datetime(df["Hora Final"],   format="%H:%M:%S", errors="coerce").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Hora Inicial","Hora Final","Data Hora Local","Nome"])  # garante base mínima

df["Inicio"] = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Inicial']}"), axis=1)
df["Fim"]    = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Final']}"), axis=1)

# eventos que viram a meia-noite
mask_cross = df["Fim"] < df["Inicio"]
df.loc[mask_cross, "Fim"] = df.loc[mask_cross, "Fim"] + pd.Timedelta(days=1)

# -------- classificação de parada --------
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
    if desc == "DESLOCAMENTO": return "Deslocamento"
    if desc == "MANOBRA":      return "Manobra"
    return "Outro"

df["Tipo Parada"] = df.apply(classifica_tipo_parada, axis=1)

# -------- utilitários --------
def agrupar_paradas(df_filtrado):
    if df_filtrado.empty:
        return pd.DataFrame(columns=["Nome","Inicio","Fim","Descrição da Operação","Duracao Min","Tipo Parada"])
    d = df_filtrado.sort_values("Inicio").reset_index(drop=True).copy()
    out = []; i = 0
    while i < len(d):
        a = d.loc[i]
        ini, fim = a["Inicio"], a["Fim"]
        op, nome, tipo = a["Descrição da Operação"], a["Nome"], a["Tipo Parada"]
        j = i + 1
        while j < len(d):
            p = d.loc[j]
            gap = (p["Inicio"] - fim).total_seconds()/60.0
            if p["Descrição da Operação"] == op and gap <= 2:
                fim = max(fim, p["Fim"]); j += 1
            else:
                break
        out.append({"Nome": nome, "Inicio": ini, "Fim": fim,
                    "Descrição da Operação": op,
                    "Duracao Min": max(0.0,(fim-ini).total_seconds()/60.0),
                    "Tipo Parada": tipo})
        i = j
    return pd.DataFrame(out)

def janela_operacional(data_str, hour_anchor=START_OF_DAY_HOUR):
    d = pd.to_datetime(data_str).date()
    ws = pd.to_datetime(f"{d} {hour_anchor:02d}:00")
    we = ws + pd.Timedelta(days=1)
    return ws, we

def night_intervals(ws, we):
    ints = []
    eve = ws.replace(hour=NIGHT_START_HOUR, minute=0, second=0, microsecond=0)
    if eve < ws: eve = eve + pd.Timedelta(days=1)
    mid = eve.normalize() + pd.Timedelta(days=1)
    if eve < we: ints.append((max(ws,eve), min(we,mid)))
    nm = ws.normalize() + pd.Timedelta(days=1)
    morning_end = nm.replace(hour=NIGHT_END_HOUR, minute=0, second=0, microsecond=0)
    ints.append((max(ws,nm), min(we,morning_end)))
    return [(s,e) for s,e in ints if e>s]

def overlap_min(a0, a1, b0, b1):
    s = max(a0,b0); e = min(a1,b1)
    return max(0.0,(e-s).total_seconds()/60.0)

def fmt_hhmm(ts):
    return "-" if ts is pd.NaT or pd.isna(ts) else pd.to_datetime(ts).strftime("%H:%M")

# ================== APP ==================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

# valores iniciais seguros
nomes = sorted(df["Nome"].dropna().unique())
primeiro_nome = nomes[0] if nomes else None

def eqs_do_operador(op):
    if op is None: return []
    eqs = sorted(df[df["Nome"]==op]["Equipamento"].dropna().unique().tolist())
    return eqs

primeiro_eq = (eqs_do_operador(primeiro_nome)[0] if eqs_do_operador(primeiro_nome) else VAL_ALL_EQUIPS)

def datas_do_operador(op, eq):
    base = df[df["Nome"]==op].copy() if op else df.iloc[0:0].copy()
    if eq and eq != VAL_ALL_EQUIPS and "Equipamento" in base.columns:
        base = base[base["Equipamento"]==eq]
    return sorted(base["Data Hora Local"].dt.date.unique().tolist())

datas_ini = datas_do_operador(primeiro_nome, primeiro_eq)
data_padrao = str(datas_ini[-1]) if len(datas_ini) else None

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-3", style={"color": "#343a40", "fontWeight": "bold"}),

        # ===== Filtros Globais + botão =====
        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(
                    id="operador-dropdown",
                    options=[{"label": n, "value": n} for n in nomes],
                    value=primeiro_nome,
                    placeholder="Selecione um Operador"), md=3),
                dbc.Col(dcc.Dropdown(
                    id="equipamento-dropdown",
                    options=( [{"label": "Todos os equipamentos", "value": VAL_ALL_EQUIPS}] +
                              [{"label": e, "value": e} for e in eqs_do_operador(primeiro_nome)] ),
                    value=primeiro_eq,
                    placeholder="Selecione um Equipamento"), md=3),
                dbc.Col(dcc.Dropdown(
                    id="data-dropdown",
                    options=[{"label": str(d), "value": str(d)} for d in datas_ini],
                    value=data_padrao,
                    placeholder="Selecione uma Data"), md=3),
                dbc.Col(dbc.Button("Ver Tabela", id="btn-ver-tabela", color="primary", className="w-100"), md=3),
            ], align="center"),
        ]), className="mb-3"),

        dcc.Tabs(id="tabs", value="tab-linha", children=[
            dcc.Tab(label="Linha do Tempo", value="tab-linha", children=[
                dbc.Card(dbc.CardBody(id="stats-div"), className="mb-3"),
                dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "550px"}))),
            ]),
            dcc.Tab(label="Resumo Diário", value="tab-resumo", children=[
                dbc.Card(dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            html.Label("Limiar para destaque (min)"),
                            dcc.Slider(id="delta-min-slider", min=0, max=120, step=5,
                                       value=DESTAQUE_PADRAO_MIN,
                                       marks={0:"0",30:"30",60:"60",90:"90",120:"120"})
                        ], md=8),
                        dbc.Col([
                            html.Label("Filtro"),
                            dcc.Checklist(
                                id="somente-destaques",
                                options=[{"label":" Mostrar somente destaques","value":"on"}],
                                value=[], inputStyle={"marginRight":"6px","marginLeft":"4px"}
                            ),
                            dbc.Button("Voltar ao Gráfico", id="btn-voltar-grafico",
                                       outline=True, color="secondary", className="mt-2")
                        ], md=4),
                    ], className="mb-2"),
                    html.Div(id="tabela-resumo-dia")
                ]), className="mt-2"),
            ]),
        ]),
    ], fluid=False)
])

# ===== navegação de abas =====
@app.callback(Output("tabs","value"), Input("btn-ver-tabela","n_clicks"), prevent_initial_call=True)
def ir_para_tabela(n): return "tab-resumo"

@app.callback(Output("tabs","value"), Input("btn-voltar-grafico","n_clicks"), prevent_initial_call=True)
def voltar_para_grafico(n): return "tab-linha"

# ===== equipamentos por operador (com 'Todos os equipamentos') =====
@app.callback(
    Output("equipamento-dropdown","options"),
    Output("equipamento-dropdown","value"),
    Input("operador-dropdown","value"),
)
def atualizar_equipamento(op):
    eqs = eqs_do_operador(op)
    opts = [{"label":"Todos os equipamentos","value":VAL_ALL_EQUIPS}] + [{"label":e,"value":e} for e in eqs]
    val = VAL_ALL_EQUIPS if not eqs else eqs[0]
    return opts, val

# ===== datas por operador + (opcional) equipamento =====
@app.callback(
    Output("data-dropdown","options"),
    Output("data-dropdown","value"),
    Input("operador-dropdown","value"),
    Input("equipamento-dropdown","value"),
)
def atualizar_datas(op, eq):
    datas = datas_do_operador(op, eq)
    opts = [{"label":str(d),"value":str(d)} for d in datas]
    val = str(datas[-1]) if len(datas) else None
    return opts, val

# ===== gráfico + stats =====
@app.callback(
    Output("grafico-linha-tempo","figure"),
    Output("stats-div","children"),
    Input("operador-dropdown","value"),
    Input("equipamento-dropdown","value"),
    Input("data-dropdown","value")
)
def atualizar_grafico(op, eq, data_str):
    if not op or not data_str:
        return px.timeline(pd.DataFrame(columns=["Inicio","Fim","Nome"]), x_start="Inicio", x_end="Fim", y="Nome"), \
               html.Div("Ajuste os filtros.", className="text-center text-muted p-3")

    ws, we = janela_operacional(data_str, START_OF_DAY_HOUR)
    base = df[(df["Nome"]==op) & (df["Fim"]>ws) & (df["Inicio"]<we)].copy()
    if eq and eq != VAL_ALL_EQUIPS:
        base = base[base["Equipamento"]==eq]

    dff = agrupar_paradas(base)
    if dff.empty:
        fig = px.timeline(pd.DataFrame(columns=["Inicio","Fim","Nome"]), x_start="Inicio", x_end="Fim", y="Nome")
        return fig, html.Div("Sem dados na janela.", className="text-center text-muted p-3")

    exp_set = {"FINAL DE EXPEDIENTE","FIM DE EXPEDIENTE"}
    dff = dff[~dff["Descrição da Operação"].str.upper().str.strip().isin(exp_set)]

    dff["Inicio_clip"] = dff["Inicio"].clip(lower=ws)
    dff["Fim_clip"]    = dff["Fim"].clip(upper=we)
    dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds()/60.0
    dff = dff[dff["Duracao Min Clip"] > 0]

    if dff.empty:
        fig = px.timeline(pd.DataFrame(columns=["Inicio","Fim","Nome"]), x_start="Inicio", x_end="Fim", y="Nome")
        return fig, html.Div("Sem dados úteis para visualizar.", className="text-center text-muted p-3")

    # dia x noite no efetivo
    ints = night_intervals(ws, we)
    def min_noite(r): return sum(overlap_min(r["Inicio_clip"], r["Fim_clip"], s, e) for s, e in ints)
    dff["Min_Noite"] = dff.apply(min_noite, axis=1)
    dff["Min_Dia"]   = dff["Duracao Min Clip"] - dff["Min_Noite"]

    efet = dff[dff["Tipo Parada"]=="Efetivo"]
    efet_noite_h = efet["Min_Noite"].sum()/60.0
    efet_dia_h   = efet["Min_Dia"].sum()/60.0
    total_h      = dff["Duracao Min Clip"].sum()/60.0

    def card(t,v,c): 
        return dbc.Col(dbc.Card(dbc.CardBody([html.H4(v,style={"color":c,"fontWeight":"bold"}),html.P(t,className="text-muted")]),
                                 className="text-center shadow-sm"), md=2, className="mb-2")
    stats = dbc.Row([
        card("Início (janela)", dff["Inicio_clip"].min().strftime("%H:%M"), "#6c757d"),
        card("Fim (janela)",    dff["Fim_clip"].max().strftime("%H:%M"),   "#6c757d"),
        card("Total Horas", f"{total_h:.2f}h", "#343a40"),
        card("Efetivo Dia",   f"{efet_dia_h:.2f}h", "#046414"),
        card("Efetivo Noite", f"{efet_noite_h:.2f}h", "#6f42c1"),
    ], justify="center")

    dff["Resumo"] = dff.apply(lambda r: (
        f"Operador: {r['Nome']}<br>Tipo: {r['Tipo Parada']}<br>"
        f"Operação: {r['Descrição da Operação']}<br>"
        f"Início: {r['Inicio_clip'].strftime('%H:%M')}<br>"
        f"Fim: {r['Fim_clip'].strftime('%H:%M')}<br>"
        f"Duração: {round(r['Duracao Min Clip'],1)} min<br>"
        f"Noite: {round(r['Min_Noite'],1)} min"
    ), axis=1)

    fig = px.timeline(
        dff, x_start="Inicio_clip", x_end="Fim_clip", y="Nome", color="Tipo Parada",
        hover_name="Resumo",
        color_discrete_map={
            "Efetivo":"#046414","Parada Gerenciável":"#FF9393","Parada Mecânica":"#A52657",
            "Parada Improdutiva":"#FF0000","Parada Essencial":"#0026FF",
            "Deslocamento":"#ffee00","Manobra":"#93c9f7","Outros":"#8C8C8C","Outro":"#222"
        }
    )
    fig.update_layout(
        title=f"<b>Atividades de {op}</b> — janela {START_OF_DAY_HOUR:02d}:00 → {START_OF_DAY_HOUR:02d}:00 (D+1)",
        plot_bgcolor='#181818', paper_bgcolor='#181818', font=dict(color="#e9e9e9"),
        xaxis_title="Horário", yaxis_title="", margin=dict(l=40,r=40,t=80,b=60), height=550,
        legend=dict(orientation="v", x=1.02, y=1)
    )
    fig.update_traces(marker=dict(line=dict(width=1, color='white')))
    fig.update_yaxes(autorange="reversed")

    for s,e in night_intervals(ws, we):
        fig.add_vrect(x0=s, x1=e, fillcolor="#3a2e5f", opacity=0.25, layer="below", line_width=0)
    fig.add_vline(x=ws, line_width=1, line_dash="dot", line_color="#9aa0a6")
    jend = ws.replace(hour=JORNADA_END_HOUR, minute=JORNADA_END_MIN)
    if ws <= jend <= we:
        fig.add_vline(x=jend, line_width=1, line_dash="dash", line_color="#9aa0a6")

    return fig, stats

# ===== Tabela diária minimalista (sempre visível) =====
@app.callback(
    Output("tabela-resumo-dia","children"),
    Input("data-dropdown","value"),
    Input("delta-min-slider","value"),
    Input("somente-destaques","value"),
)
def atualizar_resumo_dia(data_str, limiar_min, somente_destaques):
    cols = ["Hora início","Hora fim","Hora início efetivo","Hora fim efetivo","Δ início (min)","Δ fim (min)"]

    if not data_str:
        empty_df = pd.DataFrame(columns=cols)
        return dash_table.DataTable(
            data=empty_df.to_dict("records"),
            columns=[{"name":c,"id":c} for c in cols],
            page_size=20, sort_action="native", filter_action="native",
            style_table={"overflowX":"auto"},
            style_cell={"padding":"6px","fontFamily":"Inter, system-ui, sans-serif","fontSize":"14px"},
            style_header={"backgroundColor":"#f1f3f5","fontWeight":"700"},
        )

    ws, we = janela_operacional(data_str, START_OF_DAY_HOUR)
    base = df[(df["Fim"]>ws) & (df["Inicio"]<we)].copy()

    linhas=[]; tips=[]
    exp_set = {"FINAL DE EXPEDIENTE","FIM DE EXPEDIENTE"}

    for nome, g in base.groupby("Nome"):
        dff = agrupar_paradas(g)
        if dff.empty: continue
        dff = dff[~dff["Descrição da Operação"].str.upper().str.strip().isin(exp_set)]
        if dff.empty: continue

        dff["Inicio_clip"] = dff["Inicio"].clip(lower=ws)
        dff["Fim_clip"]    = dff["Fim"].clip(upper=we)
        dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds()/60.0
        dff = dff[dff["Duracao Min Clip"]>0]
        if dff.empty: continue

        h_ini = dff["Inicio_clip"].min(); h_fim = dff["Fim_clip"].max()
        ef   = dff[dff["Tipo Parada"]=="Efetivo"]
        h_ini_ef = ef["Inicio_clip"].min() if not ef.empty else pd.NaT
        h_fim_ef = ef["Fim_clip"].max()   if not ef.empty else pd.NaT

        delta_ini = (h_ini_ef - h_ini).total_seconds()/60.0 if pd.notna(h_ini_ef) else np.nan
        delta_fim = (h_fim - h_fim_ef).total_seconds()/60.0 if pd.notna(h_fim_ef) else np.nan

        row = {
            "Hora início": fmt_hhmm(h_ini),
            "Hora fim": fmt_hhmm(h_fim),
            "Hora início efetivo": fmt_hhmm(h_ini_ef),
            "Hora fim efetivo": fmt_hhmm(h_fim_ef),
            "Δ início (min)": None if pd.isna(delta_ini) else round(delta_ini,1),
            "Δ fim (min)":    None if pd.isna(delta_fim) else round(delta_fim,1),
        }
        linhas.append(row)
        tips.append({c: f"Operador: {nome}" for c in row.keys()})

    df_resumo = pd.DataFrame(linhas, columns=cols)

    if "on" in (somente_destaques or []):
        mask = (df_resumo["Δ início (min)"].abs().fillna(0) >= limiar_min) | (df_resumo["Δ fim (min)"].abs().fillna(0) >= limiar_min)
        df_resumo = df_resumo[mask]
        tips = [t for t, keep in zip(tips, mask) if keep]

    style_cond = [
        {"if":{"filter_query":f"abs({{{{Δ início (min)}}}}) >= {limiar_min}","column_id":"Δ início (min)"},
         "backgroundColor":"#fff3cd","color":"#5c4400","fontWeight":"600"},
        {"if":{"filter_query":f"abs({{{{Δ fim (min)}}}}) >= {limiar_min}","column_id":"Δ fim (min)"},
         "backgroundColor":"#ffd6d6","color":"#7a0000","fontWeight":"600"},
    ]

    return dash_table.DataTable(
        id="datatable-resumo",
        data=df_resumo.to_dict("records"),
        columns=[{"name":c,"id":c} for c in cols],
        page_size=20, sort_action="native", filter_action="native",
        tooltip_data=tips if len(tips)==len(df_resumo) else None,
        style_table={"overflowX":"auto"},
        style_cell={"padding":"6px","fontFamily":"Inter, system-ui, sans-serif","fontSize":"14px"},
        style_header={"backgroundColor":"#f1f3f5","fontWeight":"700"},
        style_data_conditional=style_cond
    )

# ================== RUN ==================
if __name__ == "__main__":
    app.run_server(debug=True)
