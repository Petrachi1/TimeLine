import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from datetime import timedelta
import unicodedata

# ===================== PARÂMETROS =====================
ARQUIVO = "Linha do tempo.xlsx"
SHEET   = "Plan1"

# ===================== CARGA & LIMPEZA =====================
df = pd.read_excel(ARQUIVO, sheet_name=SHEET)

# Equipamento (agora também pode ser filtrado)
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]

# Parsing
df["Hora Inicial"]    = pd.to_datetime(df["Hora Inicial"], format="%H:%M:%S", errors="coerce").dt.time
df["Hora Final"]      = pd.to_datetime(df["Hora Final"],   format="%H:%M:%S", errors="coerce").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Nome", "Hora Inicial", "Hora Final", "Data Hora Local"]).copy()

# Instantes absolutos
df["Inicio"] = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Inicial']}"), axis=1)
df["Fim"]    = df.apply(lambda r: pd.to_datetime(f"{r['Data Hora Local'].date()} {r['Hora Final']}"), axis=1)

# Cruza meia-noite? soma 1 dia no Fim
mask_cross = df["Fim"] < df["Inicio"]
df.loc[mask_cross, "Fim"] = df.loc[mask_cross, "Fim"] + pd.Timedelta(days=1)

# ===================== CLASSIFICAÇÃO =====================
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

# ===================== UTILIDADES =====================
def normalize_ascii_upper(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s)).encode("ASCII","ignore").decode("ASCII").upper()

def eh_fim_de_expediente(txt: str) -> bool:
    t = normalize_ascii_upper(txt)
    return ("EXPEDIENTE" in t) and (("FIM" in t) or ("FINAL" in t))

def agrupar_paradas(df_filtrado):
    """
    Colapsa blocos contíguos da MESMA operação, MESMO equipamento e MESMO operador (gap <= 2min).
    Retorna: Nome, Inicio, Fim, Descrição da Operação, Duracao Min, Tipo Parada, Equipamento.
    """
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
            if (p["Descrição da Operação"] == op) and (p["Equipamento"] == equip) and (p["Nome"] == nome) and gap <= 2:
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

def add_divisores_de_dia(fig, tmin, tmax):
    if tmin is None or tmax is None: return
    tmin = pd.to_datetime(tmin); tmax = pd.to_datetime(tmax)
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

def janela_inicial_do_dia(data_str):
    base_day = pd.to_datetime(data_str).normalize()
    return base_day, base_day + pd.Timedelta(days=1)

def janela_visivel(data_str, relayoutData, tmin_fallback=None):
    """
    Retorna (x0, x1) da janela atualmente visível no gráfico,
    usando pan/zoom/range-slider se houver, ou [data 00:00, +24h] como fallback.
    """
    if data_str:
        base_day = pd.to_datetime(data_str).normalize()
    elif tmin_fallback is not None:
        base_day = pd.to_datetime(tmin_fallback).normalize()
    else:
        base_day = pd.Timestamp.today().normalize()

    x0 = base_day; x1 = base_day + pd.Timedelta(days=1)
    rd = relayoutData or {}

    if rd.get("xaxis.autorange", False):
        return x0, x1

    if "xaxis.range[0]" in rd and "xaxis.range[1]" in rd:
        return pd.to_datetime(rd["xaxis.range[0]"]), pd.to_datetime(rd["xaxis.range[1]"])

    if "xaxis.range" in rd and isinstance(rd["xaxis.range"], (list,tuple)) and len(rd["xaxis.range"]) == 2:
        return pd.to_datetime(rd["xaxis.range"][0]), pd.to_datetime(rd["xaxis.range"][1])

    if "xaxis.rangeslider.range[0]" in rd and "xaxis.rangeslider.range[1]" in rd:
        return pd.to_datetime(rd["xaxis.rangeslider.range[0]"]), pd.to_datetime(rd["xaxis.rangeslider.range[1]"])

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

        # Filtros
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
            ], align="center"),

            html.Hr(),
            html.H5("Máquinas (equipamentos) na janela visível — desligue para ocultar do gráfico/tabela"),
            dbc.Row([
                dbc.Col(dcc.Dropdown(
                    id="equipamentos-checklist",
                    options=[], value=[],
                    multi=True, placeholder="Selecione as máquinas (padrão: todas)",
                    maxHeight=250
                ), md=8),
                dbc.Col(html.Div(id="resumo-maquinas-div"), md=4)
            ], align="center"),
        ]), className="mb-3"),

        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-3"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "600px"}))),

        # cache leve: dataframe agrupado do operador
        dcc.Store(id="store-prep"),

        # Tabela (recorte visível + operador)
        html.Br(),
        dbc.Card(dbc.CardBody([
            html.H4("Paradas improdutivas — operador & janela visível", className="mb-3"),
            html.Div(id="tabela-improdutivas")
        ]), className="mt-2"),
    ], fluid=False)
])

# Atualiza a lista de datas por operador
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

# Prepara e guarda no Store o agrupado do operador
@app.callback(
    Output("store-prep", "data"),
    Input("operador-dropdown", "value"),
)
def preparar_dados(operador):
    base = df[df["Nome"] == operador].copy()
    if base.empty:
        return {"dff": [], "tmin": None, "tmax": None}

    dff = agrupar_paradas(base)
    dff = dff[~dff["Descrição da Operação"].map(eh_fim_de_expediente)]  # remove "fim/final de expediente"

    dff_json = dff.assign(Inicio=dff["Inicio"].astype(str), Fim=dff["Fim"].astype(str)).to_dict("records")
    tmin = str(dff["Inicio"].min()) if not dff.empty else None
    tmax = str(dff["Fim"].max())    if not dff.empty else None
    return {"dff": dff_json, "tmin": tmin, "tmax": tmax}

# === NOVO: opções/valor do filtro de máquinas (com base na JANELA VISÍVEL) ===
@app.callback(
    Output("equipamentos-checklist", "options"),
    Output("equipamentos-checklist", "value"),
    Input("store-prep", "data"),
    Input("data-dropdown", "value"),
    Input("grafico-linha-tempo", "relayoutData"),
    State("equipamentos-checklist", "value"),
)
def atualizar_equipamentos(store, data_str, relayoutData, sel_prev):
    dff = pd.DataFrame(store.get("dff", []))
    if dff.empty or not data_str:
        return [], []
    dff["Inicio"] = pd.to_datetime(dff["Inicio"]); dff["Fim"] = pd.to_datetime(dff["Fim"])

    # Máquinas que têm qualquer interseção com a JANELA VISÍVEL
    x0, x1 = janela_visivel(data_str, relayoutData, store.get("tmin"))
    mask = (dff["Fim"] > x0) & (dff["Inicio"] < x1)
    usados = sorted(dff.loc[mask, "Equipamento"].dropna().unique().tolist())
    opts = [{"label": e, "value": e} for e in usados]

    # Preserva seleção anterior quando possível
    if sel_prev is None:
        return opts, usados  # 1ª render: todas
    inter = [v for v in (sel_prev or []) if v in usados]
    # Se a janela mudou e nada da seleção antiga existe mais, volta pra todas
    # (a menos que o usuário tenha limpado manualmente: sel_prev == [])
    if not inter and sel_prev != []:
        inter = usados
    return opts, inter

# Desenha figura (respeitando filtro de máquinas e mantendo pan/zoom)
@app.callback(
    Output("grafico-linha-tempo", "figure"),
    Input("store-prep", "data"),
    Input("operador-dropdown", "value"),
    Input("data-dropdown", "value"),
    Input("equipamentos-checklist", "value"),
)
def desenhar_fig(store, operador, data_str, equips_sel):
    dff = pd.DataFrame(store.get("dff", []))
    if not dff.empty:
        dff["Inicio"] = pd.to_datetime(dff["Inicio"]); dff["Fim"] = pd.to_datetime(dff["Fim"])

    if dff.empty:
        fig = px.timeline(pd.DataFrame(columns=["Inicio","Fim","Nome"]), x_start="Inicio", x_end="Fim", y="Nome")
        fig.update_layout(title="Sem dados para exibir.", uirevision=f"op:{operador}")
        return fig

    # aplica filtro de máquinas (se nenhum selecionado, fica vazio)
    if equips_sel:
        dff = dff[dff["Equipamento"].isin(equips_sel)]
    else:
        dff = dff.iloc[0:0]

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
            "Deslocamento": "#ffee00", "Manobra": "#93c9f7",
            "Outros": "#8C8C8C", "Outro": "#222"
        }
    )
    fig.update_layout(
        title=f"<b>Atividades de {operador}</b> — pan/zoom atualiza a batelada",
        plot_bgcolor="#181818", paper_bgcolor="#181818",
        font=dict(color="#e9e9e9"), xaxis_title="Horário", yaxis_title="",
        margin=dict(l=40, r=40, t=80, b=60), height=600,
        legend=dict(orientation="v", x=1.02, y=1),
        dragmode="pan",
        uirevision=f"op:{operador}"   # mantém pan/zoom ao ligar/desligar máquinas
    )
    fig.update_traces(marker=dict(line=dict(width=1, color="white")))
    fig.update_yaxes(autorange="reversed")

    # divisores 00:00 (decorativo)
    add_divisores_de_dia(fig, store.get("tmin"), store.get("tmax"))

    # range inicial pela data escolhida (apenas na troca de data/operador)
    if data_str:
        x0, x1 = janela_inicial_do_dia(data_str)
        fig.update_xaxes(range=[x0, x1], autorange=False)

    # slider + spikes
    fig.update_xaxes(rangeslider_visible=True, showspikes=True,
                     spikemode="across", spikecolor="#bbb", spikedash="dot")
    return fig

# Cards do topo (batelada = janela visível) — respeitando máquinas selecionadas
@app.callback(
    Output("stats-div", "children"),
    Input("store-prep", "data"),
    Input("operador-dropdown", "value"),
    Input("data-dropdown", "value"),
    Input("grafico-linha-tempo", "relayoutData"),
    Input("equipamentos-checklist", "value"),
)
def atualizar_cards(store, operador, data_str, relayoutData, equips_sel):
    dff = pd.DataFrame(store.get("dff", []))
    if dff.empty:
        return html.Div("Sem dados para o operador selecionado.", className="text-center text-muted p-3")
    dff["Inicio"] = pd.to_datetime(dff["Inicio"]); dff["Fim"] = pd.to_datetime(dff["Fim"])

    if equips_sel:
        dff = dff[dff["Equipamento"].isin(equips_sel)]
    else:
        return html.Div("Nenhuma máquina selecionada.", className="text-center text-muted p-3")

    x0, x1 = janela_visivel(data_str, relayoutData, store.get("tmin"))

    dff["Inicio_clip"] = dff["Inicio"].clip(lower=x0)
    dff["Fim_clip"]    = dff["Fim"].clip(upper=x1)
    dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds() / 60.0
    win = dff[dff["Duracao Min Clip"] > 0].copy()

    if win.empty:
        return html.Div("Sem atividade nessa janela.", className="text-center text-muted p-3")

    ini_batelada = win["Inicio_clip"].min().strftime("%d/%m %H:%M")
    fim_batelada = win["Fim_clip"].max().strftime("%d/%m %H:%M")
    total_horas  = win["Duracao Min Clip"].sum() / 60.0

    def soma_h(tipo):
        return win.loc[win["Tipo Parada"] == tipo, "Duracao Min Clip"].sum() / 60.0

    def card(t, v, c):
        return dbc.Col(dbc.Card(dbc.CardBody([
            html.H4(v, style={"color": c, "fontWeight": "bold"}),
            html.P(t, className="text-muted")
        ]), className="text-center shadow-sm"), md=2, className="mb-2")

    return dbc.Row([
        card("Início da batelada", ini_batelada, "#6c757d"),
        card("Fim da batelada",    fim_batelada, "#6c757d"),
        card("Total (janela)",     f"{total_horas:.2f}h", "#343a40"),
        card("Efetivo",            f"{soma_h('Efetivo'):.2f}h", "#046414"),
        card("Gerenciável",        f"{soma_h('Parada Gerenciável'):.2f}h", "#B26B00"),
        card("Mecânica",           f"{soma_h('Parada Mecânica'):.2f}h", "#A52657"),
    ], justify="center")

# Resumo por máquina (horas na janela visível)
@app.callback(
    Output("resumo-maquinas-div", "children"),
    Input("store-prep", "data"),
    Input("data-dropdown", "value"),
    Input("grafico-linha-tempo", "relayoutData"),
    Input("equipamentos-checklist", "value"),
)
def resumo_maquinas(store, data_str, relayoutData, equips_sel):
    dff = pd.DataFrame(store.get("dff", []))
    if dff.empty or not data_str or not equips_sel:
        return html.Div("Sem máquinas selecionadas.", className="text-muted")

    dff["Inicio"] = pd.to_datetime(dff["Inicio"]); dff["Fim"] = pd.to_datetime(dff["Fim"])
    dff = dff[dff["Equipamento"].isin(equips_sel)]

    x0, x1 = janela_visivel(data_str, relayoutData, store.get("tmin"))

    dff["Inicio_clip"] = dff["Inicio"].clip(lower=x0)
    dff["Fim_clip"]    = dff["Fim"].clip(upper=x1)
    dff["Min_clip"]    = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds() / 60.0
    win = dff[dff["Min_clip"] > 0].copy()

    if win.empty:
        return html.Div("Sem atividade na janela.", className="text-muted")

    g = (win.groupby("Equipamento", as_index=False)["Min_clip"].sum()
            .sort_values("Min_clip", ascending=False))
    g["Horas"] = (g["Min_clip"] / 60.0).round(2)
    g = g.drop(columns=["Min_clip"])

    table = dbc.Table.from_dataframe(
        g.rename(columns={"Equipamento": "Equip.", "Horas": "Horas (janela)"}),
        striped=True, bordered=True, hover=True, className="table-sm mb-0"
    )
    return html.Div([
        html.H6("Resumo por máquina", className="mb-2"),
        table
    ])

# Tabela improdutivas (OPERADOR + JANELA VISÍVEL + MÁQUINAS SELECIONADAS)
@app.callback(
    Output("tabela-improdutivas", "children"),
    Input("store-prep", "data"),
    Input("operador-dropdown", "value"),
    Input("data-dropdown", "value"),
    Input("grafico-linha-tempo", "relayoutData"),
    Input("equipamentos-checklist", "value"),
)
def tabela_improdutivas(store, operador, data_str, relayoutData, equips_sel):
    dff = pd.DataFrame(store.get("dff", []))
    if dff.empty:
        return html.Div("Sem dados para o operador selecionado.", className="text-center text-muted p-2")

    dff["Inicio"] = pd.to_datetime(dff["Inicio"]); dff["Fim"] = pd.to_datetime(dff["Fim"])

    # aplica filtro de máquinas
    if equips_sel:
        dff = dff[dff["Equipamento"].isin(equips_sel)]
    else:
        return html.Div("Nenhuma máquina selecionada.", className="text-center text-muted p-2")

    # janela atual
    x0, x1 = janela_visivel(data_str, relayoutData, store.get("tmin"))

    # recorte por janela
    dff["Inicio_clip"] = dff["Inicio"].clip(lower=x0)
    dff["Fim_clip"]    = dff["Fim"].clip(upper=x1)
    dff["Duracao Min Clip"] = (dff["Fim_clip"] - dff["Inicio_clip"]).dt.total_seconds() / 60.0
    win = dff[dff["Duracao Min Clip"] > 0].copy()

    # apenas improdutivas
    alvo = {"Parada Mecânica", "Parada Gerenciável", "Parada Essencial", "Parada Improdutiva"}
    win = win[win["Tipo Parada"].isin(alvo)]
    if win.empty:
        return html.Div("Sem improdutivas nessa janela.", className="text-center text-muted p-2")

    resumo = (win.groupby(["Tipo Parada","Descrição da Operação"], as_index=False)
                 .agg(Minutos=("Duracao Min Clip","sum"),
                      Ocorrências=("Duracao Min Clip","count")))
    resumo["Minutos"] = resumo["Minutos"].round(0).astype(int)

    tipo_ordem = {"Parada Gerenciável":1, "Parada Mecânica":2, "Parada Essencial":3, "Parada Improdutiva":4}
    resumo["ord"] = resumo["Tipo Parada"].map(tipo_ordem).fillna(9)
    resumo = resumo.sort_values(["ord","Minutos"], ascending=[True, False]).drop(columns=["ord"])

    resumo = resumo.rename(columns={"Tipo Parada":"Tipo","Descrição da Operação":"Apontamento"})
    return dbc.Table.from_dataframe(
        resumo[["Tipo","Apontamento","Minutos","Ocorrências"]],
        striped=True, bordered=True, hover=True, className="table-sm"
    )

# ===================== RUN =====================
if __name__ == "__main__":
    app.run_server(debug=True)
