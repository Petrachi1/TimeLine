import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from datetime import timedelta
import unicodedata

# =============== PARÂMETROS ===============
ARQUIVO = "Linha do tempo.xlsx"
SHEET   = "Plan1"

# =============== CARREGAMENTO ===============
df = pd.read_excel(ARQUIVO, sheet_name=SHEET)

# Equipamento (apenas visual; NÃO filtramos por ele)
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]

# Parsing básico
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

# =============== CLASSIFICAÇÃO ===============
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
        return "Parada Improdutiva"  # improdutivas que não são mec/gerenc/essenciais
    if desc == "DESLOCAMENTO":  return "Deslocamento"
    if desc == "MANOBRA":       return "Manobra"
    return "Outro"

df["Tipo Parada"] = df.apply(classifica_tipo_parada, axis=1)

# =============== AGRUPADORES ===============
def agrupar_paradas(df_filtrado):
    """Colapsa blocos contíguos da MESMA operação e MESMO equipamento (gap <= 2min)."""
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

def blocos_equipamento(dff_raw):
    """Faixas de uso por equipamento (independe da operação)."""
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

def normalize_ascii_upper(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s)).encode("ASCII","ignore").decode("ASCII").upper()

def eh_fim_de_expediente(txt: str) -> bool:
    t = normalize_ascii_upper(txt)
    return ("EXPEDIENTE" in t) and (("FIM" in t) or ("FINAL" in t))

# =============== APP ===============
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

# Valores iniciais
primeiro_nome = sorted(df["Nome"].dropna().unique())[0]
primeiras_datas = sorted(df[df["Nome"] == primeiro_nome]["Data Hora Local"].dt.date.unique())
data_padrao = str(primeiras_datas[-2]) if len(primeiras_datas) >= 2 else str(primeiras_datas[-1])

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),

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
        ]), className="mb-3"),

        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-3"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "600px"}))),

        # cache leve: guarda dff (agrupado) e blocos de equipamento pro operador atual
        dcc.Store(id="store-prep"),

        # >>> NOVO: Tabela no rodapé (improdutivas do dia)
        html.Br(),
        dbc.Card(dbc.CardBody([
            html.H4("Paradas improdutivas do dia (00:00 → 24:00)", className="mb-3"),
            html.Div(id="tabela-improdutivas-dia")
        ]), className="mt-2"),
    ], fluid=False)
])

# Atualiza lista de datas quando troca o operador
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

# =============== PREPARO DE DADOS (CACHE EM STORE) ===============
@app.callback(
    Output("store-prep", "data"),
    Input("operador-dropdown", "value"),
)
def preparar_dados(operador):
    base = df[df["Nome"] == operador].copy()
    if base.empty:
        return {"dff": [], "equip": [], "tmin": None, "tmax": None}

    dff = agrupar_paradas(base)
    dff = dff[~dff["Descrição da Operação"].map(eh_fim_de_expediente)]  # remove "fim/final de expediente"
    equip = blocos_equipamento(base)

    # serializa datas como ISO (Dash-friendly)
    dff_json = dff.assign(Inicio=dff["Inicio"].astype(str), Fim=dff["Fim"].astype(str)).to_dict("records")
    equip_json = equip.assign(Inicio=equip["Inicio"].astype(str), Fim=equip["Fim"].astype(str)).to_dict("records")
    tmin = str(dff["Inicio"].min()) if not dff.empty else None
    tmax = str(dff["Fim"].max())    if not dff.empty else None
    return {"dff": dff_json, "equip": equip_json, "tmin": tmin, "tmax": tmax}

# =============== FIGURE (renderiza só quando troca operador/data) ===============
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

@app.callback(
    Output("grafico-linha-tempo", "figure"),
    Input("store-prep", "data"),
    Input("operador-dropdown", "value"),
    Input("data-dropdown", "value"),
)
def desenhar_fig(store, operador, data_str):
    # reconstrói dff/equip a partir do cache
    dff = pd.DataFrame(store.get("dff", []))
    equip = pd.DataFrame(store.get("equip", []))

    if not dff.empty:
        dff["Inicio"] = pd.to_datetime(dff["Inicio"]); dff["Fim"] = pd.to_datetime(dff["Fim"])

    # figura vazia amigável
    if dff.empty:
        fig = px.timeline(pd.DataFrame(columns=["Inicio","Fim","Nome"]), x_start="Inicio", x_end="Fim", y="Nome")
        fig.update_layout(title="Sem dados para exibir.", uirevision=f"op:{operador}")
        return fig

    # gráfico principal (toda a linha do tempo → pan infinito)
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
        title=f"<b>Atividades de {operador}</b>",
        plot_bgcolor="#181818", paper_bgcolor="#181818",
        font=dict(color="#e9e9e9"), xaxis_title="Horário", yaxis_title="",
        margin=dict(l=40, r=40, t=80, b=60), height=600,
        legend=dict(orientation="v", x=1.02, y=1),
        dragmode="pan",
        uirevision=f"op:{operador}"  # mantém pan/zoom/slider
    )
    fig.update_traces(marker=dict(line=dict(width=1, color="white")))
    fig.update_yaxes(autorange="reversed")

    # faixa por equipamento
    if not equip.empty:
        equip["Inicio"] = pd.to_datetime(equip["Inicio"]); equip["Fim"] = pd.to_datetime(equip["Fim"])
        palette = px.colors.qualitative.Pastel + px.colors.qualitative.Set3 + px.colors.qualitative.Light24
        eqs = equip["Equipamento"].fillna("Sem equipamento").unique().tolist()
        cmap = {e: palette[i % len(palette)] for i, e in enumerate(eqs)}
        for _, r in equip.iterrows():
            fig.add_vrect(x0=r["Inicio"], x1=r["Fim"], fillcolor=cmap.get(r["Equipamento"], "#3a2e5f"),
                          opacity=0.12, layer="below", line_width=0)

    # divisores 00:00
    add_divisores_de_dia(fig, store.get("tmin"), store.get("tmax"))

    # range inicial (pela data selecionada) — uirevision preserva depois
    if data_str:
        x0, x1 = janela_inicial_do_dia(data_str)
        fig.update_xaxes(range=[x0, x1], autorange=False)

    # slider + spikes
    fig.update_xaxes(rangeslider_visible=True, showspikes=True,
                     spikemode="across", spikecolor="#bbb", spikedash="dot")

    return fig

# =============== TABELA: IMPRODUTIVAS DO DIA (TODOS OPERADORES) ===============
@app.callback(
    Output("tabela-improdutivas-dia", "children"),
    Input("data-dropdown", "value"),
)
def tabela_improdutivas_dia(data_str):
    if not data_str:
        return html.Div("Selecione uma data.", className="text-center text-muted p-2")

    # janela do dia
    dia0 = pd.to_datetime(data_str).normalize()
    dia1 = dia0 + pd.Timedelta(days=1)

    # pega tudo que SOBREPÕE o dia
    base = df[(df["Fim"] > dia0) & (df["Inicio"] < dia1)].copy()
    if base.empty:
        return html.Div("Nenhuma ocorrência improdutiva nesse dia.", className="text-center text-muted p-2")

    # exclui 'fim/final de expediente'
    base = base[~base["Descrição da Operação"].map(eh_fim_de_expediente)]
    if base.empty:
        return html.Div("Nenhuma ocorrência improdutiva nesse dia.", className="text-center text-muted p-2")

    # recorta para o dia
    base["Inicio"] = base["Inicio"].clip(lower=dia0)
    base["Fim"]    = base["Fim"].clip(upper=dia1)
    base = base[(base["Fim"] > base["Inicio"])]

    # agrupa blocos contíguos por Nome + Operação + Equipamento (contagem de ocorrências justa)
    dff = agrupar_paradas(base)

    # mantém apenas improdutivas (todas as não-efetivas improdutivas)
    alvo = {"Parada Mecânica","Parada Gerenciável","Parada Essencial","Parada Improdutiva"}
    dff = dff[dff["Tipo Parada"].isin(alvo)]
    if dff.empty:
        return html.Div("Nenhuma ocorrência improdutiva nesse dia.", className="text-center text-muted p-2")

    # somas + contagens
    resumo = (dff
              .groupby(["Tipo Parada","Descrição da Operação"], as_index=False)
              .agg(Minutos=("Duracao Min","sum"), Ocorrências=("Descrição da Operação","count")))

    # arredonda minutos
    resumo["Minutos"] = (resumo["Minutos"]).round(0).astype(int)

    # ordena: por tipo, depois por minutos desc
    tipo_ordem = {"Parada Gerenciável":1, "Parada Mecânica":2, "Parada Essencial":3, "Parada Improdutiva":4}
    resumo["ord"] = resumo["Tipo Parada"].map(tipo_ordem).fillna(9)
    resumo = resumo.sort_values(["ord","Minutos"], ascending=[True, False]).drop(columns=["ord"])

    # renomeia colunas p/ exibir
    resumo = resumo.rename(columns={
        "Tipo Parada":"Tipo",
        "Descrição da Operação":"Apontamento"
    })

    # tabela Bootstrap
    return dbc.Table.from_dataframe(
        resumo[["Tipo","Apontamento","Minutos","Ocorrências"]],
        striped=True, bordered=True, hover=True, className="table-sm"
    )

# =============== RUN ===============
if __name__ == "__main__":
    app.run_server(debug=True)
