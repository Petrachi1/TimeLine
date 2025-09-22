import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from datetime import timedelta, datetime

# ================== CONFIG JORNADA ==================
JORNADA_MIN = 9 * 60 + 48  # 9h48 -> 588 minutos
TOLERANCIA_MIN = 10        # tolerância de 10 min para não "picar" OK

# Caminho do Excel
arquivo = "Linha do tempo.xlsx"
df = pd.read_excel(arquivo, sheet_name="Plan1")

# Coluna de equipamento formatado
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]

# Classificadores de tipo
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
df["Hora Inicial"] = pd.to_datetime(df["Hora Inicial"], format="%H:%M:%S", errors="coerce").dt.time
df["Hora Final"] = pd.to_datetime(df["Hora Final"], format="%H:%M:%S", errors="coerce").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Hora Inicial", "Hora Final"])
df["Hora Inicial Decimal"] = df["Hora Inicial"].apply(lambda x: x.hour + x.minute / 60)
df["Hora Final Decimal"] = df["Hora Final"].apply(lambda x: x.hour + x.minute / 60)
df["Duracao Min"] = df["Hora Final Decimal"] - df["Hora Inicial Decimal"]

df["Resumo"] = df.apply(lambda row: (
    f"Operador: {row['Nome']}<br>"
    f"Tipo: {row['Tipo']}<br>"
    f"Operação: {row['Descrição da Operação']}<br>"
    f"Início: {row['Hora Inicial']}<br>"
    f"Fim: {row['Hora Final']}<br>"
    f"Duração: {round(row['Duracao Min'], 2)} min"
), axis=1)

df["Inicio"] = df.apply(lambda row: pd.to_datetime(f"{row['Data Hora Local'].date()} {row['Hora Inicial']}"), axis=1)
df["Fim"] = df.apply(lambda row: pd.to_datetime(f"{row['Data Hora Local'].date()} {row['Hora Final']}"), axis=1)

# Classificador de parada
def classifica_tipo_parada(row):
    grupo = str(row["Descrição do Grupo da Operação"]).strip().upper()
    desc = str(row["Descrição da Operação"]).strip().upper()
    gerenciaveis = ["AGUARDANDO COMBUSTIVEL", "AGUARDANDO ORDENS", "AGUARDANDO MOVIMENTACAO PIVO", "FALTA DE INSUMOS"]
    essenciais = ["REFEICAO", "BANHEIRO"]
    mecanicas = ["AGUARDANDO MECANICO", "BORRACHARIA", "EXCESSO DE TEMPERATURA DO MOTOR", "IMPLEMENTO QUEBRADO",
                 "MANUTENCAO ELETRICA", "MANUTENCAO MECANICA", "TRATOR QUEBRADO", "SEM SINAL GPS"]

    if grupo == "PRODUTIVA":
        return "Efetivo"
    elif grupo == "IMPRODUTIVA":
        if desc in gerenciaveis:
            return "Parada Gerenciável"
        elif desc in mecanicas:
            return "Parada Mecânica"
        elif desc in essenciais:
            return "Parada Essencial"
        elif desc == "OUTROS":
            return "Outros"
        else:
            return "Parada Improdutiva"
    elif desc == "DESLOCAMENTO":
        return "Deslocamento"
    elif desc == "MANOBRA":
        return "Manobra"
    else:
        return "Outro"

df["Tipo Parada"] = df.apply(classifica_tipo_parada, axis=1)

# Agrupador de paradas
def agrupar_paradas(df_filtrado):
    df_filtrado = df_filtrado.sort_values(by="Inicio").reset_index(drop=True).copy()
    agrupados = []
    i = 0
    while i < len(df_filtrado):
        atual = df_filtrado.loc[i]
        inicio_bloco = atual["Inicio"]
        fim_bloco = atual["Fim"]
        operacao = atual["Descrição da Operação"]
        nome = atual["Nome"]
        tipo = atual["Tipo Parada"]

        j = i + 1
        while j < len(df_filtrado):
            proximo = df_filtrado.loc[j]
            gap = (proximo["Inicio"] - fim_bloco).total_seconds() / 60
            mesma_operacao = proximo["Descrição da Operação"] == operacao
            if mesma_operacao and gap <= 2:
                fim_bloco = proximo["Fim"]
                j += 1
            else:
                break

        duracao_bloco = (fim_bloco - inicio_bloco).total_seconds() / 60
        novo_bloco = {
            "Nome": nome,
            "Inicio": inicio_bloco,
            "Fim": fim_bloco,
            "Descrição da Operação": operacao,
            "Duracao Min": duracao_bloco,
            "Tipo Parada": tipo,
        }
        agrupados.append(novo_bloco)
        i = j

    return pd.DataFrame(agrupados)

# ===== App
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

# valores iniciais
primeiro_nome = sorted(df["Nome"].unique())[0]
primeiro_eq = sorted(df[df["Nome"] == primeiro_nome]["Equipamento"].unique())[0]
primeiras_datas = sorted(df[(df["Nome"] == primeiro_nome) & (df["Equipamento"] == primeiro_eq)]["Data Hora Local"].dt.date.unique())
data_padrao = str(primeiras_datas[-2]) if len(primeiras_datas) >= 2 else str(primeiras_datas[-1])

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),

        # Filtros
        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(id="operador-dropdown",
                                     options=[{"label": n, "value": n} for n in sorted(df["Nome"].unique())],
                                     value=primeiro_nome, placeholder="Selecione um Operador"), md=3),
                dbc.Col(dcc.Dropdown(id="equipamento-dropdown", value=primeiro_eq,
                                     placeholder="Selecione um Equipamento"), md=4),
                dbc.Col(dcc.Dropdown(id="data-dropdown", value=data_padrao,
                                     placeholder="Selecione uma Data"), md=3),
                dbc.Col(dbc.Button([html.I(className="fa fa-arrow-left me-2"), "Retroceder 1 dia"],
                                   id="retroceder-dia", n_clicks=0, color="dark",
                                   outline=True, className="w-100"), md=2),
            ], align="center")
        ]), className="mb-4"),

        # Linha do tempo
        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-4"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "550px"}))),


        # ===== NOVO: Tabela externa (filtra só por dia; indep. de máquina/operador)
        dbc.Card(dbc.CardBody([
            html.H4("Resumo diário por operador (dia selecionado)", className="mb-3"),
            html.Div(id="tabela-resumo-dia")
        ]), className="mt-4"),
    ], fluid=False)
])

# ---------- CALLBACK: equipamentos por operador ----------
@app.callback(
    Output("equipamento-dropdown", "options"),
    Output("equipamento-dropdown", "value"),
    Input("operador-dropdown", "value"),
)
def atualizar_equipamento(operador):
    if not operador:
        return [], None
    equipamentos = sorted(df[df["Nome"] == operador]["Equipamento"].unique().tolist())
    if not equipamentos:
        return [], None
    return [{"label": eq, "value": eq} for eq in equipamentos], equipamentos[0]

# ---------- CALLBACK: datas por operador+equipamento + retroceder ----------
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
        try:
            idx = str_datas.index(data_atual)
        except ValueError:
            idx = len(str_datas) - 1
        novo_idx = max(0, idx - 1)
        valor = str_datas[novo_idx]
    else:
        valor = str(datas[-2]) if len(datas) >= 2 else str(datas[-1])

    return opcoes_datas, valor

# ---------- CALLBACK: gráfico + stats ----------
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

    data = pd.to_datetime(data_str).date()
    dff_raw = df[(df["Nome"] == operador) & (df["Equipamento"] == equipamento) & (df["Data Hora Local"].dt.date == data)].copy()
    if dff_raw.empty:
        return {}, html.Div("Nenhum dado encontrado.", className="text-center text-muted p-4")

    dff = agrupar_paradas(dff_raw)
    # excluir FINAL/FIM DE EXPEDIENTE
    mask_exp = ~dff["Descrição da Operação"].str.upper().str.strip().isin(["FINAL DE EXPEDIENTE", "FIM DE EXPEDIENTE"])
    dff = dff[mask_exp]

    hora_inicio = dff["Inicio"].min().strftime("%H:%M")
    hora_fim = dff["Fim"].max().strftime("%H:%M")
    stats = {
        "Efetivo": dff[dff["Tipo Parada"] == "Efetivo"]["Duracao Min"].sum() / 60,
        "Parada Gerenciável": dff[dff["Tipo Parada"] == "Parada Gerenciável"]["Duracao Min"].sum() / 60,
        "Parada Mecânica": dff[dff["Tipo Parada"] == "Parada Mecânica"]["Duracao Min"].sum() / 60,
        "Parada Improdutiva": dff[dff["Tipo Parada"] == "Parada Improdutiva"]["Duracao Min"].sum() / 60,
        "Parada Essencial": dff[dff["Tipo Parada"] == "Parada Essencial"]["Duracao Min"].sum() / 60,
        "Deslocamento": dff[dff["Tipo Parada"] == "Deslocamento"]["Duracao Min"].sum() / 60,
        "Manobra": dff[dff["Tipo Parada"] == "Manobra"]["Duracao Min"].sum() / 60,
    }
    total = sum(stats.values())

    def create_stat_card(title, value, color):
        return dbc.Col(dbc.Card(dbc.CardBody([
            html.H4(value, style={"color": color, "fontWeight": "bold"}),
            html.P(title, className="text-muted")
        ]), className="text-center shadow-sm"), md=2, className="mb-2")

    stats_html = dbc.Row([
        create_stat_card("Início do Expediente", hora_inicio, "#6c757d"),
        create_stat_card("Fim do Expediente", hora_fim, "#6c757d"),
        create_stat_card("Total Horas", f"{total:.2f}h", "#343a40"),
        create_stat_card("Efetivo", f"{stats['Efetivo']:.2f}h", "#046414"),
        create_stat_card("Parada Gerenciável", f"{stats['Parada Gerenciável']:.2f}h", "#FF9393"),
        create_stat_card("Parada Mecânica", f"{stats['Parada Mecânica']:.2f}h", "#A52657"),
        create_stat_card("Parada Improdutiva", f"{stats['Parada Improdutiva']:.2f}h", "#FF0000"),
        create_stat_card("Parada Essencial", f"{stats['Parada Essencial']:.2f}h", "#0026FF"),
    ], justify="center")

    dff["Resumo"] = dff.apply(lambda r: (
        f"Operador: {r['Nome']}<br>Tipo: {r['Tipo Parada']}<br>"
        f"Operação: {r['Descrição da Operação']}<br>"
        f"Início: {r['Inicio'].strftime('%H:%M')}<br>"
        f"Fim: {r['Fim'].strftime('%H:%M')}<br>"
        f"Duração: {round(r['Duracao Min'], 2)} min"
    ), axis=1)

    fig = px.timeline(
        dff, x_start="Inicio", x_end="Fim", y="Nome", color="Tipo Parada",
        hover_name="Resumo", color_discrete_map={
            "Efetivo": "#046414", "Parada Gerenciável": "#FF9393", "Parada Mecânica": "#A52657",
            "Parada Improdutiva": "#FF0000", "Parada Essencial": "#0026FF", "Deslocamento": "#ffee00",
            "Manobra": "#93c9f7", "Outros": "#8C8C8C", "Outro": "#222"
        }
    )
    fig.update_layout(
        title=f"<b>Atividades de {operador}</b> em <span style='color:#39d353'>{data_str}</span>",
        plot_bgcolor='#181818', paper_bgcolor='#181818',
        font=dict(color="#e9e9e9"), xaxis_title="Horário", yaxis_title="",
        margin=dict(l=40, r=40, t=80, b=60), height=550,
        legend=dict(orientation="v", x=1.02, y=1)
    )
    fig.update_traces(marker=dict(line=dict(width=1, color='white')))
    fig.update_yaxes(autorange="reversed")
    return fig, stats_html

# ---------- NOVO CALLBACK: Tabela do dia (independe de máquina/operador) ----------
@app.callback(
    Output("tabela-resumo-dia", "children"),
    Input("data-dropdown", "value")
)
def atualizar_resumo_dia(data_str):
    if not data_str:
        return html.Div("Selecione uma data.", className="text-center text-muted p-2")

    data = pd.to_datetime(data_str).date()
    df_dia_raw = df[df["Data Hora Local"].dt.date == data].copy()
    if df_dia_raw.empty:
        return html.Div("Nenhum dado encontrado para a data.", className="text-center text-muted p-2")

    def fmt_horas(mins: float) -> str:
        mins = max(0, float(mins))
        h = int(mins // 60)
        m = int(round(mins % 60))
        return f"{h:02d}:{m:02d}"

    linhas = []
    for nome, grupo_raw in df_dia_raw.groupby("Nome"):
        dff = agrupar_paradas(grupo_raw)
        if dff.empty:
            continue

        # excluir FINAL/FIM DE EXPEDIENTE (independente de máquina)
        mask_exp = ~dff["Descrição da Operação"].str.upper().str.strip().isin(["FINAL DE EXPEDIENTE", "FIM DE EXPEDIENTE"])
        dff = dff[mask_exp]
        if dff.empty:
            continue

        inicio = dff["Inicio"].min()
        fim = dff["Fim"].max()

        # Total = soma dos blocos (min)
        total_min = float(dff["Duracao Min"].sum())

        # % por tipo
        stats_min = dff.groupby("Tipo Parada")["Duracao Min"].sum()
        total_para_pct = float(stats_min.sum()) if float(stats_min.sum()) > 0 else 1.0

        def pct(tipo):
            return f"{(float(stats_min.get(tipo, 0.0)) / total_para_pct) * 100:.1f}%"

        # Primeiro efetivo
        efetivos = dff[dff["Tipo Parada"] == "Efetivo"]
        primeiro_efetivo = efetivos["Inicio"].min().strftime("%H:%M") if not efetivos.empty else "-"

        # Status da jornada (com tolerância)
        if total_min > JORNADA_MIN + TOLERANCIA_MIN:
            status = "Hora Extra"
        elif total_min < JORNADA_MIN - TOLERANCIA_MIN:
            status = "Suspeito"
        else:
            status = "OK"

        linhas.append({
            "Nome": nome,
            "Início": inicio.strftime("%H:%M"),
            "Fim": fim.strftime("%H:%M"),
            "Horas Trabalhadas": fmt_horas(total_min),
            "Status Jornada": status,
            "Primeiro Efetivo": primeiro_efetivo,
            "Efetivo %": pct("Efetivo"),
            "Parada Gerenciável %": pct("Parada Gerenciável"),
            "Parada Mecânica %": pct("Parada Mecânica"),
            "Parada Improdutiva %": pct("Parada Improdutiva"),
            "Parada Essencial %": pct("Parada Essencial"),
            "Deslocamento %": pct("Deslocamento"),
            "Manobra %": pct("Manobra"),
        })

    if not linhas:
        return html.Div("Sem registros consolidados para a data.", className="text-center text-muted p-2")

    df_resumo = pd.DataFrame(linhas).sort_values(["Status Jornada", "Nome"]).reset_index(drop=True)

    return dbc.Table.from_dataframe(
        df_resumo,
        striped=True,
        bordered=True,
        hover=True,
        className="table-sm"
    )

if __name__ == "__main__":
    app.run_server(debug=True)
