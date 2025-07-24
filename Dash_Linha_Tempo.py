import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc

# Caminho do Excel
arquivo = "Linha do tempo.xlsx"
df = pd.read_excel(arquivo, sheet_name="Plan1")

# Coluna de equipamento formatado
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]


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
df["Hora Inicial"] = pd.to_datetime(df["Hora Inicial"], format="%H:%M:%S").dt.time
df["Hora Final"] = pd.to_datetime(df["Hora Final"], format="%H:%M:%S").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True)
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

# DASH APP
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

app.layout = html.Div(
    style={"backgroundColor": "#f8f9fa", "padding": "20px"},
    children=[
        dbc.Container([
            html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),
            dbc.Card(dbc.CardBody([
                dbc.Row([
                    dbc.Col(dcc.Dropdown(id="operador-dropdown",
                        options=[{"label": nome, "value": nome} for nome in sorted(df["Nome"].unique())],
                        value=sorted(df["Nome"].unique())[0], placeholder="Selecione um Operador"), md=3),
                    dbc.Col(dcc.Dropdown(id="equipamento-dropdown", placeholder="Selecione um Equipamento"), md=4),
                    dbc.Col(dcc.Dropdown(id="data-dropdown", placeholder="Selecione uma Data"), md=3),
                    dbc.Col(dbc.Button([html.I(className="fa fa-arrow-left me-2"), "Retroceder 1 dia"],
                        id="retroceder-dia", n_clicks=0, color="dark", outline=True, className="w-100"), md=2),
                ], align="center")
            ]), className="mb-4"),
            dbc.Card(dbc.CardBody(id="stats-div"), className="mb-4"),
            dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "550px"}))),
        ], fluid=False)
    ]
)

@app.callback(
    Output("equipamento-dropdown", "options"),
    Output("equipamento-dropdown", "value"),
    Input("operador-dropdown", "value"),
)
def atualizar_equipamentos(operador):
    if not operador:
        return [], None
    equipamentos = df[df["Nome"] == operador]["Equipamento"].unique()
    opcoes = [{"label": eq, "value": eq} for eq in sorted(equipamentos)]
    valor = sorted(equipamentos)[0] if len(equipamentos) > 0 else None
    return opcoes, valor

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
    filtro = (df["Nome"] == operador) & (df["Equipamento"] == equipamento)
    datas_disponiveis = df[filtro]["Data Hora Local"].dt.date.unique()
    datas_ordenadas = sorted(datas_disponiveis)
    opcoes = [{"label": str(data), "value": str(data)} for data in datas_ordenadas]
    triggered_id = ctx.triggered_id
    if not triggered_id or triggered_id in ["operador-dropdown", "equipamento-dropdown"]:
        return opcoes, str(datas_ordenadas[-1]) if datas_ordenadas else None
    if triggered_id == "retroceder-dia" and data_atual:
        datas_str = [str(d) for d in datas_ordenadas]
        try:
            idx = datas_str.index(data_atual)
            novo_idx = max(idx - 1, 0)
            return opcoes, str(datas_ordenadas[novo_idx])
        except ValueError:
            return opcoes, data_atual
    return opcoes, data_atual

def create_stat_card(title, value, color, icon):
    return dbc.Col(
        dbc.Card(dbc.CardBody([
            html.H4(value, className="card-title", style={"color": color, "fontWeight": "bold"}),
            html.P(title, className="card-text text-muted")
        ]), className="text-center shadow-sm"), md=2, className="mb-2"
    )

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
    filtro = (df["Nome"] == operador) & (df["Equipamento"] == equipamento) & (df["Data Hora Local"].dt.date == data)
    hora_inicio_expediente = dff["Inicio"].min().strftime("%H:%M")
    hora_fim_expediente = dff["Fim"].max().strftime("%H:%M")
    dff_raw = df[filtro].copy()
    if dff_raw.empty:
        return {}, html.Div("Nenhum dado encontrado.", className="text-center text-muted p-4")
    dff = agrupar_paradas(dff_raw)
    dff = dff[dff["Descrição da Operação"].str.strip().str.upper() != "FINAL DE EXPEDIENTE"]
    dff["Resumo"] = dff.apply(lambda row: (
        f"Operador: {row['Nome']}<br>"
        f"Tipo: {row['Tipo Parada']}<br>"
        f"Operação: {row['Descrição da Operação']}<br>"
        f"Início: {row['Inicio'].strftime('%H:%M')}<br>"
        f"Fim: {row['Fim'].strftime('%H:%M')}<br>"
        f"Duração: {round(row['Duracao Min'], 2)} min"
    ), axis=1)

    cores = {
        "Efetivo": "#046414", "Parada Gerenciável": "#FF9393", "Parada Mecânica": "#A52657",
        "Parada Improdutiva": "#FF0000", "Parada Essencial": "#0026FF", "Deslocamento": "#ffee00",
        "Manobra": "#93c9f7", "Outros": "#8C8C8C", "Outro": "#222"
    }

    stats = {
        "Efetivo": dff[dff["Tipo Parada"] == "Efetivo"]["Duracao Min"].sum() / 60,
        "Parada Gerenciável": dff[dff["Tipo Parada"] == "Parada Gerenciável"]["Duracao Min"].sum() / 60,
        "Parada Mecânica": dff[dff["Tipo Parada"] == "Parada Mecânica"]["Duracao Min"].sum() / 60,
        "Parada Improdutiva": dff[dff["Tipo Parada"] == "Parada Improdutiva"]["Duracao Min"].sum() / 60,
        "Parada Essencial": dff[dff["Tipo Parada"] == "Parada Essencial"]["Duracao Min"].sum() / 60,
        "Deslocamento": dff[dff["Tipo Parada"] == "Deslocamento"]["Duracao Min"].sum() / 60,
    }
    total_horas = sum(stats.values())
    stats_html = dbc.Row([
        create_stat_card("Início do Expediente", hora_inicio_expediente, "#6c757d", "fa fa-sign-in-alt"),
        create_stat_card("Fim do Expediente", hora_fim_expediente, "#6c757d", "fa fa-sign-out-alt"),
        create_stat_card("Total Horas", f"{total_horas:.2f}h", "#343a40", "fa fa-clock"),
        create_stat_card("Efetivo", f"{stats['Efetivo']:.2f}h", cores["Efetivo"], "fa fa-check-circle"),
        create_stat_card("Parada Gerenciável", f"{stats['Parada Gerenciável']:.2f}h", cores["Parada Gerenciável"], "fa fa-pause-circle"),
        create_stat_card("Parada Mecânica", f"{stats['Parada Mecânica']:.2f}h", cores["Parada Mecânica"], "fa fa-wrench"),
        create_stat_card("Parada Improdutiva", f"{stats['Parada Improdutiva']:.2f}h", cores["Parada Improdutiva"], "fa fa-times-circle"),
        create_stat_card("Parada Essencial", f"{stats['Parada Essencial']:.2f}h", cores["Parada Essencial"], "fa fa-coffee"),
    ], justify="center")

    fig = px.timeline(
        dff, x_start="Inicio", x_end="Fim", y="Nome", color="Tipo Parada",
        hover_name="Resumo", color_discrete_map=cores
    )
    fig.update_layout(
        xaxis_title="Horário do Dia", yaxis_title="", height=550,
        plot_bgcolor='#181818', paper_bgcolor='#181818',
        font=dict(family="Segoe UI, Arial", size=14, color="#e9e9e9"),
        legend=dict(orientation="v", x=1.02, y=1),
        margin=dict(l=40, r=40, t=80, b=60),
        title={
            'text': f"<b>Atividades de {operador}</b> em <span style='color:#39d353'>{data_str}</span>",
            'y': 0.95, 'x': 0.5, 'xanchor': 'center', 'yanchor': 'top',
            'font': dict(size=24)
        },
    )
    fig.update_traces(marker=dict(line=dict(width=1, color='white')))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(showgrid=False)
    return fig, stats_html

if __name__ == "__main__":
    print("Iniciando Dash...")
    app.run_server(debug=True)
