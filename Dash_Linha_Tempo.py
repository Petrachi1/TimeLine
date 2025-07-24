import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from datetime import timedelta, datetime

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

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

primeiro_nome = sorted(df["Nome"].unique())[0]
primeiro_eq = sorted(df[df["Nome"] == primeiro_nome]["Equipamento"].unique())[0]
primeiras_datas = sorted(df[(df["Nome"] == primeiro_nome) & (df["Equipamento"] == primeiro_eq)]["Data Hora Local"].dt.date.unique())
data_padrao = str(primeiras_datas[-2]) if len(primeiras_datas) >= 2 else str(primeiras_datas[-1])

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),
        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(id="operador-dropdown", options=[{"label": nome, "value": nome} for nome in sorted(df["Nome"].unique())], value=primeiro_nome, placeholder="Selecione um Operador"), md=3),
                dbc.Col(dcc.Dropdown(id="equipamento-dropdown", value=primeiro_eq, placeholder="Selecione um Equipamento"), md=4),
                dbc.Col(dcc.Dropdown(id="data-dropdown", value=data_padrao, placeholder="Selecione uma Data"), md=3),
                dbc.Col(dbc.Button([html.I(className="fa fa-arrow-left me-2"), "Retroceder 1 dia"], id="retroceder-dia", n_clicks=0, color="dark", outline=True, className="w-100"), md=2),
            ], align="center")
        ]), className="mb-4"),
        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-4"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "550px"}))),
    ], fluid=False)
])

@app.callback(
    Output("equipamento-dropdown", "options"),
    Output("equipamento-dropdown", "value"),
    Output("data-dropdown", "options"),
    Output("data-dropdown", "value"),
    Input("operador-dropdown", "value"),
)
def atualizar_equipamento_e_data(operador):
    if not operador:
        return [], None, [], None

    equipamentos = sorted(df[df["Nome"] == operador]["Equipamento"].unique())
    if not equipamentos:
        return [], None, [], None

    equipamento_padrao = equipamentos[0]

    datas = sorted(df[(df["Nome"] == operador) & (df["Equipamento"] == equipamento_padrao)]["Data Hora Local"].dt.date.unique())
    opcoes_datas = [{"label": str(d), "value": str(d)} for d in datas]

    if len(datas) >= 2:
        data_padrao = str(datas[-2])
    elif datas:
        data_padrao = str(datas[-1])
    else:
        data_padrao = None

    return (
        [{"label": eq, "value": eq} for eq in equipamentos],
        equipamento_padrao,
        opcoes_datas,
        data_padrao,
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
    dff_raw = df[(df["Nome"] == operador) & (df["Equipamento"] == equipamento) & (df["Data Hora Local"].dt.date == data)].copy()
    if dff_raw.empty:
        return {}, html.Div("Nenhum dado encontrado.", className="text-center text-muted p-4")

    dff = agrupar_paradas(dff_raw)
    dff = dff[dff["Descrição da Operação"].str.upper().str.strip() != "FINAL DE EXPEDIENTE"]
    hora_inicio = dff["Inicio"].min().strftime("%H:%M")
    hora_fim = dff["Fim"].max().strftime("%H:%M")
    stats = {
        "Efetivo": dff[dff["Tipo Parada"] == "Efetivo"]["Duracao Min"].sum() / 60,
        "Parada Gerenciável": dff[dff["Tipo Parada"] == "Parada Gerenciável"]["Duracao Min"].sum() / 60,
        "Parada Mecânica": dff[dff["Tipo Parada"] == "Parada Mecânica"]["Duracao Min"].sum() / 60,
        "Parada Improdutiva": dff[dff["Tipo Parada"] == "Parada Improdutiva"]["Duracao Min"].sum() / 60,
        "Parada Essencial": dff[dff["Tipo Parada"] == "Parada Essencial"]["Duracao Min"].sum() / 60,
        "Deslocamento": dff[dff["Tipo Parada"] == "Deslocamento"]["Duracao Min"].sum() / 60,
    }
    total = sum(stats.values())

    def create_stat_card(title, value, color, icon):
        return dbc.Col(dbc.Card(dbc.CardBody([
            html.H4(value, style={"color": color, "fontWeight": "bold"}),
            html.P(title, className="text-muted")
        ]), className="text-center shadow-sm"), md=2, className="mb-2")

    stats_html = dbc.Row([
        create_stat_card("Início do Expediente", hora_inicio, "#6c757d", "fa fa-sign-in-alt"),
        create_stat_card("Fim do Expediente", hora_fim, "#6c757d", "fa fa-sign-out-alt"),
        create_stat_card("Total Horas", f"{total:.2f}h", "#343a40", "fa fa-clock"),
        create_stat_card("Efetivo", f"{stats['Efetivo']:.2f}h", "#046414", "fa fa-check-circle"),
        create_stat_card("Parada Gerenciável", f"{stats['Parada Gerenciável']:.2f}h", "#FF9393", "fa fa-pause-circle"),
        create_stat_card("Parada Mecânica", f"{stats['Parada Mecânica']:.2f}h", "#A52657", "fa fa-wrench"),
        create_stat_card("Parada Improdutiva", f"{stats['Parada Improdutiva']:.2f}h", "#FF0000", "fa fa-times-circle"),
        create_stat_card("Parada Essencial", f"{stats['Parada Essencial']:.2f}h", "#0026FF", "fa fa-coffee"),
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

if __name__ == "__main__":
    app.run_server(debug=True)
