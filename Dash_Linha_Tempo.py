import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash import ctx

# Caminho do Excel
arquivo = "Linha do tempo.xlsx"
df = pd.read_excel(arquivo, sheet_name="Plan1")

# Crie coluna de equipamento formatado
df["Equipamento"] = df["Código Equipamento"].astype(str) + " - " + df["Descrição do Equipamento"]

# Função para classificar tipo
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

df["Resumo"] = (
    "Operador: " + df["Nome"] +
    "<br>Tipo: " + df["Tipo"] +
    "<br>Operação: " + df["Descrição da Operação"] +
    "<br>Início: " + df["Hora Inicial"].astype(str) +
    "<br>Fim: " + df["Hora Final"].astype(str) +
    "<br>Duração: " + df["Duracao Min"].round(2).astype(str) + " min"
)
df["Inicio"] = df.apply(lambda row: pd.to_datetime(f"{row['Data Hora Local'].date()} {row['Hora Inicial']}"), axis=1)
df["Fim"] = df.apply(lambda row: pd.to_datetime(f"{row['Data Hora Local'].date()} {row['Hora Final']}"), axis=1)

# FUNÇÃO DE AGRUPAMENTO
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
        tipo = atual["Tipo"]

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
            "Tipo": tipo,
        }
        agrupados.append(novo_bloco)
        i = j

    return pd.DataFrame(agrupados)

app = dash.Dash(__name__)
app.title = "Linha do Tempo Operacional"

app.layout = html.Div([
    html.H2("Linha do Tempo dos Operadores"),
    html.Div([
        dcc.Dropdown(
            id="equipamento-dropdown",
            options=[{"label": eq, "value": eq} for eq in sorted(df["Equipamento"].unique())],
            value=sorted(df["Equipamento"].unique())[0],
            style={"width": "350px", "margin-right": "20px"}
        ),
        dcc.Dropdown(id="operador-dropdown", style={"width": "300px", "margin-right": "20px"}),
        dcc.Dropdown(id="data-dropdown", style={"width": "200px"}),
        html.Button("Retroceder 1 dia", id="retroceder-dia", n_clicks=0,
            style={
                "margin-left":"15px", "padding":"6px 18px", "font-size":"15px",
                "background":"#222", "color":"#fff", "border-radius":"7px",
                "border":"none", "cursor":"pointer"
            }
        ),
    ], style={"display": "flex", "flex-direction": "row", "margin-bottom": "20px"}),
    html.Div(id="stats-div"),
    dcc.Graph(id="grafico-linha-tempo")
])

# Dropdown de operador depende do equipamento selecionado
@app.callback(
    Output("operador-dropdown", "options"),
    Output("operador-dropdown", "value"),
    Input("equipamento-dropdown", "value"),
)
def atualizar_operadores(equipamento):
    operadores = df[df["Equipamento"] == equipamento]["Nome"].unique()
    opcoes = [{"label": nome, "value": nome} for nome in sorted(operadores)]
    valor = sorted(operadores)[0] if len(operadores) > 0 else None
    return opcoes, valor

# Dropdown de data depende de operador e equipamento
@app.callback(
    Output("data-dropdown", "options"),
    Output("data-dropdown", "value"),
    Input("equipamento-dropdown", "value"),
    Input("operador-dropdown", "value"),
    Input("retroceder-dia", "n_clicks"),
    State("data-dropdown", "value"),
    prevent_initial_call=False
)
def atualizar_datas(equipamento, operador, n_clicks, data_atual):
    filtro = (df["Nome"] == operador) & (df["Equipamento"] == equipamento)
    datas_disponiveis = df[filtro]["Data Hora Local"].dt.date.unique()
    datas_ordenadas = sorted(datas_disponiveis)
    opcoes = [{"label": str(data), "value": str(data)} for data in datas_ordenadas]

    # Se for a primeira vez, traz o D-1
    if ctx.triggered_id in ("operador-dropdown", "equipamento-dropdown") or n_clicks == 0:
        if len(datas_ordenadas) >= 2:
            valor_padrao = str(datas_ordenadas[-2])
        elif len(datas_ordenadas) == 1:
            valor_padrao = str(datas_ordenadas[0])
        else:
            valor_padrao = None
        return opcoes, valor_padrao

    # Se clicar no botão
    if ctx.triggered_id == "retroceder-dia":
        datas_str = [str(d) for d in datas_ordenadas]
        if data_atual in datas_str:
            idx = datas_str.index(data_atual)
            novo_idx = max(idx - 1, 0)
            return opcoes, str(datas_ordenadas[novo_idx])
        else:
            return opcoes, data_atual

    return opcoes, data_atual

# Gera o gráfico
@app.callback(
    Output("grafico-linha-tempo", "figure"),
    Output("stats-div", "children"),
    Input("equipamento-dropdown", "value"),
    Input("operador-dropdown", "value"),
    Input("data-dropdown", "value")
)
def atualizar_grafico(equipamento, operador, data_str):
    if not equipamento or not operador or not data_str:
        return {}, ""
    data = pd.to_datetime(data_str).date()
    filtro = (df["Nome"] == operador) & (df["Equipamento"] == equipamento) & (df["Data Hora Local"].dt.date == data)
    dff_raw = df[filtro].copy()

    # Agrupa as operações (incluindo FINAL DE EXPEDIENTE)
    dff = agrupar_paradas(dff_raw)
    dff = dff[dff["Descrição da Operação"].str.strip().str.upper() != "FINAL DE EXPEDIENTE"]

    dff["Resumo"] = (
        "Operador: " + dff["Nome"] +
        "<br>Tipo: " + dff["Tipo"] +
        "<br>Operação: " + dff["Descrição da Operação"] +
        "<br>Início: " + dff["Inicio"].astype(str) +
        "<br>Fim: " + dff["Fim"].astype(str) +
        "<br>Duração: " + dff["Duracao Min"].round(2).astype(str) + " min"
    )

    cores = {
        "Produtiva": "#046414",        # Verde
        "Improdutiva": "#DB3B13",      # Laranja 
        "Deslocamento": "#eebf02",     # Amarelo 
        "Manobra": "#93c9f7",          # Azul bebê
        "Outro": "#111111"             # Preto
    }

    total_horas = dff["Duracao Min"].sum() / 60
    produtivo = dff[dff["Tipo"] == "Produtiva"]["Duracao Min"].sum() / 60
    improdutivo = dff[dff["Tipo"] == "Improdutiva"]["Duracao Min"].sum() / 60
    deslocamento = dff[dff["Tipo"] == "Deslocamento"]["Duracao Min"].sum() / 60
    manobra = dff[dff["Tipo"] == "Manobra"]["Duracao Min"].sum() / 60
    inicio = dff["Inicio"].min().strftime("%H:%M") if not dff.empty else "-"
    fim = dff["Fim"].max().strftime("%H:%M") if not dff.empty else "-"
    operacoes = dff["Descrição da Operação"].nunique()

    stats_html = html.Div([
        html.Span("Total horas trabalhadas: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{total_horas:.2f}h", style={"margin-right":"18px", "color":"black"}),
        html.Span("Produtivo: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{produtivo:.2f}h", style={"margin-right":"18px", "color":"#046414"}),
        html.Span("Improdutivo: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{improdutivo:.2f}h", style={"margin-right":"18px", "color":"#DB3B13"}),
        html.Span("Deslocamento: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{deslocamento:.2f}h", style={"margin-right":"18px", "color":"#eebf02"}),
        html.Span("Manobra: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{manobra:.2f}h", style={"margin-right":"18px", "color":"#93c9f7"}),
        html.Span("Início: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{inicio}", style={"margin-right":"18px", "color":"black"}),
        html.Span("Fim: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{fim}", style={"margin-right":"18px", "color":"black"}),
        html.Span("Operações diferentes: ", style={"color":"black", "font-weight":"bold"}),
        html.Span(f"{operacoes}", style={"margin-right":"5px", "color":"black"}),
    ], style={
        "background": "#fff",
        "font-size": "18px",
        "margin-bottom": "15px",
        "padding": "10px 20px",
        "border-radius": "8px",
        "box-shadow": "0 2px 12px 0 #ccc"
    })

    fig = px.timeline(
        dff,
        x_start="Inicio",
        x_end="Fim",
        y="Nome",
        color="Tipo",
        hover_name="Resumo",
        title=f"Atividades de {operador} em {data_str}",
        color_discrete_map=cores,
    )
    fig.update_layout(
        xaxis_title="Horário",
        yaxis_title="",
        height=550,
        plot_bgcolor='#181818',
        paper_bgcolor='#181818',
        font=dict(family="Montserrat, Arial", size=16, color="#e9e9e9"),
        legend=dict(
            orientation="v",
            x=1.02,
            y=1,
            bgcolor='rgba(0,0,0,0)',
            bordercolor='rgba(0,0,0,0)'
        ),
        margin=dict(l=60, r=60, t=90, b=60),
        title={
            'text': f"<b>Atividades de {operador}</b> em <span style='color:#39d353'>{data_str}</span>",
            'y':0.92,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(size=28)
        },
    )
    fig.update_traces(marker=dict(line=dict(width=2, color='white')))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(showgrid=False)

    return fig, stats_html

if __name__ == "__main__":
    print("Iniciando Dash...")
    app.run(debug=True)
