import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, ctx
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from datetime import timedelta, datetime
import numpy as np

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
df["Hora Final"]   = pd.to_datetime(df["Hora Final"],   format="%H:%M:%S", errors="coerce").dt.time
df["Data Hora Local"] = pd.to_datetime(df["Data Hora Local"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Hora Inicial", "Hora Final", "Data Hora Local"])

df["Hora Inicial Decimal"] = df["Hora Inicial"].apply(lambda x: x.hour + x.minute / 60)
df["Hora Final Decimal"]   = df["Hora Final"].apply(lambda x: x.hour + x.minute / 60)
df["Duracao Min"]          = df["Hora Final Decimal"] - df["Hora Inicial Decimal"]

df["Inicio"] = df.apply(lambda row: pd.to_datetime(f"{row['Data Hora Local'].date()} {row['Hora Inicial']}"), axis=1)
df["Fim"]    = df.apply(lambda row: pd.to_datetime(f"{row['Data Hora Local'].date()} {row['Hora Final']}"), axis=1)

# >>> CORREÇÃO: se cruzou meia-noite, soma 1 dia ao Fim (não troca!)
mask_cross = df["Fim"] < df["Inicio"]
df.loc[mask_cross, "Fim"] = df.loc[mask_cross, "Fim"] + pd.Timedelta(days=1)

df["Resumo"] = df.apply(lambda row: (
    f"Operador: {row['Nome']}<br>"
    f"Equipamento: {row['Equipamento']}<br>"
    f"Tipo: {row['Tipo']}<br>"
    f"Operação: {row['Descrição da Operação']}<br>"
    f"Início: {pd.to_datetime(row['Inicio']).strftime('%d/%m %H:%M')}<br>"
    f"Fim: {pd.to_datetime(row['Fim']).strftime('%d/%m %H:%M')}<br>"
    f"Duração: {round((row['Fim']-row['Inicio']).total_seconds()/60.0, 1)} min"
), axis=1)

# Classificador de parada (categorias de negócio)
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

# Agrupador de paradas — agora NÃO junta se trocar de equipamento
def agrupar_paradas(df_filtrado):
    df_filtrado = df_filtrado.sort_values(by="Inicio").reset_index(drop=True).copy()
    agrupados = []
    i = 0
    while i < len(df_filtrado):
        atual = df_filtrado.loc[i]
        inicio_bloco = atual["Inicio"]
        fim_bloco    = atual["Fim"]
        operacao     = atual["Descrição da Operação"]
        nome         = atual["Nome"]
        tipo         = atual["Tipo Parada"]
        equip        = atual["Equipamento"]

        j = i + 1
        while j < len(df_filtrado):
            proximo = df_filtrado.loc[j]
            gap = (proximo["Inicio"] - fim_bloco).total_seconds() / 60
            mesma_operacao   = proximo["Descrição da Operação"] == operacao
            mesmo_equip      = proximo["Equipamento"] == equip
            if mesma_operacao and mesmo_equip and gap <= 2:
                fim_bloco = proximo["Fim"]
                j += 1
            else:
                break

        duracao_bloco = (fim_bloco - inicio_bloco).total_seconds() / 60
        agrupados.append({
            "Nome": nome,
            "Inicio": inicio_bloco,
            "Fim": fim_bloco,
            "Descrição da Operação": operacao,
            "Duracao Min": duracao_bloco,
            "Tipo Parada": tipo,
            "Equipamento": equip,
        })
        i = j

    return pd.DataFrame(agrupados)

# ==== APP ====
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME])
app.title = "Linha do Tempo Operacional"

primeiro_nome = sorted(df["Nome"].unique())[0]
# mantemos dropdowns existentes, mas o gráfico NÃO usa mais o filtro de equipamento
primeiro_eq = sorted(df[df["Nome"] == primeiro_nome]["Equipamento"].dropna().unique().tolist())[:1]
primeiro_eq = primeiro_eq[0] if primeiro_eq else None
primeiras_datas = sorted(df[df["Nome"] == primeiro_nome]["Data Hora Local"].dt.date.unique())
data_padrao = str(primeiras_datas[-2]) if len(primeiras_datas) >= 2 else str(primeiras_datas[-1])

app.layout = html.Div(style={"backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    dbc.Container([
        html.H1("Linha do Tempo dos Operadores", className="text-center mb-4", style={"color": "#343a40", "fontWeight": "bold"}),

        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(
                    id="operador-dropdown",
                    options=[{"label": nome, "value": nome} for nome in sorted(df["Nome"].unique())],
                    value=primeiro_nome,
                    placeholder="Selecione um Operador"
                ), md=4),
                dbc.Col(dcc.Dropdown(
                    id="equipamento-dropdown",
                    options=[{"label": e, "value": e} for e in sorted(df[df["Nome"] == primeiro_nome]["Equipamento"].dropna().unique())],
                    value=primeiro_eq, placeholder="(Ignorado) Equipamento"
                ), md=4),
                dbc.Col(dcc.Dropdown(
                    id="data-dropdown",
                    options=[{"label": str(d), "value": str(d)} for d in primeiras_datas],
                    value=data_padrao, placeholder="Data inicial da janela"
                ), md=4),
            ], align="center")
        ]), className="mb-4"),

        dbc.Card(dbc.CardBody(id="stats-div"), className="mb-4"),
        dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-linha-tempo", style={"height": "580px"}))),
    ], fluid=False)
])

# Atualiza lista de equipamentos (opcional/visual), mas o gráfico não filtra por eles
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
    equipamentos = sorted(df[df["Nome"] == operador]["Equipamento"].dropna().unique())
    eq_val = equipamentos[0] if len(equipamentos) else None
    datas = sorted(df[df["Nome"] == operador]["Data Hora Local"].dt.date.unique())
    opcoes_datas = [{"label": str(d), "value": str(d)} for d in datas]
    data_val = str(datas[-2]) if len(datas) >= 2 else (str(datas[-1]) if len(datas) else None)
    return [{"label": e, "value": e} for e in equipamentos], eq_val, opcoes_datas, data_val

# ======= FUNÇÕES AUXILIARES PARA SOBREPOSIÇÕES =======
def blocos_equipamento(dff_raw):
    """Colapsa blocos contínuos por Equipamento (independe de operação)."""
    d = dff_raw.sort_values("Inicio").reset_index(drop=True)
    out = []
    i = 0
    while i < len(d):
        atual = d.loc[i]
        equip = atual["Equipamento"]
        ini   = atual["Inicio"]
        fim   = atual["Fim"]
        j = i + 1
        while j < len(d):
            prox = d.loc[j]
            gap = (prox["Inicio"] - fim).total_seconds()/60.0
            if prox["Equipamento"] == equip and gap <= 2:
                fim = max(fim, prox["Fim"])
                j += 1
            else:
                break
        out.append({"Equipamento": equip, "Inicio": ini, "Fim": fim})
        i = j
    return pd.DataFrame(out)

def add_divisores_de_dia(fig, tmin, tmax):
    """Linhas verticais em cada 00:00 e anotação da data."""
    if pd.isna(tmin) or pd.isna(tmax):
        return
    start_day = pd.to_datetime(tmin.date())
    end_day   = pd.to_datetime(tmax.date()) + pd.Timedelta(days=1)
    cur = start_day
    k = 0
    while cur <= end_day:
        fig.add_vline(x=cur, line_width=1, line_dash="dot", line_color="#9aa0a6")
        # anotação no topo do dia (meio-dia)
        mid = cur + pd.Timedelta(hours=12)
        fig.add_annotation(x=mid, y=1.04, yref="paper", text=cur.strftime("Dia %d/%m"),
                           showarrow=False, font=dict(size=11, color="#aeb5bd"))
        cur += pd.Timedelta(days=1); k += 1

# ======= CALLBACK PRINCIPAL =======
@app.callback(
    Output("grafico-linha-tempo", "figure"),
    Output("stats-div", "children"),
    Input("operador-dropdown", "value"),
    Input("equipamento-dropdown", "value"),  # <<< ignorado no filtro
    Input("data-dropdown", "value")
)
def atualizar_grafico(operador, _equip_ignorado, data_str):
    if not operador:
        return {}, html.Div("Selecione um operador.", className="text-center text-muted p-4")

    # >>> NÃO filtra por equipamento nem por data: pega TODA a linha do tempo do operador
    base = df[df["Nome"] == operador].copy()
    if base.empty:
        return {}, html.Div("Nenhum dado encontrado.", className="text-center text-muted p-4")

    # Agrupa e remove 'FINAL/FIM DE EXPEDIENTE'
    dff = agrupar_paradas(base)
    exp_set = {"FINAL DE EXPEDIENTE", "FIM DE EXPEDIENTE"}
    dff = dff[~dff["Descrição da Operação"].str.upper().str.strip().isin(exp_set)]
    if dff.empty:
        return {}, html.Div("Nenhum dado útil para visualizar.", className="text-center text-muted p-4")

    # Stats simples (sobre tudo)
    hora_inicio = dff["Inicio"].min().strftime("%d/%m %H:%M")
    hora_fim    = dff["Fim"].max().strftime("%d/%m %H:%M")
    def soma_h(tipo): return float(dff[dff["Tipo Parada"] == tipo]["Duracao Min"].sum())/60.0
    stats = {
        "Efetivo": soma_h("Efetivo"),
        "Parada Gerenciável": soma_h("Parada Gerenciável"),
        "Parada Mecânica": soma_h("Parada Mecânica"),
        "Parada Improdutiva": soma_h("Parada Improdutiva"),
        "Parada Essencial": soma_h("Parada Essencial"),
        "Deslocamento": soma_h("Deslocamento"),
        "Manobra": soma_h("Manobra"),
    }
    total = sum(stats.values())

    def create_stat_card(title, value, color):
        return dbc.Col(dbc.Card(dbc.CardBody([
            html.H4(value, style={"color": color, "fontWeight": "bold"}),
            html.P(title, className="text-muted")
        ]), className="text-center shadow-sm"), md=2, className="mb-2")

    stats_html = dbc.Row([
        create_stat_card("Início", hora_inicio, "#6c757d"),
        create_stat_card("Fim",    hora_fim,    "#6c757d"),
        create_stat_card("Total Horas", f"{total:.2f}h", "#343a40"),
        create_stat_card("Efetivo", f"{stats['Efetivo']:.2f}h", "#046414"),
        create_stat_card("Parada Gerenciável", f"{stats['Parada Gerenciável']:.2f}h", "#FF9393"),
        create_stat_card("Parada Mecânica", f"{stats['Parada Mecânica']:.2f}h", "#A52657"),
        create_stat_card("Parada Improdutiva", f"{stats['Parada Improdutiva']:.2f}h", "#FF0000"),
        create_stat_card("Parada Essencial", f"{stats['Parada Essencial']:.2f}h", "#0026FF"),
    ], justify="center")

    # Hover com equipamento
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
        title=f"<b>Atividades de {operador}</b> — role/panar para ver outros dias",
        plot_bgcolor='#181818', paper_bgcolor='#181818',
        font=dict(color="#e9e9e9"), xaxis_title="Horário", yaxis_title="",
        margin=dict(l=40, r=40, t=80, b=60), height=580,
        legend=dict(orientation="v", x=1.02, y=1)
    )
    fig.update_traces(marker=dict(line=dict(width=1, color='white')))
    fig.update_yaxes(autorange="reversed")

    # ===== Sobreposição por EQUIPAMENTO (tipo a faixa de noite) =====
    blocos_eq = blocos_equipamento(base)
    # paleta para equipamentos
    palette = px.colors.qualitative.Pastel + px.colors.qualitative.Set3 + px.colors.qualitative.Light24
    eqs = blocos_eq["Equipamento"].fillna("Sem equipamento").unique().tolist()
    cmap = {e: palette[i % len(palette)] for i, e in enumerate(eqs)}

    for _, row in blocos_eq.iterrows():
        fig.add_vrect(
            x0=row["Inicio"], x1=row["Fim"],
            fillcolor=cmap.get(row["Equipamento"], "#3a2e5f"),
            opacity=0.12, layer="below", line_width=0
        )
        # anotação do equipamento no topo do bloco
        xm = row["Inicio"] + (row["Fim"] - row["Inicio"]) / 2
        fig.add_annotation(
            x=xm, y=1.02, yref="paper",
            text=str(row["Equipamento"]),
            showarrow=False, font=dict(size=10, color="#aab2bd"), align="center"
        )

    # ===== Divisores de dia (meia-noite) e rótulos "Dia dd/mm" =====
    tmin, tmax = dff["Inicio"].min(), dff["Fim"].max()
    def add_divisores_de_dia(fig, tmin, tmax):
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
    add_divisores_de_dia(fig, tmin, tmax)

    # ===== Janela inicial: 1 dia baseado no dropdown de data (mas pode panar depois) =====
    if data_str:
        dia0 = pd.to_datetime(data_str)
        x0 = pd.Timestamp(dia0.date())
        x1 = x0 + pd.Timedelta(days=1)
        fig.update_xaxes(range=[x0, x1])

    return fig, stats_html

if __name__ == "__main__":
    app.run_server(debug=True)
