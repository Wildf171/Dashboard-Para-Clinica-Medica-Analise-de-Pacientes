from flask import Flask, render_template, Response
import pandas as pd
import io

app = Flask(__name__)

# --- FUNÇÃO AUXILIAR PARA CARREGAR O DATAFRAME ---
def get_dataframe():
    # 1. Carregar planilha sem cabeçalho
    df = pd.read_excel("dados.xlsx", header=None, dtype=str, engine='openpyxl')

    # 2. Definir colunas
    df.columns = [
        "data_atendimento", "paciente", "medico_solicitante", "exame",
        "origem_paciente", "prioridade_atendimento", "cid", "total_item", "total_liquido"
    ]

    # 3. Limpeza
    df["data_atendimento"] = pd.to_datetime(df["data_atendimento"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["data_atendimento"])

    # Ajuste de moeda
    df["total_liquido"] = df["total_liquido"].astype(str).str.replace(',', '.')
    df["total_liquido"] = pd.to_numeric(df["total_liquido"], errors="coerce").fillna(0)

    # Padronização de Texto
    cols_texto = ['medico_solicitante', 'origem_paciente', 'prioridade_atendimento', 'cid', 'exame']
    for col in cols_texto:
        df[col] = df[col].fillna("NÃO INFORMADO").astype(str).str.upper()
        
    return df

def carregar_dados_dashboard():
    df = get_dataframe()

    # --- PROCESSAMENTO (TOP 10 VISUAL) ---
    medico_counts = df['medico_solicitante'].value_counts().head(10)
    exame_counts = df['exame'].value_counts().head(10)
    cid_counts = df['cid'].value_counts().head(10) # Top 10 para o gráfico
    
    origem_counts = df['origem_paciente'].value_counts()
    prioridade_counts = df['prioridade_atendimento'].value_counts()
    
    # Timeline Mensal
    timeline_resampled = df.set_index('data_atendimento').resample('ME').size()
    timeline_labels = timeline_resampled.index.strftime('%m/%Y').tolist()
    timeline_values = timeline_resampled.values.tolist()

    total_faturamento = df['total_liquido'].sum()

    dados_dashboard = {
        'medicos': { 'labels': medico_counts.index.tolist(), 'values': medico_counts.values.tolist() },
        'exames':  { 'labels': exame_counts.index.tolist(),  'values': exame_counts.values.tolist() },
        'origem':  { 'labels': origem_counts.index.tolist(), 'values': origem_counts.values.tolist() },
        'prioridade': { 'labels': prioridade_counts.index.tolist(), 'values': prioridade_counts.values.tolist() },
        'cids':    { 'labels': cid_counts.index.tolist(),    'values': cid_counts.values.tolist() },
        'timeline': { 'labels': timeline_labels, 'values': timeline_values },
        'kpis': {
            'total_exames': len(df),
            'faturamento': f"R$ {total_faturamento:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        }
    }
    return dados_dashboard

@app.route('/')
def index():
    try:
        dados = carregar_dados_dashboard()
        return render_template('dashboard.html', dados=dados)
    except Exception as e:
        return f"<div style='color:red'><h1>Erro:</h1><p>{str(e)}</p></div>"

# --- ROTA PARA DOWNLOAD (AGORA COM CID) ---
@app.route('/download/<tipo>')
def download_csv(tipo):
    df = get_dataframe()
    
    if tipo == 'medicos':
        tabela = df['medico_solicitante'].value_counts().reset_index()
        tabela.columns = ['Médico Solicitante', 'Qtd Exames']
        filename = "relatorio_medicos.csv"
        
    elif tipo == 'exames':
        tabela = df['exame'].value_counts().reset_index()
        tabela.columns = ['Exame', 'Qtd Realizada']
        filename = "relatorio_exames.csv"
        
    elif tipo == 'cids':
        # NOVA LÓGICA AQUI: Gera tabela completa de CIDs
        tabela = df['cid'].value_counts().reset_index()
        tabela.columns = ['Diagnóstico (CID)', 'Ocorrências']
        filename = "relatorio_cids.csv"
        
    elif tipo == 'geral':
        tabela = df
        filename = "base_completa.csv"
    else:
        return "Tipo inválido"

    # Gera CSV com separador ';' para abrir fácil no Excel brasileiro
    csv_data = tabela.to_csv(sep=';', index=False, encoding='utf-8-sig')

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-disposition": f"attachment; filename={filename}"}
    )

if __name__ == '__main__':
    app.run(debug=True)