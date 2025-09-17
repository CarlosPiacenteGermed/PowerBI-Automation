import os
import pandas as pd
from dotenv import load_dotenv

def calcular_sellin_ytd(mes_limite):
    # Carregar variáveis de ambiente
    load_dotenv()
    username = os.getenv("USERNAME")
    coluna_nome = os.getenv("COLUNA_NOME")  # Ex: 'Responsável'

    # Caminho da pasta Downloads e RT
    downloads_folder = f"C:\\Users\\{username}\\Downloads"
    rt_folder = os.path.join(downloads_folder, "RT")
    os.makedirs(rt_folder, exist_ok=True)

    # Caminhos dos arquivos
    bookmark_names = [name.strip() + ".xlsx" for name in os.getenv("BOOKMARKS", "").split(",") if name.strip()]
    if len(bookmark_names) != 3:
        raise ValueError("A variável BOOKMARKS no .env deve conter exatamente 3 nomes separados por vírgula.")

    path_total = os.path.join(downloads_folder, bookmark_names[0])
    path_visitado = os.path.join(downloads_folder, bookmark_names[1])
    path_nao_visitado = os.path.join(downloads_folder, bookmark_names[2])

    # Carregar os dados
    df_total = pd.read_excel(path_total, skiprows=2, engine="openpyxl")
    df_visitado = pd.read_excel(path_visitado, skiprows=2, engine="openpyxl")
    df_nao_visitado = pd.read_excel(path_nao_visitado, skiprows=2, engine="openpyxl")

    # Padronizar datas
    for df in [df_total, df_visitado, df_nao_visitado]:
        df['Mes/Ano'] = pd.to_datetime(df['Mes/Ano'])
        df['Ano'] = df['Mes/Ano'].dt.year
        df['Mes'] = df['Mes/Ano'].dt.month

    # Função para calcular YTD
    def calcular_ytd(df):
        ytd_2024 = df[(df['Ano'] == 2024) & (df['Mes'] <= mes_limite)]['PPP Realizado'].sum()
        ytd_2025 = df[(df['Ano'] == 2025) & (df['Mes'] <= mes_limite)]['PPP Realizado'].sum()
        desv_abs = ytd_2025 - ytd_2024
        desv_perc = (desv_abs / ytd_2024 * 100) if ytd_2024 != 0 else 0
        return ytd_2024, ytd_2025, desv_abs, desv_perc

    resultado = []

    # Responsáveis únicos
    responsaveis = df_total[coluna_nome].unique()

    # Calcular YTD0 total para cada responsável
    ytd0_totais = {}
    for resp in responsaveis:
        total = df_total[df_total[coluna_nome] == resp]
        ytd_2024_total, ytd_2025_total, desv_total, perc_total = calcular_ytd(total)
        ytd0_totais[resp] = ytd_2025_total

    # Ordenar responsáveis pelo YTD0 total (decrescente)
    responsaveis_ordenados = sorted(responsaveis, key=lambda r: ytd0_totais[r], reverse=True)

    for resp in responsaveis_ordenados:
        total = df_total[df_total[coluna_nome] == resp]
        visitado = df_visitado[df_visitado[coluna_nome] == resp]
        nao_visitado = df_nao_visitado[df_nao_visitado[coluna_nome] == resp]
        
        ytd_2024_total, ytd_2025_total, desv_total, perc_total = calcular_ytd(total)
        ytd_2024_vis, ytd_2025_vis, desv_vis, perc_vis = calcular_ytd(visitado)
        ytd_2024_nao, ytd_2025_nao, desv_nao, perc_nao = calcular_ytd(nao_visitado)

        repres_vis = (ytd_2025_vis / ytd_2025_total * 100) if ytd_2025_total != 0 else 0
        repres_nao = (ytd_2025_nao / ytd_2025_total * 100) if ytd_2025_total != 0 else 0
        
        # Linha Total (Responsável)
        resultado.append({
            'Responsável': resp,
            'RCD YTD-1': round(ytd_2024_total),
            'RCD YTD0': round(ytd_2025_total),
            'DESV. ABS': round(desv_total),
            'REPRES. %': '100%',
            'DESV. %': f"{round(perc_total)}%"
        })

        # Linha Visitado
        resultado.append({
            'Responsável': 'VISITADO',
            'RCD YTD-1': round(ytd_2024_vis),
            'RCD YTD0': round(ytd_2025_vis),
            'DESV. ABS': round(desv_vis),
            'REPRES. %': f"{round(repres_vis)}%",
            'DESV. %': f"{round(perc_vis)}%"
        })

        # Linha Não Visitado
        resultado.append({
            'Responsável': 'NÃO VISITADO',
            'RCD YTD-1': round(ytd_2024_nao),
            'RCD YTD0': round(ytd_2025_nao),
            'DESV. ABS': round(desv_nao),
            'REPRES. %': f"{round(repres_nao)}%",
            'DESV. %': f"{round(perc_nao)}%"
        })
        
    return pd.DataFrame(resultado)

if __name__ == "__main__":
    load_dotenv()
    username = os.getenv("USERNAME")
    coluna_nome = os.getenv("COLUNA_NOME")
    downloads_folder = f"C:\\Users\\{username}\\Downloads"
    rt_folder = os.path.join(downloads_folder, "RT")

    # Solicita o mês limite ao usuário
    mes_limite = int(input("Informe o mês limite (número de 1 a 12): "))
    df_resultado = calcular_sellin_ytd(mes_limite)

    # Salvar em Excel
    df_resultado.to_excel(os.path.join(rt_folder, f"{coluna_nome}_resultado_sellin_ytd.xlsx"), index=False)
    print("Arquivo 'resultado_sellin_ytd.xlsx' gerado com sucesso.")