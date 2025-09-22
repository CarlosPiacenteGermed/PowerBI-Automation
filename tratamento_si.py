import os
import pandas as pd
from dotenv import load_dotenv

def calcular_sellin_ytd(df_total,df_nao_visitado,df_visitado,mes_limite):
    # Carregar variáveis de ambiente
    load_dotenv()
    coluna_nome = os.getenv("COLUNA_NOME")  # Ex: 'Responsável'

    #padronizar datas
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

def calcular_desvio_percentual(valor_atual, media_l3m):
    return round(((valor_atual - media_l3m) / media_l3m * 100), 2) if media_l3m != 0 else 0

def calcular_kpis(dados_ago, dados_l3m):
    # Calcula os valores principais
    ppp = dados_ago['PPP Realizado'].sum()
    positiv = dados_ago['Positivação'].mean()
    giro = dados_ago['Giro Médio'].mean()
    sku = dados_ago['SKU-PDV'].mean()
    preco = dados_ago['Preco Médio PPP'].mean()

    # Calcula os desvios percentuais
    desv_ppp = calcular_desvio_percentual(ppp, dados_l3m['PPP Realizado'].mean())
    desv_positiv = calcular_desvio_percentual(positiv, dados_l3m['Positivação'].mean())
    desv_giro = calcular_desvio_percentual(giro, dados_l3m['Giro Médio'].mean())
    desv_sku = calcular_desvio_percentual(sku, dados_l3m['SKU-PDV'].mean())
    desv_preco = calcular_desvio_percentual(preco, dados_l3m['Preco Médio PPP'].mean())

    kpis = {
        'DEM. PPP AGO/25': ppp,
        'DEM. PPP l3m': dados_l3m['PPP Realizado'].mean(),
        'DESV. % (PPP L3M)': desv_ppp,
        'POSITIV. AGO/25': positiv,
        'DESV. % (POSITIV L3M)': desv_positiv,
        'GIRO AGO/25': giro,
        'DESV. % (GIRO L3M)': desv_giro,
        'SKU/PDV AGO/25': sku,
        'DESV. % (SKU L3M)': desv_sku,
        'P. MÉDIO AGO/25': preco,
        'DESV. % (P. MÉDIO L3M)': desv_preco
    }
    return kpis


def calcular_tabela_por_grupo_agrupado(df_ago_total, df_l3m_total, df_ago_visitado, df_l3m_visitado, df_ago_nvisitado, df_l3m_nvisitado, coluna_grupo, coluna_responsavel):
    resultado = []

    # Lista de responsáveis únicos
    responsaveis = df_ago_total[coluna_responsavel].unique()
    for responsavel in responsaveis:
        # Adiciona linha do responsável
        resultado.append({'Responsável': responsavel})

        # Provedores desse responsável
        provedores = df_ago_total[df_ago_total[coluna_responsavel] == responsavel][coluna_grupo].unique()
        for grupo in provedores:
            # Linha do grupo (total)
            dados_ago_grupo = df_ago_total[(df_ago_total[coluna_grupo] == grupo) & (df_ago_total[coluna_responsavel] == responsavel)]
            dados_l3m_grupo = df_l3m_total[(df_l3m_total[coluna_grupo] == grupo) & (df_l3m_total[coluna_responsavel] == responsavel)]
            kpis_total = calcular_kpis(dados_ago_grupo, dados_l3m_grupo)
            kpis_total['Responsável'] = grupo
            resultado.append(kpis_total)

            # Linha VISITADO
            dados_ago_vis = df_ago_visitado[(df_ago_visitado[coluna_grupo] == grupo) & (df_ago_visitado[coluna_responsavel] == responsavel)]
            dados_l3m_vis = df_l3m_visitado[(df_l3m_visitado[coluna_grupo] == grupo) & (df_l3m_visitado[coluna_responsavel] == responsavel)]
            kpis_vis = calcular_kpis(dados_ago_vis, dados_l3m_vis)
            kpis_vis['Responsável'] = 'VISITADO'
            resultado.append(kpis_vis)

            # Linha NÃO VISITADO
            dados_ago_nao = df_ago_nvisitado[(df_ago_nvisitado[coluna_grupo] == grupo) & (df_ago_nvisitado[coluna_responsavel] == responsavel)]
            dados_l3m_nao = df_l3m_nvisitado[(df_l3m_nvisitado[coluna_grupo] == grupo) & (df_l3m_nvisitado[coluna_responsavel] == responsavel)]
            kpis_nao = calcular_kpis(dados_ago_nao, dados_l3m_nao)
            kpis_nao['Responsável'] = 'NÃO VISITADO'
            resultado.append(kpis_nao)

    df_final = pd.DataFrame(resultado)

    # Reorganiza as colunas na ordem desejada
    colunas_ordenadas = [
        'Responsável',
        'DEM. PPP AGO/25',
        'DEM. PPP l3m',
        'DESV. % (PPP L3M)',
        'POSITIV. AGO/25',
        'DESV. % (POSITIV L3M)',
        'GIRO AGO/25',
        'DESV. % (GIRO L3M)',
        'SKU/PDV AGO/25',
        'DESV. % (SKU L3M)',
        'P. MÉDIO AGO/25',
        'DESV. % (P. MÉDIO L3M)'
    ]
    df_final = df_final[colunas_ordenadas]

    return df_final
if __name__ == "__main__":
    load_dotenv()
    username = os.getenv("USERNAME")
    coluna_nome = os.getenv("COLUNA_NOME")
    coluna_grupo = os.getenv("COLUNA_GRUPO")
    downloads_folder = f"C:\\Users\\{username}\\Downloads"
    rt_folder = os.path.join(downloads_folder, "RT")

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

 # Solicita o mês limite ao usuário
    mes_limite = int(input("Informe o mês limite (número de 1 a 12): "))
    meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite-1]
    df_ago_total = df_total[(df_total['Ano'] == 2025) & (df_total['Mes'] == mes_limite)]
    df_l3m_total = df_total[(df_total['Ano'] == 2025) & (df_total['Mes'].isin(meses_l3m))]
    df_ago_visitado = df_visitado[(df_visitado['Ano'] == 2025) & (df_visitado['Mes'] == mes_limite)]
    df_l3m_visitado = df_visitado[(df_visitado['Ano'] == 2025) & (df_visitado['Mes'].isin(meses_l3m))]
    df_ago_nvisitado = df_nao_visitado[(df_nao_visitado['Ano'] == 2025) & (df_nao_visitado['Mes'] == mes_limite)]
    df_l3m_nvisitado = df_nao_visitado[(df_nao_visitado['Ano'] == 2025) & (df_nao_visitado['Mes'].isin(meses_l3m))]

    df_resultado = calcular_sellin_ytd(df_total,df_nao_visitado,df_visitado,mes_limite)
        # Salvar em Excel
    df_resultado.to_excel(os.path.join(rt_folder, f"{coluna_nome}_resultado_sellin_ytd.xlsx"), index=False)
    print("Arquivo 'resultado_sellin_ytd.xlsx' gerado com sucesso.")

    # Chamada da função KPIs por grupo na ordem correta
    df_resultado_kpis_grupo = calcular_tabela_por_grupo_agrupado(
        df_ago_total, df_l3m_total,
        df_ago_visitado, df_l3m_visitado,
        df_ago_nvisitado, df_l3m_nvisitado,
        coluna_grupo, 'CONTAS DISTY'
    )

    # Reorganiza as colunas para colocar 'Responsável' primeiro
    colunas_ordenadas = ['Responsável'] + [col for col in df_resultado_kpis_grupo.columns if col != 'Responsável']
    df_resultado_kpis_grupo = df_resultado_kpis_grupo[colunas_ordenadas]

    # Salva em Excel
    df_resultado_kpis_grupo.to_excel(os.path.join(rt_folder, f"{coluna_nome}_resultado_kpis_consolidado.xlsx"), index=False)