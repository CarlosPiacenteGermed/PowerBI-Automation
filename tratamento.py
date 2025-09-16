import pandas as pd
import os
from dotenv import load_dotenv

def main():
    # Carregar variáveis de ambiente
    load_dotenv()
    username = os.getenv("USERNAME")
    coluna_nome = os.getenv("COLUNA_NOME")
    coluna_grupo = os.getenv("COLUNA_GRUPO")
   
    # Caminho da pasta Downloads e RT
    downloads_folder = f"C:\\Users\\{username}\\Downloads"
    rt_folder = os.path.join(downloads_folder, "RT")
    os.makedirs(rt_folder, exist_ok=True)  # Cria a pasta RT se não existir

    # Caminhos dos arquivos
    bookmark_names = [name.strip() + ".xlsx" for name in os.getenv("BOOKMARKS", "").split(",") if name.strip()]
    if len(bookmark_names) != 3:
        raise ValueError("A variável BOOKMARKS no .env deve conter exatamente 3 nomes separados por vírgula.")

    path_coord = os.path.join(downloads_folder, bookmark_names[0])
    path_total = os.path.join(downloads_folder, bookmark_names[1])
    path_grupo = os.path.join(downloads_folder, bookmark_names[2])
    # Carregar as bases
    df_coord = pd.read_excel(path_coord, skiprows=2, engine="openpyxl")
    df_total = pd.read_excel(path_total, skiprows=2, engine="openpyxl")
    df_grupo = pd.read_excel(path_grupo, skiprows=2, engine="openpyxl")

    # Padronizar datas
    for df in [df_coord, df_total, df_grupo]:
        df['Mes/Ano'] = pd.to_datetime(df['Mes/Ano'])
        df['Ano'] = df['Mes/Ano'].dt.year
        df['Mes'] = df['Mes/Ano'].dt.month

    df_total[coluna_nome] = 'Total'

    # ------------------ TABELA PRINCIPAL ------------------ #
    def calcular_indicadores_gerais(df, coluna_nome=coluna_nome):
        resultado = []
        for nome in df[coluna_nome].unique():
            dados = df[df[coluna_nome] == nome]

            ytd_1 = dados[dados['Ano'] == 2024]['PPP Realizado'].sum()
            ytdo = dados[(dados['Ano'] == 2025) & (dados['Mes'] <= 8)]['PPP Realizado'].sum()
            desv_abs = ytdo - ytd_1
            desv_perc = ((desv_abs / ytd_1) * 100) if ytd_1 != 0 else 0
            nome_final = 'Total' if nome == 'Total' else nome.split()[0]
            resultado.append({
                '': nome_final if coluna_nome == 'CONTAS REDE' or 'CONTAS ASSOC.' else 'Total',
                'RCD YTD-1': round(ytd_1),
                'RCD YTDO': round(ytdo),
                'DESV. ABS': round(desv_abs),
                'DESV. %': round(desv_perc, 1),
            })
        return pd.DataFrame(resultado)

    # Calcular e ordenar
    tabela_coord = calcular_indicadores_gerais(df_coord, coluna_nome)
    tabela_total = calcular_indicadores_gerais(df_total, coluna_nome)
    tabela_final = pd.concat([tabela_coord, tabela_total], ignore_index=True)

    # Colocar Total Coord. por último
    linha_total = tabela_final[tabela_final[''] == 'Total']
    tabela_final = tabela_final[tabela_final[''] != 'Total']
    tabela_final = tabela_final.sort_values(by='RCD YTDO', ascending=False)
    tabela_final = pd.concat([tabela_final, linha_total], ignore_index=True)

    # ------------------ TABELA DA IMAGEM ------------------ #
    def calcular_tabela_l3m_gerais(df):
        df_l3m = df[(df['Ano'] == 2025) & (df['Mes'].isin([6, 7, 8]))]
        df_ago = df[(df['Ano'] == 2025) & (df['Mes'] == 8)]

        resultado = []
        for nome in df_ago[coluna_nome].unique():
            dados_ago = df_ago[df_ago[coluna_nome] == nome]
            dados_l3m = df_l3m[df_l3m[coluna_nome] == nome]

            def calc_desv_percentual(valor_atual, media_l3m):
                return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m != 0 else 0

            ppp_ago = dados_ago['PPP Realizado'].sum()
            ppp_l3m = dados_l3m['PPP Realizado'].mean()
            desv_ppp = calc_desv_percentual(ppp_ago, ppp_l3m)

            positiv_ago = dados_ago['Positivação'].mean()
            positiv_l3m = dados_l3m['Positivação'].mean()
            desv_positiv = calc_desv_percentual(positiv_ago, positiv_l3m)

            giro_ago = dados_ago['Giro Médio'].mean()
            giro_l3m = dados_l3m['Giro Médio'].mean()
            desv_giro = calc_desv_percentual(giro_ago, giro_l3m)

            sku_ago = dados_ago['SKU-PDV'].mean()
            sku_l3m = dados_l3m['SKU-PDV'].mean()
            desv_sku = calc_desv_percentual(sku_ago, sku_l3m)

            preco_ago = dados_ago['Preco Médio PPP'].mean()
            preco_l3m = dados_l3m['Preco Médio PPP'].mean()
            desv_preco = calc_desv_percentual(preco_ago, preco_l3m)
            nome_final = 'Total' if nome == 'Total' else nome.split()[0]
            resultado.append({
                '': nome_final,
                'DEM. PPP AGO/25': round(ppp_ago),
                'DESV. % (PPP L3M)': round(desv_ppp, 1),
                'POSITIV. AGO/25': round(positiv_ago),
                'DESV. % (POSITIV L3M)': round(desv_positiv, 1),
                'GIRO AGO/25': round(giro_ago, 1),
                'DESV. % (GIRO L3M)': round(desv_giro, 1),
                'SKU/PDV AGO/25': round(sku_ago, 1),
                'DESV. % (SKU L3M)': round(desv_sku, 1),
                'P. MÉDIO AGO/25': round(preco_ago, 2),
                'DESV. % (P. MÉDIO L3M)': round(desv_preco, 1),
            })
        return pd.DataFrame(resultado)

    # Gerar segunda tabela
    tabela_l3m = calcular_tabela_l3m_gerais(df_coord)
    tabela_l3m_total = calcular_tabela_l3m_gerais(df_total)
    tabela_l3m_final = pd.concat([tabela_l3m, tabela_l3m_total], ignore_index=True)
    # Reordenar colocando Total Coord. por último
    linha_total = tabela_l3m_final[tabela_l3m_final[''] == 'Total']
    tabela_l3m_final = tabela_l3m_final[tabela_l3m_final[''] != 'Total']
    tabela_l3m_final = tabela_l3m_final.sort_values(by='DEM. PPP AGO/25', ascending=False)
    tabela_l3m_final = pd.concat([tabela_l3m_final, linha_total], ignore_index=True)

    # Filtrar dados de agosto/2025 e últimos 3 meses
    df_ago = df_grupo[(df_grupo['Ano'] == 2025) & (df_grupo['Mes'] == 8)]
    df_l3m = df_grupo[(df_grupo['Ano'] == 2025) & (df_grupo['Mes'].isin([6, 7, 8]))]

    # Função para calcular indicadores gerais por grupo econômico (tabela tipo "principal")
    def calcular_indicadores_gerais_grupo(df, coluna_nome = coluna_nome):
        resultado = []
        for grupo in df[coluna_nome].unique():
            dados = df[df[coluna_nome] == grupo]
            ytd_1 = dados[dados['Ano'] == 2024]['PPP Realizado'].sum()
            ytdo = dados[(dados['Ano'] == 2025) & (dados['Mes'] <= 8)]['PPP Realizado'].sum()
            desv_abs = ytdo - ytd_1
            desv_perc = ((desv_abs / ytd_1) * 100) if ytd_1 != 0 else 0
            resultado.append({
                '': grupo,
                'RCD YTD-1': round(ytd_1),
                'RCD YTDO': round(ytdo),
                'DESV. ABS': round(desv_abs),
                'DESV. %': round(desv_perc, 1),
            })
        df_resultado = pd.DataFrame(resultado)
        # Ordena por RCD YTDO e pega os 7 maiores
        df_top7 = df_resultado.sort_values(by='RCD YTDO', ascending=False).head(7)
        # Soma total geral
        total_geral = df_resultado[['RCD YTD-1', 'RCD YTDO', 'DESV. ABS']].sum()
        total_geral['DESV. %'] = ((total_geral['DESV. ABS'] / total_geral['RCD YTD-1']) * 100) if total_geral['RCD YTD-1'] != 0 else 0
        total_row = {'': 'TOTAL', **{col: round(total_geral[col], 2) for col in total_geral.index}}
        # Soma dos TOP7
        top7_sum = df_top7[['RCD YTD-1', 'RCD YTDO', 'DESV. ABS']].sum()
        top7_sum['DESV. %'] = ((top7_sum['DESV. ABS'] / top7_sum['RCD YTD-1']) * 100) if top7_sum['RCD YTD-1'] != 0 else 0
        outros_row = {'': 'OUTROS', **{col: round(total_geral[col] - top7_sum[col], 2) for col in ['RCD YTD-1', 'RCD YTDO', 'DESV. ABS']}}
        outros_row['DESV. %'] = ((outros_row['DESV. ABS'] / outros_row['RCD YTD-1']) * 100) if outros_row['RCD YTD-1'] != 0 else 0
        # Junta tudo
        df_final = pd.concat([df_top7, pd.DataFrame([outros_row, total_row])], ignore_index=True)
        return df_final
    def calcular_tabela_por_grupo(dados_ago, dados_l3m, grupos_top7):
        resultado = []
        for grupo in dados_ago[coluna_grupo].unique():
            dados_ago_grupo = dados_ago[dados_ago[coluna_grupo] == grupo]
            dados_l3m_grupo = dados_l3m[dados_l3m[coluna_grupo] == grupo]

            def calc_desv_percentual(valor_atual, media_l3m):
                return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m != 0 else 0

            ppp_ago = dados_ago_grupo['PPP Realizado'].sum()
            ppp_l3m = dados_l3m_grupo['PPP Realizado'].mean()
            desv_ppp = calc_desv_percentual(ppp_ago, ppp_l3m)

            positiv_ago = dados_ago_grupo['Positivação'].mean()
            positiv_l3m = dados_l3m_grupo['Positivação'].mean()
            desv_positiv = calc_desv_percentual(positiv_ago, positiv_l3m)

            giro_ago = dados_ago_grupo['Giro Médio'].mean()
            giro_l3m = dados_l3m_grupo['Giro Médio'].mean()
            desv_giro = calc_desv_percentual(giro_ago, giro_l3m)

            sku_ago = dados_ago_grupo['SKU-PDV'].mean()
            sku_l3m = dados_l3m_grupo['SKU-PDV'].mean()
            desv_sku = calc_desv_percentual(sku_ago, sku_l3m)

            preco_ago = dados_ago_grupo['Preco Médio PPP'].mean()
            preco_l3m = dados_l3m_grupo['Preco Médio PPP'].mean()
            desv_preco = calc_desv_percentual(preco_ago, preco_l3m)

            resultado.append({
                '': grupo,
                'DEM. PPP AGO/25': round(ppp_ago),
                'DESV. % (PPP L3M)': round(desv_ppp, 1),
                'POSITIV. AGO/25': round(positiv_ago),
                'DESV. % (POSITIV L3M)': round(desv_positiv, 1),
                'GIRO AGO/25': round(giro_ago, 1),
                'DESV. % (GIRO L3M)': round(desv_giro, 1),
                'SKU/PDV AGO/25': round(sku_ago, 1),
                'DESV. % (SKU L3M)': round(desv_sku, 1),
                'P. MÉDIO AGO/25': round(preco_ago, 2),
                'DESV. % (P. MÉDIO L3M)': round(desv_preco, 1),
            })
        df_top7 = pd.DataFrame(resultado)
        # OUTROS: média dos grupos que não estão no TOP7
        grupos_outros = set(dados_ago[coluna_grupo].unique()) - set(grupos_top7)
        if grupos_outros:
            outros_rows = []
            for grupo in grupos_outros:
                dados_ago_grupo = dados_ago[dados_ago[coluna_grupo] == grupo]
                dados_l3m_grupo = dados_l3m[dados_l3m[coluna_grupo] == grupo]
                def calc_desv_percentual(valor_atual, media_l3m):
                    return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m != 0 else 0
                outros_rows.append({
                    'DEM. PPP AGO/25': dados_ago_grupo['PPP Realizado'].sum(),
                    'DESV. % (PPP L3M)': calc_desv_percentual(
                        dados_ago_grupo['PPP Realizado'].sum(),
                        dados_l3m_grupo['PPP Realizado'].mean()
                    ),
                    'POSITIV. AGO/25': dados_ago_grupo['Positivação'].mean(),
                    'DESV. % (POSITIV L3M)': calc_desv_percentual(
                        dados_ago_grupo['Positivação'].mean(),
                        dados_l3m_grupo['Positivação'].mean()
                    ),
                    'GIRO AGO/25': dados_ago_grupo['Giro Médio'].mean(),
                    'DESV. % (GIRO L3M)': calc_desv_percentual(
                        dados_ago_grupo['Giro Médio'].mean(),
                        dados_l3m_grupo['Giro Médio'].mean()
                    ),
                    'SKU/PDV AGO/25': dados_ago_grupo['SKU-PDV'].mean(),
                    'DESV. % (SKU L3M)': calc_desv_percentual(
                        dados_ago_grupo['SKU-PDV'].mean(),
                        dados_l3m_grupo['SKU-PDV'].mean()
                    ),
                    'P. MÉDIO AGO/25': dados_ago_grupo['Preco Médio PPP'].mean(),
                    'DESV. % (P. MÉDIO L3M)': calc_desv_percentual(
                        dados_ago_grupo['Preco Médio PPP'].mean(),
                        dados_l3m_grupo['Preco Médio PPP'].mean()
                    ),
                })
            # Média dos outros
            outros_media = {k: round(pd.DataFrame(outros_rows)[k].mean(), 2) for k in outros_rows[0].keys()}
            outros_media[''] = 'OUTROS'
            df_top7 = pd.concat([df_top7, pd.DataFrame([outros_media])], ignore_index=True)
        # TOTAL: soma/média dos KPIs
        total_row = {
            '': 'TOTAL',
            'DEM. PPP AGO/25': df_top7['DEM. PPP AGO/25'].sum(),
            'DESV. % (PPP L3M)': round(df_top7['DESV. % (PPP L3M)'].mean(), 2),
            'POSITIV. AGO/25': round(df_top7['POSITIV. AGO/25'].mean(), 2),
            'DESV. % (POSITIV L3M)': round(df_top7['DESV. % (POSITIV L3M)'].mean(), 2),
            'GIRO AGO/25': round(df_top7['GIRO AGO/25'].mean(), 2),
            'DESV. % (GIRO L3M)': round(df_top7['DESV. % (GIRO L3M)'].mean(), 2),
            'SKU/PDV AGO/25': round(df_top7['SKU/PDV AGO/25'].mean(), 2),
            'DESV. % (SKU L3M)': round(df_top7['DESV. % (SKU L3M)'].mean(), 2),
            'P. MÉDIO AGO/25': round(df_top7['P. MÉDIO AGO/25'].mean(), 2),
            'DESV. % (P. MÉDIO L3M)': round(df_top7['DESV. % (P. MÉDIO L3M)'].mean(), 2),
        }
        df_top7 = pd.concat([df_top7, pd.DataFrame([total_row])], ignore_index=True)
        return df_top7
    #Gerar tabelas por coordenador
    for coordenador in df_ago[coluna_nome].unique():
        dados_coord = df_grupo[df_grupo[coluna_nome] == coordenador]
        dados_ago_coord = df_ago[df_ago[coluna_nome] == coordenador]
        dados_l3m_coord = df_l3m[df_l3m[coluna_nome] == coordenador]
        # Tabela tipo "principal" (YTD)
        tabela_ytd = calcular_indicadores_gerais_grupo(dados_coord)
        grupos_top7 = tabela_ytd[(tabela_ytd[''] != 'OUTROS') & (tabela_ytd[''] != 'TOTAL')][''].tolist()
        # Tabela tipo "imagem" (KPI) na mesma ordem do YTD
        tabela_kpi = calcular_tabela_por_grupo(dados_ago_coord, dados_l3m_coord, grupos_top7)
        primeiro_nome = coordenador.split()[0]
        tabela_kpi.to_excel(os.path.join(rt_folder, f"tabela_top7_KPI_{primeiro_nome}.xlsx"), index=False)
        tabela_ytd.to_excel(os.path.join(rt_folder, f"tabela_top7_YTD_{primeiro_nome}.xlsx"), index=False)

    # Exportar tabelas principais
    tabela_final.to_excel(os.path.join(rt_folder, "tabela_ytd.xlsx"), index=False)
    tabela_l3m_final.to_excel(os.path.join(rt_folder, "tabela_l3m_agosto.xlsx"), index=False)
    print(f"Arquivos salvos na pasta: {rt_folder}")
if __name__ == "__main__":
    main()