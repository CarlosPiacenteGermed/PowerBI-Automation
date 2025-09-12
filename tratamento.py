import pandas as pd
import os
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()
username = os.getenv("USERNAME")

# Caminhos dos arquivos
path_coord = f"C:\\Users\\{username}\\Downloads\\Material Coordenador Contas RT.xlsx"
path_total = f"C:\\Users\\{username}\\Downloads\\Material Coordenador Geral RT.xlsx"
path_grupo = f"C:\\Users\\{username}\\Downloads\\Material Coordenador RT.xlsx"

# Carregar as bases
df_coord = pd.read_excel(path_coord, skiprows=2, engine="openpyxl")
df_total = pd.read_excel(path_total, skiprows=2, engine="openpyxl")
df_grupo = pd.read_excel(path_grupo, skiprows=2, engine="openpyxl")

# Padronizar datas
for df in [df_coord, df_total, df_grupo]:
    df['Mes/Ano'] = pd.to_datetime(df['Mes/Ano'])
    df['Ano'] = df['Mes/Ano'].dt.year
    df['Mes'] = df['Mes/Ano'].dt.month

# Renomear DIEGO VICENTE RODRIGUES como 'Total Coord.'
df_total['CONTAS REDE'] = df_total['REGIONAL REDE'].replace('DIEGO VICENTE RODRIGUES', 'Total Coord.')


# ------------------ TABELA PRINCIPAL ------------------ #
def calcular_indicadores_gerais(df, coluna_nome='CONTAS REDE'):
    resultado = []
    for nome in df[coluna_nome].unique():
        dados = df[df[coluna_nome] == nome]

        ytd_1 = dados[dados['Ano'] == 2024]['PPP Realizado'].sum()
        ytdo = dados[(dados['Ano'] == 2025) & (dados['Mes'] <= 8)]['PPP Realizado'].sum()
        desv_abs = ytdo - ytd_1
        desv_perc = ((desv_abs / ytd_1) * 100) if ytd_1 != 0 else 0
        nome_final = 'Total Coord.' if nome == 'Total Coord.' else nome.split()[0]
        resultado.append({
            '': nome_final if coluna_nome == 'CONTAS REDE' else 'Total Coord.',
            'RCD YTD-1': round(ytd_1, 2),
            'RCD YTDO': round(ytdo, 2),
            'DESV. ABS': round(desv_abs, 2),
            'DESV. %': round(desv_perc, 2),
        })
    return pd.DataFrame(resultado)

# Calcular e ordenar
tabela_coord = calcular_indicadores_gerais(df_coord, 'CONTAS REDE')
tabela_total = calcular_indicadores_gerais(df_total, 'REGIONAL REDE')
tabela_final = pd.concat([tabela_coord, tabela_total], ignore_index=True)

# Colocar Total Coord. por último
linha_total = tabela_final[tabela_final[''] == 'Total Coord.']
tabela_final = tabela_final[tabela_final[''] != 'Total Coord.']
tabela_final = tabela_final.sort_values(by='RCD YTDO', ascending=False)
tabela_final = pd.concat([tabela_final, linha_total], ignore_index=True)

# ------------------ TABELA DA IMAGEM ------------------ #
def calcular_tabela_l3m_gerais(df):
    df_l3m = df[(df['Ano'] == 2025) & (df['Mes'].isin([6, 7, 8]))]
    df_ago = df[(df['Ano'] == 2025) & (df['Mes'] == 8)]

    resultado = []
    for nome in df_ago['CONTAS REDE'].unique():
        dados_ago = df_ago[df_ago['CONTAS REDE'] == nome]
        dados_l3m = df_l3m[df_l3m['CONTAS REDE'] == nome]

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
        nome_final = 'Total Coord.' if nome == 'Total Coord.' else nome.split()[0]
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
linha_total = tabela_l3m_final[tabela_l3m_final[''] == 'Total Coord.']
tabela_l3m_final = tabela_l3m_final[tabela_l3m_final[''] != 'Total Coord.']
tabela_l3m_final = tabela_l3m_final.sort_values(by='DEM. PPP AGO/25', ascending=False)
tabela_l3m_final = pd.concat([tabela_l3m_final, linha_total], ignore_index=True)


# Exibir ou salvar
# print("TABELA PRINCIPAL:")
# print(tabela_final)

# print("TABELA DA IMAGEM:")
# print(tabela_l3m_final)

# # Exportar se quiser
# tabela_final.to_excel("tabela_principal.xlsx", index=False)
# tabela_l3m_final.to_excel("tabela_l3m_agosto.xlsx", index=False)
