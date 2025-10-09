import pandas as pd
import os
from dotenv import load_dotenv
from openpyxl import load_workbook

def main(mes_limite):
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

    meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite-1]
    # ------------------ TABELA PRINCIPAL ------------------ #
    def calcular_indicadores_gerais(df, coluna_nome=coluna_nome):
        meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite-1]
        df_l3m = df[(df['Ano'] == 2025) & (df['Mes'].isin(meses_l3m))]
        df_ago = df[(df['Ano'] == 2025) & (df['Mes'] == mes_limite)]
        
        resultado = []
        for nome in df[coluna_nome].unique():
            dados = df_ago[df_ago[coluna_nome] == nome]
            dados_l3 = df_l3m[df_l3m[coluna_nome] == nome]

            ytd_1 = dados_l3[(dados_l3['Ano'] == 2025) & (dados_l3['Mes'].isin(meses_l3m))]['PPP Realizado'].fillna(0).sum()
            ytd_1=ytd_1/3
            ytdo = dados[(dados['Ano'] == 2025) & (dados['Mes'] == mes_limite)]['PPP Realizado'].fillna(0).sum()
            desv_abs = ytdo - ytd_1
            desv_perc = ((desv_abs / ytd_1) * 100) if ytd_1 != 0 else 0
            nome_final = 'Total' if nome == 'Total' else nome.split()[0]
            resultado.append({
                '': nome_final if coluna_nome == 'CONTAS REDE' or 'CONTAS ASSOC.' else 'Total',
                'RCD L3M': round(ytd_1),
                'RCD MES': round(ytdo),
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
    tabela_final = tabela_final.sort_values(by='RCD MES', ascending=False)
    tabela_final = pd.concat([tabela_final, linha_total], ignore_index=True)

    # ------------------ TABELA DA IMAGEM ------------------ #
    def calcular_tabela_l3m_gerais(df, coluna_nome=coluna_nome):
        import os
        from dotenv import load_dotenv
        load_dotenv()
        territorio_path = os.getenv("TERRITORIO_PATH" )
        path_territorio = os.path.join(downloads_folder, territorio_path)
        def calc_desv_percentual(valor_atual, media_l3m):
             return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m not in [0, None, float('nan')] and media_l3m != 0 else 0
        territorio = False
        if os.path.isfile(path_territorio):
            territorio = True
            # Importa a tabela de território
            df_territorio = pd.read_excel(path_territorio, skiprows=2,engine="openpyxl")
            
            resultado = []
            meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite-1]
            df_l3m = df[(df['Ano'] == 2025) & (df['Mes'].isin(meses_l3m))]
            df_ago = df[(df['Ano'] == 2025) & (df['Mes'] == mes_limite)]
            for nome in df_ago[coluna_nome].unique():
                 dados_ago = df_ago[df_ago[coluna_nome] == nome]
                 dados_l3m = df_l3m[df_l3m[coluna_nome] == nome]
                 dados_grupo = df_territorio[df_territorio.iloc[:, 0] == nome]
                 def calc_desv_percentual(valor_atual, media_l3m):
                    return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m != 0 else 0

                 ppp_ago = dados_grupo["DMD | OL"].fillna(0).sum()
                 ppp_l3m = dados_grupo["L3M"].fillna(0).mean()
                 desv_ppp = calc_desv_percentual(ppp_ago, ppp_l3m)

                 positiv_ago = dados_ago['Positivação'].fillna(0).mean()
                 positiv_l3m = dados_l3m['Positivação'].fillna(0).mean()
                 desv_positiv = calc_desv_percentual(positiv_ago, positiv_l3m)

                 giro_ago = dados_ago['Giro Médio'].fillna(0).mean()
                 giro_l3m = dados_l3m['Giro Médio'].fillna(0).mean()
                 desv_giro = calc_desv_percentual(giro_ago, giro_l3m)

                 sku_ago = dados_ago['SKU-PDV'].fillna(0).mean()
                 sku_l3m = dados_l3m['SKU-PDV'].fillna(0).mean()
                 desv_sku = calc_desv_percentual(sku_ago, sku_l3m)

                 preco_ago = dados_ago['Preco Médio PPP'].fillna(0).mean()
                 preco_l3m = dados_l3m['Preco Médio PPP'].fillna(0).mean()
                 desv_preco = calc_desv_percentual(preco_ago, preco_l3m)

                 nome_final = 'Total' if nome == 'Total' else nome.split()[0]
                
                 resultado.append({
                    '': nome_final,
                    'DEM. PPP MES/25': round(ppp_ago),
                    'DESV. % (PPP L3M)': round(desv_ppp, 1),
                    ' ': '',  # coluna separadora
                    'POSITIV. MES/25': round(positiv_ago),
                    'DESV. % (POSITIV L3M)': round(desv_positiv, 1),
                    'GIRO MES/25': round(giro_ago, 1),
                    'DESV. % (GIRO L3M)': round(desv_giro, 1),
                    'SKU/PDV MES/25': round(sku_ago),
                    'DESV. % (SKU L3M)': round(desv_sku, 1),
                    'P. MÉDIO MES/25': round(preco_ago, 2),
                    'DESV. % (P. MÉDIO L3M)': round(desv_preco, 1),
                })
            return pd.DataFrame(resultado)
        else:   
            meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite-1]
            df_l3m = df[(df['Ano'] == 2025) & (df['Mes'].isin(meses_l3m))]
            df_ago = df[(df['Ano'] == 2025) & (df['Mes'] == mes_limite)]

            resultado = []
            for nome in df_ago[coluna_nome].unique():
                dados_ago = df_ago[df_ago[coluna_nome] == nome]
                dados_l3m = df_l3m[df_l3m[coluna_nome] == nome]

                def calc_desv_percentual(valor_atual, media_l3m):
                    return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m != 0 else 0

                ppp_ago = dados_ago['PPP Realizado'].fillna(0).sum()
                ppp_l3m = dados_l3m['PPP Realizado'].fillna(0).mean()
                desv_ppp_abs = ppp_ago - ppp_l3m
                desv_ppp = calc_desv_percentual(ppp_ago, ppp_l3m)

                positiv_ago = dados_ago['Positivação'].fillna(0).mean()
                positiv_l3m = dados_l3m['Positivação'].fillna(0).mean()
                desv_positiv = calc_desv_percentual(positiv_ago, positiv_l3m)

                giro_ago = dados_ago['Giro Médio'].fillna(0).mean()
                giro_l3m = dados_l3m['Giro Médio'].fillna(0).mean()
                desv_giro = calc_desv_percentual(giro_ago, giro_l3m)

                sku_ago = dados_ago['SKU-PDV'].fillna(0).mean()
                sku_l3m = dados_l3m['SKU-PDV'].fillna(0).mean()
                desv_sku = calc_desv_percentual(sku_ago, sku_l3m)

                preco_ago = dados_ago['Preco Médio PPP'].fillna(0).mean()
                preco_l3m = dados_l3m['Preco Médio PPP'].fillna(0).mean()
                desv_preco = calc_desv_percentual(preco_ago, preco_l3m)

                nome_final = 'Total' if nome == 'Total' else nome.split()[0]
                resultado.append({
                    '': nome_final,
                    'DEM. PPP MES/25': round(ppp_ago),
                    'DESV. % (PPP L3M)': round(desv_ppp, 1),
                    ' ': '',  # coluna separadora
                    'POSITIV. MES/25': round(positiv_ago),
                    'DESV. % (POSITIV L3M)': round(desv_positiv, 1),
                    'GIRO MES/25': round(giro_ago, 1),
                    'DESV. % (GIRO L3M)': round(desv_giro, 1),
                    'SKU/PDV MES/25': round(sku_ago),
                    'DESV. % (SKU L3M)': round(desv_sku, 1),
                    'P. MÉDIO MES/25': round(preco_ago, 2),
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
    tabela_l3m_final = tabela_l3m_final.sort_values(by='DEM. PPP MES/25', ascending=False)
    tabela_l3m_final = pd.concat([tabela_l3m_final, linha_total], ignore_index=True)

    # Filtrar dados de agosto/2025 e últimos 3 meses
    df_ago = df_grupo[(df_grupo['Ano'] == 2025) & (df_grupo['Mes'] == mes_limite)]
    df_l3m = df[(df['Ano'] == 2025) & (df['Mes'].isin(meses_l3m))]
    
    # Função para calcular indicadores gerais por grupo econômico (tabela tipo "principal")
    def calcular_indicadores_gerais_grupo(df, coluna_grupo, nome_coordenador):
        resultado = []

        # Calcula métricas por grupo
        for grupo, dados in df.groupby(coluna_grupo):
            ytd_1 = dados[(dados['Ano'] == 2025) & (dados['Mes'].isin(meses_l3m))]['PPP Realizado'].fillna(0).sum() / 3
            ytdo  = dados[(dados['Ano'] == 2025) & (dados['Mes'] == mes_limite)]['PPP Realizado'].fillna(0).sum()

            ytd_1 = 0 if pd.isna(ytd_1) else ytd_1
            ytdo  = 0 if pd.isna(ytdo)  else ytdo

            desv_abs  = ytdo - ytd_1
            desv_perc = (desv_abs / ytd_1 * 100) if ytd_1 != 0 else 0

            resultado.append({
                '': grupo,
                'RCD L3M': round(ytd_1),
                'RCD MES': round(ytdo),
                'DESV. ABS': round(desv_abs),
                'DESV. %': round(desv_perc, 1),
            })

        # DataFrame com os grupos ordenados por RCD MES (desc)
        df_resultado = pd.DataFrame(resultado).sort_values(by='RCD MES', ascending=False).reset_index(drop=True)

        # ---- TOTAL (linha do coordenador) ----
        df_ago = df_coord[(df_coord['Ano'] == 2025) & (df_coord['Mes'] == mes_limite)]
        df_l3m = df_coord[(df_coord['Ano'] == 2025) & (df_coord['Mes'].isin(meses_l3m))]

        dados    = df_ago[df_ago[coluna_nome] == nome_coordenador]
        dados_l3 = df_l3m[df_l3m[coluna_nome] == nome_coordenador]

        ytd_1 = dados_l3['PPP Realizado'].fillna(0).sum() / 3
        ytdo  = dados['PPP Realizado'].fillna(0).sum()

        ytd_1 = 0 if pd.isna(ytd_1) else ytd_1
        ytdo  = 0 if pd.isna(ytdo)  else ytdo

        desv_abs  = ytdo - ytd_1
        desv_perc = (desv_abs / ytd_1 * 100) if ytd_1 != 0 else 0

        total_row = {
            '': 'TOTAL',
            'RCD L3M': round(ytd_1),
            'RCD MES': round(ytdo),
            'DESV. ABS': round(desv_abs),
            'DESV. %': round(desv_perc, 1),
        }

        n_groups = len(df_resultado)

        # ---- Regra de exibição ----
        if n_groups > 8:
            # Apenas 7 grupos + OUTROS + TOTAL
            df_top7 = df_resultado.head(7)

            top7_sum = df_top7[['RCD L3M', 'RCD MES', 'DESV. ABS']].sum()

            outros_vals = {
                'RCD L3M': total_row['RCD L3M'] - top7_sum['RCD L3M'],
                'RCD MES': total_row['RCD MES'] - top7_sum['RCD MES'],
                'DESV. ABS': total_row['DESV. ABS'] - top7_sum['DESV. ABS'],
            }

            # Evita divisão por zero
            outros_perc = (outros_vals['DESV. ABS'] / outros_vals['RCD L3M'] * 100) if outros_vals['RCD L3M'] != 0 else 0

            outros_row = {
                '': 'OUTROS',
                'RCD L3M': round(outros_vals['RCD L3M']),
                'RCD MES': round(outros_vals['RCD MES']),
                'DESV. ABS': round(outros_vals['DESV. ABS']),
                'DESV. %': round(outros_perc, 1),
            }

            df_final = pd.concat([df_top7, pd.DataFrame([outros_row, total_row])], ignore_index=True)

        else:
            # Até 8 grupos (<= 8): mostra todos os grupos (até 8) + TOTAL, sem OUTROS
            df_top = df_resultado.head(8)
            df_final = pd.concat([df_top, pd.DataFrame([total_row])], ignore_index=True)

        return df_final
    
    
    def calcular_tabela_por_grupo(dados_ago, dados_l3m, nome_coordenador):
        # Usa variáveis globais: df_coord, coluna_nome, coluna_grupo
        # 'grupos_top7' não é usado; a seleção é feita pela própria função.
        import os
        from dotenv import load_dotenv
        load_dotenv()
        territorio_path = os.getenv("TERRITORIO_PATH" )
        path_territorio = os.path.join(downloads_folder, territorio_path)
        def calc_desv_percentual(valor_atual, media_l3m):
             return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m not in [0, None, float('nan')] and media_l3m != 0 else 0
        territorio = False
        if os.path.isfile(path_territorio):
            territorio = True
            # Importa a tabela de território
            df_territorio = pd.read_excel(path_territorio, skiprows=2,engine="openpyxl")
            # Filtra os dados para o mês e coordenador
            
            
            # ---- 1) KPIs por grupo (mesma lógica para todos) ----
            resultado = []
            for grupo in dados_ago[coluna_grupo].unique():
                dados_ago_grupo = dados_ago[dados_ago[coluna_grupo] == grupo]
                dados_l3m_grupo = dados_l3m[dados_l3m[coluna_grupo] == grupo]
                

                dados_grupo = df_territorio[df_territorio["Nome Gd"] == grupo]
                ppp_ago = dados_grupo["DMD | OL"].fillna(0).sum()
                ppp_l3m = dados_grupo["L3M"].fillna(0).mean()
                desv_ppp = calc_desv_percentual(ppp_ago, ppp_l3m)

                positiv_ago = dados_ago_grupo['Positivação'].fillna(0).mean()
                positiv_l3m = dados_l3m_grupo['Positivação'].fillna(0).mean()
                desv_positiv = calc_desv_percentual(positiv_ago, positiv_l3m)

                giro_ago = dados_ago_grupo['Giro Médio'].fillna(0).mean()
                giro_l3m = dados_l3m_grupo['Giro Médio'].fillna(0).mean()
                desv_giro = calc_desv_percentual(giro_ago, giro_l3m)

                sku_ago = dados_ago_grupo['SKU-PDV'].fillna(0).mean()
                sku_l3m = dados_l3m_grupo['SKU-PDV'].fillna(0).mean()
                desv_sku = calc_desv_percentual(sku_ago, sku_l3m)

                preco_ago = dados_ago_grupo['Preco Médio PPP'].fillna(0).mean()
                preco_l3m = dados_l3m_grupo['Preco Médio PPP'].fillna(0).mean()
                desv_preco = calc_desv_percentual(preco_ago, preco_l3m)
                resultado.append({
                    '': grupo,
                    'DEM. OL AGO/25': round(ppp_ago),
                    'DESV. % (PPP L3M)': round(desv_ppp, 1),
                    ' ': '',  # coluna separadora
                    'POSITIV. AGO/25': round(positiv_ago),
                    'DESV. % (POSITIV L3M)': round(desv_positiv, 1),
                    'GIRO AGO/25': round(giro_ago, 1),
                    'DESV. % (GIRO L3M)': round(desv_giro, 1),
                    'SKU/PDV AGO/25': round(sku_ago),
                    'DESV. % (SKU L3M)': round(desv_sku, 1),
                    'P. MÉDIO AGO/25': round(preco_ago, 2),
                    'DESV. % (P. MÉDIO L3M)': round(desv_preco, 1),
                })
            def calc_desv_percentual(valor_atual, media_l3m):
                return ((valor_atual - media_l3m) / media_l3m * 100) if media_l3m not in [0, None, float('nan')] and media_l3m != 0 else 0
            df_grupos = pd.DataFrame(resultado).sort_values(by='DEM. OL AGO/25', ascending=False).reset_index(drop=True)
            n_groups = len(df_grupos)

            # ---- 2) Seleção TOP e cálculo de OUTROS conforme regra ----
            if n_groups > 8:
                # TOP 7
                df_top7 = df_grupos.head(7).copy()
                grupos_top7_calc = set(df_top7[''].tolist())

                # Conjunto OUTROS com a MESMA lógica dos grupos (usando os dados brutos desses grupos)
                mask_outros_ago = ~dados_ago[coluna_grupo].isin(grupos_top7_calc)
                mask_outros_l3m = ~dados_l3m[coluna_grupo].isin(grupos_top7_calc)

                dados_outros_ago = dados_ago[mask_outros_ago]
                dados_outros_l3m = dados_l3m[mask_outros_l3m]

                ppp_ago_o = dados_outros_ago['DEM. OL AGO/25'].fillna(0).sum()
                ppp_l3m_o = dados_outros_l3m['L3M'].fillna(0).mean()
                desv_ppp_o = calc_desv_percentual(ppp_ago_o, ppp_l3m_o)

                positiv_ago_o = dados_outros_ago['Positivação'].fillna(0).mean()
                positiv_l3m_o = dados_outros_l3m['Positivação'].fillna(0).mean()
                desv_positiv_o = calc_desv_percentual(positiv_ago_o, positiv_l3m_o)

                giro_ago_o = dados_outros_ago['Giro Médio'].fillna(0).mean()
                giro_l3m_o = dados_outros_l3m['Giro Médio'].fillna(0).mean()
                desv_giro_o = calc_desv_percentual(giro_ago_o, giro_l3m_o)

                sku_ago_o = dados_outros_ago['SKU-PDV'].fillna(0).mean()
                sku_l3m_o = dados_outros_l3m['SKU-PDV'].fillna(0).mean()
                desv_sku_o = calc_desv_percentual(sku_ago_o, sku_l3m_o)

                preco_ago_o = dados_outros_ago['Preco Médio PPP'].fillna(0).mean()
                preco_l3m_o = dados_outros_l3m['Preco Médio PPP'].fillna(0).mean()
                desv_preco_o = calc_desv_percentual(preco_ago_o, preco_l3m_o)

                outros_row = {
                    '': 'OUTROS',
                    'DMD | OL': round(ppp_ago_o),
                    'DESV. % (PPP L3M)': round(desv_ppp_o, 1),
                    ' ': '',
                    'POSITIV. AGO/25': round(positiv_ago_o),
                    'DESV. % (POSITIV L3M)': round(desv_positiv_o, 1),
                    'GIRO AGO/25': round(giro_ago_o, 1),
                    'DESV. % (GIRO L3M)': round(desv_giro_o, 1),
                    'SKU/PDV AGO/25': round(sku_ago_o),
                    'DESV. % (SKU L3M)': round(desv_sku_o, 1),
                    'P. MÉDIO AGO/25': round(preco_ago_o, 2),
                    'DESV. % (P. MÉDIO L3M)': round(desv_preco_o, 1),
                }

                df_final = pd.concat([df_top7, pd.DataFrame([outros_row])], ignore_index=True)

            else:
                # Até 8 grupos: mostra todos (até 8), sem OUTROS
                df_final = df_grupos.head(8).copy()

            # ---- 3) TOTAL (linha do coordenador) ----
            df_ago = df_coord[(df_coord['Ano'] == 2025) & (df_coord['Mes'] == mes_limite)]
            df_l3m = df_coord[(df_coord['Ano'] == 2025) & (df_coord['Mes'].isin(meses_l3m))]

            dados    = df_ago[df_ago[coluna_nome] == nome_coordenador]
            dados_l3 = df_l3m[df_l3m[coluna_nome] == nome_coordenador]

            ytd_1 = dados_l3['PPP Realizado'].fillna(0).sum() / 3
            ytdo  = dados['PPP Realizado'].fillna(0).sum()

            ytd_1 = 0 if pd.isna(ytd_1) else ytd_1
            ytdo  = 0 if pd.isna(ytdo)  else ytdo

            desv_abs  = ytdo - ytd_1
            desv_perc = (desv_abs / ytd_1 * 100) if ytd_1 != 0 else 0

            ppp_ago = dados['PPP Realizado'].fillna(0).sum()
            ppp_l3m = dados_l3['PPP Realizado'].fillna(0).mean()

            positiv_ago = dados['Positivação'].fillna(0).mean()
            positiv_l3m = dados_l3['Positivação'].fillna(0).mean()
            desv_positiv = ((positiv_ago - positiv_l3m) / positiv_l3m * 100) if positiv_l3m != 0 else 0

            giro_ago = dados['Giro Médio'].fillna(0).mean()
            giro_l3m = dados_l3['Giro Médio'].fillna(0).mean()
            desv_giro = ((giro_ago - giro_l3m) / giro_l3m * 100) if giro_l3m != 0 else 0

            sku_ago = dados['SKU-PDV'].fillna(0).mean()
            sku_l3m = dados_l3['SKU-PDV'].fillna(0).mean()
            desv_sku = ((sku_ago - sku_l3m) / sku_l3m * 100) if sku_l3m != 0 else 0

            preco_ago = dados['Preco Médio PPP'].fillna(0).mean()
            preco_l3m = dados_l3['Preco Médio PPP'].fillna(0).mean()
            desv_preco = ((preco_ago - preco_l3m) / preco_l3m * 100) if preco_l3m != 0 else 0

            total_row = {
                '': 'TOTAL',
                'DMD | OL': round(ppp_ago),
                'DESV. % (PPP L3M)': round(desv_perc, 1),
                ' ': '',
                'POSITIV. AGO/25': round(positiv_ago, 2),
                'DESV. % (POSITIV L3M)': round(desv_positiv, 2),
                'GIRO AGO/25': round(giro_ago, 2),
                'DESV. % (GIRO L3M)': round(desv_giro, 2),
                'SKU/PDV AGO/25': round(sku_ago, 2),
                'DESV. % (SKU L3M)': round(desv_sku, 2),
                'P. MÉDIO AGO/25': round(preco_ago, 2),
                'DESV. % (P. MÉDIO L3M)': round(desv_preco, 2),
            }

            df_final = pd.concat([df_final, pd.DataFrame([total_row])], ignore_index=True)
            return df_final
        else:
           # ---- 1) KPIs por grupo (mesma lógica para todos) ----
            resultado = []
            for grupo in dados_ago[coluna_grupo].unique():
                dados_ago_grupo = dados_ago[dados_ago[coluna_grupo] == grupo]
                dados_l3m_grupo = dados_l3m[dados_l3m[coluna_grupo] == grupo]

                ppp_ago = dados_ago_grupo['PPP Realizado'].fillna(0).sum()
                ppp_l3m = dados_l3m_grupo['PPP Realizado'].fillna(0).mean()
                desv_ppp = calc_desv_percentual(ppp_ago, ppp_l3m)

                positiv_ago = dados_ago_grupo['Positivação'].fillna(0).mean()
                positiv_l3m = dados_l3m_grupo['Positivação'].fillna(0).mean()
                desv_positiv = calc_desv_percentual(positiv_ago, positiv_l3m)

                giro_ago = dados_ago_grupo['Giro Médio'].fillna(0).mean()
                giro_l3m = dados_l3m_grupo['Giro Médio'].fillna(0).mean()
                desv_giro = calc_desv_percentual(giro_ago, giro_l3m)

                sku_ago = dados_ago_grupo['SKU-PDV'].fillna(0).mean()
                sku_l3m = dados_l3m_grupo['SKU-PDV'].fillna(0).mean()
                desv_sku = calc_desv_percentual(sku_ago, sku_l3m)

                preco_ago = dados_ago_grupo['Preco Médio PPP'].fillna(0).mean()
                preco_l3m = dados_l3m_grupo['Preco Médio PPP'].fillna(0).mean()
                desv_preco = calc_desv_percentual(preco_ago, preco_l3m)

                resultado.append({
                    '': grupo,
                    'DEM. PPP AGO/25': round(ppp_ago),
                    'DESV. % (PPP L3M)': round(desv_ppp, 1),
                    ' ': '',  # coluna separadora
                    'POSITIV. AGO/25': round(positiv_ago),
                    'DESV. % (POSITIV L3M)': round(desv_positiv, 1),
                    'GIRO AGO/25': round(giro_ago, 1),
                    'DESV. % (GIRO L3M)': round(desv_giro, 1),
                    'SKU/PDV AGO/25': round(sku_ago),
                    'DESV. % (SKU L3M)': round(desv_sku, 1),
                    'P. MÉDIO AGO/25': round(preco_ago, 2),
                    'DESV. % (P. MÉDIO L3M)': round(desv_preco, 1),
                })

            df_grupos = pd.DataFrame(resultado).sort_values(by='DEM. PPP AGO/25', ascending=False).reset_index(drop=True)
            n_groups = len(df_grupos)

            # ---- 2) Seleção TOP e cálculo de OUTROS conforme regra ----
            if n_groups > 8:
                # TOP 7
                df_top7 = df_grupos.head(7).copy()
                grupos_top7_calc = set(df_top7[''].tolist())

                # Conjunto OUTROS com a MESMA lógica dos grupos (usando os dados brutos desses grupos)
                mask_outros_ago = ~dados_ago[coluna_grupo].isin(grupos_top7_calc)
                mask_outros_l3m = ~dados_l3m[coluna_grupo].isin(grupos_top7_calc)

                dados_outros_ago = dados_ago[mask_outros_ago]
                dados_outros_l3m = dados_l3m[mask_outros_l3m]

                ppp_ago_o = dados_outros_ago['PPP Realizado'].fillna(0).sum()
                ppp_l3m_o = dados_outros_l3m['PPP Realizado'].fillna(0).mean()
                desv_ppp_o = calc_desv_percentual(ppp_ago_o, ppp_l3m_o)

                positiv_ago_o = dados_outros_ago['Positivação'].fillna(0).mean()
                positiv_l3m_o = dados_outros_l3m['Positivação'].fillna(0).mean()
                desv_positiv_o = calc_desv_percentual(positiv_ago_o, positiv_l3m_o)

                giro_ago_o = dados_outros_ago['Giro Médio'].fillna(0).mean()
                giro_l3m_o = dados_outros_l3m['Giro Médio'].fillna(0).mean()
                desv_giro_o = calc_desv_percentual(giro_ago_o, giro_l3m_o)

                sku_ago_o = dados_outros_ago['SKU-PDV'].fillna(0).mean()
                sku_l3m_o = dados_outros_l3m['SKU-PDV'].fillna(0).mean()
                desv_sku_o = calc_desv_percentual(sku_ago_o, sku_l3m_o)

                preco_ago_o = dados_outros_ago['Preco Médio PPP'].fillna(0).mean()
                preco_l3m_o = dados_outros_l3m['Preco Médio PPP'].fillna(0).mean()
                desv_preco_o = calc_desv_percentual(preco_ago_o, preco_l3m_o)

                outros_row = {
                    '': 'OUTROS',
                    'DEM. PPP AGO/25': round(ppp_ago_o),
                    'DESV. % (PPP L3M)': round(desv_ppp_o, 1),
                    ' ': '',
                    'POSITIV. AGO/25': round(positiv_ago_o),
                    'DESV. % (POSITIV L3M)': round(desv_positiv_o, 1),
                    'GIRO AGO/25': round(giro_ago_o, 1),
                    'DESV. % (GIRO L3M)': round(desv_giro_o, 1),
                    'SKU/PDV AGO/25': round(sku_ago_o),
                    'DESV. % (SKU L3M)': round(desv_sku_o, 1),
                    'P. MÉDIO AGO/25': round(preco_ago_o, 2),
                    'DESV. % (P. MÉDIO L3M)': round(desv_preco_o, 1),
                }

                df_final = pd.concat([df_top7, pd.DataFrame([outros_row])], ignore_index=True)

            else:
                # Até 8 grupos: mostra todos (até 8), sem OUTROS
                df_final = df_grupos.head(8).copy()

            # ---- 3) TOTAL (linha do coordenador) ----
            df_ago = df_coord[(df_coord['Ano'] == 2025) & (df_coord['Mes'] == mes_limite)]
            df_l3m = df_coord[(df_coord['Ano'] == 2025) & (df_coord['Mes'].isin(meses_l3m))]

            dados    = df_ago[df_ago[coluna_nome] == nome_coordenador]
            dados_l3 = df_l3m[df_l3m[coluna_nome] == nome_coordenador]

            ytd_1 = dados_l3['PPP Realizado'].fillna(0).sum() / 3
            ytdo  = dados['PPP Realizado'].fillna(0).sum()

            ytd_1 = 0 if pd.isna(ytd_1) else ytd_1
            ytdo  = 0 if pd.isna(ytdo)  else ytdo

            desv_abs  = ytdo - ytd_1
            desv_perc = (desv_abs / ytd_1 * 100) if ytd_1 != 0 else 0

            ppp_ago = dados['PPP Realizado'].fillna(0).sum()
            ppp_l3m = dados_l3['PPP Realizado'].fillna(0).mean()

            positiv_ago = dados['Positivação'].fillna(0).mean()
            positiv_l3m = dados_l3['Positivação'].fillna(0).mean()
            desv_positiv = ((positiv_ago - positiv_l3m) / positiv_l3m * 100) if positiv_l3m != 0 else 0

            giro_ago = dados['Giro Médio'].fillna(0).mean()
            giro_l3m = dados_l3['Giro Médio'].fillna(0).mean()
            desv_giro = ((giro_ago - giro_l3m) / giro_l3m * 100) if giro_l3m != 0 else 0

            sku_ago = dados['SKU-PDV'].fillna(0).mean()
            sku_l3m = dados_l3['SKU-PDV'].fillna(0).mean()
            desv_sku = ((sku_ago - sku_l3m) / sku_l3m * 100) if sku_l3m != 0 else 0

            preco_ago = dados['Preco Médio PPP'].fillna(0).mean()
            preco_l3m = dados_l3['Preco Médio PPP'].fillna(0).mean()
            desv_preco = ((preco_ago - preco_l3m) / preco_l3m * 100) if preco_l3m != 0 else 0

            total_row = {
                '': 'TOTAL',
                'DEM. PPP AGO/25': round(ppp_ago),
                'DESV. % (PPP L3M)': round(desv_perc, 1),
                ' ': '',
                'POSITIV. AGO/25': round(positiv_ago, 2),
                'DESV. % (POSITIV L3M)': round(desv_positiv, 2),
                'GIRO AGO/25': round(giro_ago, 2),
                'DESV. % (GIRO L3M)': round(desv_giro, 2),
                'SKU/PDV AGO/25': round(sku_ago, 2),
                'DESV. % (SKU L3M)': round(desv_sku, 2),
                'P. MÉDIO AGO/25': round(preco_ago, 2),
                'DESV. % (P. MÉDIO L3M)': round(desv_preco, 2),
            }

            df_final = pd.concat([df_final, pd.DataFrame([total_row])], ignore_index=True)
            return df_final


    #Gerar tabelas por coordenador
    for coordenador in df_ago[coluna_nome].unique():
        dados_coord = df_grupo[(df_grupo[coluna_nome] == coordenador)]
        dados_ago_coord = df_ago[df_ago[coluna_nome] == coordenador]
        dados_l3m_coord = df_l3m[df_l3m[coluna_nome] == coordenador]
        # Tabela tipo "principal" (YTD)
        tabela_ytd = calcular_indicadores_gerais_grupo(dados_coord, coluna_grupo, coordenador)
        grupos_top7 = tabela_ytd[(tabela_ytd[''] != 'OUTROS') & (tabela_ytd[''] != 'TOTAL')][''].tolist()
        # Tabela tipo "imagem" (KPI) na mesma ordem do YTD
        tabela_kpi = calcular_tabela_por_grupo(dados_ago_coord, dados_l3m_coord, coordenador)
        primeiro_nome = coordenador.split()[0]
        tabela_kpi.to_excel(os.path.join(rt_folder, f"tabela_top7_KPI_{primeiro_nome}.xlsx"), index=False)
        tabela_ytd.to_excel(os.path.join(rt_folder, f"tabela_top7_YTD_{primeiro_nome}.xlsx"), index=False)

        wb = load_workbook(os.path.join(rt_folder, f"tabela_top7_YTD_{primeiro_nome}.xlsx"))
        ws = wb.active  # primeira planilha (ou use wb['NomeDaAba'] se você definiu)

        # Formatos desejados
        formato_milhar = '#,##0'     # B, C, D: X.XXX.XXX (sem casas decimais, usa locale pt-BR)
        formato_percent_txt = '0.0"%"'  # E: exibe 12,3%

        # Descobre a última linha com dados
        ultima_linha = ws.max_row

        # Colunas B, C, D => separador de milhar, sem decimais
        for col in ('B', 'C', 'D'):
            for row in range(2, ultima_linha + 1):  # pula o cabeçalho
                cell = ws[f'{col}{row}']
                if cell.value is not None:
                    cell.number_format = formato_milhar

        # Coluna E => mostrar o símbolo % APÓS o número (sem mudar a escala)
        for row in range(2, ultima_linha + 1):
            cell = ws[f'E{row}']
            if cell.value is not None:
                cell.number_format = formato_percent_txt

        # Salva as formatações
        wb.save(os.path.join(rt_folder, f"tabela_top7_YTD_{primeiro_nome}.xlsx"))

         # Abre o arquivo
        wb = load_workbook(os.path.join(rt_folder, f"tabela_top7_KPI_{primeiro_nome}.xlsx"))
        ws = wb.active  # ou: ws = wb['L3M'] se você nomeou a aba

        # Formatos desejados
        formato_milhar = '#,##0'       # B e C: X.XXX.XXX (sem casas decimais; Excel pt-BR usa ponto para milhar)
        formato_percent_txt = '0.0"%"' # D, F, H, J, L: exibe 1 casa decimal + símbolo % (sem mudar escala)

        # Descobre a última linha com dados
        ultima_linha = ws.max_row

        # Colunas B, C => separador de milhar, sem decimais
        for col in ('B'):
            for row in range(2, ultima_linha + 1):  # pula o cabeçalho
                cell = ws[f'{col}{row}']
                if cell.value is not None and cell.value != "":
                    cell.number_format = formato_milhar

        # Colunas D, F, H, J, L => número seguido de % (sem mudar a escala)
        for col in ('C','D', 'F', 'H', 'J', 'L'):
            for row in range(2, ultima_linha + 1):
                cell = ws[f'{col}{row}']
                if cell.value is not None and cell.value != "":
                    cell.number_format = formato_percent_txt

        # Salva as formatações
        wb.save(os.path.join(rt_folder, f"tabela_top7_KPI_{primeiro_nome}.xlsx"))
    # Exportar tabelas principais
    tabela_final.to_excel(os.path.join(rt_folder, "tabela_ytd.xlsx"), index=False)
    tabela_l3m_final.to_excel(os.path.join(rt_folder, "tabela_l3m_agosto.xlsx"), index=False)
    
    
    # Abre o arquivo
    wb = load_workbook(os.path.join(rt_folder, "tabela_l3m_agosto.xlsx"))
    ws = wb.active  # ou: ws = wb['L3M'] se você nomeou a aba

    # Formatos desejados
    formato_milhar = '#,##0'       # B e C: X.XXX.XXX (sem casas decimais; Excel pt-BR usa ponto para milhar)
    formato_percent_txt = '0.0"%"' # D, F, H, J, L: exibe 1 casa decimal + símbolo % (sem mudar escala)

    # Descobre a última linha com dados
    ultima_linha = ws.max_row

    # Colunas B, C => separador de milhar, sem decimais
    for col in ('B'):
        for row in range(2, ultima_linha + 1):  # pula o cabeçalho
            cell = ws[f'{col}{row}']
            if cell.value is not None and cell.value != "":
                cell.number_format = formato_milhar

    # Colunas D, F, H, J, L => número seguido de % (sem mudar a escala)
    for col in ('C', 'F', 'H', 'J', 'L'):
        for row in range(2, ultima_linha + 1):
            cell = ws[f'{col}{row}']
            if cell.value is not None and cell.value != "":
                cell.number_format = formato_percent_txt

    # Salva as formatações
    wb.save(os.path.join(rt_folder, "tabela_l3m_agosto.xlsx"))



    wb = load_workbook(os.path.join(rt_folder, "tabela_ytd.xlsx"))
    ws = wb.active  # primeira planilha (ou use wb['NomeDaAba'] se você definiu)

    # Formatos desejados
    formato_milhar = '#,##0'     # B, C, D: X.XXX.XXX (sem casas decimais, usa locale pt-BR)
    formato_percent_txt = '0.0"%"'  # E: exibe 12,3%

    # Descobre a última linha com dados
    ultima_linha = ws.max_row

    # Colunas B, C, D => separador de milhar, sem decimais
    for col in ('B', 'C', 'D'):
        for row in range(2, ultima_linha + 1):  # pula o cabeçalho
            cell = ws[f'{col}{row}']
            if cell.value is not None:
                cell.number_format = formato_milhar

    # Coluna E => mostrar o símbolo % APÓS o número (sem mudar a escala)
    for row in range(2, ultima_linha + 1):
        cell = ws[f'E{row}']
        if cell.value is not None:
            cell.number_format = formato_percent_txt

    # Salva as formatações
    wb.save(os.path.join(rt_folder, "tabela_ytd.xlsx"))

    print(f"Arquivos salvos e formatados na pasta: {rt_folder}")
if __name__ == "__main__":
    mes_limite = int(input("Informe o mês limite (número de 1 a 12): "))
   
    main(mes_limite)