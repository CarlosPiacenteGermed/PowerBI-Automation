import os
import pandas as pd
from dotenv import load_dotenv


def calcular_sellin_l3m(df_total, df_nao_visitado, df_visitado, mes_limite):
    load_dotenv()
    coluna_nome = os.getenv("COLUNA_NOME")  # Ex: 'Responsável'

    # padronizar datas
    for df in [df_total, df_visitado, df_nao_visitado]:
        df['Mes/Ano'] = pd.to_datetime(df['Mes/Ano'])
        df['Ano'] = df['Mes/Ano'].dt.year
        df['Mes'] = df['Mes/Ano'].dt.month

    meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite - 1]

    def calcular_l3m_mes(df):
        soma_l3m = (
            df[(df['Ano'] == 2025) & (df['Mes'].isin(meses_l3m))]['PPP Realizado']
            .fillna(0)
            .sum()
        )
        mes = (
            df[(df['Ano'] == 2025) & (df['Mes'] == mes_limite)]['PPP Realizado']
            .fillna(0)
            .sum()
        )
        l3m = soma_l3m / 3 if soma_l3m != 0 else 0
        desv_abs = mes - l3m
        desv_perc = (desv_abs / l3m * 100) if l3m != 0 else 0
        return l3m, mes, desv_abs, desv_perc

    resultado = []
    responsaveis = df_total[coluna_nome].unique()

    # ordenar por RCD MES (total)
    rcd_mes_totais = {}
    for resp in responsaveis:
        total = df_total[df_total[coluna_nome] == resp]
        _, mes_total, _, _ = calcular_l3m_mes(total)
        rcd_mes_totais[resp] = mes_total

    responsaveis_ordenados = sorted(responsaveis, key=lambda r: rcd_mes_totais[r], reverse=True)

    for resp in responsaveis_ordenados:
        total = df_total[df_total[coluna_nome] == resp]
        visitado = df_visitado[df_visitado[coluna_nome] == resp]
        nao_visitado = df_nao_visitado[df_nao_visitado[coluna_nome] == resp]

        l3m_total, mes_total, desv_total, perc_total = calcular_l3m_mes(total)
        l3m_vis, mes_vis, desv_vis, perc_vis = calcular_l3m_mes(visitado)
        l3m_nao, mes_nao, desv_nao, perc_nao = calcular_l3m_mes(nao_visitado)

        # Garantir que VISITADO + NÃO VISITADO some 100%:
        denom = mes_vis + mes_nao
        if denom > 0:
            repres_vis = mes_vis / denom * 100
            repres_nao = mes_nao / denom * 100
        else:
            # Se não há dados de visitado/nao, manter fallback para mes_total (ou 0)
            repres_vis = (mes_vis / mes_total * 100) if mes_total != 0 else 0
            repres_nao = (mes_nao / mes_total * 100) if mes_total != 0 else 0

        resultado.append({
            'Responsável': resp,
            'RCD L3M': l3m_vis,
            'RCD MES': mes_vis,
            'DESV. ABS': desv_vis,
            'REPRES. %': 100,
            'DESV. %': perc_vis,
           
        })

        resultado.append({
            'Responsável': 'VISITADO',
            'RCD L3M': l3m_total,
            'RCD MES': mes_total,
            'DESV. ABS': desv_total,
            'REPRES. %': repres_vis,        # total sempre 100%
            'DESV. %': perc_total,
        })

        resultado.append({
            'Responsável': 'NÃO VISITADO',
            'RCD L3M': l3m_nao,
            'RCD MES': mes_nao,
            'DESV. ABS': desv_nao,
            'REPRES. %': repres_nao,
            'DESV. %': perc_nao,
        })

    df_res = pd.DataFrame(resultado)
    # garantir numéricos
    numeric_cols = ['RCD L3M', 'RCD MES', 'DESV. ABS', 'REPRES. %', 'DESV. %']
    for c in numeric_cols:
        if c in df_res.columns:
            df_res[c] = pd.to_numeric(df_res[c], errors='coerce').fillna(0)

    return df_res


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
        provedores = df_ago_total[(df_ago_total[coluna_responsavel] == responsavel)][coluna_grupo].unique()

        for grupo in provedores:
            # Linha do grupo (total)
            dados_ago_grupo = df_ago_total[(df_ago_total[coluna_grupo] == grupo) & (df_ago_total[coluna_responsavel] == responsavel)]
            dados_l3m_grupo = df_l3m_total[(df_l3m_total[coluna_grupo] == grupo) & (df_l3m_total[coluna_responsavel] == responsavel)]
            dados_ago_vis = df_ago_visitado[(df_ago_visitado[coluna_grupo] == grupo) & (df_ago_visitado[coluna_responsavel] == responsavel)]
            dados_l3m_vis = df_l3m_visitado[(df_l3m_visitado[coluna_grupo] == grupo) & (df_l3m_visitado[coluna_responsavel] == responsavel)]
            kpis_total = calcular_kpis(dados_ago_vis, dados_l3m_vis)
            kpis_total['Responsável'] = grupo
            resultado.append(kpis_total)

            # Linha VISITADO
           

            kpis_vis = calcular_kpis(dados_ago_grupo, dados_l3m_grupo)
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

    mes_limite = int(input("Informe o mês limite (número de 1 a 12): "))
    meses_l3m = [mes_limite - 3, mes_limite - 2, mes_limite - 1]

    df_ago_total = df_total[(df_total['Ano'] == 2025) & (df_total['Mes'] == mes_limite)]
    df_l3m_total = df_total[(df_total['Ano'] == 2025) & (df_total['Mes'].isin(meses_l3m))]

    df_ago_visitado = df_visitado[(df_visitado['Ano'] == 2025) & (df_visitado['Mes'] == mes_limite)]
    df_l3m_visitado = df_visitado[(df_visitado['Ano'] == 2025) & (df_visitado['Mes'].isin(meses_l3m))]

    df_ago_nvisitado = df_nao_visitado[(df_nao_visitado['Ano'] == 2025) & (df_nao_visitado['Mes'] == mes_limite)]
    df_l3m_nvisitado = df_nao_visitado[(df_nao_visitado['Ano'] == 2025) & (df_nao_visitado['Mes'].isin(meses_l3m))]

    # === Lógica condicional conforme .env ===
    if coluna_nome == "REGIONAL DISTY" and coluna_grupo == "CONTAS DISTY":
     # Tabela GR: como já está
        df_sellin_gr = calcular_sellin_l3m(df_total, df_nao_visitado, df_visitado, mes_limite)

        # Tabela GD: para cada REGIONAL DISTY, pega os clientes (CONTAS DISTY) e calcula VISITADO/NÃO VISITADO
        resultado_gd = []

        regionais = df_total[coluna_nome].unique()
        for regional in regionais:
            # Filtra clientes dessa regional
            clientes = df_total[df_total[coluna_nome] == regional][coluna_grupo].unique()
            for cliente in clientes:
                df_cliente_vis = df_visitado[(df_visitado[coluna_nome] == regional) & (df_visitado[coluna_grupo] == cliente)]
                df_cliente_naovis = df_nao_visitado[(df_nao_visitado[coluna_nome] == regional) & (df_nao_visitado[coluna_grupo] == cliente)]
                df_cliente_total = df_total[(df_total[coluna_nome] == regional) & (df_total[coluna_grupo] == cliente)]

                # Calcula os três tipos de linha
                l3m_total, mes_total, desv_total, perc_total = 0, 0, 0, 0
                l3m_vis, mes_vis, desv_vis, perc_vis = 0, 0, 0, 0
                l3m_nao, mes_nao, desv_nao, perc_nao = 0, 0, 0, 0

                if not df_cliente_total.empty:
                    df_calc = calcular_sellin_l3m(df_cliente_total, df_cliente_naovis, df_cliente_vis, mes_limite)
                    # Linha do responsável (primeira linha)
                    l3m_total, mes_total, desv_total, perc_total, repres_total= df_calc.iloc[0][['RCD L3M', 'RCD MES', 'DESV. ABS', 'DESV. %', 'REPRES. %']]
                    # VISITADO (segunda linha)
                    l3m_vis, mes_vis, desv_vis, perc_vis, repres_vis = df_calc[df_calc['Responsável'] == 'VISITADO'][['RCD L3M', 'RCD MES', 'DESV. ABS', 'DESV. %', 'REPRES. %']].values[0]
                    # NÃO VISITADO (terceira linha)
                    l3m_nao, mes_nao, desv_nao, perc_nao, repres_nao = df_calc[df_calc['Responsável'] == 'NÃO VISITADO'][['RCD L3M', 'RCD MES', 'DESV. ABS', 'DESV. %', 'REPRES. %']].values[0]

                resultado_gd.append({
                    'Regional': regional,
                    'Cliente': cliente,
                    'Tipo': 'TOTAL',
                    'RCD L3M': l3m_total,
                    'RCD MES': mes_total,
                    'REPRES %': repres_total,  # total sempre 100%
                    'DESV. ABS': desv_total,
                    'DESV. %': perc_total
                })
                resultado_gd.append({
                    'Regional': regional,
                    'Cliente': cliente,
                    'Tipo': 'VISITADO',
                    'RCD L3M': l3m_vis,
                    'RCD MES': mes_vis,
                    'REPRES %': repres_vis,  
                    'DESV. ABS': desv_vis,
                    'DESV. %': perc_vis
                })
                resultado_gd.append({
                    'Regional': regional,
                    'Cliente': cliente,
                    'Tipo': 'NÃO VISITADO',
                    'RCD L3M': l3m_nao,
                    'RCD MES': mes_nao,
                    'REPRES %': repres_nao,  
                    'DESV. ABS': desv_nao,
                    'DESV. %': perc_nao
                })

        df_sellin_gd = pd.DataFrame(resultado_gd)

        with pd.ExcelWriter(os.path.join(rt_folder, f"{coluna_nome}_sellin_GR_GD.xlsx")) as writer:
            # Formatação para tabela GR
            df_gr_fmt = df_sellin_gr.copy()
            # Colunas B,C,D: 'RCD L3M', 'RCD MES', 'DESV. ABS' -> X.XXX.XXX
            for col in ['RCD L3M', 'RCD MES', 'DESV. ABS']:
                df_gr_fmt[col] = df_gr_fmt[col].apply(lambda x: f"{int(round(x)):,}".replace(",", "."))
            # Colunas E,F: 'REPRES. %', 'DESV. %' -> X,X%
            for col in ['REPRES. %', 'DESV. %']:
                df_gr_fmt[col] = df_gr_fmt[col].apply(lambda x: f"{x:.1f}%".replace(".", ","))

            # Formatação para tabela GD
            df_gd_fmt = df_sellin_gd.copy()
            # Colunas D,E,G: 'RCD L3M', 'RCD MES', 'DESV. ABS' -> X.XXX.XXX
            for col in ['RCD L3M', 'RCD MES', 'DESV. ABS']:
                df_gd_fmt[col] = df_gd_fmt[col].apply(lambda x: f"{int(round(x)):,}".replace(",", "."))
            # Colunas F,H: 'REPRES %', 'DESV. %' -> X,X%
            for col in ['REPRES %', 'DESV. %']:
                df_gd_fmt[col] = df_gd_fmt[col].apply(lambda x: f"{x:.1f}%".replace(".", ","))

            df_gr_fmt.to_excel(writer, sheet_name="GR", index=False)
            df_gd_fmt.to_excel(writer, sheet_name="GD", index=False)
        print("Arquivo 'sellin_GR_GD.xlsx' gerado com sucesso.")


    elif coluna_nome == "CONTAS DISTY" and coluna_grupo == "PROVEDOR BM":
    # Gera apenas KPI dos clientes de cada GD
        df_resultado_kpis_grupo = calcular_tabela_por_grupo_agrupado(
            df_ago_total, df_l3m_total,
            df_ago_visitado, df_l3m_visitado,
            df_ago_nvisitado, df_l3m_nvisitado,
            coluna_grupo, 'CONTAS DISTY'
        )
        colunas_ordenadas = ['Responsável'] + [col for col in df_resultado_kpis_grupo.columns if col != 'Responsável']
        df_resultado_kpis_grupo = df_resultado_kpis_grupo[colunas_ordenadas]

        # Formatação conforme solicitado:
        # B: 'DEM. PPP AGO/25' -> X.XXX.XXX
        df_resultado_kpis_grupo['DEM. PPP AGO/25'] = df_resultado_kpis_grupo['DEM. PPP AGO/25'].apply(
            lambda x: f"{int(round(x)):,}".replace(",", ".") if pd.notnull(x) else "0"
        )
        # C: 'DEM. PPP l3m' -> X.XXX.XXX
        df_resultado_kpis_grupo['DEM. PPP l3m'] = df_resultado_kpis_grupo['DEM. PPP l3m'].apply(
            lambda x: f"{int(round(x)):,}".replace(",", ".") if pd.notnull(x) else "0"
        )
        # D,F,H,J,L: X,X%
        for col in ['DESV. % (PPP L3M)', 'DESV. % (POSITIV L3M)', 'DESV. % (GIRO L3M)', 'DESV. % (SKU L3M)', 'DESV. % (P. MÉDIO L3M)']:
            df_resultado_kpis_grupo[col] = df_resultado_kpis_grupo[col].apply(
                lambda x: f"{x:.1f}%".replace(".", ",") if pd.notnull(x) else "0,0%"
            )
        # E, I: XXX
        for col in ['POSITIV. AGO/25', 'SKU/PDV AGO/25']:
            df_resultado_kpis_grupo[col] = df_resultado_kpis_grupo[col].apply(
                lambda x: f"{int(round(x))}" if pd.notnull(x) else "0"
            )
        # G: 'GIRO AGO/25' -> X,X
        df_resultado_kpis_grupo['GIRO AGO/25'] = df_resultado_kpis_grupo['GIRO AGO/25'].apply(
            lambda x: f"{x:.1f}".replace(".", ",") if pd.notnull(x) else "0,0"
        )
        # K: 'P. MÉDIO AGO/25' -> X,XX
        df_resultado_kpis_grupo['P. MÉDIO AGO/25'] = df_resultado_kpis_grupo['P. MÉDIO AGO/25'].apply(
            lambda x: f"{x:.2f}".replace(".", ",") if pd.notnull(x) else "0,00"
        )
        df_resultado_kpis_grupo.to_excel(os.path.join(rt_folder, f"{coluna_nome}_resultado_kpis_consolidado.xlsx"), index=False)
        print("Arquivo 'resultado_kpis_consolidado.xlsx' gerado com sucesso.")

    else:
        print("Configuração de COLUNA_NOME e COLUNA_GRUPO não reconhecida.")