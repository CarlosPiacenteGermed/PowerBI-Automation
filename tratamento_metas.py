import pandas as pd
import os
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import numpy as np
from typing import Dict, List


load_dotenv()
username = os.getenv("USERNAME")

# Caminho da pasta Downloads e RT
downloads_folder = f"C:\\Users\\{username}\\Downloads"
rt_folder = os.path.join(downloads_folder, "RT")
os.makedirs(rt_folder, exist_ok=True)  # Cria a pasta RT se não existir

# Caminhos dos arquivos
bookmark_names = [name.strip() + ".xlsx" for name in os.getenv("BOOKMARKS_META", "").split(",") if name.strip()]
if len(bookmark_names) != 2:
    raise ValueError("A variável BOOKMARKS no .env deve conter exatamente 3 nomes separados por vírgula.")

path_sem_ge = os.path.join(downloads_folder, bookmark_names[0])
path_metas = os.path.join(downloads_folder, bookmark_names[1])

df_sem_ge = pd.read_excel(path_sem_ge, skiprows=2, engine="openpyxl")
df_metas = pd.read_excel(path_metas, skiprows=2, engine="openpyxl")



 #========= Helpers =========

def _ensure_columns(df: pd.DataFrame, cols: List[str], ctx: str = ""):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(f"As colunas {missing} não foram encontradas no DataFrame {ctx}.")

def _coerce_numeric(df: pd.DataFrame, cols: List[str]):
    """
    Converte colunas para numérico (robusto para formatações BR, símbolos e pontuações).
    """
    for c in cols:
        if c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]):
                continue
            s = df[c].astype(str)
            s = s.str.replace(r"[^\d\-,.\(\)]", "", regex=True)
            s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
            df[c] = pd.to_numeric(s, errors="coerce")

def _coverage(real: pd.Series, meta: pd.Series) -> pd.Series:
    """
    Retorna uma Series (alinhada ao índice) com Cobertura = Real/Meta quando Meta != 0, senão NaN.
    """
    real = pd.to_numeric(real, errors="coerce")
    meta = pd.to_numeric(meta, errors="coerce")
    cobertura = real.divide(meta)*100
    cobertura = cobertura.where(meta.ne(0) & meta.notna(), np.nan)
    return cobertura

# ========= Configuração dos nomes de colunas =========

CATEGORIAS = {
    "OL": {
        "meta": "($)META_OL",
        "real": "($)OL",
        "componentes": ["MERCANET", "PHARMALINK"],  # se ($)OL não existir, somamos estes
        "delta": "'Medidas'[($)OLDelta]",
        "cobertura": "(%)_OL",
    },
    "PPP_TOTAL": {
        "meta": "($)META_PPP",
        "real": "PPP",
        "delta": "'Medidas'[($)DeltaMDTRAssoc]",
        "cobertura": "(%)_PPP",
    },
    "PPP_N_COMBATE": {
        "meta": "'Medidas'[($)META_PPP_N_COMBATE_SO]",
        "real": "N COMBATE",
        "delta": "'Medidas'[($)Delta_N_Combate_SO]",
        "cobertura": "'Medidas'[(%)COB_N_COMBATE_SO]",
    },
    "PPP_COMBATE": {
        "meta": "'Medidas'[($)META_PPP_COMBATE_SO]",
        "real": "COMBATE",
        "delta": "'Medidas'[($)DeltaCombate_SO]",
        "cobertura": "'Medidas'[(%)COB_COMBATE_SO]",
    },
    "PPP_MIX": {
        "meta": "($)META_PPP_MIX",
        "real": "MIX",
        "delta": "'Medidas'[($)DemandaDeltaMixFocoAssoc]",
        "cobertura": "(%)_PPP_MIX",
    },
    "PPP_LANC": {
        "meta": "($)META_PPP_LAN",
        "real": "LANÇ",
        "delta": "'Medidas'[($)DemandaDeltaLaNCFocoAssoc]",
        "cobertura": "(%)_PPP_LAN",
    },
}

BASE_COL_ORDER = [
    "NOME GD", "NOME REP",
    "($)META_OL","($)OL",
    "'Medidas'[($)OLDelta]", "(%)_OL",
    "($)META_PPP", "PPP", "'Medidas'[($)DeltaMDTRAssoc]", "(%)_PPP",
    "($)META_PPP_MIX", "MIX", "'Medidas'[($)DemandaDeltaMixFocoAssoc]", "(%)_PPP_MIX",
    "($)META_PPP_LAN", "LANÇ", "'Medidas'[($)DemandaDeltaLaNCFocoAssoc]", "(%)_PPP_LAN",
]

def build_tabela_base(df_sem_ge: pd.DataFrame) -> pd.DataFrame:
    out = df_sem_ge.copy()

    # Cria ($)OL a partir de MERCANET+PHARMALINK se necessário
    cat_ol = CATEGORIAS["OL"]
    if cat_ol["real"] not in out.columns:
        comps = [c for c in cat_ol.get("componentes", []) if c in out.columns]
        if len(comps) == 2:
            _coerce_numeric(out, comps)
            out[cat_ol["real"]] = out[comps[0]].fillna(0) + out[comps[1]].fillna(0)

    # Converte numérico
    possiveis_numericas = set()
    for cfg in CATEGORIAS.values():
        possiveis_numericas.update([
            cfg.get("meta", ""), cfg.get("real", ""),
            cfg.get("delta", ""), cfg.get("cobertura", "")
        ])
        for comp in cfg.get("componentes", []):
            possiveis_numericas.add(comp)
    possiveis_numericas = [c for c in possiveis_numericas if c and c in out.columns]
    _coerce_numeric(out, possiveis_numericas)

    # Calcula deltas/coberturas ausentes
    for cfg in CATEGORIAS.values():
        meta_col, real_col = cfg["meta"], cfg["real"]
        delta_col, cob_col = cfg["delta"], cfg["cobertura"]
        if meta_col not in out.columns or real_col not in out.columns:
            continue
        if delta_col not in out.columns:
            out[delta_col] = out[real_col].fillna(0) - out[meta_col].fillna(0)
        if cob_col not in out.columns:
            out[cob_col] = _coverage(out[real_col], out[meta_col])

    for comp_col in ["MERCANET", "PHARMALINK"]:
        if comp_col not in out.columns:
            out[comp_col] = np.nan

    # Colunas obrigatórias
    for c in ["NOME GD", "NOME REP"]:
        if c not in out.columns:
            raise KeyError(f"A coluna obrigatória '{c}' não está presente após o preparo.")

    # Garante layout
    for col in BASE_COL_ORDER:
        if col not in out.columns:
            out[col] = np.nan

    out = out[BASE_COL_ORDER]
    return out

# ========= Tabelas finais =========

def _agg_por_grupo(tabela_base: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """
    Agrega por grupo (GD ou REP) gerando somatórios para metas/real das 4 categorias pedidas.
    Retorna um DF com colunas: META/REAL de cada categoria.
    """
    cat_cols = {
        "Demanda PPP": ("($)META_PPP", "PPP"),
        "Lançamentos": ("($)META_PPP_LAN", "LANÇ"),
        "Mix Foco": ("($)META_PPP_MIX", "MIX"),
        "OL": ("($)META_OL", "($)OL"),
    }

    numeric_cols = []
    for meta_col, real_col in cat_cols.values():
        if meta_col in tabela_base.columns: numeric_cols.append(meta_col)
        if real_col in tabela_base.columns: numeric_cols.append(real_col)
    _coerce_numeric(tabela_base, numeric_cols)

    agg_dict = {}
    for meta_col, real_col in cat_cols.values():
        if meta_col in tabela_base.columns:
            agg_dict[meta_col] = "sum"
        if real_col in tabela_base.columns:
            agg_dict[real_col] = "sum"

    agg = tabela_base.groupby(group_col, dropna=False).agg(agg_dict)

    # Garante colunas mesmo se não existirem na origem
    for nome_cat, (meta_col, real_col) in cat_cols.items():
        if meta_col not in agg.columns: agg[meta_col] = np.nan
        if real_col not in agg.columns: agg[real_col] = np.nan

    return agg

def build_matriz_por_grupo(tabela_base: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """
    Retorna matriz no layout solicitado:
      - Colunas: ['Demanda PPP','Lançamentos','Mix Foco','OL']
      - Linhas: para cada grupo (GD ou REP), 3 linhas: ['Objetivo (Meta)', 'Real', 'Cobertura']
    """
    agg = _agg_por_grupo(tabela_base, group_col=group_col)
    cat_cols = {
        "Demanda PPP": ("($)META_PPP", "PPP"),
        "Lançamentos": ("($)META_PPP_LAN", "LANÇ"),
        "Mix Foco": ("($)META_PPP_MIX", "MIX"),
        "OL": ("($)META_OL", "($)OL"),
    }

    # Monta um DF com colunas MultiIndex (Categoria, Métrica)
    frames = []
    for nome_cat, (meta_col, real_col) in cat_cols.items():
        objetivo = agg[meta_col]
        real = agg[real_col]
        cob = _coverage(real, objetivo)

        bloco = pd.concat(
            {
                (nome_cat, "Objetivo (Meta)"): objetivo,
                (nome_cat, "Real"): real,
                (nome_cat, "Cobertura"): cob,
            },
            axis=1,
        )
        frames.append(bloco)

    wide = pd.concat(frames, axis=1)

    # Reordena colunas de categoria e métricas
    wide = wide.reindex(
        columns=pd.MultiIndex.from_product(
            [["Demanda PPP", "Lançamentos", "Mix Foco", "OL"], ["Objetivo (Meta)", "Real", "Cobertura"]]
        )
    )

    # Transforma as métricas em linhas (stack do nível 1 das colunas)
    out = wide.stack(level=1, future_stack=True)  # index -> (group_col, 'Linha'); columns -> categorias
    out.index = out.index.set_names([group_col, "Linha"])
    return out

def build_metas_total_gd(tabela_base: pd.DataFrame) -> pd.DataFrame:
    """Matriz por GD no layout solicitado."""
    _ensure_columns(tabela_base, ["NOME GD"], "tabela_base")
    return build_matriz_por_grupo(tabela_base, group_col="NOME GD")

def build_metas_gr_completa(tabela_base: pd.DataFrame) -> pd.DataFrame:
    """Matriz por Representante (GR) no layout solicitado."""
    _ensure_columns(tabela_base, ["NOME REP"], "tabela_base")
    return build_matriz_por_grupo(tabela_base, group_col="NOME REP")

def build_metas_gr_ppp(tabela_base: pd.DataFrame) -> pd.DataFrame:
    """
    METAS GR (por representante), somente PPP:
    - Colunas: META PPP, Demanda PPP, Desv. Abs., Cobertura
    - Índice: NOME REP
    - Adiciona linha de total com soma das colunas e cálculo de desvio e cobertura
    """
    _ensure_columns(tabela_base, ["NOME REP"], "tabela_base")

    cols_need = []
    if "($)META_PPP" in tabela_base.columns: cols_need.append("($)META_PPP")
    if "PPP" in tabela_base.columns: cols_need.append("PPP")
    _coerce_numeric(tabela_base, cols_need)

    agg = (
        tabela_base.groupby("NOME REP", dropna=False)
        .agg({"($)META_PPP": "sum", "PPP": "sum"})
        .rename(columns={"($)META_PPP": "META PPP", "PPP": "Demanda PPP"})
    )
    for col in ["META PPP", "Demanda PPP"]:
        if col not in agg.columns:
            agg[col] = np.nan

    agg["Desv. Abs."] = agg["Demanda PPP"] - agg["META PPP"]
    agg["Cobertura"] = _coverage(agg["Demanda PPP"], agg["META PPP"])
    agg = agg[["META PPP", "Demanda PPP", "Desv. Abs.", "Cobertura"]]

    # Adiciona linha de total
    total_meta = agg["META PPP"].sum()
    total_demanda = agg["Demanda PPP"].sum()
    total_desvio = total_demanda - total_meta
    total_cobertura = (total_demanda/total_meta)*100

    total_row = pd.DataFrame({
        "META PPP": [total_meta],
        "Demanda PPP": [total_demanda],
        "Desv. Abs.": [total_desvio],
        "Cobertura": [total_cobertura]
    }, index=["TOTAL"])

    agg = pd.concat([agg, total_row])

    return agg

def build_top8_grupo_por_rep(df: pd.DataFrame) -> dict:
    """
    Gera, para cada representante (NOME REP), um DataFrame com:
      - Top 7 grupos na ordem vinda do Excel (order_map)
      - Linha 'Outros' (soma do restante)
      - Linha 'TOTAL'
      - COB% recalculado no final

    Observação: usa a função `_coverage(dem, meta)` já existente no seu código.
    """
    df = df.rename(columns={"($)META_PPP": "META PPP AGO/25", "PPP": "DEM. PPP"})
    df["META PPP AGO/25"] = pd.to_numeric(df["META PPP AGO/25"], errors="coerce")
    df["DEM. PPP"] = pd.to_numeric(df["DEM. PPP"], errors="coerce")

    resultado = {}
    meta_coluna = os.getenv("COLUNA_META")
    for rep, grupo_rep in df.groupby("NOME REP"):
        # Caminho do arquivo YTD do representante
        ytd_file = os.path.join(
            r"C:\Users\c0050485\Downloads\RT",
            f"tabela_top7_YTD_{str(rep).split()[0].upper()}.xlsx"
        )
        # Lê a ordem dos grupos da coluna A do arquivo YTD
        if os.path.exists(ytd_file):
            try:
                ordem_grupos = pd.read_excel(ytd_file, usecols=[0], skiprows=1, header=None).iloc[:, 0].astype(str).tolist()
            except Exception:
                ordem_grupos = []
        else:
            ordem_grupos = []

        grupo_agg = grupo_rep.groupby(meta_coluna, as_index=True).agg({
            "META PPP AGO/25": "sum",
            "DEM. PPP": "sum"
        })

        grupo_agg["DESV. ABS"] = grupo_agg["DEM. PPP"] - grupo_agg["META PPP AGO/25"]
        grupo_agg["COB%"] = _coverage(grupo_agg["DEM. PPP"], grupo_agg["META PPP AGO/25"])

        # Ordena conforme ordem_grupos, mantendo os demais ao final
        grupos_presentes = [g for g in ordem_grupos if g in grupo_agg.index]
        restantes = [g for g in grupo_agg.index if g not in grupos_presentes]
        grupo_agg = grupo_agg.loc[grupos_presentes + restantes]

        # === Top7 + Outros + Total ===
        if len(grupo_agg) <= 8:
            final_df = grupo_agg.copy()
        else:
            top7 = grupo_agg.iloc[:7]
            numeric_cols = grupo_agg.select_dtypes(include="number").columns.tolist()
            numeric_cols_no_cob = [c for c in numeric_cols if c != "COB%"]

            total = grupo_agg[numeric_cols_no_cob].sum()
            soma_top7 = top7[numeric_cols_no_cob].sum()

            outros = (total - soma_top7).to_frame().T
            outros.index = ["Outros"]

            final_df = pd.concat([top7, outros], axis=0)

            total_final = final_df[numeric_cols_no_cob].sum()
            linha_total = total_final.to_frame().T
            linha_total.index = ["TOTAL"]

            final_df = pd.concat([final_df, linha_total], axis=0)

            # Recalcula COB% para todas as linhas do final_df
            final_df["COB%"] = _coverage(final_df["DEM. PPP"], final_df["META PPP AGO/25"])

        resultado[rep] = final_df

    return resultado

# ========= Execução: preparar df_sem_ge e gerar as tabelas =========


# 2) Tabela base com todas as colunas/deltas/coberturas
tabela_base = build_tabela_base(df_sem_ge)
tabela_base_metas = build_tabela_base(df_metas)
ytd_path = r"C:\Users\c0050485\Downloads\RT\tabela_ytd.xlsx"
ytd_nomes = pd.read_excel(ytd_path, usecols=[0], skiprows=1, header=None).iloc[:, 0].astype(str).tolist()
# 3) As três saídas exatamente no layout pedido
metas_total_gd = build_metas_total_gd(tabela_base)          # Index: (NOME GD, Linha) | Colunas: categorias
metas_gr = build_metas_gr_ppp(tabela_base)                  # Index: NOME REP       | Colunas: PPP
metas_gr_completa = build_metas_gr_completa(tabela_base)    # Index: (NOME REP, Linha) | Colunas: categorias
primeiro_nome_map = {str(nome).split()[0].upper(): nome for nome in metas_gr.index if nome != "TOTAL"}
ordem_nomes = []
for primeiro_nome in ytd_nomes:
    nome_upper = str(primeiro_nome).upper()
    if nome_upper in primeiro_nome_map:
        ordem_nomes.append(primeiro_nome_map[nome_upper])

# Adiciona os restantes (que não estão na lista de primeiros nomes)
restantes = [i for i in metas_gr.index if i not in ordem_nomes and i != "TOTAL"]

# Remove TOTAL antes de reordenar
if "TOTAL" in metas_gr.index:
    total_row = metas_gr.loc[["TOTAL"]]
    metas_gr = metas_gr.drop("TOTAL")
else:
    total_row = None

# Reordena e adiciona TOTAL ao final
metas_gr = metas_gr.loc[ordem_nomes + restantes]
if total_row is not None:
    metas_gr = pd.concat([metas_gr, total_row])

# 4) Salvar em Excel (pasta RT)

# 5) Exporta METAS_PROCESSADAS_GE com abas por coordenador

out_path = os.path.join(rt_folder, "METAS_PROCESSADAS.xlsx")
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    tabela_base.to_excel(writer, sheet_name="BASE_METAS", index=False)
    metas_total_gd.to_excel(writer, sheet_name="METAS_TOTAL_GD")        # tem índice MultiIndex (NOME GD, Linha)
    metas_gr.to_excel(writer, sheet_name="METAS_GR")                    # índice: NOME REP
    metas_gr_completa.to_excel(writer, sheet_name="METAS_GR_COMPLETA") 


top8_por_rep = build_top8_grupo_por_rep(df_metas)
out_path = os.path.join(rt_folder, "METAS_GRUPO_ECONOMICO_TOP8.xlsx")
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    for rep, df_rep in top8_por_rep.items():
        aba_nome = str(rep)[:31] if isinstance(rep, str) else "REP"
        df_rep.to_excel(writer, sheet_name=aba_nome)

print(f"Arquivos salvos em: {out_path}")

 #Caminho do arquivo gerado anteriormente
arquivo_metas = os.path.join(rt_folder, "METAS_PROCESSADAS.xlsx")


# Formatos
formato_milhar        = '#,##0'    # → X.XXX.XXX (sem casas decimais)
formato_percent_ratio = '0.0%'     # → XX,X% quando o VALOR está entre 0 e 1 (ex.: 0,729)
formato_percent_abs   = '0.0"%"'   # → XX,X% quando o VALOR já está entre 0 e 100 (ex.: 72,9)

wb = load_workbook(arquivo_metas)

def _fmt_percent_por_valor(cell):
    """
    Se o valor estiver em [0, 1.0001] → usar '0.0%'
    Se o valor > 1 → usar '0.0"%"' (não escala)
    Obs.: tolerância pequena por conta de arredondamentos.
    """
    v = cell.value
    if isinstance(v, (int, float)):
        if 0 <= v <= 1.0001:
            cell.number_format = formato_percent_ratio   # 0.0%  (escala por 100)
        else:
            cell.number_format = formato_percent_abs     # 0.0"%" (não escala)
    else:
        # Se for vazio ou string, apenas define o formato desejado
        cell.number_format = formato_percent_abs

# ----------------------------------------------------
# 1) METAS_TOTAL_GD: Col A = NOME GD | Col B = Linha | Col C.. = categorias
# ----------------------------------------------------
if 'METAS_TOTAL_GD' in wb.sheetnames:
    ws = wb['METAS_TOTAL_GD']
    ultima_linha = ws.max_row
    ultima_coluna = ws.max_column

    cols_categorias = [get_column_letter(c) for c in range(3, ultima_coluna + 1)]

    for row in range(2, ultima_linha + 1):  # pula cabeçalho
        linha_val = ws[f'B{row}'].value
        linha_norm = (str(linha_val).strip().lower() if linha_val is not None else '')
        is_cobertura = (linha_norm == 'cobertura')

        for col in cols_categorias:
            cell = ws[f'{col}{row}']
            if cell.value is None or cell.value == "":
                continue
            if is_cobertura:
                _fmt_percent_por_valor(cell)      # ↩️ autodetecta 0–1 vs 0–100
            else:
                cell.number_format = formato_milhar

# ----------------------------------------------------
# 2) METAS_GR_COMPLETA: Col A = NOME REP | Col B = Linha | Col C.. = categorias
# ----------------------------------------------------
if 'METAS_GR_COMPLETA' in wb.sheetnames:
    ws = wb['METAS_GR_COMPLETA']
    ultima_linha = ws.max_row
    ultima_coluna = ws.max_column

    cols_categorias = [get_column_letter(c) for c in range(3, ultima_coluna + 1)]

    for row in range(2, ultima_linha + 1):
        linha_val = ws[f'B{row}'].value
        linha_norm = (str(linha_val).strip().lower() if linha_val is not None else '')
        is_cobertura = (linha_norm == 'cobertura')

        for col in cols_categorias:
            cell = ws[f'{col}{row}']
            if cell.value is None or cell.value == "":
                continue
            if is_cobertura:
                _fmt_percent_por_valor(cell)
            else:
                cell.number_format = formato_milhar

# ----------------------------------------------------
# 3) METAS_GR: Col A = NOME REP | Depois: META PPP, Demanda PPP, Desv. Abs., Cobertura
# ----------------------------------------------------
if 'METAS_GR' in wb.sheetnames:
    ws = wb['METAS_GR']
    ultima_linha = ws.max_row
    ultima_coluna = ws.max_column

    # Mapa de cabeçalhos
    header_to_col = {}
    for c in range(1, ultima_coluna + 1):
        letra = get_column_letter(c)
        header = ws[f'{letra}1'].value
        header_to_col[header] = letra

    # Milhar nas colunas numéricas
    for nome_col in ['META PPP', 'Demanda PPP', 'Desv. Abs.']:
        col_letter = header_to_col.get(nome_col)
        if col_letter:
            for row in range(2, ultima_linha + 1):
                cell = ws[f'{col_letter}{row}']
                if cell.value is not None and cell.value != "":
                    cell.number_format = formato_milhar

    # Cobertura com autodetecção de escala
    col_cov = header_to_col.get('Cobertura')
    if col_cov:
        for row in range(2, ultima_linha + 1):
            cell = ws[f'{col_cov}{row}']
            if cell.value is not None and cell.value != "":
                _fmt_percent_por_valor(cell)

wb.save(arquivo_metas)
print(f"Arquivo formatado: {arquivo_metas}")

# Caminho do arquivo gerado
arquivo_metas_GE = os.path.join(rt_folder, "METAS_GRUPO_ECONOMICO_TOP8.xlsx")

formato_milhar        = '#,##0'    # → X.XXX.XXX (sem casas decimais)
formato_percent_ratio = '0.0%'     # → XX,X% quando o VALOR está entre 0 e 1 (ex.: 0,729)
formato_percent_abs   = '0.0"%"'  

def _fmt_percent_por_valor(cell):
    """
    Se o valor estiver em [0, 1.0001] → usar '0.0%'
    Se o valor > 1 → usar '0.0"%"' (não escala)
    Obs.: tolerância pequena por conta de arredondamentos.
    """
    v = cell.value
    if isinstance(v, (int, float)):
        if 0 <= v <= 1.0001:
            cell.number_format = formato_percent_ratio   # 0.0%  (escala por 100)
        else:
            cell.number_format = formato_percent_abs     # 0.0"%" (não escala)
    else:
        # Se for vazio ou string, apenas define o formato desejado
        cell.number_format = formato_percent_abs

# Carrega o arquivo
wb = load_workbook(arquivo_metas_GE)

# Aplica formatação em todas as abas
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    ultima_linha = ws.max_row
    ultima_coluna = ws.max_column

    # Mapa de cabeçalhos
    header_to_col = {}
    for c in range(1, ultima_coluna + 1):
        letra = get_column_letter(c)
        header = ws[f'{letra}1'].value
        header_to_col[header] = letra

    # Milhar nas colunas numéricas
    for nome_col in ['META PPP AGO/25', 'DEM. PPP', 'DESV. ABS']:
        col_letter = header_to_col.get(nome_col)
        if col_letter:
            for row in range(2, ultima_linha + 1):
                cell = ws[f'{col_letter}{row}']
                if cell.value is not None and cell.value != "":
                    cell.number_format = formato_milhar

    # Cobertura com autodetecção de escala
    col_cov = header_to_col.get('COB%')
    if col_cov:
        for row in range(2, ultima_linha + 1):
            cell = ws[f'{col_cov}{row}']
            if cell.value is not None and cell.value != "":
                _fmt_percent_por_valor(cell)

# Salva o arquivo
wb.save(arquivo_metas_GE)
print(f"Arquivo formatado: {arquivo_metas_GE}")