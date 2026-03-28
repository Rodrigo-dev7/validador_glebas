"""
=============================================================================
VALIDADOR DE GLEBAS - SICOR
=============================================================================
Valida o erro:
  "SICOR: A gleba informada não corresponde a uma área válida"

Baseado no documento: Treinamento_GLEBAS_-_2026 (FSBR)

Causas do erro:
  1. Polígono NÃO fechado: o primeiro e o último ponto devem ser iguais
  2. Gleba com menos de 3 pontos únicos (mínimo exigido pelo SICOR)
  3. Pontos duplicados no meio da sequência (3+ pontos repetidos)

Formato esperado da planilha Excel:
  Coluna A: Número da Gleba     (ex: 1, 1, 1, 2, 2 ...)
  Coluna B: Sequência do Ponto  (ex: 1, 2, 3, 1, 2 ...)
  Coluna C: Latitude            (ex: -14.43539142600)
  Coluna D: Longitude           (ex: -44.33006286500)

Dependências:
  pip install pandas openpyxl xlrd

Uso:
  python validador_glebas.py                        # usa arquivo padrão
  python validador_glebas.py meu_arquivo.xlsx
  python validador_glebas.py meu_arquivo.xls
=============================================================================
"""

import sys
import os
import pandas as pd
from collections import defaultdict


# ===========================================================================
# CONFIGURAÇÕES
# ===========================================================================

# Nome do arquivo padrão (se não passar argumento na linha de comando)
ARQUIVO_PADRAO = "Coordenadas_Nicolas____.xls"

# Nomes possíveis das colunas (o programa tenta reconhecer automaticamente)
NOMES_COLUNA_GLEBA = ["gleba", "num_gleba", "nr_gleba", "sequencial_gleba", "gleba_seq", "sq_glb"]
NOMES_COLUNA_PONTO = ["ponto", "seq_ponto", "ordem_ponto", "nr_ponto", "sequencial_ponto", "sq_cgl"]
NOMES_COLUNA_LAT   = ["latitude", "lat", "nr_lat"]
NOMES_COLUNA_LON   = ["longitude", "lon", "lng", "nr_lon"]

# Tolerância para considerar dois pontos iguais (evita erros de float)
TOLERANCIA_COORDENADA = 1e-8


# ===========================================================================
# FUNÇÕES AUXILIARES
# ===========================================================================

def detectar_colunas(df: pd.DataFrame) -> dict:
    """
    Tenta detectar automaticamente quais colunas correspondem a
    gleba, ponto, latitude e longitude — seja pelo nome do cabeçalho
    ou pela posição (A, B, C, D).
    """
    colunas = {c.lower().strip().replace(" ", "_"): c for c in df.columns}
    mapeamento = {}

    def buscar(nomes_candidatos, rotulo):
        for nome in nomes_candidatos:
            if nome in colunas:
                return colunas[nome]
        return None

    mapeamento["gleba"]     = buscar(NOMES_COLUNA_GLEBA, "gleba")
    mapeamento["ponto"]     = buscar(NOMES_COLUNA_PONTO, "ponto")
    mapeamento["latitude"]  = buscar(NOMES_COLUNA_LAT,   "latitude")
    mapeamento["longitude"] = buscar(NOMES_COLUNA_LON,   "longitude")

    # Fallback: usar posição (0=gleba, 1=ponto, 2=lat, 3=lon)
    colunas_lista = list(df.columns)
    if mapeamento["gleba"]     is None and len(colunas_lista) > 0:
        mapeamento["gleba"]     = colunas_lista[0]
    if mapeamento["ponto"]     is None and len(colunas_lista) > 1:
        mapeamento["ponto"]     = colunas_lista[1]
    if mapeamento["latitude"]  is None and len(colunas_lista) > 2:
        mapeamento["latitude"]  = colunas_lista[2]
    if mapeamento["longitude"] is None and len(colunas_lista) > 3:
        mapeamento["longitude"] = colunas_lista[3]

    return mapeamento


def pontos_iguais(lat1, lon1, lat2, lon2) -> bool:
    """Compara dois pontos com tolerância de float."""
    try:
        return (abs(float(lat1) - float(lat2)) < TOLERANCIA_COORDENADA and
                abs(float(lon1) - float(lon2)) < TOLERANCIA_COORDENADA)
    except (ValueError, TypeError):
        return False


def carregar_planilha(caminho: str) -> pd.DataFrame:
    """Carrega o arquivo Excel (.xls ou .xlsx) em um DataFrame."""
    ext = os.path.splitext(caminho)[1].lower()

    if ext == ".xlsx":
        df = pd.read_excel(caminho, engine="openpyxl", header=0, dtype=str)
    elif ext == ".xls":
        df = pd.read_excel(caminho, engine="xlrd", header=0, dtype=str)
    else:
        raise ValueError(f"Formato não suportado: '{ext}'. Use .xls ou .xlsx")

    # Remove linhas completamente vazias
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ===========================================================================
# VALIDAÇÕES
# ===========================================================================

def validar_area_invalida(df: pd.DataFrame, cols: dict) -> list:
    """
    Valida o erro:
      'SICOR: A gleba informada não corresponde a uma área válida'

    Regras verificadas (conforme Treinamento GLEBAS 2026):

      R1 - Polígono fechado: o primeiro e último ponto devem ter
           coordenadas idênticas (lat e lon iguais).

      R2 - Mínimo de pontos únicos: além do ponto de fechamento,
           a gleba precisa ter ao menos 3 pontos distintos.
           (um triângulo é o menor polígono possível → 3 únicos + 1 repetido = 4 linhas)

      R3 - Ponto de fechamento duplicado em excesso: se mais de 2 pontos
           têm as mesmas coordenadas do primeiro ponto (ou qualquer
           duplicata excessiva no interior), a gleba é inválida.

    Retorna lista de dicts com os erros encontrados.
    """
    erros = []

    col_gleba = cols["gleba"]
    col_ponto = cols["ponto"]
    col_lat   = cols["latitude"]
    col_lon   = cols["longitude"]

    # Agrupa as linhas por número de gleba
    grupos = defaultdict(list)
    for idx, row in df.iterrows():
        num_gleba = str(row[col_gleba]).strip()
        if num_gleba and num_gleba.lower() not in ("nan", "none", ""):
            grupos[num_gleba].append({
                "linha_excel": idx + 2,   # +2: cabeçalho (linha 1) + índice 0-based
                "seq_ponto"  : row[col_ponto],
                "lat"        : row[col_lat],
                "lon"        : row[col_lon],
            })

    for num_gleba, pontos in grupos.items():
        n = len(pontos)
        prefixo = f"Gleba {num_gleba}"

        # Coleta coordenadas válidas
        coords_validas = []
        for p in pontos:
            try:
                lat = float(str(p["lat"]).replace(",", "."))
                lon = float(str(p["lon"]).replace(",", "."))
                coords_validas.append((lat, lon, p["linha_excel"]))
            except (ValueError, TypeError):
                erros.append({
                    "gleba"       : num_gleba,
                    "linha_excel" : p["linha_excel"],
                    "seq_ponto"   : p["seq_ponto"],
                    "tipo_erro"   : "COORDENADA INVÁLIDA",
                    "detalhe"     : (
                        f"Latitude='{p['lat']}' ou Longitude='{p['lon']}' "
                        f"não são números válidos."
                    ),
                })

        if not coords_validas:
            erros.append({
                "gleba"       : num_gleba,
                "linha_excel" : pontos[0]["linha_excel"],
                "seq_ponto"   : "-",
                "tipo_erro"   : "SEM COORDENADAS",
                "detalhe"     : f"{prefixo} não possui nenhuma coordenada numérica.",
            })
            continue

        # ---------------------------------------------------------------
        # R1 — Polígono fechado: primeiro == último
        # ---------------------------------------------------------------
        primeiro = coords_validas[0]
        ultimo   = coords_validas[-1]

        poligono_fechado = pontos_iguais(primeiro[0], primeiro[1],
                                         ultimo[0],   ultimo[1])

        if not poligono_fechado:
            erros.append({
                "gleba"       : num_gleba,
                "linha_excel" : ultimo[2],
                "seq_ponto"   : pontos[-1]["seq_ponto"],
                "tipo_erro"   : "POLÍGONO NÃO FECHADO",
                "detalhe"     : (
                    f"{prefixo}: o último ponto (linha {ultimo[2]}) "
                    f"[{ultimo[0]:.11f}, {ultimo[1]:.11f}] "
                    f"é DIFERENTE do primeiro ponto (linha {primeiro[2]}) "
                    f"[{primeiro[0]:.11f}, {primeiro[1]:.11f}]. "
                    "O SICOR exige que primeiro e último ponto sejam idênticos."
                ),
            })

        # ---------------------------------------------------------------
        # R2 — Mínimo de 3 pontos únicos
        # (desconsiderando o ponto de fechamento se existir)
        # ---------------------------------------------------------------
        coords_sem_fechamento = coords_validas[:-1] if poligono_fechado else coords_validas

        # Conta pontos únicos
        pontos_unicos = set(
            (round(lat, 8), round(lon, 8))
            for lat, lon, _ in coords_sem_fechamento
        )
        qtd_unicos = len(pontos_unicos)

        if qtd_unicos < 3:
            erros.append({
                "gleba"       : num_gleba,
                "linha_excel" : pontos[0]["linha_excel"],
                "seq_ponto"   : "-",
                "tipo_erro"   : "PONTOS INSUFICIENTES",
                "detalhe"     : (
                    f"{prefixo}: possui apenas {qtd_unicos} ponto(s) único(s) "
                    f"(mínimo exigido: 3). Total de linhas na planilha: {n}. "
                    "Um polígono válido precisa de ao menos 3 vértices distintos."
                ),
            })

        # ---------------------------------------------------------------
        # R3 — Pontos duplicados em excesso no interior
        # (3 ou mais ocorrências do mesmo par lat/lon)
        # ---------------------------------------------------------------
        contagem_coords = defaultdict(list)
        for lat, lon, linha in coords_validas:
            chave = (round(lat, 8), round(lon, 8))
            contagem_coords[chave].append(linha)

        for coord, linhas in contagem_coords.items():
            # O ponto de fechamento aparece 2x (início e fim) — isso é normal
            # Mais de 2x indica duplicata problemática
            if len(linhas) > 2:
                erros.append({
                    "gleba"       : num_gleba,
                    "linha_excel" : linhas[0],
                    "seq_ponto"   : "-",
                    "tipo_erro"   : "PONTO DUPLICADO EM EXCESSO",
                    "detalhe"     : (
                        f"{prefixo}: o ponto [{coord[0]:.11f}, {coord[1]:.11f}] "
                        f"aparece {len(linhas)} vezes (linhas: {linhas}). "
                        "Apenas o ponto de fechamento pode se repetir (1ª e última linha)."
                    ),
                })

    return erros


# ===========================================================================
# RELATÓRIO
# ===========================================================================

SEPARADOR = "=" * 75

def imprimir_relatorio(erros: list, arquivo: str, total_glebas: int):
    """Exibe o relatório de erros no terminal."""
    print()
    print(SEPARADOR)
    print("  RELATÓRIO DE VALIDAÇÃO DE GLEBAS — SICOR")
    print(f"  Arquivo : {arquivo}")
    print(f"  Glebas  : {total_glebas}")
    print(f"  Erros   : {len(erros)}")
    print(SEPARADOR)

    if not erros:
        print()
        print("  ✅  Nenhum erro encontrado! Todas as glebas estão válidas.")
        print()
        print(SEPARADOR)
        return

    # Agrupa por tipo de erro
    por_tipo = defaultdict(list)
    for e in erros:
        por_tipo[e["tipo_erro"]].append(e)

    for tipo, lista in sorted(por_tipo.items()):
        print()
        print(f"  ❌  {tipo}  ({len(lista)} ocorrência(s))")
        print("-" * 75)
        for e in lista:
            print(f"  Gleba       : {e['gleba']}")
            print(f"  Linha Excel : {e['linha_excel']}")
            if str(e.get("seq_ponto", "-")) not in ("-", "nan", "None", ""):
                print(f"  Seq. Ponto  : {e['seq_ponto']}")
            print(f"  Detalhe     : {e['detalhe']}")
            print()

    print(SEPARADOR)
    print()
    print("  RESUMO POR TIPO DE ERRO:")
    for tipo, lista in sorted(por_tipo.items()):
        print(f"    • {tipo}: {len(lista)} gleba(s)")
    print()
    print("  CORREÇÃO SUGERIDA:")
    print("    • POLÍGONO NÃO FECHADO   → Adicione no final da gleba uma linha")
    print("                               com as mesmas coordenadas da 1ª linha.")
    print("    • PONTOS INSUFICIENTES   → A gleba precisa de ao menos 3 vértices")
    print("                               distintos (4 linhas com fechamento).")
    print("    • PONTO DUPLICADO        → Remova os pontos repetidos no interior")
    print("                               da sequência da gleba.")
    print()
    print(SEPARADOR)


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    # Determina o arquivo a validar
    if len(sys.argv) > 1:
        caminho = sys.argv[1]
    else:
        caminho = ARQUIVO_PADRAO

    if not os.path.exists(caminho):
        print(f"\n❌  Arquivo não encontrado: '{caminho}'")
        print(f"    Uso: python validador_glebas.py <arquivo.xls ou arquivo.xlsx>")
        sys.exit(1)

    print(f"\n🔍  Lendo arquivo: {caminho} ...")

    # Carrega a planilha
    try:
        df = carregar_planilha(caminho)
    except Exception as ex:
        print(f"\n❌  Erro ao abrir o arquivo: {ex}")
        print("    Verifique se 'xlrd' (para .xls) ou 'openpyxl' (para .xlsx) está instalado:")
        print("    pip install xlrd openpyxl")
        sys.exit(1)

    print(f"    {len(df)} linhas de dados carregadas.")
    print(f"    Colunas encontradas: {list(df.columns)}")

    # Detecta as colunas
    cols = detectar_colunas(df)
    print(f"\n    Mapeamento de colunas detectado:")
    for rotulo, nome_col in cols.items():
        print(f"      {rotulo:<12}: '{nome_col}'")

    # Valida
    erros = validar_area_invalida(df, cols)

    # Total de glebas únicas
    col_gleba = cols["gleba"]
    total_glebas = df[col_gleba].dropna().astype(str).str.strip()
    total_glebas = total_glebas[~total_glebas.str.lower().isin(["nan", "none", ""])].nunique()

    # Imprime relatório
    imprimir_relatorio(erros, caminho, total_glebas)


if __name__ == "__main__":
    main()
