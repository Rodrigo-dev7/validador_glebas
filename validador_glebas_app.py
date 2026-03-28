"""
=============================================================================
VALIDADOR DE GLEBAS — Aplicativo Desktop (CustomTkinter)
Banco do Nordeste / SICOR — Treinamento GLEBAS 2026 (FSBR)
=============================================================================
Dependências:
    pip install pandas openpyxl xlrd customtkinter

Para gerar o executável (.exe):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name ValidadorGlebas validador_glebas_app.py
=============================================================================
"""

import os
import threading
from collections import defaultdict
from datetime import datetime

import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk
import pandas as pd


# ============================================================
# TEMA
# ============================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

AZUL_VIVO   = "#3B82F6"
AZUL_CLARO  = "#60A5FA"
VERDE       = "#22C55E"
VERDE_DIM   = "#16A34A"
VERMELHO    = "#EF4444"
AMARELO     = "#F59E0B"
CINZA_CARD  = "#1E2533"
CINZA_FUNDO = "#161B27"
CINZA_BORDA = "#2D3748"
BRANCO      = "#F1F5F9"
DIM         = "#64748B"
ROXO        = "#7C3AED"


# ============================================================
# LÓGICA DE VALIDAÇÃO
# ============================================================

NOMES_GLEBA = ["gleba","num_gleba","nr_gleba","sequencial_gleba","gleba_seq","sq_glb"]
NOMES_PONTO = ["ponto","seq_ponto","ordem_ponto","nr_ponto","sequencial_ponto","sq_cgl"]
NOMES_LAT   = ["latitude","lat","nr_lat"]
NOMES_LON   = ["longitude","lon","lng","nr_lon"]
TOLERANCIA  = 1e-8


def detectar_colunas(df):
    cols = {c.lower().strip().replace(" ", "_"): c for c in df.columns}
    def buscar(nomes):
        for n in nomes:
            if n in cols:
                return cols[n]
        return None
    lista = list(df.columns)
    return {
        "gleba"    : buscar(NOMES_GLEBA) or (lista[0] if len(lista) > 0 else None),
        "ponto"    : buscar(NOMES_PONTO) or (lista[1] if len(lista) > 1 else None),
        "latitude" : buscar(NOMES_LAT)   or (lista[2] if len(lista) > 2 else None),
        "longitude": buscar(NOMES_LON)   or (lista[3] if len(lista) > 3 else None),
    }


def pontos_iguais(la1, lo1, la2, lo2):
    try:
        return (abs(float(la1) - float(la2)) < TOLERANCIA and
                abs(float(lo1) - float(lo2)) < TOLERANCIA)
    except Exception:
        return False


def carregar_planilha(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".xlsx":
        df = pd.read_excel(caminho, engine="openpyxl", header=0, dtype=str)
    elif ext == ".xls":
        df = pd.read_excel(caminho, engine="xlrd", header=0, dtype=str)
    else:
        raise ValueError(f"Formato '{ext}' não suportado. Use .xls ou .xlsx")
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def validar(df, cols):
    erros  = []
    grupos = defaultdict(list)

    for idx, row in df.iterrows():
        g = str(row[cols["gleba"]]).strip()
        if g and g.lower() not in ("nan", "none", ""):
            grupos[g].append({
                "linha": idx + 2,
                "seq"  : row[cols["ponto"]],
                "lat"  : row[cols["latitude"]],
                "lon"  : row[cols["longitude"]],
            })

    for num, pontos in grupos.items():
        coords = []
        for p in pontos:
            try:
                lat = float(str(p["lat"]).replace(",", "."))
                lon = float(str(p["lon"]).replace(",", "."))
                coords.append((lat, lon, p["linha"]))
            except Exception:
                erros.append({
                    "gleba": num, "linha": p["linha"], "seq": p["seq"],
                    "tipo" : "COORDENADA INVÁLIDA",
                    "msg"  : f"Lat='{p['lat']}' ou Lon='{p['lon']}' não é numérico.",
                })

        if not coords:
            continue

        # R1 — Polígono fechado
        fechado = pontos_iguais(*coords[0][:2], *coords[-1][:2])
        if not fechado:
            erros.append({
                "gleba": num, "linha": coords[-1][2], "seq": pontos[-1]["seq"],
                "tipo" : "POLÍGONO NÃO FECHADO",
                "msg"  : ("Último ponto ≠ Primeiro ponto. "
                          "Adicione ao final uma linha igual à 1ª."),
            })

        # R2 — Mínimo de 3 pontos únicos
        sem_fech = coords[:-1] if fechado else coords
        unicos   = {(round(la, 8), round(lo, 8)) for la, lo, _ in sem_fech}
        if len(unicos) < 3:
            erros.append({
                "gleba": num, "linha": pontos[0]["linha"], "seq": "-",
                "tipo" : "PONTOS INSUFICIENTES",
                "msg"  : f"Apenas {len(unicos)} vértice(s) único(s). Mínimo: 3.",
            })

        # R3 — Duplicatas em excesso
        cont = defaultdict(list)
        for la, lo, ln in coords:
            cont[(round(la, 8), round(lo, 8))].append(ln)
        for coord, linhas in cont.items():
            if len(linhas) > 2:
                erros.append({
                    "gleba": num, "linha": linhas[0], "seq": "-",
                    "tipo" : "PONTO DUPLICADO EM EXCESSO",
                    "msg"  : (f"Ponto [{coord[0]:.6f}, {coord[1]:.6f}] "
                              f"aparece {len(linhas)}× (linhas {linhas[:4]})."),
                })

    return erros, grupos


# ============================================================
# COMPONENTES VISUAIS
# ============================================================

class CartaoEstatistica(ctk.CTkFrame):
    def __init__(self, parent, label, valor="—", cor=AZUL_VIVO, **kw):
        super().__init__(parent, fg_color=CINZA_CARD, corner_radius=12, **kw)
        self._cor = cor
        self.lbl_valor = ctk.CTkLabel(
            self, text=valor,
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color=cor
        )
        self.lbl_valor.pack(padx=16, pady=(14, 2))
        ctk.CTkLabel(
            self, text=label,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=DIM
        ).pack(padx=16, pady=(0, 14))

    def atualizar(self, valor, cor=None):
        self.lbl_valor.configure(text=str(valor), text_color=cor or self._cor)


class ZonaDrop(ctk.CTkFrame):
    def __init__(self, parent, on_click, **kw):
        super().__init__(
            parent, fg_color=CINZA_CARD, corner_radius=16,
            border_width=2, border_color=CINZA_BORDA, **kw
        )
        self._on_click = on_click
        self._tem_arq  = False

        self.icone = ctk.CTkLabel(self, text="📂", font=ctk.CTkFont(size=40))
        self.icone.pack(pady=(24, 6))

        self.lbl_principal = ctk.CTkLabel(
            self, text="Clique aqui para selecionar o arquivo Excel",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=BRANCO
        )
        self.lbl_principal.pack()

        self.lbl_sub = ctk.CTkLabel(
            self, text="Formatos aceitos:  .xls   •   .xlsx",
            font=ctk.CTkFont(size=11), text_color=DIM
        )
        self.lbl_sub.pack(pady=(4, 24))

        for w in [self, self.icone, self.lbl_principal, self.lbl_sub]:
            w.bind("<Button-1>", lambda e: self._on_click())
            w.bind("<Enter>",    self._enter)
            w.bind("<Leave>",    self._leave)

    def _enter(self, e=None):
        if not self._tem_arq:
            self.configure(border_color=AZUL_VIVO, fg_color="#232D40")

    def _leave(self, e=None):
        if not self._tem_arq:
            self.configure(border_color=CINZA_BORDA, fg_color=CINZA_CARD)

    def set_arquivo(self, nome):
        self._tem_arq = True
        self.configure(border_color=VERDE, fg_color="#1a2e22")
        self.icone.configure(text="✅")
        self.lbl_principal.configure(text=nome, text_color=VERDE)
        self.lbl_sub.configure(text="Arquivo selecionado  •  Clique para trocar")

    def reset(self):
        self._tem_arq = False
        self.configure(border_color=CINZA_BORDA, fg_color=CINZA_CARD)
        self.icone.configure(text="📂")
        self.lbl_principal.configure(
            text="Clique aqui para selecionar o arquivo Excel", text_color=BRANCO)
        self.lbl_sub.configure(
            text="Formatos aceitos:  .xls   •   .xlsx", text_color=DIM)


class BadgeTipo(ctk.CTkFrame):
    CORES = {
        "POLÍGONO NÃO FECHADO"      : ("#7F1D1D", VERMELHO),
        "PONTOS INSUFICIENTES"      : ("#78350F", AMARELO),
        "PONTO DUPLICADO EM EXCESSO": ("#713F12", "#FCD34D"),
        "COORDENADA INVÁLIDA"       : ("#7F1D1D", VERMELHO),
    }

    def __init__(self, parent, tipo, **kw):
        bg, fg = self.CORES.get(tipo, ("#1e293b", DIM))
        super().__init__(parent, fg_color=bg, corner_radius=6, **kw)
        ctk.CTkLabel(
            self, text=tipo,
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=fg
        ).pack(padx=8, pady=3)


# ============================================================
# JANELA PRINCIPAL
# ============================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Validador de Glebas — SICOR")
        self.geometry("1020x740")
        self.minsize(800, 580)
        self.configure(fg_color=CINZA_FUNDO)

        self.update_idletasks()
        w, h = 1020, 740
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        self._arquivo      = None
        self._erros_cache  = []
        self._grupos_cache = {}

        self._build()

    # ── LAYOUT ───────────────────────────────────────────────
    def _build(self):
        # Sidebar
        self.sidebar = ctk.CTkFrame(
            self, width=240, fg_color="#111827", corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        # Conteúdo
        self.content = ctk.CTkFrame(self, fg_color=CINZA_FUNDO, corner_radius=0)
        self.content.pack(side="left", fill="both", expand=True)
        self._build_content()

    # ── SIDEBAR ──────────────────────────────────────────────
    def _build_sidebar(self):
        sb = self.sidebar

        ctk.CTkFrame(sb, fg_color=AZUL_VIVO, height=4, corner_radius=0).pack(fill="x")

        # Logo
        logo = ctk.CTkFrame(sb, fg_color="transparent")
        logo.pack(fill="x", padx=20, pady=(20, 8))
        ctk.CTkLabel(logo, text="🌾", font=ctk.CTkFont(size=32)).pack(anchor="w")
        ctk.CTkLabel(
            logo, text="Validador\nde Glebas",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=BRANCO, justify="left"
        ).pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(
            logo, text="SICOR  •  BNB",
            font=ctk.CTkFont(size=11), text_color=DIM
        ).pack(anchor="w")

        ctk.CTkFrame(sb, fg_color=CINZA_BORDA, height=1, corner_radius=0
                     ).pack(fill="x", padx=16, pady=16)

        # Cartões de estatística
        ctk.CTkLabel(
            sb, text="RESUMO",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=DIM
        ).pack(anchor="w", padx=20, pady=(0, 10))

        sf = ctk.CTkFrame(sb, fg_color="transparent")
        sf.pack(fill="x", padx=12)

        self.card_glebas = CartaoEstatistica(sf, "Glebas",   "—", AZUL_CLARO)
        self.card_glebas.pack(fill="x", pady=(0, 8))

        self.card_erros = CartaoEstatistica(sf, "Erros",    "—", DIM)
        self.card_erros.pack(fill="x", pady=(0, 8))

        self.card_ok = CartaoEstatistica(sf, "Sem erro", "—", VERDE)
        self.card_ok.pack(fill="x")

        ctk.CTkFrame(sb, fg_color=CINZA_BORDA, height=1, corner_radius=0
                     ).pack(fill="x", padx=16, pady=20)

        # Botões
        ctk.CTkLabel(
            sb, text="AÇÕES",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=DIM
        ).pack(anchor="w", padx=20, pady=(0, 10))

        btn_cfg = dict(height=38, corner_radius=10,
                       font=ctk.CTkFont(size=12, weight="bold"))

        self.btn_buscar = ctk.CTkButton(
            sb, text="🗂  Buscar arquivo",
            fg_color=AZUL_VIVO, hover_color="#2563EB",
            command=self._abrir_dialogo, **btn_cfg)
        self.btn_buscar.pack(fill="x", padx=14, pady=(0, 8))

        self.btn_validar = ctk.CTkButton(
            sb, text="▶  Validar",
            fg_color=VERDE_DIM, hover_color="#15803D",
            command=self._iniciar_validacao,
            state="disabled", **btn_cfg)
        self.btn_validar.pack(fill="x", padx=14, pady=(0, 8))

        self.btn_exportar = ctk.CTkButton(
            sb, text="💾  Exportar relatório",
            fg_color=ROXO, hover_color="#6D28D9",
            command=self._exportar,
            state="disabled", **btn_cfg)
        self.btn_exportar.pack(fill="x", padx=14, pady=(0, 8))

        self.btn_limpar = ctk.CTkButton(
            sb, text="✕  Limpar",
            fg_color="#374151", hover_color="#4B5563",
            command=self._limpar, **btn_cfg)
        self.btn_limpar.pack(fill="x", padx=14)

        # Status
        self.lbl_status = ctk.CTkLabel(
            sb, text="Aguardando arquivo...",
            font=ctk.CTkFont(size=10), text_color=DIM,
            wraplength=200, justify="left")
        self.lbl_status.pack(side="bottom", padx=16, pady=16, anchor="w")

    # ── CONTEÚDO ─────────────────────────────────────────────
    def _build_content(self):
        ct = self.content

        # Cabeçalho
        hdr = ctk.CTkFrame(ct, fg_color="transparent")
        hdr.pack(fill="x", padx=24, pady=(20, 12))
        ctk.CTkLabel(
            hdr,
            text="Validação de Coordenadas Geodésicas",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=BRANCO
        ).pack(anchor="w")
        ctk.CTkLabel(
            hdr,
            text='Detecta: "SICOR: A gleba informada não corresponde a uma área válida"',
            font=ctk.CTkFont(size=11), text_color=DIM
        ).pack(anchor="w", pady=(2, 0))

        # Zona de drop
        self.zona_drop = ZonaDrop(ct, on_click=self._abrir_dialogo)
        self.zona_drop.pack(fill="x", padx=24, pady=(0, 12))

        # Progress bar
        self.progress = ctk.CTkProgressBar(
            ct, height=4, corner_radius=2,
            fg_color=CINZA_BORDA, progress_color=AZUL_VIVO)
        self.progress.pack(fill="x", padx=24, pady=(0, 12))
        self.progress.set(0)

        # Tabview
        self.tabview = ctk.CTkTabview(
            ct,
            fg_color=CINZA_CARD,
            segmented_button_fg_color="#111827",
            segmented_button_selected_color=AZUL_VIVO,
            segmented_button_selected_hover_color="#2563EB",
            segmented_button_unselected_color="#111827",
            segmented_button_unselected_hover_color="#1F2937",
            text_color=BRANCO,
            corner_radius=12
        )
        self.tabview.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        self.tabview.add("📋  Relatório")
        self.tabview.add("🔍  Por Gleba")
        self.tabview.add("ℹ️  Como usar")

        self._build_aba_relatorio()
        self._build_aba_glebas()
        self._build_aba_ajuda()

    def _build_aba_relatorio(self):
        aba = self.tabview.tab("📋  Relatório")
        self.txt_relatorio = ctk.CTkTextbox(
            aba, fg_color="transparent", text_color=BRANCO,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word", state="disabled",
            scrollbar_button_color=CINZA_BORDA,
            scrollbar_button_hover_color=AZUL_VIVO,
        )
        self.txt_relatorio.pack(fill="both", expand=True, padx=4, pady=4)
        self._escrever_relatorio(self._texto_boas_vindas())

    def _build_aba_glebas(self):
        aba = self.tabview.tab("🔍  Por Gleba")
        self.frame_glebas_scroll = ctk.CTkScrollableFrame(
            aba, fg_color="transparent",
            scrollbar_button_color=CINZA_BORDA,
            scrollbar_button_hover_color=AZUL_VIVO,
        )
        self.frame_glebas_scroll.pack(fill="both", expand=True)
        ctk.CTkLabel(
            self.frame_glebas_scroll,
            text="Nenhum arquivo validado ainda.",
            font=ctk.CTkFont(size=13), text_color=DIM
        ).pack(pady=40)

    def _build_aba_ajuda(self):
        aba = self.tabview.tab("ℹ️  Como usar")
        txt = ctk.CTkTextbox(
            aba, fg_color="transparent", text_color=BRANCO,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            wrap="word", scrollbar_button_color=CINZA_BORDA,
        )
        txt.pack(fill="both", expand=True, padx=4, pady=4)
        txt.insert("end", """
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  COMO USAR O VALIDADOR
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1.  Clique na área central OU em "🗂 Buscar arquivo"
      para selecionar o arquivo Excel.

  2.  Clique em "▶ Validar" para iniciar a análise.

  3.  Veja os resultados na aba "📋 Relatório".

  4.  Para ver resultado por gleba: aba "🔍 Por Gleba".

  5.  Para salvar: clique em "💾 Exportar relatório".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  FORMATO DO ARQUIVO EXCEL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  O arquivo deve ter 4 colunas:

  ┌────────┬────────┬──────────────────┬──────────────────┐
  │ Gleba  │ Ponto  │    Latitude      │   Longitude      │
  ├────────┼────────┼──────────────────┼──────────────────┤
  │   1    │   1    │ -14.43539142600  │ -44.33006286500  │
  │   1    │   2    │ -14.43388207100  │ -44.33622437000  │
  │   1    │   3    │ -14.43252498700  │ -44.33593356600  │
  │   1    │   4    │ -14.43539142600  │ -44.33006286500  │ ← igual ao 1º
  └────────┴────────┴──────────────────┴──────────────────┘

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ERROS DETECTADOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  🔴 POLÍGONO NÃO FECHADO
     Primeiro e último ponto devem ter as mesmas
     coordenadas. Copie a 1ª linha e cole no final.

  🟠 PONTOS INSUFICIENTES
     Mínimo de 3 vértices únicos (4 linhas com
     o ponto de fechamento).

  🟡 PONTO DUPLICADO EM EXCESSO
     Uma coordenada aparece 3+ vezes na sequência.
     Somente o ponto de fechamento pode repetir.

  🔴 COORDENADA INVÁLIDA
     Latitude ou longitude não é um número válido.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  REFERÊNCIA NORMATIVA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  • MCR 2-1-2  —  Manual de Crédito Rural
  • Normativo 3102-32-05  —  SICOR / BACEN
  • Treinamento GLEBAS 2026  —  FSBR
""")
        txt.configure(state="disabled")

    # ── AÇÕES ────────────────────────────────────────────────
    def _abrir_dialogo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo de coordenadas",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        if caminho:
            self._arquivo = caminho
            self.zona_drop.set_arquivo(os.path.basename(caminho))
            self.btn_validar.configure(state="normal")
            self.lbl_status.configure(text=f"📄  {os.path.basename(caminho)}")

    def _iniciar_validacao(self):
        if not self._arquivo:
            return
        self.btn_validar.configure(state="disabled")
        self.btn_exportar.configure(state="disabled")
        self.lbl_status.configure(text="Processando...")
        self.progress.configure(mode="indeterminate", progress_color=AZUL_VIVO)
        self.progress.start()
        threading.Thread(target=self._rodar_validacao, daemon=True).start()

    def _rodar_validacao(self):
        try:
            df = carregar_planilha(self._arquivo)
            cols = detectar_colunas(df)
            erros, grupos = validar(df, cols)
            self.after(0, lambda: self._exibir_resultado(erros, grupos))
        except Exception as ex:
            self.after(0, lambda: self._exibir_erro(str(ex)))

    def _exibir_erro(self, msg):
        self._parar_progress()
        self.btn_validar.configure(state="normal")
        self.lbl_status.configure(text="❌  Erro ao processar")
        self._escrever_relatorio(
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "  ❌  ERRO AO ABRIR O ARQUIVO\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            f"  {msg}\n\n"
            "  Verifique:\n"
            "  • Arquivo aberto no Excel? — feche-o\n"
            "  • pip install xlrd      (para .xls)\n"
            "  • pip install openpyxl  (para .xlsx)\n"
        )

    def _exibir_resultado(self, erros, grupos):
        self._parar_progress()
        self._erros_cache  = erros
        self._grupos_cache = grupos

        total      = len(grupos)
        n_err      = len(erros)
        glebas_err = len({e["gleba"] for e in erros})
        glebas_ok  = total - glebas_err

        self.card_glebas.atualizar(total, AZUL_CLARO)
        self.card_erros.atualizar(n_err, VERMELHO if n_err else VERDE)
        self.card_ok.atualizar(glebas_ok, VERDE)

        if not erros:
            self.lbl_status.configure(text=f"✅  {total} gleba(s) sem erros!")
            self.progress.configure(progress_color=VERDE)
        else:
            self.lbl_status.configure(text=f"⚠  {n_err} erro(s) em {glebas_err} gleba(s)")
            self.progress.configure(progress_color=VERMELHO)

        self.progress.set(1)
        self.btn_validar.configure(state="normal")
        if erros:
            self.btn_exportar.configure(state="normal")

        self._escrever_relatorio(self._montar_texto(erros, grupos))
        self._atualizar_aba_glebas(erros, grupos)

    def _parar_progress(self):
        self.progress.stop()
        self.progress.configure(mode="determinate")

    def _montar_texto(self, erros, grupos):
        horario = datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
        nome    = os.path.basename(self._arquivo)
        linhas  = [
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n",
            "  RELATÓRIO DE VALIDAÇÃO — SICOR\n",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n",
            f"  Arquivo  :  {nome}\n",
            f"  Horário  :  {horario}\n",
            f"  Glebas   :  {len(grupos)}\n",
            f"  Erros    :  {len(erros)}\n\n",
            "─" * 50 + "\n",
        ]

        if not erros:
            linhas += [
                "\n  ✅  TODAS AS GLEBAS SÃO VÁLIDAS!\n",
                "      Nenhuma inconsistência detectada.\n\n",
            ]
        else:
            ICONES = {
                "POLÍGONO NÃO FECHADO"      : "🔴",
                "PONTOS INSUFICIENTES"      : "🟠",
                "PONTO DUPLICADO EM EXCESSO": "🟡",
                "COORDENADA INVÁLIDA"       : "🔴",
            }
            por_tipo = defaultdict(list)
            for e in erros:
                por_tipo[e["tipo"]].append(e)

            for tipo, lista in sorted(por_tipo.items()):
                linhas += [
                    f"\n  {ICONES.get(tipo,'⚠')}  {tipo}  ({len(lista)}×)\n",
                    "─" * 50 + "\n",
                ]
                for e in lista:
                    seq = str(e.get("seq", "-"))
                    seq_txt = f"  |  Ponto #{seq}" if seq not in ("-","nan","None","") else ""
                    linhas += [
                        f"  Gleba {e['gleba']}  |  Linha Excel: {e['linha']}{seq_txt}\n",
                        f"  → {e['msg']}\n\n",
                    ]

            linhas += [
                "─" * 50 + "\n\n",
                "  COMO CORRIGIR:\n\n",
                "  • POLÍGONO NÃO FECHADO     → Copie a 1ª linha e cole como última.\n",
                "  • PONTOS INSUFICIENTES     → Adicione mais vértices (mín. 3 únicos).\n",
                "  • PONTO DUPLICADO          → Remova coordenadas repetidas no meio.\n",
                "  • COORDENADA INVÁLIDA      → Corrija o valor numérico na planilha.\n\n",
            ]

        return "".join(linhas)

    def _atualizar_aba_glebas(self, erros, grupos):
        for w in self.frame_glebas_scroll.winfo_children():
            w.destroy()

        glebas_com_erro = defaultdict(list)
        for e in erros:
            glebas_com_erro[e["gleba"]].append(e)

        for num, pontos in sorted(grupos.items(), key=lambda x: x[0]):
            errs    = glebas_com_erro.get(num, [])
            tem_err = bool(errs)

            card = ctk.CTkFrame(
                self.frame_glebas_scroll,
                fg_color="#2a1a1a" if tem_err else "#1a2a1a",
                corner_radius=10, border_width=1,
                border_color=VERMELHO if tem_err else VERDE
            )
            card.pack(fill="x", padx=6, pady=5)

            topo = ctk.CTkFrame(card, fg_color="transparent")
            topo.pack(fill="x", padx=14, pady=(12, 4))

            ctk.CTkLabel(
                topo,
                text=f"{'❌' if tem_err else '✅'}  Gleba {num}",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color=VERMELHO if tem_err else VERDE
            ).pack(side="left")

            ctk.CTkLabel(
                topo, text=f"{len(pontos)} linha(s)",
                font=ctk.CTkFont(size=11), text_color=DIM
            ).pack(side="right")

            if errs:
                for e in errs:
                    lf = ctk.CTkFrame(card, fg_color="transparent")
                    lf.pack(fill="x", padx=14, pady=2)
                    BadgeTipo(lf, e["tipo"]).pack(side="left", padx=(0, 8))
                    ctk.CTkLabel(
                        lf,
                        text=f"Linha {e['linha']} — {e['msg']}",
                        font=ctk.CTkFont(size=11),
                        text_color="#FCA5A5",
                        wraplength=520, justify="left"
                    ).pack(side="left", anchor="w")
                ctk.CTkFrame(card, fg_color="transparent", height=6).pack()
            else:
                ctk.CTkLabel(
                    card, text="Gleba válida.",
                    font=ctk.CTkFont(size=11), text_color=VERDE
                ).pack(anchor="w", padx=14, pady=(0, 12))

    def _escrever_relatorio(self, texto):
        self.txt_relatorio.configure(state="normal")
        self.txt_relatorio.delete("0.0", "end")
        self.txt_relatorio.insert("end", texto)
        self.txt_relatorio.configure(state="disabled")
        self.txt_relatorio.see("0.0")

    def _texto_boas_vindas(self):
        return (
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "  BEM-VINDO AO VALIDADOR DE GLEBAS — SICOR\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "  Detecta o erro:\n\n"
            "  \"SICOR: A gleba informada não corresponde\n"
            "   a uma área válida\"\n\n"
            "─" * 50 + "\n\n"
            "  COMO COMEÇAR:\n\n"
            "  1.  Clique na área central  OU  em\n"
            "      🗂 Buscar arquivo na barra lateral.\n\n"
            "  2.  Clique em  ▶ Validar.\n\n"
            "  3.  O resultado aparecerá aqui.\n\n"
            "─" * 50 + "\n\n"
            "  VERIFICAÇÕES REALIZADAS:\n\n"
            "  🔴  Polígono não fechado\n"
            "  🟠  Pontos insuficientes (mín. 3 únicos)\n"
            "  🟡  Ponto duplicado em excesso\n"
            "  🔴  Coordenada inválida\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        )

    def _limpar(self):
        self._arquivo      = None
        self._erros_cache  = []
        self._grupos_cache = {}
        self.zona_drop.reset()
        self.btn_validar.configure(state="disabled")
        self.btn_exportar.configure(state="disabled")
        self.card_glebas.atualizar("—", AZUL_CLARO)
        self.card_erros.atualizar("—", DIM)
        self.card_ok.atualizar("—", VERDE)
        self.progress.set(0)
        self.progress.configure(progress_color=AZUL_VIVO)
        self.lbl_status.configure(text="Aguardando arquivo...")
        self._escrever_relatorio(self._texto_boas_vindas())
        for w in self.frame_glebas_scroll.winfo_children():
            w.destroy()
        ctk.CTkLabel(
            self.frame_glebas_scroll,
            text="Nenhum arquivo validado ainda.",
            font=ctk.CTkFont(size=13), text_color=DIM
        ).pack(pady=40)

    def _exportar(self):
        if not self._erros_cache:
            return
        caminho = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            initialfile="relatorio_glebas.txt",
            title="Salvar relatório"
        )
        if not caminho:
            return
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(self._montar_texto(self._erros_cache, self._grupos_cache))
        self.lbl_status.configure(text=f"💾  Salvo: {os.path.basename(caminho)}")


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
