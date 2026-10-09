#!/usr/bin/env python3
"""
Histórico Escolar — busca de aluno, conferência e edição de notas, campos de
estágio e período do estágio, e impressão do diploma / histórico (modelo
DIPLOMA - PADRÃO). A edição fica no módulo editor_dados.py.
"""
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

# Esconde a janela de console quando empacotado com --console
if sys.platform == "win32":
    try:
        import ctypes
        _hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if _hwnd:
            ctypes.windll.user32.ShowWindow(_hwnd, 0)
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # texto nítido em telas HiDPI
        except Exception:
            pass
    except Exception:
        pass

import diploma_core as dc
import editor_dados
import diploma_docx

# ------------------------------------------------------------ aparência ----
AZUL = "#1F4E79"
AZUL_CLARO = "#2E75B6"
FUNDO = "#F3F6FA"
CARD = "#FFFFFF"
TEXTO = "#1E2A36"
CINZA = "#6B7785"
VERMELHO = "#B3261E"
VERDE = "#1E7B45"
LINHA_ALT = "#EEF3F9"
FONTE = "Segoe UI" if sys.platform == "win32" else "Helvetica"

TODOS_CURSOS = "Todos os cursos"
TODAS_TURMAS = "Todas as turmas"
LIMITE_RESULTADOS = 500

# rótulo da disciplina de estágio -> coluna do CSV (para abrir o editor no campo certo)
ROTULO_PARA_CAMPO = {rot: col for rot, col in dc.ESTAGIO_DIPLOMA + dc.ESTAGIO_EXTRAS}


def abrir_arquivo(caminho, imprimir=False):
    """abre (ou manda imprimir) um arquivo com o programa padrão do sistema"""
    if sys.platform == "win32":
        if imprimir:
            os.startfile(str(caminho), "print")
        else:
            os.startfile(str(caminho))
    elif sys.platform == "darwin":
        subprocess.run(["lp" if imprimir else "open", str(caminho)], check=False)
    else:
        subprocess.run(["lp" if imprimir else "xdg-open", str(caminho)], check=False)


class App(tk.Tk):
    def __init__(self, base, duplicadas, csv_path):
        super().__init__()
        self.base = base
        self.csv_path = Path(csv_path)
        csv_nome = self.csv_path.name
        self.resultados = []
        self.sel = None
        self._timer = None

        self.title("Histórico Escolar — Busca de Aluno")
        w = min(1200, self.winfo_screenwidth() - 40)
        h = min(780, self.winfo_screenheight() - 90)
        self.geometry(f"{w}x{h}")
        self.minsize(900, 560)
        self.configure(bg=FUNDO)

        self.var_busca = tk.StringVar()
        self.var_curso = tk.StringVar(value=TODOS_CURSOS)
        self.var_turma = tk.StringVar(value=TODAS_TURMAS)
        self.var_ano = tk.StringVar()
        self.var_ini = tk.StringVar()
        self.var_fim = tk.StringVar()
        self.var_status = tk.StringVar()

        self._estilos()
        self._montar_ui()
        self._carregar_filtros()
        self.var_status.set(
            f"Base carregada: {len(base)} alunos  ({csv_nome})"
            + (f" — {duplicadas} linhas repetidas foram ignoradas" if duplicadas else "")
        )
        self.var_busca.trace_add("write", lambda *_: self._agendar_busca())
        self.ent_busca.focus_set()

    # --------------------------------------------------------- estilos ----
    def _estilos(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure(".", font=(FONTE, 10), background=FUNDO, foreground=TEXTO)
        st.configure("TFrame", background=FUNDO)
        st.configure("Card.TFrame", background=CARD)
        st.configure("TLabel", background=FUNDO, foreground=TEXTO)
        st.configure("Card.TLabel", background=CARD)
        st.configure("Muted.TLabel", background=CARD, foreground=CINZA, font=(FONTE, 9))
        st.configure("Titulo.TLabel", background=CARD, font=(FONTE, 17, "bold"), foreground=AZUL)
        st.configure("TEntry", padding=6, fieldbackground="white")
        st.configure("TCombobox", padding=5)
        st.configure("TButton", padding=(12, 7), font=(FONTE, 10))
        st.configure("Primary.TButton", background=AZUL, foreground="white",
                     font=(FONTE, 10, "bold"), padding=(14, 9))
        st.map("Primary.TButton", background=[("active", AZUL_CLARO), ("disabled", "#A9B7C6")],
               foreground=[("disabled", "#EEF2F6")])
        st.configure("Treeview", rowheight=25, font=(FONTE, 10), background="white",
                     fieldbackground="white", borderwidth=0)
        st.configure("Treeview.Heading", font=(FONTE, 10, "bold"), background="#DCE6F1",
                     foreground=AZUL, padding=6)
        st.map("Treeview", background=[("selected", AZUL_CLARO)], foreground=[("selected", "white")])
        st.configure("TNotebook", background=FUNDO, borderwidth=0)
        st.configure("TNotebook.Tab", padding=(16, 8), font=(FONTE, 10))
        st.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", AZUL)])
        st.configure("Horizontal.TProgressbar", troughcolor="#DCE3EB", background=VERDE, thickness=8)

    # -------------------------------------------------------------- UI ----
    def _montar_ui(self):
        # --- cabeçalho
        topo = tk.Frame(self, bg=AZUL)
        topo.pack(fill="x")
        tk.Label(topo, text="Histórico Escolar", bg=AZUL, fg="white",
                 font=(FONTE, 17, "bold")).pack(side="left", padx=(20, 8), pady=8)
        tk.Label(topo, text="Centro de Formação Técnica em Enfermagem Irmã Dulce", bg=AZUL,
                 fg="#C9DAEC", font=(FONTE, 10)).pack(side="left", pady=(8, 0))

        # --- busca e filtros
        busca = ttk.Frame(self, padding=(20, 10, 20, 4))
        busca.pack(fill="x")
        ttk.Label(busca, text="Buscar aluno (nome ou nº da matrícula)",
                  font=(FONTE, 9), foreground=CINZA).grid(row=0, column=0, sticky="w")
        self.ent_busca = ttk.Entry(busca, textvariable=self.var_busca, font=(FONTE, 12), width=44)
        self.ent_busca.grid(row=1, column=0, sticky="we", padx=(0, 12))
        ttk.Label(busca, text="Curso", font=(FONTE, 9), foreground=CINZA).grid(row=0, column=1, sticky="w")
        self.cb_curso = ttk.Combobox(busca, textvariable=self.var_curso, state="readonly", width=28)
        self.cb_curso.grid(row=1, column=1, padx=(0, 12))
        ttk.Label(busca, text="Turma", font=(FONTE, 9), foreground=CINZA).grid(row=0, column=2, sticky="w")
        self.cb_turma = ttk.Combobox(busca, textvariable=self.var_turma, state="readonly", width=14)
        self.cb_turma.grid(row=1, column=2, padx=(0, 12))
        ttk.Button(busca, text="Limpar", command=self.limpar).grid(row=1, column=3)
        busca.columnconfigure(0, weight=1)
        for cb in (self.cb_curso, self.cb_turma):
            cb.bind("<<ComboboxSelected>>", lambda e: self.buscar())

        # --- rodapé (empacotado primeiro para nunca ser cortado)
        ttk.Label(self, textvariable=self.var_status, foreground=CINZA,
                  padding=(20, 0, 20, 8)).pack(side="bottom", fill="x")

        # --- ações
        acoes = ttk.Frame(self, padding=(20, 12, 20, 6))
        acoes.pack(side="bottom", fill="x")
        self.btn_diploma = ttk.Button(acoes, text="Imprimir diploma completo (frente e verso)",
                                      style="Primary.TButton", command=lambda: self.imprimir(True))
        self.btn_diploma.pack(side="left")
        self.btn_hist = ttk.Button(acoes, text="Imprimir só o histórico (verso)",
                                   command=lambda: self.imprimir(False))
        self.btn_hist.pack(side="left", padx=8)
        self.btn_ver = ttk.Button(acoes, text="Visualizar PDF", command=self.visualizar)
        self.btn_ver.pack(side="left")
        self.btn_salvar = ttk.Button(acoes, text="Salvar PDF…", command=self.salvar)
        self.btn_salvar.pack(side="left", padx=8)
        self.btn_docx = ttk.Button(acoes, text="Salvar DOCX…", command=self.salvar_docx)
        self.btn_docx.pack(side="left")
        self.btn_editar = ttk.Button(acoes, text="Editar notas e estágios", command=self.editar)
        self.btn_editar.pack(side="left", padx=8)
        self._botoes = (self.btn_diploma, self.btn_hist, self.btn_ver, self.btn_salvar,
                        self.btn_docx, self.btn_editar)
        self._estado_botoes(False)

        

        # --- dados para impressão (uma linha só)
        imp = ttk.Frame(self, style="Card.TFrame", padding=(16, 8))
        imp.pack(side="bottom", fill="x", padx=20, pady=(8, 0))
        ttk.Label(imp, text="Dados para impressão", style="Card.TLabel",
                  font=(FONTE, 10, "bold"), foreground=AZUL).pack(side="left", padx=(0, 18))
        ttk.Label(imp, text="Ano letivo de conclusão", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        ttk.Entry(imp, textvariable=self.var_ano, width=7).pack(side="left", padx=(0, 18))
        ttk.Label(imp, text="Período do estágio: de", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        ttk.Entry(imp, textvariable=self.var_ini, width=11).pack(side="left")
        ttk.Label(imp, text="a", style="Muted.TLabel").pack(side="left", padx=6)
        ttk.Entry(imp, textvariable=self.var_fim, width=11).pack(side="left", padx=(0, 14))
        ttk.Label(imp, text="(pode corrigir antes de imprimir)", style="Muted.TLabel").pack(side="left")

        # --- corpo
        corpo = ttk.PanedWindow(self, orient="horizontal")
        corpo.pack(side="top", fill="both", expand=True, padx=20, pady=(6, 0))

        esq = ttk.Frame(corpo, style="Card.TFrame", padding=10)
        corpo.add(esq, weight=2)
        self.lbl_resultados = ttk.Label(esq, text="Resultados", style="Muted.TLabel")
        self.lbl_resultados.pack(anchor="w", pady=(0, 6))
        cols = ("nome", "mat", "turma", "ano")
        self.lista = ttk.Treeview(esq, columns=cols, show="headings", selectmode="browse")
        for c, t, w, a in (("nome", "Aluno", 210, "w"), ("mat", "Matrícula", 110, "center"),
                           ("turma", "Turma", 62, "center"), ("ano", "Concl.", 62, "center")):
            self.lista.heading(c, text=t)
            self.lista.column(c, width=w, anchor=a, stretch=(c == "nome"))
        sb = ttk.Scrollbar(esq, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=sb.set)
        self.lista.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.lista.bind("<<TreeviewSelect>>", self.selecionar)
        self.lista.tag_configure("par", background=LINHA_ALT)

        dir_ = ttk.Frame(corpo, padding=(14, 0, 0, 0))
        corpo.add(dir_, weight=5)

        # cartão do aluno
        card = ttk.Frame(dir_, style="Card.TFrame", padding=(16, 10))
        card.pack(fill="x")
        self.lbl_nome = ttk.Label(card, text="Selecione um aluno na lista", style="Titulo.TLabel")
        self.lbl_nome.pack(anchor="w")
        self.lbl_sub = ttk.Label(card, text="Use a busca acima para encontrar o aluno.", style="Muted.TLabel")
        self.lbl_sub.pack(anchor="w", pady=(2, 6))
        lin = ttk.Frame(card, style="Card.TFrame")
        lin.pack(fill="x")
        self.lbl_prog = ttk.Label(lin, text="", style="Card.TLabel", font=(FONTE, 9, "bold"))
        self.lbl_prog.pack(side="left")
        self.barra = ttk.Progressbar(lin, maximum=100, length=220)
        self.barra.pack(side="left", padx=10)
        self.lbl_aviso = ttk.Label(card, text="", style="Card.TLabel", foreground=VERMELHO,
                                   font=(FONTE, 9), wraplength=620, justify="left")
        self.lbl_aviso.pack(anchor="w", pady=(4, 0))

        # abas
        self.abas = ttk.Notebook(dir_)
        self.abas.pack(fill="both", expand=True, pady=(10, 0))
        self._aba_notas()
        self._aba_estagios()
        self._aba_dados()
        self._aba_bruta()


    def _arvore(self, pai, colunas, rodape=False):
        """cria uma Treeview com barra de rolagem; colunas = [(id, título, largura, ancora)]"""
        quadro = ttk.Frame(pai, style="Card.TFrame", padding=6)
        if rodape:  # linha fixa abaixo da lista (empacotada antes para não ser cortada)
            self.lbl_periodo = ttk.Label(quadro, text="", style="Card.TLabel", font=(FONTE, 10, "bold"))
            self.lbl_periodo.pack(side="bottom", anchor="w", pady=(6, 0))
        tv = ttk.Treeview(quadro, columns=[c[0] for c in colunas], show="headings", selectmode="browse")
        for cid, tit, larg, anc in colunas:
            tv.heading(cid, text=tit)
            tv.column(cid, width=larg, anchor=anc, stretch=(anc == "w"))
        sb = ttk.Scrollbar(quadro, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return quadro, tv

    def _aba_notas(self):
        quadro, self.tv_notas = self._arvore(self.abas, [
            ("mod", "Módulo", 70, "center"), ("comp", "Componente curricular", 330, "w"),
            ("teoria", "Teoria", 80, "center"), ("est", "Estágio", 80, "center")])
        self.abas.add(quadro, text="  Notas finais  ")
        self.tv_notas.tag_configure("falta", foreground=VERMELHO)
        self.tv_notas.tag_configure("m1", background="#FFFFFF")
        self.tv_notas.tag_configure("m2", background=LINHA_ALT)
        self.tv_notas.bind("<Double-1>", self._duplo_clique_notas)

    def _aba_estagios(self):
        quadro, self.tv_est = self._arvore(self.abas, [
            ("disc", "Disciplina", 190, "w"), ("local", "Campo de estágio", 420, "w")], rodape=True)
        self.abas.add(quadro, text="  Estágios  ")
        self.tv_est.tag_configure("falta", foreground=CINZA)
        self.tv_est.bind("<Double-1>", self._duplo_clique_estagios)

    def _aba_dados(self):
        quadro = ttk.Frame(self.abas, style="Card.TFrame", padding=16)
        self.abas.add(quadro, text="  Dados do aluno  ")
        self.lbls_dados = {}
        campos = [("nome", "Nome"), ("matricula", "Matrícula"), ("rg", "RG"), ("nascimento", "Nascimento"),
                  ("naturalidade", "Naturalidade"), ("estado", "Estado"), ("pai", "Pai"), ("mae", "Mãe"),
                  ("curso", "Curso"), ("turma", "Turma"), ("turno", "Turno"), ("inicio", "Início do curso"),
                  ("conclusao", "Data de conclusão"), ("endereco", "Endereço")]
        for i, (k, rot) in enumerate(campos):
            ttk.Label(quadro, text=rot, style="Muted.TLabel").grid(row=i, column=0, sticky="nw", pady=3, padx=(0, 16))
            v = ttk.Label(quadro, text="", style="Card.TLabel", wraplength=520, justify="left")
            v.grid(row=i, column=1, sticky="w", pady=3)
            self.lbls_dados[k] = v

    def _aba_bruta(self):
        quadro, self.tv_bruta = self._arvore(self.abas, [
            ("campo", "Campo original da base", 240, "w"), ("valor", "Valor", 200, "w")])
        self.abas.add(quadro, text="  Ficha original  ")

    # ---------------------------------------------------------- busca ----
    def _carregar_filtros(self):
        cursos = {}
        for r in self.base:
            cursos.setdefault(dc.normalizar(r["_curso"]), r["_curso"])
        self.cb_curso["values"] = [TODOS_CURSOS] + sorted(cursos.values(), key=dc.normalizar)
        turmas = sorted({r["_turma"] for r in self.base})
        self.cb_turma["values"] = [TODAS_TURMAS] + turmas

    def _agendar_busca(self):
        if self._timer:
            self.after_cancel(self._timer)
        self._timer = self.after(250, self.buscar)

    def limpar(self):
        self.var_busca.set("")
        self.var_curso.set(TODOS_CURSOS)
        self.var_turma.set(TODAS_TURMAS)
        self.buscar()
        self.ent_busca.focus_set()

    def buscar(self):
        self._timer = None
        tokens = dc.normalizar(self.var_busca.get()).split()
        curso, turma = self.var_curso.get(), self.var_turma.get()
        sem_filtro = curso == TODOS_CURSOS and turma == TODAS_TURMAS
        self.lista.delete(*self.lista.get_children())
        self.resultados = []
        self._mostrar(None)
        if not tokens and sem_filtro:
            self.lbl_resultados.config(text="Digite um nome ou matrícula para buscar")
            return
        achados = []
        for r in self.base:
            if curso != TODOS_CURSOS and dc.normalizar(r["_curso"]) != dc.normalizar(curso):
                continue
            if turma != TODAS_TURMAS and r["_turma"] != turma:
                continue
            if tokens and not all(t in r["_busca"] for t in tokens):
                continue
            achados.append(r)
        achados.sort(key=lambda r: dc.normalizar(r.get("NOME", "")))
        total = len(achados)
        self.resultados = achados[:LIMITE_RESULTADOS]
        for i, r in enumerate(self.resultados):
            self.lista.insert("", "end", iid=str(i), tags=("par",) if i % 2 else (),
                              values=(r.get("NOME", "").strip(), r.get("NUM_MAT", "").strip(),
                                      r["_turma"], dc.ano_conclusao(r)))
        if not total:
            self.lbl_resultados.config(text="Nenhum aluno encontrado")
        elif total > LIMITE_RESULTADOS:
            self.lbl_resultados.config(text=f"{total} alunos — mostrando os {LIMITE_RESULTADOS} primeiros (refine a busca)")
        else:
            self.lbl_resultados.config(text=f"{total} aluno(s) encontrado(s)")
        if self.resultados and total <= 40:
            self.lista.selection_set("0")

    def selecionar(self, _e=None):
        sel = self.lista.selection()
        if sel:
            self._mostrar(self.resultados[int(sel[0])])

    # --------------------------------------------------------- detalhe ----
    def _estado_botoes(self, ativo, diploma_ok=True):
        for b in self._botoes:
            b.state(["!disabled"] if ativo else ["disabled"])
        if ativo and not diploma_ok:
            self.btn_diploma.state(["disabled"])
            self.btn_hist.state(["disabled"])
            self.btn_docx.state(["disabled"])

    def _mostrar(self, r):
        self.sel = r
        self.tv_notas.delete(*self.tv_notas.get_children())
        self.tv_est.delete(*self.tv_est.get_children())
        self.tv_bruta.delete(*self.tv_bruta.get_children())
        for v in self.lbls_dados.values():
            v.config(text="")
        self.lbl_periodo.config(text="")
        self.lbl_aviso.config(text="")
        self.lbl_prog.config(text="")
        self.barra["value"] = 0
        if r is None:
            self.lbl_nome.config(text="Selecione um aluno na lista")
            self.lbl_sub.config(text="Use a busca acima para encontrar o aluno.")
            for v in (self.var_ano, self.var_ini, self.var_fim):
                v.set("")
            self._estado_botoes(False)
            return

        d = dc.dados_pessoais(r)
        self.lbl_nome.config(text=d["nome"])
        sub = [x for x in (d["curso"], f"Turma {d['turma']}" if d["turma"] else "",
                           f"Matrícula {d['matricula']}" if d["matricula"] else "", d["turno"]) if x]
        self.lbl_sub.config(text="   •   ".join(sub))
        for k, v in self.lbls_dados.items():
            v.config(text=d.get(k) or "—")

        tecnico = dc.is_tecnico(r)
        # notas
        itens = dc.historico_do_aluno(r)
        achadas, esperadas, faltas = dc.resumo_faltas(itens)
        for i, it in enumerate(itens):
            t = it["nota_t"] or "— falta —"
            e = (it["nota_e"] or "— falta —") if it["ch_e"] else ""
            tags = ["m1" if it["modulo"] % 2 else "m2"]
            if it["falta_t"]:
                tags.append("falta")
            self.tv_notas.insert("", "end", values=(f"Mód. {it['modulo']:02d}", it["nome"], t, e), tags=tags)
        # estágios
        for rot, local in dc.estagios_do_aluno(r):
            self.tv_est.insert("", "end", values=(rot, local or "— não informado —"),
                               tags=("" if local else "falta",))
        ini, fim = dc.periodo_estagio(r)
        self.lbl_periodo.config(text="Período do estágio:  " + (f"{ini or '—'}  a  {fim or '—'}"))
        # ficha original
        for campo, valor in r.items():
            if campo.startswith("_") or campo in dc.CAMPOS_NAO_BRUTOS:
                continue
            valor = (valor or "").strip()
            if valor:
                self.tv_bruta.insert("", "end", values=(campo, valor))

        # indicador de completude
        if tecnico:
            self.lbl_prog.config(text=f"{achadas} de {esperadas} notas encontradas na base")
            self.barra["value"] = 100 * achadas / esperadas
            if faltas:
                self.lbl_aviso.config(text="Faltam: " + "; ".join(faltas[:4]) +
                                      (f" … e mais {len(faltas) - 4}" if len(faltas) > 4 else ""))
            else:
                self.lbl_aviso.config(text="Todas as notas do diploma foram encontradas.", foreground=VERDE)
        else:
            self.lbl_prog.config(text="Este aluno não é do curso Técnico em Enfermagem")
            self.lbl_aviso.config(text="O modelo de diploma padrão só se aplica ao Técnico em Enfermagem.")
        if faltas or not tecnico:
            self.lbl_aviso.config(foreground=VERMELHO)

        # campos de impressão
        self.var_ano.set(dc.ano_conclusao(r))
        self.var_ini.set(ini)
        self.var_fim.set(fim)
        self._estado_botoes(True, tecnico)
        self.var_status.set(f"{d['nome']}: {achadas} de {esperadas} notas encontradas.")

    # --------------------------------------------------------- edição ----
    def _duplo_clique_notas(self, event):
        """duplo clique numa nota abre o editor já editando aquela célula"""
        tv = self.tv_notas
        if not self.sel or tv.identify_region(event.x, event.y) != "cell":
            return
        iid = tv.identify_row(event.y)
        # colunas #3 (teoria) e #4 (estágio) -> posição em dc.CURRICULO (col_t=5, col_e=6)
        pos = {"#3": 5, "#4": 6}.get(tv.identify_column(event.x))
        campo = dc.CURRICULO[tv.index(iid)][pos] if iid and pos else None
        self.editar(aba=0, foco=("nota", campo) if campo else None)

    def _duplo_clique_estagios(self, event):
        tv = self.tv_est
        if not self.sel or tv.identify_region(event.x, event.y) != "cell":
            return
        iid = tv.identify_row(event.y)
        campo = ROTULO_PARA_CAMPO.get(tv.item(iid, "values")[0]) if iid else None
        self.editar(aba=1, foco=("local", campo) if campo else None)

    def _locais_conhecidos(self):
        """nomes de locais já usados na base, para o autocompletar do editor"""
        usados = {dc.limpa(r.get(c, "")) for r in self.base
                  for _rot, c in dc.ESTAGIO_DIPLOMA + dc.ESTAGIO_EXTRAS}
        usados.discard("")
        return sorted(usados, key=dc.normalizar)

    def editar(self, aba=0, foco=None):
        if not self.sel:
            messagebox.showinfo("Aviso", "Selecione um aluno primeiro.")
            return
        editor_dados.EditorAluno(
            self, linha=self.sel, csv_path=self.csv_path,
            locais_conhecidos=self._locais_conhecidos(), ao_salvar=self._apos_edicao,
            aba_inicial=aba, foco=foco)

    def _apos_edicao(self, qtd, backup, n_linhas):
        self._mostrar(self.sel)          # redesenha notas, estágios, período e indicadores
        msg = (f"{qtd} alteração(ões) salva(s) para {dc.limpa(self.sel.get('NOME', ''))}. "
               "Cópia de segurança na pasta 'backups'.")
        if n_linhas > 1:
            msg += f" ({n_linhas} linhas idênticas no arquivo foram atualizadas.)"
        self.var_status.set(msg)

    # ------------------------------------------------------- impressão ----
    def _nome_arquivo(self, sufixo, ext=".pdf"):
        nome = (self.sel.get("NOME", "aluno") or "aluno").strip()
        nome = "".join(c for c in nome if c.isalnum() or c in " _-").strip().replace(" ", "_")
        return f"{sufixo}_{nome or 'aluno'}{ext}"

    def _confirmar_faltas(self):
        _, _, faltas = dc.resumo_faltas(dc.historico_do_aluno(self.sel))
        if not faltas:
            return True
        resumo = "\n".join("• " + f for f in faltas[:10]) + (f"\n… e mais {len(faltas) - 10}" if len(faltas) > 10 else "")
        return messagebox.askyesno(
            "Notas faltando",
            f"Não encontrei {len(faltas)} nota(s) na base para este aluno:\n\n{resumo}\n\n"
            "Elas sairão em branco no histórico. Deseja continuar?")

    def _gerar(self, frente, destino=None, conferir=True):
        if not self.sel:
            messagebox.showinfo("Aviso", "Selecione um aluno primeiro.")
            return None
        if not dc.is_tecnico(self.sel):
            messagebox.showinfo("Aviso", "O modelo de diploma padrão é só para o curso Técnico em Enfermagem.")
            return None
        if conferir and not self._confirmar_faltas():
            return None
        caminho = Path(destino) if destino else Path(tempfile.gettempdir()) / self._nome_arquivo(
            "diploma" if frente else "historico")
        opc = {"ano": self.var_ano.get(), "per_ini": self.var_ini.get(), "per_fim": self.var_fim.get()}
        try:
            dc.gerar_pdf(self.sel, caminho, opc, incluir_frente=frente)
        except ImportError:
            messagebox.showerror("Falta um componente",
                                 "Para gerar o PDF é preciso instalar o ReportLab:\n\npip install reportlab")
            return None
        except Exception as e:
            messagebox.showerror("Erro ao gerar o PDF", str(e))
            return None
        return caminho

    def imprimir(self, frente):
        caminho = self._gerar(frente)
        if not caminho:
            return
        try:
            abrir_arquivo(caminho, imprimir=True)
            self.var_status.set(f"Enviado para impressão: {caminho.name}")
        except Exception:
            # se o leitor de PDF não aceitar impressão direta, abre o arquivo para imprimir de lá
            try:
                abrir_arquivo(caminho)
                self.var_status.set("PDF aberto — use Ctrl+P para imprimir.")
            except Exception as e:
                messagebox.showerror("Não foi possível imprimir",
                                     f"{e}\n\nUse 'Salvar PDF…' e imprima o arquivo manualmente.")

    def visualizar(self):
        frente = dc.is_tecnico(self.sel) if self.sel else False
        caminho = self._gerar(frente, conferir=False)
        if caminho:
            try:
                abrir_arquivo(caminho)
                self.var_status.set(f"Aberto: {caminho}")
            except Exception as e:
                messagebox.showerror("Erro ao abrir o PDF", str(e))

    def salvar(self):
        if not self.sel:
            return
        destino = filedialog.asksaveasfilename(
            defaultextension=".pdf", initialfile=self._nome_arquivo("diploma"),
            filetypes=[("Arquivo PDF", "*.pdf")])
        if destino and self._gerar(True, destino, conferir=False):
            self.var_status.set(f"PDF salvo em {destino}")
            messagebox.showinfo("Salvo", f"PDF salvo em:\n{destino}")

    def salvar_docx(self):
        """gera o diploma (frente e verso) em Word, a partir do modelo DIPLOMA - PADRÃO.docx"""
        if not self.sel:
            messagebox.showinfo("Aviso", "Selecione um aluno primeiro.")
            return
        if not dc.is_tecnico(self.sel):
            messagebox.showinfo("Aviso", "O modelo de diploma padrão é só para o curso Técnico em Enfermagem.")
            return
        if not self._confirmar_faltas():
            return
        modelo = diploma_docx.localizar_modelo()
        if modelo is None:
            messagebox.showinfo("Modelo do diploma",
                                "Não encontrei o arquivo «DIPLOMA - PADRÃO.docx» ao lado do programa.\n"
                                "Selecione o modelo na próxima janela.")
            escolhido = filedialog.askopenfilename(title="Selecione o modelo DIPLOMA - PADRÃO.docx",
                                                   filetypes=[("Documento do Word", "*.docx")])
            if not escolhido:
                return
            modelo = Path(escolhido)
        destino = filedialog.asksaveasfilename(
            defaultextension=".docx", initialfile=self._nome_arquivo("diploma", ".docx"),
            filetypes=[("Documento do Word", "*.docx")])
        if not destino:
            return
        opc = {"ano": self.var_ano.get(), "per_ini": self.var_ini.get(), "per_fim": self.var_fim.get()}
        try:
            diploma_docx.gerar_docx(self.sel, destino, opc, modelo=modelo)
        except ImportError:
            messagebox.showerror("Falta um componente",
                                 "Para gerar o DOCX é preciso instalar o lxml:\n\npip install lxml")
            return
        except diploma_docx.ModeloInvalido as e:
            messagebox.showerror("Modelo diferente do esperado", str(e))
            return
        except PermissionError:
            messagebox.showerror("Arquivo em uso",
                                 "Não consegui gravar o arquivo. Se ele estiver aberto no Word, feche-o e tente de novo.")
            return
        except Exception as e:
            messagebox.showerror("Erro ao gerar o DOCX", str(e))
            return
        self.var_status.set(f"DOCX salvo em {destino}")
        if messagebox.askyesno("Salvo", f"DOCX salvo em:\n{destino}\n\nAbrir agora no Word?"):
            try:
                abrir_arquivo(destino)
            except Exception as e:
                messagebox.showerror("Erro ao abrir o DOCX", str(e))


def main():
    raiz = tk.Tk()
    raiz.withdraw()
    caminho = dc.localizar_csv()
    if caminho is None:
        messagebox.showinfo("Base de dados",
                            "Não encontrei o arquivo FichaAvaliacao_completa.csv ao lado do programa.\n"
                            "Selecione o arquivo na próxima janela.")
        escolhido = filedialog.askopenfilename(title="Selecione a base de dados",
                                               filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if not escolhido:
            sys.exit(0)
        caminho = Path(escolhido)
    try:
        base, dup = dc.carregar_base(caminho)
    except Exception as e:
        messagebox.showerror("Erro ao carregar dados", str(e))
        sys.exit(1)
    raiz.destroy()
    App(base, dup, Path(caminho)).mainloop()


if __name__ == "__main__":
    main()
