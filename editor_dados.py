#!/usr/bin/env python3
"""
Módulo de edição dos dados que saem no diploma / histórico:
  * notas finais de teoria e de estágio de cada componente curricular;
  * campo (local) de estágio de cada disciplina;
  * período do estágio ("No Período de ...").

Usado por buscar_aluno_gui.py (botão "Editar notas e estágios" ou duplo clique
numa linha das abas Notas finais / Estágios).

Como as alterações são gravadas
  * antes de qualquer gravação é feita uma cópia de segurança do CSV na pasta
    "backups" (ao lado do CSV; ficam as 30 mais recentes);
  * o CSV é relido do disco e só as linhas do aluno são alteradas. O programa
    principal descarta linhas repetidas ao carregar, mas aqui elas NÃO são
    apagadas do arquivo: todas as cópias idênticas do aluno recebem a mesma
    alteração (assim continuam idênticas);
  * a gravação é feita em arquivo temporário e só então substitui o original,
    então uma falha no meio nunca deixa o CSV pela metade;
  * cada campo alterado é registrado em "historico_alteracoes.csv"
    (data/hora, aluno, matrícula, campo, valor anterior, valor novo).
"""
import csv
import os
import re
import shutil
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

import diploma_core as dc

NOME_LOG = "historico_alteracoes.csv"
PASTA_BACKUP = "backups"
MAX_BACKUPS = 30

SEP_PERIODO = " à "          # separador usado na base: "10/01/11 à 07/10/11"
DATA_ABERTA = "00/00/00"     # "00/00/00" no término = sem data (sai em branco no diploma)
DATAS_ABERTAS = {"00/00/00", "00/00/0000"}

COR_ALTERADO = "#FFF3B0"
COR_FALTA = "#B3261E"
COR_CINZA = "#6B7785"

# áreas de estágio: (rótulo, coluna, entra_no_quadro_padrão)
AREAS = ([(rot, col, True) for rot, col in dc.ESTAGIO_DIPLOMA]
         + [(rot, col, False) for rot, col in dc.ESTAGIO_EXTRAS])


# =====================================================================
#  Validação (funções puras, sem interface)
# =====================================================================
def validar_nota(txt):
    """Retorna (valor_normalizado, erro). Texto vazio é válido (apaga a nota).

    Aceita vírgula ou ponto; grava com ponto, como a base já faz (8,5 -> 8.5).
    Inteiros ganham ".0" para seguir o padrão da base (8 -> 8.0).
    """
    t = (txt or "").strip().replace(",", ".")
    if t == "":
        return "", None
    if not re.fullmatch(r"\d{1,2}(\.\d{1,4})?", t):
        return None, "use apenas números, por exemplo 8.5 (ou 8,5)"
    if float(t) > 10:
        return None, "a nota deve estar entre 0 e 10"
    inteira, _, decimal = t.partition(".")
    return f"{int(inteira)}.{decimal or '0'}", None


def _ano_completo(aa):
    return 2000 + aa if aa <= 69 else 1900 + aa


def normalizar_data(txt, permitir_aberta=False):
    """Valida e devolve 'dd/mm/aa' ou 'dd/mm/aaaa' (o ano fica como foi digitado).

    Levanta ValueError com a explicação quando a data não serve.
    """
    t = re.sub(r"[-.]", "/", (txt or "").strip())
    if permitir_aberta and t in DATAS_ABERTAS:
        return DATA_ABERTA
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})", t)
    if not m:
        raise ValueError("use o formato dd/mm/aa, por exemplo 10/01/11")
    dia, mes, ano_txt = int(m.group(1)), int(m.group(2)), m.group(3)
    ano = int(ano_txt) if len(ano_txt) == 4 else _ano_completo(int(ano_txt))
    try:
        if not 1900 <= ano <= 2099:
            raise ValueError
        datetime(ano, mes, dia)
    except ValueError:
        raise ValueError(f"a data {t} não existe no calendário")
    return f"{dia:02d}/{mes:02d}/{ano_txt}"


def _data_obj(data):
    d, m, a = data.split("/")
    ano = int(a) if len(a) == 4 else _ano_completo(int(a))
    return datetime(ano, int(m), int(d))


def montar_periodo(inicio, fim, sem_termino):
    """Retorna (texto_para_a_base, erro). Tudo vazio = sem período ('')."""
    inicio, fim = (inicio or "").strip(), (fim or "").strip()
    if not inicio and not fim and not sem_termino:
        return "", None
    if not inicio:
        return None, "informe a data de início"
    try:
        i = normalizar_data(inicio)
    except ValueError as e:
        return None, f"início: {e}"
    if sem_termino:
        return f"{i}{SEP_PERIODO}{DATA_ABERTA}", None
    if not fim:
        return None, "informe a data de término ou marque 'sem data de término'"
    try:
        f = normalizar_data(fim, permitir_aberta=True)
    except ValueError as e:
        return None, f"término: {e}"
    if f != DATA_ABERTA and _data_obj(i) > _data_obj(f):
        return None, "o término é anterior ao início"
    return f"{i}{SEP_PERIODO}{f}", None


def separar_periodo(texto):
    """'10/01/11 à 07/10/11' -> ('10/01/11', '07/10/11', False)."""
    texto = (texto or "").strip()
    if not texto:
        return "", "", False
    partes = [p.strip() for p in texto.split("à")]
    if len(partes) != 2:
        # formato desconhecido: mostra o texto cru para a pessoa corrigir
        return texto, "", False
    ini, fim = partes
    if fim in DATAS_ABERTAS:
        return ini, "", True
    return ini, fim, False


# =====================================================================
#  Gravação segura
# =====================================================================
class RegistroAlterado(Exception):
    """O aluno não foi encontrado no arquivo como estava quando foi carregado."""


def _fazer_backup(csv_path: Path):
    if not csv_path.exists():
        return None
    pasta = csv_path.parent / PASTA_BACKUP
    pasta.mkdir(exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = pasta / f"{csv_path.stem}_{carimbo}{csv_path.suffix}"
    shutil.copy2(csv_path, destino)
    antigos = sorted(pasta.glob(f"{csv_path.stem}_*{csv_path.suffix}"))
    for velho in antigos[:-MAX_BACKUPS]:      # mantém só os mais recentes
        try:
            velho.unlink()
        except OSError:
            pass
    return destino


def _ler_csv(csv_path):
    with open(csv_path, encoding="utf-8", newline="") as f:
        leitor = csv.DictReader(f, delimiter=";")
        linhas = list(leitor)
        return linhas, list(leitor.fieldnames)


def _gravar_csv(csv_path: Path, linhas, fieldnames):
    """backup + arquivo temporário + conferência + substituição do original"""
    backup = _fazer_backup(csv_path)
    tmp = csv_path.with_name(csv_path.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";",
                               lineterminator="\n", restval="", extrasaction="ignore")
            w.writeheader()
            w.writerows(linhas)
        with open(tmp, encoding="utf-8", newline="") as f:     # confere o que foi gravado
            n = sum(1 for _ in csv.DictReader(f, delimiter=";"))
        if n != len(linhas):
            raise IOError(f"verificação falhou: {n} linhas gravadas, esperado {len(linhas)}")
        os.replace(tmp, csv_path)
    except Exception:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        raise
    return backup


def _registrar_historico(csv_path: Path, linha, alteracoes, rotulos):
    caminho = csv_path.parent / NOME_LOG
    novo = not caminho.exists()
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(caminho, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        if novo:
            w.writerow(["data_hora", "aluno", "matricula", "campo", "descricao",
                        "valor_anterior", "valor_novo"])
        for campo, (antigo, novo_valor) in alteracoes.items():
            w.writerow([agora, (linha.get("NOME") or "").strip(),
                        (linha.get("NUM_MAT") or "").strip(), campo,
                        rotulos.get(campo, campo), antigo, novo_valor])


def aplicar_alteracoes(csv_path, linha, alteracoes, rotulos):
    """Grava {campo: (antigo, novo)} no CSV para o aluno `linha`.

    Relê o arquivo e altera as linhas que são idênticas ao registro carregado
    (inclusive cópias repetidas). Só depois de gravar com sucesso é que `linha`
    (o registro em memória do programa) é atualizada.
    Retorna (backup, qtd_linhas_alteradas, aviso_ou_None).
    """
    csv_path = Path(csv_path)
    linhas, fieldnames = _ler_csv(csv_path)
    faltam = [c for c in alteracoes if c not in fieldnames]
    if faltam:
        raise RegistroAlterado("O arquivo não tem a(s) coluna(s): " + ", ".join(faltam))
    # como o aluno estava quando foi carregado (antes desta edição)
    chave = {c: (linha.get(c) or "") for c in fieldnames}
    alvo = [r for r in linhas if all((r.get(c) or "") == chave[c] for c in fieldnames)]
    if not alvo:
        raise RegistroAlterado(
            "Este aluno não foi encontrado no arquivo do jeito que estava ao abrir o programa.\n"
            "O arquivo foi alterado por outro programa? Feche e abra o programa de novo.")
    for r in alvo:
        for campo, (_antigo, novo) in alteracoes.items():
            r[campo] = novo
    backup = _gravar_csv(csv_path, linhas, fieldnames)
    for campo, (_antigo, novo) in alteracoes.items():
        linha[campo] = novo
    aviso = None
    try:
        _registrar_historico(csv_path, linha, alteracoes, rotulos)
    except Exception as e:      # o histórico não pode desfazer o que já foi salvo
        aviso = f"Dados salvos, mas não consegui gravar o histórico: {e}"
    return backup, len(alvo), aviso


# =====================================================================
#  Widget de período (início / término / sem término)
# =====================================================================
class CampoPeriodo(ttk.Frame):
    def __init__(self, master, texto_inicial="", ao_mudar=None):
        super().__init__(master)
        self._ao_mudar = None
        self.v_ini = tk.StringVar()
        self.v_fim = tk.StringVar()
        self.v_and = tk.BooleanVar()

        self.e_ini = ttk.Entry(self, textvariable=self.v_ini, width=11, justify="center")
        self.e_fim = ttk.Entry(self, textvariable=self.v_fim, width=11, justify="center")
        self.chk = ttk.Checkbutton(self, text="sem data de término (sai em branco)",
                                   variable=self.v_and, command=self._alternar)
        ttk.Label(self, text="de").grid(row=0, column=0, padx=(0, 6))
        self.e_ini.grid(row=0, column=1)
        ttk.Label(self, text="a").grid(row=0, column=2, padx=6)
        self.e_fim.grid(row=0, column=3)
        self.chk.grid(row=0, column=4, padx=(14, 0))

        self.definir(texto_inicial)
        self._ao_mudar = ao_mudar
        for v in (self.v_ini, self.v_fim, self.v_and):
            v.trace_add("write", self._notificar)

    def _notificar(self, *_):
        if self._ao_mudar:
            self._ao_mudar()

    def _alternar(self):
        self.e_fim.configure(state="disabled" if self.v_and.get() else "normal")

    def definir(self, texto):
        ini, fim, sem = separar_periodo(texto)
        self.v_ini.set(ini)
        self.v_fim.set(fim)
        self.v_and.set(sem)
        self._alternar()

    def estado(self):
        return (self.v_ini.get().strip(), self.v_fim.get().strip(), bool(self.v_and.get()))

    def obter(self):
        return montar_periodo(*self.estado())


# =====================================================================
#  Janela de edição
# =====================================================================
COLUNAS_EDITAVEIS = ("teoria", "est")


class EditorAluno(tk.Toplevel):
    """Janela modal para editar notas finais, campos de estágio e período de um aluno.

    Parâmetros (todos nomeados):
      linha              - registro do aluno (o mesmo dict usado pelo programa)
      csv_path           - arquivo da base
      locais_conhecidos  - nomes de locais já usados na base (autocompletar)
      ao_salvar          - função(qtd_campos, backup, qtd_linhas) chamada após salvar
      aba_inicial        - 0 = Notas finais, 1 = Estágios
      foco               - (tipo, campo): ("nota", "ANATOMIA") ou ("local", "LOCIR")
    """

    def __init__(self, master, *, linha, csv_path, locais_conhecidos=(), ao_salvar=None,
                 aba_inicial=0, foco=None):
        super().__init__(master)
        self.linha = linha
        self.csv_path = Path(csv_path)
        self.locais = list(locais_conhecidos)
        self.ao_salvar = ao_salvar
        self._edicao = None
        self._col_atual = "teoria"

        nome = dc.limpa(linha.get("NOME", ""))
        self.title(f"Editar notas e estágios — {nome}")
        self.minsize(760, 440)
        self.transient(master)
        self.configure(bg=ttk.Style(self).lookup("TFrame", "background") or self.cget("bg"))

        # (iid, coluna) -> campo do CSV   |   campo -> rótulo legível
        self.celulas, self.rotulos = {}, {}
        for i, (mod, comp, _t, _p, _e, col_t, col_e) in enumerate(dc.CURRICULO):
            self.celulas[(str(i), "teoria")] = col_t
            self.rotulos[col_t] = f"{comp} (nota final de teoria)"
            if col_e:
                self.celulas[(str(i), "est")] = col_e
                self.rotulos[col_e] = f"{comp} (nota de estágio)"
        for rot, col, _padrao in AREAS:
            self.rotulos[col] = f"Campo de estágio — {rot}"
        self.rotulos[dc.CAMPO_PERIODO] = "Período do estágio"

        self.campos_nota = list(dict.fromkeys(self.celulas.values()))
        self.originais = {c: (linha.get(c) or "").strip() for c in self.campos_nota}
        self.valores = dict(self.originais)

        self._build_ui(nome)
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.bind("<Control-s>", lambda e: self.salvar())

        self._centralizar(master, 940, 650)
        self.nb.select(aba_inicial)
        try:
            self.wait_visibility()
            self.grab_set()
        except tk.TclError:
            pass
        self.focus_set()
        if foco:
            self.after(60, lambda: self._aplicar_foco(*foco))

    def _centralizar(self, master, largura, altura):
        master.update_idletasks()
        tela_l, tela_a = self.winfo_screenwidth(), self.winfo_screenheight()
        largura = min(largura, tela_l - 20)
        altura = min(altura, tela_a - 90)          # deixa espaço para a barra de tarefas
        x = master.winfo_rootx() + (master.winfo_width() - largura) // 2
        y = master.winfo_rooty() + (master.winfo_height() - altura) // 2
        x = max(0, min(x, tela_l - largura - 10))
        y = max(0, min(y, tela_a - altura - 60))
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    def _aplicar_foco(self, tipo, campo):
        if tipo == "nota":
            for (iid, col), c in self.celulas.items():
                if c == campo:
                    self.nb.select(0)
                    self.tree.selection_set(iid)
                    self.tree.see(iid)
                    self._iniciar_edicao(iid, col)
                    return
        elif tipo == "local" and campo in self.var_locais:
            self.nb.select(1)
            self.combos[campo].focus_set()

    # ------------------------------------------------------------ UI ----
    def _build_ui(self, nome):
        topo = ttk.Frame(self, padding=(16, 12, 16, 0))
        topo.pack(fill="x")
        ttk.Label(topo, text=nome, font=("Helvetica", 13, "bold")).pack(anchor="w")
        ttk.Label(topo, foreground=COR_CINZA,
                  text=f"Matrícula {dc.limpa(self.linha.get('NUM_MAT', ''))}   •   "
                       f"Turma {dc.limpa(self.linha.get('TURMA', ''))}   •   "
                       f"{dc.limpa(self.linha.get('CURSO', ''))}").pack(anchor="w")

        # rodapé primeiro (side="bottom"): se a tela for pequena, o que encolhe é a tabela
        rodape = ttk.Frame(self, padding=(16, 6, 16, 14))
        rodape.pack(side="bottom", fill="x")
        self.var_msg = tk.StringVar()
        ttk.Label(rodape, textvariable=self.var_msg, foreground=COR_FALTA).pack(side="left")
        ttk.Button(rodape, text="Cancelar", command=self._fechar).pack(side="right")
        ttk.Button(rodape, text="💾  Salvar alterações  (Ctrl+S)", style="Primary.TButton",
                   command=self.salvar).pack(side="right", padx=8)

        self.nb = ttk.Notebook(self)
        self.nb.pack(side="top", fill="both", expand=True, padx=16, pady=10)
        self._build_aba_notas()
        self._build_aba_estagios()

    # ---- aba de notas ---------------------------------------------------
    def _build_aba_notas(self):
        aba = ttk.Frame(self.nb, padding=10)
        self.nb.add(aba, text="  Notas finais  ")
        self.var_dica = tk.StringVar()
        self.var_contador = tk.StringVar()

        acoes = ttk.Frame(aba)
        acoes.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Button(acoes, text="Restaurar linha", command=self._restaurar_selecionado).pack(side="left")
        ttk.Label(acoes, textvariable=self.var_contador, foreground=COR_CINZA).pack(side="right")
        ttk.Label(acoes, foreground=COR_CINZA,
                  text="   Enter confirma e desce  •  Tab: próxima célula  •  "
                       "Esc cancela  •  vazio apaga").pack(side="left")
        ttk.Label(aba, textvariable=self.var_dica, font=("Helvetica", 10, "bold")).pack(
            side="bottom", anchor="w", pady=(8, 0))

        quadro = ttk.Frame(aba)
        quadro.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(quadro, columns=("mod", "comp", "teoria", "est"),
                                 show="headings", selectmode="browse")
        for cid, tit, larg, anc in (("mod", "Módulo", 70, "center"),
                                    ("comp", "Componente curricular", 330, "w"),
                                    ("teoria", "Teoria  (duplo clique)", 170, "center"),
                                    ("est", "Estágio  (duplo clique)", 170, "center")):
            self.tree.heading(cid, text=tit)
            self.tree.column(cid, width=larg, anchor=anc, stretch=(cid == "comp"))
        self.tree.tag_configure("alterado", background=COR_ALTERADO)
        self.tree.tag_configure("falta", foreground=COR_FALTA)
        self._sb = ttk.Scrollbar(quadro, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self._ao_rolar)
        self.tree.pack(side="left", fill="both", expand=True)
        self._sb.pack(side="right", fill="y")

        for i, (mod, comp, *_resto) in enumerate(dc.CURRICULO):
            self.tree.insert("", "end", iid=str(i), values=("", "", "", ""))
            self._atualizar_linha(str(i))
        self.tree.bind("<Double-1>", self._duplo_clique)
        self.tree.bind("<Return>", lambda e: self._editar_selecionado())
        self.tree.bind("<F2>", lambda e: self._editar_selecionado())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._atualizar_dica())

        self._atualizar_contador()

    def _ao_rolar(self, a, b):
        self._sb.set(a, b)
        if self._edicao:                  # o campo de edição não acompanha a rolagem
            self._cancelar_edicao()

    def _mudou(self, campo):
        return self.valores[campo] != self.originais[campo]

    def _txt_celula(self, campo):
        if campo is None:
            return "—"
        v = self.valores[campo]
        txt = v if v else "(falta)"
        if self._mudou(campo):
            txt += f"   (era {self.originais[campo] or 'vazio'})"
        return txt

    def _atualizar_linha(self, iid):
        i = int(iid)
        mod, comp = dc.CURRICULO[i][0], dc.CURRICULO[i][1]
        c_t = self.celulas[(iid, "teoria")]
        c_e = self.celulas.get((iid, "est"))
        tags = []
        if any(self._mudou(c) for c in (c_t, c_e) if c):
            tags.append("alterado")
        if not self.valores[c_t] or (c_e and not self.valores[c_e]):
            tags.append("falta")
        self.tree.item(iid, tags=tags, values=(
            f"Mód. {mod:02d}", comp, self._txt_celula(c_t), self._txt_celula(c_e)))
        self._atualizar_contador()
        self._atualizar_dica()

    def _atualizar_contador(self):
        n = sum(1 for c in self.campos_nota if self._mudou(c))
        self.var_contador.set(f"{n} nota(s) alterada(s)" if n else "nenhuma nota alterada")

    def _como_sai(self, campo):
        return dc.fmt_nota(dc.parse_nota(self.valores[campo])) or "em branco"

    def _atualizar_dica(self):
        sel = self.tree.selection()
        if not sel:
            self.var_dica.set("")
            return
        iid = sel[0]
        partes = [f"teoria {self._como_sai(self.celulas[(iid, 'teoria')])}"]
        c_e = self.celulas.get((iid, "est"))
        if c_e:
            partes.append(f"estágio {self._como_sai(c_e)}")
        self.var_dica.set("Sai no diploma:  " + "   •   ".join(partes) + "     (1 casa decimal)")

    # ---- edição direto na tabela ---------------------------------------
    def _duplo_clique(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col = {"#3": "teoria", "#4": "est"}.get(self.tree.identify_column(event.x))
        if iid and col:
            self.tree.selection_set(iid)
            self._iniciar_edicao(iid, col)

    def _editar_selecionado(self):
        sel = self.tree.selection()
        if sel:
            self._iniciar_edicao(sel[0], self._col_atual)

    def _iniciar_edicao(self, iid, col):
        self._cancelar_edicao()
        campo = self.celulas.get((iid, col))
        if campo is None:
            self.var_msg.set("Este componente não tem nota de estágio.")
            return
        self._col_atual = col
        self.tree.see(iid)
        self.update_idletasks()
        bbox = self.tree.bbox(iid, col)
        if not bbox:
            return
        x, y, w, h = bbox
        var = tk.StringVar(value=self.valores[campo])
        ent = ttk.Entry(self.tree, textvariable=var, justify="center")
        ent.place(x=x, y=y, width=w, height=h)
        ent.focus_set()
        ent.select_range(0, "end")
        ent.icursor("end")
        ent.bind("<Return>", lambda e: self._confirmar_edicao("baixo"))
        ent.bind("<KP_Enter>", lambda e: self._confirmar_edicao("baixo"))
        ent.bind("<Tab>", lambda e: self._confirmar_edicao("proxima") or "break")
        ent.bind("<Escape>", lambda e: self._cancelar_edicao())
        ent.bind("<FocusOut>", lambda e: self._confirmar_edicao(None, ao_sair=True))
        self._edicao = (iid, col, campo, ent, var)
        self.var_msg.set("")

    def _cancelar_edicao(self):
        if not self._edicao:
            return
        ent = self._edicao[3]
        self._edicao = None
        ent.destroy()

    def _confirmar_edicao(self, avancar, ao_sair=False):
        if not self._edicao:
            return
        iid, col, campo, ent, var = self._edicao
        texto = var.get().strip()
        # texto igual ao valor atual: não valida nem normaliza (nunca cria alteração falsa)
        valor, erro = (texto, None) if texto == self.valores[campo] else validar_nota(texto)
        nome = dc.CURRICULO[int(iid)][1]
        if erro:
            if ao_sair:                   # perdeu o foco com valor inválido: descarta
                self._cancelar_edicao()
                self.var_msg.set(f"{nome}: {erro} (alteração descartada)")
            else:                         # Enter/Tab: mantém aberto para corrigir
                self.var_msg.set(f"{nome}: {erro}")
                ent.focus_set()
            return
        self._edicao = None
        ent.destroy()
        self.valores[campo] = valor
        self.var_msg.set("")
        self._atualizar_linha(iid)
        if avancar:
            prox = self._proxima_celula(iid, col, avancar)
            if prox:
                self.tree.selection_set(prox[0])
                self._iniciar_edicao(*prox)
            else:
                self.tree.focus_set()

    def _proxima_celula(self, iid, col, modo):
        ordem = [(str(i), c) for i in range(len(dc.CURRICULO)) for c in COLUNAS_EDITAVEIS
                 if (str(i), c) in self.celulas]
        if modo == "baixo":               # mesma coluna, próxima linha que tenha a célula
            ordem = [t for t in ordem if t[1] == col]
        try:
            return ordem[ordem.index((iid, col)) + 1]
        except (ValueError, IndexError):
            return None

    def _restaurar_selecionado(self):
        sel = self.tree.selection()
        if sel:
            self._cancelar_edicao()
            for col in COLUNAS_EDITAVEIS:
                campo = self.celulas.get((sel[0], col))
                if campo:
                    self.valores[campo] = self.originais[campo]
            self._atualizar_linha(sel[0])

    # ---- aba de estágios -------------------------------------------------
    def _build_aba_estagios(self):
        aba = ttk.Frame(self.nb, padding=12)
        self.nb.add(aba, text="  Estágios  ")

        quadro = ttk.LabelFrame(aba, text="Campo de estágio de cada disciplina", padding=10)
        quadro.pack(fill="x")
        quadro.columnconfigure(1, weight=1)
        self.var_locais, self.combos, self._locais_iniciais = {}, {}, {}
        r = 0
        for padrao, titulo in ((True, "No quadro do diploma"),
                               (False, "Fora do quadro padrão (só saem no diploma se preenchidos)")):
            ttk.Label(quadro, text=titulo, foreground=COR_CINZA).grid(
                row=r, column=0, columnspan=2, sticky="w", pady=(0 if r == 0 else 10, 4))
            r += 1
            for rot, col, e_padrao in AREAS:
                if e_padrao != padrao:
                    continue
                ttk.Label(quadro, text=rot).grid(row=r, column=0, sticky="w", padx=(0, 14), pady=3)
                inicial = dc.limpa(self.linha.get(col, ""))
                var = tk.StringVar(value=inicial)
                cb = ttk.Combobox(quadro, textvariable=var, values=self.locais, width=60)
                cb.grid(row=r, column=1, sticky="ew", pady=3)
                self.var_locais[col], self.combos[col] = var, cb
                self._locais_iniciais[col] = inicial
                r += 1

        per = ttk.LabelFrame(aba, text="Período do estágio  (No Período de ... a ...)", padding=10)
        per.pack(fill="x", pady=(12, 0))
        self.per = CampoPeriodo(per, self.linha.get(dc.CAMPO_PERIODO, ""),
                                ao_mudar=self._atualizar_previa)
        self.per.pack(anchor="w")
        self._per_inicial = self.per.estado()
        self.var_previa = tk.StringVar()
        ttk.Label(per, textvariable=self.var_previa, font=("Helvetica", 10, "bold")).pack(
            anchor="w", pady=(8, 0))
        self._atualizar_previa()

        ttk.Label(aba, foreground=COR_CINZA, wraplength=860, justify="left",
                  text="Apagar o local de uma disciplina deixa o campo em branco no diploma. "
                       "O mesmo período vale para todas as disciplinas de estágio.").pack(
            anchor="w", pady=(10, 0))

    def _atualizar_previa(self):
        texto, erro = self.per.obter()
        if erro:
            self.var_previa.set(f"Corrija o período: {erro}")
            return
        ini, fim = dc.periodo_estagio({dc.CAMPO_PERIODO: texto})
        self.var_previa.set(f"Sai no diploma:  No Período de: {ini or '__/__/__'} a {fim or '__/__/____'}")

    # ---- salvar ------------------------------------------------------------
    def _coletar_alteracoes(self):
        """Retorna ({campo: (antigo, novo)}, [erros])."""
        alteracoes, erros = {}, []
        for c in self.campos_nota:
            if self._mudou(c):
                alteracoes[c] = (self.originais[c], self.valores[c])

        for rot, col, _padrao in AREAS:
            novo = " ".join(self.var_locais[col].get().split())     # tira espaços duplicados
            if novo != self._locais_iniciais[col]:
                alteracoes[col] = ((self.linha.get(col) or "").strip(), novo)

        if self.per.estado() != self._per_inicial:
            texto, erro = self.per.obter()
            if erro:
                erros.append(f"Período do estágio: {erro}")
            else:
                antigo = (self.linha.get(dc.CAMPO_PERIODO) or "").strip()
                if texto != antigo:
                    alteracoes[dc.CAMPO_PERIODO] = (antigo, texto)
        return alteracoes, erros

    def salvar(self):
        self._confirmar_edicao(None)                  # fecha uma edição pendente
        if self._edicao:                              # continuou inválida
            return
        alteracoes, erros = self._coletar_alteracoes()
        if erros:
            messagebox.showwarning("Corrija antes de salvar", "\n".join(erros), parent=self)
            return
        if not alteracoes:
            messagebox.showinfo("Nada a salvar", "Nenhum campo foi alterado.", parent=self)
            return

        linhas = [f"• {self.rotulos.get(c, c)}:  {a or '(vazio)'}  →  {n or '(vazio)'}"
                  for c, (a, n) in alteracoes.items()]
        if len(linhas) > 15:
            linhas = linhas[:15] + [f"… e mais {len(alteracoes) - 15} alteração(ões)"]
        if not messagebox.askyesno(
                "Confirmar alterações",
                f"Salvar {len(alteracoes)} alteração(ões) no cadastro de "
                f"{dc.limpa(self.linha.get('NOME', ''))}?\n\n" + "\n".join(linhas),
                parent=self):
            return

        try:
            backup, n_linhas, aviso = aplicar_alteracoes(
                self.csv_path, self.linha, alteracoes, self.rotulos)
        except PermissionError:
            messagebox.showerror(
                "Não foi possível salvar",
                f"O arquivo '{self.csv_path.name}' está bloqueado.\n"
                "Se ele estiver aberto no Excel (ou outro programa), feche-o e tente de novo.\n"
                "Nada foi alterado.", parent=self)
            return
        except RegistroAlterado as e:
            messagebox.showerror("Não foi possível salvar", f"{e}\n\nNada foi alterado.", parent=self)
            return
        except Exception as e:
            messagebox.showerror("Não foi possível salvar",
                                 f"{e}\n\nNada foi alterado.", parent=self)
            return

        if aviso:
            messagebox.showwarning("Atenção", aviso, parent=self)
        self._cancelar_edicao()
        self.grab_release()
        self.destroy()
        if self.ao_salvar:
            self.ao_salvar(len(alteracoes), backup, n_linhas)

    def _fechar(self):
        self._cancelar_edicao()
        alteracoes, erros = self._coletar_alteracoes()
        if (alteracoes or erros) and not messagebox.askyesno(
                "Descartar alterações?",
                "Há alterações que ainda não foram salvas. Fechar mesmo assim?", parent=self):
            return
        self.grab_release()
        self.destroy()
