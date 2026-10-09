#!/usr/bin/env python3
"""
Núcleo do programa: leitura da base (CSV), regras de associação das notas às
linhas do diploma padrão e geração do PDF (frente = diploma, verso = histórico
com notas, campos de estágio e período), seguindo o modelo DIPLOMA - PADRÃO.docx.
"""
import csv
import re
import sys
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from xml.sax.saxutils import escape

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

# =====================================================================
#  CURRÍCULO DO DIPLOMA PADRÃO  (ajuste aqui se o modelo mudar)
#  (módulo, componente, CH teoria, CH prática, CH estágio,
#   coluna da NOTA FINAL de teoria no CSV, coluna da nota de estágio no CSV)
#  A nota final é a coluna "base" do CSV (já é a média das notas I, II, III...).
# =====================================================================
CURRICULO = [
    (1, "Anatomia e Fisiologia Humana",            80, 0, 0,   "ANATOMIA",   None),
    (1, "Microbiologia e Parasitologia",           50, 0, 0,   "MICRO",      None),
    (1, "Higiene e Profilaxia",                    30, 0, 0,   "HIGIENE",    None),
    (1, "Nutrição e Dietética",                    30, 0, 0,   "NUTRICAO",   None),
    (1, "Epidemiologia",                           30, 0, 0,   "ESTUDOS",    None),     # a confirmar
    (1, "Ética e Psicologia Aplicada",             40, 0, 0,   "ETICA",      None),
    (1, "Língua Portuguesa",                       60, 0, 0,   "HUMANIZACA", None),     # a confirmar
    (1, "Farmacologia",                            30, 0, 0,   "Farmaco",    None),
    (1, "Matemática Aplicada",                     40, 0, 0,   "MAT",        None),
    (2, "Introdução à Enfermagem",                120, 50, 150, "INTRODUCAO", "INTRODUE"),
    (2, "Enfermagem em Clínica Médica 01",         30, 50, 50,  "Medica01",   "MEDICAE"),
    (2, "Enfermagem em Centro Cirúrgico e CME",    40, 25, 60,  "Cirurgica01", "CIRURGICAE"),
    (2, "Enfermagem Materno Infantil 01",          40, 0, 50,   "Materno01",  "MATERNOE"),
    (2, "Noções de Administração na Unidade de Enf. 1", 20, 0, 0, "Adm01",    None),
    (3, "Enfermagem em Clínica Médica 2",          30, 0, 60,   "Medica02",   "Medica02E"),
    (3, "Enfermagem em Clínica Cirúrgica",         30, 0, 40,   "Cirurgica02", "Cirurgica02E"),
    (3, "Enfermagem Materno Infantil 2",           50, 0, 40,   "Materno02",  "Materno02E"),
    (3, "Enfermagem Pediátrica",                   50, 25, 0,   "PEDIATRICA", None),
    (3, "Noções de Administração na Unidade de Enf. 2", 20, 0, 0, "Adm02",    None),
    (3, "Enfermagem em Saúde Coletiva",            65, 25, 0,   "SAUDE",      None),
    (4, "Enfermagem em Saúde Mental",              40, 0, 50,   "PSIQUIATRI", "PSIQUIATRE"),
    (4, "Enfermagem em Emergência",                60, 0, 50,   "EMERGENCIA", "EMERGENCIE"),
    (4, "Enfermagem em Geriátria",                 90, 50, 0,   "GER",        None),
    (4, "Enfermagem Médica 3",                     30, 0, 50,   "Medica03",   "Medica03E"),
]

# Campos de estágio exatamente como no quadro "Disciplinas | Campo de Estágios"
ESTAGIO_DIPLOMA = [
    ("Introdução à Enf.", "LOINT"),
    ("Enf. Médica", "LOMED"),
    ("Enf. Cirúrgica", "LOCIR"),
    ("Enf. Materno", "LOMATR"),
    ("Enf. Emergência", "LOEMER"),
]
# Existem na base mas não fazem parte do quadro padrão: só entram se preenchidos
ESTAGIO_EXTRAS = [
    ("Enf. Pediátrica", "LOPED"),
    ("Enf. Saúde Mental", "LOPSQ"),
    ("Enf. Saúde Coletiva", "LOSAU"),
]
CAMPO_PERIODO = "PERINT"

CURSO_DIPLOMA = "Técnico em Enfermagem"

CAMPOS_NAO_BRUTOS = {"Foto", "OBSERVA", "Inativo", "TurmaInativa", "MATRICULA"}


# ---------------------------------------------------------------- util ----
def normalizar(txt):
    txt = unicodedata.normalize("NFKD", txt or "")
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return txt.lower().strip()


def limpa(v):
    """texto sem espaços; campos só com '_' (placeholder do sistema antigo) viram vazio"""
    v = (v or "").strip()
    return "" if (not v or set(v) <= {"_", " "}) else v


def parse_nota(v):
    """Lê a nota do CSV. Usa o valor em ponto flutuante (como o sistema antigo),
    o que reproduz o diploma modelo: 6,25 -> 6,3 e 7,05 -> 7,0."""
    v = (v or "").strip().replace(",", ".")
    if not v:
        return None
    try:
        return Decimal(float(v))
    except (InvalidOperation, ValueError):
        return None


def fmt_nota(d):
    """nota com 1 casa decimal e vírgula (arredondamento escolar: 0,05 sobe)"""
    if d is None:
        return ""
    q = d.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return f"{q:.1f}".replace(".", ",")


def fmt_ch(n):
    return f"{n:,}".replace(",", ".")


def is_tecnico(linha):
    return normalizar(linha.get("CURSO", "")) == normalizar(CURSO_DIPLOMA)


# ------------------------------------------------------------- base CSV ----
def localizar_csv():
    p = BASE_DIR / "FichaAvaliacao_completa.csv"
    if p.exists():
        return p
    achados = sorted(BASE_DIR.glob("FichaAvaliacao_completa*.csv"))
    return achados[0] if achados else None


def carregar_base(caminho):
    """lê o CSV, remove linhas totalmente idênticas (duplicatas) e prepara a busca"""
    with open(caminho, encoding="utf-8") as f:
        linhas = list(csv.DictReader(f, delimiter=";"))
    vistos, base, duplicadas = set(), [], 0
    for r in linhas:
        chave = tuple(r.items())
        if chave in vistos:
            duplicadas += 1
            continue
        vistos.add(chave)
        r["_busca"] = normalizar(r.get("NOME", "")) + " " + (r.get("NUM_MAT", "") or "").strip()
        r["_curso"] = (r.get("CURSO", "") or "").strip() or "(sem curso)"
        r["_turma"] = (r.get("TURMA", "") or "").strip() or "(sem turma)"
        base.append(r)
    return base, duplicadas


# ----------------------------------------------------- regras do diploma ----
def historico_do_aluno(linha):
    """
    Linhas do quadro de componentes curriculares do diploma padrão.
    Cada item: dict com módulo, nome, nota/CH de teoria, CH de prática,
    nota/CH de estágio e indicadores de nota faltando.
    """
    itens = []
    for mod, nome, ch_t, ch_p, ch_e, col_t, col_e in CURRICULO:
        nota_t = parse_nota(linha.get(col_t, "")) if col_t else None
        nota_e = parse_nota(linha.get(col_e, "")) if col_e else None
        itens.append({
            "modulo": mod, "nome": nome,
            "ch_t": ch_t, "ch_p": ch_p, "ch_e": ch_e,
            "nota_t": fmt_nota(nota_t), "nota_e": fmt_nota(nota_e),
            "falta_t": nota_t is None,
            "falta_e": bool(ch_e) and nota_e is None,
        })
    return itens


# O diploma padrão imprime "Total de Cargas Horárias 1.880", mas 1.105 + 225 + 600
# somam 1.930. Mantemos o valor do modelo; altere aqui se o modelo for corrigido.
TOTAL_GERAL_CH = 1880


def totais_curriculo():
    t = sum(i[2] for i in CURRICULO)
    p = sum(i[3] for i in CURRICULO)
    e = sum(i[4] for i in CURRICULO)
    return t, p, e, TOTAL_GERAL_CH


def resumo_faltas(itens):
    """(notas encontradas, notas esperadas, lista com o que falta)"""
    esperadas, achadas, faltas = 0, 0, []
    for it in itens:
        esperadas += 1
        if it["falta_t"]:
            faltas.append(f"{it['nome']} (teoria)")
        else:
            achadas += 1
        if it["ch_e"]:
            esperadas += 1
            if it["falta_e"]:
                faltas.append(f"{it['nome']} (estágio)")
            else:
                achadas += 1
    return achadas, esperadas, faltas


def estagios_do_aluno(linha):
    """[(disciplina, local)] — as 5 do quadro padrão + extras preenchidos"""
    saida = [(rot, limpa(linha.get(c, ""))) for rot, c in ESTAGIO_DIPLOMA]
    for rot, c in ESTAGIO_EXTRAS:
        v = limpa(linha.get(c, ""))
        if v:
            saida.append((rot, v))
    return saida


def periodo_estagio(linha):
    """(início, fim) a partir de 'dd/mm/aa à dd/mm/aa'; datas 00/00/00 viram vazio"""
    datas = re.findall(r"\d{2}/\d{2}/\d{2,4}", linha.get(CAMPO_PERIODO, "") or "")
    datas = ["" if d.startswith("00/00") else d for d in datas]
    ini = datas[0] if len(datas) > 0 else ""
    fim = datas[1] if len(datas) > 1 else ""
    return ini, fim


def ano_conclusao(linha):
    a = (linha.get("ANOCONC", "") or "").strip()
    if a:
        return a
    m = re.search(r"(\d{4})\s*$", linha.get("DATACONC", "") or "")
    return m.group(1) if m else ""


def _naturalidade(linha):
    nat = limpa(linha.get("NATURALIDA", ""))
    nat = re.sub(r"[\s\-/,]+[A-Z]{2}$", "", nat)  # "Salvador BA" -> "Salvador"
    return nat


def _artigo_estado(estado):
    return "da" if normalizar(estado) in ("bahia", "paraiba") else "do"


def dados_pessoais(linha):
    return {
        "nome": limpa(linha.get("NOME", "")),
        "rg": limpa(linha.get("RG", "")),
        "pai": limpa(linha.get("PAI", "")),
        "mae": limpa(linha.get("MAE", "")),
        "naturalidade": _naturalidade(linha),
        "estado": limpa(linha.get("ESTADO", "")),
        "nascimento": limpa(linha.get("NASCIDO", "")),
        "matricula": limpa(linha.get("NUM_MAT", "")),
        "turma": limpa(linha.get("TURMA", "")),
        "turno": limpa(linha.get("Turno", "")),
        "curso": limpa(linha.get("CURSO", "")),
        "inicio": limpa(linha.get("DataInici", "")),
        "conclusao": limpa(linha.get("DATACONC", "")),
        "endereco": limpa(linha.get("EDNALUNO", "")),
    }


# =====================================================================
#  PDF
# =====================================================================
def _logos():
    for nome in ("brasao.jpg", "logo_escola.jpg"):
        for base in (BASE_DIR, Path(getattr(sys, "_MEIPASS", BASE_DIR))):
            p = base / nome
            if p.exists():
                yield nome, str(p)
                break


def gerar_pdf(linha, caminho, opcoes=None, incluir_frente=True):
    """
    Gera o PDF (A4 paisagem). Páginas: [diploma (frente)] + histórico (verso).
    opcoes: {"ano": "2025", "per_ini": "06/03/24", "per_fim": "30/11/2024"}
    """
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Paragraph

    opcoes = opcoes or {}
    d = dados_pessoais(linha)
    ano = (opcoes.get("ano") or ano_conclusao(linha)).strip()
    p_ini, p_fim = periodo_estagio(linha)
    p_ini = (opcoes.get("per_ini") if opcoes.get("per_ini") is not None else p_ini).strip()
    p_fim = (opcoes.get("per_fim") if opcoes.get("per_fim") is not None else p_fim).strip()
    logos = dict(_logos())

    W, H = landscape(A4)
    c = canvas.Canvas(str(caminho), pagesize=(W, H))
    c.setTitle(f"Diploma - {d['nome']}")
    c.setAuthor("Centro de Formação Técnica em Enfermagem Irmã Dulce")

    # ---- helpers (coordenadas com origem no topo da página) -------------
    def Y(y):
        return H - y

    def box(x, y, w, h, lw=0.7):
        c.setLineWidth(lw)
        c.rect(x, Y(y + h), w, h)

    def texto(x, y, s, fonte="Helvetica", tam=8, alinh="left"):
        c.setFont(fonte, tam)
        if alinh == "center":
            c.drawCentredString(x, Y(y), s)
        elif alinh == "right":
            c.drawRightString(x, Y(y), s)
        else:
            c.drawString(x, Y(y), s)

    def celula(x, y, w, h, s, fonte="Helvetica", tam=8, alinh="center", borda=True):
        if borda:
            box(x, y, w, h, 0.5)
        base_y = y + h / 2 + tam * 0.34
        if alinh == "center":
            texto(x + w / 2, base_y, s, fonte, tam, "center")
        else:
            texto(x + 3, base_y, s, fonte, tam)

    def paragrafo(html, x, y, w, estilo):
        p = Paragraph(html, estilo)
        _, ph = p.wrap(w, 1000)
        p.drawOn(c, x, Y(y) - ph)
        return ph

    # ================= PÁGINA 1 — DIPLOMA (frente) =====================
    if incluir_frente:
        if "brasao.jpg" in logos:
            c.drawImage(logos["brasao.jpg"], 112, Y(52 + 82), 82, 82, mask="auto")
        if "logo_escola.jpg" in logos:
            c.drawImage(logos["logo_escola.jpg"], 640, Y(52 + 78), 104, 78, mask="auto")
        cab = [
            ("República Federativa do Brasil", 8.5),
            ("Centro de Formação Técnica em Enfermagem Irmã Dulce", 8.5),
            ("Avenida Joana Angélica, nº 177, Nazaré, Salvador, Bahia  Tel: 3321-9366", 8.5),
            ("Parecer CEE 340/2023 – Resolução 227/2023 D.O. 16 de setembro de 2023", 8),
            ("COREN – Ba n. 0038 – RE / RT", 7),
        ]
        yy = 70
        for s, t in cab:
            texto(W / 2, yy, s, "Helvetica", t, "center")
            yy += 12
        texto(W / 2, 232, "Diploma", "Helvetica-Bold", 56, "center")

        est = ParagraphStyle("corpo", fontName="Helvetica", fontSize=12, leading=21, alignment=TA_CENTER)
        branco = "_" * 22
        nome = escape(d["nome"] or branco)
        rg = escape(d["rg"] or "_" * 14)
        pai, mae = escape(d["pai"]), escape(d["mae"])
        filiacao = " e ".join(x for x in (pai, mae) if x) or ("_" * 30)
        nat = escape(d["naturalidade"] or "_" * 14)
        uf = escape(d["estado"] or "_" * 10)
        nasc = escape(d["nascimento"] or "__/__/____")
        art = _artigo_estado(d["estado"])
        corpo1 = (
            "A diretora do Centro de Formação Técnica em Enfermagem Irmã Dulce de acordo com a Lei "
            "9.394/96 e Resolução CNE/CEB 4/99 e como dispositivo do Regimento Escolar confere o título "
            f"de Técnico em Enfermagem a <b>{nome}</b>, RG n <b>{rg}</b>, filho(a) <b>{filiacao}</b>, "
            f"natural de <b>{nat}</b>, estado {art} <b>{uf}</b>, nascido (a) em <b>{nasc}</b>, por "
            f"ter concluído o curso de Técnico em Enfermagem no ano letivo de <b>{escape(ano) or '______'}</b>."
        )
        ph = paragrafo(corpo1, 36, 262, W - 72, est)
        paragrafo("O presente Diploma outorga-lhe os direitos e prerrogativas estabelecidas nas Leis "
                  "vigentes do País.", 36, 262 + ph + 26, W - 72, est)
        texto(W / 2, 430, "Salvador, ................ de......................................."
                          "...........................................de...............................",
              "Helvetica", 11, "center")
        # assinaturas
        ya = 510
        for cx, linhas in (
            (205, [("Concluinte", 8)]),
            (438, [("Secretário", 8), ("Luciano Kleber Carneiro Ribeiro", 6.5),
                   ("Secretário Escolar", 6.5), ("Aut. NTE 26-508 Val. 12/2027", 6.5)]),
            (668, [("Diretora", 8), ("Ana Paula Costa Santos", 6.5),
                   ("Diretora Geral", 6.5), ("Aut. NTE 26-521 Val. 12/2027", 6.5)]),
        ):
            c.setLineWidth(0.5)
            c.line(cx - 95, Y(ya), cx + 95, Y(ya))
            yl = ya + 11
            for s, t in linhas:
                texto(cx, yl, s, "Helvetica", t, "center")
                yl += 9.5
        c.showPage()

    # ================= PÁGINA 2 — HISTÓRICO (verso) =====================
    box(14, 14, W - 28, H - 28, 0.9)

    # cabeçalho: título + objetivos
    box(34, 22, 158, 36)
    for i, s in enumerate(("Título: Técnico em Enfermagem", "Área Profissional: Saúde",
                           "Eixo Tecnológico Ambiente e Saúde")):
        texto(39, 33 + i * 10, s, "Helvetica", 6.5)
    box(198, 22, W - 22 - 198, 36)
    est_obj = ParagraphStyle("obj", fontName="Helvetica", fontSize=6.5, leading=8.2)
    paragrafo("- Identificar e avaliar os processos de Saúde X Doença, Dor X Sintomatologia – Direcionar a "
              "assistência de Enfermagem observando os sintomas e as histórias das Doenças.<br/>"
              "- Assistir ao portador de doenças degenerativas – Compreender e relacionar os fundamentos "
              "da Administração da Unidade de internação.", 203, 26, W - 22 - 198 - 10, est_obj)

    # faixa com os dados do aluno
    X0, LMOD = 34, 14
    TX = X0 + LMOD                       # início da tabela
    w_nome, w_nota, w_ch = 226, 40, 60
    TR = TX + w_nome + 3 * (w_nota + w_ch)   # fim da tabela
    box(X0, 62, TR - X0, 27)
    l1 = f"Aluno(a): {d['nome']}"
    texto(X0 + 4, 73, l1, "Helvetica-Bold", 7.8)
    partes = []
    if d["matricula"]:
        partes.append(f"Matrícula: {d['matricula']}")
    if d["rg"]:
        partes.append(f"RG: {d['rg']}")
    if d["nascimento"]:
        partes.append(f"Nascimento: {d['nascimento']}")
    texto(X0 + 4, 84, "     ".join(partes), "Helvetica", 7)
    extra = []
    if d["turma"]:
        extra.append(f"Turma: {d['turma']}")
    if ano:
        extra.append(f"Conclusão: {ano}")
    texto(TR - 4, 73, "   ".join(extra), "Helvetica", 7, "right")

    # cabeçalho da tabela
    yg = 94
    hg, hc = 13, 15
    for i, g in enumerate(("Teoria", "Prática", "Estágio")):
        gx = TX + w_nome + i * (w_nota + w_ch)
        celula(gx, yg, w_nota + w_ch, hg, g, "Helvetica-Bold", 7.8)
    celula(TX, yg + hg, w_nome, hc, "Componentes Curriculares", "Helvetica-Bold", 8, "left")
    for i in range(3):
        gx = TX + w_nome + i * (w_nota + w_ch)
        celula(gx, yg + hg, w_nota, hc, "Nota", "Helvetica-Bold", 7.2)
        celula(gx + w_nota, yg + hg, w_ch, hc, "Carga Horária", "Helvetica-Bold", 7.2)

    # linhas
    itens = historico_do_aluno(linha)
    y = yg + hg + hc
    RH = 16.8
    sem = "__"
    por_modulo = {}
    for it in itens:
        por_modulo.setdefault(it["modulo"], []).append(y)
        celula(TX, y, w_nome, RH, it["nome"], "Helvetica", 7.8, "left")
        xs = TX + w_nome
        # teoria
        celula(xs, y, w_nota, RH, it["nota_t"], "Helvetica", 8)
        celula(xs + w_nota, y, w_ch, RH, str(it["ch_t"]), "Helvetica", 8)
        # prática (o diploma padrão não traz nota de prática)
        xs += w_nota + w_ch
        celula(xs, y, w_nota, RH, sem, "Helvetica", 8)
        celula(xs + w_nota, y, w_ch, RH, str(it["ch_p"]) if it["ch_p"] else sem, "Helvetica", 8)
        # estágio
        xs += w_nota + w_ch
        celula(xs, y, w_nota, RH, it["nota_e"] if it["ch_e"] else sem, "Helvetica", 8)
        celula(xs + w_nota, y, w_ch, RH, str(it["ch_e"]) if it["ch_e"] else sem, "Helvetica", 8)
        y += RH
    # rótulos de módulo (vertical)
    for mod, ys in por_modulo.items():
        top, bottom = ys[0], ys[-1] + RH
        box(X0, top, LMOD, bottom - top, 0.7)
        c.saveState()
        c.translate(X0 + LMOD / 2 + 2.6, Y((top + bottom) / 2))
        c.rotate(90)
        c.setFont("Helvetica", 7)
        c.drawCentredString(0, 0, f"Módulo {mod:02d}")
        c.restoreState()

    # total de cargas horárias
    t, p, e, tot = totais_curriculo()
    celula(TX, y, w_nome, RH, "Total de Cargas Horárias", "Helvetica", 7.8, "left")
    xs = TX + w_nome
    for valor in (t, p, e):
        celula(xs, y, w_nota, RH, sem, "Helvetica", 8)
        celula(xs + w_nota, y, w_ch, RH, fmt_ch(valor), "Helvetica", 8)
        xs += w_nota + w_ch
    y += RH + 4
    celula(TX, y, 150, RH, f"Total de Cargas Horárias {fmt_ch(tot)}", "Helvetica", 7.8, "left")

    # coluna da direita: autenticação + estágios
    RX, RW = TR + 8, W - 22 - (TR + 8)
    box(RX, 62, RW, 150)
    texto(RX + RW / 2, 76, "Campo Reservado para Autentificação", "Helvetica-Bold", 8.5, "center")
    texto(RX + RW / 2, 88, "REGISTRO COM VALIDADE NACIONAL", "Helvetica-Bold", 7.5, "center")
    box(RX, 218, RW, 105)

    # quadro de campos de estágio
    estagios = estagios_do_aluno(linha)
    w1, w2 = 84, RW - 84
    sty = ParagraphStyle("est", fontName="Helvetica", fontSize=7.4, leading=8.8)
    sty_b = ParagraphStyle("estb", parent=sty, fontName="Helvetica-Bold")
    yt = 334
    celula(RX, yt, w1, 14, "Disciplinas", "Helvetica-Bold", 7.6, "left")
    celula(RX + w1, yt, w2, 14, "Campo de Estágios", "Helvetica-Bold", 7.6)
    yt += 14
    for rot, local in estagios:
        p1 = Paragraph(escape(rot), sty_b)
        p2 = Paragraph(escape(local.upper()), sty)
        h = max(p1.wrap(w1 - 6, 100)[1], p2.wrap(w2 - 6, 100)[1]) + 6
        box(RX, yt, w1, h, 0.5)
        box(RX + w1, yt, w2, h, 0.5)
        p1.drawOn(c, RX + 3, Y(yt + 3) - p1.height)
        p2.drawOn(c, RX + w1 + 3, Y(yt + 3) - p2.height)
        yt += h
    ini = p_ini or "__/__/__"
    fim = p_fim or "__/__/____"
    box(RX, yt, RW, 16, 0.5)
    texto(RX + RW / 2, yt + 11, f"No Período de: {ini} a {fim}", "Helvetica", 7.8, "center")

    c.showPage()
    c.save()
    return caminho
