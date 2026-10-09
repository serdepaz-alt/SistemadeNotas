#!/usr/bin/env python3
"""
Geração do diploma / histórico em DOCX a partir do modelo "DIPLOMA - PADRÃO.docx".

O modelo (um diploma de exemplo já preenchido) é usado como molde: este módulo
abre o arquivo, troca SOMENTE o texto dos campos variáveis e grava uma cópia.
Fontes, tamanhos, caixas de texto, tabelas, imagens e posições não são alterados,
então o resultado tem exatamente o layout do modelo.

O que é preenchido (as mesmas regras do PDF, em diploma_core.py):
  * frente : nome, RG, filiação, naturalidade, estado, nascimento, ano letivo;
  * verso  : nota final de teoria e de estágio de cada componente, campos de
             estágio (as disciplinas extras preenchidas ganham uma linha a mais)
             e o período "No Período de ...".

Atenção: as caixas "REGISTRO" e "Cod. Aut" do modelo trazem números do aluno de
exemplo (livro, folha, registro, código). Esses números são apagados (viram
linhas em branco) para nunca sair o registro de outra pessoa no diploma.
O registro é feito depois, à mão ou no Word.
"""
import copy
import re
import sys
import zipfile
from pathlib import Path

import diploma_core as dc

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V = "urn:schemas-microsoft-com:vml"
NS = {"w": W}


def _q(tag):
    return f"{{{W}}}{tag}"


# Algumas células de nota do modelo usam a fonte inexistente "CIDFont+F2" (resíduo da
# conversão de PDF para Word), por isso parecem diferentes das outras. Trocar a fonte
# deixa as linhas mais baixas e desalinha os rótulos "Módulo 0X" (caixas de altura fixa),
# então o padrão é manter o modelo exatamente como está.
UNIFORMIZAR_FONTE_NOTAS = False

# espaço que a caixa «Campo de Estágios» pode crescer antes de bater no fim da página;
# acima disso a fonte da caixa é reduzida (10 -> 8 pt) para nada ser cortado
LIMITE_CRESCIMENTO_CAIXA_PT = 48.0

# parágrafos vazios (espaçadores) retirados antes do verso; ver _estabilizar_paginacao
ESPACADORES_A_RETIRAR = 4

# ---- textos do aluno de exemplo que existem no modelo (âncoras) -------------
EX_NOME = "Maiara Amaral Neves da Silva"
EX_RG = "14.968.154-24"
EX_PAI = "Adilson Marques Neves"
EX_MAE = "Renilda do Espirito Santo Amaral"
EX_NAT = "Salvador"
EX_UF = "Bahia"
EX_NASC = "31/08/1992"
EX_ANO_ANTES = "no ano letivo de 20"      # seguido por dois runs: "2" e "5"


class ModeloNaoEncontrado(Exception):
    pass


class ModeloInvalido(Exception):
    """O DOCX não tem a estrutura esperada do modelo DIPLOMA - PADRÃO."""


# ---------------------------------------------------------------- modelo ----
def localizar_modelo():
    """procura o DOCX do modelo ao lado do programa (aceita variações do nome)"""
    pastas = [dc.BASE_DIR, Path(getattr(sys, "_MEIPASS", dc.BASE_DIR))]
    for pasta in pastas:
        achados = sorted(
            p for p in pasta.glob("*.docx")
            if not p.name.startswith("~$")
            and "diploma" in dc.normalizar(p.name) and "padr" in dc.normalizar(p.name))
        if achados:
            return achados[0]
    return None


# ------------------------------------------------------------ utilitários ---
def _txt_run(r):
    return "".join(t.text or "" for t in r.findall("w:t", NS))


def _set_run(r, texto):
    """troca o texto do run mantendo toda a formatação (rPr)"""
    for t in r.findall("w:t", NS):
        r.remove(t)
    if texto == "":
        return
    from lxml import etree
    t = etree.SubElement(r, _q("t"))
    t.text = texto
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _runs_texto(p):
    """runs do parágrafo que têm texto e não estão dentro de caixas de texto"""
    saida = []
    for r in p.iter(_q("r")):
        aninhado = False
        for a in r.iterancestors():          # só descarta caixas que estão DENTRO do parágrafo
            if a is p:
                break
            if a.tag == _q("txbxContent"):
                aninhado = True
                break
        if not aninhado and r.find("w:t", NS) is not None:
            saida.append(r)
    return saida


def _texto_par(p):
    return "".join(_txt_run(r) for r in _runs_texto(p))


def _trocar(runs, antigo, novo, a_partir=0):
    for i in range(a_partir, len(runs)):
        if _txt_run(runs[i]) == antigo:
            _set_run(runs[i], novo)
            return i
    raise ModeloInvalido(
        f"Não encontrei o trecho «{antigo}» no modelo. O modelo foi alterado? "
        "Ele precisa ser o DIPLOMA - PADRÃO original (com o aluno de exemplo).")


def _par_com(body, trecho):
    achados = [p for p in body.findall("w:p", NS) if trecho in _texto_par(p)]
    if len(achados) != 1:
        raise ModeloInvalido(f"Esperava 1 parágrafo com «{trecho}» no modelo e achei {len(achados)}.")
    return achados[0]


def _limpar_fonte_residual(r):
    """remove a fonte inexistente 'CIDFont+F2' (resíduo de conversão de PDF)"""
    rpr = r.find("w:rPr", NS)
    if rpr is None:
        return
    rf = rpr.find("w:rFonts", NS)
    if rf is not None and (rf.get(_q("ascii")) or "").startswith("CIDFont"):
        rpr.remove(rf)
    if len(rpr) == 0:
        r.remove(rpr)


def _celula_texto(cel, texto, uniformizar_fonte=False):
    """troca o texto da 1ª linha da célula; reaproveita o 1º run (formatação)"""
    from lxml import etree
    p = cel.find("w:p", NS)
    runs = [r for r in p.findall("w:r", NS)
            if r.find("w:pict", NS) is None and r.find("w:drawing", NS) is None]
    for r in runs[1:]:
        p.remove(r)
    if texto == "":                      # célula sem nota fica vazia, como no modelo
        if uniformizar_fonte:
            for r in runs:
                p.remove(r)
        elif runs:
            _set_run(runs[0], "")
        return
    if runs:
        r = runs[0]
    else:
        r = etree.SubElement(p, _q("r"))
    _set_run(r, texto)
    if uniformizar_fonte:
        _limpar_fonte_residual(r)


def _cel_texto_lido(cel):
    p = cel.find("w:p", NS)
    return "".join(_txt_run(r) for r in p.findall("w:r", NS)).strip() if p is not None else ""


def _shape_de(elem):
    for a in elem.iterancestors():
        if a.tag == f"{{{V}}}shape":
            return a
    return None


def _sem_cache_grafico(shape):
    """o atributo gfxdata guarda uma cópia antiga do conteúdo da caixa; sem ele o
    Word usa o texto atual (que é o que alteramos)"""
    if shape is not None and "gfxdata" in shape.attrib:
        del shape.attrib["gfxdata"]


# --------------------------------------------------- estimativa de altura ---
def _linhas(texto, largura_pt=124.0, tam=10.0):
    try:
        from reportlab.pdfbase.pdfmetrics import stringWidth
        medir = lambda s: stringWidth(s, "Helvetica", tam) * 0.80      # ~Calibri
    except Exception:
        medir = lambda s: len(s) * tam * 0.55
    linhas, atual = 1, 0.0
    for palavra in texto.split():
        if atual and atual + medir(palavra) > largura_pt:
            linhas += 1
            atual = 0.0
        atual += medir(palavra + " ")
    return linhas


def _altura_linha_tabela(n_linhas):
    return max(15.25, 3.6 + 11.8 * n_linhas)


# ===================================================================== edição
def _preencher_frente(body, d, ano):
    # nome e RG
    p = _par_com(body, "confere o título")
    runs = _runs_texto(p)
    i = _trocar(runs, EX_NOME, d["nome"] or "_" * 22)
    _trocar(runs, EX_RG, d["rg"] or "_" * 14, i)

    # filiação, naturalidade, estado e nascimento
    p = _par_com(body, "nascido (a) em")
    runs = _runs_texto(p)
    pai, mae = d["pai"], d["mae"]
    j = _trocar(runs, EX_PAI, pai or mae or "_" * 30)
    if _txt_run(runs[j + 1]) != " e " or _txt_run(runs[j + 2]) != EX_MAE:
        raise ModeloInvalido("A frase da filiação do modelo não tem o formato esperado.")
    if pai and mae:
        _set_run(runs[j + 2], mae)
    else:                                  # só um (ou nenhum) responsável: sem o " e "
        _set_run(runs[j + 1], "")
        _set_run(runs[j + 2], "")
    j += 2
    j = _trocar(runs, EX_NAT, d["naturalidade"] or "_" * 14, j)
    k = next((n for n in range(j, len(runs)) if _txt_run(runs[n]) == ", estado "), None)
    if k is None or _txt_run(runs[k + 1]) != "d" or _txt_run(runs[k + 2]) != "a":
        raise ModeloInvalido("A frase «estado da ...» do modelo não tem o formato esperado.")
    _set_run(runs[k + 2], dc._artigo_estado(d["estado"])[1])        # da / do
    _trocar(runs, EX_UF, d["estado"] or "_" * 10, k + 2)
    _trocar(runs, EX_NASC, d["nascimento"] or "__/__/____")

    # ano letivo
    p = _par_com(body, "no ano letivo de")
    runs = _runs_texto(p)
    n = next((x for x in range(len(runs)) if _txt_run(runs[x]) == EX_ANO_ANTES), None)
    if n is None or _txt_run(runs[n + 1]) != "2" or _txt_run(runs[n + 2]) != "5":
        raise ModeloInvalido("A frase do ano letivo do modelo não tem o formato esperado.")
    _set_run(runs[n], "no ano letivo de " + (ano or "______"))
    _set_run(runs[n + 1], "")
    _set_run(runs[n + 2], "")


def _preencher_notas(body, itens, uniformizar_fonte):
    tbl = body.find("w:tbl", NS)
    linhas = tbl.findall("w:tr", NS) if tbl is not None else []
    if len(linhas) != len(dc.CURRICULO) + 2 or \
            _cel_texto_lido(linhas[0].find("w:tc", NS)) != "Componentes Curriculares":
        raise ModeloInvalido("A tabela de notas do modelo não tem as 26 linhas esperadas.")
    for i, it in enumerate(itens, start=1):
        cels = linhas[i].findall("w:tc", NS)
        esperado = (str(it["ch_t"]), str(it["ch_p"]) if it["ch_p"] else "__",
                    str(it["ch_e"]) if it["ch_e"] else "__")
        lido = (_cel_texto_lido(cels[2]).replace(".", ""), _cel_texto_lido(cels[4]),
                _cel_texto_lido(cels[6]) if i != 18 else esperado[2])
        if lido != esperado:
            raise ModeloInvalido(
                f"A linha «{it['nome']}» do modelo tem cargas horárias {lido}, mas o "
                f"CURRICULO do programa diz {esperado}. Ajuste um dos dois.")
        _celula_texto(cels[1], it["nota_t"], uniformizar_fonte)
        if it["ch_e"]:                      # só há nota de estágio onde há carga horária
            _celula_texto(cels[5], it["nota_e"], uniformizar_fonte)


def _altura_tabela(linhas, tam):
    total = 0.0
    for lin in linhas[1:]:
        cels = lin.findall("w:tc", NS)
        total += _altura_linha_tabela(_linhas(_cel_texto_lido(cels[1]), tam=tam,
                                              largura_pt=124.0))
    return total


def _fonte_run(r, meio_pt):
    from lxml import etree
    rpr = r.find("w:rPr", NS)
    if rpr is None:
        rpr = etree.Element(_q("rPr"))
        r.insert(0, rpr)
    for nome in ("sz", "szCs"):
        e = rpr.find(f"w:{nome}", NS)
        if e is None:
            e = etree.SubElement(rpr, _q(nome))
        e.set(_q("val"), str(meio_pt))


def _preencher_estagios(root, estagios, per_ini, per_fim):
    from lxml import etree
    tabela = None
    for t in root.iter(_q("tbl")):
        primeira = t.find("w:tr/w:tc", NS)
        if primeira is not None and _cel_texto_lido(primeira) == "Disciplinas" \
                and any(a.tag == _q("txbxContent") for a in t.iterancestors()):
            tabela = t
            break
    if tabela is None:
        raise ModeloInvalido("Não encontrei a tabela «Disciplinas | Campo de Estágios» no modelo.")
    linhas = tabela.findall("w:tr", NS)
    padrao = dc.ESTAGIO_DIPLOMA
    if len(linhas) != len(padrao) + 1:
        raise ModeloInvalido("A tabela de campos de estágio do modelo não tem as 5 linhas esperadas.")

    altura_antes = _altura_tabela(linhas, 10.0)       # como o modelo está agora
    modelo_extra = linhas[-1]

    for lin, (rot, local) in zip(linhas[1:], estagios[:len(padrao)]):
        _celula_texto(lin.findall("w:tc", NS)[1], local.upper())
    for rot, local in estagios[len(padrao):]:           # disciplinas extras preenchidas
        nova = copy.deepcopy(modelo_extra)
        cels = nova.findall("w:tc", NS)
        _celula_texto(cels[0], rot)
        _celula_texto(cels[1], local.upper())
        tabela.append(nova)
    linhas = tabela.findall("w:tr", NS)

    # a caixa de texto tem altura fixa: cresce o que as linhas novas / quebras exigirem
    altura_depois = _altura_tabela(linhas, 10.0)
    reduzir = (altura_depois - altura_antes) > LIMITE_CRESCIMENTO_CAIXA_PT
    if reduzir:                                         # caso extremo: fonte menor na caixa
        for r in tabela.iter(_q("r")):
            _fonte_run(r, 16)
        altura_depois = _altura_tabela(linhas, 8.0)
    shape = _shape_de(tabela)
    _sem_cache_grafico(shape)
    delta = altura_depois - altura_antes
    if shape is not None and delta > 0.5:
        estilo = shape.get("style", "")
        m = re.search(r"(?<![-\w])height:([\d.]+)pt", estilo)
        if m:
            nova_altura = float(m.group(1)) + delta
            shape.set("style", estilo[:m.start(1)] + f"{nova_altura:.2f}" + estilo[m.end(1):])

    # linha do período
    caixa = next(a for a in tabela.iterancestors() if a.tag == _q("txbxContent"))
    alvo = [p for p in caixa.findall("w:p", NS) if _texto_par(p).startswith("No Período de")]
    if len(alvo) != 1:
        raise ModeloInvalido("Não encontrei a linha «No Período de:» no modelo.")
    runs = _runs_texto(alvo[0])
    _set_run(runs[1], f"{per_ini or '__/__/__'} à {per_fim or '__/__/____'}")
    for r in runs[2:]:
        r.getparent().remove(r)
    if reduzir:
        for r in runs[:2]:
            _fonte_run(r, 16)


def _estabilizar_paginacao(body):
    """No modelo, o verso só começa na página 2 porque a frente "enche" a página 1 com
    parágrafos vazios; assim a posição de TODO o verso muda conforme o nº de linhas do
    texto da frente (que varia com o tamanho dos nomes). Aqui a quebra de página passa a
    ser explícita (é onde o Word já iniciava o verso) e os espaçadores vazios finais são
    retirados, o que dá folga para frentes mais longas sem criar página em branco."""
    from lxml import etree
    verso = None
    for p in body.findall("w:p", NS):
        for t in p.iter(_q("t")):
            if (t.text or "").startswith("Título: Técnico em Enfermagem"):
                verso = p
                break
        if verso is not None:
            break
    if verso is None:
        raise ModeloInvalido("Não encontrei o início do verso (caixa «Título: Técnico em Enfermagem»).")

    # espaçadores: parágrafos vazios, sem objetos, logo antes do verso
    vazios = []
    ant = verso.getprevious()
    while ant is not None and ant.tag == _q("p") and len(vazios) < ESPACADORES_A_RETIRAR:
        vazio = (not any(True for _ in ant.iter(_q("t")))
                 and ant.find(".//w:pict", NS) is None and ant.find(".//w:drawing", NS) is None
                 and ant.find(".//w:sectPr", NS) is None)
        if not vazio:
            break
        vazios.append(ant)
        ant = ant.getprevious()
    for v in vazios:
        body.remove(v)

    ppr = verso.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(_q("pPr"))
        verso.insert(0, ppr)
    if ppr.find("w:pageBreakBefore", NS) is None:
        pos = 0
        for i, filho in enumerate(ppr):
            if filho.tag in (_q("pStyle"), _q("keepNext"), _q("keepLines")):
                pos = i + 1
        ppr.insert(pos, etree.Element(_q("pageBreakBefore")))


def _limpar_registro(root):
    """apaga os números do aluno de exemplo nas caixas REGISTRO / Cod. Aut"""
    def em_branco(m):
        return "_" * len(m.group())
    for p in root.iter(_q("p")):
        if not any(a.tag == _q("txbxContent") for a in p.iterancestors()):
            continue
        t = "".join(x.text or "" for x in p.iter(_q("t"))).strip()
        if re.match(r"^N[oº°]?\s*do Registro", t) or t.startswith("Folha") or t.startswith("Cod"):
            for x in p.iter(_q("t")):
                if x.text:
                    x.text = re.sub(r"\d+[A-Z]*", em_branco, x.text)
            _sem_cache_grafico(_shape_de(p))


# ====================================================================== API ==
def gerar_docx(linha, destino, opcoes=None, modelo=None, uniformizar_fonte=UNIFORMIZAR_FONTE_NOTAS):
    """
    Gera o DOCX do diploma (frente + verso) do aluno `linha` em `destino`.
    opcoes: {"ano": "2025", "per_ini": "06/03/24", "per_fim": "30/11/2024"}
            (os mesmos campos "Dados para impressão" usados no PDF)
    uniformizar_fonte: troca a fonte inexistente 'CIDFont+F2' (resíduo de conversão de
            PDF presente em algumas células de nota do modelo) pela fonte das demais.
            Padrão False: mantém a formatação de cada célula como no modelo, o que
            preserva a altura das linhas (e o alinhamento dos rótulos "Módulo 0X").
    """
    from lxml import etree

    modelo = Path(modelo) if modelo else localizar_modelo()
    if not modelo or not Path(modelo).exists():
        raise ModeloNaoEncontrado(
            "Não encontrei o modelo «DIPLOMA - PADRÃO.docx». Coloque-o na mesma pasta do programa.")

    if not dc.is_tecnico(linha):
        raise ValueError("O diploma padrão só vale para o curso Técnico em Enfermagem.")
    opcoes = opcoes or {}
    d = dc.dados_pessoais(linha)
    ano = (opcoes.get("ano") or dc.ano_conclusao(linha)).strip()
    per_ini, per_fim = dc.periodo_estagio(linha)
    if opcoes.get("per_ini") is not None:
        per_ini = opcoes["per_ini"].strip()
    if opcoes.get("per_fim") is not None:
        per_fim = opcoes["per_fim"].strip()

    with zipfile.ZipFile(modelo) as zin:
        infos = zin.infolist()
        dados = {i.filename: zin.read(i.filename) for i in infos}
    if "word/document.xml" not in dados:
        raise ModeloInvalido("O arquivo não é um documento do Word válido.")

    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = etree.fromstring(dados["word/document.xml"], parser)
    body = root.find("w:body", NS)

    _preencher_frente(body, d, ano)
    _preencher_notas(body, dc.historico_do_aluno(linha), uniformizar_fonte)
    _preencher_estagios(root, dc.estagios_do_aluno(linha), per_ini, per_fim)
    _limpar_registro(root)
    _estabilizar_paginacao(body)

    dados["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)

    destino = Path(destino)
    tmp = destino.with_name(destino.name + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in infos:                       # mesma ordem do original
                novo = zipfile.ZipInfo(info.filename, info.date_time)
                novo.compress_type = zipfile.ZIP_DEFLATED if info.compress_type else zipfile.ZIP_STORED
                novo.external_attr = info.external_attr
                zout.writestr(novo, dados[info.filename])
        tmp.replace(destino)
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise
    return destino
