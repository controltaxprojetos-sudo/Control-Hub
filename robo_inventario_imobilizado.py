#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robô de inventário do ativo imobilizado - Grupo Qualitá  (100% local: sem API, sem tokens)

Lê os PDFs de notas fiscais (DANFE de NF-e modelo 55) da pasta e subpastas, extrai o texto com
as posições na página (PyMuPDF) e, para PDFs digitalizados, usa OCR (Tesseract, opcional).
Gera o Excel Inventario_Imobilizado_Grupo_Qualita_DD-MM-AA_HHhMM.xlsx.

Instalação (Prompt de Comando):
    py -m pip install pymupdf openpyxl rapidocr-onnxruntime pillow numpy tzdata
Uso:
    py robo_inventario_imobilizado.py            (ele pergunta a pasta)
    py robo_inventario_imobilizado.py --pasta "\\\\servidor\\pasta" --saida "C:\\Saida"
"""
from __future__ import annotations
import argparse
import io
import re
import sys
import unicodedata
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import os
import traceback
from datetime import timedelta, timezone


def _pausar():
    try:
        input("\nPressione Enter para fechar...")
    except Exception:
        pass


try:
    import fitz  # PyMuPDF
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as _e:
    print(f"Falta instalar um componente: {_e}")
    print("Abra o Prompt de Comando e rode:")
    print("    py -m pip install pymupdf openpyxl rapidocr-onnxruntime pillow numpy tzdata")
    _pausar()
    sys.exit(1)

# --------------------------------------------------------------------------- #
# CONFIGURAÇÃO
# --------------------------------------------------------------------------- #
PASTA_PADRAO = r"\\público\itinga"
TOLERANCIA = Decimal("0.05")

# Classificação do destinatário. A raiz do CNPJ (8 primeiros dígitos) tem prioridade.
# Preencha com as raízes reais; "Outra do grupo" serve para empresas sem nome-chave.
CNPJ_RAIZ_GRUPO = {
    # "12345678": "Itinga",
    # "87654321": "Outra do grupo",
}
# Fallback por palavra-chave na razão social (sem acento, maiúsculas).
NOME_GRUPO = {
    "Itinga": ["ITINGA"],
    "Splendour": ["SPLENDOUR"],
    "Granulito": ["GRANULITO"],
    "Qualitá": ["QUALITA"],
}
EMPRESAS_RESUMO = ["Itinga", "Splendour", "Granulito", "Qualitá", "Outra do grupo", "Fora do grupo"]

TIPOS = ["DANFE NF-e (55)", "NFC-e (65)", "NF modelo 1/1A", "NFS-e", "Outro"]
TIPOS_COM_CHAVE = {"DANFE NF-e (55)", "NFC-e (65)"}
MODELO_POR_TIPO = {"DANFE NF-e (55)": "55", "NFC-e (65)": "65"}

# Tipos de pendência
P_ILEGIVEL = "Arquivo ilegível / sem nota"
P_NAO_PROC = "Arquivo não processado"
P_CHAVE_INV = "Chave inválida"
P_CHAVE_INC = "Chave incoerente com a nota"
P_DUP = "Nota duplicada"
P_TOTAL = "Divergência de total"
P_FORA = "Destinatário fora do grupo"
P_SEM_ID = "Bem sem identificação"
P_ID_PADRAO = "Identificador fora do padrão"
P_NAO_SUP = "Documento não suportado (lançar manualmente)"
P_LEIT = "Item com leitura incompleta"
ORDEM_PEND = [P_ILEGIVEL, P_NAO_PROC, P_NAO_SUP, P_LEIT, P_CHAVE_INV, P_CHAVE_INC, P_DUP, P_TOTAL, P_FORA, P_SEM_ID, P_ID_PADRAO]

# --------------------------------------------------------------------------- #
# COLUNAS DO INVENTÁRIO (ordem e nomes conforme a especificação)
# --------------------------------------------------------------------------- #
# (cabeçalho, tipo, largura)
COLUNAS = [
    ("Arquivo de origem", "text", 50),
    ("Tipo de documento", "text", 18),
    ("Chave de acesso", "text", 48),
    ("Validação da chave", "text", 16),
    ("Número da NF", "int", 12),
    ("Série", "int", 8),
    ("Data de emissão", "date", 12),
    ("Emitente: razão social", "text", 36),
    ("Emitente: CNPJ", "text", 20),
    ("Destinatário: razão social", "text", 36),
    ("Destinatário: CNPJ", "text", 20),
    ("Empresa do grupo", "text", 16),
    ("Item", "int", 6),
    ("Descrição completa do produto/equipamento", "text", 70),
    ("NCM", "text", 12),
    ("CFOP", "text", 8),
    ("CST ou CSOSN (ICMS)", "text", 10),
    ("Quantidade", "qty", 11),
    ("Unidade", "text", 9),
    ("Valor unitário", "money", 15),
    ("Valor total do item", "money", 15),
    ("Desconto rateado", "money", 14),
    ("IPI do item", "money", 14),
    ("Frete rateado", "money", 14),
    ("Seguro rateado", "money", 14),
    ("ICMS-ST do item (ou rateado)", "money", 15),
    ("Outras despesas acessórias rateadas", "money", 15),
    ("Valor unitário de custo", "money", 16),
    ("Valor total de custo do item", "money", 16),
    ("Base de cálculo do ICMS", "money", 15),
    ("Alíquota de ICMS", "pct", 10),
    ("ICMS destacado", "money", 14),
    ("Alíquota de IPI", "pct", 10),
    ("Número de série", "text", 22),
    ("Chassi", "text", 20),
    ("RENAVAM", "text", 14),
    ("Placa", "text", 10),
    ("Valor total da NF", "money", 15),
    ("Observações / pendências", "text", 60),
]
IDX = {nome: i for i, (nome, _, _) in enumerate(COLUNAS)}
# Letras das colunas usadas em fórmulas
L = {k: get_column_letter(i + 1) for k, i in {
    "chave": 2, "numero": 4, "serie": 5, "emit_cnpj": 8, "empresa": 11, "ncm": 14,
    "qtd": 17, "vtotal": 20, "desc": 21, "ipi": 22, "frete": 23, "seg": 24, "st": 25,
    "outras": 26, "custo_un": 27, "custo_tot": 28, "bc": 29, "icms": 31,
}.items()}

FMT_MONEY = '"R$" #,##0.00'
FMT_PCT = "0.00%"
FMT_DATE = "DD/MM/YYYY"
ARIAL = "Arial"

# --------------------------------------------------------------------------- #
# UTILITÁRIOS
# --------------------------------------------------------------------------- #
ZERO = Decimal("0")
CENT = Decimal("0.01")


def D(x):
    if x is None or x == "":
        return None
    try:
        return Decimal(str(x))
    except Exception:
        return None


def q2(d):
    return d.quantize(CENT, rounding=ROUND_HALF_UP)


def allocate(total: Decimal, pesos):
    """Reparte 'total' proporcionalmente aos pesos, em centavos; o resto vai ao último peso > 0."""
    soma = sum(pesos)
    if soma == 0:
        pesos = [Decimal(1)] * len(pesos)
        soma = Decimal(len(pesos))
    partes = [q2(total * p / soma) for p in pesos]
    resto = total - sum(partes)
    if resto != 0:
        alvo = max(i for i, p in enumerate(pesos) if p > 0)
        partes[alvo] += resto
    return partes


def so_digitos(s):
    return re.sub(r"\D", "", str(s or ""))


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or ""))
    return "".join(c for c in s if not unicodedata.combining(c)).upper()


def txt(s):
    s = str(s).strip() if s is not None else ""
    return s or None


def num_ou_texto(x):
    s = txt(x)
    if s is None:
        return None
    return int(s) if s.isdigit() else s


def fmt_cnpj(d):
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}" if len(d) == 14 else (d or None)


def cnpj_valido(c):
    if len(c) != 14 or c == c[0] * 14:
        return False

    def dv(base):
        pesos = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2] if len(base) == 12 else [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        r = sum(int(d) * p for d, p in zip(base, pesos)) % 11
        return "0" if r < 2 else str(11 - r)
    d1 = dv(c[:12])
    return c[12:] == d1 + dv(c[:12] + d1)


def chave_dv_ok(ch):
    s, peso = 0, 2
    for d in reversed(ch[:43]):
        s += int(d) * peso
        peso = 2 if peso == 9 else peso + 1
    r = s % 11
    return (0 if r in (0, 1) else 11 - r) == int(ch[43])


def classificar(cnpj, razao):
    if cnpj[:8] in CNPJ_RAIZ_GRUPO and len(cnpj) == 14:
        return CNPJ_RAIZ_GRUPO[cnpj[:8]]
    nome = norm(razao)
    for emp, kws in NOME_GRUPO.items():
        if any(k in nome for k in kws):
            return emp
    return "Fora do grupo"


# --------------------------------------------------------------------------- #
# EXTRAÇÃO LOCAL (sem API, sem tokens): texto do PDF com coordenadas + OCR opcional
# --------------------------------------------------------------------------- #
RE_BR = re.compile(r"^(?:\d{1,3}(?:\.\d{3})+|\d+),\d+$")
RE_DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
RE_CNPJ_FMT = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
RE_CNPJ_CPF_FMT = re.compile(r"^(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})$")

# Marcadores de fim de tabela de produtos (começo da linha, sem acento e sem pontuação)
SEC_FIM = ("CALCULODOISSQN", "DADOSADICIONAIS", "CALCULODOIMPOSTO", "TRANSPORTADOR",
           "FATURA", "INFORMACOESCOMPLEMENTARES", "RESERVADOAOFISCO")

# Rótulos da área "CÁLCULO DO IMPOSTO" (tokens normalizados; "*" = prefixo)
CAMPOS_TOTAIS = {
    "valor_produtos": [["VALOR", "TOTAL", "DOS", "PRODUTOS"], ["V", "TOTAL", "PRODUTOS"], ["TOTAL", "DOS", "PRODUTOS"]],
    "frete": [["VALOR", "DO", "FRETE"], ["V", "FRETE"]],
    "seguro": [["VALOR", "DO", "SEGURO"], ["V", "SEGURO"]],
    "desconto": [["DESCONTO"]],
    "outras_despesas": [["OUTRAS", "DESPESAS*"]],
    "ipi": [["VALOR", "TOTAL", "DO", "IPI"], ["V", "TOTAL", "IPI"], ["VALOR", "DO", "IPI"]],
    "icms_st": [["VALOR", "DO", "ICMS", "SUBST*"], ["V", "ICMS", "SUBST*"]],
    "valor_nf": [["VALOR", "TOTAL", "DA", "NOTA"], ["V", "TOTAL", "DA", "NOTA"], ["TOTAL", "DA", "NOTA"]],
}

RE_SERIE = re.compile(
    r"\b(?:N[ºO°]?\.?\s*DE\s*S[ÉE]RIE|N[ºO°]\.?\s*S[ÉE]RIE|S[ÉE]RIE|N/S|NS)\b\s*[:\-\.]?\s*"
    r"([A-Z0-9][A-Z0-9\-\./]{1,24}(?:\s*[,;]\s*[A-Z0-9][A-Z0-9\-\./]{1,24})*)", re.I)
RE_CHASSI = re.compile(r"\bCHASSI?S?\b\.?\s*[:\-]?\s*([A-Z0-9]{8,20}(?:\s*[,;]\s*[A-Z0-9]{8,20})*)", re.I)
RE_RENAVAM = re.compile(r"\bRENAVAM\b\s*[:\-]?\s*(\d{8,11}(?:\s*[,;]\s*\d{8,11})*)", re.I)
RE_PLACA = re.compile(r"\bPLACAS?\b\s*[:\-]?\s*([A-Z]{3}-?\d[A-Z0-9]\d{2}(?:\s*[,;]\s*[A-Z]{3}-?\d[A-Z0-9]\d{2})*)", re.I)


def nrm(t):
    s = unicodedata.normalize("NFKD", str(t or "")).upper()
    return re.sub(r"[^A-Z0-9]", "", s)


def br(t):
    return Decimal(t.replace(".", "").replace(",", "."))


class Pal:
    __slots__ = ("x0", "y0", "x1", "y1", "t", "n")

    def __init__(self, x0, y0, x1, y1, t):
        self.x0, self.y0, self.x1, self.y1, self.t = x0, y0, x1, y1, t
        self.n = nrm(t)


def ocr_disponivel():
    try:
        import pytesseract
        for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
            if Path(c).exists():
                pytesseract.pytesseract.tesseract_cmd = c
                break
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def ocr_pagina(page):
    import pytesseract
    from PIL import Image
    zoom = 300 / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    try:
        d = pytesseract.image_to_data(img, lang="por", output_type=pytesseract.Output.DICT)
    except Exception:
        d = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    out = []
    for i, t in enumerate(d["text"]):
        t = t.strip()
        if t and float(d["conf"][i]) >= 0:
            x, y, w, h = d["left"][i], d["top"][i], d["width"][i], d["height"][i]
            out.append(Pal(x / zoom, y / zoom, (x + w) / zoom, (y + h) / zoom, t))
    return out


def palavras_pagina(page, ocr_ok):
    """Devolve (palavras, usou_ocr). usou_ocr=None quando não há texto e não há OCR."""
    ws = [Pal(*w[:5]) for w in page.get_text("words")]
    if sum(len(w.t) for w in ws) >= 60:
        return ws, False
    if ocr_ok:
        return ocr_pagina(page), True
    return [], None


def agrupar_linhas(ws, tol=3.0):
    ws = sorted(ws, key=lambda w: ((w.y0 + w.y1) / 2, w.x0))
    linhas, ys = [], []
    for w in ws:
        yc = (w.y0 + w.y1) / 2
        if linhas and abs(yc - ys[-1]) <= tol:
            linhas[-1].append(w)
        else:
            linhas.append([w]); ys.append(yc)
    for l in linhas:
        l.sort(key=lambda w: w.x0)
    return linhas


def achar_seq(linhas, seq, ini=0):
    """Acha a sequência de tokens normalizados numa linha. Devolve dict ou None."""
    for li in range(ini, len(linhas)):
        ws = [w for w in linhas[li] if w.n]
        ns = [w.n for w in ws]
        for i in range(len(ns) - len(seq) + 1):
            if all(ns[i + k] == s or (s.endswith("*") and ns[i + k].startswith(s[:-1]))
                   for k, s in enumerate(seq)):
                nxt = ws[i + len(seq)] if i + len(seq) < len(ws) else None
                return {"li": li, "w0": ws[i], "w1": ws[i + len(seq) - 1], "nxt": nxt}
    return None


def valor_abaixo(linhas, a, filtro=lambda w: True, maxdy=30):
    """Palavras logo abaixo do rótulo, dentro da caixa dele (até o próximo rótulo da mesma linha)."""
    x0 = a["w0"].x0 - 3
    xmax = a["nxt"].x0 - 1 if a["nxt"] else 1e9
    ylab = a["w0"].y1
    for lj in range(a["li"] + 1, len(linhas)):
        if min(w.y0 for w in linhas[lj]) - ylab > maxdy:
            break
        cand = [w for w in linhas[lj] if w.x0 >= x0 and w.x0 < xmax and filtro(w)]
        if cand:
            return cand
    return []


def achar_chave(linhas):
    for ws in linhas:
        run = []
        for w in ws + [None]:
            if w is not None and re.fullmatch(r"\d{4}", w.t):
                run.append(w.t)
                continue
            if len(run) == 11:
                return "".join(run)
            run = []
            if w is not None and re.fullmatch(r"\d{44}", w.t):
                return w.t
    return None


def extrair_totais(paginas):
    for linhas in paginas:
        m = achar_seq(linhas, ["CALCULO", "DO", "IMPOSTO"])
        if not m:
            continue
        out = {}
        for nome, seqs in CAMPOS_TOTAIS.items():
            out[nome] = None
            for seq in seqs:
                a = achar_seq(linhas, seq, m["li"])
                if a:
                    ws = valor_abaixo(linhas, a, lambda w: RE_BR.match(w.t))
                    if ws:
                        out[nome] = br(ws[0].t)
                    break
        if out.get("valor_nf") is not None or out.get("valor_produtos") is not None:
            return out
    return {}


def _item_da_linha(ws):
    """Reconhece linha de item: ... NCM(8) CST CFOP UN <9 números no formato BR>."""
    for p in range(len(ws) - 2):
        if (re.fullmatch(r"\d{8}", ws[p].t) and re.fullmatch(r"\d{2,4}", ws[p + 1].t)
                and re.fullmatch(r"\d{4}", ws[p + 2].t)):
            j = p + 3
            un = None
            if j < len(ws) and not RE_BR.match(ws[j].t):
                un = ws[j].t
                j += 1
            nums = [br(w.t) for w in ws[j:] if RE_BR.match(w.t)]
            antes = ws[:p]
            if len(antes) >= 2:
                codigo, desc_w = antes[0].t, antes[1:]
            else:
                codigo, desc_w = None, antes
            it = {"codigo": codigo, "ncm": ws[p].t, "cst_csosn": ws[p + 1].t, "cfop": ws[p + 2].t,
                  "unidade": un, "_desc": [" ".join(w.t for w in desc_w)] if desc_w else [],
                  "_dx0": desc_w[0].x0 if desc_w else (ws[0].x1 + 2), "_dx1": ws[p].x0,
                  "leitura_incompleta": False, "observacoes": None}
            if len(nums) >= 9:
                (it["quantidade"], it["valor_unitario"], it["valor_total"], it["desconto"], it["bc_icms"],
                 it["valor_icms"], it["ipi"], it["aliq_icms"], it["aliq_ipi"]) = nums[:9]
            else:
                it["leitura_incompleta"] = True
                it["observacoes"] = f"Colunas numéricas lidas: {len(nums)} de 9; conferir com o PDF"
                for k, c in enumerate(("quantidade", "valor_unitario", "valor_total")):
                    it[c] = nums[k] if k < len(nums) else None
            return it
    return None


def extrair_itens(paginas):
    itens = []
    for linhas in paginas:
        atual = None
        for ws in linhas:
            ln = "".join(w.n for w in ws)
            ns = [w.n for w in ws]
            if ln.startswith(SEC_FIM) or (any(n.startswith("NCM") for n in ns) and "CFOP" in ns):
                atual = None
                continue
            it = _item_da_linha(ws)
            if it:
                itens.append(it); atual = it
            elif atual:
                cont = [w.t for w in ws if w.x0 >= atual["_dx0"] - 3 and w.x1 <= atual["_dx1"] + 1]
                if cont:
                    atual["_desc"].append(" ".join(cont))
    return itens


def _lista(rx, texto):
    out = []
    for m in rx.finditer(texto):
        for v in re.split(r"\s*[,;]\s*", m.group(1).strip()):
            v = v.strip().upper()
            if v and v not in out:
                out.append(v)
    return out


def identificadores(texto, qtd):
    ser = [s for s in _lista(RE_SERIE, texto) if any(c.isdigit() for c in s)]
    cha, ren, pla = _lista(RE_CHASSI, texto), _lista(RE_RENAVAM, texto), _lista(RE_PLACA, texto)
    listas = {"numero_serie": ser, "chassi": cha, "renavam": ren, "placa": pla}
    if not any(listas.values()):
        return []
    n = int(qtd) if qtd is not None and qtd == qtd.to_integral_value() and qtd > 1 else 1
    if n > 1 and any(len(l) == n for l in listas.values()):
        return [{k: (l[i] if len(l) == n else None) for k, l in listas.items()} for i in range(n)]
    m = max(len(l) for l in listas.values())
    return [{k: (l[i] if i < len(l) else None) for k, l in listas.items()} for i in range(m)]


def texto_complementar(paginas):
    partes = []
    for linhas in paginas:
        a = achar_seq(linhas, ["INFORMACOES", "COMPLEMENTARES"])
        if not a:
            continue
        b = achar_seq(linhas, ["RESERVADO", "AO", "FISCO"], a["li"])
        xmax = b["w0"].x0 - 1 if b and b["li"] == a["li"] else 1e9
        for ws in linhas[a["li"] + 1:]:
            t = " ".join(w.t for w in ws if w.x0 < xmax)
            if t:
                partes.append(t)
    return " ".join(partes)


def montar_nota(grupo):
    pag = grupo["paginas"]
    p1 = pag[0]
    chave = grupo["chave"]
    obs = []
    tipo = "Outro"
    if chave:
        tipo = {"55": "DANFE NF-e (55)", "65": "NFC-e (65)"}.get(chave[20:22], "Outro")
    nota = {"tipo_documento": tipo, "chave_acesso": chave, "legivel": True}
    if tipo != "DANFE NF-e (55)":
        nota["nao_suportado"] = True
        nota["observacoes"] = ("Documento não reconhecido como DANFE de NF-e modelo 55 (NFC-e, NFS-e ou outro "
                               "layout); lançar manualmente")
        return nota

    # cabeçalho
    topo = " ".join(w.t for ws in p1[:40] for w in ws)
    m = re.search(r"RECEBEMOS DE\s+(.+?)\s+OS PRODUTOS", topo, re.I | re.S)
    nota["emitente_razao"] = m.group(1).strip() if m else None
    dest = achar_seq(p1, ["DESTINATARIO*"])
    ini_dest = dest["li"] if dest else len(p1)
    antes = " ".join(w.t for ws in p1[:ini_dest] for w in ws)
    m = RE_CNPJ_FMT.search(antes)
    if m:
        nota["emitente_cnpj"] = m.group(0)
    else:
        nota["emitente_cnpj"] = chave[6:20]
        obs.append("CNPJ do emitente obtido da chave de acesso")
    if not nota["emitente_razao"]:
        obs.append("Razão social do emitente não localizada")
    m = re.search(r"N[ºO°]\.?\s*:?\s*(\d{1,3}(?:\.\d{3})+|\d+)", antes, re.I)
    if m:
        nota["numero"] = m.group(1).replace(".", "")
    else:
        nota["numero"] = str(int(chave[25:34])); obs.append("Número da NF obtido da chave de acesso")
    m = re.search(r"S[ÉE]RIE\s*:?\s*(\d+)", antes, re.I)
    if m:
        nota["serie"] = m.group(1)
    else:
        nota["serie"] = str(int(chave[22:25])); obs.append("Série obtida da chave de acesso")

    # destinatário e data
    nota["destinatario_razao"], nota["destinatario_cnpj"], nota["data_emissao"] = None, None, None
    if dest:
        a = achar_seq(p1, ["NOME", "RAZAO", "SOCIAL"], dest["li"])
        if a:
            ws = valor_abaixo(p1, a)
            nota["destinatario_razao"] = " ".join(w.t for w in ws) or None
        depois = [w for ws in p1[dest["li"]:] for w in ws]
        for w in depois:
            if RE_CNPJ_CPF_FMT.match(w.t) or re.fullmatch(r"\d{14}", w.t):
                nota["destinatario_cnpj"] = w.t
                break
        a = achar_seq(p1, ["DATA", "DA", "EMISSAO"], dest["li"])
        ws = valor_abaixo(p1, a, lambda w: RE_DATA.match(w.t)) if a else []
        if not ws:
            ws = [w for w in depois if RE_DATA.match(w.t)][:1]
        if ws:
            d, mth, y = ws[0].t.split("/")
            nota["data_emissao"] = f"{y}-{mth}-{d}"
    for campo, rot in (("destinatario_razao", "Razão social do destinatário"),
                       ("destinatario_cnpj", "CNPJ do destinatário"), ("data_emissao", "Data de emissão")):
        if not nota[campo]:
            obs.append(f"{rot} não localizado(a)")

    nota["totais"] = extrair_totais(pag)
    if not nota["totais"]:
        obs.append("Bloco de totais não localizado")
    comp = texto_complementar(pag)
    nota["informacoes_complementares"] = comp

    itens = extrair_itens(pag)
    for k, it in enumerate(itens, 1):
        it["numero_item"] = k
        texto = " ".join(it.pop("_desc"))
        it.pop("_dx0"); it.pop("_dx1")
        cod = it.pop("codigo")
        it["descricao"] = f"{cod} - {texto}" if cod and texto else (texto or cod or "")
        ids = identificadores(texto, it.get("quantidade"))
        if not ids and comp and len(itens) == 1:
            ids = identificadores(comp, it.get("quantidade"))
        it["identificadores"] = ids
        for campo in ("frete", "seguro", "icms_st", "outras_despesas"):
            it[campo] = None
    if comp and len(itens) > 1 and identificadores(comp, None):
        obs.append("Há série/chassi/RENAVAM nas informações complementares; conferir vínculo com os itens")
    nota["itens"] = itens
    if grupo.get("ocr"):
        obs.append("Texto obtido por OCR; conferir")
    nota["observacoes"] = "; ".join(obs) or None
    return nota


# --------------------------------------------------------------------------- #
# EXTRAÇÃO POR OCR (PDFs digitalizados): RapidOCR (somente pip) ou Tesseract (reserva)
# --------------------------------------------------------------------------- #
OCR_DPI = 300
_OCR = {"motor": None, "obj": None}


class Cx:
    """Caixa de texto OCR (coordenadas em pixels da página já na orientação correta)."""
    __slots__ = ("x0", "y0", "x1", "y1", "t", "n", "conf")

    def __init__(self, x0, y0, x1, y1, t, conf=1.0):
        self.x0, self.y0, self.x1, self.y1, self.t, self.conf = x0, y0, x1, y1, t.strip(), conf
        self.n = nrm(t)

    @property
    def yc(self):
        return (self.y0 + self.y1) / 2

    @property
    def xc(self):
        return (self.x0 + self.x1) / 2

    @property
    def h(self):
        return self.y1 - self.y0


def _motor_ocr():
    """Escolhe o motor de OCR disponível. Devolve 'rapid', 'tess' ou None."""
    if _OCR["motor"] is not None:
        return _OCR["motor"] or None
    try:
        from rapidocr_onnxruntime import RapidOCR
        _OCR["obj"], _OCR["motor"] = RapidOCR(), "rapid"
        return "rapid"
    except Exception:
        pass
    if ocr_disponivel():
        _OCR["motor"] = "tess"
        return "tess"
    _OCR["motor"] = ""
    return None


def _pagina_para_array(page):
    import numpy as np
    from PIL import Image
    pix = page.get_pixmap(dpi=OCR_DPI)
    img = np.array(Image.open(io.BytesIO(pix.tobytes("png"))).convert("L"))
    try:
        import cv2
        img = cv2.medianBlur(img, 3)  # reduz o ruído de scanner/fax
        img = _desinclina(img)
    except Exception:
        pass
    return img


def _desinclina(img):
    """Corrige a inclinação do scanner (até ~4°) maximizando a variância dos perfis horizontais."""
    import numpy as np
    import cv2
    pq = cv2.resize(img, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    bin_ = (pq < 150).astype(np.uint8)
    h, w = bin_.shape

    def pontua(ang):
        M = cv2.getRotationMatrix2D((w / 2, h / 2), ang, 1.0)
        r = cv2.warpAffine(bin_, M, (w, h), flags=cv2.INTER_NEAREST)
        return float(np.var(r.sum(axis=1)))
    grossos = [(pontua(a), a) for a in np.arange(-4.0, 4.01, 0.5)]
    a0 = max(grossos)[1]
    finos = [(pontua(a), a) for a in np.arange(a0 - 0.5, a0 + 0.51, 0.1)]
    ang = max(finos)[1]
    if abs(ang) < 0.15:
        return img
    H, W = img.shape
    M = cv2.getRotationMatrix2D((W / 2, H / 2), ang, 1.0)
    return cv2.warpAffine(img, M, (W, H), flags=cv2.INTER_LINEAR, borderValue=255)


def _ocr_rapid(img):
    res, _ = _OCR["obj"](img)
    out = []
    for b, t, c in (res or []):
        xs, ys = [p[0] for p in b], [p[1] for p in b]
        out.append(Cx(min(xs), min(ys), max(xs), max(ys), t, float(c)))
    return out


def _ocr_tess(img):
    """Tesseract: testa 0 e 180 graus e junta palavras vizinhas em frases."""
    import pytesseract
    from PIL import Image
    melhor = None
    for rot in (0, 180):
        im = Image.fromarray(img).rotate(rot)
        try:
            d = pytesseract.image_to_data(im, lang="por", config="--psm 6", output_type=pytesseract.Output.DICT)
        except Exception:
            d = pytesseract.image_to_data(im, config="--psm 6", output_type=pytesseract.Output.DICT)
        ws = [(d["left"][i], d["top"][i], d["left"][i] + d["width"][i], d["top"][i] + d["height"][i],
               d["text"][i].strip(), float(d["conf"][i])) for i in range(len(d["text"]))
              if d["text"][i].strip() and float(d["conf"][i]) >= 0]
        sc = sum(w[5] for w in ws)
        if melhor is None or sc > melhor[0]:
            melhor = (sc, ws, rot)
    _, ws, rot = melhor
    ws.sort(key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    frases, atual = [], None
    for w in ws:
        if atual and abs((w[1] + w[3]) / 2 - (atual[1] + atual[3]) / 2) < 0.5 * (w[3] - w[1]) \
                and 0 <= w[0] - atual[2] <= 1.2 * (w[3] - w[1]):
            atual[2], atual[4] = w[2], atual[4] + " " + w[4]
            atual[1], atual[3] = min(atual[1], w[1]), max(atual[3], w[3])
            atual[5] = min(atual[5], w[5])
        else:
            if atual:
                frases.append(atual)
            atual = list(w)
    if atual:
        frases.append(atual)
    return [Cx(f[0], f[1], f[2], f[3], f[4], f[5] / 100) for f in frases], rot


ANCORAS = {
    "DEST": "DESTINATARIOREMETENTE", "IMPOSTO": "CALCULODOIMPOSTO", "TRANSP": "TRANSPORTADORVOLUMESTRANSPORTADOS",
    "PROD": "DADOSDOPRODUTOSERVICOS", "ISSQN": "CALCULODOISSQN", "ADIC": "DADOSADICIONAIS",
}


DICAS = {"DEST": ("DESTIN", "REMET"), "IMPOSTO": ("IMPOS", "POSTO"), "TRANSP": ("TRANSP", "VOLUMES", "PORTADOR"),
         "PROD": ("PRODUT", "SERVI"), "ISSQN": ("ISSQN", "ISSON", "ISSN", "SQN"), "ADIC": ("ADICION", "ADIC")}


def _pont_ancora(n, nome):
    alvo = ANCORAS[nome]
    if len(n) < 8 or len(n) > len(alvo) + 6:
        return 0
    sc = max(_sim(n, alvo), _sim(n, alvo[:len(n)]))
    if any(d in n for d in DICAS[nome]):
        sc += 0.2
    elif sc < 0.8:
        return 0
    return sc if sc >= 0.74 else 0


def _casa_ancora(n, alvo):
    nome = next((k for k, v in ANCORAS.items() if v == alvo), None)
    return bool(nome and _pont_ancora(n, nome))


def _ancora(cxs, nome):
    c = [(x, _pont_ancora(x.n, nome)) for x in cxs]
    if nome == "IMPOSTO":
        c = [(x, p) for x, p in c if "ISS" not in x.n]
    c = [(x, p) for x, p in c if p]
    return max(c, key=lambda xp: (round(xp[1], 2), -xp[0].y0))[0] if c else None


def _gira180(cxs, W, H):
    return [Cx(W - c.x1, H - c.y1, W - c.x0, H - c.y0, c.t, c.conf) for c in cxs]


def _de_cabeca_para_baixo(cxs):
    """Vota pela ordem vertical das seções: DESTINATÁRIO vem antes das demais no formulário."""
    d = _ancora(cxs, "DEST")
    if not d:
        return False
    votos = 0
    for k in ("IMPOSTO", "TRANSP", "PROD", "ADIC", "ISSQN"):
        a = _ancora(cxs, k)
        if a:
            votos += 1 if d.y0 < a.y0 else -1
    return votos < 0


def ocr_pagina_caixas(page):
    """Devolve (caixas na orientação correta, motor, imagem na orientação correta) ou (None, None, None)."""
    import numpy as np
    motor = _motor_ocr()
    if motor is None:
        return None, None, None
    img = _pagina_para_array(page)
    if motor == "rapid":
        cxs = _ocr_rapid(img)
        vert = sum(1 for c in cxs if c.h > 2 * (c.x1 - c.x0) and len(c.t) > 3)
        if cxs and vert > 0.4 * len(cxs):  # página deitada: gira 90° e refaz
            melhor = None
            for k in (1, 3):
                im = np.rot90(img, k).copy()
                r = _ocr_rapid(im)
                sc = sum(1 for c in r if any(_casa_ancora(c.n, al) for al in ANCORAS.values()))
                if melhor is None or sc > melhor[0]:
                    melhor = (sc, r, im)
            cxs, img = melhor[1], melhor[2]
        H, W = img.shape
        if _de_cabeca_para_baixo(cxs):
            cxs, img = _gira180(cxs, W, H), np.rot90(img, 2).copy()
    else:
        cxs, rot = _ocr_tess(img)
        H, W = img.shape
        if rot == 180:
            cxs, img = _gira180(cxs, W, H), np.rot90(img, 2).copy()
    return cxs, motor, img


# ------------------------- leitura de números ------------------------------- #
_TROCA = str.maketrans({"O": "0", "o": "0", "l": "1", "I": "1", "|": "1", "S": "5", "B": "8", "Z": "2"})


def _limpa_num(s):
    return re.sub(r"[^\d.,]", "", s.translate(_TROCA))


def _interpreta(s):
    """'1.234,56' / '1,234.56' / '8892.0600' -> Decimal (último separador = decimal)."""
    m = re.fullmatch(r"(\d[\d.,]*?)[.,](\d{1,4})", s)
    if m:
        return Decimal(re.sub(r"[.,]", "", m.group(1)) + "." + m.group(2))
    d = re.sub(r"[.,]", "", s)
    return Decimal(d) if d else None


def cands_num(texto):
    """Interpretações plausíveis (a primeira é a principal) de um número lido por OCR."""
    s = _limpa_num(texto)
    out = []
    if not re.search(r"\d", s):
        return out
    for v in (s, s[:-1], s[1:]):
        try:
            x = _interpreta(v) if v else None
        except Exception:
            x = None
        if x is not None and x not in out:
            out.append(x)
    if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})+[.,]\d", s):  # ex.: 8.892,961 (lixo no fim)
        x = _interpreta(s[:-1])
        if x in out:
            out.remove(x)
        out.insert(0, x)
    return out


def num1(texto):
    c = cands_num(texto)
    return c[0] if c else None


RE_NUMTOK = re.compile(r"^[\dOolIS|B]{1,3}(?:[.,][\dOolIS|B]{3})*[.,]?[\dOolIS|B]*\D?$")
RE_CNPJ_OCR = re.compile(r"\d{2}[.:,\s]?\d{3}[.:,\s]?\d{3}[/\s]?\d{4}[-\s]?\d{2}")


def _cnpjs(texto):
    return [so_digitos(m.group(0)) for m in RE_CNPJ_OCR.finditer(texto)]


def _escolhe_cnpj(cands, pref=None):
    cands = [c for c in cands if len(c) == 14]
    if pref and pref in cands:
        return pref
    for c in cands:
        if cnpj_valido(c):
            return c
    return cands[0] if cands else None


# --------------------------- chave de acesso -------------------------------- #
def _rotulo_chave(cxs):
    return next((c for c in cxs if len(c.n) >= 9 and (_sim(c.n, "CHAVEDEACESSO") >= 0.75 or "CHAVEDEACES" in c.n)), None)


def achar_chave_ocr(cxs, img=None):
    def dig(c):
        t = c.t.translate(_TROCA) if sum(ch.isdigit() for ch in c.t) >= 0.7 * len(re.sub(r"\s", "", c.t)) else c.t
        return so_digitos(t)
    lab = _rotulo_chave(cxs)
    if lab:
        zona = [c for c in cxs if lab.y1 - 5 <= c.y0 <= lab.y1 + 130 and c.x1 >= lab.x0 - 150]
    else:
        H = max((c.y1 for c in cxs), default=0)
        zona = [c for c in cxs if c.y0 < 0.3 * H]
    cand = []
    for c in zona:
        d = dig(c)
        if len(d) >= 40:
            cand.append((0 if len(d) == 44 else 1, c.y0, d))
    for ln in agrupar_cx(zona, 0.9):  # chave quebrada em várias caixas
        d = "".join(dig(c) for c in ln if len(dig(c)) >= 3)
        if 40 <= len(d) <= 48:
            cand.append((0 if len(d) == 44 else 1, ln[0].y0, d))
    melhor = sorted(cand)[0][2] if cand else None
    if melhor and len(melhor) != 44 and lab is None:
        melhor = None
    if melhor and len(melhor) == 44 and chave_dv_ok(melhor):
        return melhor
    if img is not None and _OCR["motor"] == "rapid" and (lab is not None or _parece_danfe(cxs)):
        H, W = img.shape
        if lab is not None:
            reg = (max(0, int(lab.x0 - 60)), max(0, int(lab.y1 - 5)), min(W, int(lab.x0 + 1100)), min(H, int(lab.y1 + 130)))
        else:
            reg = (int(W * 0.45), int(H * 0.04), W, int(H * 0.16))
        r = _chave_focada(img, reg)
        if r:
            return r
    return melhor


def _parece_danfe(cxs):
    return any(c.n == "DANFE" or "DOCUMENTOAUXILIAR" in c.n or ("AUTENTICIDADE" in c.n and "NFE" in c.n)
               or _sim(c.n, "NOTAFISCALELETRONICA") > 0.8 for c in cxs)


def _chave_focada(img, regiao):
    """Relê só a área da chave (várias escalas/filtros) e vota pelas leituras de 44 dígitos, preferindo DV válido."""
    import numpy as np
    import cv2
    x0, y0, x1, y1 = regiao
    rec = img[y0:y1, x0:x1]
    if rec.size == 0:
        return None
    visto = {}
    for esc in (1.0, 1.25, 1.5, 0.8, 2.0):
        im = cv2.resize(rec, None, fx=esc, fy=esc, interpolation=cv2.INTER_CUBIC)
        for v in (im, cv2.GaussianBlur(im, (3, 3), 0)):
            v = cv2.copyMakeBorder(v, 120, 120, 80, 80, cv2.BORDER_CONSTANT, value=255)
            for c in _ocr_rapid(v):
                d = so_digitos(c.t.translate(_TROCA)) if sum(ch.isdigit() for ch in c.t) >= 0.6 * len(c.t.replace(" ", "")) else ""
                if len(d) == 44:
                    visto[d] = visto.get(d, 0) + 1
    ordenadas = sorted(visto.items(), key=lambda kv: -kv[1])
    for d, _ in ordenadas:
        if chave_dv_ok(d):
            return d
    return ordenadas[0][0] if ordenadas else None


def agrupar_cx(cxs, frac=0.6):
    cs = sorted(cxs, key=lambda c: (c.yc, c.x0))
    linhas = []
    for c in cs:
        if linhas and abs(c.yc - linhas[-1][0]) <= frac * max(c.h, 10):
            linhas[-1][1].append(c)
        else:
            linhas.append([c.yc, [c]])
    return [sorted(l[1], key=lambda c: c.x0) for l in linhas]


# ------------------------------ totais -------------------------------------- #
ROTULOS_TOT = {
    "valor_produtos": ["VALORTOTALDOSPRODUTOS"], "frete": ["VALORDOFRETE"], "seguro": ["VALORDOSEGURO"],
    "desconto": ["DESCONTO"], "outras_despesas": ["OUTRASDESPESASACESSORIAS", "OUTRASDESPESAS"],
    "ipi": ["VALORTOTALDOIPI", "VALORDOIPI"], "icms_st": ["VALORDOICMSSUBSTITUICAO", "VALORDOICMSST"],
    "valor_nf": ["VALORTOTALDANOTA"], "bc_icms": ["BASEDECALCULODOICMS"], "v_icms": ["VALORDOICMS"],
}


def _sim(a, b):
    import difflib
    return difflib.SequenceMatcher(None, a, b).ratio()


def _rotulo(c):
    """Qual rótulo de totais a caixa representa (tolerando erros de OCR)."""
    if len(c.n) < 5 or len(c.n) > 34:
        return None
    melhor = (0.8, None)
    for k, alvos in ROTULOS_TOT.items():
        for a in alvos:
            r = _sim(c.n, a)
            if r > melhor[0] or (r == melhor[0] and melhor[1] is None):
                melhor = (r, k)
    return melhor[1]


def extrair_totais_ocr(cxs, y_ini, y_fim):
    rot = []
    for c in cxs:
        if y_ini <= c.y0 <= y_fim:
            k = _rotulo(c)
            if k:
                rot.append((k, c))
    nums = [c for c in cxs if y_ini <= c.y0 <= y_fim + 60 and RE_NUMTOK.match(c.t.replace(" ", ""))
            and re.search(r"\d[.,]\d", c.t)]
    out = {k: [] for k in ROTULOS_TOT}
    usados = set()
    for linha in agrupar_cx([c for _, c in rot], 0.8):
        grupo = sorted(((k, c) for k, c in rot if c in linha), key=lambda kc: kc[1].x0)
        for i, (k, c) in enumerate(grupo):
            x0 = c.x0 - 25
            x1 = grupo[i + 1][1].x0 - 25 if i + 1 < len(grupo) else 1e9
            vs = [n for n in nums if id(n) not in usados and n.y0 > c.y1 - 8 and n.y0 - c.y1 < 120 and x0 <= n.xc < x1]
            if vs and not out[k]:
                v = min(vs, key=lambda n: n.y0)
                out[k] = cands_num(v.t)
                usados.add(id(v))
    out["_sobras"] = [cands_num(n.t) for n in nums if id(n) not in usados]
    # leitura posicional (rótulos ilegíveis): linha 1 = BC, ICMS, BC ST, ICMS ST, produtos; linha 2 = frete, seguro,
    # desconto, outras, IPI, total da NF
    linhas = [ln for ln in agrupar_cx(nums, 0.8) if len(ln) >= 5]
    pos = {}
    if linhas:
        a = linhas[0]
        for k, c in zip(("bc_icms", "v_icms", "bc_st", "icms_st", "valor_produtos"), a[:5]):
            pos[k] = cands_num(c.t)
        b = [ln for ln in linhas[1:] if len(ln) >= 6]
        if b:
            for k, c in zip(("frete", "seguro", "desconto", "outras_despesas", "ipi", "valor_nf"), b[-1][:6]):
                pos[k] = cands_num(c.t)
    out["_pos"] = pos
    return out


def resolve_totais(c):
    """Tenta a leitura por rótulos e, se não fechar, a posicional."""
    r = _resolve(c)
    if not r["_fechado"] and c.get("_pos"):
        r2 = _resolve(dict(c["_pos"], _sobras=[]))
        if r2["_fechado"]:
            return r2
    return r


def _resolve(c):
    """Escolhe entre as leituras candidatas a combinação que fecha: produtos+...-desconto = total da NF."""
    import itertools
    chaves = ["valor_produtos", "frete", "seguro", "outras_despesas", "ipi", "icms_st", "desconto", "valor_nf"]
    listas = [(c.get(k) or [ZERO if k not in ("valor_produtos", "valor_nf") else None])[:3] for k in chaves]
    esc = {k: (l[0] if l else None) for k, l in zip(chaves, listas)}
    fechado = False
    if all(l and l[0] is not None for l in listas if True) or True:
        for combo in itertools.product(*[l if l else [None] for l in listas]):
            if any(v is None for v in combo):
                continue
            vp, fr, sg, ou, ip, st, ds, nf = combo
            if abs(vp + fr + sg + ou + ip + st - ds - nf) <= TOLERANCIA:
                esc = dict(zip(chaves, combo))
                fechado = True
                break
    if not fechado:  # um campo sem rótulo lido: tenta um valor solto da área que feche a conta
        vazios = [k for k in chaves if not c.get(k) and k not in ("valor_produtos", "valor_nf")]
        for k in vazios:
            for sobra in c.get("_sobras", []):
                for v in sobra[:2]:
                    t = dict(esc)
                    t[k] = v
                    if None not in t.values() and abs(t["valor_produtos"] + t["frete"] + t["seguro"] + t["outras_despesas"]
                                                      + t["ipi"] + t["icms_st"] - t["desconto"] - t["valor_nf"]) <= TOLERANCIA:
                        esc, fechado = t, True
                        break
                if fechado:
                    break
            if fechado:
                break
    esc["_fechado"] = fechado
    return esc


# ------------------------------- itens -------------------------------------- #
RE_NCM = re.compile(r"^\d{7,9}$")
RE_CFOP = re.compile(r"^[1-7]\d{3}$")


def _tok_num(c):
    t = c.t.replace(" ", "")
    return bool(re.search(r"\d", t)) and bool(RE_NUMTOK.match(t))


def _divide_numeros(cxs):
    """Caixas com vários números ('1,0000  80.000,0000') viram uma caixa por número."""
    out = []
    for c in cxs:
        partes = c.t.split()
        if len(partes) > 1 and all(re.search(r"\d", p) and RE_NUMTOK.match(p) for p in partes):
            larg = (c.x1 - c.x0) / max(1, len(c.t))
            pos = 0
            for p in partes:
                k = c.t.index(p, pos)
                out.append(Cx(c.x0 + k * larg, c.y0, c.x0 + (k + len(p)) * larg, c.y1, p, c.conf))
                pos = k + len(p)
        else:
            out.append(c)
    return out


def _cfop(t):
    m = re.fullmatch(r"\d?([1-7]\d{3})", t)
    return m.group(1) if m else None


def extrair_itens_ocr(cxs, y_ini, y_fim):
    corpo = _divide_numeros([c for c in cxs if y_ini <= c.yc <= y_fim])
    linhas = agrupar_cx(corpo, 0.7)
    ancoras = []
    for li, ws in enumerate(linhas):
        p = next((i for i, c in enumerate(ws) if RE_NCM.match(c.t) and any(_cfop(w.t) for w in ws[i + 1:i + 3])), None)
        if p is not None:
            ancoras.append((li, p))
    itens = []
    for k, (li, p) in enumerate(ancoras):
        ws = linhas[li]
        resto = ws[p + 1:]
        cst = cfop = un = None
        j = 0
        for q in range(min(2, len(resto))):
            if _cfop(resto[q].t):
                cfop = _cfop(resto[q].t)
                if q == 1:
                    cst = resto[0].t
                j = q + 1
                break
        if j < len(resto) and not _tok_num(resto[j]):
            un, j = resto[j].t, j + 1
        nums = [c for c in resto[j:] if _tok_num(c)]
        antes = ws[:p]
        desc_x0 = antes[0].x0 if antes else ws[p].x0 - 400
        codigo = antes[0].t if len(antes) >= 2 else None
        desc = [" ".join(c.t for c in (antes[1:] if codigo else antes))]
        it = {"ncm": ws[p].t, "cst_csosn": cst, "cfop": cfop, "unidade": un, "leitura_incompleta": False,
              "observacoes": None, "_nums": [cands_num(c.t) for c in nums]}
        # continuação da descrição: linhas seguintes até o próximo item, dentro da coluna de descrição
        fim = ancoras[k + 1][0] if k + 1 < len(ancoras) else len(linhas)
        for lj in range(li + 1, fim):
            cont = [c.t for c in linhas[lj] if c.x0 >= desc_x0 - 15 and c.x1 <= ws[p].x0 + 5 and not _tok_num(c)]
            if cont:
                desc.append(" ".join(cont))
        it["_desc"] = [d for d in desc if d]
        it["_codigo"] = codigo
        itens.append(it)
    for it in itens:
        _aplica_colunas(it)
    return itens


def _aplica_colunas(it):
    nums = [n for n in it.pop("_nums") if n]
    obs = []
    if len(nums) >= 9:
        perfil = ["quantidade", "valor_unitario", "valor_total", "desconto", "bc_icms", "valor_icms", "ipi", "aliq_icms", "aliq_ipi"]
    elif len(nums) == 8:
        perfil = ["quantidade", "valor_unitario", "valor_total", "bc_icms", "valor_icms", "ipi", "aliq_icms", "aliq_ipi"]
    else:
        perfil = ["quantidade", "valor_unitario", "valor_total"][:len(nums)]
        it["leitura_incompleta"] = True
        obs.append(f"Colunas numéricas lidas por OCR: {len(nums)}; conferir com o PDF")
    for k, v in zip(perfil, nums):
        it[k] = v[0]
    q, u, t = it.get("quantidade"), it.get("valor_unitario"), it.get("valor_total")
    if q is not None and u is not None and t is not None:
        melhor = None
        for qq in (nums[0] if nums else []):
            for uu in nums[1]:
                for tt in nums[2]:
                    e = abs(qq * uu - tt)
                    if melhor is None or e < melhor[0]:
                        melhor = (e, qq, uu, tt)
        if melhor and melhor[0] <= Decimal("0.05"):
            _, it["quantidade"], it["valor_unitario"], it["valor_total"] = melhor
        else:
            obs.append("Quantidade x valor unitário não fecha com o total do item (OCR); conferir")
            it["leitura_incompleta"] = True
    for a in ("aliq_icms", "aliq_ipi"):
        if it.get(a) is not None and it[a] > 100:
            it[a] = None
    if obs:
        it["observacoes"] = "; ".join(obs)


# ------------------------------ nota ---------------------------------------- #
RE_VIN = re.compile(r"\b(?=[A-Z0-9]*\d)(?=[A-Z0-9]*[A-Z])[A-HJ-NPR-Z0-9]{17}\b")


def _texto_secao(cxs, y0, y1):
    return " ".join(c.t for ln in agrupar_cx([c for c in cxs if y0 <= c.yc < y1]) for c in ln)


def montar_nota_ocr(grupo):
    cxs = [c for pg in grupo["paginas"] for c in pg]  # pg: caixas (1ª página define o cabeçalho)
    p1 = grupo["paginas"][0]
    chave = grupo["chave"]
    obs = ["Texto obtido por OCR; conferir com o PDF"]
    H = max(c.y1 for c in p1)
    a = {k: _ancora(p1, k) for k in ANCORAS}
    ydest = a["DEST"].y0 if a["DEST"] else H * 0.25
    topo = [c for c in p1 if c.y1 < ydest + 5]
    tipo = "Outro"
    if chave and len(chave) == 44:
        tipo = {"55": "DANFE NF-e (55)", "65": "NFC-e (65)"}.get(chave[20:22], "Outro")
    elif a["DEST"] and (a["IMPOSTO"] or a["TRANSP"] or a["PROD"]):
        eh_danfe = any(c.n == "DANFE" or "DOCUMENTOAUXILIAR" in c.n or "NOTAFISCALELETRONICA" in c.n
                       or _sim(c.n, "NOTAFISCALELETRONICA") > 0.8
                       or ("AUTENTICIDADE" in c.n and "NFE" in c.n) for c in p1) or _rotulo_chave(p1) is not None
        tipo = "DANFE NF-e (55)" if eh_danfe else "NF modelo 1/1A"
    nota = {"tipo_documento": tipo, "chave_acesso": chave, "legivel": True}
    if tipo not in ("DANFE NF-e (55)", "NF modelo 1/1A"):
        nota["nao_suportado"] = True
        nota["observacoes"] = "Documento não reconhecido como DANFE modelo 55 nem NF modelo 1/1A; lançar manualmente"
        return nota

    # emitente
    txt_topo = " ".join(c.t for ln in agrupar_cx(topo) for c in ln)
    txt_pag = " ".join(c.t for ln in agrupar_cx(p1) for c in ln)
    m = re.search(r"RECEBEMOS\s*DE\s*(.+?)\s*OS\s*PRODUTOS", txt_topo, re.I) or \
        re.search(r"(?:RECEB|RECEM|ROC)\w*\s+\w{1,2}\s+(.+?)\s+Os\s+pr", txt_pag, re.I)
    razao = m.group(1).strip() if m else None
    if not razao:
        danfe = next((c for c in topo if c.n == "DANFE"), None)
        xmax = danfe.x0 - 20 if danfe else 1e9
        nat = next((c for c in topo if _sim(c.n, "NATUREZADAOPERACAO") >= 0.75), None)
        bloco = [c for c in topo if c.x1 < xmax and c.conf >= 0.6 and re.search(r"[A-Za-z]{3}", c.t)
                 and (nat is None or c.y1 < nat.y0)
                 and not re.search(r"RECEB|IDENTIFIC|DATA|ASSINATURA|SERIE|N[º°O]|VENDA|CFOP|AVENIDA|RUA|CEP|FONE|CNPJ|INSC|SEFAZ|AUTORIZ|PORTAL|CONSULTA|WWW|FAZENDA|REMESSA|PROTOCOLO|SIMPLES", c.n)
                 and c.y0 > 0.04 * H]
        if bloco:
            suf = [c for c in bloco if re.search(r"LTDA|S/?A$|EIRELI|\bME$|EPP|COMERCIO|INDUSTRIA|MINERA", c.n)]
            if suf:  # linha com sufixo societário + linhas vizinhas de mesma altura logo acima
                base = max(suf, key=lambda c: c.h)
                grandes = sorted((c for c in bloco if base.y0 - 2.2 * base.h <= c.y0 <= base.y1 and c.h >= 0.8 * base.h),
                                 key=lambda c: (c.y0, c.x0))
            else:
                hmax = max(c.h for c in bloco)
                grandes = sorted((c for c in bloco if c.h >= 0.85 * hmax), key=lambda c: (c.y0, c.x0))
            razao = " ".join(c.t for c in grandes[:3])
    razao = re.sub(r"(?<=[A-Z])(LTDA|S/?A|EIRELI|ME|EPP)\b", r" \1", razao or "") or None
    nota["emitente_razao"] = razao
    if not razao:
        obs.append("Razão social do emitente não localizada")
    ecn = _escolhe_cnpj([x for c in topo for x in _cnpjs(c.t)], chave[6:20] if chave and len(chave) == 44 else None)
    nota["emitente_cnpj"] = ecn

    # número e série (OCR); a chave, quando válida, prevalece
    mnum = None
    for c in sorted(topo, key=lambda c: c.y0):
        mm = re.search(r"N\s*[º°oO0]\.?\s*[:\-]?\s*(\d{1,3}(?:\.\d{3})+|\d{3,9})", c.t)
        if mm:
            mnum = mm.group(1).replace(".", "")
            break
    if not mnum:  # número solto, perto do rótulo "SÉRIE"
        rs = next((c for c in topo if c.n.startswith("SERIE") and len(c.n) <= 8), None)
        if rs:
            vs = [c for c in topo if re.fullmatch(r"\d{3,9}", c.t) and rs.y0 - 190 <= c.y0 <= rs.y1 and abs(c.xc - rs.xc) < 350]
            if vs:
                mnum = min(vs, key=lambda c: abs(c.y0 - rs.y0)).t
    ms = None
    for c in sorted(topo, key=lambda c: c.y0):
        mm = re.search(r"S[ÉE]RIE\s*:?\s*([\dOo]{1,3})\b", c.t, re.I)
        if mm:
            ms = mm.group(1).translate(_TROCA)
            break
    nota["numero"] = str(int(mnum)) if mnum else None
    nota["serie"] = str(int(ms)) if ms else None
    if chave and len(chave) == 44 and chave_dv_ok(chave):
        if not nota["emitente_cnpj"] or not cnpj_valido(nota["emitente_cnpj"]) or nota["emitente_cnpj"] != chave[6:20]:
            if nota["emitente_cnpj"] != chave[6:20]:
                obs.append("CNPJ do emitente obtido da chave de acesso (OCR divergente ou ilegível)")
            nota["emitente_cnpj"] = chave[6:20]
        n_chave, s_chave = str(int(chave[25:34])), str(int(chave[22:25]))
        if nota["numero"] != n_chave:
            if nota["numero"]:
                obs.append(f"Número lido por OCR ({nota['numero']}) difere da chave; usado o da chave")
            nota["numero"] = n_chave
        if nota["serie"] != s_chave:
            nota["serie"] = s_chave
    elif tipo == "DANFE NF-e (55)":
        if not nota["numero"] and chave and len(chave) == 44:
            nota["numero"] = str(int(chave[25:34]))
    if not nota["numero"]:
        obs.append("Número da NF não localizado")

    # destinatário
    nota["destinatario_razao"], nota["destinatario_cnpj"], nota["data_emissao"] = None, None, None
    yfim_dest = min([x.y0 for x in (a["IMPOSTO"], a["TRANSP"], a["PROD"]) if x and x.y0 > ydest] or [ydest + 500])
    reg = [c for c in p1 if ydest <= c.y0 < yfim_dest]
    rot_nome = next((c for c in sorted(reg, key=lambda c: c.y0)
                    if len(c.n) <= 22 and (re.search(r"NOME.*RAZ|RAZAOSOC", c.n) or _sim(c.n, "NOMERAZAOSOCIAL") >= 0.72)), None)
    rot_cnpj = next((c for c in reg if re.search(r"CNPJ|CPF", c.n) and len(c.n) < 12), None)
    if rot_nome:
        xmax = rot_cnpj.x0 - 10 if rot_cnpj and rot_cnpj.x0 > rot_nome.x1 else 1e9
        vs = [c for c in reg if rot_nome.y1 - 8 < c.y0 < rot_nome.y1 + 90 and c.x0 >= rot_nome.x0 - 25 and c.x1 < xmax + 40
              and not _cnpjs(c.t) and not re.fullmatch(r"[\d/.\- ]+", c.t)]
        if vs:
            y1 = min(vs, key=lambda c: c.y0).yc
            linha = sorted((c for c in vs if abs(c.yc - y1) < 25), key=lambda c: c.x0)
            nm = " ".join(c.t for c in linha)
            nm = re.sub(r"\s+\d{5,}$", "", nm)
            nota["destinatario_razao"] = re.sub(r"(?<=[A-Z])(LTDA|S/?A|EIRELI|ME|EPP)\b", r" \1", nm) or None
    if not nota["destinatario_razao"]:  # reserva: caixa com nome de empresa do grupo
        g = next((c for c in sorted(reg, key=lambda c: c.y0) if len(c.n) < 60 and
                  any(k in c.n for ks in NOME_GRUPO.values() for k in ks)), None)
        if g:
            nm = re.sub(r"\s+\d{5,}$", "", g.t)
            nota["destinatario_razao"] = re.sub(r"(?<=[A-Z])(LTDA|S/?A|EIRELI|ME|EPP)\b", r" \1", nm)
    nota["destinatario_cnpj"] = _escolhe_cnpj([x for c in reg for x in _cnpjs(c.t)])
    datas = []
    for c in reg:
        for mm in re.finditer(r"(\d{2})/(\d{2})/(\d{4})", c.t):
            datas.append((c.y0, mm.groups()))
    rot_data = next((c for c in reg if "DATADAEMISSAO" in c.n or "DATAEMISSAO" in c.n), None)
    if rot_data:
        vs = [(abs(c.xc - rot_data.xc), c) for c in reg if rot_data.y1 - 8 < c.y0 < rot_data.y1 + 70
              and re.search(r"\d{2}/\d{2}/\d{4}", c.t)]
        if vs:
            mm = re.search(r"(\d{2})/(\d{2})/(\d{4})", min(vs, key=lambda v: v[0])[1].t)
            nota["data_emissao"] = f"{mm.group(3)}-{mm.group(2)}-{mm.group(1)}"
    if not nota["data_emissao"] and datas:
        y, (dd, mm_, aa) = sorted(datas, key=lambda d: d[0])[0]
        nota["data_emissao"] = f"{aa}-{mm_}-{dd}"
    if not nota["data_emissao"] and tipo == "DANFE NF-e (55)":
        prot = [c for c in p1 if c.y0 < ydest and re.search(r"\d{2}/\d{2}/\d{4}", c.t) and re.search(r"\d{10,}", c.t)]
        if prot:
            mm = re.search(r"(\d{2})/(\d{2})/(\d{4})", prot[0].t)
            nota["data_emissao"] = f"{mm.group(3)}-{mm.group(2)}-{mm.group(1)}"
            obs.append("Data de emissão ilegível; usada a data do protocolo de autorização (conferir)")
    for campo, rot in (("destinatario_razao", "Razão social do destinatário"),
                       ("destinatario_cnpj", "CNPJ do destinatário"), ("data_emissao", "Data de emissão")):
        if not nota[campo]:
            obs.append(f"{rot} não localizado(a)")

    # totais
    todas = [c for pg in grupo["paginas"] for c in pg]
    yi = a["IMPOSTO"].y0 if a["IMPOSTO"] else ydest
    yf = min([x.y0 for x in (a["TRANSP"], a["PROD"], a["ISSQN"]) if x and x.y0 > yi + 10] or [yi + 420])
    tc = extrair_totais_ocr(p1, yi, yf)
    tot = resolve_totais(tc)
    if tot.get("valor_nf") is None and tot.get("valor_produtos") is None:
        obs.append("Bloco de totais não localizado")
        nota["totais"] = {}
    else:
        if not tot.pop("_fechado", False):
            obs.append("Totais lidos por OCR não fecham entre si (produtos+frete+...=total); conferir valores")
        tot.pop("_fechado", None)
        nota["totais"] = {k: v for k, v in tot.items() if k in ("valor_produtos", "frete", "seguro", "desconto",
                                                                 "outras_despesas", "ipi", "icms_st", "valor_nf")}
    # itens (todas as páginas da nota)
    itens = []
    for pg in grupo["paginas"]:
        ap, ai, aa = _ancora(pg, "PROD"), _ancora(pg, "ISSQN"), _ancora(pg, "ADIC")
        ys = [x.y0 for x in (ai, aa) if x]
        y_ini = ap.y1 if ap else 0
        y_fim = min([y for y in ys if y > y_ini] or [max(c.y1 for c in pg)])
        if tipo == "NF modelo 1/1A":  # na modelo 1 os itens ficam antes de "Cálculo dos impostos"
            ai2 = _ancora(pg, "IMPOSTO")
            if ai2 and ai2.y0 > y_ini:
                y_fim = ai2.y0
        itens += extrair_itens_ocr(pg, y_ini, y_fim)
    # informações complementares (para chassi/série)
    comp = ""
    pi = _ancora(p1, "ADIC")
    if pi:
        comp = _texto_secao(p1, pi.y1, max(c.y1 for c in p1) + 1)
        comp = re.sub(r"RESERVADO\s*AO\s*FISCO|INFORMA..ES\s*COMPLEMENTARES", " ", comp, flags=re.I)
    nota["informacoes_complementares"] = comp
    for k, it in enumerate(itens, 1):
        it["numero_item"] = k
        texto = " ".join(it.pop("_desc", []))
        cod = it.pop("_codigo", None)
        it["descricao"] = f"{cod} - {texto}" if cod and texto else (texto or cod or "")
        fonte = texto
        ids = identificadores(fonte, it.get("quantidade"))
        if not ids and comp and len(itens) == 1:
            ids = identificadores(comp, it.get("quantidade"))
            vin = RE_VIN.findall(comp.upper())
            if vin and not any(i.get("chassi") for i in ids):
                ids = (ids or [{}])
                ids[0] = dict(ids[0], chassi=vin[0])
        it["identificadores"] = ids
        for campo in ("frete", "seguro", "icms_st", "outras_despesas"):
            it[campo] = None
    if comp and len(itens) > 1 and (identificadores(comp, None) or RE_VIN.search(comp.upper())):
        obs.append("Há série/chassi/RENAVAM nas informações complementares; conferir vínculo com os itens")
    if not itens and nota["totais"].get("valor_produtos") is not None:
        vp = nota["totais"]["valor_produtos"]
        itens = [{"numero_item": 1, "descricao": "ITENS NÃO LIDOS POR OCR (manuscrito/ilegível) - conferir no PDF",
                  "quantidade": Decimal(1), "valor_unitario": vp, "valor_total": vp, "ncm": None, "cfop": None,
                  "cst_csosn": None, "unidade": None, "identificadores": [], "leitura_incompleta": True,
                  "observacoes": "Item não identificado; valor total dos produtos da NF lançado em linha única",
                  "frete": None, "seguro": None, "icms_st": None, "outras_despesas": None}]
    _reconcilia_itens(itens, nota["totais"].get("valor_produtos"))
    nota["itens"] = itens
    nota["observacoes"] = "; ".join(obs)
    return nota


def _reconcilia_itens(itens, vprod):
    """Se qtd x unitário fecha com o total dos produtos mas o total lido do item não, adota qtd x unitário."""
    if not itens or vprod is None:
        return
    lidos = [i.get("valor_total") for i in itens]
    if None not in lidos and abs(sum(lidos) - vprod) <= TOLERANCIA:
        return
    alt = []
    for i in itens:
        q, u = i.get("quantidade"), i.get("valor_unitario")
        alt.append(q2(q * u) if q is not None and u is not None else i.get("valor_total"))
    if None not in alt and abs(sum(alt) - vprod) <= TOLERANCIA:
        for i, v in zip(itens, alt):
            if i["valor_total"] != v:
                i["valor_total"] = v
                i["observacoes"] = ((i.get("observacoes") or "") + "; Valor total do item = quantidade x valor unitário "
                                    "(leitura OCR do total não fechava)").strip("; ")
    elif len(itens) == 1 and not (itens[0].get("quantidade") and itens[0].get("valor_unitario") and
                                  abs(itens[0]["quantidade"] * itens[0]["valor_unitario"] - vprod) <= TOLERANCIA):
        i = itens[0]  # item único: o total do item é o total dos produtos da NF; quantidade/unitário não confiáveis
        i["valor_total"], i["quantidade"], i["valor_unitario"] = vprod, Decimal(1), vprod
        i["leitura_incompleta"] = True
        i["observacoes"] = ((i.get("observacoes") or "") + "; Quantidade e valor unitário não lidos: item único, "
                            "valor total = total dos produtos da NF e quantidade 1 assumida (conferir no PDF)").strip("; ")


def classifica_pagina_ocr(cxs):
    """True se a página parece parte de uma nota fiscal."""
    ancoras = sum(1 for k in ANCORAS if _ancora(cxs, k))
    termos = ["CNPJ", "RAZAOSOCIAL", "DANFE", "NOTAFISCAL", "VALORTOTAL", "INSCRICAOESTADUAL", "NATUREZADAOPERACAO"]
    return ancoras + sum(1 for k in termos if any(k in c.n for c in cxs)) >= 4


def extrair_pdf_local(caminho: Path, ocr_ok, diag_dir):
    doc = fitz.open(str(caminho))
    if doc.needs_pass and not doc.authenticate(""):
        raise RuntimeError("PDF protegido por senha")
    paginas, avisos, diag = [], [], []
    for n, pg in enumerate(doc, 1):
        ws = [Pal(*w[:5]) for w in pg.get_text("words")]
        if sum(len(w.t) for w in ws) >= 60:  # PDF com camada de texto
            linhas = agrupar_linhas(ws)
            paginas.append({"modo": "texto", "dados": linhas, "n": n})
            diag.append(f"=== página {n} (texto) ===")
            diag += [" | ".join(w.t for w in l) for l in linhas]
            continue
        if not ocr_ok:
            avisos.append(f"página {n} digitalizada e nenhum OCR instalado")
            continue
        cxs, motor, img_up = ocr_pagina_caixas(pg)
        diag.append(f"=== página {n} (OCR {motor}) ===")
        diag += [" | ".join(f"{c.t}" for c in ln) for ln in agrupar_cx(cxs)]
        if not classifica_pagina_ocr(cxs):
            avisos.append(f"página {n} não parece ser nota fiscal (anexo, CRV ou ilegível)")
            continue
        paginas.append({"modo": "ocr", "dados": cxs, "n": n, "img": img_up})
    if diag_dir is not None and diag:
        diag_dir.mkdir(parents=True, exist_ok=True)
        nome = re.sub(r"[^\w\-]+", "_", str(caminho))[-120:]
        (diag_dir / f"{nome}.txt").write_text("\n".join(diag), encoding="utf-8")
    if not paginas:
        return {"notas": [], "arquivo_ilegivel": True, "observacao_arquivo": "; ".join(avisos)}

    grupos = []
    for p in paginas:
        if p["modo"] == "texto":
            ch = achar_chave(p["dados"])
            inicia = bool(ch) and (not grupos or grupos[-1]["chave"] != ch)
        else:
            ch = achar_chave_ocr(p["dados"], p.get("img"))
            inicia = bool(ch and (not grupos or grupos[-1]["chave"] != ch)) or (
                not ch and (_ancora(p["dados"], "DEST") is not None or _ancora(p["dados"], "IMPOSTO") is not None))
        if inicia or not grupos:
            grupos.append({"chave": ch, "paginas": [p["dados"]], "ocr": p["modo"] == "ocr", "modo": p["modo"]})
        else:
            grupos[-1]["paginas"].append(p["dados"])
    notas = [(montar_nota_ocr(g) if g["modo"] == "ocr" else montar_nota(g)) for g in grupos]
    return {"notas": notas, "arquivo_ilegivel": False, "observacao_arquivo": "; ".join(avisos)}


# --------------------------------------------------------------------------- #
# PROCESSAMENTO DAS NOTAS
# --------------------------------------------------------------------------- #
class Contexto:
    def __init__(self):
        self.linhas = []        # cada item: lista com 39 valores (None nas colunas de fórmula)
        self.meta = []          # (empresa, ncm, custo, icms, ipi) por linha, para o resumo no console
        self.notas = []         # dados para a aba Conferência
        self.pend = []          # (tipo, arquivo, nf, chave, item, detalhe)
        self.vistos = {}        # chave/identificador -> arquivo

    def add_pend(self, tipo, arq, nf=None, chave=None, item=None, detalhe=""):
        self.pend.append((tipo, arq, nf, chave, item, detalhe))


def componente(itens, campo, total_nota, pesos):
    """Valor por item se a nota o traz; senão rateia o total da nota pelo valor total do item."""
    vals = [D(i.get(campo)) for i in itens]
    if any(v is not None for v in vals) and sum((v or ZERO) for v in vals) > 0:
        return [v if v is not None else ZERO for v in vals]
    t = D(total_nota)
    if t:
        return allocate(t, pesos)
    return [ZERO] * len(itens)


def processar_nota(ctx: Contexto, arq: str, nota: dict):
    tipo = nota.get("tipo_documento") or "Outro"
    obs_nota = []
    if txt(nota.get("observacoes")):
        obs_nota.append(nota["observacoes"].strip())

    emit_cnpj = so_digitos(nota.get("emitente_cnpj"))
    dest_cnpj = so_digitos(nota.get("destinatario_cnpj"))
    numero = num_ou_texto(nota.get("numero"))
    serie = num_ou_texto(nota.get("serie"))
    ref_nf = numero if numero is not None else "?"

    # ---- chave de acesso ----
    chave, val_chave = "", ""
    if tipo in TIPOS_COM_CHAVE:
        chave = so_digitos(nota.get("chave_acesso"))
        if len(chave) != 44:
            val_chave = "ilegível"
            ctx.add_pend(P_CHAVE_INV, arq, ref_nf, chave or None, None, "Chave ausente, incompleta ou ilegível")
        elif not chave_dv_ok(chave):
            val_chave = "DV inválido"
            ctx.add_pend(P_CHAVE_INV, arq, ref_nf, chave, None, "Dígito verificador (módulo 11) não confere")
        else:
            val_chave = "OK"
        if len(chave) == 44:
            inc = []
            if emit_cnpj and chave[6:20] != emit_cnpj:
                inc.append(f"CNPJ da chave {chave[6:20]} x emitente {emit_cnpj}")
            if chave[20:22] != MODELO_POR_TIPO[tipo]:
                inc.append(f"modelo da chave {chave[20:22]} x tipo {tipo}")
            if isinstance(serie, int) and int(chave[22:25]) != serie:
                inc.append(f"série da chave {int(chave[22:25])} x nota {serie}")
            if isinstance(numero, int) and int(chave[25:34]) != numero:
                inc.append(f"número da chave {int(chave[25:34])} x nota {numero}")
            if inc:
                ctx.add_pend(P_CHAVE_INC, arq, ref_nf, chave, None, "; ".join(inc))
                obs_nota.append("Chave incoerente: " + "; ".join(inc))

    if emit_cnpj and len(emit_cnpj) == 14 and not cnpj_valido(emit_cnpj):
        obs_nota.append("CNPJ do emitente com DV inválido")
    if dest_cnpj and len(dest_cnpj) == 14 and not cnpj_valido(dest_cnpj):
        obs_nota.append("CNPJ do destinatário com DV inválido")
    if nota.get("legivel") is False:
        ctx.add_pend(P_ILEGIVEL, arq, ref_nf, chave or None, None, "Nota parcialmente ilegível: " + (nota.get("observacoes") or ""))

    # ---- duplicidade ----
    ident = chave if len(chave) == 44 else f"{tipo}|{emit_cnpj}|{numero}|{serie}"
    if ident in ctx.vistos:
        ctx.add_pend(P_DUP, arq, ref_nf, chave or None, None, f"Já lançada a partir de: {ctx.vistos[ident]}")
        return
    ctx.vistos[ident] = arq

    # ---- data e empresa ----
    data = None
    if txt(nota.get("data_emissao")):
        try:
            data = datetime.strptime(nota["data_emissao"][:10], "%Y-%m-%d").date()
        except ValueError:
            obs_nota.append(f"Data de emissão não interpretada: {nota['data_emissao']}")
    empresa = classificar(dest_cnpj, nota.get("destinatario_razao"))
    if empresa == "Fora do grupo":
        ctx.add_pend(P_FORA, arq, ref_nf, chave or None, None,
                     f"{nota.get('destinatario_razao') or '?'} - {fmt_cnpj(dest_cnpj) or 'CNPJ ausente'}")
        obs_nota.append("Destinatário fora do grupo")

    # ---- rateios ----
    itens = nota.get("itens") or []
    tot = nota.get("totais") or {}
    vtot = [D(i.get("valor_total")) or ZERO for i in itens]
    comp = {
        "desconto": componente(itens, "desconto", tot.get("desconto"), vtot),
        "ipi": componente(itens, "ipi", tot.get("ipi"), vtot),
        "frete": componente(itens, "frete", tot.get("frete"), vtot),
        "seguro": componente(itens, "seguro", tot.get("seguro"), vtot),
        "icms_st": componente(itens, "icms_st", tot.get("icms_st"), vtot),
        "outras": componente(itens, "outras_despesas", tot.get("outras_despesas"), vtot),
    }
    vnf = D(tot.get("valor_nf"))

    linhas_nota = []
    soma_custo = ZERO
    for k, item in enumerate(itens):
        qtd = D(item.get("quantidade"))
        ids = [i for i in (item.get("identificadores") or [])
               if any(txt(i.get(c)) for c in ("numero_serie", "chassi", "renavam", "placa"))]
        obs_item = []
        if item.get("leitura_incompleta"):
            ctx.add_pend(P_LEIT, arq, ref_nf, chave or None, item.get("numero_item"),
                         "Colunas numéricas do item não reconhecidas por completo")
        if txt(item.get("observacoes")):
            obs_item.append(item["observacoes"].strip())

        n = 1
        if qtd is not None and qtd == qtd.to_integral_value() and qtd > 1 and len(ids) == int(qtd):
            n = int(qtd)
            obs_item.append(f"Desdobrado de quantidade {int(qtd)}")
        elif len(ids) > 1:
            obs_item.append(f"{len(ids)} identificadores para quantidade {qtd}; não desdobrado: " +
                            "; ".join(str({c: v for c, v in i.items() if v}) for i in ids))

        def parte(valor):
            if valor is None:
                return [None] * n
            return allocate(valor, [Decimal(1)] * n) if n > 1 else [valor]

        p_vtot = parte(vtot[k])
        p_desc, p_ipi = parte(comp["desconto"][k]), parte(comp["ipi"][k])
        p_frete, p_seg = parte(comp["frete"][k]), parte(comp["seguro"][k])
        p_st, p_out = parte(comp["icms_st"][k]), parte(comp["outras"][k])
        p_bc, p_icms = parte(D(item.get("bc_icms"))), parte(D(item.get("valor_icms")))

        for u in range(n):
            if n > 1:
                idf = ids[u]
            elif len(ids) == 1:
                idf = ids[0]
            else:
                idf = {}
            serie_b, chassi, renavam, placa = (txt(idf.get(c)) for c in ("numero_serie", "chassi", "renavam", "placa"))
            obs = list(obs_nota) + list(obs_item)
            if chassi and not re.fullmatch(r"[A-Za-z0-9]{17}", chassi):
                obs.append("Chassi fora do padrão de 17 caracteres")
                ctx.add_pend(P_ID_PADRAO, arq, ref_nf, chave or None, item.get("numero_item"), f"Chassi: {chassi}")
            if renavam and not re.fullmatch(r"\d{11}", renavam):
                obs.append("RENAVAM fora do padrão de 11 dígitos")
                ctx.add_pend(P_ID_PADRAO, arq, ref_nf, chave or None, item.get("numero_item"), f"RENAVAM: {renavam}")
            if not (serie_b or chassi or renavam):
                obs.append("Bem sem número de série, chassi ou RENAVAM")
                ctx.add_pend(P_SEM_ID, arq, ref_nf, chave or None, item.get("numero_item"),
                             (txt(item.get("descricao")) or "")[:150])

            qtd_linha = Decimal(1) if n > 1 else qtd
            custo = (p_vtot[u] or ZERO) + p_ipi[u] + p_frete[u] + p_seg[u] + p_st[u] + p_out[u] - p_desc[u]
            soma_custo += custo

            linha = [None] * len(COLUNAS)
            linha[IDX["Arquivo de origem"]] = arq
            linha[IDX["Tipo de documento"]] = tipo
            linha[IDX["Chave de acesso"]] = chave or None
            linha[IDX["Validação da chave"]] = val_chave or None
            linha[IDX["Número da NF"]] = numero
            linha[IDX["Série"]] = serie
            linha[IDX["Data de emissão"]] = data
            linha[IDX["Emitente: razão social"]] = txt(nota.get("emitente_razao"))
            linha[IDX["Emitente: CNPJ"]] = fmt_cnpj(emit_cnpj)
            linha[IDX["Destinatário: razão social"]] = txt(nota.get("destinatario_razao"))
            linha[IDX["Destinatário: CNPJ"]] = fmt_cnpj(dest_cnpj)
            linha[IDX["Empresa do grupo"]] = empresa
            linha[IDX["Item"]] = item.get("numero_item") or (k + 1)
            linha[IDX["Descrição completa do produto/equipamento"]] = txt(item.get("descricao"))
            linha[IDX["NCM"]] = txt(item.get("ncm"))
            linha[IDX["CFOP"]] = txt(item.get("cfop"))
            linha[IDX["CST ou CSOSN (ICMS)"]] = txt(item.get("cst_csosn"))
            linha[IDX["Quantidade"]] = qtd_linha
            linha[IDX["Unidade"]] = txt(item.get("unidade"))
            linha[IDX["Valor unitário"]] = D(item.get("valor_unitario"))
            linha[IDX["Valor total do item"]] = p_vtot[u]
            linha[IDX["Desconto rateado"]] = p_desc[u]
            linha[IDX["IPI do item"]] = p_ipi[u]
            linha[IDX["Frete rateado"]] = p_frete[u]
            linha[IDX["Seguro rateado"]] = p_seg[u]
            linha[IDX["ICMS-ST do item (ou rateado)"]] = p_st[u]
            linha[IDX["Outras despesas acessórias rateadas"]] = p_out[u]
            linha[IDX["Base de cálculo do ICMS"]] = p_bc[u]
            ai, ap = D(item.get("aliq_icms")), D(item.get("aliq_ipi"))
            linha[IDX["Alíquota de ICMS"]] = ai / 100 if ai is not None else None
            linha[IDX["ICMS destacado"]] = p_icms[u]
            linha[IDX["Alíquota de IPI"]] = ap / 100 if ap is not None else None
            linha[IDX["Número de série"]] = serie_b
            linha[IDX["Chassi"]] = chassi
            linha[IDX["RENAVAM"]] = renavam
            linha[IDX["Placa"]] = placa
            linha[IDX["Valor total da NF"]] = vnf
            linhas_nota.append((linha, obs))
            ctx.meta.append((empresa, txt(item.get("ncm")), custo, p_icms[u] or ZERO, p_ipi[u]))

    # ---- prova por nota ----
    diverge = False
    if not itens:
        ctx.add_pend(P_TOTAL, arq, ref_nf, chave or None, None, "Nota sem itens lidos")
        diverge = True
    elif vnf is None:
        ctx.add_pend(P_TOTAL, arq, ref_nf, chave or None, None, "Valor total da NF ilegível/ausente")
        diverge = True
    elif abs(soma_custo - vnf) > TOLERANCIA:
        ctx.add_pend(P_TOTAL, arq, ref_nf, chave or None, None,
                     f"Soma do custo {soma_custo:.2f} x valor da NF {vnf:.2f} (dif. {soma_custo - vnf:.2f})")
        diverge = True
    for linha, obs in linhas_nota:
        if diverge:
            obs.append("Divergência entre soma do custo e valor total da NF")
        linha[IDX["Observações / pendências"]] = "; ".join(dict.fromkeys(obs)) or None
        ctx.linhas.append(linha)

    ctx.notas.append({
        "chave": chave or None, "numero": numero, "serie": serie,
        "emitente": txt(nota.get("emitente_razao")), "emit_cnpj": fmt_cnpj(emit_cnpj),
        "destinatario": txt(nota.get("destinatario_razao")), "vnf": vnf, "arquivo": arq,
    })


def processar_pasta(pasta: Path, diag_dir):
    ctx = Contexto()
    pdfs = sorted(p for p in pasta.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
    motor = _motor_ocr()
    ocr_ok = motor is not None
    print("OCR: " + {"rapid": "RapidOCR", "tess": "Tesseract"}.get(motor, "NENHUM instalado - PDFs digitalizados irão para Pendências"))
    if not ocr_ok:
        print("Instale o OCR com:  py -m pip install rapidocr-onnxruntime")
    lidos = 0
    for i, p in enumerate(pdfs, 1):
        arq = str(p)
        print(f"[{i}/{len(pdfs)}] {arq}", flush=True)
        try:
            r = extrair_pdf_local(p, ocr_ok, diag_dir)
        except Exception as e:
            ctx.add_pend(P_NAO_PROC, arq, None, None, None, f"{type(e).__name__}: {e}")
            continue
        lidos += 1
        if r.get("arquivo_ilegivel") or not r.get("notas"):
            ctx.add_pend(P_ILEGIVEL, arq, None, None, None, r.get("observacao_arquivo") or "Nenhuma nota identificada")
            continue
        if r.get("observacao_arquivo"):
            ctx.add_pend(P_ILEGIVEL, arq, None, None, None, "Páginas ignoradas: " + r["observacao_arquivo"])
        for nota in r["notas"]:
            if nota.get("nao_suportado"):
                ctx.add_pend(P_NAO_SUP, arq, num_ou_texto(nota.get("numero")), nota.get("chave_acesso"), None,
                             nota.get("observacoes") or "")
                continue
            processar_nota(ctx, arq, nota)
    return ctx, len(pdfs), lidos


# --------------------------------------------------------------------------- #
# EXCEL
# --------------------------------------------------------------------------- #
def _fonte(bold=False):
    return Font(name=ARIAL, size=10, bold=bold, color="000000")


def _escrever(ws, linha, col, valor, kind):
    c = ws.cell(row=linha, column=col)
    if valor is None:
        c.font = _fonte()
        return c
    if kind == "text":
        c.value = str(valor)
        c.number_format = "@"
    elif kind == "date":
        c.value = valor
        c.number_format = FMT_DATE
    elif kind == "money":
        c.value = float(valor)
        c.number_format = FMT_MONEY
    elif kind == "pct":
        c.value = float(valor)
        c.number_format = FMT_PCT
    elif kind == "qty":
        c.value = float(valor)
        c.number_format = "General"
    else:  # int
        c.value = valor
    c.font = _fonte()
    return c


def _cabecalho(ws, titulos, larguras=None):
    for j, t in enumerate(titulos, 1):
        c = ws.cell(row=1, column=j, value=t)
        c.font = _fonte(bold=True)
        if larguras:
            ws.column_dimensions[get_column_letter(j)].width = larguras[j - 1]
    ws.freeze_panes = "A2"


def gerar_excel(ctx: Contexto, destino: Path):
    wb = Workbook()
    inv = wb.active
    inv.title = "Inventário"
    n = len(ctx.linhas)
    ultima = max(n + 1, 2)

    # ---- Inventário ----
    _cabecalho(inv, [c[0] for c in COLUNAS], [c[2] for c in COLUNAS])
    for r, linha in enumerate(ctx.linhas, start=2):
        for j, (valor, (_, kind, _)) in enumerate(zip(linha, COLUNAS), start=1):
            _escrever(inv, r, j, valor, kind)
        # Fórmulas de custo (colunas 28 e 29)
        c28 = inv.cell(row=r, column=28,
                       value=f"=IF({L['qtd']}{r}=0,0,({L['vtotal']}{r}+{L['ipi']}{r}+{L['frete']}{r}+{L['seg']}{r}"
                             f"+{L['st']}{r}+{L['outras']}{r}-{L['desc']}{r})/{L['qtd']}{r})")
        c29 = inv.cell(row=r, column=29, value=f"={L['custo_un']}{r}*{L['qtd']}{r}")
        for c in (c28, c29):
            c.number_format = FMT_MONEY
            c.font = _fonte()
    inv.auto_filter.ref = f"A1:{get_column_letter(len(COLUNAS))}{ultima}"
    # Linha de totais
    lt = ultima + 1
    inv.cell(row=lt, column=1, value="TOTAIS").font = _fonte(bold=True)
    for nome in ["Quantidade", "Valor total do item", "Desconto rateado", "IPI do item", "Frete rateado",
                 "Seguro rateado", "ICMS-ST do item (ou rateado)", "Outras despesas acessórias rateadas",
                 "Valor total de custo do item", "Base de cálculo do ICMS", "ICMS destacado"]:
        j = IDX[nome] + 1
        col = get_column_letter(j)
        c = inv.cell(row=lt, column=j, value=f"=SUM({col}2:{col}{ultima})")
        c.number_format = "General" if nome == "Quantidade" else FMT_MONEY
        c.font = _fonte(bold=True)

    rng = lambda col: f"'Inventário'!${col}$2:${col}${ultima}"

    # ---- Resumo ----
    rs = wb.create_sheet("Resumo")
    for j, w in enumerate([34, 20, 20, 20], 1):
        rs.column_dimensions[get_column_letter(j)].width = w
    heads = ["Custo total", "ICMS destacado", "IPI"]
    cols_soma = [L["custo_tot"], L["icms"], L["ipi"]]

    def bloco(linha_ini, titulo, chaves, col_crit, especial_vazio=None):
        c = rs.cell(row=linha_ini, column=1, value=titulo); c.font = _fonte(bold=True)
        for j, h in enumerate(heads, 2):
            rs.cell(row=linha_ini, column=j, value=h).font = _fonte(bold=True)
        r = linha_ini + 1
        for ch in chaves:
            a = rs.cell(row=r, column=1, value=ch if ch is not None else "(sem NCM)")
            a.number_format = "@"; a.font = _fonte()
            for j, cs in enumerate(cols_soma, 2):
                crit = f"$A{r}" if ch is not None else '"="'
                f = rs.cell(row=r, column=j, value=f"=SUMIFS({rng(cs)},{rng(col_crit)},{crit})")
                f.number_format = FMT_MONEY; f.font = _fonte()
            r += 1
        t = rs.cell(row=r, column=1, value="TOTAL"); t.font = _fonte(bold=True)
        for j in range(2, 5):
            col = get_column_letter(j)
            f = rs.cell(row=r, column=j, value=f"=SUM({col}{linha_ini + 1}:{col}{r - 1})")
            f.number_format = FMT_MONEY; f.font = _fonte(bold=True)
        return r + 2

    prox = bloco(1, "Empresa do grupo", EMPRESAS_RESUMO, L["empresa"])
    ncms = sorted({m[1] for m in ctx.meta if m[1]})
    if any(m[1] is None for m in ctx.meta):
        ncms.append(None)
    bloco(prox, "NCM", ncms, L["ncm"])

    # ---- Conferência por NF ----
    cf = wb.create_sheet("Conferência por NF")
    tit = ["Chave de acesso", "Número da NF", "Série", "Emitente", "Emitente: CNPJ", "Destinatário",
           "Valor total da NF", "Soma do custo dos itens", "Diferença", "Situação"]
    _cabecalho(cf, tit, [48, 12, 8, 36, 20, 36, 16, 18, 14, 14])
    for r, nt in enumerate(ctx.notas, start=2):
        _escrever(cf, r, 1, nt["chave"], "text")
        _escrever(cf, r, 2, nt["numero"], "int")
        _escrever(cf, r, 3, nt["serie"], "int")
        _escrever(cf, r, 4, nt["emitente"], "text")
        _escrever(cf, r, 5, nt["emit_cnpj"], "text")
        _escrever(cf, r, 6, nt["destinatario"], "text")
        _escrever(cf, r, 7, nt["vnf"], "money")
        crit = []
        if nt["numero"] is not None:
            crit.append(f"{rng(L['numero'])},$B{r}")
        if nt["serie"] is not None:
            crit.append(f"{rng(L['serie'])},$C{r}")
        if nt["emit_cnpj"]:
            crit.append(f"{rng(L['emit_cnpj'])},$E{r}")
        if not crit:  # sem identificador: usa o arquivo de origem
            crit.append(f"{rng('A')},\"{nt['arquivo']}\"")
        h = cf.cell(row=r, column=8, value=f"=SUMIFS({rng(L['custo_tot'])},{','.join(crit)})")
        i = cf.cell(row=r, column=9, value=f"=H{r}-G{r}")
        j = cf.cell(row=r, column=10, value=f'=IF(G{r}="","SEM TOTAL LIDO",IF(ABS(I{r})<={TOLERANCIA},"OK","DIVERGENTE"))')
        for c in (h, i):
            c.number_format = FMT_MONEY
        for c in (h, i, j):
            c.font = _fonte()
    cf.auto_filter.ref = f"A1:J{max(len(ctx.notas) + 1, 2)}"

    # ---- Pendências ----
    pd = wb.create_sheet("Pendências")
    _cabecalho(pd, ["Tipo de pendência", "Arquivo de origem", "Número da NF", "Chave de acesso", "Item", "Detalhe"],
               [30, 50, 12, 48, 8, 80])
    ordenadas = sorted(ctx.pend, key=lambda p: (ORDEM_PEND.index(p[0]) if p[0] in ORDEM_PEND else 99, p[1]))
    for r, p in enumerate(ordenadas, start=2):
        for j, v in enumerate(p, start=1):
            kind = "int" if j in (3, 5) and isinstance(v, int) else "text"
            _escrever(pd, r, j, v, kind)
    pd.auto_filter.ref = f"A1:F{max(len(ordenadas) + 1, 2)}"

    wb.calculation.fullCalcOnLoad = True
    wb.save(destino)


# --------------------------------------------------------------------------- #
# PRINCIPAL
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Inventário do imobilizado - Grupo Qualitá")
    ap.add_argument("--pasta", default=None, help="Pasta dos PDFs (se omitida, o robô pergunta)")
    ap.add_argument("--saida", default=".", help="Pasta onde o Excel será gravado")
    args = ap.parse_args()

    pasta_txt = args.pasta
    if not pasta_txt:
        resp = input(f"Qual a pasta dos PDFs? (Enter para usar {PASTA_PADRAO}): ")
        pasta_txt = resp.strip().strip('"').strip("'") or PASTA_PADRAO
    pasta, saida = Path(pasta_txt), Path(args.saida)
    if not pasta.exists():
        sys.exit(f"Pasta não encontrada: {pasta}")
    saida.mkdir(parents=True, exist_ok=True)

    ctx, achados, lidos = processar_pasta(pasta, saida / "_diagnostico")

    try:
        from zoneinfo import ZoneInfo
        agora = datetime.now(ZoneInfo("America/Sao_Paulo"))
    except Exception:  # sem tzdata no Windows: Brasília = UTC-3
        agora = datetime.now(timezone(timedelta(hours=-3)))
    destino = saida / f"Inventario_Imobilizado_Grupo_Qualita_{agora:%d-%m-%y_%Hh%M}.xlsx"
    gerar_excel(ctx, destino)

    # ---- resumo final ----
    print("\n" + "=" * 60)
    print(f"Arquivo gerado: {destino}")
    print(f"PDFs encontrados: {achados} | lidos: {lidos}")
    print(f"Notas lançadas: {len(ctx.notas)} | itens/unidades lançados: {len(ctx.linhas)}")
    print("\nCusto total por empresa:")
    for emp in EMPRESAS_RESUMO:
        v = sum((m[2] for m in ctx.meta if m[0] == emp), ZERO)
        print(f"  {emp:<16} R$ {v:>16,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    print("\nPendências por tipo:")
    for t in ORDEM_PEND:
        print(f"  {t:<46} {sum(1 for p in ctx.pend if p[0] == t)}")
    print(f"  {'TOTAL':<46} {len(ctx.pend)}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            print(f"\nErro: {e.code}")
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário.")
    except Exception:
        print("\nO robô parou por um erro:\n")
        traceback.print_exc()
    finally:
        _pausar()
