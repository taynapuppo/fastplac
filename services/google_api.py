import io
import os
import re
import json
import time
import copy
import html as _html_mod
import zipfile
import tempfile
import requests as http_requests
import google.auth.transport.requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload
from pypdf import PdfWriter, PdfReader
from lxml import etree

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]

_BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
_SECRETS_DIR = os.path.join(_BASE_DIR, "secrets")
TOKEN_FILE   = os.path.join(_SECRETS_DIR, "token.json")
CREDS_FILE   = os.path.join(_SECRETS_DIR, "client_secret.json")


# ─────────────────────────────────────────────────────────────
#  Autenticação
# ─────────────────────────────────────────────────────────────

def get_services():
    creds = None

    try:
        import streamlit as st
        if "google" in st.secrets and "token" in st.secrets["google"]:
            token_info = json.loads(st.secrets["google"]["token"])
            creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    except Exception:
        pass

    if creds is None and os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
        try:
            with open(TOKEN_FILE, "w") as f:
                f.write(creds.to_json())
        except Exception:
            pass

    if not creds or not creds.valid:
        try:
            import streamlit as st
            if "google" in st.secrets and "client_secret" in st.secrets["google"]:
                client_info = json.loads(st.secrets["google"]["client_secret"])
                with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tmp:
                    json.dump(client_info, tmp)
                    tmp_path = tmp.name
                flow = InstalledAppFlow.from_client_secrets_file(tmp_path, SCOPES)
                os.unlink(tmp_path)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
        except Exception:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)

        creds = flow.run_local_server(port=0)
        os.makedirs(_SECRETS_DIR, exist_ok=True)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    drive  = build("drive",  "v3", credentials=creds)
    slides = build("slides", "v1", credentials=creds)
    return drive, slides, creds


# ─────────────────────────────────────────────────────────────
#  Retry para Drive API
# ─────────────────────────────────────────────────────────────

def _execute_with_retry(request, max_retries: int = 7):
    for attempt in range(max_retries):
        try:
            return request.execute()
        except HttpError as e:
            if e.resp.status in (429, 500, 503) and attempt < max_retries - 1:
                wait = min(2 ** (attempt + 2), 60)
                print(f"[RETRY {e.resp.status}] Aguardando {wait}s (tentativa {attempt + 1}/{max_retries})...")
                time.sleep(wait)
            else:
                raise


# ─────────────────────────────────────────────────────────────
#  Download de template como PPTX (com cache por tipo)
# ─────────────────────────────────────────────────────────────

def _download_template_pptx(creds, template_id: str) -> bytes:
    """Baixa um Google Slides como PPTX."""
    if creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
    url  = f"https://docs.google.com/presentation/d/{template_id}/export/pptx"
    resp = http_requests.get(
        url,
        headers={"Authorization": f"Bearer {creds.token}"},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.content


# ─────────────────────────────────────────────────────────────
#  Preenchimento LOCAL de placeholders via XML direto no ZIP
# ─────────────────────────────────────────────────────────────

def _xml_escape(text: str) -> str:
    """Escapa caracteres especiais XML no valor substituído (usado só em fallback)."""
    return _html_mod.escape(str(text), quote=False)


def format_value(key: str, val: str) -> str:
    """Formata o texto de entrada. Para Dimensões, garante espaços ao redor da barra '/' para permitir quebra de linha."""
    val_str = str(val).strip() if val not in (None, "") else ""
    if not val_str:
        return ""
    k = key.lower()
    if k == "dimensões" or k.startswith("peso") or k.startswith("carga"):
        # Formata barras como ' / ' para permitir quebra de linha fluida
        # (espaço não-quebrável antes da barra: a linha nunca começa com '/')
        val_str = re.sub(r'\s*/\s*', '\u00a0/ ', val_str)
    return val_str


# ─────────────────────────────────────────────────────────────
#  Ajuste automático de fonte na faixa BRANCA da placa
# ─────────────────────────────────────────────────────────────
#
#  Como funciona:
#   1. Localiza no slide a faixa branca (retângulo branco largo) e onde
#      a faixa azul começa — é o espaço disponível. Nada no template muda.
#   2. Para cada caixa de texto dentro da faixa branca que recebeu algum
#      {{placeholder}}, mede o texto com as larguras reais da Montserrat Bold
#      e procura o MAIOR tamanho (até o tamanho original do template, 45pt)
#      em que o texto cabe na largura e na altura disponíveis.
#   3. Insere as quebras de linha explicitamente, mantendo sempre o "mm"
#      colado ao último número (nunca sozinho na linha de baixo).
#
#  Títulos ("DIMENSÕES DA CARGA", "CARGA MÁXIMA | NÍVEL"...) ficam no tamanho
#  original; só diminuem se nem no tamanho mínimo o conteúdo couber.
#  A faixa azul, o cliente e o logo nunca são alterados.

# Tipos que ficam EXATAMENTE como no template (só troca os {{campos}}, sem mexer
# em fonte, posição ou quebra de linha). Basta o nome do tipo CONTER o texto abaixo,
# sem diferenciar maiúsculas/acentos — ex.: "mezanino" pega "Placa Mezanino".
TIPOS_SEM_AJUSTE_FONTE = {"flow rack", "mezanino"}

LADO_DIREITO_UNIFORME = True  # título e valores do lado direito da seta com o mesmo tamanho
CENTRALIZAR_VERTICAL = True   # centraliza o texto na altura da faixa branca (False = posição original do template)
# Tipos com dois valores empilhados (um em cima do outro): mantêm a posição do template,
# senão a centralização sobrepõe os valores. Comparação sem diferenciar maiúsculas.
TIPOS_SEM_CENTRALIZAR = {"Placa Mezanino"}

FONTE_MIN_PT = 18      # menor tamanho permitido
PASSO_PT     = 1       # passo da busca
FOLGA_LARGURA = 0.95   # usa 95% da largura (margem de segurança p/ renderização do Google)
LINHA_FATOR   = 1.22   # altura de linha da Montserrat (ascender+descender / em)

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A = f"{{{_NS_A}}}"
_P = f"{{{_NS_P}}}"

# Larguras de avanço da Montserrat Bold (unidades por 1000 em)
_MONT_CHARS = (' !"#$%&\'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`'
               'abcdefghijklmnopqrstuvwxyz{|}~ÀÁÂÃÇÉÊÍÓÔÕÚÜàáâãçéêíóôõúü°ºª²³·–—×\u00a0')
_MONT_W = [283, 289, 437, 720, 638, 877, 728, 230, 357, 358, 434, 599, 262, 386, 262, 392, 679, 392,
           590, 592, 689, 595, 637, 620, 660, 637, 262, 262, 599, 599, 599, 589, 1035, 766, 765, 724,
           826, 671, 639, 771, 808, 328, 541, 740, 604, 955, 808, 844, 732, 844, 735, 638, 618, 788,
           746, 1163, 714, 676, 671, 368, 392, 368, 600, 500, 600, 617, 690, 591, 692, 631, 387, 700,
           691, 301, 307, 655, 301, 1049, 691, 655, 690, 690, 431, 531, 435, 687, 598, 937, 595, 598,
           543, 391, 309, 391, 599, 766, 766, 766, 766, 724, 671, 671, 328, 844, 844, 844, 788, 788,
           617, 617, 617, 617, 591, 631, 631, 301, 655, 655, 655, 687, 687, 418, 427, 412, 430, 430,
           302, 500, 1000, 599, 283]
_CHAR_W = dict(zip(_MONT_CHARS, _MONT_W))
_CHAR_W_PADRAO = 700  # caractere desconhecido: estimativa conservadora

EMU_POR_PT = 12700


def _largura_emu(texto: str, sz_centesimos: int) -> float:
    pt = sz_centesimos / 100.0
    return sum(_CHAR_W.get(c, _CHAR_W_PADRAO) for c in texto) / 1000.0 * pt * EMU_POR_PT


# ── Geometria do slide ──────────────────────────────────────

def _xfrm_de(el):
    return el.find(f"{_P}spPr/{_A}xfrm") if el.tag == f"{_P}sp" else el.find(f"{_P}grpSpPr/{_A}xfrm")


def _coletar_formas(sptree):
    """Retorna [(elemento_sp, x, y, w, h, cor_preenchimento)] em coordenadas absolutas (EMU)."""
    formas = []

    def walk(node, ax, bx, ay, by):
        for ch in node:
            if ch.tag == f"{_P}grpSp":
                xf = _xfrm_de(ch)
                if xf is None:
                    walk(ch, ax, bx, ay, by)
                    continue
                off, ext = xf.find(f"{_A}off"), xf.find(f"{_A}ext")
                choff, chext = xf.find(f"{_A}chOff"), xf.find(f"{_A}chExt")
                ox, oy = int(off.get("x")), int(off.get("y"))
                ex, ey = int(ext.get("cx")), int(ext.get("cy"))
                cx0 = int(choff.get("x")) if choff is not None else 0
                cy0 = int(choff.get("y")) if choff is not None else 0
                cex = int(chext.get("cx")) if chext is not None and int(chext.get("cx")) else ex or 1
                cey = int(chext.get("cy")) if chext is not None and int(chext.get("cy")) else ey or 1
                sx, sy = ex / cex, ey / cey
                # filho: abs = a*(off + (x - chOff)*s) + b
                walk(ch, ax * sx, ax * (ox - cx0 * sx) + bx, ay * sy, ay * (oy - cy0 * sy) + by)
            elif ch.tag == f"{_P}sp":
                xf = _xfrm_de(ch)
                if xf is None:
                    continue
                off, ext = xf.find(f"{_A}off"), xf.find(f"{_A}ext")
                if off is None or ext is None:
                    continue
                x = ax * int(off.get("x")) + bx
                y = ay * int(off.get("y")) + by
                w = ax * int(ext.get("cx"))
                h = ay * int(ext.get("cy"))
                cor = None
                sf = ch.find(f"{_P}spPr/{_A}solidFill/{_A}srgbClr")
                if sf is not None:
                    cor = (sf.get("val") or "").upper()
                formas.append((ch, x, y, w, h, cor))

    walk(sptree, 1.0, 0.0, 1.0, 0.0)
    return formas


def _faixa_branca(formas, slide_w):
    """(topo, base) da faixa branca visível, ou None se não encontrada."""
    largas = [f for f in formas if f[5] and f[3] >= 0.6 * slide_w]
    brancas = [f for f in largas if f[5] == "FFFFFF"]
    if not brancas:
        return None
    b = max(brancas, key=lambda f: f[3] * f[4])
    topo, base = b[2], b[2] + b[4]
    for f in largas:
        if f[5] == "FFFFFF" or f is b:
            continue
        f_top, f_bot = f[2], f[2] + f[4]
        if topo < f_top < base:            # faixa colorida começando por cima da branca (ex.: azul)
            base = min(base, f_top)
        elif f_top <= topo < f_bot < base:  # faixa colorida invadindo o topo
            topo = f_bot
    return topo, base


# ── Layout de texto ────────────────────────────────────────

def _sz_do_run(rpr, padrao):
    if rpr is not None and rpr.get("sz"):
        return int(rpr.get("sz"))
    return padrao


def _info_paragrafo(p):
    """Extrai sequência de itens ('t', texto, sz) / ('br',) e parâmetros de espaçamento."""
    ppr = p.find(f"{_A}pPr")
    ln_pct, ln_pts, bef, aft = 1.0, None, 0.0, 0.0
    if ppr is not None:
        v = ppr.find(f"{_A}lnSpc/{_A}spcPct")
        if v is not None: ln_pct = int(v.get("val")) / 100000.0
        v = ppr.find(f"{_A}lnSpc/{_A}spcPts")
        if v is not None: ln_pts = int(v.get("val")) / 100.0
        v = ppr.find(f"{_A}spcBef/{_A}spcPts")
        if v is not None: bef = int(v.get("val")) / 100.0
        v = ppr.find(f"{_A}spcAft/{_A}spcPts")
        if v is not None: aft = int(v.get("val")) / 100.0
    end = p.find(f"{_A}endParaRPr")
    sz_end = _sz_do_run(end, 1800)
    itens = []
    for ch in p:
        if ch.tag in (f"{_A}r", f"{_A}fld"):
            t = ch.find(f"{_A}t")
            itens.append(("t", (t.text or "") if t is not None else "", _sz_do_run(ch.find(f"{_A}rPr"), sz_end)))
        elif ch.tag == f"{_A}br":
            itens.append(("br",))
    return itens, ln_pct, ln_pts, bef, aft, sz_end


def _quebrar_linhas(itens, larg_max, escala_sz):
    """
    Quebra gulosa por espaços comuns (NBSP não quebra).
    Retorna (linhas, posicoes_quebra) — posicoes são índices no texto
    concatenado do parágrafo onde há um espaço a ser trocado por quebra.
    """
    # achata em caracteres com tamanho
    chars = []  # (char, sz) ou ('\n', None) para <a:br>
    for it in itens:
        if it[0] == "br":
            chars.append(("\n", None))
        else:
            sz = escala_sz(it[2])
            chars.extend((c, sz) for c in it[1])

    linhas, quebras = [], []
    larg_linha, linha_txt = 0.0, ""
    pos = 0           # posição no texto (sem contar <a:br>)
    ult_esp = None    # (pos, larg_ate_espaco, idx_linha_txt)
    for c, sz in chars:
        if c == "\n":
            linhas.append(linha_txt.rstrip(" "))
            larg_linha, linha_txt, ult_esp = 0.0, "", None
            continue
        w = _largura_emu(c, sz)
        if c == " ":
            ult_esp = (pos, larg_linha, len(linha_txt))
        if c != " " and larg_linha + w > larg_max and linha_txt.strip():
            if ult_esp is not None:
                p_esp, _, idx = ult_esp
                quebras.append(p_esp)
                linhas.append(linha_txt[:idx].rstrip(" "))
                resto = linha_txt[idx + 1:]
                linha_txt = resto
                larg_linha = sum(_largura_emu(ch, sz) for ch in resto)  # aproximação (mesmo sz)
            else:
                # palavra única maior que a linha: o renderizador quebra no meio
                linhas.append(linha_txt)
                linha_txt, larg_linha = "", 0.0
            ult_esp = None
        linha_txt += c
        larg_linha += w
        pos += 1
    linhas.append(linha_txt.rstrip(" "))
    return linhas, quebras


def _medir_forma(paragrafos, sz_para, larg_max):
    """
    paragrafos: lista de _info_paragrafo; sz_para: função(idx_par, sz_original)->sz.
    Retorna (altura_total_emu, cabe_na_largura, quebras_por_paragrafo).
    """
    altura, cabe, todas_quebras = 0.0, True, []
    for i, (itens, ln_pct, ln_pts, bef, aft, sz_end) in enumerate(paragrafos):
        f = (lambda s, i=i: sz_para(i, s))
        linhas, quebras = _quebrar_linhas(itens, larg_max, f)
        todas_quebras.append(quebras)
        tams = [f(it[2]) for it in itens if it[0] == "t" and it[1]] or [f(sz_end)]
        maior = max(tams) / 100.0
        for ln in linhas:
            if ln and _largura_emu(ln, max(tams)) > larg_max * 1.001 and " " not in ln.strip():
                cabe = False  # palavra única não cabe
        lh = ln_pts if ln_pts else maior * LINHA_FATOR * ln_pct
        altura += (len(linhas) * lh + bef + aft) * EMU_POR_PT
    return altura, cabe, todas_quebras


def _inserir_quebras(p, posicoes):
    """Troca os espaços nas posições indicadas por <a:br/> (copiando a formatação do run)."""
    if not posicoes:
        return
    alvo = set(posicoes)
    pos = 0
    for r in list(p):
        if r.tag not in (f"{_A}r", f"{_A}fld"):
            continue
        t = r.find(f"{_A}t")
        txt = (t.text or "") if t is not None else ""
        ini = pos
        pos += len(txt)
        if r.tag != f"{_A}r":
            continue
        cortes = sorted(q - ini for q in alvo if ini <= q < pos)
        if not cortes:
            continue
        partes, ant = [], 0
        for c in cortes:
            partes.append(txt[ant:c])
            ant = c + 1  # remove o espaço
        partes.append(txt[ant:])
        rpr = r.find(f"{_A}rPr")
        pai, idx = r.getparent(), r.getparent().index(r)
        t.text = partes[0]
        for parte in partes[1:]:
            br = etree.Element(f"{_A}br")
            if rpr is not None:
                br.append(copy.deepcopy(rpr))
            novo = copy.deepcopy(r)
            novo.find(f"{_A}t").text = parte
            idx += 1; pai.insert(idx, br)
            idx += 1; pai.insert(idx, novo)


def _aplicar_sz(p, sz, sz_padrao, forcar=False):
    """Limita o tamanho dos runs do parágrafo a `sz` (forcar=True: usa exatamente `sz`)."""
    for r in p.findall(f"{_A}r"):
        if r.find(f"{_A}rPr") is None:
            r.insert(0, etree.Element(f"{_A}rPr"))
    for el in p.iter(f"{_A}rPr", f"{_A}endParaRPr"):
        el.set("sz", str(sz if forcar else min(_sz_do_run(el, sz_padrao), sz)))


def _ajustar_forma(sp, x, y, w, h, faixa, slide_w, paras_alterados, permitir_centralizar=True):
    body = sp.find(f"{_P}txBody")
    if body is None:
        return
    bpr = body.find(f"{_A}bodyPr")
    g = (lambda k, d: int(bpr.get(k)) if bpr is not None and bpr.get(k) is not None else d)
    lins, rins, tins, bins = g("lIns", 91440), g("rIns", 91440), g("tIns", 45720), g("bIns", 45720)
    anchor = (bpr.get("anchor") if bpr is not None else None) or "t"
    sem_quebra = bpr is not None and bpr.get("wrap") == "none"

    topo, base = faixa
    margem_y = 0.04 * (base - topo)
    margem_x = 0.02 * slide_w

    x0 = x + lins
    x1 = min(x + w - rins, slide_w - margem_x)   # caixas que "vazam" do slide são limitadas à borda
    larg = (x1 - x0) * FOLGA_LARGURA

    # Centralização vertical: a caixa passa a ocupar toda a faixa branca (com margem)
    # e o texto fica ancorado no meio. Só para caixas fora de grupo (coordenadas diretas).
    centralizar = CENTRALIZAR_VERTICAL and permitir_centralizar and bpr is not None and sp.getparent() is not None and sp.getparent().tag == f"{_P}spTree"
    if centralizar:
        novo_y = topo + margem_y
        novo_h = (base - topo) - 2 * margem_y
        alt = novo_h - tins - bins
    elif anchor == "ctr":
        c = (y + tins + y + h - bins) / 2
        alt = 2 * min(c - (topo + margem_y), (base - margem_y) - c)
    elif anchor == "b":
        alt = (y + h - bins) - (topo + margem_y)
    else:
        alt = (base - margem_y) - (y + tins)
    if larg <= 0 or alt <= 0:
        return

    pars = body.findall(f"{_A}p")
    # Parágrafos vazios no fim da caixa só ocupam altura (e tiram o texto do centro):
    # são removidos, mantendo sempre pelo menos um parágrafo.
    while len(pars) > 1 and not "".join(t.text or "" for t in pars[-1].iter(f"{_A}t")).strip() \
            and pars[-1].find(f"{_A}br") is None:
        body.remove(pars[-1]); pars.pop()
    infos = [_info_paragrafo(p) for p in pars]

    # Parágrafos variáveis: os que receberam placeholder + legenda (F x P x A)
    variaveis = set()
    for i, p in enumerate(pars):
        txt = "".join(t.text or "" for t in p.iter(f"{_A}t")).upper()
        if p in paras_alterados or re.search(r'F\s*X\s*P\s*X\s*A', txt):
            variaveis.add(i)
    if not variaveis:
        return

    # Lado DIREITO da seta: título ("CARGA MÁXIMA | NÍVEL"...) e valores ficam todos
    # do MESMO tamanho, o maior que couber (até o maior tamanho da caixa no template).
    uniforme = LADO_DIREITO_UNIFORME and x0 > slide_w * 0.4
    if uniforme:
        variaveis = {i for i, inf in enumerate(infos) if any(it[0] == "t" and it[1].strip() for it in inf[0])} or variaveis

    larg_busca = float("inf") if sem_quebra else larg
    sz_max = max(max([it[2] for it in infos[i][0] if it[0] == "t"] or [infos[i][5]]) for i in variaveis)
    if uniforme:
        sz_max = max(sz_max, max(max([it[2] for it in inf[0] if it[0] == "t"] or [inf[5]]) for inf in infos))
    sz_min = FONTE_MIN_PT * 100

    def testar(sz_var, fator_fixo=1.0):
        def f(i, s):
            if i in variaveis:
                return sz_var if uniforme else min(s, sz_var)
            return int(round(s * fator_fixo / 50.0) * 50)
        altura, cabe, quebras = _medir_forma(infos, f, larg_busca)
        return (cabe and altura <= alt), quebras, f

    escolha = None
    # 1) títulos fixos, conteúdo variável
    sz = sz_max
    while sz >= sz_min:
        ok, quebras, f = testar(sz)
        if ok:
            escolha = (sz, 1.0, quebras); break
        sz -= PASSO_PT * 100
    # 2) não coube nem no mínimo: reduz os títulos junto
    if escolha is None:
        fator = 0.95
        while fator >= 0.5:
            ok, quebras, f = testar(sz_min, fator)
            if ok:
                escolha = (sz_min, fator, quebras); break
            fator -= 0.05
    if escolha is None:
        _, quebras, _ = testar(sz_min, 0.5)
        escolha = (sz_min, 0.5, quebras)

    sz_var, fator, quebras = escolha

    if centralizar:
        xf = sp.find(f"{_P}spPr/{_A}xfrm")
        xf.find(f"{_A}off").set("y", str(int(round(novo_y))))
        xf.find(f"{_A}ext").set("cy", str(int(round(novo_h))))
        bpr.set("anchor", "ctr")
        for fit in (f"{_A}spAutoFit", f"{_A}normAutofit", f"{_A}noAutofit"):
            el = bpr.find(fit)
            if el is not None:
                bpr.remove(el)
        bpr.append(etree.Element(f"{_A}noAutofit"))
    for i, p in enumerate(pars):
        if i in variaveis:
            _aplicar_sz(p, sz_var, infos[i][5], forcar=uniforme)
            if not sem_quebra:
                _inserir_quebras(p, quebras[i])
        elif fator < 1.0:
            for el in p.iter(f"{_A}rPr", f"{_A}endParaRPr"):
                if el.get("sz"):
                    el.set("sz", str(int(round(int(el.get("sz")) * fator / 50.0) * 50)))


# ── Substituição de placeholders ───────────────────────────

# unidades que ficam sempre coladas ao número anterior (nunca sozinhas na linha)
_UNIDADES = r'(mm|kg|t)'
_RE_MM = re.compile(r'(\d)[ \t]+' + _UNIDADES + r'\b', re.IGNORECASE)


def _substituir_paragrafo(p, replacements) -> bool:
    """Substitui {{chave}} nos runs de um parágrafo preservando a formatação. Retorna True se alterou."""
    ts = [t for t in p.iter(f"{_A}t")]
    if not ts:
        return False
    texts = [t.text or "" for t in ts]
    original = list(texts)

    for key, value in replacements.items():
        rgx   = re.compile(re.escape(f'{{{{{key}}}}}'), re.IGNORECASE)
        spans = [m.span() for m in rgx.finditer(''.join(texts))]
        if not spans:
            continue
        rep = str(value).strip() if value not in (None, "") else ""

        starts, pos = [], 0
        for t in texts:
            starts.append(pos); pos += len(t)

        def _run_of(off):
            r = 0
            for i, s in enumerate(starts):
                if off >= s: r = i
                else: break
            return r

        for a, b in reversed(spans):
            ri, rj = _run_of(a), _run_of(b - 1)
            oa, ob = a - starts[ri], b - starts[rj]
            if ri == rj:
                texts[ri] = texts[ri][:oa] + rep + texts[ri][ob:]
            else:
                texts[ri] = texts[ri][:oa] + rep
                for k in range(ri + 1, rj): texts[k] = ''
                texts[rj] = texts[rj][ob:]

    if texts == original:
        return False

    # "mm" sempre colado ao número (espaço não-quebrável), inclusive entre runs diferentes
    for i in range(len(texts)):
        texts[i] = _RE_MM.sub('\\1\u00a0\\2', texts[i])
        m = re.match(r'^[ \t]+' + _UNIDADES + r'\b', texts[i], re.IGNORECASE)
        if m:
            ant = next((texts[j] for j in range(i - 1, -1, -1) if texts[j]), "")
            if ant[-1:].isdigit():
                texts[i] = '\u00a0' + texts[i].lstrip(" \t")

    for t, txt in zip(ts, texts):
        t.text = txt
    return True


def _normalizar_tipo(tipo) -> str:
    """Compara nomes de tipo ignorando maiúsculas, acentos e espaços extras."""
    import unicodedata
    t = unicodedata.normalize("NFKD", str(tipo or "")).encode("ascii", "ignore").decode()
    return " ".join(t.lower().split())


def _processar_slide(xml_bytes: bytes, replacements: dict, slide_w: int, ajustar: bool, centralizar: bool = True) -> bytes:
    root = etree.fromstring(xml_bytes)
    alterados = set()
    for p in root.iter(f"{_A}p"):
        if _substituir_paragrafo(p, replacements):
            alterados.add(p)

    if ajustar and alterados:
        sptree = root.find(f"{_P}cSld/{_P}spTree")
        formas = _coletar_formas(sptree)
        faixa  = _faixa_branca(formas, slide_w)
        if faixa is None:
            print("[AVISO] Faixa branca não encontrada no template — fonte não ajustada.")
        else:
            topo, base = faixa
            alvos = [f for f in formas
                     if topo <= f[2] + f[4] / 2 <= base
                     and any(p in alterados for p in f[0].iter(f"{_A}p"))]
            for sp, x, y, w, h, _ in alvos:
                # Caixas empilhadas (outra caixa alvo na mesma coluna, ex.: Mezanino com
                # m² e plano) nunca são centralizadas — senão uma fica em cima da outra.
                empilhada = any(
                    o[0] is not sp and min(x + w, o[1] + o[3]) - max(x, o[1]) > 0.5 * min(w, o[3])
                    for o in alvos
                )
                _ajustar_forma(sp, x, y, w, h, faixa, slide_w, alterados,
                               permitir_centralizar=centralizar and not empilhada)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _fill_pptx_placeholders(pptx_bytes: bytes, data: dict, tipo: str = "") -> bytes:
    replacements = {k: (format_value(k, str(v)) if v not in (None, "") else "") for k, v in data.items()}
    tipo_norm = _normalizar_tipo(tipo)
    ajustar = not any(_normalizar_tipo(t) in tipo_norm for t in TIPOS_SEM_AJUSTE_FONTE)
    print(f"[FastPlac] tipo={tipo!r} -> ajuste de fonte/posição: {'SIM' if ajustar else 'NÃO (igual ao template)'}")
    centralizar = _normalizar_tipo(tipo) not in {_normalizar_tipo(t) for t in TIPOS_SEM_CENTRALIZAR}

    src_buf = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(src_buf, 'r') as src_zip:
        slide_w = 17995900
        try:
            pres = etree.fromstring(src_zip.read("ppt/presentation.xml"))
            sld = pres.find(f"{_P}sldSz")
            if sld is not None:
                slide_w = int(sld.get("cx"))
        except Exception:
            pass

        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for name in src_zip.namelist():
                raw = src_zip.read(name)
                if re.match(r'ppt/slides/slide\d+\.xml$', name):
                    raw = _processar_slide(raw, replacements, slide_w, ajustar, centralizar)
                out_zip.writestr(name, raw)

    out_buf.seek(0)
    return out_buf.read()


# ─────────────────────────────────────────────────────────────
#  Duplicação LOCAL do primeiro slide (sem API)
# ─────────────────────────────────────────────────────────────

def _duplicate_first_slide_pptx(pptx_bytes: bytes, extra_copies: int) -> bytes:
    """
    Duplica o primeiro slide N vezes dentro de um PPTX, operando
    diretamente no ZIP/XML sem chamar nenhuma API.
    """
    if extra_copies <= 0:
        return pptx_bytes

    NS_PPTX   = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_REL    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_CT     = "http://schemas.openxmlformats.org/package/2006/content-types"
    REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    CT_SLIDE  = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    src_buf = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(src_buf, "r") as src_zip:
        names = src_zip.namelist()

        slide_names = sorted(
            [n for n in names if re.match(r"ppt/slides/slide[0-9]+\.xml$", n)],
            key=lambda x: int(re.search(r"[0-9]+", x).group()),
        )
        if not slide_names:
            return pptx_bytes

        first_slide     = slide_names[0]
        first_slide_num = int(re.search(r"[0-9]+", first_slide).group())
        max_slide_num   = max(int(re.search(r"[0-9]+", n).group()) for n in slide_names)

        pres_root = etree.fromstring(src_zip.read("ppt/presentation.xml"))
        rels_root = etree.fromstring(src_zip.read("ppt/_rels/presentation.xml.rels"))
        ct_root   = etree.fromstring(src_zip.read("[Content_Types].xml"))

        sldIdLst = pres_root.find(f"{{{NS_PPTX}}}sldIdLst")
        max_id   = max(
            (int(e.get("id", 256)) for e in sldIdLst.findall(f"{{{NS_PPTX}}}sldId")),
            default=256,
        )
        max_rel_id = max(
            (int(re.sub(r"\D", "", e.get("Id", "0")) or 0) for e in rels_root),
            default=100,
        )

        extra_files = {}
        for i in range(extra_copies):
            new_num      = max_slide_num + i + 1
            new_path     = f"ppt/slides/slide{new_num}.xml"
            new_rel_path = f"ppt/slides/_rels/slide{new_num}.xml.rels"
            rel_id       = f"rId{max_rel_id + i + 1}"

            extra_files[new_path] = src_zip.read(first_slide)

            first_rel = f"ppt/slides/_rels/slide{first_slide_num}.xml.rels"
            if first_rel in names:
                extra_files[new_rel_path] = src_zip.read(first_rel)

            max_id += 1
            el = etree.SubElement(sldIdLst, f"{{{NS_PPTX}}}sldId")
            el.set("id", str(max_id))
            el.set(f"{{{NS_REL}}}id", rel_id)

            rel_el = etree.SubElement(rels_root, "Relationship")
            rel_el.set("Id", rel_id)
            rel_el.set("Type", REL_SLIDE)
            rel_el.set("Target", f"slides/slide{new_num}.xml")

            ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
            ov.set("PartName", f"/{new_path}")
            ov.set("ContentType", CT_SLIDE)

        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for name in names:
                if name == "ppt/presentation.xml":
                    out_zip.writestr(name, etree.tostring(pres_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                elif name == "ppt/_rels/presentation.xml.rels":
                    out_zip.writestr(name, etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                elif name == "[Content_Types].xml":
                    out_zip.writestr(name, etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                else:
                    out_zip.writestr(name, src_zip.read(name))

            for name, data in extra_files.items():
                out_zip.writestr(name, data)

    out_buf.seek(0)
    return out_buf.read()


# ─────────────────────────────────────────────────────────────
#  Mesclagem LOCAL de múltiplos PPTX
# ─────────────────────────────────────────────────────────────

def _merge_pptx(pptx_bytes_list: list[bytes]) -> bytes:
    """
    Mescla múltiplos PPTX trabalhando direto no ZIP/XML.
    Preserva slides, mídias e relacionamentos de cada arquivo.
    """
    NS_PPTX   = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_REL    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_CT     = "http://schemas.openxmlformats.org/package/2006/content-types"
    REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    CT_SLIDE  = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    zips       = [zipfile.ZipFile(io.BytesIO(b), "r") for b in pptx_bytes_list]
    base_zip   = zips[0]
    base_names = set(base_zip.namelist())
    out_buf    = io.BytesIO()

    slide_count = len([n for n in base_names if re.match(r"ppt/slides/slide[0-9]+\.xml$", n)])
    media_names = {n for n in base_names if n.startswith("ppt/media/")}
    extra_files = {}
    new_pairs   = []

    # Cada template numera as imagens do seu jeito (no Porta Paletes o QR do site é
    # image1.png, no Mezanino image1.png é o QR "Fale conosco"...). Antes, imagens com
    # o mesmo nome eram consideradas iguais e os slides dos outros templates passavam
    # a apontar para a imagem errada — os QR codes trocavam de lugar/tamanho.
    # Agora as imagens são comparadas pelo CONTEÚDO e renomeadas quando necessário.
    import hashlib
    hash_para_nome = {hashlib.md5(base_zip.read(n)).hexdigest(): n for n in media_names}
    contador_media = [0]

    def _nome_unico(ext):
        while True:
            contador_media[0] += 1
            nome = f"ppt/media/fp_img{contador_media[0]}{ext}"
            if nome not in media_names:
                return nome

    for src_zip in zips[1:]:
        src_names  = set(src_zip.namelist())
        src_slides = sorted(
            [n for n in src_names if re.match(r"ppt/slides/slide[0-9]+\.xml$", n)],
            key=lambda x: int(re.search(r"[0-9]+", x).group()),
        )
        mapa_media = {}   # "imageX.png" deste template -> nome final no arquivo mesclado
        for name in sorted(src_names):
            if not name.startswith("ppt/media/"):
                continue
            dados = src_zip.read(name)
            h = hashlib.md5(dados).hexdigest()
            if h not in hash_para_nome:
                final = name if name not in media_names else _nome_unico(os.path.splitext(name)[1])
                extra_files[final] = dados
                media_names.add(final)
                hash_para_nome[h] = final
            mapa_media[name.split("/")[-1]] = hash_para_nome[h].split("/")[-1]
        for slide_path in src_slides:
            slide_count += 1
            new_path = f"ppt/slides/slide{slide_count}.xml"
            extra_files[new_path] = src_zip.read(slide_path)
            rel_src = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            rel_dst = new_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            if rel_src in src_names:
                rels_xml = src_zip.read(rel_src).decode("utf-8")
                rels_xml = re.sub(
                    r'Target="\.\./media/([^"]+)"',
                    lambda m: f'Target="../media/{mapa_media.get(m.group(1), m.group(1))}"',
                    rels_xml,
                )
                extra_files[rel_dst] = rels_xml.encode("utf-8")
            new_pairs.append((new_path, f"rId{100 + slide_count}"))

    pres_root = etree.fromstring(base_zip.read("ppt/presentation.xml"))
    sldIdLst  = pres_root.find(f"{{{NS_PPTX}}}sldIdLst")
    max_id    = max(
        (int(e.get("id", 256)) for e in sldIdLst.findall(f"{{{NS_PPTX}}}sldId")),
        default=256,
    )
    for _, rel_id in new_pairs:
        max_id += 1
        el = etree.SubElement(sldIdLst, f"{{{NS_PPTX}}}sldId")
        el.set("id", str(max_id))
        el.set(f"{{{NS_REL}}}id", rel_id)

    rels_root = etree.fromstring(base_zip.read("ppt/_rels/presentation.xml.rels"))
    for slide_path, rel_id in new_pairs:
        el = etree.SubElement(rels_root, "Relationship")
        el.set("Id", rel_id)
        el.set("Type", REL_SLIDE)
        el.set("Target", slide_path.replace("ppt/", ""))

    ct_root = etree.fromstring(base_zip.read("[Content_Types].xml"))
    # garante o tipo de conteúdo para extensões de imagem vindas de outros templates
    exts = {e.get("Extension", "").lower() for e in ct_root.findall(f"{{{NS_CT}}}Default")}
    tipos_img = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
                 "svg": "image/svg+xml", "emf": "image/x-emf", "wmf": "image/x-wmf"}
    for n in extra_files:
        ext = os.path.splitext(n)[1].lstrip(".").lower()
        if n.startswith("ppt/media/") and ext and ext not in exts and ext in tipos_img:
            d = etree.SubElement(ct_root, f"{{{NS_CT}}}Default")
            d.set("Extension", ext); d.set("ContentType", tipos_img[ext])
            exts.add(ext)
    for slide_path, _ in new_pairs:
        ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
        ov.set("PartName", f"/{slide_path}")
        ov.set("ContentType", CT_SLIDE)

    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for name in base_zip.namelist():
            if name == "ppt/presentation.xml":
                out_zip.writestr(name, etree.tostring(pres_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "ppt/_rels/presentation.xml.rels":
                out_zip.writestr(name, etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "[Content_Types].xml":
                out_zip.writestr(name, etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            else:
                out_zip.writestr(name, base_zip.read(name))
        for name, data in extra_files.items():
            out_zip.writestr(name, data)

    for z in zips:
        z.close()
    out_buf.seek(0)
    return out_buf.read()


# ─────────────────────────────────────────────────────────────
#  Drive helpers
# ─────────────────────────────────────────────────────────────

def rename_file(drive, file_id: str, new_name: str) -> str:
    result = _execute_with_retry(
        drive.files().update(
            fileId=file_id,
            body={"name": new_name},
            fields="id, webViewLink",
        )
    )
    return result.get("webViewLink", "")


def delete_file(drive, file_id: str):
    try:
        drive.files().delete(fileId=file_id).execute()
    except HttpError as e:
        print(f"[WARN] Não foi possível deletar {file_id}: {e}")


def upload_pdf(drive, pdf_bytes: bytes, nome: str, folder_id: str) -> str:
    metadata = {"name": f"{nome}.pdf", "parents": [folder_id], "mimeType": "application/pdf"}
    media    = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf")
    arquivo  = _execute_with_retry(
        drive.files().create(body=metadata, media_body=media, fields="id, webViewLink")
    )
    return arquivo.get("webViewLink", "")


def export_as_pdf(creds, presentation_id: str) -> bytes:
    if creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
    url  = f"https://docs.google.com/presentation/d/{presentation_id}/export/pdf"
    resp = http_requests.get(
        url,
        headers={"Authorization": f"Bearer {creds.token}"},
        timeout=120,
    )
    resp.raise_for_status()
    return resp.content


def merge_pdfs(pdf_bytes_list: list[bytes]) -> bytes:
    writer = PdfWriter()
    for pdf_bytes in pdf_bytes_list:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ─────────────────────────────────────────────────────────────
#  Função principal
# ─────────────────────────────────────────────────────────────

def gerar_pdf_consolidado(
    placas: list[dict],
    folder_id: str,
    template_ids: dict,
    nome_arquivo: str = "Placas",
    progress_callback=None,
) -> tuple[bytes, str, list[dict], str]:
    """
    1. Baixa cada template UMA vez por tipo (cache em memória)
    2. Preenche placeholders diretamente no XML do ZIP (preserva imagens/QR codes)
    3. Duplica slides localmente se Qtd > 1
    4. Mescla todos os PPTX localmente
    5. Faz 1 upload do PPTX mesclado → exporta como PDF
    6. Mantém o PPTX mesclado como Slides consolidado no Drive
    7. Faz 1 upload do PDF final
    """
    drive, _, creds = get_services()
    total = len(placas)

    # ── 1. Download dos templates (por tipo, com cache) ──
    if progress_callback:
        progress_callback(0.02, "Baixando templates...")

    tipos_unicos: list[str]       = list({p["tipo"] for p in placas})
    template_cache: dict[str, bytes] = {}

    for i, tipo in enumerate(tipos_unicos):
        if progress_callback:
            progress_callback(
                0.02 + (i / len(tipos_unicos)) * 0.15,
                f"Baixando template {i + 1}/{len(tipos_unicos)}: {tipo}",
            )
        template_cache[tipo] = _download_template_pptx(creds, template_ids[tipo])
        time.sleep(0.3)

    # ── 2 + 3. Preenche e duplica cada placa localmente ──
    filled_pptx_list: list[bytes] = []
    slides_info:      list[dict]  = []

    for idx, placa in enumerate(placas):
        tipo  = placa["tipo"]
        dados = placa["dados"]
        qtd   = max(1, int(dados.get("Quantidade de Placas") or 1))

        if progress_callback:
            pct = 0.17 + (idx / total) * 0.55
            progress_callback(pct, f"Preenchendo placa {idx + 1}/{total}: {tipo}")

        filled = _fill_pptx_placeholders(template_cache[tipo], dados, tipo=tipo)

        if qtd > 1:
            filled = _duplicate_first_slide_pptx(filled, qtd - 1)

        filled_pptx_list.append(filled)

        cliente = dados.get("Cliente", "")
        slides_info.append({"tipo": tipo, "cliente": cliente, "link": ""})

    # ── 4. Mescla todos os PPTX localmente ──
    if progress_callback:
        progress_callback(0.73, f"Mesclando {total} apresentações...")

    merged_pptx = _merge_pptx(filled_pptx_list)

    # ── 5. Faz upload do PPTX mesclado como Google Slides ──
    if progress_callback:
        progress_callback(0.80, "Enviando para o Google Drive...")

    nome_slides = f"Slides - {nome_arquivo}"
    metadata = {
        "name":     nome_slides,
        "parents":  [folder_id],
        "mimeType": "application/vnd.google-apps.presentation",
    }
    media = MediaIoBaseUpload(
        io.BytesIO(merged_pptx),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        resumable=True,
    )
    result = _execute_with_retry(
        drive.files().create(body=metadata, media_body=media, fields="id, webViewLink")
    )
    slides_id   = result["id"]
    link_slides = result.get("webViewLink", "")

    # ── 6. Exporta como PDF a partir do Slides consolidado ──
    if progress_callback:
        progress_callback(0.90, "Exportando PDF...")

    time.sleep(3)
    pdf_bytes = export_as_pdf(creds, slides_id)

    # ── 7. Faz upload do PDF ──
    if progress_callback:
        progress_callback(0.96, "Salvando PDF na pasta de concluídos...")

    link_pdf = upload_pdf(drive, pdf_bytes, nome_arquivo, folder_id)

    if progress_callback:
        progress_callback(1.0, "Concluído!")

    return pdf_bytes, link_pdf, slides_info, link_slides