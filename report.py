"""
services/report.py  –  Gerador de relatório PDF das placas
"""
import io
import os
import base64
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

AZUL        = colors.HexColor("#242480")
AZUL_CLARO  = colors.HexColor("#eef0fb")
CINZA       = colors.HexColor("#888888")
CINZA_LINHA = colors.HexColor("#f0f2f8")
BRANCO      = colors.white
PRETO       = colors.HexColor("#1a1a1a")


def _estilo_campo(label: str, valor: str) -> Paragraph:
    return Paragraph(
        f"<font color='#888888' size='7.5'>{label}</font><br/>"
        f"<b><font size='9' color='#1a1a1a'>{valor}</font></b>",
        ParagraphStyle("c", fontName="Helvetica", fontSize=9, leading=14)
    )


def gerar_relatorio(placas: list[dict], nome_cliente: str = "", nome_pedido: str = "") -> bytes:
    buffer   = io.BytesIO()
    largura_util = A4[0] - 36*mm

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=5*mm,  bottomMargin=18*mm,
    )

    # ── Estilos ──
    titulo    = ParagraphStyle("titulo",    fontName="Helvetica-Bold", fontSize=20, textColor=AZUL, spaceAfter=5*mm)
    subtitulo = ParagraphStyle("subtitulo", fontName="Helvetica",      fontSize=8.5, textColor=CINZA, spaceAfter=2*mm)
    secao     = ParagraphStyle("secao",     fontName="Helvetica-Bold", fontSize=10,  textColor=BRANCO)
    qtd_st    = ParagraphStyle("qtd",       fontName="Helvetica-Bold", fontSize=9,   textColor=BRANCO, alignment=TA_RIGHT)
    rodape    = ParagraphStyle("rodape",    fontName="Helvetica",      fontSize=7.5, textColor=CINZA,  alignment=TA_CENTER)

    elementos = []

    # ── Header ──
    agora    = datetime.now().strftime("%d/%m/%Y às %H:%M")
    info_txt = "Relatório de Placas"
    if nome_cliente: info_txt += f"  ·  Cliente: <b>{nome_cliente}</b>"
    if nome_pedido:  info_txt += f"  ·  Pedido: <b>{nome_pedido}</b>"
    info_txt += f"  ·  Gerado em {agora}"

    elementos.append(Paragraph("FastPlac", titulo))
    elementos.append(Spacer(1, 1*mm))
    elementos.append(Paragraph(info_txt, subtitulo))
    elementos.append(HRFlowable(width="100%", thickness=1.5, color=AZUL, spaceAfter=5*mm))

    # ── Resumo ──
    total_placas = sum(int(p["dados"].get("Quantidade de Placas") or 1) for p in placas)
    resumo_data  = [
        ["Total de tipos", "Total de placas", "Cliente", "N° do Pedido"],
        [str(len(placas)), str(total_placas), nome_cliente or "—", nome_pedido or "—"],
    ]
    col_w4 = largura_util / 4
    tabela_resumo = Table(resumo_data, colWidths=[col_w4]*4)
    tabela_resumo.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR",     (0, 0), (-1, 0), BRANCO),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0), 8),
        ("BACKGROUND",    (0, 1), (-1, 1), AZUL_CLARO),
        ("FONTNAME",      (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 1), (-1, 1), 10),
        ("TEXTCOLOR",     (0, 1), (-1, 1), AZUL),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#d0d5e8")),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    elementos.append(tabela_resumo)
    elementos.append(Spacer(1, 8*mm))

    # ── Detalhes por placa ──
    IGNORAR = {"Quantidade de Placas"}
    col_w3  = largura_util / 3

    for idx, placa in enumerate(placas):
        tipo  = placa["tipo"]
        dados = placa["dados"]
        qtd   = int(dados.get("Quantidade de Placas") or 1)

        campos = [
            (k, str(v)) for k, v in dados.items()
            if k not in IGNORAR and str(v).strip() not in ("", "None")
        ]

        # Monta linhas em 3 colunas
        linhas = []
        for i in range(0, len(campos), 3):
            trio = campos[i:i+3]
            celulas = [_estilo_campo(label, valor) for label, valor in trio]
            while len(celulas) < 3:
                celulas.append(Paragraph("", ParagraphStyle("vazio")))
            linhas.append(celulas)

        # Cabeçalho da placa
        cab = Table(
            [[Paragraph(f"{idx + 1}.  {tipo}", secao), Paragraph(f"Qtd: {qtd}", qtd_st)]],
            colWidths=[largura_util * 0.78, largura_util * 0.22],
        )
        cab.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), AZUL),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING",   (0, 0), (0, 0), 8),
            ("RIGHTPADDING",  (1, 0), (1, 0), 8),
        ]))

        # Tabela de campos
        tab_campos = Table(linhas, colWidths=[col_w3, col_w3, col_w3])
        row_bgs = [BRANCO if r % 2 == 0 else CINZA_LINHA for r in range(len(linhas))]
        tab_campos.setStyle(TableStyle([
            ("ROWBACKGROUNDS", (0, 0), (-1, -1), row_bgs),
            ("GRID",           (0, 0), (-1, -1), 0.5, colors.HexColor("#e0e4f5")),
            ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",     (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING",  (0, 0), (-1, -1), 8),
            ("LEFTPADDING",    (0, 0), (-1, -1), 8),
        ]))

        # KeepTogether mantém cabeçalho + campos juntos na mesma página
        bloco = KeepTogether([cab, tab_campos, Spacer(1, 7*mm)])
        elementos.append(bloco)

    # ── Rodapé ──
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=CINZA, spaceBefore=2*mm))
    elementos.append(Spacer(1, 2*mm))
    ano = datetime.now().year
    elementos.append(Paragraph(
        f"{ano} © FastPlac by Águia Sistemas  ·  Relatório gerado automaticamente",
        rodape
    ))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.read()