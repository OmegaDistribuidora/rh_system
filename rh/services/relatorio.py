from io import BytesIO
from datetime import datetime, date, timedelta
from collections import defaultdict, Counter

from django.http import HttpResponse
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm

from ..models import Admissao, Desligamento


def add_header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4

    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawString(doc.leftMargin, h - 1*cm, "Relatório TURNOVER - Ômega Distribuidora")
    canvas.line(doc.leftMargin, h - 1.2*cm, w - doc.rightMargin, h - 1.2*cm)

    canvas.setFont("Helvetica", 8)
    canvas.drawString(doc.leftMargin, 1*cm, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    canvas.drawCentredString(w / 2, 1*cm, "Confidencial – Uso Interno")
    canvas.drawRightString(w - doc.rightMargin, 1*cm, f"Página {doc.page}")

    canvas.restoreState()


_styles = getSampleStyleSheet()
H1 = _styles["Heading1"]
H2 = _styles["Heading2"]

BODY9 = ParagraphStyle(
    name="Body9",
    parent=_styles["Normal"],
    fontSize=9,
    leading=11,
    spaceAfter=4,
    textColor=colors.black,
)


def build_table(data, col_widths, header_color):

    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),

        ("ALIGN", (0, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 1), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),

        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#F3F3F3")]),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),

        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def gerar_relatorio_pdf():
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5*cm,
        rightMargin=1.5*cm,
        topMargin=1.7*cm,
        bottomMargin=1.5*cm,
    )
    elements = []

    hoje = date.today()
    inicio_semana = hoje - timedelta(days=hoje.weekday())
    inicio_mes = hoje.replace(day=1)
    inicio_ano = hoje.replace(month=1, day=1)

    adm_qs = Admissao.objects.all().order_by("-data_admissao", "-id")
    deslig_qs = Desligamento.objects.all().order_by("-demissao", "-id")

    codigo_para_supervisor = {}
    for a in adm_qs:
        sup_txt = (a.supervisor_responsavel or "").strip().upper()
        if a.codigo and a.codigo not in codigo_para_supervisor:
            codigo_para_supervisor[a.codigo] = sup_txt

    total_adm_semana = adm_qs.filter(data_admissao__gte=inicio_semana).count()
    total_adm_mes = adm_qs.filter(data_admissao__gte=inicio_mes).count()
    total_adm_ano = adm_qs.filter(data_admissao__gte=inicio_ano).count()

    total_desl_semana = deslig_qs.filter(demissao__gte=inicio_semana).count()
    total_desl_mes = deslig_qs.filter(demissao__gte=inicio_mes).count()
    total_desl_ano = deslig_qs.filter(demissao__gte=inicio_ano).count()

    adm_por_sup_count = Counter()
    for a in adm_qs:
        sup_txt = (a.supervisor_responsavel or "").strip().upper()
        adm_por_sup_count[sup_txt] += 1

    deslig_por_sup_count = Counter()
    for d in deslig_qs:
        sup_txt = (d.supervisor_responsavel or "").strip().upper() if getattr(d, "supervisor_responsavel", None) else ""
        if not sup_txt and d.codigo:
            sup_txt = codigo_para_supervisor.get(d.codigo, "")
        deslig_por_sup_count[sup_txt] += 1

    elements.append(Paragraph("📌 Resumo Executivo", H1))

    resumo = [
        ["Admissões Semana", total_adm_semana],
        ["Admissões Mês", total_adm_mes],
        ["Admissões Ano", total_adm_ano],
        ["Desligamentos Semana", total_desl_semana],
        ["Desligamentos Mês", total_desl_mes],
        ["Desligamentos Ano", total_desl_ano],
    ]

    card_data = []
    linha = []
    for i, (label, value) in enumerate(resumo, 1):
        cell = Table(
            [
                [Paragraph(f"<b>{label}</b>", BODY9)],
                [Paragraph(f"<font size=16><b>{value}</b></font>", BODY9)],
            ],
            colWidths=[170],
            rowHeights=[20, 25],
        )
        cell.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
        ]))
        linha.append(cell)

        if i % 3 == 0:
            card_data.append(linha)
            linha = []
    if linha:
        card_data.append(linha)

    card_table = Table(card_data, colWidths=[170, 170, 170])
    card_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
        ("BOX", (0, 0), (-1, -1), 1, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    elements.append(card_table)
    elements.append(Spacer(1, 25))

    dados_adm = [["Supervisor", "Qtd Admissões"]]
    for sup_txt, qtd in sorted(adm_por_sup_count.items(), key=lambda x: (x[0] or "")):
        if not sup_txt:
            continue
        dados_adm.append([sup_txt, qtd])
    dados_adm.append(["TOTAL", sum(adm_por_sup_count.values())])

    elements.append(build_table(dados_adm, col_widths=[270, 140], header_color=colors.HexColor("#1565C0")))
    elements.append(Spacer(1, 20))

    dados_desl = [["Supervisor", "Qtd Desligamentos"]]
    for sup_txt, qtd in sorted(deslig_por_sup_count.items(), key=lambda x: (x[0] or "")):
        if not sup_txt:
            continue
        dados_desl.append([sup_txt, qtd])
    dados_desl.append(["TOTAL", sum(deslig_por_sup_count.values())])

    elements.append(build_table(dados_desl, col_widths=[270, 140], header_color=colors.HexColor("#C62828")))
    elements.append(PageBreak())

    elements.append(Paragraph("✅ ADMISSÕES", H1))
    elements.append(Spacer(1, 6))

    adm_por_sup = defaultdict(list)
    for a in adm_qs:
        sup_txt = (a.supervisor_responsavel or "").strip().upper()
        if not sup_txt:
            continue
        adm_por_sup[sup_txt].append(a)

    for sup_txt in sorted(adm_por_sup.keys(), key=lambda s: (s or "")):
        elements.append(Paragraph(f"👤 Supervisor: {sup_txt}", H2))
        dados = [["Nome", "Código", "Cargo", "Data", "Status"]]

        for a in adm_por_sup[sup_txt]:
            dados.append([
                Paragraph(a.nome or "—", BODY9),
                a.codigo or "—",
                Paragraph(a.cargo or "—", BODY9),
                a.data_admissao.strftime("%d/%m/%Y") if a.data_admissao else "—",
                a.status or "—",
            ])

        elements.append(build_table(dados, col_widths=[170, 60, 140, 70, 70], header_color=colors.HexColor("#1565C0")))
        elements.append(Spacer(1, 10))

    elements.append(PageBreak())

    elements.append(Paragraph("❌ DESLIGAMENTOS", H1))
    elements.append(Spacer(1, 6))

    deslig_por_sup = defaultdict(list)
    for d in deslig_qs:
        sup_txt = (d.supervisor_responsavel or "").strip().upper() if getattr(d, "supervisor_responsavel", None) else ""
        if not sup_txt and d.codigo:
            sup_txt = codigo_para_supervisor.get(d.codigo, "")
        if not sup_txt:
            continue
        deslig_por_sup[sup_txt].append(d)

    for sup_txt in sorted(deslig_por_sup.keys(), key=lambda s: (s or "")):
        elements.append(Paragraph(f"👤 Supervisor: {sup_txt}", H2))
        dados = [["Nome", "Área", "Data", "Motivo", "Status"]]

        for d in deslig_por_sup[sup_txt]:
            dados.append([
                Paragraph(d.nome or "—", BODY9),
                Paragraph(getattr(d, "area_atuacao", "") or "—", BODY9),
                d.demissao.strftime("%d/%m/%Y") if d.demissao else "—",
                Paragraph(getattr(d, "motivo", "") or "—", BODY9),
                getattr(d, "status", "—") or "—",
            ])

        elements.append(build_table(dados, col_widths=[160, 80, 60, 160, 50], header_color=colors.HexColor("#C62828")))
        elements.append(Spacer(1, 10))

    doc.build(elements, onFirstPage=add_header_footer, onLaterPages=add_header_footer)
    buffer.seek(0)
    return HttpResponse(buffer, content_type="application/pdf")
