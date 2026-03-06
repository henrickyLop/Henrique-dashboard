"""
Gerador de Dashboard PDF — Concrelongo
Recebe o dicionário do parser e gera o PDF com gráficos e tabelas.
"""

import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from collections import defaultdict

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Spacer, Table, TableStyle,
    Paragraph, Image as RLImage, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Paleta ──
BG     = '#07090f'
S1     = '#0e1119'
S2     = '#131720'
GOLD   = '#f0c040'
GREEN  = '#3be8b0'
BLUE   = '#5c9eff'
PURPLE = '#c084fc'
ORANGE = '#fb923c'
RED    = '#ff5c5c'
MUT    = '#5a6278'
TXT    = '#dde3ee'
CORES  = [GOLD, GREEN, BLUE, PURPLE, ORANGE, RED, '#38bdf8', '#a3e635',
          '#f472b6', '#34d399', '#fbbf24', '#60a5fa']


def hm2min(h: str) -> int:
    hh, mm = map(int, h.split(':'))
    return hh * 60 + mm


def ps(name, **kw):
    return ParagraphStyle(name, **kw)


def fig_to_rl_image(fig, avail_w_cm, dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=dpi, bbox_inches='tight',
                facecolor=S1, edgecolor='none')
    plt.close(fig)
    buf.seek(0)
    img = RLImage(buf)
    ow, oh = img.imageWidth, img.imageHeight
    img.drawWidth  = avail_w_cm * cm
    img.drawHeight = (oh / ow) * avail_w_cm * cm
    return img


def style_ax(ax, title=''):
    ax.set_facecolor(S1)
    ax.tick_params(colors=MUT, labelsize=7)
    for sp in ['top', 'right']:
        ax.spines[sp].set_visible(False)
    for sp in ['bottom', 'left']:
        ax.spines[sp].set_color('#1e2535')
    if title:
        ax.set_title(title, color=TXT, fontsize=8, fontweight='bold', pad=6)
    ax.grid(axis='y', color='#1e2535', linewidth=0.5, zorder=0)
    ax.grid(axis='x', visible=False)


# ════════════════════════════════════════
# GRÁFICOS
# ════════════════════════════════════════

def chart_por_dia(por_dia, avail_w):
    dias = sorted(por_dia.keys(), key=lambda d: (d[6:], d[3:5], d[:2]))
    vols = [sum(r.volume    for r in por_dia[d]) for d in dias]
    recs = [sum(r.vlr_total for r in por_dia[d]) for d in dias]

    fig, ax1 = plt.subplots(figsize=(10, 3.8), facecolor=S1)
    ax1.set_facecolor(S1)
    x = np.arange(len(dias))
    n = len(dias)
    bar_cols = [CORES[i % len(CORES)] for i in range(n)]
    ax1.bar(x, vols, color=[c + '66' for c in bar_cols],
            edgecolor=bar_cols, linewidth=1.5, width=0.55, zorder=3)
    ax1.set_xticks(x)
    ax1.set_xticklabels(dias, color=MUT, fontsize=7.5)
    ax1.set_ylabel('Volume (m³)', color=MUT, fontsize=7)
    for sp in ['top', 'right']: ax1.spines[sp].set_visible(False)
    for sp in ['bottom', 'left']: ax1.spines[sp].set_color('#1e2535')
    ax1.grid(axis='y', color='#1e2535', linewidth=0.5, zorder=0)
    ax1.tick_params(colors=MUT, labelsize=7.5)

    ax2 = ax1.twinx()
    ax2.set_facecolor(S1)
    ax2.plot(x, [r / 1000 for r in recs], color=GOLD, marker='o',
             markersize=5, linewidth=2, zorder=4)
    ax2.set_ylabel('Receita (R$ mil)', color=GOLD, fontsize=7)
    ax2.tick_params(colors=GOLD, labelsize=7)
    for sp in ['top', 'right', 'left', 'bottom']:
        ax2.spines[sp].set_visible(False)

    for i, v in enumerate(vols):
        ax1.text(i, v + max(vols) * 0.01, f'{v:.0f}m³',
                 ha='center', va='bottom', color=TXT, fontsize=7, fontweight='bold')

    ax1.set_title('Volume (m³) e Receita por Dia', color=TXT, fontsize=9, fontweight='bold', pad=8)
    fig.patch.set_facecolor(S1)
    plt.tight_layout(pad=0.8)
    return fig_to_rl_image(fig, avail_w)


def chart_clientes(por_cliente, avail_w):
    cli_vol = sorted(
        [(c, sum(r.volume for r in v)) for c, v in por_cliente.items()],
        key=lambda x: -x[1]
    )[:12]
    nomes = [c[:22] for c, _ in cli_vol]
    vols  = [v for _, v in cli_vol]

    fig, ax = plt.subplots(figsize=(10, max(3.5, len(nomes) * 0.42)), facecolor=S1)
    ax.set_facecolor(S1)
    y = np.arange(len(nomes))
    bar_cols = [CORES[i % len(CORES)] for i in range(len(nomes))]
    ax.barh(y, vols,
            color=[c + '55' for c in bar_cols],
            edgecolor=bar_cols, linewidth=1.5, height=0.55)
    ax.set_yticks(y)
    ax.set_yticklabels(nomes, color=TXT, fontsize=7.5)
    ax.invert_yaxis()
    for i, v in enumerate(vols):
        ax.text(v + max(vols) * 0.01, i, f'{v:.1f} m³',
                va='center', color=TXT, fontsize=7.5, fontweight='bold')
    style_ax(ax, 'Volume por Cliente (m³)')
    plt.tight_layout(pad=0.8)
    return fig_to_rl_image(fig, avail_w)


def chart_equipamentos(por_equip, avail_w):
    eq_data = sorted(
        [(e, sum(r.volume for r in v), len(v)) for e, v in por_equip.items()],
        key=lambda x: -x[1]
    )
    nomes = [e for e, *_ in eq_data]
    vols  = [v for _, v, _ in eq_data]
    viags = [n for _, _, n in eq_data]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, max(4, len(nomes) * 0.45)), facecolor=S1)
    bar_cols = [CORES[i % len(CORES)] for i in range(len(nomes))]
    y = np.arange(len(nomes))

    for ax, data, title, unit in [
        (ax1, vols,  'Volume Total (m³)',  'm³'),
        (ax2, viags, 'Nº de Viagens',      'x'),
    ]:
        ax.set_facecolor(S1)
        col = bar_cols if ax == ax1 else [PURPLE] * len(nomes)
        ax.barh(y, data,
                color=[c + '55' for c in col],
                edgecolor=col, linewidth=1.5, height=0.55)
        ax.set_yticks(y)
        ax.set_yticklabels(nomes, color=TXT, fontsize=7.5)
        ax.invert_yaxis()
        for i, v in enumerate(data):
            ax.text(v + max(data) * 0.01, i, f'{v}{unit}',
                    va='center', color=TXT, fontsize=7.5, fontweight='bold')
        style_ax(ax, title)

    plt.tight_layout(pad=1)
    return fig_to_rl_image(fig, avail_w)


def chart_carga_media(por_equip, avail_w):
    eq_data = sorted(
        [(e, sum(r.volume for r in v) / len(v)) for e, v in por_equip.items()],
        key=lambda x: -x[1]
    )
    nomes = [e for e, _ in eq_data]
    meds  = [round(m, 1) for _, m in eq_data]

    fig, ax = plt.subplots(figsize=(10, 3.5), facecolor=S1)
    ax.set_facecolor(S1)
    x = np.arange(len(nomes))
    bar_cols = [GREEN if m >= 7.5 else GOLD if m >= 6 else RED for m in meds]
    ax.bar(x, meds, color=[c + '55' for c in bar_cols],
           edgecolor=bar_cols, linewidth=1.5, width=0.55)
    ax.set_xticks(x)
    ax.set_xticklabels(nomes, color=TXT, fontsize=7.5, rotation=30 if len(nomes) > 8 else 0)
    for i, v in enumerate(meds):
        ax.text(i, v + 0.05, f'{v}', ha='center', va='bottom',
                color=TXT, fontsize=7.5, fontweight='bold')
    ax.axhline(y=7.5, color=GREEN, linestyle='--', linewidth=1, alpha=0.4)
    style_ax(ax, 'Carga Média por Betoneira (m³/viagem)  |  Verde ≥7,5  |  Amarelo ≥6  |  Vermelho <6')
    ax.set_ylabel('m³/viagem', fontsize=7, color=MUT)
    plt.tight_layout(pad=0.8)
    return fig_to_rl_image(fig, avail_w)


def chart_horas(remessas, avail_w):
    counts = [0] * 24
    for r in remessas:
        h = int(r.hora.split(':')[0])
        counts[h] += 1

    horas = list(range(6, 21))
    vals  = counts[6:21]

    fig, ax = plt.subplots(figsize=(10, 3.5), facecolor=S1)
    ax.set_facecolor(S1)
    bar_cols = [GREEN if v >= 8 else GOLD if v >= 4 else RED for v in vals]
    ax.bar(horas, vals, color=[c + '66' for c in bar_cols],
           edgecolor=bar_cols, linewidth=1.5, width=0.7)
    ax.set_xticks(horas)
    ax.set_xticklabels([f'{h:02d}h' for h in horas], color=MUT, fontsize=7.5)
    for h, v in zip(horas, vals):
        if v > 0:
            ax.text(h, v + 0.05, str(v), ha='center', va='bottom',
                    color=TXT, fontsize=7.5, fontweight='bold')
    style_ax(ax, 'Distribuição de Saídas por Hora do Dia')
    ax.set_ylabel('Remessas', fontsize=7, color=MUT)
    plt.tight_layout(pad=0.8)
    return fig_to_rl_image(fig, avail_w)


def chart_gaps(por_dia, avail_w):
    gaps = []
    for d, rems in por_dia.items():
        rems_sorted = sorted(rems, key=lambda r: hm2min(r.hora))
        for i in range(1, len(rems_sorted)):
            g = hm2min(rems_sorted[i].hora) - hm2min(rems_sorted[i - 1].hora)
            if 0 <= g <= 300:
                gaps.append(g)

    buckets = [0] * 6
    labels  = ['0–10 min', '10–20', '20–30', '30–45', '45–60', '> 60']
    bcols   = [GREEN, GOLD, ORANGE, RED, PURPLE, BLUE]
    for g in gaps:
        if g <= 10:   buckets[0] += 1
        elif g <= 20: buckets[1] += 1
        elif g <= 30: buckets[2] += 1
        elif g <= 45: buckets[3] += 1
        elif g <= 60: buckets[4] += 1
        else:         buckets[5] += 1

    fig, ax = plt.subplots(figsize=(8, 3.5), facecolor=S1)
    ax.set_facecolor(S1)
    x = np.arange(len(labels))
    ax.bar(x, buckets, color=[c + '66' for c in bcols],
           edgecolor=bcols, linewidth=1.5, width=0.6)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, color=MUT, fontsize=7.5)
    for i, v in enumerate(buckets):
        if v > 0:
            ax.text(i, v + 0.05, str(v), ha='center', va='bottom',
                    color=TXT, fontsize=7.5, fontweight='bold')
    style_ax(ax, 'Distribuição de Gaps Entre Remessas Consecutivas')
    ax.set_ylabel('Ocorrências', fontsize=7, color=MUT)
    plt.tight_layout(pad=0.8)
    return fig_to_rl_image(fig, avail_w)


def chart_projecao(dados, avail_w):
    """Gera projeção para o mês do período."""
    from datetime import datetime, timedelta
    import calendar

    # Determina o mês/ano do período
    pi = dados['periodo_inicio']
    dia_ini = datetime.strptime(pi, '%d/%m/%Y')
    ano, mes = dia_ini.year, dia_ini.month
    nome_mes = ['', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'][mes]

    # Calcula média por dia útil (Seg–Sex) a partir dos dados reais
    por_dia = dados['por_dia']
    dias_uteis_obs = []
    for d, rems in por_dia.items():
        try:
            dt = datetime.strptime(d, '%d/%m/%Y')
            dow = dt.weekday()  # 0=Seg, 6=Dom
            if dow < 5:  # Seg–Sex
                dias_uteis_obs.append(sum(r.volume for r in rems))
        except:
            pass

    media_dia_util = sum(dias_uteis_obs) / len(dias_uteis_obs) if dias_uteis_obs else 100
    media_sabado   = media_dia_util * 0.5

    # Monta calendário do mês
    total_dias = calendar.monthrange(ano, mes)[1]
    primeiro_dia = datetime(ano, mes, 1).weekday()  # 0=Seg

    proj_vols = []
    proj_labels = []
    dow_abbr = ['S', 'T', 'Q', 'Q', 'S', 'S', 'D']
    for d in range(1, total_dias + 1):
        dow = (primeiro_dia + d - 1) % 7
        if dow == 6:    vol = 0
        elif dow == 5:  vol = media_sabado
        else:           vol = media_dia_util
        proj_vols.append(round(vol, 1))
        proj_labels.append(f"{d}\n{dow_abbr[dow]}")

    acum = []
    ac = 0
    for v in proj_vols:
        ac += v
        acum.append(round(ac, 1))

    x = np.arange(total_dias)
    bar_cols = []
    for d in range(1, total_dias + 1):
        dow = (primeiro_dia + d - 1) % 7
        if dow == 6:   bar_cols.append('#ffffff08')
        elif dow == 5: bar_cols.append(ORANGE + '55')
        else:          bar_cols.append(GREEN + '55')

    edge_cols = []
    for d in range(1, total_dias + 1):
        dow = (primeiro_dia + d - 1) % 7
        if dow == 6:   edge_cols.append('#333')
        elif dow == 5: edge_cols.append(ORANGE)
        else:          edge_cols.append(GREEN)

    fig, ax1 = plt.subplots(figsize=(16, 4), facecolor=S1)
    ax1.set_facecolor(S1)
    ax1.bar(x, proj_vols, color=bar_cols, edgecolor=edge_cols, linewidth=1, width=0.7)
    ax1.set_xticks(x)
    ax1.set_xticklabels(proj_labels, color=MUT, fontsize=6)
    ax1.set_ylabel('m³/dia', color=MUT, fontsize=7)
    ax1.tick_params(colors=MUT)
    for sp in ['top', 'right']: ax1.spines[sp].set_visible(False)
    for sp in ['bottom', 'left']: ax1.spines[sp].set_color('#1e2535')
    ax1.grid(axis='y', color='#1e2535', linewidth=0.5)

    ax2 = ax1.twinx()
    ax2.set_facecolor(S1)
    ax2.plot(x, acum, color=ORANGE, linewidth=2)
    ax2.set_ylabel('Acumulado (m³)', color=ORANGE, fontsize=7)
    ax2.tick_params(colors=ORANGE, labelsize=7)
    for sp in ax2.spines.values(): sp.set_visible(False)

    total_proj = acum[-1]
    ticket = dados['totais']['ticket_medio']
    ax1.set_title(
        f'Projeção {nome_mes}/{ano}  |  Verde=Seg–Sex ({media_dia_util:.0f}m³)  |  '
        f'Laranja=Sáb ({media_sabado:.0f}m³)  |  Total: {total_proj:.0f}m³  ≈  '
        f'R${total_proj * ticket / 1000:.0f}k',
        color=TXT, fontsize=8, fontweight='bold', pad=8
    )
    fig.patch.set_facecolor(S1)
    plt.tight_layout(pad=0.8)
    return fig_to_rl_image(fig, avail_w), total_proj, ticket


def chart_cenarios(total_proj, ticket, avail_w):
    cenarios = [total_proj * 0.85, total_proj, total_proj * 1.15]
    labels   = ['Conservador\n(−15%)', 'Realista', 'Otimista\n(+15%)']
    bcols    = [BLUE, GREEN, GOLD]
    recs     = [v * ticket / 1000 for v in cenarios]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4), facecolor=S1)
    x = np.arange(3)
    for ax, data, title, fmt in [
        (ax1, cenarios, 'Volume (m³)',    lambda v: f'{v:.0f} m³'),
        (ax2, recs,     'Receita (R$ mil)', lambda v: f'R${v:.0f}k'),
    ]:
        ax.set_facecolor(S1)
        ax.bar(x, data, color=[c + '55' for c in bcols],
               edgecolor=bcols, linewidth=2, width=0.5)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, color=TXT, fontsize=8)
        for i, v in enumerate(data):
            ax.text(i, v + max(data) * 0.01, fmt(v),
                    ha='center', color=TXT, fontsize=8, fontweight='bold')
        style_ax(ax, title)

    plt.tight_layout(pad=1)
    return fig_to_rl_image(fig, avail_w)


# ════════════════════════════════════════
# TABELAS
# ════════════════════════════════════════

def tabela_equipamentos(por_equip, avail_w):
    eq_data = sorted(
        [(e, sum(r.volume for r in v), len(v), sum(r.vlr_total for r in v))
         for e, v in por_equip.items()],
        key=lambda x: -x[1]
    )
    max_vol = eq_data[0][1] if eq_data else 1

    rows = [['#', 'Betoneira', 'Viagens', 'Volume (m³)', 'Carga Média', 'Receita']]
    for i, (e, vol, viag, rec) in enumerate(eq_data):
        rows.append([
            str(i + 1),
            e,
            str(viag),
            f'{vol:.1f} m³',
            f'{vol / viag:.1f} m³/vg',
            f'R$ {rec:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'),
        ])

    col_ws = [0.8 * cm, 2.5 * cm, 2 * cm, 3 * cm, 3 * cm, avail_w * cm - 11.3 * cm]
    tbl = Table(rows, colWidths=col_ws)
    tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor(S2)),
        ('TEXTCOLOR',     (0, 0), (-1, 0),  colors.HexColor(MUT)),
        ('FONTNAME',      (0, 0), (-1, 0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0), (-1, -1), 7.5),
        ('FONTNAME',      (0, 1), (-1, -1), 'Helvetica'),
        ('ROWBACKGROUNDS',(0, 1), (-1, -1), [colors.HexColor(S1), colors.HexColor(S2)]),
        ('TEXTCOLOR',     (0, 1), (-1, -1), colors.HexColor(TXT)),
        ('TEXTCOLOR',     (3, 1), (3, -1),  colors.HexColor(GREEN)),
        ('TEXTCOLOR',     (5, 1), (5, -1),  colors.HexColor(GOLD)),
        ('LINEBELOW',     (0, 0), (-1, 0),  1, colors.HexColor(GOLD)),
        ('GRID',          (0, 0), (-1, -1), 0.3, colors.HexColor('#1e2535')),
        ('TOPPADDING',    (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING',   (0, 0), (-1, -1), 6),
    ]))
    return tbl


# ════════════════════════════════════════
# MONTAGEM DO PDF
# ════════════════════════════════════════

PAGE_W, PAGE_H = A4
MARGIN = 1.8 * cm


def page_bg(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor(BG))
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor(GOLD))
    canvas.rect(0, PAGE_H - 3, PAGE_W, 3, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor(MUT))
    canvas.setFont('Helvetica', 7)
    canvas.drawString(MARGIN, 0.8 * cm,
                      f'Concrelongo · Dashboard Produção · {doc.periodo_inicio} a {doc.periodo_fim}')
    canvas.drawRightString(PAGE_W - MARGIN, 0.8 * cm, f'Pág. {doc.page}')
    canvas.restoreState()


def gerar_dashboard_pdf(dados: dict) -> bytes:
    """Recebe o dicionário do parser e retorna os bytes do PDF."""

    avail_w = (PAGE_W - 2 * MARGIN) / cm

    buf = io.BytesIO()

    class DocComPeriodo(SimpleDocTemplate):
        pass

    doc = DocComPeriodo(
        buf, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm
    )
    doc.periodo_inicio = dados['periodo_inicio']
    doc.periodo_fim    = dados['periodo_fim']

    tot = dados['totais']

    def kpi_table(items):
        n = len(items)
        col_w = (PAGE_W - 2 * MARGIN) / n
        h_row = [Paragraph(lbl, ps('kl', fontSize=7, textColor=colors.HexColor(MUT),
                                    fontName='Helvetica-Bold')) for lbl, _, _ in items]
        v_row = [Paragraph(val, ps('kv', fontSize=15, textColor=colors.HexColor(GOLD),
                                    fontName='Helvetica-Bold')) for _, val, _ in items]
        i_row = [Paragraph(hint, ps('kh', fontSize=7, textColor=colors.HexColor(MUT),
                                     fontName='Helvetica')) for _, _, hint in items]
        tbl = Table([h_row, v_row, i_row], colWidths=[col_w] * n)
        tbl.setStyle(TableStyle([
            ('BACKGROUND',  (0, 0), (-1, -1), colors.HexColor(S1)),
            ('BOX',         (0, 0), (-1, -1), 0.5, colors.HexColor('#1e2535')),
            ('INNERGRID',   (0, 0), (-1, -1), 0.3, colors.HexColor('#1e2535')),
            ('LINEABOVE',   (0, 0), (-1, 0),  2,   colors.HexColor(GOLD)),
            ('TOPPADDING',  (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING',(0, 0),(-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ]))
        return tbl

    def sec(title):
        return [
            HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#1e2535')),
            Paragraph(f'▸  {title}',
                      ps('s', fontSize=10, textColor=colors.HexColor(GOLD),
                         fontName='Helvetica-Bold', spaceBefore=14, spaceAfter=8)),
        ]

    story = []

    # ── Cabeçalho ──
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph('CONCRELONGO',
                            ps('T', fontSize=26, textColor=colors.HexColor(GOLD),
                               fontName='Helvetica-Bold', spaceAfter=3)))
    story.append(Paragraph('Dashboard de Produção Analítica por Programação',
                            ps('ST', fontSize=12, textColor=colors.HexColor(TXT),
                               fontName='Helvetica', spaceAfter=3)))
    story.append(Paragraph(
        f'Vendedor: {dados["vendedor"]}  ·  '
        f'Período: {dados["periodo_inicio"]} a {dados["periodo_fim"]}',
        ps('PI', fontSize=9, textColor=colors.HexColor(MUT), fontName='Helvetica', spaceAfter=14)
    ))
    story.append(HRFlowable(width='100%', thickness=1, color=colors.HexColor(GOLD)))
    story.append(Spacer(1, 0.4 * cm))

    # ── KPIs ──
    story.append(kpi_table([
        ('VOLUME TOTAL',   f'{tot["volume"]:.1f} m³', f'{tot["remessas"]} remessas'),
        ('RECEITA TOTAL',  f'R$ {tot["geral"]:,.0f}'.replace(',', '.'), 'concreto + bomba'),
        ('VLR. CONCRETO',  f'R$ {tot["concreto"]:,.0f}'.replace(',', '.'), f'{tot["concreto"]/tot["geral"]*100:.0f}% da receita'),
        ('VLR. BOMBA',     f'R$ {tot["bomba"]:,.0f}'.replace(',', '.'),    f'{tot["bomba"]/tot["geral"]*100:.1f}% da receita'),
        ('CLIENTES',       str(tot['clientes']), 'clientes atendidos'),
        ('TICKET MÉDIO',   f'R$ {tot["ticket_medio"]:.2f}'.replace('.', ','), 'por m³'),
    ]))
    story.append(Spacer(1, 0.4 * cm))

    # ── 1. Por dia ──
    if len(dados['por_dia']) > 0:
        story += sec('1. PRODUÇÃO POR DIA')
        story.append(chart_por_dia(dados['por_dia'], avail_w))
        story.append(Spacer(1, 0.4 * cm))

    # ── 2. Clientes ──
    story += sec('2. VOLUME POR CLIENTE')
    story.append(chart_clientes(dados['por_cliente'], avail_w))
    story.append(Spacer(1, 0.4 * cm))

    # ── 3. Equipamentos ──
    story += sec('3. PRODUTIVIDADE POR BETONEIRA')
    story.append(chart_equipamentos(dados['por_equip'], avail_w))
    story.append(Spacer(1, 0.4 * cm))
    story.append(chart_carga_media(dados['por_equip'], avail_w))
    story.append(Spacer(1, 0.4 * cm))
    story.append(tabela_equipamentos(dados['por_equip'], avail_w))
    story.append(Spacer(1, 0.4 * cm))

    # ── 4. Ociosidade ──
    story += sec('4. OCIOSIDADE NA EXPEDIÇÃO')
    story.append(chart_horas(dados['remessas'], avail_w))
    story.append(Spacer(1, 0.4 * cm))
    story.append(chart_gaps(dados['por_dia'], avail_w))
    story.append(Spacer(1, 0.4 * cm))

    # ── 5. Projeção ──
    story += sec('5. PROJEÇÃO MENSAL')
    img_proj, total_proj, ticket = chart_projecao(dados, avail_w)
    story.append(img_proj)
    story.append(Spacer(1, 0.4 * cm))
    story.append(chart_cenarios(total_proj, ticket, avail_w))

    doc.build(story, onFirstPage=page_bg, onLaterPages=page_bg)
    return buf.getvalue()
