"""
Parser do Relatório de Produção Analítico por Programação — Concrelongo
"""
import re
import pdfplumber
from dataclasses import dataclass
from typing import List
from collections import defaultdict


@dataclass
class Remessa:
    numero: int
    hora: str
    volume: float
    equipamento: str
    vlr_concreto: float
    vlr_bomba: float
    vlr_hr_extra: float
    vlr_m3_faltante: float
    vlr_outros: float
    vlr_demais: float
    vlr_total: float
    vlr_unit: float
    cliente_id: str = ""
    cliente_nome: str = ""
    obra: str = ""
    traco: str = ""
    fck: float = 0.0
    data: str = ""
    vendedor: str = ""


def parse_float(s):
    if not s: return 0.0
    try: return float(s.strip().replace(".", "").replace(",", "."))
    except: return 0.0

def extrair_fck(traco):
    m = re.search(r'FCK\s+([\d,\.]+)', traco or "")
    if m:
        try: return float(m.group(1).replace(",", "."))
        except: pass
    return 0.0

def extrair_texto(caminho_pdf):
    linhas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t: linhas.extend(t.splitlines())
    return linhas

RE_REMESSA = re.compile(
    r'^(\d{5,6})\s+(\d{2}:\d{2})\s+([\d,]+)\s+(BT\w+)\s+'
    r'([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+'
    r'([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)'
)
RE_CLIENTE = re.compile(r'^Cliente:\s*(\d+)\s*-\s*(.+?)\s+Prog\.:.*?Data:\s*(\d{2}/\d{2}/\d{4})')
RE_OBRA    = re.compile(r'^Obra:\s*(.+?)\s+Usina:')
RE_TRACO   = re.compile(r'^Tra[çc]o:\s*(FCK\s+[\d,\.]+.*?)\s+Bomba:')
RE_PERIODO = re.compile(r'Per[íi]odo:\s*(\d{2}/\d{2}/\d{4})\s+a\s+(\d{2}/\d{2}/\d{4})')
RE_VEND    = re.compile(r'^Vendedor:\s*\d+\s*-\s*(.+)')


def parse_relatorio(caminho_pdf):
    linhas = extrair_texto(caminho_pdf)
    remessas = []
    periodo_inicio = periodo_fim = vendedor_global = ""
    ctx = dict(cliente_id="", cliente_nome="", obra="", traco="", fck=0.0, data="")

    for linha in linhas:
        s = linha.strip()

        mp = RE_PERIODO.search(s)
        if mp and not periodo_inicio:
            periodo_inicio, periodo_fim = mp.group(1), mp.group(2)
            continue

        mv = RE_VEND.match(s)
        if mv and not vendedor_global:
            vendedor_global = mv.group(1).strip()
            continue

        mc = RE_CLIENTE.match(s)
        if mc:
            ctx.update(cliente_id=mc.group(1).strip(),
                       cliente_nome=mc.group(2).strip(),
                       data=mc.group(3).strip())
            continue

        mo = RE_OBRA.match(s)
        if mo:
            ctx["obra"] = mo.group(1).strip()
            continue

        mt = RE_TRACO.match(s)
        if mt:
            ctx["traco"] = mt.group(1).strip()
            ctx["fck"]   = extrair_fck(ctx["traco"])
            continue

        mr = RE_REMESSA.match(s)
        if mr:
            remessas.append(Remessa(
                numero=int(mr.group(1)), hora=mr.group(2),
                volume=parse_float(mr.group(3)), equipamento=mr.group(4),
                vlr_concreto=parse_float(mr.group(5)),
                vlr_bomba=parse_float(mr.group(6)),
                vlr_hr_extra=parse_float(mr.group(7)),
                vlr_m3_faltante=parse_float(mr.group(8)),
                vlr_outros=parse_float(mr.group(9)),
                vlr_demais=parse_float(mr.group(10)),
                vlr_total=parse_float(mr.group(11)),
                vlr_unit=parse_float(mr.group(12)),
                **ctx, vendedor=vendedor_global
            ))

    por_dia     = defaultdict(list)
    por_cliente = defaultdict(list)
    por_equip   = defaultdict(list)
    por_fck     = defaultdict(list)
    for r in remessas:
        por_dia[r.data].append(r)
        por_cliente[r.cliente_nome].append(r)
        por_equip[r.equipamento].append(r)
        por_fck[r.fck].append(r)

    tv = sum(r.volume for r in remessas)
    tc = sum(r.vlr_concreto for r in remessas)
    tb = sum(r.vlr_bomba for r in remessas)
    to = sum(r.vlr_outros for r in remessas)
    td = sum(r.vlr_demais for r in remessas)
    tg = sum(r.vlr_total for r in remessas)

    return {
        "periodo_inicio": periodo_inicio,
        "periodo_fim":    periodo_fim,
        "vendedor":       vendedor_global,
        "remessas":       remessas,
        "por_dia":        dict(por_dia),
        "por_cliente":    dict(por_cliente),
        "por_equip":      dict(por_equip),
        "por_fck":        dict(por_fck),
        "totais": {
            "volume": tv, "concreto": tc, "bomba": tb,
            "outros": to, "demais": td, "geral": tg,
            "remessas": len(remessas), "clientes": len(por_cliente),
            "ticket_medio": tg / tv if tv > 0 else 0,
        }
    }
