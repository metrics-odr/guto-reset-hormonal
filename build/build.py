#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard de Controle de Tráfego Pago — Funil VSL (Meta Ads × Compradores).

Lê duas abas de uma planilha Google (export CSV público) e emite os REGISTROS
BRUTOS (meta[] / sales[]) dentro do HTML. Todo o cálculo/filtro/gráfico roda no
navegador (ver build/template.html). Somente leitura; nunca escreve nas planilhas.

Funil VSL / tráfego direto — não há etapa de "Leads"/"MQL":
    Gasto → Impressões → Cliques → Page Views → Checkouts → Vendas → Faturamento

Teste local: python build/build.py --meta-file meta.csv --sales-file sales.csv --out dist/index.html
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request
from datetime import datetime, timezone, timedelta

# ==========================================================================
# CONFIGURAÇÃO DO CLIENTE
# ==========================================================================
SPREADSHEET_ID = "1SmqXOiITq97O1gtYvUfIbAAjXP3dOM50UNZ_4VepypU"
GID_META  = "1195145852"   # aba Meta Ads
GID_SALES = "1836439885"   # aba Compradores

TAX_FACTOR = 1.13806       # imposto Meta: toggle ON aplica ×1,13806 (+13,806%)

# Produto principal do funil (base de Vendas/CAC/ConvCHK/Ticket).
# Casamento por prefixo (sem acento, minúsculas) — inclui a variante "- 35%".
MAIN_PRODUCT_PREFIX = "protocolo reset hormonal"

# Campanhas de terceiros que aparecem na conta e não devem ser contabilizadas
# (nome de Campaign Name contendo qualquer um destes termos, sem distinção de
# maiúsculas/acentos/colchetes). Removidas de TODAS as visualizações do dash.
EXCLUDED_CAMPAIGN_TERMS = ["ibec", "piso", "motor"]

# Rótulos exibidos na interface (lidos pelo template.html):
CLIENT_NAME  = "Reset Hormonal"
CLIENT_SUB   = "Controle de Tráfego · VSL"
TAX_LABEL    = "Imposto Meta ×1,13806"
MAIN_PRODUCT = "Protocolo Reset Hormonal"
# ==========================================================================

EXPORT_URL = "https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"
BRT = timezone(timedelta(hours=-3))   # horário de Brasília (exibição)


# --------------------------------------------------------------------------- #
# Leitura (só leitura das planilhas)
# --------------------------------------------------------------------------- #
def fetch_csv(url: str) -> list[list[str]]:
    req = urllib.request.Request(url, headers={"User-Agent": "dash-vsl-bot/1.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
    return list(csv.reader(io.StringIO(raw)))


def read_csv_file(path: str) -> list[list[str]]:
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        return list(csv.reader(f))


def load_rows(url: str, local: str | None) -> list[list[str]]:
    return read_csv_file(local) if local else fetch_csv(url)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def norm(s: str | None) -> str:
    return strip_accents((s or "").strip().lower())


def to_float(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(v: str) -> str | None:
    if not v:
        return None
    s = str(v).strip()
    if not s:
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d/%m/%y", "%b %d, %Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def is_test_row(rowtext: str) -> bool:
    return "<test lead" in rowtext.lower()


def is_paid(status: str) -> bool:
    """Considera venda apenas status pago/aprovado. Sem coluna Status -> conta."""
    sn = norm(status)
    if not sn:
        return True
    return any(k in sn for k in ("pag", "aprov", "paid", "conclu", "complet", "ativ"))


def is_main_product(prod: str) -> bool:
    return norm(prod).startswith(MAIN_PRODUCT_PREFIX)


def is_excluded_campaign(camp: str) -> bool:
    """Campanhas de terceiros (não do cliente) — ver EXCLUDED_CAMPAIGN_TERMS."""
    cn = norm(camp)
    return any(term in cn for term in EXCLUDED_CAMPAIGN_TERMS)


# ----- Máscara de PII (a página publicada é pública) ----- #
def mask_email(e: str) -> str:
    e = (e or "").strip()
    if "@" not in e:
        return "—"
    user, dom = e.split("@", 1)
    keep = user[:2] if len(user) > 2 else user[:1]
    return f"{keep}****@{dom}"


def first_last_initial(name: str) -> str:
    parts = (name or "").strip().split()
    if not parts:
        return "—"
    return parts[0] if len(parts) == 1 else f"{parts[0]} {parts[-1][:1]}."


# --------------------------------------------------------------------------- #
# Indexação de colunas (por nome, com fallback posicional)
# --------------------------------------------------------------------------- #
def header_index(header, wanted, fallback):
    idx = {}
    hn = [norm(h) for h in header]
    for key, aliases in wanted.items():
        found = None
        for a in aliases:
            a = norm(a)
            for i, h in enumerate(hn):
                if h == a or (a and a in h):
                    found = i
                    break
            if found is not None:
                break
        idx[key] = found if found is not None else fallback.get(key)
    return idx


def cell(row, i):
    if i is None or i < 0 or i >= len(row):
        return ""
    return (row[i] or "").strip()


# --------------------------------------------------------------------------- #
# Processamento -> registros brutos
# --------------------------------------------------------------------------- #
def process(meta_rows, sales_rows):
    # ---------------- Aba META ADS ----------------
    # Colunas reais: Day · Campaign Name · Ad Set Name · Ad Name · Amount Spent ·
    #   Impressions · Link Clicks · Landing Page Views · Checkouts Initiated · ...
    mheader = meta_rows[0] if meta_rows else []
    midx = header_index(
        mheader,
        {"day": ["day", "data"], "campaign": ["campaign name", "campaign"],
         "adset": ["ad set name", "adset", "ad set"], "ad": ["ad name"],
         "spent": ["amount spent", "valor gasto", "gasto"], "impr": ["impressions", "impress"],
         "clicks": ["link clicks", "clicks", "cliques"],
         "pv": ["landing page views", "page views", "pageview", "landing"],
         "ck": ["checkouts initiated", "checkouts", "initiate checkout", "checkout"]},
        {"day": 0, "campaign": 1, "adset": 2, "ad": 3, "spent": 4, "impr": 5,
         "clicks": 6, "pv": 7, "ck": 8},
    )

    meta = []
    ad_names = set()   # nomes de anúncio REAIS do Meta -> marca a venda como tráfego pago
    ad_fallback = {}   # norm(ad_name) -> (camp, adset) do 1º Meta; só usado se a venda
                       # não trouxer utm_campaign/utm_medium próprios (raro)
    for row in meta_rows[1:]:
        if not any((c or "").strip() for c in row):
            continue
        if is_test_row(" ".join(str(c) for c in row)):
            continue
        camp = cell(row, midx["campaign"]) or "(sem campanha)"
        adset = cell(row, midx["adset"]) or "(sem conjunto)"
        ad_raw = cell(row, midx["ad"])
        ad = ad_raw or "(sem anúncio)"
        # Só nomes de anúncio REAIS entram no índice. O rótulo "(sem anúncio)" (Ad Name
        # vazio) nunca é indexado, senão a 1ª linha do Meta sem Ad Name capturaria a
        # chave e toda venda sem UTM Content cairia nessa campanha.
        an = norm(ad_raw)
        if an:
            ad_names.add(an)
            ad_fallback.setdefault(an, (camp, adset))
        meta.append({
            "d": parse_date(cell(row, midx["day"])),
            "camp": camp, "adset": adset, "ad": ad,
            "sp": round(to_float(cell(row, midx["spent"])), 4),
            "im": to_float(cell(row, midx["impr"])),
            "cl": to_float(cell(row, midx["clicks"])),
            "pv": to_float(cell(row, midx["pv"])),
            "ck": to_float(cell(row, midx["ck"])),
        })

    # ---------------- Aba COMPRADORES ----------------
    # Colunas reais: Data de Criação · Cliente / Nome · Cliente / E-mail · Produto ·
    #   Valor da Venda · UTM Content · UTM Campaign · UTM Medium · UTM Source · Status
    sheader = sales_rows[0] if sales_rows else []
    sidx = header_index(
        sheader,
        {"created": ["data de criacao", "data", "created", "created_time"],
         "name": ["cliente / nome", "nome", "full_name"],
         "email": ["cliente / e-mail", "e-mail", "email"],
         "prod": ["produto", "product"],
         "val": ["valor da venda", "valor", "value", "amount"],
         "utm_content": ["utm content", "utm_content"],
         "utm_campaign": ["utm campaign", "utm_campaign"],
         "utm_medium": ["utm medium", "utm_medium"],
         "status": ["status"]},
        {"created": 0, "name": 1, "email": 2, "prod": 3, "val": 4,
         "utm_content": 5, "utm_campaign": 6, "utm_medium": 7, "status": 9},
    )

    sales = []
    for row in sales_rows[1:]:
        if not any((c or "").strip() for c in row):
            continue
        if is_test_row(" ".join(str(c) for c in row)):
            continue
        if not is_paid(cell(row, sidx["status"])):
            continue
        prod = cell(row, sidx["prod"])
        ad_raw = cell(row, sidx["utm_content"])
        ad = ad_raw or "(sem anúncio)"
        # Só casa com o Meta quando há UTM Content real. UTM Content vazio -> an=""
        # (nunca está em ad_names), então a venda não é marcada como tráfego pago.
        an = norm(ad_raw)
        matched = an in ad_names   # an vazio jamais entra em ad_names (guardado acima)
        main = is_main_product(prod)
        # Atribuição ao tráfego rastreado: produto principal OU anúncio que existe
        # no Meta (captura orderbumps/upsells que carregam a UTM do anúncio).
        attributed = main or matched
        if not attributed:
            continue
        # Campanha/Conjunto vêm da PRÓPRIA UTM da venda (convenção do funil:
        # Campaign Name = utm_campaign, Ad Set Name = utm_medium). O mesmo criativo
        # (Ad Name = utm_content) é reutilizado em várias campanhas, então NÃO se pode
        # inferir a campanha só pelo nome do anúncio — isso jogava todas as vendas de
        # um criativo na 1ª campanha que o usou. Só se a venda não trouxer utm_campaign
        # recorremos ao fallback pelo nome do anúncio (melhor esforço).
        camp = cell(row, sidx["utm_campaign"])
        adset = cell(row, sidx["utm_medium"])
        if matched and (not camp or not adset):
            fc, fa = ad_fallback.get(an, ("", ""))
            camp = camp or fc
            adset = adset or fa
        camp = camp or "(sem campanha)"
        adset = adset or "(sem conjunto)"
        sales.append({
            "d": parse_date(cell(row, sidx["created"])),
            "camp": camp, "adset": adset, "ad": ad,
            "prod": prod or "—",
            "val": round(to_float(cell(row, sidx["val"])), 2),
            "main": 1 if main else 0,
            # meta=1 quando a venda casa com um Ad Name real do Meta (tráfego pago).
            # Vendas do produto principal sem UTM de anúncio (orgânico/direto) têm meta=0.
            "meta": 1 if matched else 0,
            "nm": first_last_initial(cell(row, sidx["name"])),
            "em": mask_email(cell(row, sidx["email"])),
        })

    # Filtro final: remove campanhas de terceiros (não são do cliente e infla o
    # gasto/contagem de forma irreal). Aplicado por último, sem alterar a lógica
    # de junção/atribuição acima — só "esconde" quem casa com EXCLUDED_CAMPAIGN_TERMS.
    meta = [m for m in meta if not is_excluded_campaign(m["camp"])]
    sales = [s for s in sales if not is_excluded_campaign(s["camp"])]

    dates = sorted({d for d in ([m["d"] for m in meta if m["d"]] + [s["d"] for s in sales if s["d"]])})
    now_brt = datetime.now(BRT)
    return {
        "build": {
            "generated_at_brt": now_brt.strftime("%d/%m/%Y %H:%M"),
            "build_id": datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S"),
            "today": now_brt.strftime("%Y-%m-%d"),
            "date_min": dates[0] if dates else None,
            "date_max": dates[-1] if dates else None,
            "tax_factor": TAX_FACTOR,
            "client_name": CLIENT_NAME,
            "client_sub": CLIENT_SUB,
            "tax_label": TAX_LABEL,
            "main_product": MAIN_PRODUCT,
        },
        "meta": meta,
        "sales": sales,
    }


# --------------------------------------------------------------------------- #
# Render
# --------------------------------------------------------------------------- #
def render(data, template_path):
    with open(template_path, "r", encoding="utf-8") as f:
        tpl = f.read()
    tpl = tpl.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    tpl = tpl.replace("__BUILD_ID__", data["build"]["build_id"])
    tpl = tpl.replace("__GENERATED_BRT__", data["build"]["generated_at_brt"])
    return tpl


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--meta-file")
    ap.add_argument("--sales-file")
    ap.add_argument("--template", default="build/template.html")
    ap.add_argument("--out", default="dist/index.html")
    args = ap.parse_args()

    meta_rows = load_rows(EXPORT_URL.format(sid=SPREADSHEET_ID, gid=GID_META), args.meta_file)
    sales_rows = load_rows(EXPORT_URL.format(sid=SPREADSHEET_ID, gid=GID_SALES), args.sales_file)
    data = process(meta_rows, sales_rows)

    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(render(data, args.template))

    b = data["build"]
    vendas = sum(s["main"] for s in data["sales"])
    fat = sum(s["val"] for s in data["sales"])
    print("== build ok ==", file=sys.stderr)
    print(f"  periodo : {b['date_min']} -> {b['date_max']}", file=sys.stderr)
    print(f"  meta    : {len(data['meta'])} linhas", file=sys.stderr)
    print(f"  sales   : {len(data['sales'])} linhas (funil) · Vendas(principal): {vendas} · Fat: R$ {fat:,.2f}", file=sys.stderr)
    print(f"  out     : {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
