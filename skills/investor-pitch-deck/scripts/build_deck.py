#!/usr/bin/env python3
"""Build a 13-slide investor pitch deck (.pptx) from a JSON content spec.

usage: build_deck.py --spec deck.json --out pitch.pptx [--theme boardroom|ember|forest]
Exit code 0 = built and consistent; 2 = a text did not fit (shorten it) or numbers are inconsistent.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import *  # noqa: F401,F403
import deckkit as K

HERE = os.path.dirname(os.path.abspath(__file__))
THEMES = os.path.join(HERE, "..", "templates", "themes.json")
CW = K.CONTENT_W
M = K.MARGIN


# ------------------------------------------------------------------ derived numbers
def derive(spec):
    """Compute every number that must be consistent, and return (derived, problems)."""
    problems, d = [], {}
    mk = spec["market"]
    for k in ("tam", "sam", "som"):
        mk[k]["value"] = mk[k]["accounts"] * mk[k]["acv"]
        d[k] = mk[k]["value"]
    if not (mk["tam"]["value"] >= mk["sam"]["value"] >= mk["som"]["value"]):
        problems.append("Market sizes must satisfy TAM >= SAM >= SOM (got %s / %s / %s)"
                        % (K.money(d["tam"]), K.money(d["sam"]), K.money(d["som"])))
    tr = spec["traction"]
    mrr = tr["values"]
    d["mrr_last"] = K.money(mrr[-1] * 1000)
    d["arr"] = K.money(mrr[-1] * 12 * 1000, 2)
    d["growth"] = "%.1fx" % (mrr[-1] / mrr[0])
    d["months"] = len(mrr)
    fin = spec["financials"]
    n = len(fin["years"])
    for key in ("revenue", "costs", "gross_margin"):
        if len(fin[key]) != n:
            problems.append("financials.%s needs %d values" % (key, n))
    fin["ebitda"] = [round(r - c, 2) for r, c in zip(fin["revenue"], fin["costs"])]
    cum, be = 0.0, None
    for yr, e in zip(fin["years"], fin["ebitda"]):
        if e >= 0:
            be = be or yr
        else:
            cum += -e
    d["breakeven"] = be or "beyond forecast"
    d["cum_burn"] = cum
    ask = spec["ask"]
    pct = sum(u["pct"] for u in ask["use_of_funds"])
    if abs(pct - 100) > 0.01:
        problems.append("use_of_funds percentages sum to %s, must be 100" % pct)
    for u in ask["use_of_funds"]:
        u["amount"] = ask["amount_m"] * u["pct"] / 100.0
    if cum > ask["amount_m"] + 1e-9:
        problems.append("Forecast cumulative burn $%.1fM exceeds the raise $%.1fM; fix the plan or the ask"
                        % (cum, ask["amount_m"]))
    d["raise"] = "$%.1fM" % ask["amount_m"]
    d["tam_title"], d["som_title"] = K.money(d["tam"]), K.money(d["som"])
    d["rev_last"] = "$%gM" % fin["revenue"][-1]
    return d, problems


def sub(text, d):
    """Fill {tokens} such as {arr}, {growth}, {mrr_last} from derived numbers."""
    try:
        return text.format(**d)
    except (KeyError, IndexError):
        return text


# ------------------------------------------------------------------ slides
def s_title(prs, th, spec, d, page, foot):
    s = K.blank_slide(prs, th, bg=th["primary"])
    # decorative geometry (right side)
    K.oval(s, 8.6, -1.2, 7.2, 7.2, fill=K.mix(th["primary"], th["accent"], 0.28))
    K.oval(s, 10.3, 3.4, 4.6, 4.6, fill=K.mix(th["primary"], th["accent"], 0.55))
    K.rect(s, M, 2.35, 0.9, 0.09, th["accent"])
    K.add_text(s, M, 2.6, 8.2, 1.4, spec["company"], th, max_pt=60, min_pt=36, bold=True,
               color=th["primary_ink"], what="company name")
    K.add_text(s, M, 4.05, 7.6, 1.2, spec["tagline"], th, max_pt=24, min_pt=16,
               color=K.mix(th["primary_ink"], th["primary"], 0.15), what="tagline")
    K.add_text(s, M, 6.3, 9, 0.5, spec["subtitle"], th, max_pt=14, min_pt=10,
               color=K.mix(th["primary_ink"], th["primary"], 0.3), what="subtitle")
    K.notes(s, spec.get("notes", {}).get("title"))


def numbered_rows(s, th, items, x, y, w, row_h, gap=0.12, title_pt=19, body_pt=14, num_d=0.55):
    for i, it in enumerate(items):
        ry = y + i * (row_h + gap)
        K.oval(s, x, ry + 0.05, num_d, num_d, fill=th["accent"])
        K.add_text(s, x, ry + 0.05, num_d, num_d, str(i + 1), th, size=16, bold=True,
                   color=th["accent_ink"], align="c", anchor="m")
        K.add_text(s, x + num_d + 0.3, ry, w - num_d - 0.3, row_h,
                   [{"text": it["head"], "bold": True, "scale": title_pt / body_pt, "gap": 0},
                    {"text": it["body"], "color": th["muted"], "gap": 4}],
                   th, max_pt=body_pt, min_pt=11, what="row '%s'" % it["head"])


def s_problem(prs, th, spec, d, page, foot):
    p = spec["problem"]
    s = K.content_slide(prs, th, p["title"], page, foot, "Problem", spec["notes"].get("problem"))
    numbered_rows(s, th, p["points"], M, 1.8, 7.5, 1.4, gap=0.18)
    K.rect(s, 8.6, 1.8, 4.13, 4.95, th["primary"])
    K.add_text(s, 8.95, 2.35, 3.45, 1.3, p["stat"]["value"], th, max_pt=64, min_pt=36, bold=True,
               color=th["accent2"], what="problem stat")
    K.add_text(s, 8.95, 3.85, 3.45, 2.6, p["stat"]["label"], th, max_pt=20, min_pt=12,
               color=th["primary_ink"], what="problem stat label")


def s_solution(prs, th, spec, d, page, foot):
    p = spec["solution"]
    s = K.content_slide(prs, th, p["title"], page, foot, "Solution", spec["notes"].get("solution"))
    K.add_text(s, M, 1.8, CW, 0.95, p["statement"], th, max_pt=22, min_pt=15, color=th["muted"],
               what="solution statement")
    gap = 0.3
    w = (CW - 2 * gap) / 3
    for i, c in enumerate(p["pillars"]):
        x = M + i * (w + gap)
        K.card(s, x, 3.0, w, 3.75, th)
        K.rect(s, x, 3.0, w, 0.09, th["accent"])
        K.add_text(s, x + 0.3, 3.35, w - 0.6, 3.2,
                   [{"text": c["head"], "bold": True, "scale": 24 / 19, "gap": 0},
                    {"text": c["body"], "gap": 8}], th, max_pt=19, min_pt=11, what="pillar '%s'" % c["head"])


def s_why_now(prs, th, spec, d, page, foot):
    p = spec["why_now"]
    s = K.content_slide(prs, th, p["title"], page, foot, "Why now", spec["notes"].get("why_now"))
    rh, gap = 1.5, 0.22
    for i, it in enumerate(p["items"]):
        y = 1.8 + i * (rh + gap)
        K.rect(s, M, y, 2.7, rh, th["primary"] if i % 2 == 0 else th["accent"])
        fg = th["primary_ink"] if i % 2 == 0 else th["accent_ink"]
        K.add_text(s, M + 0.2, y, 2.3, rh, it["marker"], th, max_pt=34, min_pt=18, bold=True, color=fg,
                   align="c", anchor="m", what="why-now marker")
        K.rect(s, M + 2.7, y, CW - 2.7, rh, th["surface"])
        K.add_text(s, M + 3.05, y + 0.12, CW - 3.5, rh - 0.24,
                   [{"text": it["head"], "bold": True, "scale": 19 / 14, "gap": 0},
                    {"text": it["body"], "color": th["muted"], "gap": 4}], th, max_pt=14, min_pt=11,
                   anchor="m", what="why-now '%s'" % it["head"])


def s_market(prs, th, spec, d, page, foot):
    mk = spec["market"]
    s = K.content_slide(prs, th, sub(mk["title"], d), page, foot, "Market", spec["notes"].get("market"))
    cx, bottom = 3.15, 6.75
    rings = [("tam", 4.9, K.mix(th["primary"], th["bg"], 0.88), th["ink"]),
             ("sam", 3.4, K.mix(th["accent"], th["bg"], 0.55), th["ink"]),
             ("som", 1.9, th["primary"], th["primary_ink"])]
    for key, dia, fill, fg in rings:
        K.oval(s, cx - dia / 2, bottom - dia, dia, dia, fill=fill)
    for key, dia, fill, fg in rings:
        top = bottom - dia
        band = dia / 2 if key == "som" else 1.2
        yy = top + (0.0 if key == "som" else 0.28)
        if key == "som":
            yy = top + 0.45
            band = 1.0
        K.add_text(s, cx - 1.1, yy, 2.2, 0.35, key.upper(), th, size=12, bold=True, color=fg, align="c",
                   what=key)
        K.add_text(s, cx - 1.1, yy + 0.32, 2.2, 0.5, K.money(mk[key]["value"]), th, size=24 if key != "som" else 22,
                   bold=True, color=fg, align="c")
    # bottom-up table
    header = ["Bottom-up sizing", "Accounts", "Annual value", "Market"]
    rows = []
    for key in ("tam", "sam", "som"):
        m = mk[key]
        rows.append([{"text": "%s: %s" % (key.upper(), m["label"]), "bold": False},
                     "{:,}".format(m["accounts"]), "$%s" % "{:,}".format(m["acv"]),
                     {"text": K.money(m["value"]), "bold": True, "color": th["accent"]}])
    widths = [2.75, 1.05, 1.25, 1.1]
    size, groups = K.plan_table(header, rows, widths, 3.3, max_pt=13, min_pt=10)
    if len(groups) > 1:
        raise K.FitError("market table too long; shorten the labels")
    yb = K.add_table(s, 6.6, 1.85, widths, header, rows, th, size, aligns=["l", "r", "r", "r"], min_row_h=0.62)
    K.add_text(s, 6.6, yb + 0.25, 6.13, 6.75 - yb - 0.25,
               [{"text": "How we sized it", "bold": True, "color": th["accent"], "scale": 1.05}] +
               [{"text": b, "bullet": True} for b in mk["assumptions"]], th, max_pt=13, min_pt=10,
               gap_pt=5, what="market assumptions")


def s_product(prs, th, spec, d, page, foot):
    p = spec["product"]
    s = K.content_slide(prs, th, p["title"], page, foot, "Product", spec["notes"].get("product"))
    gap = 0.35
    w = (CW - 2 * gap) / 3
    for i, st in enumerate(p["steps"]):
        x = M + i * (w + gap)
        K.card(s, x, 1.8, w, 2.45, th)
        K.oval(s, x + 0.25, 2.0, 0.5, 0.5, fill=th["accent"])
        K.add_text(s, x + 0.25, 2.0, 0.5, 0.5, str(i + 1), th, size=15, bold=True, color=th["accent_ink"],
                   align="c", anchor="m")
        K.add_text(s, x + 0.95, 2.0, w - 1.2, 2.1,
                   [{"text": st["head"], "bold": True, "scale": 22 / 16, "gap": 0},
                    {"text": st["body"], "color": th["muted"], "gap": 6}], th, max_pt=16, min_pt=10.5,
                   what="step '%s'" % st["head"])
        if i < 2:
            K.rect(s, x + w + 0.08, 2.85, 0.19, 0.4, th["accent"], shape=K.MSO_SHAPE.CHEVRON)
    K.add_text(s, M, 4.6, 6, 0.3, "WHAT MAKES IT DIFFERENT", th, size=11, bold=True, color=th["accent"])
    K.rect(s, M, 4.95, CW, 0.015, th["line"])
    n = len(p["capabilities"])
    cg = 0.3
    cw = (CW - (n - 1) * cg) / n
    for i, c in enumerate(p["capabilities"]):
        x = M + i * (cw + cg)
        K.add_text(s, x, 5.15, cw, 1.65,
                   [{"text": c["head"], "bold": True, "scale": 19 / 15, "gap": 0},
                    {"text": c["body"], "color": th["muted"], "gap": 4}], th, max_pt=15, min_pt=10,
                   what="capability '%s'" % c["head"])


def s_traction(prs, th, spec, d, page, foot):
    t = spec["traction"]
    s = K.content_slide(prs, th, sub(t["title"], d), page, foot, "Traction", spec["notes"].get("traction"))
    K.add_text(s, M, 1.75, 8, 0.3, t["chart_title"], th, size=13, bold=True)
    pc = [K.mix(th["primary"], th["bg"], 0.35)] * (len(t["values"]) - 1) + [th["accent"]]
    K.add_chart(s, "column", M, 2.1, 8.1, 4.45, t["categories"], [("MRR ($K)", t["values"])], th,
                number_format='"$"#,##0"K"', legend=False, colors=[th["primary"]], point_colors=pc,
                font_pt=11, gap=45, axis_visible=False, label_pt=11)
    K.add_text(s, M, 6.6, 8.1, 0.3, t["source"], th, size=10, color=th["muted"])
    for i, st in enumerate(t["stats"]):
        y = 1.8 + i * 1.7
        K.card(s, 9.1, y, 3.63, 1.55, th)
        K.big_stat(s, 9.4, y + 0.14, 3.1, sub(st["value"], d), sub(st["label"], d), th, value_pt=34)


def s_model(prs, th, spec, d, page, foot):
    b = spec["business_model"]
    s = K.content_slide(prs, th, b["title"], page, foot, "Business model", spec["notes"].get("business_model"))
    tw, tg = 2.45, 0.2
    for i, t in enumerate(b["tiers"]):
        x = M + i * (tw + tg)
        feat = t.get("featured")
        K.rect(s, x, 1.8, tw, 4.95, th["primary"] if feat else th["surface"])
        fg = th["primary_ink"] if feat else th["ink"]
        sub_fg = K.mix(th["primary_ink"], th["primary"], 0.2) if feat else th["muted"]
        K.rect(s, x, 1.8, tw, 0.09, th["accent2"] if feat else th["accent"])
        K.add_text(s, x + 0.25, 2.1, tw - 0.5, 0.4, t["name"], th, size=18, bold=True, color=fg)
        K.add_text(s, x + 0.25, 2.55, tw - 0.5, 0.75, t["price"], th, max_pt=30, min_pt=18, bold=True,
                   color=th["accent2"] if feat else th["accent"], what="price")
        K.add_text(s, x + 0.25, 3.35, tw - 0.5, 0.5, t["for"], th, max_pt=12, min_pt=9, color=sub_fg, what="tier audience")
        K.add_text(s, x + 0.25, 4.0, tw - 0.5, 2.55, [{"text": f, "bullet": True} for f in t["features"]], th,
                   max_pt=15, min_pt=10, color=fg, gap_pt=9, what="tier features")
    px = M + 3 * tw + 2 * tg + 0.35
    pw = K.SLIDE_W - M - px
    K.card(s, px, 1.8, pw, 4.95, th, fill=th["bg"], line=th["line"])
    K.add_text(s, px + 0.3, 2.0, pw - 0.6, 0.35, "UNIT ECONOMICS", th, size=11, bold=True, color=th["accent"])
    ue = b["unit_economics"]
    for i, u in enumerate(ue):
        y = 2.5 + i * (4.1 / len(ue))
        K.add_text(s, px + 0.3, y, 1.55, 0.6, u["value"], th, max_pt=26, min_pt=16, bold=True, color=th["accent"],
                   anchor="m", what="unit econ value")
        K.add_text(s, px + 1.9, y, pw - 2.15, 0.6, u["label"], th, max_pt=13, min_pt=9, color=th["ink"],
                   anchor="m", what="unit econ label")


def dot(s, th, x, y, level, is_us):
    dd = 0.3
    c = th["accent"] if is_us else th["primary"]
    if level == 2:
        K.oval(s, x - dd / 2, y - dd / 2, dd, dd, fill=c)
    elif level == 1:
        K.oval(s, x - dd / 2, y - dd / 2, dd, dd, fill=K.mix(c, th["bg"], 0.62), line=c, line_w=1.25)
    else:
        K.oval(s, x - dd / 2, y - dd / 2, dd, dd, fill=None, line=K.mix(th["muted"], th["bg"], 0.35), line_w=1.25)


def s_competition(prs, th, spec, d, page, foot):
    c = spec["competition"]
    s = K.content_slide(prs, th, c["title"], page, foot, "Competition", spec["notes"].get("competition"))
    crit = c["criteria"]
    lw = 3.1
    cwid = (CW - lw) / len(crit)
    y0, hh, rh = 1.8, 0.85, 0.6
    for j, cr in enumerate(crit):
        K.add_text(s, M + lw + j * cwid + 0.08, y0, cwid - 0.16, hh, cr, th, max_pt=13, min_pt=10, bold=True,
                   align="c", anchor="b", what="criterion")
    K.rect(s, M, y0 + hh + 0.05, CW, 0.02, th["ink"])
    for i, row in enumerate(c["rows"]):
        y = y0 + hh + 0.12 + i * rh
        us = i == 0
        if us:
            K.rect(s, M, y, CW, rh - 0.04, K.mix(th["accent"], th["bg"], 0.88))
        elif i % 2 == 0:
            K.rect(s, M, y, CW, rh - 0.04, th["surface"])
        K.add_text(s, M + 0.15, y, lw - 0.3, rh - 0.04, row["name"], th, max_pt=15, min_pt=10, bold=us,
                   color=th["ink"], anchor="m", what="competitor name")
        for j, lv in enumerate(row["scores"]):
            dot(s, th, M + lw + j * cwid + cwid / 2, y + (rh - 0.04) / 2, lv, us)
    ly = y0 + hh + 0.12 + len(c["rows"]) * rh + 0.2
    for k, (lv, txt) in enumerate([(2, "Full"), (1, "Partial"), (0, "None")]):
        lx = M + 0.2 + k * 1.5
        dot(s, th, lx, ly + 0.15, lv, False)
        K.add_text(s, lx + 0.25, ly, 1.1, 0.3, txt, th, size=12, color=th["muted"], anchor="m")
    K.add_text(s, M, ly + 0.42, CW, 0.4, c["takeaway"], th, max_pt=16, min_pt=11, bold=True, color=th["ink"],
               what="competition takeaway")


def s_gtm(prs, th, spec, d, page, foot):
    g = spec["gtm"]
    s = K.content_slide(prs, th, g["title"], page, foot, "Go-to-market", spec["notes"].get("gtm"))
    gap = 0.25
    w = (CW - 2 * gap) / 3
    for i, ph in enumerate(g["phases"]):
        x = M + i * (w + gap)
        fill = [th["primary"], K.mix(th["primary"], th["accent"], 0.45), th["accent"]][i]
        fg = th["primary_ink"] if i != 1 else th["primary_ink"]
        K.rect(s, x, 1.8, w, 0.85, fill, shape=K.MSO_SHAPE.PENTAGON if False else K.MSO_SHAPE.RECTANGLE)
        K.add_text(s, x + 0.25, 1.86, w - 0.5, 0.3, ph["when"].upper(), th, size=11, bold=True,
                   color=K.mix(fg, fill, 0.2))
        K.add_text(s, x + 0.25, 2.14, w - 0.5, 0.45, ph["name"], th, size=20, bold=True, color=fg)
        K.rect(s, x, 2.65, w, 2.95, th["surface"])
        K.add_text(s, x + 0.25, 2.85, w - 0.5, 2.6, [{"text": b, "bullet": True} for b in ph["bullets"]], th,
                   max_pt=16, min_pt=10.5, gap_pt=9, what="phase '%s'" % ph["name"])
    mw = (CW - (len(g["metrics"]) - 1) * gap) / len(g["metrics"])
    for i, m in enumerate(g["metrics"]):
        x = M + i * (mw + gap)
        K.rect(s, x, 5.85, mw, 0.95, th["bg"], line=th["line"])
        K.add_text(s, x + 0.2, 5.85, 1.5, 0.95, m["value"], th, max_pt=26, min_pt=16, bold=True, color=th["accent"],
                   anchor="m", what="gtm metric")
        K.add_text(s, x + 1.75, 5.85, mw - 1.9, 0.95, m["label"], th, max_pt=13, min_pt=9, anchor="m",
                   what="gtm metric label")


def initials(name):
    parts = [p for p in name.replace("-", " ").split() if p[0].isalpha()]
    return (parts[0][0] + parts[-1][0]).upper()


def s_team(prs, th, spec, d, page, foot):
    t = spec["team"]
    s = K.content_slide(prs, th, t["title"], page, foot, "Team", spec["notes"].get("team"))
    n = len(t["people"])
    gap = 0.25
    w = (CW - (n - 1) * gap) / n
    for i, p in enumerate(t["people"]):
        x = M + i * (w + gap)
        K.card(s, x, 1.8, w, 3.95, th)
        K.oval(s, x + w / 2 - 0.5, 2.05, 1.0, 1.0, fill=[th["primary"], th["accent"]][i % 2])
        K.add_text(s, x + w / 2 - 0.5, 2.05, 1.0, 1.0, initials(p["name"]), th, size=24, bold=True,
                   color=th["primary_ink"] if i % 2 == 0 else th["accent_ink"], align="c", anchor="m")
        K.add_text(s, x + 0.2, 3.25, w - 0.4, 0.4, p["name"], th, max_pt=17, min_pt=12, bold=True, align="c",
                   what="person name")
        K.add_text(s, x + 0.2, 3.65, w - 0.4, 0.35, p["role"], th, max_pt=13, min_pt=10, bold=True,
                   color=th["accent"], align="c", what="person role")
        K.add_text(s, x + 0.2, 4.1, w - 0.4, 1.55, p["bio"], th, max_pt=15, min_pt=10, color=th["muted"],
                   align="c", what="bio for %s" % p["name"])
    K.add_text(s, M, 6.0, CW, 0.8, [{"text": "Advisors and backers", "bold": True, "color": th["accent"],
                                     "scale": 0.9, "gap": 0}, {"text": t["advisors"], "gap": 3}],
               th, max_pt=14, min_pt=10, what="advisors")


def s_financials(prs, th, spec, d, page, foot):
    f = spec["financials"]
    s = K.content_slide(prs, th, sub(f["title"], d), page, foot, "Financial summary", spec["notes"].get("financials"))
    K.add_text(s, M, 1.75, 7.4, 0.3, "Revenue vs. total costs ($M)", th, size=13, bold=True)
    K.add_chart(s, "column", M, 2.1, 7.5, 4.55, f["years"],
                [("Revenue", f["revenue"]), ("Total costs", f["costs"])], th, number_format='"$"0.0"M"',
                colors=[th["accent"], th["primary"]], font_pt=11, gap=55, axis_visible=False, label_pt=10.5)
    rows = []
    for y, g, e in zip(f["years"], f["gross_margin"], f["ebitda"]):
        col = th["good"] if e >= 0 else th["bad"]
        rows.append([{"text": y, "bold": True}, "%d%%" % g,
                     {"text": ("%s$%.1fM" % ("" if e >= 0 else "-", abs(e))), "color": col, "bold": True}])
    widths = [1.2, 1.55, 1.6]
    size, groups = K.plan_table(["Year", "Gross margin", "EBITDA"], rows, widths, 3.2, max_pt=14, min_pt=11,
                                min_row_h=0.5)
    yb = K.add_table(s, 8.4, 1.85, widths, ["Year", "Gross margin", "EBITDA"], rows, th, size,
                     aligns=["l", "r", "r"], min_row_h=0.5)
    K.add_text(s, 8.4, yb + 0.25, 4.33, 6.75 - yb - 0.25,
               [{"text": "EBITDA break-even: %s" % d["breakeven"], "bold": True, "scale": 1.15},
                {"text": "Cumulative burn before break-even: $%.1fM, covered by this round." % d["cum_burn"]
                 if d["cum_burn"] <= spec["ask"]["amount_m"] else "", "color": th["muted"]},
                {"text": f["note"], "color": th["muted"], "scale": 0.85}],
               th, max_pt=14, min_pt=10, gap_pt=8, what="financial notes")


def s_ask(prs, th, spec, d, page, foot):
    a = spec["ask"]
    s = K.content_slide(prs, th, sub(a["title"], d), page, foot, "The ask", spec["notes"].get("ask"))
    K.add_text(s, M, 1.65, 5.6, 1.25, d["raise"], th, size=72, bold=True, color=th["accent"])
    K.add_text(s, M, 2.95, 5.6, 0.45, a["round"], th, size=18, bold=True)
    K.add_text(s, M, 3.6, 5.6, 0.35, "WHAT THIS ROUND GETS US", th, size=11, bold=True, color=th["accent"])
    K.add_text(s, M, 4.0, 5.6, 2.75, [{"text": m, "bullet": True} for m in a["milestones"]], th, max_pt=20,
               min_pt=11, gap_pt=12, what="milestones")
    K.add_text(s, 6.8, 1.75, 5.9, 0.3, "Use of funds", th, size=13, bold=True)
    cats = ["%s (%s)" % (u["label"], K.money(u["amount"] * 1e6, 2 if u["amount"] * 1e6 % 1e5 else 1))
            for u in a["use_of_funds"]]
    K.add_chart(s, "doughnut", 6.8, 2.1, 5.93, 4.2, cats, [("Use of funds", [u["pct"] for u in a["use_of_funds"]])],
                th, number_format='0"%"', font_pt=12, label_pt=13,
                point_colors=[th["primary"], th["accent"], th["accent2"], K.mix(th["primary"], th["bg"], 0.45)])
    K.add_text(s, 6.8, 6.4, 5.93, 0.4, a["contact"], th, max_pt=14, min_pt=10, color=th["muted"], what="contact")


SLIDES = [s_title, s_problem, s_solution, s_why_now, s_market, s_product, s_traction, s_model, s_competition,
          s_gtm, s_team, s_financials, s_ask]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec")
    ap.add_argument("--out")
    ap.add_argument("--theme")
    ap.add_argument("--list-themes", action="store_true")
    a = ap.parse_args()
    if a.list_themes:
        for k, v in json.load(open(THEMES)).items():
            print(k, "-", v["label"])
        return
    if not a.spec or not a.out:
        ap.error("--spec and --out are required")
    spec = json.load(open(a.spec, encoding="utf-8"))
    th = K.load_theme(THEMES, a.theme or spec.get("theme"))
    spec.setdefault("notes", {})
    errors = []
    ph = K.find_placeholders(K.walk_strings(spec))
    if ph:
        print("WARNING: unresolved placeholders (ask the user or remove): %s" % ph)
    d, problems = derive(spec)
    errors += problems
    prs = K.new_prs()
    foot = "%s  |  %s  |  Confidential" % (spec["company"], spec["ask"]["round"])
    for i, fn in enumerate(SLIDES, 1):
        try:
            fn(prs, th, spec, d, i, foot)
        except K.FitError as e:
            errors.append("slide %d (%s): %s" % (i, fn.__name__[2:], e))
    K.save(prs, a.out)
    if errors:
        print("\nPROBLEMS TO FIX:")
        for e in errors:
            print(" -", e)
        sys.exit(2)
    print("OK. Derived:", {k: v for k, v in d.items() if isinstance(v, (str, int, float))})


if __name__ == "__main__":
    main()
