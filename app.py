import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import plotly.io as pio
from datetime import datetime, date
from io import StringIO, BytesIO

# ── Brand colours ──────────────────────────────────────────────────────────
CHARCOAL   = "#5C5452"
ORANGE     = "#F47920"
OFF_WHITE  = "#FAF8F7"
LIGHT_GREY = "#E8E5E3"
MID_GREY   = "#6B6866"

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HDL Business Forecast",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ─────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  /* Background */
  .stApp {{ background-color: {OFF_WHITE}; }}

  /* KPI cards */
  .kpi-card {{
    background: white;
    border-radius: 10px;
    padding: 20px 24px;
    border-left: 5px solid {ORANGE};
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }}
  .kpi-label {{
    font-size: 0.78rem;
    color: #4A4846;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 6px;
  }}
  .kpi-value {{
    font-size: 1.9rem;
    font-weight: 700;
    color: {CHARCOAL};
    line-height: 1.1;
  }}
  .kpi-delta {{
    font-size: 0.82rem;
    margin-top: 6px;
  }}
  .delta-up   {{ color: #27ae60; }}
  .delta-down {{ color: #e74c3c; }}
  .delta-flat {{ color: #4A4846; }}

  /* Section headers */
  .section-header {{
    font-size: 1.05rem;
    font-weight: 600;
    color: {CHARCOAL};
    border-bottom: 2px solid {ORANGE};
    padding-bottom: 6px;
    margin: 28px 0 16px;
  }}

  /* Upload area */
  .upload-hint {{
    font-size: 0.85rem;
    color: {MID_GREY};
    margin-top: 8px;
  }}

  /* Hide Streamlit branding */
  #MainMenu, footer {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────

REQUIRED_COLS = {
    "ConsultantName", "ProjectNumber", "ProjectDescription",
    "DateAccepted", "Price", "Cost", "OutstandingPrice",
    "OutstandingCost", "Status",
}

def load_csv(file_obj) -> tuple[pd.DataFrame | None, str]:
    """Parse uploaded CSV; return (df, error_msg)."""
    try:
        raw = file_obj.read().decode("utf-8-sig")
        df = pd.read_csv(StringIO(raw))
    except Exception as e:
        return None, f"Could not read file: {e}"

    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        return None, f"Missing columns: {', '.join(sorted(missing))}"

    df["DateAccepted"] = pd.to_datetime(df["DateAccepted"], dayfirst=True, errors="coerce")
    for col in ["Price", "Cost", "OutstandingPrice", "OutstandingCost"]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(r"[$,]", "", regex=True), errors="coerce").fillna(0)

    df["YearMonth"] = df["DateAccepted"].dt.to_period("M")
    return df, ""


def fmt_dollar(val: float) -> str:
    if abs(val) >= 1_000_000:
        return f"${val/1_000_000:.2f}M"
    if abs(val) >= 1_000:
        return f"${val/1_000:.1f}K"
    return f"${val:,.0f}"


def delta_html(current: float, prior: float, is_currency: bool = True) -> str:
    diff = current - prior
    if prior == 0:
        pct_str = "—"
    else:
        pct = (diff / abs(prior)) * 100
        pct_str = f"{pct:+.1f}%"
    diff_str = fmt_dollar(diff) if is_currency else f"{diff:+,.0f}"
    if diff > 0:
        return f'<span class="delta-up">▲ {diff_str} ({pct_str}) vs prior month</span>'
    if diff < 0:
        return f'<span class="delta-down">▼ {diff_str} ({pct_str}) vs prior month</span>'
    return f'<span class="delta-flat">→ No change vs prior month</span>'


def kpi_card(label: str, value: str, delta_html_str: str = "") -> str:
    delta_section = f'<div class="kpi-delta">{delta_html_str}</div>' if delta_html_str else ""
    return f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      {delta_section}
    </div>
    """


DARK_TEXT = "#1A1817"   # near-black for chart text

def plotly_layout(fig: go.Figure, title: str = "") -> go.Figure:
    fig.update_layout(
        title=dict(text=title, font=dict(color=DARK_TEXT, size=13, family="sans-serif"), x=0),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color=DARK_TEXT, family="sans-serif"),
        margin=dict(l=12, r=20, t=48 if title else 16, b=16),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.04, xanchor="right", x=1,
            font=dict(size=12, color=DARK_TEXT), bgcolor="rgba(0,0,0,0)", borderwidth=0,
        ),
        xaxis=dict(
            showgrid=False, zeroline=False,
            showline=False, tickfont=dict(size=12, color=DARK_TEXT),
            tickangle=0,
        ),
        yaxis=dict(
            gridcolor=LIGHT_GREY, gridwidth=1,
            zeroline=False, showline=False,
            tickfont=dict(size=12, color=DARK_TEXT),
        ),
        bargap=0.3,
        bargroupgap=0.08,
    )
    return fig


# ── PDF report generator ───────────────────────────────────────────────────

def _delta_info(cur: float, prev: float, is_currency: bool = True):
    """Return (display_text, hex_colour) for a month-over-month delta."""
    if prev == 0:
        return "—", "#6B6866"
    diff = cur - prev
    pct  = (diff / abs(prev)) * 100
    val  = fmt_dollar(diff) if is_currency else f"{diff:+,.0f}"
    if diff > 0:
        return f"▲ {val} ({pct:+.1f}%)", "#27ae60"
    if diff < 0:
        return f"▼ {val} ({pct:+.1f}%)", "#e74c3c"
    return "→ No change", "#6B6866"


def generate_pdf_report(
    latest_month, prior_month,
    k_proj_cur,  k_val_cur,  k_urev_cur,  k_ucost_cur,
    k_proj_prev, k_val_prev, k_urev_prev, k_ucost_prev,
    fig1, fig2, fig3,
    df_table:               pd.DataFrame,
    by_consultant:          pd.DataFrame,
    total_uncl_active:      float,
    count_active:           int,
    total_uncl_cost_active: float,
    df_active_full:         pd.DataFrame,
) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors as RC
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        Image as RLImage, HRFlowable, PageBreak,
    )

    # ── Colours ────────────────────────────────────────────────────────────
    C_CHARCOAL  = RC.HexColor("#5C5452")
    C_ORANGE    = RC.HexColor("#F47920")
    C_OFF_WHITE = RC.HexColor("#FAF8F7")
    C_LIGHT_GRY = RC.HexColor("#E8E5E3")
    C_DARK      = RC.HexColor("#1A1817")
    C_MID       = RC.HexColor("#6B6866")
    C_WHITE     = RC.white

    PAGE_W, PAGE_H = A4
    MARGIN   = 1.5 * cm
    USABLE_W = PAGE_W - 2 * MARGIN

    # ── Style helpers ──────────────────────────────────────────────────────
    def ps(name, **kw):
        base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=C_DARK)
        base.update(kw)
        return ParagraphStyle(name, **base)

    S_TITLE   = ps("T",  fontName="Helvetica-Bold", fontSize=20, textColor=C_CHARCOAL, leading=26)
    S_SUB     = ps("S",  fontSize=10, textColor=C_MID)
    S_SEC     = ps("Se", fontName="Helvetica-Bold", fontSize=11, textColor=C_CHARCOAL)
    S_KVAL    = ps("KV", fontName="Helvetica-Bold", fontSize=17, textColor=C_CHARCOAL,
                   alignment=TA_CENTER, leading=22)
    S_KLBL    = ps("KL", fontSize=7, textColor=RC.HexColor("#4A4846"),
                   alignment=TA_CENTER, leading=10)
    S_KDLT    = ps("KD", fontSize=7, alignment=TA_CENTER, leading=10)
    S_TH      = ps("TH", fontName="Helvetica-Bold", fontSize=8, textColor=C_WHITE,
                   alignment=TA_CENTER, leading=11)
    S_TD      = ps("TD", fontSize=8, textColor=C_DARK, leading=11)
    S_TD_R    = ps("TR", fontSize=8, textColor=C_DARK, alignment=TA_RIGHT, leading=11)
    S_FOOT    = ps("F",  fontSize=7, textColor=C_MID, alignment=TA_CENTER)
    S_NOTE    = ps("N",  fontSize=7.5, textColor=C_MID)

    # ── Chart → Image (matplotlib, no external binary needed) ─────────────
    def chart_img(fig, w_cm, h_cm):
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        import numpy as np

        C_CHR = "#5C5452"
        C_ORG = "#F47920"
        C_LG  = "#E8E5E3"

        dpi    = 150
        fig_w  = w_cm / 2.54
        fig_h  = h_cm / 2.54
        mfig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi)
        mfig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        # Determine chart type from first trace
        traces = fig.data
        if not traces:
            ax.text(0.5, 0.5, "No data", ha="center", va="center",
                    transform=ax.transAxes, color=C_CHR)
        else:
            t0 = traces[0]
            is_hbar = (hasattr(t0, "orientation") and t0.orientation == "h")

            if is_hbar:
                # Horizontal bar
                ys  = list(t0.y) if t0.y is not None else []
                xs  = list(t0.x) if t0.x is not None else []
                pos = np.arange(len(ys))
                ax.barh(pos, xs, color=C_ORG, height=0.55)
                ax.set_yticks(pos)
                ax.set_yticklabels(ys, fontsize=7.5, color="#1A1817")
                for i, v in enumerate(xs):
                    ax.text(v * 0.02, i, f"${v/1000:.0f}K",
                            va="center", fontsize=7, color="white", fontweight="bold")
                ax.set_xlabel("")
                ax.xaxis.set_visible(False)
                ax.spines[["top","right","bottom"]].set_visible(False)
                ax.spines["left"].set_color(C_LG)

            elif len(traces) >= 2 and hasattr(traces[1], "mode") and "lines" in (traces[1].mode or ""):
                # Bar + line combo (chart 1)
                xs    = list(traces[0].x) if traces[0].x is not None else []
                bar_y = list(traces[0].y) if traces[0].y is not None else []
                lin_y = list(traces[1].y) if traces[1].y is not None else []
                pos   = np.arange(len(xs))
                ax.bar(pos, bar_y, color=C_CHR, width=0.55, label=traces[0].name)
                ax2 = ax.twinx()
                ax2.plot(pos, lin_y, color=C_ORG, linewidth=2,
                         marker="o", markersize=5, label=traces[1].name)
                ax.set_xticks(pos); ax.set_xticklabels(xs, fontsize=7.5, rotation=0, color="#1A1817")
                ax.yaxis.set_tick_params(labelsize=7.5, labelcolor="#1A1817")
                ax2.yaxis.set_tick_params(labelsize=7.5, labelcolor=C_ORG)
                ax.spines[["top","right"]].set_visible(False)
                ax2.spines[["top","left"]].set_visible(False)
                ax.spines[["bottom","left"]].set_color(C_LG)
                ax.grid(axis="y", color=C_LG, linewidth=0.5)
                h1 = mpatches.Patch(color=C_CHR, label=traces[0].name)
                h2 = mpatches.Patch(color=C_ORG, label=traces[1].name)
                ax.legend(handles=[h1, h2], fontsize=7, loc="upper left",
                          frameon=False)

            else:
                # Stacked / grouped bar
                xs  = list(traces[0].x) if traces[0].x is not None else []
                pos = np.arange(len(xs))
                colors = [C_CHR, C_ORG, "#9C9896"]
                bottom = np.zeros(len(xs))
                handles = []
                for i, tr in enumerate(traces):
                    ys = np.array(list(tr.y) if tr.y is not None else [0]*len(xs), dtype=float)
                    c  = colors[i % len(colors)]
                    ax.bar(pos, ys, bottom=bottom, color=c, width=0.6, label=tr.name)
                    bottom += ys
                    handles.append(mpatches.Patch(color=c, label=tr.name))
                ax.set_xticks(pos)
                ax.set_xticklabels(xs, fontsize=6.5, rotation=-35, ha="left", color="#1A1817")
                ax.yaxis.set_tick_params(labelsize=7.5, labelcolor="#1A1817")
                ax.spines[["top","right"]].set_visible(False)
                ax.spines[["bottom","left"]].set_color(C_LG)
                ax.grid(axis="y", color=C_LG, linewidth=0.5)
                ax.legend(handles=handles, fontsize=7, loc="upper left", frameon=False)

        # Title
        title = fig.layout.title.text if fig.layout.title and fig.layout.title.text else ""
        if title:
            ax.set_title(title, fontsize=8, color="#1A1817", loc="left", pad=6)

        mfig.tight_layout(pad=0.4)
        buf = BytesIO()
        mfig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                     facecolor="white")
        plt.close(mfig)
        buf.seek(0)
        return RLImage(buf, width=w_cm * cm, height=h_cm * cm)

    # ── KPI block ──────────────────────────────────────────────────────────
    def kpi_block(label, value, cur, prev, is_currency=True):
        d_txt, d_hex = _delta_info(cur, prev, is_currency) if prior_month else ("", "#6B6866")
        delta_para = Paragraph(
            f'<font color="{d_hex}">{d_txt}</font>', S_KDLT
        ) if d_txt else Paragraph("", S_KDLT)
        return [
            Paragraph(label.upper(), S_KLBL),
            Paragraph(value, S_KVAL),
            delta_para,
        ]

    # ── Section header ─────────────────────────────────────────────────────
    def sec(title):
        return [
            Paragraph(title, S_SEC),
            HRFlowable(width="100%", thickness=1.5, color=C_ORANGE,
                       spaceBefore=3, spaceAfter=8),
        ]

    # ── Build story ────────────────────────────────────────────────────────
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
        title=f"HDL Business Forecast — {latest_month.strftime('%B %Y')}",
    )
    story = []

    # Header bar
    hdr = Table(
        [[
            Paragraph("<b>HDL BUSINESS FORECAST</b>",
                      ps("HB", fontName="Helvetica-Bold", fontSize=15,
                         textColor=C_WHITE, leading=20)),
            Paragraph(
                f"End of Month Report<br/>"
                f"<font size='10'>{latest_month.strftime('%B %Y')}</font>",
                ps("HD", fontSize=11, textColor=C_WHITE,
                   alignment=TA_RIGHT, leading=16)),
        ]],
        colWidths=[USABLE_W * 0.6, USABLE_W * 0.4],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), C_CHARCOAL),
        ("TOPPADDING",    (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING",   (0, 0), (0,  -1), 14),
        ("RIGHTPADDING",  (-1,0), (-1, -1), 14),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(hdr)
    story.append(HRFlowable(width="100%", thickness=4, color=C_ORANGE,
                             spaceBefore=0, spaceAfter=12))

    # ── KPIs ───────────────────────────────────────────────────────────────
    story += sec("Key Performance Indicators")
    cw = USABLE_W / 4
    kpi_tbl = Table(
        [[
            kpi_block("Projects Accepted",    f"{k_proj_cur:,}",
                      k_proj_cur, k_proj_prev, False),
            kpi_block("Total Value Accepted",  fmt_dollar(k_val_cur),
                      k_val_cur,  k_val_prev),
            kpi_block("Unclaimed Revenue",     fmt_dollar(k_urev_cur),
                      k_urev_cur, k_urev_prev),
            kpi_block("Unclaimed Cost",        fmt_dollar(k_ucost_cur),
                      k_ucost_cur,k_ucost_prev),
        ]],
        colWidths=[cw] * 4,
    )
    kpi_tbl.setStyle(TableStyle([
        ("BOX",           (0,0),(0,-1), 0.75, C_ORANGE),
        ("BOX",           (1,0),(1,-1), 0.75, C_ORANGE),
        ("BOX",           (2,0),(2,-1), 0.75, C_ORANGE),
        ("BOX",           (3,0),(3,-1), 0.75, C_ORANGE),
        ("LINEAFTER",     (0,0),(2,-1), 0.5,  C_LIGHT_GRY),
        ("BACKGROUND",    (0,0),(-1,-1), C_WHITE),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
        ("TOPPADDING",    (0,0),(-1,-1), 10),
        ("BOTTOMPADDING", (0,0),(-1,-1), 10),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 14))

    # ── Trend charts ───────────────────────────────────────────────────────
    story += sec("Trends")
    half_w = (USABLE_W - 0.3 * cm) / 2
    chart_h = 5.5
    try:
        img1 = chart_img(fig1, half_w / cm, chart_h)
        img2 = chart_img(fig2, half_w / cm, chart_h)
        ct = Table([[img1, img2]], colWidths=[half_w + 0.15*cm, half_w + 0.15*cm])
        ct.setStyle(TableStyle([
            ("LEFTPADDING",   (0,0),(-1,-1), 0),
            ("RIGHTPADDING",  (0,0),(-1,-1), 0),
            ("TOPPADDING",    (0,0),(-1,-1), 0),
            ("BOTTOMPADDING", (0,0),(-1,-1), 0),
        ]))
        story.append(ct)
    except Exception as e:
        story.append(Paragraph(
            f"Charts unavailable — install kaleido: pip install kaleido ({e})", S_NOTE))
    story.append(Spacer(1, 14))

    # ── Active pipeline snapshot ───────────────────────────────────────────
    story += sec("Active Pipeline")
    pw = USABLE_W / 3
    pip_tbl = Table(
        [[
            [Paragraph("UNCLAIMED (PRODUCT + INSTALL)", S_KLBL),
             Paragraph(fmt_dollar(total_uncl_active), S_KVAL)],
            [Paragraph("ACTIVE PROJECTS", S_KLBL),
             Paragraph(f"{count_active:,}", S_KVAL)],
            [Paragraph("UNCLAIMED (INSTALL ONLY)", S_KLBL),
             Paragraph(fmt_dollar(total_uncl_cost_active), S_KVAL)],
        ]],
        colWidths=[pw] * 3,
    )
    pip_tbl.setStyle(TableStyle([
        ("BOX",           (0,0),(0,-1), 0.75, C_ORANGE),
        ("BOX",           (1,0),(1,-1), 0.75, C_ORANGE),
        ("BOX",           (2,0),(2,-1), 0.75, C_ORANGE),
        ("LINEAFTER",     (0,0),(1,-1), 0.5,  C_LIGHT_GRY),
        ("BACKGROUND",    (0,0),(-1,-1), C_WHITE),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
        ("TOPPADDING",    (0,0),(-1,-1), 8),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(pip_tbl)
    story.append(Spacer(1, 10))

    # Pipeline bar chart
    try:
        story.append(chart_img(fig3, USABLE_W / cm, 4.5))
    except Exception:
        pass
    story.append(Spacer(1, 14))

    # ── Active projects table ──────────────────────────────────────────────
    story += sec(f"Active Projects — {latest_month.strftime('%B %Y')}")

    if df_table.empty:
        story.append(Paragraph("No active projects this period.", S_NOTE))
    else:
        cols = list(df_table.columns)
        # Widths: Proj# | Description | Consultant | Date | Value | UnclRev | UnclCost
        raw_ws = [2.0, 5.2, 2.8, 2.0, 2.0, 2.0, 2.0]
        col_ws = [w * cm for w in raw_ws[:len(cols)]]
        # Stretch last col to fill
        diff = USABLE_W - sum(col_ws)
        if diff:
            col_ws[-1] += diff

        money_cols = {"Value ($)", "Unclaimed Rev ($)", "Unclaimed Cost ($)"}
        rows = [[Paragraph(c, S_TH) for c in cols]]
        for _, row in df_table.iterrows():
            rows.append([
                Paragraph(str(row[c]), S_TD_R if c in money_cols else S_TD)
                for c in cols
            ])

        data_tbl = Table(rows, colWidths=col_ws, repeatRows=1)
        data_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1,  0), C_CHARCOAL),
            ("LINEBELOW",     (0, 0), (-1,  0), 1.5, C_ORANGE),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [C_WHITE, C_OFF_WHITE]),
            ("GRID",          (0, 0), (-1, -1), 0.25, C_LIGHT_GRY),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 4),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(data_tbl)

    story.append(Spacer(1, 20))

    # ── Pipeline by consultant — one page per consultant ───────────────────
    consultant_order_pdf = (
        by_consultant.sort_values("OutstandingPrice", ascending=False)["ConsultantName"].tolist()
    )

    pipe_cols_pdf    = ["ProjectNumber", "ProjectDescription", "DateAccepted",
                        "Price", "OutstandingPrice", "OutstandingCost"]
    pipe_rename_pdf  = {
        "ProjectNumber":      "Project #",
        "ProjectDescription": "Description",
        "DateAccepted":       "Date Accepted",
        "Price":              "Value ($)",
        "OutstandingPrice":   "Unclaimed Rev ($)",
        "OutstandingCost":    "Unclaimed Cost ($)",
    }
    pipe_money = {"Value ($)", "Unclaimed Rev ($)", "Unclaimed Cost ($)"}

    # Column widths for the per-consultant table
    pipe_col_ws = [w * cm for w in [2.0, 6.5, 2.2, 2.0, 2.2, 2.2]]
    diff_pipe = USABLE_W - sum(pipe_col_ws)
    if diff_pipe:
        pipe_col_ws[1] += diff_pipe  # absorb remainder into Description col

    for consultant in consultant_order_pdf:
        story.append(PageBreak())

        # Consultant header bar
        uncl_total = by_consultant.loc[
            by_consultant["ConsultantName"] == consultant, "OutstandingPrice"
        ].values[0]
        uncl_cost_total = df_active_full[
            df_active_full["ConsultantName"] == consultant
        ]["OutstandingCost"].sum()

        chdr = Table(
            [[
                Paragraph(f"<b>{consultant}</b>",
                          ps("CN", fontName="Helvetica-Bold", fontSize=14,
                             textColor=C_WHITE, leading=18)),
                Paragraph(
                    f"Active Pipeline<br/>"
                    f"<font size='9'>{fmt_dollar(uncl_total)} unclaimed</font>",
                    ps("CS", fontSize=10, textColor=C_WHITE,
                       alignment=TA_RIGHT, leading=14)),
            ]],
            colWidths=[USABLE_W * 0.6, USABLE_W * 0.4],
        )
        chdr.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), C_CHARCOAL),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("LEFTPADDING",   (0, 0), (0,  -1), 14),
            ("RIGHTPADDING",  (-1,0), (-1, -1), 14),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(chdr)
        story.append(HRFlowable(width="100%", thickness=3, color=C_ORANGE,
                                 spaceBefore=0, spaceAfter=12))

        # Mini KPI strip
        df_c_all = df_active_full[df_active_full["ConsultantName"] == consultant]
        proj_count = len(df_c_all)
        fully_uncl = (df_c_all["OutstandingPrice"] >= df_c_all["Price"]).sum()

        ckpi = Table(
            [[
                [Paragraph("ACTIVE PROJECTS", S_KLBL),
                 Paragraph(f"{proj_count:,}", S_KVAL)],
                [Paragraph("UNCLAIMED (PRODUCT + INSTALL)", S_KLBL),
                 Paragraph(fmt_dollar(uncl_total), S_KVAL)],
                [Paragraph("UNCLAIMED (INSTALL ONLY)", S_KLBL),
                 Paragraph(fmt_dollar(uncl_cost_total), S_KVAL)],
                [Paragraph("NOT YET INVOICED", S_KLBL),
                 Paragraph(f"{fully_uncl:,}", S_KVAL)],
            ]],
            colWidths=[USABLE_W / 4] * 4,
        )
        ckpi.setStyle(TableStyle([
            ("BOX",           (0,0),(0,-1), 0.75, C_ORANGE),
            ("BOX",           (1,0),(1,-1), 0.75, C_ORANGE),
            ("BOX",           (2,0),(2,-1), 0.75, C_ORANGE),
            ("BOX",           (3,0),(3,-1), 0.75, C_ORANGE),
            ("LINEAFTER",     (0,0),(2,-1), 0.5,  C_LIGHT_GRY),
            ("BACKGROUND",    (0,0),(-1,-1), C_WHITE),
            ("LEFTPADDING",   (0,0),(-1,-1), 10),
            ("RIGHTPADDING",  (0,0),(-1,-1), 10),
            ("TOPPADDING",    (0,0),(-1,-1), 8),
            ("BOTTOMPADDING", (0,0),(-1,-1), 8),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ]))
        story.append(ckpi)
        story.append(Spacer(1, 14))

        # Projects table
        story += sec("Active Jobs")
        df_c = df_active_full[
            df_active_full["ConsultantName"] == consultant
        ][pipe_cols_pdf].copy()

        # Sort: fully unclaimed first, then by OutstandingPrice desc
        df_c["_fu"] = (df_c["OutstandingPrice"] >= df_c["Price"]).astype(int)
        df_c = df_c.sort_values(["_fu", "OutstandingPrice"], ascending=[False, False])
        df_c = df_c.drop(columns="_fu")
        df_c["DateAccepted"] = pd.to_datetime(df_c["DateAccepted"], errors="coerce").dt.strftime("%d %b %Y")
        df_c = df_c.rename(columns=pipe_rename_pdf)

        if df_c.empty:
            story.append(Paragraph("No active projects.", S_NOTE))
        else:
            rows_c = [[Paragraph(col, S_TH) for col in df_c.columns]]
            for _, row in df_c.iterrows():
                rows_c.append([
                    Paragraph(
                        f"${row[col]:,.0f}" if col in pipe_money else str(row[col]),
                        S_TD_R if col in pipe_money else S_TD,
                    )
                    for col in df_c.columns
                ])
            tbl_c = Table(rows_c, colWidths=pipe_col_ws, repeatRows=1)
            tbl_c.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1,  0), C_CHARCOAL),
                ("LINEBELOW",     (0, 0), (-1,  0), 1.5, C_ORANGE),
                ("ROWBACKGROUNDS",(0, 1), (-1, -1), [C_WHITE, C_OFF_WHITE]),
                ("GRID",          (0, 0), (-1, -1), 0.25, C_LIGHT_GRY),
                ("TOPPADDING",    (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING",   (0, 0), (-1, -1), 4),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
                ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ]))
            story.append(tbl_c)

    story.append(Spacer(1, 16))

    # ── Footer ─────────────────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.75, color=C_LIGHT_GRY,
                             spaceBefore=8, spaceAfter=6))
    story.append(Paragraph(
        f"Generated {datetime.now().strftime('%d %B %Y, %I:%M %p')}  ·  "
        f"HDL Business Forecast Dashboard  ·  Confidential",
        S_FOOT,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── Session state ──────────────────────────────────────────────────────────
if "df_raw" not in st.session_state:
    st.session_state["df_raw"] = None


# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        f'<div style="font-size:1.4rem;font-weight:700;color:{CHARCOAL};'
        f'border-bottom:3px solid {ORANGE};padding-bottom:10px;margin-bottom:18px;">'
        f'HDL Forecast</div>',
        unsafe_allow_html=True,
    )

    df_raw: pd.DataFrame | None = st.session_state["df_raw"]

    if df_raw is not None:
        st.markdown("**Filters**")

        sel_consultants = df_raw["ConsultantName"].dropna().unique().tolist()  # not exposed as filter

        valid_dates = df_raw["DateAccepted"].dropna()
        min_date = valid_dates.min().date() if not valid_dates.empty else date(2023, 1, 1)
        max_date = valid_dates.max().date() if not valid_dates.empty else date.today()

        date_range = st.date_input(
            "Date range", value=(min_date, max_date),
            min_value=min_date, max_value=max_date, key="flt_dates"
        )

        sel_statuses = df_raw["Status"].dropna().unique().tolist()  # not exposed as filter

        st.divider()
        if st.button("🗑 Clear uploaded data", use_container_width=True):
            st.session_state["df_raw"] = None
            st.rerun()
    else:
        st.info("Upload a CSV on the **Upload** tab to enable filters.")


# ── Tabs ───────────────────────────────────────────────────────────────────
tab_upload, tab_insights, tab_pipeline = st.tabs(["📂  Upload", "📊  Insights", "📋  Active Pipeline"])


# ══════════════════════════════════════════════════════════════════════════
#  TAB 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown(
        f'<div style="max-width:640px;margin:40px auto 0;">'
        f'<div style="font-size:1.5rem;font-weight:700;color:{CHARCOAL};margin-bottom:6px;">'
        f'Import ProMaster Business Forecast</div>'
        f'<div style="color:{MID_GREY};font-size:0.9rem;margin-bottom:28px;">'
        f'Drop your HDLBusinessForecast CSV export below. The file must include: '
        f'ConsultantName, ProjectNumber, ProjectDescription, DateAccepted, Price, Cost, '
        f'OutstandingPrice, OutstandingCost, Status.</div>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Drag & drop or browse",
        type=["csv"],
        label_visibility="collapsed",
        key="csv_upload",
    )

    if uploaded is not None:
        df_parsed, err = load_csv(uploaded)
        if err:
            st.error(f"**Upload failed:** {err}")
        else:
            st.session_state["df_raw"] = df_parsed
            st.success(
                f"✅ Loaded **{len(df_parsed):,} rows** across "
                f"**{df_parsed['ConsultantName'].nunique()} consultants** — "
                f"switch to the **Insights** tab."
            )
            with st.expander("Preview (first 10 rows)"):
                st.dataframe(df_parsed.head(10), use_container_width=True)

    elif st.session_state["df_raw"] is not None:
        st.info("Data already loaded. Switch to **Insights** or upload a new file to replace it.")

    st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  TAB 2 — INSIGHTS
# ══════════════════════════════════════════════════════════════════════════
with tab_insights:

    df_raw = st.session_state["df_raw"]

    if df_raw is None:
        st.markdown(
            f'<div style="text-align:center;margin-top:80px;color:{MID_GREY};">'
            f'<div style="font-size:3rem;">📂</div>'
            f'<div style="font-size:1.1rem;margin-top:12px;">No data loaded yet.</div>'
            f'<div style="font-size:0.88rem;margin-top:6px;">Upload a CSV on the <b>Upload</b> tab first.</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.stop()

    # ── Apply sidebar filters ──────────────────────────────────────────────
    sel_consultants = st.session_state.get("flt_consultant", df_raw["ConsultantName"].unique().tolist())
    sel_statuses    = st.session_state.get("flt_status",     df_raw["Status"].unique().tolist())
    date_range      = st.session_state.get("flt_dates",      None)

    df = df_raw[
        df_raw["ConsultantName"].isin(sel_consultants) &
        df_raw["Status"].isin(sel_statuses)
    ].copy()

    if date_range and len(date_range) == 2:
        start_dt = pd.Timestamp(date_range[0])
        end_dt   = pd.Timestamp(date_range[1])
        df = df[(df["DateAccepted"] >= start_dt) & (df["DateAccepted"] <= end_dt)]

    if df.empty:
        st.warning("No data matches your current filters.")
        st.stop()

    # ── Derived periods ────────────────────────────────────────────────────
    all_periods  = df["YearMonth"].dropna().sort_values().unique()
    latest_month = all_periods[-1]
    prior_month  = all_periods[-2] if len(all_periods) > 1 else None

    df_cur  = df[df["YearMonth"] == latest_month]
    df_prev = df[df["YearMonth"] == prior_month] if prior_month else pd.DataFrame()

    # ── EOM KPIs ───────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="section-header">EOM KPIs — {latest_month.strftime("%B %Y")}</div>',
        unsafe_allow_html=True,
    )

    k_projects_cur   = len(df_cur)
    k_value_cur      = df_cur["Price"].sum()
    k_uncl_rev_cur   = df_cur["OutstandingPrice"].sum()
    k_uncl_cost_cur  = df_cur["OutstandingCost"].sum()

    k_projects_prev  = len(df_prev)
    k_value_prev     = df_prev["Price"].sum()   if not df_prev.empty else 0
    k_uncl_rev_prev  = df_prev["OutstandingPrice"].sum()  if not df_prev.empty else 0
    k_uncl_cost_prev = df_prev["OutstandingCost"].sum()   if not df_prev.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(kpi_card(
            "Projects Accepted",
            f"{k_projects_cur:,}",
            delta_html(k_projects_cur, k_projects_prev, is_currency=False) if prior_month else "",
        ), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card(
            "Total Value Accepted",
            fmt_dollar(k_value_cur),
            delta_html(k_value_cur, k_value_prev) if prior_month else "",
        ), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi_card(
            "Unclaimed Revenue",
            fmt_dollar(k_uncl_rev_cur),
            delta_html(k_uncl_rev_cur, k_uncl_rev_prev) if prior_month else "",
        ), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi_card(
            "Unclaimed Cost",
            fmt_dollar(k_uncl_cost_cur),
            delta_html(k_uncl_cost_cur, k_uncl_cost_prev) if prior_month else "",
        ), unsafe_allow_html=True)

    # ── Trend data — 6 months for grouped chart, 12 for stacked ───────────
    all_sorted = sorted(all_periods)
    recent_6  = all_sorted[-6:]
    recent_12 = all_sorted[-12:]

    def make_monthly(periods):
        subset = df[df["YearMonth"].isin(periods)]
        m = (
            subset.groupby("YearMonth")
            .agg(ValueAccepted=("Price", "sum"), UnclaimedRevenue=("OutstandingPrice", "sum"))
            .reset_index()
        )
        m["Month"] = m["YearMonth"].dt.strftime("%b '%y")
        return m

    monthly6  = make_monthly(recent_6)
    monthly12 = make_monthly(recent_12)

    # ── Charts row ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Trends</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    # Chart 1 — bar (value) + line (unclaimed) — last 6 months
    with col_a:
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            name="Value Accepted",
            x=monthly6["Month"], y=monthly6["ValueAccepted"],
            marker_color=CHARCOAL,
            marker_line_width=0,
        ))
        fig1.add_trace(go.Scatter(
            name="Unclaimed Revenue",
            x=monthly6["Month"], y=monthly6["UnclaimedRevenue"],
            mode="lines+markers",
            line=dict(color=ORANGE, width=2.5),
            marker=dict(size=7, color=ORANGE),
            yaxis="y",
        ))
        fig1 = plotly_layout(fig1, "Value Accepted & Unclaimed Revenue — Last 6 Months")
        fig1.update_yaxes(tickprefix="$", tickformat="~s")
        st.plotly_chart(fig1, use_container_width=True)

    # Chart 2 — stacked bar: claimed vs unclaimed — last 12 months
    with col_b:
        monthly12["Claimed"] = (monthly12["ValueAccepted"] - monthly12["UnclaimedRevenue"]).clip(lower=0)

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            name="Claimed",
            x=monthly12["Month"], y=monthly12["Claimed"],
            marker_color=CHARCOAL, marker_line_width=0,
        ))
        fig2.add_trace(go.Bar(
            name="Unclaimed",
            x=monthly12["Month"], y=monthly12["UnclaimedRevenue"],
            marker_color=ORANGE, marker_line_width=0,
        ))
        fig2.update_layout(barmode="stack")
        fig2 = plotly_layout(fig2, "Claimed vs Unclaimed — Last 12 Months")
        fig2.update_yaxes(tickprefix="$", tickformat="~s")
        fig2.update_xaxes(tickangle=-35, tickfont=dict(size=11, color=DARK_TEXT))
        st.plotly_chart(fig2, use_container_width=True)

    # ── Active pipeline charts ─────────────────────────────────────────────
    st.markdown('<div class="section-header">Active Pipeline</div>', unsafe_allow_html=True)

    df_active = df[df["Status"].str.strip().str.lower() == "active"]

    col_c, col_d = st.columns([2, 1])

    # Chart 3 — horizontal bar: unclaimed by consultant
    with col_c:
        by_consultant = (
            df_active.groupby("ConsultantName")["OutstandingPrice"]
            .sum()
            .reset_index()
            .sort_values("OutstandingPrice", ascending=True)
        )
        fig3 = go.Figure(go.Bar(
            y=by_consultant["ConsultantName"],
            x=by_consultant["OutstandingPrice"],
            orientation="h",
            marker_color=ORANGE,
            marker_line_width=0,
            text=by_consultant["OutstandingPrice"].apply(fmt_dollar),
            textposition="inside",
            insidetextanchor="end",
            textfont=dict(color="white", size=11, family="sans-serif"),
        ))
        fig3 = plotly_layout(fig3, "Unclaimed Revenue by Consultant (Active Projects)")
        fig3.update_xaxes(showticklabels=False, showgrid=False, zeroline=False)
        fig3.update_yaxes(title="", tickfont=dict(size=12, color=DARK_TEXT))
        fig3.update_layout(margin=dict(l=8, r=20, t=48, b=8), bargap=0.35)
        st.plotly_chart(fig3, use_container_width=True)

    # Active pipeline snapshot KPIs
    with col_d:
        total_uncl_active      = df_active["OutstandingPrice"].sum()
        total_uncl_cost_active = df_active["OutstandingCost"].sum()
        count_active           = df_active["ProjectNumber"].nunique()

        st.markdown(
            kpi_card("Unclaimed (Product + Install)", fmt_dollar(total_uncl_active)),
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
        st.markdown(
            kpi_card("Active Projects", f"{count_active:,}"),
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
        st.markdown(
            kpi_card("Unclaimed (Install Only)", fmt_dollar(total_uncl_cost_active)),
            unsafe_allow_html=True,
        )

    # ── Current month projects table ───────────────────────────────────────
    st.markdown(
        f'<div class="section-header">Active Projects — {latest_month.strftime("%B %Y")}</div>',
        unsafe_allow_html=True,
    )

    table_cols = [
        "ProjectNumber", "ProjectDescription", "ConsultantName",
        "DateAccepted", "Price", "OutstandingPrice", "OutstandingCost",
    ]
    df_table = df_cur[df_cur["Status"].str.strip().str.lower() == "active"][table_cols].copy()
    df_table["DateAccepted"] = df_table["DateAccepted"].dt.strftime("%d %b %Y")
    df_table = df_table.rename(columns={
        "ProjectNumber":      "Project #",
        "ProjectDescription": "Description",
        "ConsultantName":     "Consultant",
        "DateAccepted":       "Date Accepted",
        "Price":              "Value ($)",
        "OutstandingPrice":   "Unclaimed Rev ($)",
        "OutstandingCost":    "Unclaimed Cost ($)",
    })

    st.dataframe(
        df_table.style
            .format({
                "Value ($)":          "${:,.0f}",
                "Unclaimed Rev ($)":  "${:,.0f}",
                "Unclaimed Cost ($)": "${:,.0f}",
            })
            .set_properties(**{"color": "#1A1817", "background-color": "white"})
            .apply(
                lambda col: [
                    f"background-color: {LIGHT_GREY}; color: #1A1817;"
                    if col.name in ("Unclaimed Rev ($)", "Unclaimed Cost ($)") else ""
                    for _ in col
                ],
                axis=0,
            ),
        use_container_width=True,
        height=min(400, 40 + len(df_table) * 35),
    )

    # ── Downloads ──────────────────────────────────────────────────────────
    dl_col1, dl_col2 = st.columns([1, 1])

    with dl_col1:
        csv_out = df_table.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇ Download table as CSV",
            data=csv_out,
            file_name=f"HDL_EOM_{latest_month.strftime('%Y_%m')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with dl_col2:
        with st.spinner("Building PDF report…"):
            try:
                pdf_bytes = generate_pdf_report(
                    latest_month, prior_month,
                    k_projects_cur,  k_value_cur,  k_uncl_rev_cur,  k_uncl_cost_cur,
                    k_projects_prev, k_value_prev, k_uncl_rev_prev, k_uncl_cost_prev,
                    fig1, fig2, fig3,
                    df_table,
                    by_consultant,
                    total_uncl_active, count_active,
                    total_uncl_cost_active,
                    df_active,
                )
                st.download_button(
                    "📄 Download Management Report (PDF)",
                    data=pdf_bytes,
                    file_name=f"HDL_EOM_Report_{latest_month.strftime('%Y_%m')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"PDF generation failed: {e}")


# ══════════════════════════════════════════════════════════════════════════
#  TAB 3 — ACTIVE PIPELINE (all active projects, by consultant)
# ══════════════════════════════════════════════════════════════════════════
with tab_pipeline:

    df_raw = st.session_state["df_raw"]

    if df_raw is None:
        st.markdown(
            f'<div style="text-align:center;margin-top:80px;color:{MID_GREY};">'
            f'<div style="font-size:3rem;">📂</div>'
            f'<div style="font-size:1.1rem;margin-top:12px;">No data loaded yet.</div>'
            f'<div style="font-size:0.88rem;margin-top:6px;">Upload a CSV on the <b>Upload</b> tab first.</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.stop()

    # All active projects across the full dataset (date filter still applies)
    date_range_p = st.session_state.get("flt_dates", None)
    df_p = df_raw[df_raw["Status"].str.strip().str.lower() == "active"].copy()

    if date_range_p and len(date_range_p) == 2:
        df_p = df_p[
            (df_p["DateAccepted"] >= pd.Timestamp(date_range_p[0])) &
            (df_p["DateAccepted"] <= pd.Timestamp(date_range_p[1]))
        ]

    if df_p.empty:
        st.warning("No active projects found.")
        st.stop()

    # Sort within each project:
    # 1. Fully unclaimed first (OutstandingPrice == Price — nothing invoiced yet)
    # 2. Then by OutstandingPrice descending
    df_p["_fully_unclaimed"] = (df_p["OutstandingPrice"] >= df_p["Price"]).astype(int)
    df_p = df_p.sort_values(["_fully_unclaimed", "OutstandingPrice"], ascending=[False, False])

    # Consultant order: highest total unclaimed first
    consultant_order = (
        df_p.groupby("ConsultantName")["OutstandingPrice"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    pipeline_cols = [
        "ProjectNumber", "ProjectDescription", "DateAccepted",
        "Price", "OutstandingPrice", "OutstandingCost",
    ]
    pipeline_rename = {
        "ProjectNumber":      "Project #",
        "ProjectDescription": "Description",
        "DateAccepted":       "Date Accepted",
        "Price":              "Value ($)",
        "OutstandingPrice":   "Unclaimed Rev ($)",
        "OutstandingCost":    "Unclaimed Cost ($)",
    }

    total_active   = df_p["ProjectNumber"].nunique()
    total_unclaimed = df_p["OutstandingPrice"].sum()

    # Summary strip
    st.markdown(
        f'<div style="display:flex;gap:32px;margin-bottom:24px;">'
        f'<div class="kpi-card" style="flex:1;">'
        f'<div class="kpi-label">Total Active Projects</div>'
        f'<div class="kpi-value">{total_active:,}</div>'
        f'</div>'
        f'<div class="kpi-card" style="flex:1;">'
        f'<div class="kpi-label">Total Unclaimed Revenue</div>'
        f'<div class="kpi-value">{fmt_dollar(total_unclaimed)}</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # One section per consultant
    for consultant in consultant_order:
        df_c = df_p[df_p["ConsultantName"] == consultant][pipeline_cols].copy()
        df_c["DateAccepted"] = df_c["DateAccepted"].dt.strftime("%d %b %Y")
        df_c = df_c.rename(columns=pipeline_rename)

        consultant_total = df_c["Unclaimed Rev ($)"].sum()
        project_count    = len(df_c)

        st.markdown(
            f'<div class="section-header">'
            f'{consultant}'
            f'<span style="font-weight:400;font-size:0.85rem;color:{MID_GREY};margin-left:12px;">'
            f'{project_count} project{"s" if project_count != 1 else ""} &nbsp;·&nbsp; '
            f'{fmt_dollar(consultant_total)} unclaimed'
            f'</span></div>',
            unsafe_allow_html=True,
        )

        st.dataframe(
            df_c.style
                .format({
                    "Value ($)":          "${:,.0f}",
                    "Unclaimed Rev ($)":  "${:,.0f}",
                    "Unclaimed Cost ($)": "${:,.0f}",
                })
                .set_properties(**{"color": "#1A1817", "background-color": "white"})
                .apply(
                    lambda col: [
                        f"background-color: {LIGHT_GREY}; color: #1A1817;"
                        if col.name in ("Unclaimed Rev ($)", "Unclaimed Cost ($)") else ""
                        for _ in col
                    ],
                    axis=0,
                )
                .apply(
                    # Highlight rows where nothing has been invoiced yet
                    lambda row: [
                        f"background-color: #FFF4EC; color: #1A1817;"
                        if row["Unclaimed Rev ($)"] >= row["Value ($)"] else ""
                        for _ in row
                    ],
                    axis=1,
                ),
            use_container_width=True,
            hide_index=True,
            height=min(500, 44 + len(df_c) * 35),
        )
