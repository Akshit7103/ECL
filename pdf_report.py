"""PDF Report Generator for ECL Automation -Board-presentation quality."""

import io
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from fpdf import FPDF

# ── Chart style ──────────────────────────────────────────────────────────────
plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Helvetica", "Arial", "DejaVu Sans"],
    "font.size": 9,
    "axes.titlesize": 11,
    "axes.labelsize": 9,
    "xtick.labelsize": 8,
    "ytick.labelsize": 8,
    "legend.fontsize": 8,
    "figure.facecolor": "white",
    "axes.facecolor": "#fafbfc",
    "axes.edgecolor": "#cbd5e0",
    "grid.color": "#e2e8f0",
    "grid.alpha": 0.7,
    "axes.grid": True,
    "grid.linewidth": 0.5,
})

NAVY   = "#0f2b46"
BLUE   = "#2b6cb0"
SKY    = "#4299e1"
GREEN  = "#38a169"
RED    = "#e53e3e"
AMBER  = "#d69e2e"
GREY   = "#718096"
PURPLE = "#805ad5"

C_NAVY  = (15, 43, 70)
C_BLUE  = (43, 108, 176)
C_SKY   = (66, 153, 225)
C_WHITE = (255, 255, 255)
C_LGREY = (240, 244, 248)
C_GREEN = (56, 161, 105)
C_RED   = (229, 62, 62)
C_AMBER = (214, 158, 46)
C_GREY  = (113, 128, 150)
C_GOLD  = (255, 242, 204)
C_LBUE  = (219, 234, 254)


# ═══════════════════════════════════════════════════════════════════════════════
# CHART BUILDERS  (return matplotlib Figure objects)
# ═══════════════════════════════════════════════════════════════════════════════

def _fig_odr_trend(data):
    active = [r for r in data["odr_summary"] if r["odr"] is not None and r["status"] != "no_data"]
    fig, ax = plt.subplots(figsize=(7, 2.8))
    periods = [r["period"] for r in active]
    odrs = [r["odr"] * 100 for r in active]
    colors = [AMBER if r["status"] == "partial" else BLUE for r in active]
    ax.plot(periods, odrs, color=BLUE, linewidth=2, zorder=3)
    ax.scatter(periods, odrs, c=colors, s=40, zorder=4, edgecolors="white", linewidths=1)
    ax.fill_between(range(len(periods)), odrs, alpha=0.08, color=BLUE)
    ax.set_ylabel("ODR (%)")
    ax.set_title("Observed Default Rate Trend", fontweight="bold", color=NAVY)
    plt.xticks(rotation=30, ha="right")
    fig.tight_layout()
    return fig


def _fig_ttc_bars(data):
    grades = [r for r in data["ttc_rho"] if r["ttc"] < 1]
    fig, ax = plt.subplots(figsize=(7, 2.8))
    bars = ax.bar(
        [g["grade"] for g in grades],
        [g["ttc"] * 100 for g in grades],
        color=[SKY, BLUE, AMBER, RED],
        edgecolor="white", linewidth=1.2, width=0.55, zorder=3,
    )
    for bar, g in zip(bars, grades):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.15,
                f'{g["ttc"]*100:.2f}%', ha="center", va="bottom", fontsize=8, color=NAVY, fontweight="bold")
    ax.set_ylabel("TTC PD (%)")
    ax.set_title("Through-The-Cycle PD by Grade", fontweight="bold", color=NAVY)
    fig.tight_layout()
    return fig


def _fig_correlation(data):
    cc = data.get("corr_curve", [])
    pts = [p for i, p in enumerate(cc) if i % 3 == 0 or i < 20]
    fig, ax = plt.subplots(figsize=(7, 2.8))
    ax.plot([p["pd"]*100 for p in pts], [p["rho"] for p in pts], color=PURPLE, linewidth=2)
    ax.fill_between([p["pd"]*100 for p in pts], [p["rho"] for p in pts], alpha=0.06, color=PURPLE)
    # Mark the grades
    for r in data["ttc_rho"]:
        if r["ttc"] < 1:
            ax.scatter([r["ttc"]*100], [r["rho"]], color=RED, s=50, zorder=5, edgecolors="white", linewidths=1.5)
            ax.annotate(r["grade"], (r["ttc"]*100, r["rho"]), textcoords="offset points",
                        xytext=(8, 4), fontsize=7, color=NAVY, fontweight="bold")
    ax.set_xlabel("PD (%)")
    ax.set_ylabel("Asset Correlation (rho)")
    ax.set_title("Basel II Retail IRB Asset Correlation", fontweight="bold", color=NAVY)
    fig.tight_layout()
    return fig


def _fig_fan_chart(data):
    s = data["scenarios"]
    yrs = [y for y in s["years"] if int(y) >= 2022]
    idx = [s["years"].index(y) for y in yrs]
    fig, ax = plt.subplots(figsize=(7, 3))
    up_vals = [s["Upturn"][i] for i in idx]
    base_vals = [s["Base"][i] for i in idx]
    down_vals = [s["Downturn"][i] for i in idx]
    ax.fill_between(range(len(yrs)), up_vals, down_vals, alpha=0.1, color=BLUE, label="Scenario Range")
    ax.plot(range(len(yrs)), base_vals, color=BLUE, linewidth=2.5, label="Base", zorder=3)
    ax.plot(range(len(yrs)), up_vals, color=GREEN, linewidth=1.2, linestyle="--", label="Upturn")
    ax.plot(range(len(yrs)), down_vals, color=RED, linewidth=1.2, linestyle="--", label="Downturn")
    ax.set_xticks(range(len(yrs)))
    ax.set_xticklabels(yrs, rotation=30, ha="right")
    ax.set_ylabel("GDP Z-Factor")
    ax.set_title("GDP Z-Factor Scenarios (Fan Chart)", fontweight="bold", color=NAVY)
    ax.legend(loc="best", framealpha=0.9)
    fig.tight_layout()
    return fig


def _fig_pd_comparison(data):
    pd_data = data["vasicek_pd"]
    yrs = pd_data["years"]
    grades = [r for r in data["ttc_rho"] if r["ttc"] < 1]
    colors = [BLUE, SKY, AMBER, RED]
    fig, axes = plt.subplots(1, 3, figsize=(7, 2.8), sharey=True)
    for si, scen in enumerate(["Base", "Upturn", "Downturn"]):
        ax = axes[si]
        for gi, g in enumerate(grades):
            vals = [v * 100 for v in pd_data[scen][g["grade"]]]
            ax.plot(range(len(yrs)), vals, color=colors[gi], linewidth=1.8, label=g["grade"])
        ax.set_xticks(range(len(yrs)))
        ax.set_xticklabels(yrs, rotation=45, ha="right", fontsize=7)
        ax.set_title(scen, fontweight="bold", fontsize=10,
                     color={"Base": NAVY, "Upturn": GREEN, "Downturn": RED}[scen])
        if si == 0:
            ax.set_ylabel("PD (%)")
    axes[0].legend(loc="upper right", fontsize=7, framealpha=0.9)
    fig.suptitle("Vasicek PD -All Scenarios", fontweight="bold", color=NAVY, fontsize=11, y=1.02)
    fig.tight_layout()
    return fig


def _fig_pd_base(data):
    pd_data = data["vasicek_pd"]
    yrs = pd_data["years"]
    grades = [r for r in data["ttc_rho"] if r["ttc"] < 1]
    colors = [BLUE, SKY, AMBER, RED]
    fig, ax = plt.subplots(figsize=(7, 3))
    for gi, g in enumerate(grades):
        vals = [v * 100 for v in pd_data["Base"][g["grade"]]]
        ax.plot(yrs, vals, color=colors[gi], linewidth=2, marker="o", markersize=4, label=g["grade"])
    ax.set_ylabel("PD (%)")
    ax.set_title("Vasicek PD Forecast -Base Scenario", fontweight="bold", color=NAVY)
    ax.legend(loc="best", framealpha=0.9)
    plt.xticks(rotation=30, ha="right")
    fig.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════════════════════
# PDF REPORT CLASS
# ═══════════════════════════════════════════════════════════════════════════════

class ECLReport(FPDF):

    def __init__(self, data, company="", prepared_by=""):
        super().__init__("P", "mm", "A4")
        self.data = data
        self.company = company or "Credit Risk Division"
        self.prepared_by = prepared_by or "ECL Automation Engine"
        self.report_date = datetime.now().strftime("%B %d, %Y")
        self.set_auto_page_break(auto=True, margin=22)
        self.set_margins(15, 15, 15)

    # ── Header / Footer ──────────────────────────────────────────────────

    def header(self):
        if self.page_no() <= 1:
            return
        self.set_fill_color(*C_NAVY)
        self.rect(0, 0, 210, 10, "F")
        # Thin accent line
        self.set_fill_color(*C_BLUE)
        self.rect(0, 10, 210, 0.8, "F")
        self.set_xy(15, 1.5)
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(*C_WHITE)
        self.cell(90, 7, "Expected Credit Loss Report", ln=0)
        self.cell(90, 7, self.report_date, ln=0, align="R")
        self.ln(12)

    def footer(self):
        if self.page_no() <= 1:
            return
        self.set_y(-14)
        self.set_draw_color(*C_LGREY)
        self.line(15, self.get_y(), 195, self.get_y())
        self.ln(2)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*C_GREY)
        self.cell(90, 6, f"Page {self.page_no() - 1}  |  {self.company}", ln=0)
        self.cell(90, 6, "Confidential", ln=0, align="R")

    # ── Helpers ──────────────────────────────────────────────────────────

    def _section(self, title):
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(*C_NAVY)
        self.cell(0, 10, title, ln=1)
        self.set_draw_color(*C_BLUE)
        self.set_line_width(0.6)
        self.line(15, self.get_y(), 80, self.get_y())
        self.set_line_width(0.2)
        self.ln(4)

    def _subsection(self, title):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*C_BLUE)
        self.cell(0, 7, title, ln=1)
        self.ln(1)

    def _body(self, text):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(50, 50, 50)
        self.multi_cell(0, 4.5, text)
        self.ln(2)

    def _embed(self, fig, w=170):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=180, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        buf.seek(0)
        # Disable auto page break to prevent recursion during image placement
        self.set_auto_page_break(auto=False)
        # If less than 80mm left on page, start a new one
        if self.get_y() > 210:
            self.add_page()
        x = (210 - w) / 2
        self.image(buf, x=x, w=w)
        self.ln(4)
        self.set_auto_page_break(auto=True, margin=22)

    def _table(self, headers, rows, col_widths=None, highlight_last=False):
        usable = 180
        if col_widths is None:
            col_widths = [usable / len(headers)] * len(headers)
        else:
            total = sum(col_widths)
            col_widths = [w / total * usable for w in col_widths]

        # Disable auto page break — manage breaks manually per row
        self.set_auto_page_break(auto=False)

        def _draw_header():
            self.set_font("Helvetica", "B", 8)
            self.set_fill_color(*C_NAVY)
            self.set_text_color(*C_WHITE)
            for i, h in enumerate(headers):
                align = "L" if i == 0 else "C"
                self.cell(col_widths[i], 7, str(h), border=1, ln=0, align=align, fill=True)
            self.ln()

        _draw_header()

        # Rows
        self.set_font("Helvetica", "", 8)
        for ri, row in enumerate(rows):
            # Manual page break check (leave 22mm for footer)
            if self.get_y() > 275 - 22:
                self.add_page()
                _draw_header()
                self.set_font("Helvetica", "", 8)

            is_last = ri == len(rows) - 1 and highlight_last
            if is_last:
                self.set_fill_color(*C_LBUE)
                self.set_font("Helvetica", "B", 8)
            elif ri % 2 == 0:
                self.set_fill_color(*C_LGREY)
            else:
                self.set_fill_color(*C_WHITE)
            self.set_text_color(30, 30, 30)

            for i, val in enumerate(row):
                align = "L" if i == 0 else "C"
                self.cell(col_widths[i], 6, str(val), border=1, ln=0, align=align, fill=True)
            self.ln()
            if is_last:
                self.set_font("Helvetica", "", 8)

        self.set_auto_page_break(auto=True, margin=22)
        self.ln(3)

    def _metric_row(self, metrics):
        """Render a row of metric cards."""
        card_w = 42
        gap = 2
        start_x = (210 - (card_w * len(metrics) + gap * (len(metrics)-1))) / 2
        y = self.get_y()
        for i, (label, value, sub, color) in enumerate(metrics):
            x = start_x + i * (card_w + gap)
            # Card background
            self.set_fill_color(*C_WHITE)
            self.set_draw_color(*color)
            self.set_line_width(0.5)
            self.rect(x, y, card_w, 20, "D")
            # Top accent
            self.set_fill_color(*color)
            self.rect(x, y, card_w, 2.5, "F")
            # Label
            self.set_xy(x + 2, y + 4)
            self.set_font("Helvetica", "", 6.5)
            self.set_text_color(*C_GREY)
            self.cell(card_w - 4, 3, label.upper(), ln=0)
            # Value
            self.set_xy(x + 2, y + 8)
            self.set_font("Helvetica", "B", 14)
            self.set_text_color(*C_NAVY)
            self.cell(card_w - 4, 7, str(value), ln=0)
            # Sub
            self.set_xy(x + 2, y + 16)
            self.set_font("Helvetica", "", 6)
            self.set_text_color(*C_GREY)
            self.cell(card_w - 4, 3, sub, ln=0)
        self.set_line_width(0.2)
        self.set_y(y + 24)

    # ── Pages ────────────────────────────────────────────────────────────

    def build_cover(self):
        self.set_auto_page_break(auto=False)
        self.add_page()
        # Navy background block
        self.set_fill_color(*C_NAVY)
        self.rect(0, 0, 210, 155, "F")
        # Blue accent strip
        self.set_fill_color(*C_BLUE)
        self.rect(0, 155, 210, 3, "F")

        # Logo / icon area
        self.set_xy(15, 30)
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(150, 190, 230)
        self.cell(0, 8, self.company.upper(), ln=1)

        # Main title
        self.set_xy(15, 55)
        self.set_font("Helvetica", "B", 32)
        self.set_text_color(*C_WHITE)
        self.cell(0, 14, "Expected Credit Loss", ln=1)
        self.set_x(15)
        self.cell(0, 14, "Report", ln=1)

        # Subtitle
        self.set_xy(15, 92)
        self.set_font("Helvetica", "", 14)
        self.set_text_color(180, 210, 240)
        self.cell(0, 8, "IFRS 9 / Ind AS 109 Compliant ECL Computation", ln=1)

        # Thin line
        self.set_draw_color(100, 160, 220)
        self.set_line_width(0.3)
        self.line(15, 108, 120, 108)

        # Metadata on cover
        self.set_xy(15, 114)
        self.set_font("Helvetica", "", 10)
        self.set_text_color(180, 210, 240)
        cfg = self.data.get("config_used", {})
        lines = [
            f"Report Date:   {self.report_date}",
            f"Prepared By:   {self.prepared_by}",
            f"Shock Factor:  +/-{cfg.get('shock', 0.1)*100:.0f}%",
            f"GDP LTM: {cfg.get('gdp_ltm', '-')}   |   GDP SD: {cfg.get('gdp_sd', '-')}",
            f"Hist. Cutoff:  {cfg.get('hist_cutoff', '-')}",
        ]
        for line in lines:
            self.cell(0, 6, line, ln=1)
            self.set_x(15)

        # Bottom section (white area)
        self.set_xy(15, 168)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*C_NAVY)
        self.cell(0, 7, "Computation Summary", ln=1)

        self.set_x(15)
        self.set_font("Helvetica", "", 9)
        self.set_text_color(80, 80, 80)
        odr = self.data.get("odr_summary", [])
        active = [r for r in odr if r["status"] != "no_data"]
        odrs = [r["odr"] for r in active if r["odr"] is not None]
        avg_odr = sum(odrs) / len(odrs) * 100 if odrs else 0
        grades = len(self.data.get("ttc_rho", []))
        forecast_yrs = self.data.get("vasicek_pd", {}).get("years", [])

        summary_lines = [
            f"Periods Analyzed: {len(odr)} ({len(active)} with active data)",
            f"Average ODR: {avg_odr:.4f}%",
            f"Risk Grades: {grades} DPD bucket categories",
            f"Forecast Horizon: {forecast_yrs[0] if forecast_yrs else '-'} to {forecast_yrs[-1] if forecast_yrs else '-'}",
            f"Scenarios: Base, Upturn, Downturn",
        ]
        for line in summary_lines:
            self.cell(0, 5.5, line, ln=1)
            self.set_x(15)

        # Confidential stamp
        self.set_xy(15, 275)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*C_GREY)
        self.cell(0, 5, "CONFIDENTIAL  |  For internal use only  |  Generated by ECL Automation Engine", ln=0)
        self.set_auto_page_break(auto=True, margin=22)

    def build_executive_summary(self):
        self.add_page()
        self._section("Executive Summary")

        odr = self.data.get("odr_summary", [])
        active = [r for r in odr if r["status"] != "no_data"]
        odrs = [r["odr"] for r in active if r["odr"] is not None]
        avg_odr = sum(odrs) / len(odrs) * 100 if odrs else 0
        latest = odrs[-1] * 100 if odrs else 0
        grades = len(self.data.get("ttc_rho", []))

        self._metric_row([
            ("Total Periods", str(len(odr)), f"{len(active)} active", C_BLUE),
            ("Average ODR", f"{avg_odr:.2f}%", "Across active periods", C_GREEN),
            ("Latest ODR", f"{latest:.2f}%", active[-1]["period"] if active else "-", C_AMBER),
            ("Risk Grades", str(grades), "DPD categories", C_RED),
        ])
        self.ln(2)

        self._body(
            f"The portfolio was analyzed across {len(odr)} annual periods, of which {len(active)} contained active loan data. "
            f"The average observed default rate is {avg_odr:.4f}%, with the most recent period ({active[-1]['period'] if active else 'N/A'}) "
            f"recording an ODR of {latest:.4f}%. "
            f"Loans are classified into {grades} DPD-based risk grades for through-the-cycle PD estimation."
        )

        self._subsection("ODR Trend")
        self._embed(_fig_odr_trend(self.data), w=160)

        self._subsection("TTC PD by Grade")
        self._embed(_fig_ttc_bars(self.data), w=160)

    def build_odr_analysis(self):
        self.add_page()
        self._section("Observed Default Rate Analysis")

        self._body(
            "The table below presents the observed default rate for each annual period. "
            "Partial periods (fewer than 12 months of data) are flagged, as the ODR may be understated. "
            "Periods with no active loans are excluded from the TTC calculation."
        )

        self._subsection("ODR Summary")
        headers = ["Period", "ODR", "Months", "Observations", "Status"]
        rows = []
        for r in self.data["odr_summary"]:
            odr_str = f'{r["odr"]*100:.4f}%' if r["odr"] is not None else "N/A"
            rows.append([r["period"], odr_str, str(r["months"]), f'{r["total_obs"]:,}', r["status"].title()])
        self._table(headers, rows, col_widths=[25, 20, 15, 20, 15])

        # ODR by grade
        self._subsection("ODR by Grade & Period")
        obg = self.data.get("odr_by_grade", {})
        if obg:
            grades = list(obg.keys())
            periods = [r["period"] for r in obg[grades[0]]] if grades else []
            headers2 = ["Grade"] + periods
            rows2 = []
            for g in grades:
                row = [g] + [f'{r["odr"]*100:.2f}%' for r in obg[g]]
                rows2.append(row)
            widths = [18] + [18] * len(periods)
            self._table(headers2, rows2, col_widths=widths)

    def build_ttc_correlation(self):
        self.add_page()
        self._section("TTC PD & Asset Correlation")

        self._body(
            "Through-the-cycle PD is computed as the arithmetic mean of observed default rates "
            "across all active annual periods for each risk grade. Asset correlation (rho) is derived using "
            "the Basel II Retail IRB formula: rho = 0.03 * W + 0.16 * (1-W), where W = (1 - e^(-35*TTC)) / (1 - e^(-35))."
        )

        self._subsection("Model Inputs")
        headers = ["Grade", "TTC PD", "TTC (%)", "Asset Corr. (rho)"]
        rows = []
        for r in self.data["ttc_rho"]:
            rows.append([
                r["grade"],
                f'{r["ttc"]:.6f}',
                f'{r["ttc"]*100:.4f}%',
                f'{r["rho"]:.6f}',
            ])
        self._table(headers, rows, col_widths=[20, 25, 20, 25])

        self._subsection("Asset Correlation Curve")
        self._embed(_fig_correlation(self.data), w=155)

    def build_macro_analysis(self):
        self.add_page()
        self._section("Macroeconomic Analysis")

        cfg = self.data.get("config_used", {})
        self._body(
            f"Macroeconomic variables are sourced from the IMF World Economic Outlook (WEO). "
            f"Z-factors are computed as Z = (Value - LTM) / SD using a calibration window ending at {cfg.get('hist_cutoff', 2024)}. "
            f"GDP growth is the primary scenario driver with a shock factor of +/-{cfg.get('shock', 0.1)*100:.0f}%."
        )

        self._subsection("MAV Parameters")
        headers = ["Variable", "WEO Code", "LTM", "SD"]
        rows = [[r["mev"], r["code"], f'{r["ltm"]:.2f}', f'{r["sd"]:.2f}'] for r in self.data["mav_params"]]
        self._table(headers, rows, col_widths=[35, 20, 15, 15])

        self._subsection("GDP Z-Factor Scenarios")
        self._embed(_fig_fan_chart(self.data), w=160)

        self._body(
            f"The fan chart shows the base GDP Z-factor path with upturn and downturn bands. "
            f"A positive Z-factor indicates above-average GDP growth relative to the long-term mean ({cfg.get('gdp_ltm', '-')}), "
            f"which reduces default probability through the Vasicek model."
        )

    def build_pd_results(self):
        self.add_page()
        self._section("Vasicek PD Results")

        cfg = self.data.get("config_used", {})
        self._body(
            "Point-in-time PD is computed using the Vasicek single-factor model: "
            "PD = Phi((Phi_inv(TTC) - sqrt(rho) * Z) / sqrt(1 - rho)), "
            f"where Z is the GDP Z-factor under each scenario (shock = +/-{cfg.get('shock', 0.1)*100:.0f}%)."
        )

        pd_data = self.data["vasicek_pd"]
        yrs = pd_data["years"]

        for scen in ["Base", "Upturn", "Downturn"]:
            scen_color = {"Base": C_BLUE, "Upturn": C_GREEN, "Downturn": C_RED}[scen]
            self._subsection(f"{scen} Scenario")
            headers = ["Grade", "TTC"] + yrs
            rows = []
            for r in self.data["ttc_rho"]:
                g = r["grade"]
                row = [g, f'{r["ttc"]*100:.2f}%']
                for v in pd_data[scen][g]:
                    row.append(f'{v*100:.2f}%')
                rows.append(row)
            widths = [14, 14] + [14] * len(yrs)
            self._table(headers, rows, col_widths=widths)

        self._subsection("PD Comparison -All Scenarios")
        self._embed(_fig_pd_comparison(self.data), w=170)

        # Base scenario detail chart
        if self.get_y() > 200:
            self.add_page()
        self._subsection("Base Scenario PD Forecast")
        self._embed(_fig_pd_base(self.data), w=155)

    def build_parameters(self):
        self.add_page()
        self._section("Model Parameters & Methodology")

        self._subsection("Configuration")
        cfg = self.data.get("config_used", {})
        params = [
            ["Parameter", "Value"],
            ["Scenario Shock", f'+/-{cfg.get("shock", 0.1)*100:.0f}%'],
            ["TM Start Year", str(cfg.get("tm_start_year", 2020))],
            ["Historical Cutoff", str(cfg.get("hist_cutoff", 2024))],
            ["GDP Long-Term Mean", str(cfg.get("gdp_ltm", "-"))],
            ["GDP Standard Deviation", str(cfg.get("gdp_sd", "-"))],
            ["Calibration Window", "2019 - 2025 (n=7, population std)"],
            ["Extrapolation Method", "OLS linear (2025-2027 trend)"],
            ["Forecast Horizon", " - ".join(self.data["vasicek_pd"]["years"][:1] + self.data["vasicek_pd"]["years"][-1:])],
        ]
        self._table(params[0], params[1:], col_widths=[35, 30])

        self._subsection("Methodology")
        methods = [
            ("Observed Default Rate (ODR)",
             "Annual transition matrix aggregated across all months. Default states: 90+ DPD, Write-Off (WO), ARC. "
             "ODR = Default transitions / Total active transitions."),
            ("Through-The-Cycle PD (TTC)",
             "Arithmetic mean of grade-level ODR across all active annual periods. "
             "Represents the long-run average default probability under normal economic conditions."),
            ("Asset Correlation (rho)",
             "Basel II Retail IRB formula: rho = 0.03 * W + 0.16 * (1-W), "
             "where W = (1 - exp(-35*TTC)) / (1 - exp(-35)). Range: [0.03, 0.16]."),
            ("Z-Factor",
             "Standardized macroeconomic index: Z = (Value - LTM) / SD. "
             "Positive Z indicates favorable conditions; negative indicates adverse."),
            ("Vasicek PD",
             "Single-factor conditional PD: PD = Phi((Phi_inv(TTC) - sqrt(rho) * Z) / sqrt(1 - rho)). "
             "Converts TTC to point-in-time PD using the macro Z-factor."),
            ("Scenario Construction",
             "Base = Z (unshocked). Upturn = Z + |Z| * shock. Downturn = Z - |Z| * shock."),
        ]
        for title, desc in methods:
            self.set_font("Helvetica", "B", 9)
            self.set_text_color(*C_NAVY)
            self.cell(0, 5, title, ln=1)
            self.set_font("Helvetica", "", 8)
            self.set_text_color(70, 70, 70)
            self.multi_cell(0, 4, desc)
            self.ln(2)

        # Disclaimer
        self.ln(6)
        self.set_draw_color(*C_LGREY)
        self.line(15, self.get_y(), 195, self.get_y())
        self.ln(3)
        self.set_font("Helvetica", "I", 7.5)
        self.set_text_color(*C_GREY)
        self.multi_cell(0, 3.5,
            "Disclaimer: This report is generated by the ECL Automation Engine for internal use. "
            "The computations follow standard IFRS 9 / Ind AS 109 methodology but should be reviewed "
            "and validated by qualified risk professionals before use in regulatory submissions. "
            "Past performance does not guarantee future results. Model outputs are sensitive to input data quality "
            "and parameter assumptions."
        )

    # ── Build all ────────────────────────────────────────────────────────

    def build(self):
        self.build_cover()
        self.build_executive_summary()
        self.build_odr_analysis()
        self.build_ttc_correlation()
        self.build_macro_analysis()
        self.build_pd_results()
        self.build_parameters()


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════

def generate_report(data: dict, output_path: str, company: str = "", prepared_by: str = ""):
    report = ECLReport(data, company, prepared_by)
    report.build()
    report.output(output_path)
