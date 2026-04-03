import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import re
from datetime import datetime

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="SAP Ore Analyzer",
    page_icon="⏱️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# CUSTOM CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        text-align: center;
        border: 1px solid #475569;
    }
    .kpi-card h3 {
        color: #94a3b8;
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.3rem;
    }
    .kpi-card p {
        color: #f1f5f9;
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0;
    }
    .kpi-card .sub {
        color: #64748b;
        font-size: 0.75rem;
        margin-top: 0.2rem;
    }
    .badge-ok { color: #4ade80; }
    .badge-warn { color: #fbbf24; }
    .badge-over { color: #f87171; }
    section[data-testid="stSidebar"] { background: #0f172a; }
    section[data-testid="stSidebar"] .stMarkdown h1,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 { color: #e2e8f0; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# COLUMN MAPPING (SAP export structure)
# ──────────────────────────────────────────────
COL_WBS = "Testo breve"                             # A
COL_CREATED = "Creato il"                            # B
COL_REPARTO = "Testo contab."                        # C
COL_DATE = "Data"                                    # G
COL_PERSON = "Nome del dipendente o del candidato"   # H
COL_PERIOD = "Periodo"                               # I
COL_YEAR = "Esercizio"                               # J
COL_HOURS = "Numero (un. di mis.)"                   # K
COL_NETWORK = "Network"                              # M
COL_DESC = "Descrizione attività"                    # N

# Virtual columns (created at load time)
COL_ACT_TYPE = "_tipo_attivita"
COL_ACT_LABEL = "_tipo_attivita_label"
COL_ACT_DETAIL = "_desc_dettaglio"

# ──────────────────────────────────────────────
# ACTIVITY TYPE MAPPING
# ──────────────────────────────────────────────
ACTIVITY_TYPES = {
    "HW":    "Progettazione HW / Fornitori elettrici",
    "SW":    "Progettazione SW",
    "REV":   "Revisione / Collaudo commesse precedenti",
    "RI":    "Riunioni / Coordinamento UTE",
    "PS":    "Test macchina / FAT",
    "C":     "Cantiere / Sviluppo in loco",
    "NC":    "Gestione NC elettriche",
    "APRE":  "Assistenza cantiere (montaggio/collaudo)",
    "APOST": "Assistenza post vendita",
    "ITEC":  "Studi prevendita / Offerta",
    "RD":    "R&D",
    "V":     "Attività varie (archivio, DB, nuovi prodotti)",
    "GEST":  "Gestione (fornitori, offerte, DOC, manualistica)",
}

TYPE_COLORS = {
    "HW":    "#3b82f6",
    "SW":    "#8b5cf6",
    "REV":   "#06b6d4",
    "RI":    "#f59e0b",
    "PS":    "#10b981",
    "C":     "#ef4444",
    "NC":    "#f97316",
    "APRE":  "#ec4899",
    "APOST": "#14b8a6",
    "ITEC":  "#a855f7",
    "RD":    "#6366f1",
    "V":     "#64748b",
    "GEST":  "#84cc16",
    "N/D":   "#334155",
}


def parse_activity_type(desc: str) -> tuple:
    """
    Parse activity prefix from description.
    Handles: 'SW - Software QG1', 'SW-Software', 'SW Software', 'APOST - qualcosa'
    Returns (prefix, detail) e.g. ('SW', 'Software QG1')
    """
    if not isinstance(desc, str) or not desc.strip():
        return ("N/D", "")

    desc = desc.strip()

    # Known prefixes sorted longest-first to avoid partial matches
    known = sorted(ACTIVITY_TYPES.keys(), key=len, reverse=True)

    for prefix in known:
        # Match: PREFIX followed by separator (- or space) or end of string
        pattern = rf"^{re.escape(prefix)}\s*[-–—]\s*(.*)$"
        m = re.match(pattern, desc, re.IGNORECASE)
        if m:
            return (prefix.upper(), m.group(1).strip())

        # Also match PREFIX + space + text (no dash)
        pattern2 = rf"^{re.escape(prefix)}\s+(.+)$"
        m2 = re.match(pattern2, desc, re.IGNORECASE)
        if m2:
            return (prefix.upper(), m2.group(1).strip())

        # Exact match (just the prefix, nothing else)
        if desc.upper() == prefix.upper():
            return (prefix.upper(), "")

    return ("N/D", desc)


# ──────────────────────────────────────────────
# DATA LOADING & CLEANING
# ──────────────────────────────────────────────
@st.cache_data
def load_and_clean(uploaded_file) -> pd.DataFrame:
    """Load SAP Excel export, drop summary rows, parse activity types."""
    df = pd.read_excel(uploaded_file)

    df = df.dropna(subset=[COL_WBS])
    df = df.dropna(subset=[COL_REPARTO, COL_PERSON])

    df[COL_HOURS] = pd.to_numeric(df[COL_HOURS], errors="coerce").fillna(0)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    df[COL_WBS] = df[COL_WBS].astype(str).str.strip()
    df[COL_REPARTO] = df[COL_REPARTO].astype(str).str.strip().str.upper()
    df[COL_PERSON] = df[COL_PERSON].astype(str).str.strip().str.title()
    df[COL_DESC] = df[COL_DESC].astype(str).str.strip() if COL_DESC in df.columns else ""

    # Parse activity type from description
    parsed = df[COL_DESC].apply(parse_activity_type)
    df[COL_ACT_TYPE] = parsed.apply(lambda x: x[0])
    df[COL_ACT_DETAIL] = parsed.apply(lambda x: x[1])
    df[COL_ACT_LABEL] = df[COL_ACT_TYPE].map(
        lambda t: f"{t} — {ACTIVITY_TYPES.get(t, 'Non classificato')}"
    )

    return df


def apply_filters(df, reparti, wbs_list, persone, act_types):
    """Apply cascading filters."""
    filtered = df.copy()
    if reparti:
        filtered = filtered[filtered[COL_REPARTO].isin(reparti)]
    if wbs_list:
        filtered = filtered[filtered[COL_WBS].isin(wbs_list)]
    if persone:
        filtered = filtered[filtered[COL_PERSON].isin(persone)]
    if act_types:
        filtered = filtered[filtered[COL_ACT_TYPE].isin(act_types)]
    return filtered


# ──────────────────────────────────────────────
# CHARTS
# ──────────────────────────────────────────────
def make_target_chart(actual, target):
    pct = (actual / target * 100) if target > 0 else 0
    if pct <= 80:
        color, status = "#4ade80", "OK"
    elif pct <= 100:
        color, status = "#fbbf24", "ATTENZIONE"
    else:
        color, status = "#f87171", "SUPERATO"

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=[""], x=[target], orientation="h",
        marker=dict(color="rgba(100,116,139,0.3)", line=dict(color="#64748b", width=1)),
        name="Target",
        text=[f"Target: {target:.1f}h"], textposition="inside",
        textfont=dict(color="#94a3b8", size=13),
    ))
    fig.add_trace(go.Bar(
        y=[""], x=[actual], orientation="h",
        marker=dict(color=color, line=dict(color=color, width=1)),
        name="Reale",
        text=[f"Reale: {actual:.1f}h ({pct:.0f}%)"], textposition="inside",
        textfont=dict(color="#0f172a", size=14, family="Arial Black"),
    ))
    fig.update_layout(
        barmode="overlay", height=140,
        margin=dict(l=20, r=20, t=35, b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#94a3b8")),
        xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(color="#64748b"),
                   title=None, range=[0, max(actual, target) * 1.15]),
        yaxis=dict(showticklabels=False),
        title=dict(text=f"<b>Status: {status}</b>  —  Δ = {target - actual:+.1f}h",
                   font=dict(color=color, size=14), x=0.5),
    )
    return fig


def make_heatmap(df, index_col, columns_col, title_text):
    """Generic heatmap: index_col (y) × columns_col (x), value = hours."""
    pivot = df.pivot_table(
        index=index_col, columns=columns_col, values=COL_HOURS,
        aggfunc="sum", fill_value=0,
    )
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=True).index]

    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale=[
            [0, "#0f172a"], [0.01, "#1e3a5f"], [0.25, "#2563eb"],
            [0.5, "#7c3aed"], [0.75, "#db2777"], [1, "#f43f5e"],
        ],
        text=pivot.values,
        texttemplate="%{text:.1f}h",
        textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>%{x}<br>Ore: %{z:.1f}h<extra></extra>",
        colorbar=dict(
            title=dict(text="Ore", font=dict(color="#94a3b8")),
            tickfont=dict(color="#94a3b8"),
        ),
    ))

    height = max(350, len(pivot.index) * 45 + 100)
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=40, b=80),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(tickangle=-45, tickfont=dict(color="#cbd5e1", size=10),
                   side="bottom", title=None),
        yaxis=dict(tickfont=dict(color="#cbd5e1", size=11), title=None,
                   autorange="reversed"),
        title=dict(text=f"<b>{title_text}</b>",
                   font=dict(color="#e2e8f0", size=15), x=0.5),
    )
    return fig


def make_hours_by_person_chart(df):
    person_hours = df.groupby(COL_PERSON)[COL_HOURS].sum().sort_values(ascending=True)

    fig = go.Figure(go.Bar(
        y=person_hours.index, x=person_hours.values, orientation="h",
        marker=dict(color=person_hours.values,
                    colorscale=[[0, "#2563eb"], [1, "#f43f5e"]]),
        text=[f"{v:.1f}h" for v in person_hours.values],
        textposition="auto", textfont=dict(color="#f1f5f9", size=12),
    ))
    height = max(300, len(person_hours) * 35 + 80)
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=True, gridcolor="rgba(100,116,139,0.2)",
                   tickfont=dict(color="#64748b"),
                   title=dict(text="Ore", font=dict(color="#94a3b8"))),
        yaxis=dict(tickfont=dict(color="#cbd5e1", size=11)),
        title=dict(text="<b>Ore per Persona</b>",
                   font=dict(color="#e2e8f0", size=15), x=0.5),
    )
    return fig


def make_donut_activity(df):
    """Donut chart: hours distribution by activity type."""
    type_hours = (
        df.groupby(COL_ACT_TYPE)[COL_HOURS]
        .sum()
        .sort_values(ascending=False)
    )

    labels = [f"{t} — {ACTIVITY_TYPES.get(t, '?')}" for t in type_hours.index]
    colors = [TYPE_COLORS.get(t, "#334155") for t in type_hours.index]

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=type_hours.values,
        hole=0.5,
        marker=dict(colors=colors, line=dict(color="#0f172a", width=2)),
        textinfo="label+percent",
        textfont=dict(size=11, color="#e2e8f0"),
        hovertemplate="<b>%{label}</b><br>Ore: %{value:.1f}h<br>%{percent}<extra></extra>",
        insidetextorientation="radial",
    )])

    fig.update_layout(
        height=450,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(font=dict(color="#cbd5e1", size=10), bgcolor="rgba(0,0,0,0)"),
        title=dict(text="<b>Distribuzione Ore per Tipo Attività</b>",
                   font=dict(color="#e2e8f0", size=15), x=0.5),
        annotations=[dict(
            text=f"<b>{type_hours.sum():.0f}h</b>",
            x=0.5, y=0.5, font=dict(size=22, color="#f1f5f9"),
            showarrow=False,
        )],
    )
    return fig


def make_stacked_bar_activity(df):
    """Stacked bar: hours per person, colored by activity type."""
    pivot = df.pivot_table(
        index=COL_PERSON, columns=COL_ACT_TYPE, values=COL_HOURS,
        aggfunc="sum", fill_value=0,
    )
    person_order = pivot.sum(axis=1).sort_values(ascending=True).index
    pivot = pivot.loc[person_order]

    fig = go.Figure()
    for act_type in pivot.columns:
        fig.add_trace(go.Bar(
            y=pivot.index,
            x=pivot[act_type],
            name=f"{act_type} — {ACTIVITY_TYPES.get(act_type, '?')}",
            orientation="h",
            marker=dict(color=TYPE_COLORS.get(act_type, "#334155")),
            hovertemplate=f"<b>{act_type}</b><br>"
                          "%{y}<br>Ore: %{x:.1f}h<extra></extra>",
        ))

    height = max(350, len(pivot.index) * 40 + 100)
    fig.update_layout(
        barmode="stack",
        height=height,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=True, gridcolor="rgba(100,116,139,0.2)",
                   tickfont=dict(color="#64748b"),
                   title=dict(text="Ore", font=dict(color="#94a3b8"))),
        yaxis=dict(tickfont=dict(color="#cbd5e1", size=11)),
        legend=dict(font=dict(color="#cbd5e1", size=10),
                    bgcolor="rgba(0,0,0,0)", orientation="h",
                    yanchor="bottom", y=1.02),
        title=dict(text="<b>Ore per Persona — Breakdown Tipo Attività</b>",
                   font=dict(color="#e2e8f0", size=15), x=0.5),
    )
    return fig


# ──────────────────────────────────────────────
# REPORT EXPORT
# ──────────────────────────────────────────────
def generate_excel_report(
    df_filtered, reparti, wbs_list, persone, act_types,
    actual, target, figures,
):
    """Generate a multi-sheet Excel report."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#1e293b",
            "font_color": "#e2e8f0", "border": 1,
        })

        # ── Sheet 1: Riepilogo ──
        summary_data = {
            "Parametro": [
                "Reparti selezionati", "WBS selezionate",
                "Persone selezionate", "Tipi attività selezionati",
                "Ore Totali Reali", "Target", "Delta (Target - Reale)",
                "Periodo dati", "Report generato il",
            ],
            "Valore": [
                ", ".join(reparti) if reparti else "Tutti",
                ", ".join(wbs_list) if wbs_list else "Tutte",
                ", ".join(persone) if persone else "Tutte",
                ", ".join(act_types) if act_types else "Tutti",
                f"{actual:.2f} h",
                f"{target:.2f} h" if target > 0 else "Non impostato",
                f"{target - actual:+.2f} h" if target > 0 else "N/A",
                _get_date_range(df_filtered),
                datetime.now().strftime("%d/%m/%Y %H:%M"),
            ],
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name="Riepilogo", index=False)
        ws = writer.sheets["Riepilogo"]
        for i, v in enumerate(df_summary.columns):
            ws.write(0, i, v, header_fmt)
        ws.set_column("A:A", 28)
        ws.set_column("B:B", 50)

        # ── Sheet 2: Ore per Persona ──
        person_summary = (
            df_filtered.groupby(COL_PERSON)[COL_HOURS]
            .agg(["sum", "count"])
            .rename(columns={"sum": "Ore Totali", "count": "Giorni Lavorati"})
            .sort_values("Ore Totali", ascending=False)
            .reset_index()
        )
        person_summary.columns = ["Persona", "Ore Totali", "Giorni Lavorati"]
        person_summary["Media Ore/Giorno"] = (
            person_summary["Ore Totali"] / person_summary["Giorni Lavorati"]
        ).round(2)
        person_summary.to_excel(writer, sheet_name="Ore per Persona", index=False)
        ws = writer.sheets["Ore per Persona"]
        for i, v in enumerate(person_summary.columns):
            ws.write(0, i, v, header_fmt)
        ws.set_column("A:A", 30)
        ws.set_column("B:D", 16)

        # ── Sheet 3: Ore per WBS ──
        wbs_summary = (
            df_filtered.groupby(COL_WBS)[COL_HOURS]
            .agg(["sum", "count"])
            .rename(columns={"sum": "Ore Totali", "count": "Registrazioni"})
            .sort_values("Ore Totali", ascending=False)
            .reset_index()
        )
        wbs_summary.columns = ["WBS", "Ore Totali", "Registrazioni"]
        wbs_summary.to_excel(writer, sheet_name="Ore per WBS", index=False)
        ws = writer.sheets["Ore per WBS"]
        for i, v in enumerate(wbs_summary.columns):
            ws.write(0, i, v, header_fmt)
        ws.set_column("A:A", 25)
        ws.set_column("B:C", 16)

        # ── Sheet 4: Ore per Tipo Attività ──
        act_summary = (
            df_filtered.groupby(COL_ACT_TYPE)[COL_HOURS]
            .agg(["sum", "count"])
            .rename(columns={"sum": "Ore Totali", "count": "Registrazioni"})
            .sort_values("Ore Totali", ascending=False)
            .reset_index()
        )
        act_summary.columns = ["Tipo", "Ore Totali", "Registrazioni"]
        act_summary["Descrizione"] = act_summary["Tipo"].map(
            lambda t: ACTIVITY_TYPES.get(t, "Non classificato")
        )
        act_summary["% sul Totale"] = (
            act_summary["Ore Totali"] / act_summary["Ore Totali"].sum() * 100
        ).round(1)
        act_summary = act_summary[["Tipo", "Descrizione", "Ore Totali", "Registrazioni", "% sul Totale"]]
        act_summary.to_excel(writer, sheet_name="Ore per Tipo Attività", index=False)
        ws = writer.sheets["Ore per Tipo Attività"]
        for i, v in enumerate(act_summary.columns):
            ws.write(0, i, v, header_fmt)
        ws.set_column("A:A", 10)
        ws.set_column("B:B", 45)
        ws.set_column("C:E", 16)

        # ── Sheet 5: Dettaglio righe ──
        export_cols = [COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON, COL_HOURS,
                       COL_ACT_TYPE, COL_ACT_DETAIL, COL_DESC, COL_NETWORK]
        available_cols = [c for c in export_cols if c in df_filtered.columns]
        df_detail = df_filtered[available_cols].copy()
        rename_map = {COL_ACT_TYPE: "Tipo Attività", COL_ACT_DETAIL: "Dettaglio Attività"}
        df_detail = df_detail.rename(columns=rename_map)
        df_detail.to_excel(writer, sheet_name="Dettaglio", index=False)
        ws = writer.sheets["Dettaglio"]
        for i, v in enumerate(df_detail.columns):
            ws.write(0, i, v, header_fmt)

        # ── Sheet 6: Grafici (come immagini) ──
        ws_charts = wb.add_worksheet("Grafici")
        row_offset = 1
        ws_charts.write(0, 0, "Grafici generati dal report",
                        wb.add_format({"bold": True, "font_size": 14}))

        for label, fig in figures.items():
            if fig is not None:
                try:
                    img_bytes = fig.to_image(format="png", width=1000, height=500, scale=2)
                    img_buf = io.BytesIO(img_bytes)
                    ws_charts.write(row_offset, 0, label, wb.add_format({"bold": True}))
                    ws_charts.insert_image(row_offset + 1, 0, label,
                                           {"image_data": img_buf, "x_scale": 0.5, "y_scale": 0.5})
                    row_offset += 28
                except Exception:
                    ws_charts.write(row_offset, 0, f"{label}: errore generazione immagine")
                    row_offset += 2

    return buf.getvalue()


def _get_date_range(df):
    dates = df[COL_DATE].dropna()
    if dates.empty:
        return "N/A"
    return f"{dates.min().strftime('%d/%m/%Y')} → {dates.max().strftime('%d/%m/%Y')}"


# ══════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════
def main():
    with st.sidebar:
        st.markdown("## ⏱️ SAP Ore Analyzer")
        st.markdown("---")

        uploaded_file = st.file_uploader(
            "📂 Carica export SAP (.xlsx)",
            type=["xlsx", "xls"],
            help="File Excel esportato da SAP con le ore giornaliere",
        )

        if uploaded_file is None:
            st.info("Carica un file per iniziare")
            st.stop()

        try:
            df_raw = load_and_clean(uploaded_file)
        except Exception as e:
            st.error(f"Errore nel caricamento: {e}")
            st.stop()

        st.success(f"✅ {len(df_raw)} righe caricate")
        st.markdown("---")

        # ── FILTRO 1: Reparto (C) ──
        st.markdown("### 🏭 Reparto")
        all_reparti = sorted(df_raw[COL_REPARTO].unique())
        default_reparti = [r for r in ["UTE", "UTES"] if r in all_reparti]
        sel_reparti = st.multiselect(
            "Seleziona reparti", options=all_reparti,
            default=default_reparti if default_reparti else all_reparti,
            help="Filtra per codice reparto (colonna C)",
        )

        df_after_rep = df_raw[df_raw[COL_REPARTO].isin(sel_reparti)] if sel_reparti else df_raw

        # ── FILTRO 2: WBS (A) ──
        st.markdown("### 📋 WBS")
        all_wbs = sorted(df_after_rep[COL_WBS].unique())
        sel_wbs = st.multiselect(
            "Seleziona WBS", options=all_wbs, default=[],
            help="Lascia vuoto per tutte le WBS del reparto selezionato",
            placeholder="Tutte le WBS",
        )

        df_after_wbs = df_after_rep
        if sel_wbs:
            df_after_wbs = df_after_rep[df_after_rep[COL_WBS].isin(sel_wbs)]

        # ── FILTRO 3: Tipo Attività (parsed from N) ──
        st.markdown("### 🏷️ Tipo Attività")
        all_act_types = sorted(df_after_wbs[COL_ACT_TYPE].unique())
        sel_act_types = st.multiselect(
            "Seleziona tipo attività", options=all_act_types, default=[],
            help="Prefisso attività (SW, HW, RI, ecc.). Lascia vuoto per tutti.",
            placeholder="Tutti i tipi",
            format_func=lambda t: f"{t} — {ACTIVITY_TYPES.get(t, '?')}",
        )

        df_after_act = df_after_wbs
        if sel_act_types:
            df_after_act = df_after_wbs[df_after_wbs[COL_ACT_TYPE].isin(sel_act_types)]

        # ── FILTRO 4: Persone (H) ──
        st.markdown("### 👷 Persone")
        all_persons = sorted(df_after_act[COL_PERSON].unique())
        sel_persons = st.multiselect(
            "Seleziona persone", options=all_persons, default=[],
            help="Lascia vuoto per tutte le persone",
            placeholder="Tutte le persone",
        )

        st.markdown("---")

        # ── TARGET ──
        st.markdown("### 🎯 Target Ore")
        target = st.number_input(
            "Inserisci target (h)", min_value=0.0, value=0.0,
            step=10.0, format="%.1f",
            help="Ore budget previste per il confronto",
        )

    # ── Apply all filters ──
    df_filtered = apply_filters(df_raw, sel_reparti, sel_wbs, sel_persons, sel_act_types)

    if df_filtered.empty:
        st.warning("⚠️ Nessun dato corrisponde ai filtri selezionati.")
        st.stop()

    total_hours = df_filtered[COL_HOURS].sum()
    n_persons = df_filtered[COL_PERSON].nunique()
    n_wbs = df_filtered[COL_WBS].nunique()
    n_days = df_filtered[COL_DATE].nunique()
    n_act_types = df_filtered[COL_ACT_TYPE].nunique()
    date_range = _get_date_range(df_filtered)

    # ══════════════════════════════════════════
    # MAIN CONTENT
    # ══════════════════════════════════════════
    st.markdown("# ⏱️ Analisi Ore SAP")
    st.caption(f"Periodo: **{date_range}**  ·  Righe filtrate: **{len(df_filtered)}**")

    # ── KPI CARDS ──
    cols = st.columns(5)
    with cols[0]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Ore Totali</h3>
            <p>{total_hours:,.1f}h</p>
            <div class="sub">{len(df_filtered)} registrazioni</div>
        </div>""", unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Persone</h3>
            <p>{n_persons}</p>
            <div class="sub">coinvolte</div>
        </div>""", unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>WBS Attive</h3>
            <p>{n_wbs}</p>
            <div class="sub">work breakdown</div>
        </div>""", unsafe_allow_html=True)
    with cols[3]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Tipi Attività</h3>
            <p>{n_act_types}</p>
            <div class="sub">categorie</div>
        </div>""", unsafe_allow_html=True)
    with cols[4]:
        avg_day = total_hours / n_days if n_days > 0 else 0
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Media / Giorno</h3>
            <p>{avg_day:,.1f}h</p>
            <div class="sub">{n_days} giorni lavorativi</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("")

    # ── TARGET vs REALE ──
    fig_target = None
    if target > 0:
        st.markdown("### 🎯 Target vs Reale")
        fig_target = make_target_chart(total_hours, target)
        st.plotly_chart(fig_target, use_container_width=True)

    # ── DONUT + STACKED BAR (side by side) ──
    st.markdown("### 🏷️ Analisi per Tipo Attività")
    col_left, col_right = st.columns([1, 1])

    fig_donut = make_donut_activity(df_filtered)
    with col_left:
        st.plotly_chart(fig_donut, use_container_width=True)

    fig_stacked = make_stacked_bar_activity(df_filtered)
    with col_right:
        st.plotly_chart(fig_stacked, use_container_width=True)

    # ── HEATMAP Persone × WBS ──
    st.markdown("### 🔥 Heatmap Persone × WBS")
    fig_heatmap_wbs = make_heatmap(df_filtered, COL_PERSON, COL_WBS,
                                    "Heatmap Ore — Persone × WBS")
    st.plotly_chart(fig_heatmap_wbs, use_container_width=True)

    # ── HEATMAP Persone × Tipo Attività ──
    st.markdown("### 🔥 Heatmap Persone × Tipo Attività")
    fig_heatmap_act = make_heatmap(df_filtered, COL_PERSON, COL_ACT_TYPE,
                                    "Heatmap Ore — Persone × Tipo Attività")
    st.plotly_chart(fig_heatmap_act, use_container_width=True)

    # ── ORE PER PERSONA ──
    st.markdown("### 👷 Ore per Persona")
    fig_person = make_hours_by_person_chart(df_filtered)
    st.plotly_chart(fig_person, use_container_width=True)

    # ── DATA TABLE ──
    with st.expander("📊 Dati filtrati (espandi per vedere)", expanded=False):
        display_cols = [COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON,
                        COL_HOURS, COL_ACT_TYPE, COL_ACT_DETAIL, COL_DESC]
        available_display = [c for c in display_cols if c in df_filtered.columns]
        st.dataframe(
            df_filtered[available_display].sort_values(COL_DATE),
            use_container_width=True, height=400,
            column_config={
                COL_ACT_TYPE: st.column_config.TextColumn("Tipo Att."),
                COL_ACT_DETAIL: st.column_config.TextColumn("Dettaglio"),
            },
        )

    # ── EXPORT ──
    st.markdown("---")
    st.markdown("### 📥 Esporta Report")

    all_figures = {
        "Target vs Reale": fig_target,
        "Distribuzione Tipo Attività": fig_donut,
        "Breakdown Persona x Tipo": fig_stacked,
        "Heatmap Persone x WBS": fig_heatmap_wbs,
        "Heatmap Persone x Tipo Attività": fig_heatmap_act,
        "Ore per Persona": fig_person,
    }

    report_bytes = generate_excel_report(
        df_filtered,
        sel_reparti, sel_wbs, sel_persons, sel_act_types,
        total_hours, target,
        all_figures,
    )

    st.download_button(
        label="📥 Scarica Report Excel",
        data=report_bytes,
        file_name=f"SAP_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.document",
        type="primary",
        use_container_width=True,
    )


if __name__ == "__main__":
    import sys
    from streamlit.web import cli as stcli
    from streamlit.runtime import exists as runtime_exists

    if not runtime_exists():
        sys.argv = ["streamlit", "run", __file__]
        sys.exit(stcli.main())
    else:
        main()