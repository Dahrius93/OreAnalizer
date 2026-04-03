import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
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
    /* KPI cards */
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

    /* Status badges */
    .badge-ok { color: #4ade80; }
    .badge-warn { color: #fbbf24; }
    .badge-over { color: #f87171; }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #0f172a;
    }
    section[data-testid="stSidebar"] .stMarkdown h1,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# COLUMN MAPPING (SAP export structure)
# ──────────────────────────────────────────────
COL_WBS = "Testo breve"                          # A
COL_CREATED = "Creato il"                         # B
COL_REPARTO = "Testo contab."                     # C
COL_DATE = "Data"                                 # G
COL_PERSON = "Nome del dipendente o del candidato"  # H
COL_PERIOD = "Periodo"                            # I
COL_YEAR = "Esercizio"                            # J
COL_HOURS = "Numero (un. di mis.)"                # K
COL_NETWORK = "Network"                           # M
COL_DESC = "Descrizione attività"                 # N

REQUIRED_COLS = [COL_WBS, COL_REPARTO, COL_PERSON, COL_HOURS]


# ──────────────────────────────────────────────
# DATA LOADING & CLEANING
# ──────────────────────────────────────────────
@st.cache_data
def load_and_clean(uploaded_file) -> pd.DataFrame:
    """Load SAP Excel export and drop summary/subtotal rows."""
    df = pd.read_excel(uploaded_file)

    # Drop rows where WBS (col A) is empty → SAP subtotal / total rows
    df = df.dropna(subset=[COL_WBS])
    # Also drop rows where Reparto or Person is empty
    df = df.dropna(subset=[COL_REPARTO, COL_PERSON])

    # Normalize types
    df[COL_HOURS] = pd.to_numeric(df[COL_HOURS], errors="coerce").fillna(0)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    df[COL_WBS] = df[COL_WBS].astype(str).str.strip()
    df[COL_REPARTO] = df[COL_REPARTO].astype(str).str.strip().str.upper()
    df[COL_PERSON] = df[COL_PERSON].astype(str).str.strip().str.title()

    return df


def apply_filters(df: pd.DataFrame, reparti, wbs_list, persone) -> pd.DataFrame:
    """Apply cascading filters."""
    filtered = df.copy()
    if reparti:
        filtered = filtered[filtered[COL_REPARTO].isin(reparti)]
    if wbs_list:
        filtered = filtered[filtered[COL_WBS].isin(wbs_list)]
    if persone:
        filtered = filtered[filtered[COL_PERSON].isin(persone)]
    return filtered


# ──────────────────────────────────────────────
# CHARTS
# ──────────────────────────────────────────────
def make_target_chart(actual: float, target: float) -> go.Figure:
    """Horizontal bar: actual vs target with color coding."""
    pct = (actual / target * 100) if target > 0 else 0
    if pct <= 80:
        color = "#4ade80"
        status = "OK"
    elif pct <= 100:
        color = "#fbbf24"
        status = "ATTENZIONE"
    else:
        color = "#f87171"
        status = "SUPERATO"

    fig = go.Figure()

    # Target bar (background)
    fig.add_trace(go.Bar(
        y=[""],
        x=[target],
        orientation="h",
        marker=dict(color="rgba(100,116,139,0.3)", line=dict(color="#64748b", width=1)),
        name="Target",
        text=[f"Target: {target:.1f}h"],
        textposition="inside",
        textfont=dict(color="#94a3b8", size=13),
    ))

    # Actual bar
    fig.add_trace(go.Bar(
        y=[""],
        x=[actual],
        orientation="h",
        marker=dict(color=color, line=dict(color=color, width=1)),
        name="Reale",
        text=[f"Reale: {actual:.1f}h ({pct:.0f}%)"],
        textposition="inside",
        textfont=dict(color="#0f172a", size=14, family="Arial Black"),
    ))

    fig.update_layout(
        barmode="overlay",
        height=140,
        margin=dict(l=20, r=20, t=35, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#94a3b8")),
        xaxis=dict(
            showgrid=False, zeroline=False,
            tickfont=dict(color="#64748b"), title=None,
            range=[0, max(actual, target) * 1.15],
        ),
        yaxis=dict(showticklabels=False),
        title=dict(
            text=f"<b>Status: {status}</b>  —  Δ = {target - actual:+.1f}h",
            font=dict(color=color, size=14),
            x=0.5,
        ),
    )
    return fig


def make_heatmap(df: pd.DataFrame) -> go.Figure:
    """Heatmap: Persone (y) × WBS (x), valore = ore totali."""
    pivot = df.pivot_table(
        index=COL_PERSON, columns=COL_WBS, values=COL_HOURS,
        aggfunc="sum", fill_value=0,
    )

    # Sort: most hours on top
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=True).index]

    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale=[
            [0, "#0f172a"],
            [0.01, "#1e3a5f"],
            [0.25, "#2563eb"],
            [0.5, "#7c3aed"],
            [0.75, "#db2777"],
            [1, "#f43f5e"],
        ],
        text=pivot.values,
        texttemplate="%{text:.1f}h",
        textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>WBS: %{x}<br>Ore: %{z:.1f}h<extra></extra>",
        colorbar=dict(
            title=dict(text="Ore", font=dict(color="#94a3b8")),
            tickfont=dict(color="#94a3b8"),
        ),
    ))

    n_wbs = len(pivot.columns)
    n_persons = len(pivot.index)
    height = max(350, n_persons * 45 + 100)

    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=40, b=80),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            tickangle=-45, tickfont=dict(color="#cbd5e1", size=10),
            side="bottom", title=None,
        ),
        yaxis=dict(
            tickfont=dict(color="#cbd5e1", size=11), title=None,
            autorange="reversed",
        ),
        title=dict(
            text="<b>Heatmap Ore — Persone × WBS</b>",
            font=dict(color="#e2e8f0", size=15),
            x=0.5,
        ),
    )
    return fig


def make_hours_by_person_chart(df: pd.DataFrame) -> go.Figure:
    """Horizontal bar chart: hours per person."""
    person_hours = (
        df.groupby(COL_PERSON)[COL_HOURS]
        .sum()
        .sort_values(ascending=True)
    )

    fig = go.Figure(go.Bar(
        y=person_hours.index,
        x=person_hours.values,
        orientation="h",
        marker=dict(
            color=person_hours.values,
            colorscale=[[0, "#2563eb"], [1, "#f43f5e"]],
        ),
        text=[f"{v:.1f}h" for v in person_hours.values],
        textposition="auto",
        textfont=dict(color="#f1f5f9", size=12),
    ))

    height = max(300, len(person_hours) * 35 + 80)

    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(
            showgrid=True, gridcolor="rgba(100,116,139,0.2)",
            tickfont=dict(color="#64748b"), title=dict(text="Ore", font=dict(color="#94a3b8")),
        ),
        yaxis=dict(tickfont=dict(color="#cbd5e1", size=11)),
        title=dict(
            text="<b>Ore per Persona</b>",
            font=dict(color="#e2e8f0", size=15),
            x=0.5,
        ),
    )
    return fig


# ──────────────────────────────────────────────
# REPORT EXPORT
# ──────────────────────────────────────────────
def generate_excel_report(
    df_filtered: pd.DataFrame,
    reparti, wbs_list, persone,
    actual: float, target: float,
    fig_target: go.Figure | None,
    fig_heatmap: go.Figure | None,
    fig_person: go.Figure | None,
) -> bytes:
    """Generate a multi-sheet Excel report."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── Sheet 1: Riepilogo ──
        summary_data = {
            "Parametro": [
                "Reparti selezionati",
                "WBS selezionate",
                "Persone selezionate",
                "Ore Totali Reali",
                "Target",
                "Delta (Target - Reale)",
                "Periodo dati",
                "Report generato il",
            ],
            "Valore": [
                ", ".join(reparti) if reparti else "Tutti",
                ", ".join(wbs_list) if wbs_list else "Tutte",
                ", ".join(persone) if persone else "Tutte",
                f"{actual:.2f} h",
                f"{target:.2f} h" if target > 0 else "Non impostato",
                f"{target - actual:+.2f} h" if target > 0 else "N/A",
                _get_date_range(df_filtered),
                datetime.now().strftime("%d/%m/%Y %H:%M"),
            ],
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name="Riepilogo", index=False)

        ws_summary = writer.sheets["Riepilogo"]
        header_fmt = wb.add_format({"bold": True, "bg_color": "#1e293b", "font_color": "#e2e8f0", "border": 1})
        for col_num, value in enumerate(df_summary.columns):
            ws_summary.write(0, col_num, value, header_fmt)
        ws_summary.set_column("A:A", 28)
        ws_summary.set_column("B:B", 50)

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

        ws_person = writer.sheets["Ore per Persona"]
        for col_num, value in enumerate(person_summary.columns):
            ws_person.write(0, col_num, value, header_fmt)
        ws_person.set_column("A:A", 30)
        ws_person.set_column("B:D", 16)

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

        ws_wbs = writer.sheets["Ore per WBS"]
        for col_num, value in enumerate(wbs_summary.columns):
            ws_wbs.write(0, col_num, value, header_fmt)
        ws_wbs.set_column("A:A", 25)
        ws_wbs.set_column("B:C", 16)

        # ── Sheet 4: Dettaglio righe ──
        export_cols = [COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON, COL_HOURS, COL_DESC, COL_NETWORK]
        available_cols = [c for c in export_cols if c in df_filtered.columns]
        df_detail = df_filtered[available_cols].copy()
        df_detail.to_excel(writer, sheet_name="Dettaglio", index=False)

        ws_detail = writer.sheets["Dettaglio"]
        for col_num, value in enumerate(df_detail.columns):
            ws_detail.write(0, col_num, value, header_fmt)

        # ── Sheet 5: Grafici (come immagini) ──
        ws_charts = wb.add_worksheet("Grafici")
        row_offset = 1
        ws_charts.write(0, 0, "Grafici generati dal report", wb.add_format({"bold": True, "font_size": 14}))

        for label, fig in [("Target vs Reale", fig_target), ("Heatmap", fig_heatmap), ("Ore per Persona", fig_person)]:
            if fig is not None:
                try:
                    img_bytes = fig.to_image(format="png", width=1000, height=500, scale=2)
                    img_buf = io.BytesIO(img_bytes)
                    ws_charts.write(row_offset, 0, label, wb.add_format({"bold": True}))
                    ws_charts.insert_image(row_offset + 1, 0, label, {"image_data": img_buf, "x_scale": 0.5, "y_scale": 0.5})
                    row_offset += 28
                except Exception:
                    ws_charts.write(row_offset, 0, f"{label}: errore generazione immagine")
                    row_offset += 2

    return buf.getvalue()


def _get_date_range(df: pd.DataFrame) -> str:
    dates = df[COL_DATE].dropna()
    if dates.empty:
        return "N/A"
    return f"{dates.min().strftime('%d/%m/%Y')} → {dates.max().strftime('%d/%m/%Y')}"


# ══════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════
def main():
    # ── Sidebar ──
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

        # Load data
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
        # Pre-select UTE/UTES if present
        default_reparti = [r for r in ["UTE", "UTES"] if r in all_reparti]
        sel_reparti = st.multiselect(
            "Seleziona reparti",
            options=all_reparti,
            default=default_reparti if default_reparti else all_reparti,
            help="Filtra per codice reparto (colonna C)",
        )

        # Cascading: filter WBS based on selected reparti
        df_after_rep = df_raw[df_raw[COL_REPARTO].isin(sel_reparti)] if sel_reparti else df_raw

        # ── FILTRO 2: WBS (A) ──
        st.markdown("### 📋 WBS")
        all_wbs = sorted(df_after_rep[COL_WBS].unique())
        sel_wbs = st.multiselect(
            "Seleziona WBS",
            options=all_wbs,
            default=[],
            help="Lascia vuoto per tutte le WBS del reparto selezionato",
            placeholder="Tutte le WBS",
        )

        # Cascading: filter persons based on reparti + WBS
        df_after_wbs = df_after_rep
        if sel_wbs:
            df_after_wbs = df_after_rep[df_after_rep[COL_WBS].isin(sel_wbs)]

        # ── FILTRO 3: Persone (H) ──
        st.markdown("### 👷 Persone")
        all_persons = sorted(df_after_wbs[COL_PERSON].unique())
        sel_persons = st.multiselect(
            "Seleziona persone",
            options=all_persons,
            default=[],
            help="Lascia vuoto per tutte le persone",
            placeholder="Tutte le persone",
        )

        st.markdown("---")

        # ── TARGET ──
        st.markdown("### 🎯 Target Ore")
        target = st.number_input(
            "Inserisci target (h)",
            min_value=0.0,
            value=0.0,
            step=10.0,
            format="%.1f",
            help="Ore budget previste per il confronto",
        )

    # ── Apply all filters ──
    df_filtered = apply_filters(df_raw, sel_reparti, sel_wbs, sel_persons)

    if df_filtered.empty:
        st.warning("⚠️ Nessun dato corrisponde ai filtri selezionati.")
        st.stop()

    total_hours = df_filtered[COL_HOURS].sum()
    n_persons = df_filtered[COL_PERSON].nunique()
    n_wbs = df_filtered[COL_WBS].nunique()
    n_days = df_filtered[COL_DATE].nunique()
    date_range = _get_date_range(df_filtered)

    # ══════════════════════════════════════════
    # MAIN CONTENT
    # ══════════════════════════════════════════
    st.markdown(f"# ⏱️ Analisi Ore SAP")
    st.caption(f"Periodo: **{date_range}**  ·  Righe filtrate: **{len(df_filtered)}**")

    # ── KPI CARDS ──
    cols = st.columns(4)
    with cols[0]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Ore Totali</h3>
            <p>{total_hours:,.1f}h</p>
            <div class="sub">{len(df_filtered)} registrazioni</div>
        </div>
        """, unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Persone</h3>
            <p>{n_persons}</p>
            <div class="sub">coinvolte</div>
        </div>
        """, unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f"""
        <div class="kpi-card">
            <h3>WBS Attive</h3>
            <p>{n_wbs}</p>
            <div class="sub">work breakdown</div>
        </div>
        """, unsafe_allow_html=True)
    with cols[3]:
        avg_day = total_hours / n_days if n_days > 0 else 0
        st.markdown(f"""
        <div class="kpi-card">
            <h3>Media / Giorno</h3>
            <p>{avg_day:,.1f}h</p>
            <div class="sub">{n_days} giorni lavorativi</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    # ── TARGET vs REALE ──
    fig_target = None
    if target > 0:
        st.markdown("### 🎯 Target vs Reale")
        fig_target = make_target_chart(total_hours, target)
        st.plotly_chart(fig_target, use_container_width=True)

    # ── HEATMAP ──
    st.markdown("### 🔥 Heatmap Persone × WBS")
    fig_heatmap = make_heatmap(df_filtered)
    st.plotly_chart(fig_heatmap, use_container_width=True)

    # ── ORE PER PERSONA ──
    st.markdown("### 👷 Ore per Persona")
    fig_person = make_hours_by_person_chart(df_filtered)
    st.plotly_chart(fig_person, use_container_width=True)

    # ── DATA TABLE ──
    with st.expander("📊 Dati filtrati (espandi per vedere)", expanded=False):
        display_cols = [COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON, COL_HOURS, COL_DESC]
        available_display = [c for c in display_cols if c in df_filtered.columns]
        st.dataframe(
            df_filtered[available_display].sort_values(COL_DATE),
            use_container_width=True,
            height=400,
        )

    # ── EXPORT ──
    st.markdown("---")
    st.markdown("### 📥 Esporta Report")

    report_bytes = generate_excel_report(
        df_filtered,
        sel_reparti, sel_wbs, sel_persons,
        total_hours, target,
        fig_target, fig_heatmap, fig_person,
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
    main()