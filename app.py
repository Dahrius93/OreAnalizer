# App.py — entry point dell'applicazione Streamlit.
# Gestisce: configurazione pagina, CSS custom, sidebar con filtri e target,
#           rendering dei grafici e del pulsante di export.
# Tutta la logica di dati, grafici e report è delegata ai moduli dedicati.

import streamlit as st
from datetime import datetime

from config import (
    COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON, COL_HOURS,
    COL_DESC, COL_ACT_TYPE, COL_ACT_DETAIL,
    ACTIVITY_TYPES,
)
from data import load_and_clean, apply_filters, _get_date_range
from charts import (
    make_target_chart, make_heatmap,
    make_hours_by_person_chart, make_donut_activity, make_stacked_bar_activity,
)
from report import generate_excel_report

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


# ══════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════
def main():
    with st.sidebar:
        st.markdown("## ⏱️ SAP Ore Analyzer")
        st.markdown("v1.2.1")
        st.markdown("---")

        uploaded_file = st.file_uploader(
            "📂 Carica export SAP (.xlsx)",
            type=["xlsx", "xls"],
            help="File Excel esportato da SAP con le ore giornaliere",
        )

        if uploaded_file is None:
            st.info("Carica un file per iniziare")
            # La welcome page viene mostrata nel contenuto principale sotto
        else:
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

            # ── FILTRO 5: Data ──
            st.markdown("### 📅 Periodo")
            date_min = df_raw[COL_DATE].min().date()
            date_max = df_raw[COL_DATE].max().date()
            sel_date_start = st.date_input(
                "Da", value=date_min, min_value=date_min, max_value=date_max,
                format="DD/MM/YYYY",
            )
            sel_date_end = st.date_input(
                "A", value=date_max, min_value=date_min, max_value=date_max,
                format="DD/MM/YYYY",
            )

            st.markdown("---")

            # ── TARGET ──
            st.markdown("### 🎯 Target Ore")
            target = st.number_input(
                "Inserisci target (h)", min_value=0.0, value=0.0,
                step=10.0, format="%.1f",
                help="Ore budget previste per il confronto",
            )

    # ── Welcome screen quando nessun file è caricato ──
    if uploaded_file is None:
        st.markdown("# ⏱️ SAP Ore Analyzer")
        st.markdown("Benvenuto! Carica un export SAP per iniziare l'analisi delle ore.")
        st.markdown("---")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            <div class="kpi-card">
                <h3>Passo 1</h3>
                <p style="font-size:2rem">📂</p>
                <div class="sub">Carica il file Excel esportato da SAP dalla barra laterale</div>
            </div>""", unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div class="kpi-card">
                <h3>Passo 2</h3>
                <p style="font-size:2rem">🔍</p>
                <div class="sub">Applica i filtri per Reparto, WBS, Tipo Attività e Persone</div>
            </div>""", unsafe_allow_html=True)
        with col3:
            st.markdown("""
            <div class="kpi-card">
                <h3>Passo 3</h3>
                <p style="font-size:2rem">📊</p>
                <div class="sub">Analizza i grafici e scarica il report Excel</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("")
        st.markdown("---")
        st.markdown("### 📋 Cosa puoi fare con questa app")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("""
**Analisi disponibili:**
- **KPI principali** — ore totali, persone, WBS attive, media giornaliera
- **Target vs Reale** — confronta le ore consuntivate con il budget
- **Donut per tipo attività** — distribuzione percentuale SW, HW, RI, ecc.
- **Stacked bar** — breakdown per persona e tipo attività
""")
        with col_b:
            st.markdown("""
**Visualizzazioni avanzate:**
- **Heatmap Persone × WBS** — identifica chi lavora su cosa
- **Heatmap Persone × Tipo Attività** — distribuzione del lavoro per categoria
- **Ore per Persona** — confronto diretto tra le risorse
- **Export Excel** — report completo con grafici incorporati
""")

        st.markdown("---")
        st.info("**Formato atteso:** file `.xlsx` o `.xls` esportato da SAP con colonne WBS, Reparto, Data, Persona, Ore, Descrizione e Tipo Attività.")
        st.stop()

    # ── Apply all filters ──
    df_filtered = apply_filters(df_raw, sel_reparti, sel_wbs, sel_persons, sel_act_types, sel_date_start, sel_date_end)

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
                COL_DATE: st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
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
