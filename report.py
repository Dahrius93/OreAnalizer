# report.py — generazione del report Excel multi-foglio scaricabile.
# Richiede xlsxwriter (scrittura) e kaleido (esportazione grafici come PNG).
# Espone: generate_excel_report.

import io
from datetime import datetime

import pandas as pd

from config import (
    COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON, COL_HOURS,
    COL_DESC, COL_ACT_TYPE, COL_ACT_DETAIL, COL_NETWORK,
    ACTIVITY_TYPES,
)
from data import _get_date_range


def generate_excel_report(
    df_filtered, reparti, wbs_list, persone, act_types,
    actual, target, figures,
):
    """Genera il report Excel multi-foglio e restituisce i byte del file.

    Fogli prodotti:
      1. Riepilogo       — parametri filtro, ore totali, target, periodo
      2. Ore per Persona — totale ore, giorni lavorati, media ore/giorno
      3. Ore per WBS     — totale ore e registrazioni per WBS
      4. Ore per Tipo Attività — ore, registrazioni, % per tipo
      5. Dettaglio       — righe filtrate con colonne principali
      6. Grafici         — immagini PNG dei grafici Plotly (richiede kaleido)
    """
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
