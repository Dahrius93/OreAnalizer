# charts.py — generazione di tutti i grafici Plotly dell'applicazione.
# Ogni funzione riceve un DataFrame già filtrato e restituisce una Figure Plotly.
# Espone: make_target_chart, make_heatmap, make_hours_by_person_chart,
#         make_donut_activity, make_stacked_bar_activity.

import plotly.graph_objects as go

from config import COL_HOURS, COL_PERSON, COL_ACT_TYPE, ACTIVITY_TYPES, TYPE_COLORS


def make_target_chart(actual, target):
    """Barra orizzontale sovrapposta che confronta ore reali con il target.
    Il colore del titolo indica lo stato: verde ≤80%, giallo ≤100%, rosso >100%.
    """
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
    """Barra orizzontale delle ore totali per persona, ordinate crescenti."""
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
