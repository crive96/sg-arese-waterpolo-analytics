from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = BASE_DIR / "SG_Arese_2025_26_Analisi.xlsx"
LOGO_PATH = BASE_DIR / "example_logo arese.png"
TEAM_PHOTO_PATH = BASE_DIR / "prima_squadra_25_26.jpg"

ARESE_NAVY = "#07111F"
ARESE_DEEP = "#111A45"
ARESE_BLUE = "#1848D8"
ARESE_CYAN = "#20C7E8"
ARESE_AQUA = "#7DE8F2"
ARESE_ORANGE = "#FF7A00"
ARESE_WHITE = "#F8FAFC"
ARESE_MUTED = "#A7B3C5"
ARESE_GREEN = "#19D3A2"
ARESE_RED = "#FF4B5C"
ARESE_GOLD = "#FFC857"

PLOTLY_TEMPLATE = "plotly_dark"
PLAYER_COLOR_MAP = {
    "1-2": ARESE_BLUE,
    "3": ARESE_ORANGE,
    "4-5": ARESE_CYAN,
    "6": ARESE_AQUA,
    "GK": ARESE_GOLD,
}
RESULT_COLOR_MAP = {"V": ARESE_GREEN, "N": ARESE_GOLD, "P": ARESE_RED}


st.set_page_config(
    page_title="SG Arese Waterpolo | Analytics",
    page_icon="SG",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_css() -> None:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:
                radial-gradient(circle at 20% 20%, rgba(32, 199, 232, 0.16), transparent 28%),
                radial-gradient(circle at 80% 5%, rgba(255, 122, 0, 0.10), transparent 22%),
                linear-gradient(135deg, {ARESE_NAVY} 0%, #080A23 48%, {ARESE_DEEP} 100%);
            color: {ARESE_WHITE};
        }}
        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #050A13 0%, #0A1430 100%);
            border-right: 1px solid rgba(32, 199, 232, 0.22);
        }}
        [data-testid="stSidebar"] * {{ color: {ARESE_WHITE}; }}
        h1, h2, h3 {{ color: {ARESE_WHITE}; letter-spacing: -0.03em; }}
        .hero {{
            padding: 28px 30px;
            border: 1px solid rgba(32, 199, 232, 0.35);
            border-radius: 28px;
            background:
                linear-gradient(120deg, rgba(17, 26, 69, 0.92), rgba(7, 17, 31, 0.82)),
                radial-gradient(circle at 75% 35%, rgba(24, 72, 216, 0.42), transparent 35%);
            box-shadow: 0 24px 60px rgba(0, 0, 0, 0.38);
            margin-bottom: 22px;
        }}
        .hero-kicker {{
            color: {ARESE_CYAN};
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-size: 0.78rem;
        }}
        .hero-title {{
            color: {ARESE_WHITE};
            font-weight: 900;
            font-size: clamp(2.2rem, 5vw, 4.2rem);
            line-height: 0.95;
            margin: 8px 0 12px 0;
        }}
        .hero-subtitle {{
            color: {ARESE_MUTED};
            font-size: 1.05rem;
            max-width: 900px;
        }}
        .metric-card {{
            padding: 18px 18px 16px 18px;
            border-radius: 20px;
            border: 1px solid rgba(125, 232, 242, 0.22);
            background: linear-gradient(145deg, rgba(17, 26, 69, 0.88), rgba(7, 17, 31, 0.72));
            box-shadow: 0 14px 35px rgba(0, 0, 0, 0.28);
            min-height: 118px;
        }}
        .metric-label {{ color: {ARESE_MUTED}; font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.11em; }}
        .metric-value {{ color: {ARESE_WHITE}; font-size: 2.15rem; font-weight: 900; margin-top: 7px; }}
        .metric-note {{ color: {ARESE_CYAN}; font-size: 0.85rem; margin-top: 4px; }}
        .section-card {{
            padding: 18px;
            border: 1px solid rgba(32, 199, 232, 0.20);
            border-radius: 22px;
            background: rgba(7, 17, 31, 0.48);
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
        }}
        .pill {{
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            color: {ARESE_WHITE};
            font-weight: 800;
            font-size: 0.75rem;
            margin-right: 6px;
            border: 1px solid rgba(255,255,255,0.12);
        }}
        .pill-win {{ background: rgba(25, 211, 162, 0.20); color: {ARESE_GREEN}; }}
        .pill-draw {{ background: rgba(255, 200, 87, 0.18); color: {ARESE_GOLD}; }}
        .pill-loss {{ background: rgba(255, 75, 92, 0.18); color: {ARESE_RED}; }}
        div[data-testid="stDataFrame"] {{ border: 1px solid rgba(32, 199, 232, 0.16); border-radius: 16px; }}
        .sidebar-photo img {{
            border-radius: 18px;
            border: 1px solid rgba(32, 199, 232, 0.26);
            box-shadow: 0 14px 32px rgba(0,0,0,0.35);
        }}
        .stTabs [data-baseweb="tab-list"] {{ gap: 8px; }}
        .stTabs [data-baseweb="tab"] {{
            background: rgba(17, 26, 69, 0.70);
            border-radius: 999px;
            color: {ARESE_WHITE};
            border: 1px solid rgba(32, 199, 232, 0.18);
        }}
        .stTabs [aria-selected="true"] {{
            background: linear-gradient(90deg, {ARESE_BLUE}, {ARESE_CYAN});
            color: white;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_data(workbook_path: Path) -> dict[str, pd.DataFrame]:
    sheets = [
        "Partite",
        "Squadra_Stagione",
        "Squadra_Partita",
        "Giocatori_Stagione",
        "Portieri",
        "Giocatori_Partita",
        "Valutazione_Partita",
        "Valutazione_Stagione",
        "Classifiche",
        "Metodo",
    ]
    return {sheet: pd.read_excel(workbook_path, sheet_name=sheet) for sheet in sheets}


def clean_number_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if col in {"giocatore", "nome_standard", "ruolo", "ruolo_macro", "mano", "ruolo_base", "ruolo_label", "jolly", "ruoli_jolly", "note_ruolo", "match_id", "data", "ora", "luogo", "avversario", "casa_trasferta", "esito", "sheet_name", "opponent", "andata_ritorno"}:
            continue
        converted = pd.to_numeric(out[col], errors="coerce")
        if converted.notna().any():
            out[col] = converted
    return out


def prepare_data(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    prepared = {name: clean_number_columns(df) for name, df in data.items()}
    for name in ["Partite", "Squadra_Partita", "Giocatori_Partita", "Valutazione_Partita"]:
        if name in prepared and "data" in prepared[name].columns:
            prepared[name]["data"] = pd.to_datetime(prepared[name]["data"], errors="coerce")

    players = prepared["Giocatori_Stagione"].copy()
    players["produzione_offensiva"] = players[["gol_tot", "assist_tot"]].fillna(0).sum(axis=1)
    players["azioni_positive_difesa"] = players[["palla_rubata_tot", "et_pos_tot", "tr_pos_tot"]].fillna(0).sum(axis=1)
    players["rischio_tot"] = players[["palla_persa_tot", "errore_tot", "controfallo_tot"]].fillna(0).sum(axis=1)
    players["impatto_positivo"] = players[["gol_tot", "assist_tot", "palla_rubata_tot", "et_pos_tot", "tr_pos_tot"]].fillna(0).sum(axis=1)
    players["profilo"] = players["ruolo_base"].fillna("") + " · " + players["ruolo_label"].fillna("")
    prepared["Giocatori_Stagione"] = players

    gk = prepared["Portieri"].copy()
    gk["carico_portiere"] = gk[["parata_tot", "gol_subito_tot"]].fillna(0).sum(axis=1)
    prepared["Portieri"] = gk

    matches = prepared["Partite"].copy()
    matches["diff"] = matches["arese_gol"] - matches["avversario_gol"]
    matches["risultato"] = matches["arese_gol"].astype(str) + "-" + matches["avversario_gol"].astype(str)
    matches["punti"] = matches["esito"].map({"V": 3, "N": 1, "P": 0}).fillna(0)
    matches["punti_cumulati"] = matches.sort_values("giornata")["punti"].cumsum()
    prepared["Partite"] = matches
    return prepared


def format_pct(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{value:.1%}"


def mean_pct(series: pd.Series) -> float | None:
    values = pd.to_numeric(series, errors="coerce").dropna()
    if values.empty:
        return None
    return float(values.mean())


def format_num(value: float | int | None, decimals: int = 1) -> str:
    if value is None or pd.isna(value):
        return "-"
    if decimals == 0:
        return f"{value:,.0f}".replace(",", ".")
    return f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_table(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if col.startswith("pct_") or col in {"% tiro media", "% tiro stagionale", "% parate"}:
            out[col] = out[col].apply(lambda value: format_pct(value) if pd.notna(value) else "-")
        elif pd.api.types.is_float_dtype(out[col]):
            out[col] = out[col].apply(lambda value: format_num(value, 2) if pd.notna(value) else "-")
    return out


def metric_card(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label">{label}</div>
          <div class="metric-value">{value}</div>
          <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def team_special_situations(matches: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "giornata",
        "data",
        "avversario",
        "casa_trasferta",
        "esito",
        "gol_fatti",
        "gol_subiti",
        "pct_et_pos",
        "pct_et_neg",
        "pct_tr_pos",
        "pct_tr_neg",
        "tr_plus_team",
        "tr_minus_team",
        "palla_al_centro",
    ]
    return matches[[col for col in cols if col in matches.columns]].sort_values("giornata").copy()


def hero(title: str, subtitle: str, kicker: str = "SG Arese Waterpolo Analytics") -> None:
    st.markdown(
        f"""
        <div class="hero">
          <div class="hero-kicker">{kicker}</div>
          <div class="hero-title">{title}</div>
          <div class="hero-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def styled_plotly(fig: go.Figure, height: int = 430, *, show_legend: bool | None = None, top_margin: int = 62) -> go.Figure:
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(7,17,31,0.38)",
        font={"color": ARESE_WHITE, "family": "Inter, Arial, sans-serif"},
        title={"font": {"size": 19, "color": ARESE_WHITE}},
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
        margin={"l": 24, "r": 24, "t": top_margin, "b": 42},
        showlegend=show_legend,
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.08)", zerolinecolor="rgba(255,255,255,0.14)", tickangle=0, automargin=True)
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.08)", zerolinecolor="rgba(255,255,255,0.14)", tickangle=0, automargin=True)
    return fig


def add_selected_to_top(frame: pd.DataFrame, source: pd.DataFrame, selected_player: str | None, sort_col: str, top: int | None) -> pd.DataFrame:
    if not selected_player or not top or "giocatore" not in frame.columns or "giocatore" not in source.columns or selected_player in frame["giocatore"].tolist():
        return frame
    selected_row = source[source["giocatore"] == selected_player]
    if selected_row.empty:
        return frame
    return pd.concat([frame, selected_row], ignore_index=True).drop_duplicates("giocatore", keep="first").sort_values(sort_col, ascending=False)


def plot_bar(df: pd.DataFrame, x: str, y: str, title: str, color: str | None = None, top: int | None = None, height: int = 430, *, selected_player: str | None = None, show_legend: bool | None = None) -> go.Figure:
    frame = df.copy()
    if top:
        frame = frame.head(top)
        frame = add_selected_to_top(frame, df, selected_player, y, top)
    frame = frame.sort_values(y, ascending=True)
    is_pct = y.startswith("pct_") or "%" in y
    if selected_player and "giocatore" in frame.columns:
        frame["highlight"] = frame["giocatore"].eq(selected_player).map({True: "Selezionato", False: "Altri"})
        color = "highlight"
        color_map = {"Selezionato": ARESE_ORANGE, "Altri": ARESE_CYAN}
    else:
        color_map = PLAYER_COLOR_MAP
    fig = px.bar(
        frame,
        x=y,
        y=x,
        color=color,
        color_discrete_map=color_map,
        text=y,
        title=title,
        orientation="h",
    )
    fig.update_traces(textposition="outside", marker_line_width=0, opacity=0.92, cliponaxis=False)
    if is_pct:
        fig.update_traces(texttemplate="%{x:.2%}", hovertemplate="%{y}<br>%{x:.2%}<extra></extra>")
        fig.update_xaxes(tickformat=".0%")
    fig.update_layout(showlegend=bool(color), yaxis_title="", xaxis_title="")
    return styled_plotly(fig, height=height, show_legend=show_legend if show_legend is not None else bool(color))


def scatter_players(
    df: pd.DataFrame,
    x: str,
    y: str,
    size: str,
    color: str,
    title: str,
    hover_extra: Iterable[str] = (),
    selected_player: str | None = None,
) -> go.Figure:
    frame = df.dropna(subset=[x, y]).copy()
    if frame.empty:
        frame = pd.DataFrame({"giocatore": [], x: [], y: [], "__bubble_size": [], color: []})
    if size in frame:
        raw_size = pd.to_numeric(frame[size], errors="coerce").fillna(0)
        frame["__bubble_size"] = raw_size.abs()
        if frame["__bubble_size"].max() == 0:
            frame["__bubble_size"] = 1
    else:
        frame["__bubble_size"] = 1
    hover_cols = list(dict.fromkeys([*hover_extra, size]))
    hover_cols = [col for col in hover_cols if col in frame.columns and col not in {x, y, "giocatore"}]
    hover_data = {col: True for col in hover_cols}
    hover_data["__bubble_size"] = False
    fig = px.scatter(
        frame,
        x=x,
        y=y,
        size="__bubble_size",
        color=color,
        hover_name="giocatore",
        text="giocatore",
        color_discrete_map=PLAYER_COLOR_MAP if color == "ruolo_base" else None,
        hover_data=hover_data,
        title=title,
        size_max=42,
    )
    fig.update_traces(textposition="top center", marker={"line": {"width": 1, "color": "rgba(255,255,255,0.45)"}})
    if selected_player and "giocatore" in frame.columns and selected_player in frame["giocatore"].tolist():
        selected_frame = frame[frame["giocatore"] == selected_player]
        fig.add_trace(
            go.Scatter(
                x=selected_frame[x],
                y=selected_frame[y],
                mode="markers+text",
                text=selected_frame["giocatore"],
                textposition="top center",
                name=f"Selezionato: {selected_player}",
                marker={"size": 24, "color": ARESE_ORANGE, "line": {"width": 3, "color": ARESE_WHITE}},
                hovertemplate=f"{selected_player}<br>{x}: %{{x}}<br>{y}: %{{y}}<extra></extra>",
            )
        )
    fig.update_layout(xaxis_title=x, yaxis_title=y)
    return styled_plotly(fig, height=520)


def filter_players_sidebar(players: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("### Filtri giocatori")
    role_options = sorted(players["ruolo_base"].dropna().unique().tolist())
    selected_roles = st.sidebar.multiselect("Ruoli", role_options, default=role_options)
    hand_options = sorted(players["mano"].dropna().unique().tolist())
    selected_hands = st.sidebar.multiselect("Mano", hand_options, default=hand_options)
    jolly_options = sorted(players["jolly"].dropna().unique().tolist())
    selected_jolly = st.sidebar.multiselect("Jolly", jolly_options, default=jolly_options)
    min_games = st.sidebar.slider("Min partite", 0, int(players["partite_nel_foglio"].max()), 3)
    return players[
        players["ruolo_base"].isin(selected_roles)
        & players["mano"].isin(selected_hands)
        & players["jolly"].isin(selected_jolly)
        & (players["partite_nel_foglio"].fillna(0) >= min_games)
    ].copy()


def sidebar_header() -> str:
    if TEAM_PHOTO_PATH.exists():
        st.sidebar.markdown('<div class="sidebar-photo">', unsafe_allow_html=True)
        st.sidebar.image(str(TEAM_PHOTO_PATH), use_column_width=True)
        st.sidebar.markdown('</div>', unsafe_allow_html=True)
    elif LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), use_column_width=True)
    st.sidebar.markdown("## SG Arese Analytics")
    st.sidebar.caption("Serie C Lombardia · stagione 2025-26")
    return st.sidebar.radio(
        "Menu",
        [
            "Overview",
            "Squadra",
            "Giocatori",
            "Scheda Giocatore",
            "Confronto Giocatori",
            "Scatter Lab",
            "Portieri",
            "Classifiche",
            "Metodo",
        ],
        label_visibility="collapsed",
    )


def style_results_table(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    def row_style(row: pd.Series) -> list[str]:
        color = RESULT_COLOR_MAP.get(row.get("esito"), "rgba(255,255,255,0.08)")
        return [f"background-color: {color}22; color: {ARESE_WHITE};" for _ in row]

    return df.style.apply(row_style, axis=1).map(
        lambda value: f"background-color: {RESULT_COLOR_MAP.get(value, 'transparent')}; color: #06111F; font-weight: 800;" if value in RESULT_COLOR_MAP else "",
        subset=["esito"],
    )


def overview_page(data: dict[str, pd.DataFrame]) -> None:
    matches = data["Partite"]
    team_match = data["Squadra_Partita"]
    players = data["Giocatori_Stagione"]
    movement = players[players["ruolo"] == "movimento"].copy()
    hero(
        "La stagione in una dashboard",
        "Risultati, giocatori, portieri e confronti interattivi con il look SG Arese: blu notte, acqua e arancione full-time.",
    )

    wins = int((matches["esito"] == "V").sum())
    draws = int((matches["esito"] == "N").sum())
    losses = int((matches["esito"] == "P").sum())
    cols = st.columns(5)
    with cols[0]:
        metric_card("Record", f"{wins}-{draws}-{losses}", "V-N-P")
    with cols[1]:
        metric_card("Gol fatti", format_num(matches["arese_gol"].sum(), 0), f"{format_num(matches['arese_gol'].mean(), 1)} / partita")
    with cols[2]:
        metric_card("Gol subiti", format_num(matches["avversario_gol"].sum(), 0), f"{format_num(matches['avversario_gol'].mean(), 1)} / partita")
    with cols[3]:
        metric_card("Diff. reti", format_num(matches["diff"].sum(), 0), "stagione completa")
    with cols[4]:
        metric_card("Punti", format_num(matches["punti"].sum(), 0), "classifica Arese")

    special = team_special_situations(team_match)
    s1, s2, s3 = st.columns(3)
    with s1:
        metric_card("% ET+", format_pct(mean_pct(special["pct_et_pos"])), "media dati disponibili")
    with s2:
        metric_card("% ET-", format_pct(mean_pct(special["pct_et_neg"])), "media dati disponibili")
    with s3:
        metric_card("Palle al centro", format_num(special["palla_al_centro"].sum(), 0), "totale rilevato")

    left, right = st.columns([1.45, 1])
    with left:
        fig = px.line(
            matches.sort_values("giornata"),
            x="giornata",
            y=["arese_gol", "avversario_gol"],
            markers=True,
            title="Gol fatti e subiti per giornata",
            color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE],
        )
        fig.update_traces(line={"width": 4}, marker={"size": 9})
        st.plotly_chart(styled_plotly(fig, 450), use_container_width=True)
    with right:
        top = movement.sort_values("gol_tot", ascending=False).head(7)
        fig = plot_bar(top, "giocatore", "gol_tot", "Top marcatori", "ruolo_base", height=450, show_legend=False)
        st.plotly_chart(fig, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        top_assist = movement.sort_values("assist_tot", ascending=False).head(7)
        st.plotly_chart(plot_bar(top_assist, "giocatore", "assist_tot", "Top assist", "ruolo_base", None, height=430, show_legend=False), use_container_width=True)
    with c2:
        top_val = movement.sort_values("indice_valutazione_0_100", ascending=False).head(7)
        st.plotly_chart(plot_bar(top_val, "giocatore", "indice_valutazione_0_100", "Indice valutazione", "ruolo_base", None, height=430, show_legend=False), use_container_width=True)
    with c3:
        top_pct = movement[movement["tiri_tot"] > 0].sort_values("pct_tiri_stagionale", ascending=False).head(7)
        st.plotly_chart(plot_bar(top_pct, "giocatore", "pct_tiri_stagionale", "% tiro stagionale", "ruolo_base", None, height=430, show_legend=False), use_container_width=True)

    st.markdown(
        "<span class='pill pill-win'>Vittoria</span><span class='pill pill-draw'>Pareggio</span><span class='pill pill-loss'>Sconfitta</span>",
        unsafe_allow_html=True,
    )
    results_table = matches[["giornata", "data", "avversario", "casa_trasferta", "risultato", "esito", "statistiche_disponibili"]].sort_values("giornata")
    st.dataframe(
        style_results_table(format_table(results_table)),
        use_container_width=True,
        hide_index=True,
        height=520,
    )


def team_page(data: dict[str, pd.DataFrame]) -> None:
    matches = data["Partite"]
    team_match = data["Squadra_Partita"]
    hero("Squadra", "Andamento risultati, split casa/trasferta, punti cumulati e quarti disponibili.")

    tabs = st.tabs(["Timeline", "Split", "Quarti", "Superiorità & Rigori", "Statistiche Team"])
    with tabs[0]:
        fig = px.bar(
            matches.sort_values("giornata"),
            x="giornata",
            y="diff",
            color="esito",
            color_discrete_map=RESULT_COLOR_MAP,
            hover_data=["avversario", "risultato", "casa_trasferta"],
            title="Differenza reti per giornata",
        )
        fig.add_scatter(x=matches["giornata"], y=matches["punti_cumulati"], name="Punti cumulati", mode="lines+markers", yaxis="y2", line={"color": ARESE_CYAN, "width": 4})
        fig.update_layout(yaxis2={"overlaying": "y", "side": "right", "title": "Punti cumulati"})
        st.plotly_chart(styled_plotly(fig, 520), use_container_width=True)
    with tabs[1]:
        split = matches.groupby("casa_trasferta", dropna=False).agg(
            partite=("match_id", "count"),
            vittorie=("esito", lambda s: int((s == "V").sum())),
            pareggi=("esito", lambda s: int((s == "N").sum())),
            sconfitte=("esito", lambda s: int((s == "P").sum())),
            gol_fatti=("arese_gol", "sum"),
            gol_subiti=("avversario_gol", "sum"),
        ).reset_index()
        fig = px.bar(split, x="casa_trasferta", y=["vittorie", "pareggi", "sconfitte"], barmode="group", title="Record casa/trasferta", color_discrete_sequence=[ARESE_GREEN, ARESE_GOLD, ARESE_RED])
        st.plotly_chart(styled_plotly(fig), use_container_width=True)
        st.dataframe(format_table(split), use_container_width=True, hide_index=True)
    with tabs[2]:
        q = matches[matches["parziali_quarti_disponibili"] == "SI"].copy()
        quarter_rows = []
        for quarter in range(1, 5):
            quarter_rows.append({"quarto": f"Q{quarter}", "GF": q[f"q{quarter}_arese"].sum(), "GS": q[f"q{quarter}_avversario"].sum()})
        quarter_df = pd.DataFrame(quarter_rows)
        fig = px.bar(quarter_df, x="quarto", y=["GF", "GS"], barmode="group", title="Rendimento per quarti", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE])
        st.plotly_chart(styled_plotly(fig), use_container_width=True)
    with tabs[3]:
        special = team_special_situations(team_match)
        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            metric_card("% ET+", format_pct(mean_pct(special["pct_et_pos"])), "realizzate")
        with k2:
            metric_card("% ET-", format_pct(mean_pct(special["pct_et_neg"])), "subite")
        with k3:
            metric_card("% TR+", format_pct(mean_pct(special["pct_tr_pos"])), "realizzati")
        with k4:
            metric_card("% TR-", format_pct(mean_pct(special["pct_tr_neg"])), "subiti")
        with k5:
            metric_card("Palle al centro", format_num(special["palla_al_centro"].sum(), 0), "totale")

        pct_long = special.melt(
            id_vars=["giornata", "avversario"],
            value_vars=["pct_et_pos", "pct_et_neg", "pct_tr_pos", "pct_tr_neg"],
            var_name="metrica",
            value_name="percentuale",
        ).dropna(subset=["percentuale"])
        label_map = {"pct_et_pos": "% ET+", "pct_et_neg": "% ET-", "pct_tr_pos": "% TR+", "pct_tr_neg": "% TR-"}
        pct_long["metrica"] = pct_long["metrica"].map(label_map)
        fig = px.line(pct_long, x="giornata", y="percentuale", color="metrica", markers=True, title="Percentuali ET/TR per giornata", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE, ARESE_AQUA, ARESE_GOLD])
        fig.update_traces(line={"width": 4}, marker={"size": 8})
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(styled_plotly(fig, 500), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(special, x="giornata", y="palla_al_centro", title="Palle al centro per giornata", color="esito", color_discrete_map=RESULT_COLOR_MAP, hover_data=["avversario"])
            st.plotly_chart(styled_plotly(fig, 420), use_container_width=True)
        with c2:
            fig = px.bar(special, x="giornata", y=["tr_plus_team", "tr_minus_team"], barmode="group", title="Rigori squadra: TR+ e TR-", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE])
            st.plotly_chart(styled_plotly(fig, 420), use_container_width=True)

        st.caption("Le medie percentuali sono calcolate solo sulle partite con dato disponibile.")
        st.dataframe(format_table(special), use_container_width=True, hide_index=True, height=520)
    with tabs[4]:
        st.dataframe(format_table(team_match.sort_values("giornata")), use_container_width=True, hide_index=True)


def players_page(data: dict[str, pd.DataFrame]) -> None:
    players = data["Giocatori_Stagione"]
    movement = players[players["ruolo"] == "movimento"].copy()
    filtered = filter_players_sidebar(movement)
    hero("Giocatori", "Classifiche, profili ruolo e metriche offensive/difensive dei giocatori di movimento.")

    profile_names = filtered.sort_values("giocatore")["giocatore"].tolist()
    profile_options = ["Nessuno"] + profile_names
    saved_profile = st.session_state.get("selected_player_profile")
    default_profile = saved_profile if saved_profile in profile_names else "Nessuno"
    selected_option = st.selectbox("Fissa un giocatore nei grafici", profile_options, index=profile_options.index(default_profile)) if profile_names else "Nessuno"
    selected_profile = None if selected_option == "Nessuno" else selected_option
    if selected_profile:
        st.session_state["selected_player_profile"] = selected_profile
        st.info(f"Giocatore fissato nei grafici: `{selected_profile}`. La voce `Scheda Giocatore` si aprira' su questo profilo.")
    else:
        st.session_state.pop("selected_player_profile", None)
        st.caption("Nessun giocatore fissato: i grafici mostrano la vista standard per ruolo.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Giocatori filtrati", str(len(filtered)), "movimento")
    with c2:
        metric_card("Gol", format_num(filtered["gol_tot"].sum(), 0), "totale filtro")
    with c3:
        metric_card("Assist", format_num(filtered["assist_tot"].sum(), 0), "totale filtro")
    with c4:
        metric_card("% tiro stagionale", format_pct(filtered["gol_tot"].sum() / filtered["tiri_tot"].sum() if filtered["tiri_tot"].sum() else None), "gol / tiri")

    tab1, tab2, tab3, tab4 = st.tabs(["Top KPI", "Scatter rapidi", "ET / Rigori", "Tabella"])
    with tab1:
        a, b = st.columns(2)
        with a:
            st.plotly_chart(plot_bar(filtered.sort_values("gol_tot", ascending=False), "giocatore", "gol_tot", "Gol totali", "ruolo_base", 10, height=440, selected_player=selected_profile), use_container_width=True)
            st.plotly_chart(plot_bar(filtered.sort_values("assist_tot", ascending=False), "giocatore", "assist_tot", "Assist totali", "ruolo_base", 10, height=440, selected_player=selected_profile), use_container_width=True)
        with b:
            st.plotly_chart(plot_bar(filtered.sort_values("indice_valutazione_0_100", ascending=False), "giocatore", "indice_valutazione_0_100", "Indice valutazione 0-100", "ruolo_base", 10, height=440, selected_player=selected_profile), use_container_width=True)
            pct = filtered[filtered["tiri_tot"] > 0].sort_values("pct_tiri_stagionale", ascending=False)
            st.plotly_chart(plot_bar(pct, "giocatore", "pct_tiri_stagionale", "% tiro stagionale", "ruolo_base", 10, height=440, selected_player=selected_profile), use_container_width=True)
    with tab2:
        st.plotly_chart(scatter_players(filtered, "gol_tot", "assist_tot", "partite_nel_foglio", "ruolo_base", "Produzione offensiva: gol vs assist", ["pct_tiri_stagionale", "valutazione_grezza_totale"], selected_player=selected_profile), use_container_width=True)
        st.plotly_chart(scatter_players(filtered, "pct_tiri_media", "pct_tiri_stagionale", "tiri_tot", "mano", "Consistenza vs efficienza tiro", ["gol_tot", "assist_tot", "partite_nel_foglio"], selected_player=selected_profile), use_container_width=True)
    with tab3:
        a, b = st.columns(2)
        with a:
            st.plotly_chart(plot_bar(filtered.sort_values("et_pos_tot", ascending=False), "giocatore", "et_pos_tot", "ET+ conquistate/realizzate", "ruolo_base", 10, height=420, selected_player=selected_profile), use_container_width=True)
            st.plotly_chart(plot_bar(filtered.sort_values("tr_pos_tot", ascending=False), "giocatore", "tr_pos_tot", "TR+ conquistati/realizzati", "ruolo_base", 10, height=420, selected_player=selected_profile), use_container_width=True)
        with b:
            st.plotly_chart(plot_bar(filtered.sort_values("et_neg_tot", ascending=False), "giocatore", "et_neg_tot", "ET- subite/concesse", "ruolo_base", 10, height=420, selected_player=selected_profile), use_container_width=True)
            st.plotly_chart(plot_bar(filtered.sort_values("tr_neg_tot", ascending=False), "giocatore", "tr_neg_tot", "TR- subiti/concessi", "ruolo_base", 10, height=420, selected_player=selected_profile), use_container_width=True)
        st.plotly_chart(scatter_players(filtered, "et_pos_tot", "et_neg_tot", "partite_nel_foglio", "ruolo_base", "ET+: conquistate vs subite", ["et_pos_media", "et_neg_media", "valutazione_grezza_totale"], selected_player=selected_profile), use_container_width=True)
        st.plotly_chart(scatter_players(filtered, "tr_pos_tot", "tr_neg_tot", "partite_nel_foglio", "ruolo_base", "Rigori: TR+ vs TR-", ["tr_pos_media", "tr_neg_media", "valutazione_grezza_totale"], selected_player=selected_profile), use_container_width=True)
        et_cols = ["giocatore", "ruolo_label", "partite_nel_foglio", "et_pos_tot", "et_pos_media", "et_neg_tot", "et_neg_media", "tr_pos_tot", "tr_pos_media", "tr_neg_tot", "tr_neg_media", "valutazione_grezza_media"]
        st.dataframe(format_table(filtered[et_cols].sort_values(["et_pos_tot", "tr_pos_tot"], ascending=False)), use_container_width=True, hide_index=True, height=420)
    with tab4:
        cols = ["giocatore", "ruolo_label", "mano", "jolly", "partite_nel_foglio", "gol_tot", "assist_tot", "tiri_tot", "pct_tiri_media", "pct_tiri_stagionale", "palla_rubata_tot", "palla_persa_tot", "valutazione_grezza_totale", "valutazione_grezza_media", "indice_valutazione_0_100"]
        table = filtered[cols].sort_values("indice_valutazione_0_100", ascending=False)
        st.dataframe(format_table(table), use_container_width=True, hide_index=True, height=640)


def comparison_page(data: dict[str, pd.DataFrame]) -> None:
    players = data["Giocatori_Stagione"]
    matches = data["Giocatori_Partita"]
    evaluations = data["Valutazione_Partita"]
    movement = players[players["ruolo"] == "movimento"].sort_values("giocatore")
    hero("Confronto Giocatori", "Due profili a confronto: KPI, radar e andamento partita-per-partita.")

    names = movement["giocatore"].tolist()
    col1, col2 = st.columns(2)
    with col1:
        player_a = st.selectbox("Giocatore A", names, index=0)
    with col2:
        player_b = st.selectbox("Giocatore B", names, index=min(1, len(names) - 1))

    row_a = movement[movement["giocatore"] == player_a].iloc[0]
    row_b = movement[movement["giocatore"] == player_b].iloc[0]
    metrics = [
        ("Gol", "gol_tot"),
        ("Assist", "assist_tot"),
        ("Tiri", "tiri_tot"),
        ("% tiro media", "pct_tiri_media"),
        ("% tiro stagionale", "pct_tiri_stagionale"),
        ("Rubate", "palla_rubata_tot"),
        ("Perse", "palla_persa_tot"),
        ("Val. media", "valutazione_grezza_media"),
        ("Indice", "indice_valutazione_0_100"),
    ]
    kpi_df = pd.DataFrame(
        [{"metrica": label, player_a: row_a[col], player_b: row_b[col], "delta": row_a[col] - row_b[col]} for label, col in metrics]
    )

    c1, c2 = st.columns([1, 1.15])
    with c1:
        st.dataframe(format_table(kpi_df), use_container_width=True, hide_index=True, height=470)
    with c2:
        radar_cols = ["gol_tot", "assist_tot", "palla_rubata_tot", "pct_tiri_stagionale", "valutazione_grezza_media"]
        radar_labels = ["Gol", "Assist", "Rubate", "% tiro", "Val. media"]
        radar_source = movement[["giocatore"] + radar_cols].copy()
        for col in radar_cols:
            max_value = radar_source[col].max()
            radar_source[col] = radar_source[col] / max_value * 100 if max_value else 0
        a_norm = radar_source[radar_source["giocatore"] == player_a].iloc[0]
        b_norm = radar_source[radar_source["giocatore"] == player_b].iloc[0]
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(r=[a_norm[c] for c in radar_cols], theta=radar_labels, fill="toself", name=player_a, line={"color": ARESE_CYAN}))
        fig.add_trace(go.Scatterpolar(r=[b_norm[c] for c in radar_cols], theta=radar_labels, fill="toself", name=player_b, line={"color": ARESE_ORANGE}))
        fig.update_layout(polar={"radialaxis": {"visible": True, "range": [0, 100]}}, title="Radar normalizzato sul gruppo movimento")
        st.plotly_chart(styled_plotly(fig, 470), use_container_width=True)

    detail = matches[matches["giocatore"].isin([player_a, player_b])].merge(
        evaluations[["match_id", "giocatore", "valutazione_finale"]], on=["match_id", "giocatore"], how="left"
    )
    fig = px.line(detail.sort_values("giornata"), x="giornata", y="valutazione_finale", color="giocatore", markers=True, title="Valutazione partita-per-partita", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE])
    fig.update_traces(line={"width": 4}, marker={"size": 9})
    st.plotly_chart(styled_plotly(fig), use_container_width=True)
    st.dataframe(format_table(detail[["giornata", "data", "avversario", "giocatore", "gol", "assist", "tiri", "pct_tiri", "palla_rubata", "palla_persa", "errore", "valutazione_finale"]].sort_values(["giornata", "giocatore"])), use_container_width=True, hide_index=True, height=560)


def player_profile_page(data: dict[str, pd.DataFrame]) -> None:
    players = data["Giocatori_Stagione"].copy()
    matches = data["Giocatori_Partita"].copy()
    evaluations = data["Valutazione_Partita"].copy()
    hero("Scheda Giocatore", "Profilo individuale, andamento partita-per-partita e paragone con la media del ruolo.")

    names = players.sort_values(["ruolo", "giocatore"])["giocatore"].tolist()
    saved_selection = st.session_state.get("selected_player_profile")
    default_player = saved_selection if saved_selection in names else "NOA" if "NOA" in names else names[0]
    selected = st.selectbox("Seleziona giocatore", names, index=names.index(default_player))
    st.session_state["selected_player_profile"] = selected
    row = players[players["giocatore"] == selected].iloc[0]
    same_role = players[players["ruolo_base"] == row["ruolo_base"]].copy()
    role_label = row["ruolo_label"] if pd.notna(row["ruolo_label"]) else row["ruolo_base"]

    st.markdown(
        f"<span class='pill'>{row['ruolo_base']}</span><span class='pill'>{role_label}</span><span class='pill'>Mano {row['mano']}</span><span class='pill'>Jolly {row['jolly']}</span>",
        unsafe_allow_html=True,
    )

    if row["ruolo"] == "portiere":
        kpis = [
            ("Partite stimate", format_num(row["partite_giocate_stimate"], 0), "giocate"),
            ("Parate", format_num(row["parata_tot"], 0), f"{format_num(row['parata_media'], 2)} / partita"),
            ("Gol subiti", format_num(row["gol_subito_tot"], 0), f"{format_num(row['gol_subito_media'], 2)} / partita"),
            ("% parate", format_pct(row["pct_parate_media"]), "media"),
            ("Val. media", format_num(row["valutazione_grezza_media"], 2), "grezza"),
            ("Indice", format_num(row["indice_valutazione_0_100"], 1), "0-100"),
        ]
        comparison_metrics = [
            ("Parate/partita", "parata_media"),
            ("Gol subiti/partita", "gol_subito_media"),
            ("% parate", "pct_parate_media"),
            ("Valutazione media", "valutazione_grezza_media"),
            ("Indice 0-100", "indice_valutazione_0_100"),
        ]
    else:
        kpis = [
            ("Partite", format_num(row["partite_nel_foglio"], 0), "nel foglio"),
            ("Gol", format_num(row["gol_tot"], 0), f"{format_num(row['gol_media'], 2)} / partita"),
            ("Assist", format_num(row["assist_tot"], 0), f"{format_num(row['assist_media'], 2)} / partita"),
            ("Tiri", format_num(row["tiri_tot"], 0), f"{format_num(row['tiri_media'], 2)} / partita"),
            ("ET+", format_num(row["et_pos_tot"], 0), f"{format_num(row['et_pos_media'], 2)} / partita"),
            ("ET-", format_num(row["et_neg_tot"], 0), f"{format_num(row['et_neg_media'], 2)} / partita"),
            ("TR+", format_num(row["tr_pos_tot"], 0), f"{format_num(row['tr_pos_media'], 2)} / partita"),
            ("TR-", format_num(row["tr_neg_tot"], 0), f"{format_num(row['tr_neg_media'], 2)} / partita"),
            ("% tiro media", format_pct(row["pct_tiri_media"]), "partita-per-partita"),
            ("% tiro stagionale", format_pct(row["pct_tiri_stagionale"]), "gol / tiri"),
            ("Val. media", format_num(row["valutazione_grezza_media"], 2), "grezza"),
            ("Indice", format_num(row["indice_valutazione_0_100"], 1), "0-100"),
        ]
        comparison_metrics = [
            ("Gol/partita", "gol_media"),
            ("Assist/partita", "assist_media"),
            ("Tiri/partita", "tiri_media"),
            ("% tiro media", "pct_tiri_media"),
            ("% tiro stagionale", "pct_tiri_stagionale"),
            ("ET+/partita", "et_pos_media"),
            ("ET-/partita", "et_neg_media"),
            ("TR+/partita", "tr_pos_media"),
            ("TR-/partita", "tr_neg_media"),
            ("Rubate/partita", "palla_rubata_media"),
            ("Perse/partita", "palla_persa_media"),
            ("Valutazione media", "valutazione_grezza_media"),
            ("Indice 0-100", "indice_valutazione_0_100"),
        ]

    kpi_cols = st.columns(4)
    for idx, (label, value, note) in enumerate(kpis):
        with kpi_cols[idx % 4]:
            metric_card(label, value, note)

    comparison_rows = []
    for label, col in comparison_metrics:
        player_value = row[col]
        role_avg = same_role[col].mean()
        comparison_rows.append({"metrica": label, selected: player_value, "media ruolo": role_avg, "delta": player_value - role_avg})
    comparison_df = pd.DataFrame(comparison_rows)

    left, right = st.columns([1, 1.15])
    with left:
        st.markdown(f"### Confronto con media ruolo: {role_label}")
        st.dataframe(format_table(comparison_df), use_container_width=True, hide_index=True, height=420)
    with right:
        norm = comparison_df.copy()
        radar_pool = players[players["ruolo"] == row["ruolo"]].copy()
        norm[selected] = 0.0
        norm["media ruolo"] = 0.0
        metric_lookup = dict(comparison_metrics)
        for idx, metric_row in comparison_df.iterrows():
            source_col = metric_lookup[metric_row["metrica"]]
            scale = radar_pool[source_col].abs().max()
            scale = scale if scale and pd.notna(scale) else 1
            norm.loc[idx, selected] = metric_row[selected] / scale * 100
            norm.loc[idx, "media ruolo"] = metric_row["media ruolo"] / scale * 100
        norm[selected] = norm[selected].clip(-100, 100)
        norm["media ruolo"] = norm["media ruolo"].clip(-100, 100)
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(r=norm[selected], theta=norm["metrica"], fill="toself", name=selected, line={"color": ARESE_CYAN}))
        fig.add_trace(go.Scatterpolar(r=norm["media ruolo"], theta=norm["metrica"], fill="toself", name="Media ruolo", line={"color": ARESE_ORANGE}))
        fig.update_layout(polar={"radialaxis": {"visible": True, "range": [-100, 100]}}, title="Radar vs media ruolo, normalizzato sulla squadra")
        st.plotly_chart(styled_plotly(fig, 420), use_container_width=True)

    delta_df = comparison_df.sort_values("delta")
    fig = px.bar(delta_df, x="delta", y="metrica", orientation="h", title="Delta rispetto alla media ruolo", color="delta", color_continuous_scale=[ARESE_RED, ARESE_GOLD, ARESE_CYAN])
    fig.update_layout(yaxis_title="", xaxis_title="Delta")
    st.plotly_chart(styled_plotly(fig, 470), use_container_width=True)

    detail = matches[matches["giocatore"] == selected].merge(
        evaluations[["match_id", "giocatore", "valutazione_finale"]], on=["match_id", "giocatore"], how="left"
    ).sort_values("giornata")
    if row["ruolo"] == "portiere":
        fig = px.line(detail, x="giornata", y=["parata", "gol_subito", "valutazione_finale"], markers=True, title="Andamento partita-per-partita", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE, ARESE_GOLD])
        detail_cols = ["giornata", "data", "avversario", "parata", "gol_subito", "pct_parate", "assist", "errore", "valutazione_finale"]
    else:
        fig = px.line(detail, x="giornata", y=["gol", "assist", "tiri", "valutazione_finale"], markers=True, title="Andamento partita-per-partita", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE, ARESE_AQUA, ARESE_GOLD])
        detail_cols = ["giornata", "data", "avversario", "gol", "assist", "tiri", "pct_tiri", "et_pos", "et_neg", "tr_pos", "tr_neg", "palla_rubata", "palla_persa", "errore", "valutazione_finale"]
    fig.update_traces(line={"width": 4}, marker={"size": 8})
    st.plotly_chart(styled_plotly(fig, 470), use_container_width=True)
    st.dataframe(format_table(detail[detail_cols]), use_container_width=True, hide_index=True, height=520)


def scatter_lab_page(data: dict[str, pd.DataFrame]) -> None:
    players = data["Giocatori_Stagione"]
    hero("Scatter Lab", "Laboratorio interattivo per esplorare relazioni tra KPI. Scegli assi, dimensione bolle, colore e filtri.")

    dataset = st.radio("Dataset", ["Movimento", "Portieri", "Tutti"], horizontal=True)
    if dataset == "Movimento":
        base = players[players["ruolo"] == "movimento"].copy()
    elif dataset == "Portieri":
        base = players[players["ruolo"] == "portiere"].copy()
    else:
        base = players.copy()

    numeric_cols = [col for col in base.columns if pd.api.types.is_numeric_dtype(base[col])]
    preferred = [col for col in ["gol_tot", "assist_tot", "tiri_tot", "pct_tiri_media", "pct_tiri_stagionale", "et_pos_tot", "et_neg_tot", "tr_pos_tot", "tr_neg_tot", "valutazione_grezza_totale", "valutazione_grezza_media", "indice_valutazione_0_100", "rischio_tot", "impatto_positivo", "azioni_positive_difesa", "parata_tot", "gol_subito_tot", "pct_parate_media"] if col in numeric_cols]
    numeric_cols = preferred + [col for col in numeric_cols if col not in preferred]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        x = st.selectbox("Asse X", numeric_cols, index=numeric_cols.index("tiri_tot") if "tiri_tot" in numeric_cols else 0)
    with c2:
        y = st.selectbox("Asse Y", numeric_cols, index=numeric_cols.index("pct_tiri_stagionale") if "pct_tiri_stagionale" in numeric_cols else min(1, len(numeric_cols) - 1))
    with c3:
        size = st.selectbox("Dimensione", numeric_cols, index=numeric_cols.index("partite_nel_foglio") if "partite_nel_foglio" in numeric_cols else 0)
    with c4:
        color = st.selectbox("Colore", ["ruolo_base", "mano", "jolly", "ruolo"], index=0)

    min_games = st.slider("Min partite nel foglio", 0, int(base["partite_nel_foglio"].max()), 0)
    frame = base[base["partite_nel_foglio"].fillna(0) >= min_games].copy()
    st.plotly_chart(scatter_players(frame, x, y, size, color, f"{y} vs {x}", ["partite_nel_foglio", "gol_tot", "assist_tot", "tiri_tot"]), use_container_width=True)

    st.markdown("### Preset consigliati")
    presets = {
        "Produzione offensiva": ("gol_tot", "assist_tot", "partite_nel_foglio", "ruolo_base"),
        "Volume vs efficienza tiro": ("tiri_tot", "pct_tiri_stagionale", "gol_tot", "ruolo_base"),
        "Consistenza vs efficienza": ("pct_tiri_media", "pct_tiri_stagionale", "tiri_tot", "mano"),
        "Impatto totale vs medio": ("valutazione_grezza_totale", "valutazione_grezza_media", "partite_nel_foglio", "ruolo_base"),
        "Rischio vs rendimento": ("rischio_tot", "impatto_positivo", "valutazione_grezza_totale", "ruolo_base"),
        "Difesa vs attacco": ("azioni_positive_difesa", "produzione_offensiva", "indice_valutazione_0_100", "jolly"),
        "ET conquistate vs subite": ("et_pos_tot", "et_neg_tot", "partite_nel_foglio", "ruolo_base"),
        "Rigori conquistati vs subiti": ("tr_pos_tot", "tr_neg_tot", "partite_nel_foglio", "ruolo_base"),
    }
    cols = st.columns(3)
    for idx, (label, values) in enumerate(presets.items()):
        px_, py_, ps_, pc_ = values
        with cols[idx % 3]:
            if all(v in frame.columns for v in [px_, py_, ps_, pc_]):
                st.plotly_chart(scatter_players(frame, px_, py_, ps_, pc_, label, ["partite_nel_foglio"]), use_container_width=True)


def goalkeepers_page(data: dict[str, pd.DataFrame]) -> None:
    goalkeepers = data["Portieri"].copy()
    hero("Portieri", "Carico di lavoro, efficienza, gol subiti e valutazione dei portieri.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Portieri", str(len(goalkeepers)), "in rosa")
    with c2:
        metric_card("Parate", format_num(goalkeepers["parata_tot"].sum(), 0), "totale")
    with c3:
        metric_card("Gol subiti", format_num(goalkeepers["gol_subito_tot"].sum(), 0), "totale")
    with c4:
        weighted = goalkeepers["parata_tot"].sum() / (goalkeepers["parata_tot"].sum() + goalkeepers["gol_subito_tot"].sum())
        metric_card("% parate stimata", format_pct(weighted), "parate / tiri in porta")

    a, b = st.columns(2)
    with a:
        fig = px.bar(goalkeepers.sort_values("parata_tot", ascending=False), x="giocatore", y="parata_tot", title="Parate totali", color="giocatore", color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE, ARESE_GOLD])
        st.plotly_chart(styled_plotly(fig, 470), use_container_width=True)
    with b:
        fig = px.scatter(goalkeepers, x="gol_subito_tot", y="pct_parate_media", size="parata_tot", color="giocatore", text="giocatore", title="Carico vs efficienza", hover_data=["partite_giocate_stimate", "parata_media", "gol_subito_media"], size_max=55, color_discrete_sequence=[ARESE_CYAN, ARESE_ORANGE, ARESE_GOLD])
        fig.update_traces(textposition="top center")
        st.plotly_chart(styled_plotly(fig, 470), use_container_width=True)

    fig = px.scatter(goalkeepers, x="gol_subito_media", y="parata_media", size="partite_giocate_stimate", color="indice_valutazione_0_100", text="giocatore", title="Media parate vs media gol subiti", color_continuous_scale=[ARESE_RED, ARESE_GOLD, ARESE_CYAN], size_max=50)
    fig.update_traces(textposition="top center")
    st.plotly_chart(styled_plotly(fig, 520), use_container_width=True)
    st.dataframe(format_table(goalkeepers[["giocatore", "partite_nel_foglio", "partite_giocate_stimate", "parata_tot", "parata_media", "gol_subito_tot", "gol_subito_media", "pct_parate_media", "valutazione_grezza_totale", "valutazione_grezza_media", "indice_valutazione_0_100"]]), use_container_width=True, hide_index=True, height=220)


def rankings_page(data: dict[str, pd.DataFrame]) -> None:
    hero("Classifiche", "Classifiche interne principali derivate dal workbook Excel.")
    players = data["Giocatori_Stagione"]
    movement = players[players["ruolo"] == "movimento"].copy()
    goalkeepers = players[players["ruolo"] == "portiere"].copy()
    tabs = st.tabs(["Marcatori", "Assist", "% tiro", "Valutazione", "Portieri", "Errori"])
    with tabs[0]:
        st.dataframe(format_table(movement.sort_values("gol_tot", ascending=False)[["giocatore", "gol_tot", "gol_media", "partite_nel_foglio", "ruolo_label"]]), use_container_width=True, hide_index=True)
    with tabs[1]:
        st.dataframe(format_table(movement.sort_values("assist_tot", ascending=False)[["giocatore", "assist_tot", "assist_media", "partite_nel_foglio", "ruolo_label"]]), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(format_table(movement[movement["tiri_tot"] > 0].sort_values("pct_tiri_stagionale", ascending=False)[["giocatore", "pct_tiri_media", "pct_tiri_stagionale", "gol_tot", "tiri_tot", "partite_nel_foglio"]]), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(format_table(movement.sort_values("indice_valutazione_0_100", ascending=False)[["giocatore", "valutazione_grezza_totale", "valutazione_grezza_media", "indice_valutazione_0_100", "partite_nel_foglio"]]), use_container_width=True, hide_index=True)
    with tabs[4]:
        st.dataframe(format_table(goalkeepers.sort_values("indice_valutazione_0_100", ascending=False)[["giocatore", "parata_tot", "gol_subito_tot", "pct_parate_media", "valutazione_grezza_totale", "indice_valutazione_0_100", "partite_giocate_stimate"]]), use_container_width=True, hide_index=True)
    with tabs[5]:
        st.caption("ERRORE è una metrica soggettiva: utile come segnale, non come giudizio assoluto.")
        st.dataframe(format_table(players.sort_values("errore_tot", ascending=False)[["giocatore", "ruolo", "errore_tot", "errore_media", "partite_nel_foglio"]]), use_container_width=True, hide_index=True)


def method_page(data: dict[str, pd.DataFrame]) -> None:
    hero("Metodo", "Note di calcolo, perimetro dati e avvertenze sulle metriche.")
    st.dataframe(data["Metodo"], use_container_width=True, hide_index=True)


def main() -> None:
    inject_css()
    page = sidebar_header()
    if not WORKBOOK_PATH.exists():
        st.error(f"Workbook non trovato: {WORKBOOK_PATH}")
        st.stop()

    data = prepare_data(load_data(WORKBOOK_PATH))
    st.sidebar.markdown("---")
    st.sidebar.caption(f"Dati: `{WORKBOOK_PATH.name}`")
    st.sidebar.caption("Le statistiche giocatore usano le partite con video/statistiche disponibili.")

    if page == "Overview":
        overview_page(data)
    elif page == "Squadra":
        team_page(data)
    elif page == "Giocatori":
        players_page(data)
    elif page == "Scheda Giocatore":
        player_profile_page(data)
    elif page == "Confronto Giocatori":
        comparison_page(data)
    elif page == "Scatter Lab":
        scatter_lab_page(data)
    elif page == "Portieri":
        goalkeepers_page(data)
    elif page == "Classifiche":
        rankings_page(data)
    elif page == "Metodo":
        method_page(data)


if __name__ == "__main__":
    main()
