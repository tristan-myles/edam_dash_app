"""
MoH use case prioritisation dashboard
Compares scores across groups for use cases on 5 criteria.
Run with: python dash_app_render_ready.py
"""

import base64
import getpass
import glob
import os
import re
import socket
import sys
from dataclasses import dataclass

import dash
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from dash import Input, Output, dash_table, dcc, html
from flask import Response, request
from werkzeug.security import check_password_hash, generate_password_hash

try:
    from pyngrok import conf as _ngrok_conf
    from pyngrok import ngrok as _ngrok
except ImportError:
    _ngrok_conf = None
    _ngrok = None

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCORE_COLS = [
    "Policy Relevance",
    "Decision Impact",
    "Data Availability",
    "Feasibility",
    "Capacity Building",
]
RADAR_ROTATION_DEGREES = 144

INPUT_COLS = [
    "Partner/Office",
    "Workshop Group",
    "Use Case",
    "Policy Question",
    "Policy Relevance",
    "Decision Impact",
    "Data Availability",
    "Feasibility",
    "Capacity Building",
    "Similar Efforts",
    "Total",
]

# Criterion colours - matches the Excel Comprehensive Summary header colours
CRITERION_COLOURS = {
    "Policy Relevance":  "#1D4ED8",
    "Decision Impact":   "#059669",
    "Data Availability": "#D97706",
    "Feasibility":       "#7C3AED",
    "Capacity Building": "#DC2626",
}

# Group colours - matches the Excel output (Ethiopian flag colours for 1–5,
# arbitrary distinct colours for 6–10)
GROUP_COLOURS = {
    "Group 1":  "#078930",
    "Group 2":  "#FCDD09",
    "Group 3":  "#DA121A",
    "Group 4":  "#0F47AF",
    "Group 5":  "#8E44AD",
    "Group 6":  "#E67E22",
    "Group 7":  "#C0392B",
    "Group 8":  "#16A085",
    "Group 9":  "#1A1A1A",
    "Group 10": "#2C3E50",
}
LIGHT_BG_GROUPS = {"Group 2"}
DEFAULT_TUNNEL_PORT = 8080
DEFAULT_LAN_PORT = 8080

def build_parser():
    parser = __import__("argparse").ArgumentParser(
        description="MoH use case prioritisation dashboard"
    )
    parser.add_argument("--data_dir", "-d", default=os.path.dirname(__file__),
                        help="Directory containing group xlsx files (default: script directory)")
    parser.add_argument("--username", default="workshop", help="Basic-auth username (default: workshop)")
    parser.add_argument("--password", default=None, help="Basic-auth password (prompted if omitted)")
    parser.add_argument("--lan", action="store_true", help="Serve on all network interfaces (local network access)")
    parser.add_argument("--tunnel", action="store_true", help="Expose via ngrok tunnel for remote access")
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable Dash debug mode locally",
    )
    parser.add_argument("--port", type=int, default=None,
                        help=f"Port to serve on (default: {DEFAULT_TUNNEL_PORT} with --tunnel/--lan, otherwise choose a free port)")
    parser.add_argument("--ngrok-token", default=os.environ.get("NGROK_AUTHTOKEN"),
                        help="Ngrok auth token (or set NGROK_AUTHTOKEN)")
    parser.add_argument("--ngrok-domain", default=os.environ.get("NGROK_DOMAIN"),
                        help="Reserved ngrok domain to use for the public URL (or set NGROK_DOMAIN)")
    parser.add_argument("--ngrok-region", default=os.environ.get("NGROK_REGION"),
                        choices=["us", "eu", "ap", "au", "sa", "jp", "in"],
                        help="Ngrok tunnel region to try, e.g. eu, in, or ap (or set NGROK_REGION)")
    return parser


@dataclass
class RuntimeArgs:
    data_dir: str = os.environ.get("DATA_DIR", os.path.dirname(__file__))
    username: str = os.environ.get("DASH_BASIC_AUTH_USERNAME", "workshop")
    password: str | None = os.environ.get("DASH_BASIC_AUTH_PASSWORD")
    lan: bool = False
    tunnel: bool = False
    debug: bool = False
    port: int | None = int(os.environ["PORT"]) if os.environ.get("PORT") else None
    ngrok_token: str | None = os.environ.get("NGROK_AUTHTOKEN")
    ngrok_domain: str | None = os.environ.get("NGROK_DOMAIN")
    ngrok_region: str | None = os.environ.get("NGROK_REGION")


def parse_runtime_args():
    parser = build_parser()
    if __name__ == "__main__":
        args, _unknown = parser.parse_known_args()
        return args
    return RuntimeArgs()


_args = parse_runtime_args()
DATA_DIR = _args.data_dir

# ---------------------------------------------------------------------------
# Logo
# ---------------------------------------------------------------------------

_LOGO_PATH = os.path.join(os.path.dirname(__file__), "combined_v2.png")
LOGO_SRC = None
if os.path.exists(_LOGO_PATH):
    with open(_LOGO_PATH, "rb") as _f:
        LOGO_SRC = "data:image/png;base64," + base64.b64encode(_f.read()).decode()

# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def filename_group(path):
    """Return the group label from everything before the first hyphen."""
    stem = os.path.splitext(os.path.basename(path))[0]
    return stem.split("-", 1)[0].strip()


def group_number(group):
    """Return a leading group number from labels such as '3' or '3 Infectious Diseases'."""
    match = re.match(r"^\s*(\d+)\b", str(group))
    return match.group(1) if match else None


def group_colour(group):
    number = group_number(group)
    key = f"Group {number}" if number else str(group)
    return GROUP_COLOURS.get(key, "#CCCCCC")


def text_colour_for_background(hex_colour):
    hex_colour = hex_colour.lstrip("#")
    if len(hex_colour) != 6:
        return "#111827"
    red, green, blue = (int(hex_colour[i:i + 2], 16) for i in (0, 2, 4))
    luminance = (0.299 * red + 0.587 * green + 0.114 * blue) / 255
    return "#111827" if luminance > 0.62 else "#ffffff"


def build_group_colour_map(group_values):
    palette = list(GROUP_COLOURS.values())
    colour_map = {}
    for index, group in enumerate(group_values):
        colour = group_colour(group)
        if colour == "#CCCCCC" and group_number(group) is None:
            colour = palette[index % len(palette)]
        colour_map[group] = colour
    return colour_map


def datatable_filter_text(value):
    return str(value).replace('"', '\\"')


def separate_policy_questions(value):
    text = str(value).replace("\\n", "\n").strip()
    text = re.sub(r"\s+(?=\d+\.\s)", "\n\n", text)
    return re.sub(r"\n{3,}", "\n\n", text)


def policy_questions_for_hover(values):
    questions = []
    seen = set()
    for value in values.dropna():
        text = separate_policy_questions(value)
        if text and text not in seen:
            seen.add(text)
            questions.append(text)
    return "<br><br>".join(questions).replace("\n", "<br>") if questions else "No policy question provided"


def sort_key(value):
    text = str(value)
    number = group_number(text)
    return (0, int(number), text.casefold()) if number else (1, text.casefold())


def load_data():
    files = sorted(
        path for path in glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
        if "-" in os.path.basename(path) and not os.path.basename(path).startswith("~$")
    )
    if not files:
        raise FileNotFoundError(
            f"No input Excel files matching '*-*.xlsx' were found in {DATA_DIR!r}. "
            "For Render, include the workbook files in the repo or set DATA_DIR to the mounted data directory."
        )
    frames = []
    for path in files:
        group = filename_group(path)
        df = pd.read_excel(path, engine="openpyxl", usecols=range(len(INPUT_COLS)))
        df.columns = INPUT_COLS
        df = df[df["Use Case"].notna() & (df["Use Case"] != "Use case")].copy()
        for col in SCORE_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        df["Total"] = df[SCORE_COLS].sum(axis=1)
        df["Workshop Group"] = int(group) if group.isdigit() else group
        df["Group"] = group
        df["Use Case Short"] = df["Use Case"].str.replace(r"^\d+\s*-\s*", "", regex=True).str.strip()
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


df_all = load_data()
use_cases   = sorted(df_all["Use Case Short"].unique())
groups      = sorted(df_all["Group"].unique(), key=sort_key)
partners    = sorted(df_all["Partner/Office"].dropna().unique())
wk_groups   = sorted(df_all["Workshop Group"].dropna().unique(), key=sort_key)

GROUP_COLOUR_MAP = build_group_colour_map(groups)

# ---------------------------------------------------------------------------
# Filter helper
# ---------------------------------------------------------------------------

def apply_filters(partners_sel, groups_sel, use_cases_sel, scoring_groups_sel=None,
                  top_n=None, top_n_mode="group"):
    df = df_all
    if partners_sel:
        df = df[df["Partner/Office"].isin(partners_sel)]
    if groups_sel:
        df = df[df["Workshop Group"].isin(groups_sel)]
    if use_cases_sel:
        df = df[df["Use Case Short"].isin(use_cases_sel)]
    if scoring_groups_sel is not None:
        df = df[df["Group"].isin(scoring_groups_sel)]
    if top_n and top_n > 0 and top_n_mode == "overall":
        top_ucs = (df.groupby("Use Case Short")["Total"].mean()
                   .nlargest(int(top_n)).index)
        df = df[df["Use Case Short"].isin(top_ucs)]
    if top_n and top_n > 0 and top_n_mode == "group":
        top_n = int(top_n)
        top_by_group = (
            df.groupby(["Group", "Use Case Short"], as_index=False)["Total"]
            .mean()
            .sort_values(["Group", "Total"], ascending=[True, False])
            .groupby("Group", group_keys=False)
            .head(top_n)
        )
        keep = set(zip(top_by_group["Group"], top_by_group["Use Case Short"]))
        row_keys = list(zip(df["Group"], df["Use Case Short"]))
        df = df[[key in keep for key in row_keys]]
    return df


def get_lan_ip():
    """Return the likely LAN IP, avoiding hostname aliases such as 127.0.1.1."""
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
        try:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
        except OSError:
            return socket.gethostbyname(socket.gethostname())


def get_lan_urls(port):
    """Return likely LAN URLs, ignoring loopback and Docker-style addresses."""
    urls = []
    for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
        ip = info[4][0]
        if ip.startswith(("127.", "172.17.")):
            continue
        url = f"http://{ip}:{port}"
        if url not in urls:
            urls.append(url)
    fallback = f"http://{get_lan_ip()}:{port}"
    if fallback not in urls:
        urls.insert(0, fallback)
    return urls


def choose_port(requested_port=None):
    """Return a requested port if available, otherwise an OS-selected free port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("", requested_port or 0))
        return s.getsockname()[1]


def start_public_tunnel(port):
    """Create and return a public ngrok URL for the local Dash server."""
    if _ngrok is None:
        print("ERROR: pyngrok is not installed, so no public tunnel can be created.")
        print("Install it with: pip install pyngrok")
        sys.exit(1)

    if _args.ngrok_token:
        _ngrok.set_auth_token(_args.ngrok_token)

    pyngrok_config = None
    if _args.ngrok_region:
        pyngrok_config = _ngrok_conf.PyngrokConfig(region=_args.ngrok_region)

    connect_kwargs = {}
    if _args.ngrok_domain:
        connect_kwargs["domain"] = _args.ngrok_domain

    try:
        tunnel = _ngrok.connect(port, pyngrok_config=pyngrok_config, **connect_kwargs)
    except Exception as exc:
        print("ERROR: Could not start the ngrok public tunnel.")
        print(f"Reason: {exc}")
        print("Check that your ngrok auth token is configured and that the port is not blocked.")
        print("You can set the token with: export NGROK_AUTHTOKEN=your_token_here")
        sys.exit(1)

    return getattr(tunnel, "public_url", str(tunnel))

# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------

CARD_STYLE = {
    "backgroundColor": "#ffffff",
    "borderRadius": "8px",
    "boxShadow": "0 1px 4px rgba(0,0,0,0.12)",
    "padding": "16px",
    "marginBottom": "20px",
}

SECTION_LABEL = {
    "fontWeight": "600",
    "fontSize": "22px",
    "textAlign": "center",
    "marginBottom": "12px",
    "color": "#374151",
}
RESPONSIVE_GRAPH_CONFIG = {"responsive": True}
TOTAL_SCORE_HEIGHT = "clamp(736px, 86vh, 943px)"
AVERAGE_CHART_HEIGHT = "clamp(500px, 60vh, 700px)"
FILL_GRAPH_STYLE = {"flex": "1 1 0", "height": "100%", "minHeight": "0", "width": "100%"}
TOP_ROW_STYLE = {
    "display": "grid",
    "gridTemplateColumns": "repeat(auto-fit, minmax(min(100%, 560px), 1fr))",
    "gap": "20px",
    "marginBottom": "20px",
    "alignItems": "stretch",
}

app = dash.Dash(__name__, title="MoH prioritisation dashboard")
server = app.server

# ---------------------------------------------------------------------------
# Authentication  (password hash set in __main__ before app.run)
# ---------------------------------------------------------------------------

_pwd_hash = generate_password_hash(_args.password) if _args.password else None


@app.server.before_request
def _require_auth():
    if _pwd_hash is None:
        return
    auth = request.authorization
    if not auth or auth.username != _args.username or not check_password_hash(_pwd_hash, auth.password):
        return Response(
            "Access denied - please provide your workshop credentials.",
            401,
            {"WWW-Authenticate": 'Basic realm="FMoH Workshop"'},
        )

# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------

BODY_TEXT_SIZE = "16px"
TABLE_TEXT_SIZE = 16

_filter_label = {"fontWeight": "600", "fontSize": BODY_TEXT_SIZE, "color": "#6b7280",
                 "marginBottom": "4px"}
_filter_control = {"fontSize": BODY_TEXT_SIZE}
_filter_input = {
    **_filter_control,
    "width": "80px",
    "padding": "6px",
    "borderRadius": "4px",
    "border": "1px solid #d1d5db",
}

app.layout = html.Div(
    style={"fontFamily": "Inter, sans-serif", "backgroundColor": "#f3f4f6",
           "minHeight": "100vh", "padding": "24px"},
    children=[
        html.H1("MoH use case prioritisation",
                style={"color": "#111827", "marginBottom": "4px", "textAlign": "center"}),
        html.P(f"{len(use_cases)} use cases · {len(groups)} themes · scores 1–5 per criterion",
               style={"color": "#6b7280", "fontSize": BODY_TEXT_SIZE, "marginBottom": "16px"}),

        dcc.Store(id="active-groups", data=groups),

        # ── Global filters ─────────────────────────────────────────────────
        html.Div(
            style={**CARD_STYLE, "display": "flex", "flexDirection": "column", "gap": "10px"},
            children=[
                # Compact row: Partner/Office + Theme
                html.Div(style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "12px"},
                         children=[
                    html.Div([
                        html.P("Partner / office", style=_filter_label),
                        dcc.Dropdown(id="partner-filter",
                                     options=[{"label": p, "value": p} for p in partners],
                                     multi=True, placeholder="All…",
                                     style=_filter_control),
                    ]),
                    html.Div([
                        html.P("Theme", style=_filter_label),
                        dcc.Dropdown(id="group-filter",
                                     options=[{"label": g, "value": g} for g in wk_groups],
                                     multi=True, placeholder="All…",
                                     style=_filter_control),
                    ]),
                ]),
                # Full-width row: Use Case + Top N filters
                html.Div(style={"display": "grid", "gridTemplateColumns": "1fr auto auto", "gap": "12px",
                                "alignItems": "flex-end"}, children=[
                    html.Div([
                        html.P("Use case", style=_filter_label),
                        dcc.Dropdown(id="usecase-filter",
                                     options=[{"label": u, "value": u} for u in use_cases],
                                     multi=True, placeholder="All use cases…",
                                     style=_filter_control),
                    ]),
                    html.Div([
                        html.P("Show top", style=_filter_label),
                        dcc.Dropdown(
                            id="leaderboard-top-n-mode",
                            options=[
                                {"label": "Per theme", "value": "group"},
                                {"label": "Overall", "value": "overall"},
                            ],
                            value="group",
                            clearable=False,
                            style={"width": "150px", "fontSize": BODY_TEXT_SIZE},
                        ),
                    ]),
                    html.Div([
                        html.P("Top N", style=_filter_label),
                        dcc.Input(
                            id="leaderboard-top-n",
                            type="number", value=None, placeholder="All",
                            min=1, step=1,
                            style=_filter_input,
                        ),
                    ]),
                ]),
            ],
        ),

        # ── Total score ─────────────────────────────────────────────────────
        html.Div(
            style={
                **CARD_STYLE,
                "minWidth": 0,
                "minHeight": TOTAL_SCORE_HEIGHT,
                "display": "flex",
                "flexDirection": "column",
            },
            children=[
                html.P("Total score across all criteria", style=SECTION_LABEL),
                html.Div(
                    style={"display": "flex", "gap": "12px", "marginBottom": "10px",
                           "alignItems": "flex-end", "flexWrap": "wrap"},
                    children=[
                        html.Div([
                            html.P("Sort by", style=_filter_label),
                            dcc.Dropdown(
                                id="leaderboard-sort-by",
                                options=[
                                    {"label": "Total score", "value": "Total"},
                                    {"label": "Criterion score range", "value": "Range"},
                                ],
                                value="Total", clearable=False,
                                style={"width": "210px", "fontSize": BODY_TEXT_SIZE},
                            ),
                        ]),
                        html.Div([
                            html.P("Order", style=_filter_label),
                            dcc.Dropdown(
                                id="leaderboard-order",
                                options=[
                                    {"label": "Descending", "value": "desc"},
                                    {"label": "Ascending",  "value": "asc"},
                                ],
                                value="desc", clearable=False,
                                style={"width": "210px", "fontSize": BODY_TEXT_SIZE},
                            ),
                        ]),
                    ],
                ),
                dcc.Graph(
                    id="leaderboard-chart",
                    responsive=True,
                    config=RESPONSIVE_GRAPH_CONFIG,
                    style=FILL_GRAPH_STYLE,
                ),
            ],
        ),

        # ── Average score charts ────────────────────────────────────────────
        html.Div(
            style=TOP_ROW_STYLE,
            children=[
                html.Div(
                    style={
                        **CARD_STYLE,
                        "marginBottom": 0,
                        "minWidth": 0,
                        "minHeight": AVERAGE_CHART_HEIGHT,
                        "display": "flex",
                        "flexDirection": "column",
                    },
                    children=[
                        html.P("Average score per criterion with min/max use case range error bars", style=SECTION_LABEL),
                        dcc.RadioItems(
                            id="criteria-split-mode",
                            options=[
                                {"label": "Overall", "value": "overall"},
                                {"label": "Split by theme", "value": "group"},
                                {"label": "Split by use case", "value": "use_case"},
                            ],
                            value="overall",
                            inline=True,
                            style={"fontSize": BODY_TEXT_SIZE, "color": "#374151", "marginBottom": "8px"},
                            labelStyle={"marginRight": "14px"},
                        ),
                        dcc.Graph(
                            id="criteria-bar",
                            responsive=True,
                            config=RESPONSIVE_GRAPH_CONFIG,
                            style=FILL_GRAPH_STYLE,
                        ),
                    ],
                ),
                html.Div(
                    style={
                        **CARD_STYLE,
                        "marginBottom": 0,
                        "minWidth": 0,
                        "minHeight": AVERAGE_CHART_HEIGHT,
                        "display": "flex",
                        "flexDirection": "column",
                    },
                    children=[
                        html.P("Average score per criterion", style=SECTION_LABEL),
                        dcc.RadioItems(
                            id="radar-split-mode",
                            options=[
                                {"label": "Overall", "value": "overall"},
                                {"label": "Split by theme", "value": "group"},
                                {"label": "Split by use case", "value": "use_case"},
                            ],
                            value="group",
                            inline=True,
                            style={"fontSize": BODY_TEXT_SIZE, "color": "#374151", "marginBottom": "8px"},
                            labelStyle={"marginRight": "14px"},
                        ),
                        dcc.Graph(
                            id="all-scores-radar-chart",
                            responsive=True,
                            config=RESPONSIVE_GRAPH_CONFIG,
                            style=FILL_GRAPH_STYLE,
                        ),
                    ],
                ),
            ],
        ),

        # ── Agreement heatmap ──────────────────────────────────────────────
        html.Div(style=CARD_STYLE, children=[
            html.P(
                "Score heatmap per use case by criterion",
                style=SECTION_LABEL,
            ),
            html.Div(
                style={"display": "flex", "gap": "12px", "marginBottom": "10px",
                       "alignItems": "flex-end"},
                children=[
                    html.Div([
                        html.P("Sort by", style=_filter_label),
                        dcc.Dropdown(
                            id="heatmap-sort-by",
                            options=(
                                [
                                    {"label": "Total score", "value": "Total"},
                                    {"label": "Criterion score SD", "value": "SD"},
                                    {"label": "Criterion score range", "value": "Range"},
                                ] +
                                [{"label": c, "value": c} for c in SCORE_COLS]
                            ),
                            value="Total", clearable=False,
                            style={"width": "210px", "fontSize": BODY_TEXT_SIZE},
                        ),
                    ]),
                    html.Div([
                        html.P("Order", style=_filter_label),
                        dcc.Dropdown(
                            id="heatmap-order",
                            options=[
                                {"label": "Descending", "value": "desc"},
                                {"label": "Ascending",  "value": "asc"},
                            ],
                            value="desc", clearable=False,
                            style={"width": "210px", "fontSize": BODY_TEXT_SIZE},
                        ),
                    ]),
                ],
            ),
            dcc.Graph(id="agreement-heatmap", style={"height": "820px"}),
        ]),

        # ── Stats table ────────────────────────────────────────────────────
        html.Div(style=CARD_STYLE, children=[
            html.P("Summary table - scores per use case across themes", style=SECTION_LABEL),
            html.Div(
                style={"display": "flex", "gap": "12px", "marginBottom": "10px",
                       "alignItems": "flex-end"},
                children=[
                    html.Div([
                        html.P("Sort by", style=_filter_label),
                        dcc.Dropdown(
                            id="table-sort-by",
                            options=(
                                [
                                    {"label": "Overall rank", "value": "Overall rank"},
                                    {"label": "Theme rank", "value": "Theme rank"},
                                    {"label": "Total", "value": "Total"},
                                ] +
                                [{"label": c, "value": c} for c in SCORE_COLS] +
                                [
                                    {"label": "Theme", "value": "Theme"},
                                ]
                            ),
                            value="Total", clearable=False,
                            style={"width": "120px", "fontSize": BODY_TEXT_SIZE},
                        ),
                    ]),
                    html.Div([
                        html.P("Order", style=_filter_label),
                        dcc.Dropdown(
                            id="table-order",
                            options=[
                                {"label": "Descending", "value": "desc"},
                                {"label": "Ascending",  "value": "asc"},
                            ],
                            value="desc", clearable=False,
                            style={"width": "120px", "fontSize": BODY_TEXT_SIZE},
                        ),
                    ]),
                ],
            ),
            html.Div(id="stats-table"),
        ]),

        # ── Footer ─────────────────────────────────────────────────────────
        html.Div(
            style={"textAlign": "center", "padding": "32px 0 16px"},
            children=[html.Img(src=LOGO_SRC, style={"maxHeight": "80px", "objectFit": "contain"})]
            if LOGO_SRC else [],
        ),

    ],
)

# ---------------------------------------------------------------------------
# Callbacks
# ---------------------------------------------------------------------------

@app.callback(
    Output("partner-filter", "options"),
    Output("group-filter", "options"),
    Output("usecase-filter", "options"),
    Input("partner-filter", "value"),
    Input("group-filter", "value"),
    Input("usecase-filter", "value"),
)
def update_filter_options(partners_sel, groups_sel, use_cases_sel):
    # Partner options: data filtered by the other two
    df_p = df_all
    if groups_sel:
        df_p = df_p[df_p["Workshop Group"].isin(groups_sel)]
    if use_cases_sel:
        df_p = df_p[df_p["Use Case Short"].isin(use_cases_sel)]

    # Workshop group options: filtered by partner + use case
    df_g = df_all
    if partners_sel:
        df_g = df_g[df_g["Partner/Office"].isin(partners_sel)]
    if use_cases_sel:
        df_g = df_g[df_g["Use Case Short"].isin(use_cases_sel)]

    # Use case options: filtered by partner + workshop group
    df_u = df_all
    if partners_sel:
        df_u = df_u[df_u["Partner/Office"].isin(partners_sel)]
    if groups_sel:
        df_u = df_u[df_u["Workshop Group"].isin(groups_sel)]

    return (
        [{"label": p, "value": p} for p in sorted(df_p["Partner/Office"].dropna().unique())],
        [{"label": g, "value": g} for g in sorted(df_g["Workshop Group"].dropna().unique(), key=sort_key)],
        [{"label": u, "value": u} for u in sorted(df_u["Use Case Short"].unique())],
    )


@app.callback(
    Output("leaderboard-chart", "figure"),
    Input("partner-filter", "value"),
    Input("group-filter", "value"),
    Input("usecase-filter", "value"),
    Input("active-groups", "data"),
    Input("leaderboard-sort-by", "value"),
    Input("leaderboard-order", "value"),
    Input("leaderboard-top-n-mode", "value"),
    Input("leaderboard-top-n", "value"),
)
def leaderboard(partners_sel, groups_sel, use_cases_sel, active_groups,
                sort_by, order, top_n_mode, top_n):
    df = apply_filters(
        partners_sel, groups_sel, use_cases_sel, active_groups, top_n, top_n_mode
    )
    grp = df.groupby(["Use Case Short", "Group"])
    criterion_means = grp[SCORE_COLS].mean()
    stats = pd.DataFrame({
        "Total": grp["Total"].mean(),
        "SD": criterion_means.std(axis=1),
        "Range": criterion_means.max(axis=1) - criterion_means.min(axis=1),
    }).reset_index()
    policy_questions = (
        df.groupby(["Use Case Short", "Group"])["Policy Question"]
        .apply(policy_questions_for_hover)
        .rename("Policy questions")
        .reset_index()
    )
    stats = stats.merge(policy_questions, on=["Use Case Short", "Group"], how="left")
    stats["SD"] = stats["SD"].fillna(0)

    # For horizontal bar charts, category order must be unique per use case.
    # Plotly collapses duplicate y labels across themes, so sorting the raw
    # theme/use-case rows can make ascending and descending appear identical.
    use_case_order = (
        stats.groupby("Use Case Short")[sort_by]
        .mean()
        .sort_values(ascending=(order == "desc"))
        .index
        .tolist()
    )
    stats = stats.sort_values(
        ["Use Case Short", "Group"],
        key=lambda s: s.map({value: index for index, value in enumerate(use_case_order)})
        if s.name == "Use Case Short" else s.map(sort_key),
    )

    fig = go.Figure()
    for group in sorted(stats["Group"].unique(), key=sort_key):
        group_stats = stats[stats["Group"] == group]
        fig.add_trace(go.Bar(
            y=group_stats["Use Case Short"],
            x=group_stats["Total"],
            name=group,
            orientation="h",
            width=0.86,
            marker_color=GROUP_COLOUR_MAP.get(group, "#CCCCCC"),
            hovertemplate=(
                "%{y}<br>Theme: %{fullData.name}<br>"
                "Total score: %{x:.2f}<br><br>"
                "Policy questions:<br>%{customdata}<extra></extra>"
            ),
            customdata=group_stats["Policy questions"],
        ))
    fig.update_layout(
        xaxis_title="Total score",
        yaxis_title=None,
        showlegend=True,
        legend_title="Theme",
        bargap=0.46,
        bargroupgap=0.22,
        hoverlabel=dict(font_size=16),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            font=dict(size=16),
            title_font=dict(size=16),
        ),
        margin=dict(l=260, r=20, t=50, b=40),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(size=16),
    )
    fig.update_xaxes(range=[0, 25], dtick=5, gridcolor="#f3f4f6")
    fig.update_yaxes(
        categoryorder="array",
        categoryarray=use_case_order,
        automargin=True,
        ticklabelposition="outside left",
        tickfont=dict(size=16),
    )
    return fig


@app.callback(
    Output("criteria-bar", "figure"),
    Input("criteria-split-mode", "value"),
    Input("partner-filter", "value"),
    Input("group-filter", "value"),
    Input("usecase-filter", "value"),
    Input("active-groups", "data"),
    Input("leaderboard-top-n-mode", "value"),
    Input("leaderboard-top-n", "value"),
)
def criteria_bar(split_mode, partners_sel, groups_sel, use_cases_sel, active_groups,
                 top_n_mode, top_n):
    df = apply_filters(
        partners_sel, groups_sel, use_cases_sel, active_groups, top_n, top_n_mode
    )
    if df.empty:
        return go.Figure()

    melted = pd.melt(df, id_vars=["Group", "Use Case Short"], value_vars=SCORE_COLS,
                     var_name="Criterion", value_name="Score")

    fig = go.Figure()
    if split_mode in {"group", "use_case"}:
        series_col = "Group" if split_mode == "group" else "Use Case Short"
        legend_title = "Theme" if split_mode == "group" else "Use case"
        stats = (
            melted.groupby([series_col, "Criterion"])["Score"]
            .agg(mean="mean", min="min", max="max")
            .reset_index()
        )
        series_values = sorted(
            stats[series_col].unique(),
            key=sort_key if split_mode == "group" else lambda value: str(value).casefold(),
        )
        for series_value in series_values:
            group_stats = (
                stats[stats[series_col] == series_value]
                .set_index("Criterion")
                .reindex(SCORE_COLS)
                .reset_index()
            )
            bar_kwargs = {}
            if split_mode == "group":
                bar_kwargs["marker_color"] = GROUP_COLOUR_MAP.get(series_value, "#CCCCCC")
            fig.add_trace(go.Bar(
                x=group_stats["Criterion"],
                y=group_stats["mean"],
                name=series_value,
                width=0.2,
                error_y=dict(
                    type="data",
                    symmetric=False,
                    array=(group_stats["max"] - group_stats["mean"]).fillna(0),
                    arrayminus=(group_stats["mean"] - group_stats["min"]).fillna(0),
                    visible=True,
                    color="#111827",
                    thickness=1.4,
                    width=4,
                ),
                hovertemplate=(
                    "%{x}<br>%{fullData.name}<br>"
                    "Mean: %{y:.2f}<br>"
                    "Min: %{customdata[0]:.2f}<br>"
                    "Max: %{customdata[1]:.2f}<extra></extra>"
                ),
                customdata=np.column_stack([group_stats["min"], group_stats["max"]]),
                **bar_kwargs,
            ))
        showlegend = True
        barmode = "group"
    else:
        legend_title = "Theme"
        stats = (
            melted.groupby("Criterion")["Score"]
            .agg(mean="mean", min="min", max="max")
            .reindex(SCORE_COLS)
            .reset_index()
        )
        fig.add_trace(go.Bar(
            x=stats["Criterion"],
            y=stats["mean"],
            marker_color=CRITERION_COLOURS["Data Availability"],
            width=0.72,
            error_y=dict(
                type="data",
                symmetric=False,
                array=(stats["max"] - stats["mean"]).fillna(0),
                arrayminus=(stats["mean"] - stats["min"]).fillna(0),
                visible=True,
                color="#111827",
                thickness=1.4,
                width=5,
            ),
            hovertemplate=(
                "%{x}<br>Mean: %{y:.2f}<br>"
                "Min: %{customdata[0]:.2f}<br>"
                "Max: %{customdata[1]:.2f}<extra></extra>"
            ),
            customdata=np.column_stack([stats["min"], stats["max"]]),
        ))
        showlegend = False
        barmode = "relative"

    fig.update_layout(
        autosize=True,
        barmode=barmode,
        bargap=0.12,
        bargroupgap=0.04,
        yaxis=dict(range=[0, 5], title="Mean score (1–5)"),
        showlegend=showlegend,
        legend_title=legend_title,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
        margin=dict(l=0, r=0, t=50, b=40),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(size=16),
    )
    fig.update_yaxes(gridcolor="#f3f4f6")
    return fig


@app.callback(
    Output("all-scores-radar-chart", "figure"),
    Input("radar-split-mode", "value"),
    Input("partner-filter", "value"),
    Input("group-filter", "value"),
    Input("usecase-filter", "value"),
    Input("active-groups", "data"),
    Input("leaderboard-top-n-mode", "value"),
    Input("leaderboard-top-n", "value"),
)
def all_scores_deep_dive(split_mode, partners_sel, groups_sel, use_cases_sel, active_groups,
                         top_n_mode, top_n):
    df = apply_filters(
        partners_sel, groups_sel, use_cases_sel, active_groups, top_n, top_n_mode
    )
    if df.empty:
        return go.Figure()

    radar_fig = go.Figure()

    if split_mode == "overall":
        values = df[SCORE_COLS].mean().tolist()
        values += [values[0]]
        radar_fig.add_trace(go.Scatterpolar(
            r=values, theta=SCORE_COLS + [SCORE_COLS[0]],
            name="Overall", fill="toself", opacity=0.5,
            line_color=CRITERION_COLOURS["Data Availability"],
        ))
        legend_title = None
    elif split_mode == "use_case":
        use_case_values = sorted(df["Use Case Short"].unique(), key=lambda value: str(value).casefold())
        for use_case in use_case_values:
            uc = df[df["Use Case Short"] == use_case]
            values = uc[SCORE_COLS].mean().tolist()
            values += [values[0]]
            radar_fig.add_trace(go.Scatterpolar(
                r=values, theta=SCORE_COLS + [SCORE_COLS[0]],
                name=use_case, fill="toself", opacity=0.35,
            ))
        legend_title = "Use case"
    else:
        active_groups = sorted(df["Group"].unique(), key=sort_key)
        for group in active_groups:
            g = df[df["Group"] == group]
            values = g[SCORE_COLS].mean().tolist()
            values += [values[0]]
            radar_fig.add_trace(go.Scatterpolar(
                r=values, theta=SCORE_COLS + [SCORE_COLS[0]],
                name=group, fill="toself", opacity=0.5,
                line_color=GROUP_COLOUR_MAP.get(group, "#CCCCCC"),
            ))
        legend_title = "Theme"

    radar_fig.update_layout(
        autosize=True,
        polar=dict(
            bgcolor="#ffffff",
            angularaxis=dict(
                categoryorder="array",
                categoryarray=SCORE_COLS,
                direction="clockwise",
                rotation=RADAR_ROTATION_DEGREES,
                gridcolor="#000000",
                linecolor="#000000",
                tickfont=dict(color="#000000", size=16),
            ),
            radialaxis=dict(
                visible=True,
                range=[0, 5],
                gridcolor="#000000",
                linecolor="#000000",
                tickfont=dict(color="#000000", size=16),
            ),
        ),
        margin=dict(l=20, r=20, t=45, b=10),
        paper_bgcolor="#ffffff",
        legend_title=legend_title,
        font=dict(size=16),
    )

    return radar_fig


@app.callback(
    Output("agreement-heatmap", "figure"),
    Input("partner-filter", "value"),
    Input("group-filter", "value"),
    Input("usecase-filter", "value"),
    Input("active-groups", "data"),
    Input("heatmap-sort-by", "value"),
    Input("heatmap-order", "value"),
    Input("leaderboard-top-n-mode", "value"),
    Input("leaderboard-top-n", "value"),
)
def agreement_heatmap(partners_sel, groups_sel, use_cases_sel, active_groups,
                      sort_by, order, top_n_mode, top_n):
    df = apply_filters(
        partners_sel, groups_sel, use_cases_sel, active_groups, top_n, top_n_mode
    )
    melted = pd.melt(df, id_vars=["Use Case Short"], value_vars=SCORE_COLS + ["Total"],
                     var_name="Criterion", value_name="Score")
    means = melted.groupby(["Use Case Short", "Criterion"])["Score"].mean().unstack("Criterion").round(2)
    means = means[SCORE_COLS + ["Total"]]
    ascending = (order == "asc")
    if sort_by == "SD":
        sort_key = means[SCORE_COLS].std(axis=1)
    elif sort_by == "Range":
        sort_key = means[SCORE_COLS].max(axis=1) - means[SCORE_COLS].min(axis=1)
    else:
        sort_key = means[sort_by]
    means = means.loc[sort_key.sort_values(ascending=ascending).index]

    uc_list = means.index.tolist()
    n_rows = len(uc_list)

    # Numeric x-axis: Total first, with Policy Relevance nudged right for label spacing.
    x_total = -0.35
    x_crit = [0.25] + list(range(1, len(SCORE_COLS)))

    # Custom colorscale: z=0 → white (Total sentinel), z=1-5 → RdYlGn
    # With zmin=0, zmax=5: position = z/5
    colorscale = [
        [0.0,  "#ffffff"],  # z=0  → white (Total column)
        [0.2,  "#d73027"],  # z=1  → red
        [0.4,  "#fc8d59"],  # z=2  → orange-red
        [0.6,  "#ffffbf"],  # z=3  → yellow
        [0.8,  "#91cf60"],  # z=4  → light green
        [1.0,  "#1a9850"],  # z=5  → green
    ]

    # Text: value for Total, empty for criteria cells
    text_matrix = np.column_stack([
        means["Total"].round(1).astype(str).values,
        np.full((n_rows, len(SCORE_COLS)), ""),
    ])

    z_matrix = np.column_stack([
        np.zeros(n_rows),          # Total → z=0 → white
        means[SCORE_COLS].values,
    ])

    fig = go.Figure(go.Heatmap(
        z=z_matrix,
        x=[x_total] + x_crit,
        y=uc_list,
        text=text_matrix,
        texttemplate="%{text}",
        textfont=dict(color="black", size=16),
        colorscale=colorscale, zmin=0, zmax=5,
        colorbar=dict(title="Score", tickvals=[1,2,3,4,5]),
        hovertemplate="Use case: %{y}<br>%{x}<br>%{text}<extra></extra>",
    ))
    fig.update_xaxes(
        tickvals=[x_total] + x_crit,
        ticktext=["Total"] + SCORE_COLS,
        tickangle=0,
        tickfont=dict(size=16),
    )
    fig.update_layout(
        margin=dict(l=260, r=0, t=20, b=40),
        paper_bgcolor="#ffffff",
        yaxis=dict(
            tickfont=dict(size=16),
            automargin=True,
            ticklabelposition="outside left",
        ),
        font=dict(size=16),
    )
    return fig


@app.callback(
    Output("stats-table", "children"),
    Input("partner-filter", "value"),
    Input("group-filter", "value"),
    Input("usecase-filter", "value"),
    Input("active-groups", "data"),
    Input("table-sort-by", "value"),
    Input("table-order", "value"),
    Input("leaderboard-top-n-mode", "value"),
    Input("leaderboard-top-n", "value"),
)
def stats_table(partners_sel, groups_sel, use_cases_sel, active_groups,
                sort_by, order, top_n_mode, top_n):
    df = apply_filters(
        partners_sel, groups_sel, use_cases_sel, active_groups, top_n, top_n_mode
    )
    if df.empty:
        return html.P(
            "No matching rows.",
            style={"color": "#6b7280", "fontSize": BODY_TEXT_SIZE, "margin": "8px 0"},
        )

    tbl = df[["Group", "Use Case Short", "Policy Question"] + SCORE_COLS + ["Total"]].copy()
    for col in SCORE_COLS + ["Total"]:
        tbl[col] = tbl[col].round(2)
    tbl = tbl.rename(columns={
        "Group": "Theme",
        "Use Case Short": "Use case",
        "Policy Question": "Policy question",
    })
    tbl["Overall rank"] = tbl["Total"].rank(method="dense", ascending=False).astype(int)
    tbl["Theme rank"] = (
        tbl.groupby("Theme")["Total"]
        .rank(method="dense", ascending=False)
        .astype(int)
    )
    tbl["Policy question"] = tbl["Policy question"].fillna("").map(separate_policy_questions)
    tbl = tbl[
        ["Theme", "Use case", "Policy question"] +
        SCORE_COLS + ["Total", "Overall rank", "Theme rank"]
    ]
    tbl = tbl.sort_values(
        sort_by,
        ascending=(order == "asc"),
        key=lambda s: s.map(sort_key) if s.name == "Theme" else s,
    )
    columns = [{"name": col, "id": col} for col in tbl.columns]
    compact_score_columns = SCORE_COLS + ["Total", "Overall rank", "Theme rank"]

    return dash_table.DataTable(
        data=tbl.to_dict("records"),
        columns=columns,
        page_action="none",
        style_table={"overflowX": "auto", "maxHeight": "820px", "overflowY": "auto"},
        style_cell={
            "fontFamily": "Inter, sans-serif",
            "fontSize": f"{TABLE_TEXT_SIZE}px",
            "padding": "8px",
            "textAlign": "left",
            "whiteSpace": "normal",
            "height": "auto",
            "lineHeight": "1.35",
        },
        style_header={
            "backgroundColor": "#3b82f6",
            "color": "white",
            "fontWeight": "600",
            "whiteSpace": "normal",
            "height": "auto",
        },
        style_data_conditional=[
            {"if": {"row_index": "even"}, "backgroundColor": "#f9fafb"},
            {"if": {"row_index": "odd"}, "backgroundColor": "#ffffff"},
            *[
                {
                    "if": {
                        "filter_query": f'{{Theme}} = "{datatable_filter_text(theme)}"',
                        "column_id": "Theme",
                    },
                    "backgroundColor": GROUP_COLOUR_MAP.get(theme, "#CCCCCC"),
                    "color": text_colour_for_background(GROUP_COLOUR_MAP.get(theme, "#CCCCCC")),
                    "fontWeight": "600",
                }
                for theme in groups
            ],
        ],
        style_cell_conditional=[
            {"if": {"column_id": "Theme"}, "width": "7%", "minWidth": "70px", "maxWidth": "90px"},
            {"if": {"column_id": "Use case"}, "width": "21%", "minWidth": "220px", "maxWidth": "340px"},
            {
                "if": {"column_id": "Policy question"},
                "width": "34%",
                "minWidth": "320px",
                "maxWidth": "560px",
                "whiteSpace": "pre-line",
            },
            *[
                {
                    "if": {"column_id": col},
                    "width": "5%",
                    "minWidth": "58px",
                    "maxWidth": "76px",
                    "textAlign": "center",
                }
                for col in compact_score_columns
            ],
        ],
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if _args.tunnel:
        _password = _args.password or getpass.getpass("Dashboard password: ")
        _pwd_hash = generate_password_hash(_password)

    _requested_port = _args.port or (DEFAULT_TUNNEL_PORT if (_args.tunnel or _args.lan) else None)
    try:
        _port = choose_port(_requested_port)
    except OSError as exc:
        print(f"ERROR: Port {_requested_port} is not available.")
        print(f"Reason: {exc}")
        print("Try another port, for example: --port 8081")
        sys.exit(1)

    _local_url = f"http://127.0.0.1:{_port}"
    _lan_urls = get_lan_urls(_port)
    _debug = bool(_args.debug)

    if _args.tunnel:
        _public_url = start_public_tunnel(_port)
        print(f"Public URL  : {_public_url}")
        print(f"Local URL   : {_local_url}")
        for i, _lan_url in enumerate(_lan_urls):
            print(f"{'LAN URL' if i == 0 else 'LAN URL alt'}  : {_lan_url}")
        print(f"Username    : {_args.username}")
        print("Share the public URL and credentials with your participants.")
        _host = "0.0.0.0"
    elif _args.lan:
        for i, _lan_url in enumerate(_lan_urls):
            print(f"{'LAN URL' if i == 0 else 'LAN URL alt'}  : {_lan_url}")
        print(f"Local URL   : {_local_url}")
        _host = "0.0.0.0"
    else:
        print(f"Local URL   : {_local_url}")
        _host = "127.0.0.1"

    app.run(
        host=_host,
        debug=_debug,
        port=_port,
        threaded=True,
        dev_tools_ui=_debug,
        dev_tools_props_check=_debug,
    )
