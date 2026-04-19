import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path
import numpy as np
import base64
import io

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

st.set_page_config(page_title="Tonga National Fuel Supply & Consumption Dashboard", layout="wide")

BASE_DIR = Path(__file__).resolve().parent


def resolve_existing_path(candidates):
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return path
    return Path(candidates[0])


DEFAULT_FILE = resolve_existing_path(
    [
        BASE_DIR / "Oil_Data_Consolidated.xlsx",
        Path(r"c:\Users\halea\Nextcloud\Documents\Petroleum\Oil_Data_Consolidated.xlsx"),
        BASE_DIR / "Oil_Data_Consolidated_updated.xlsx",
    ]
)
BACKGROUND_FILE = resolve_existing_path(
    [
        BASE_DIR / "background.jpg",
        BASE_DIR.parent / "doe-website" / "background.jpg",
        Path(r"c:\Users\halea\Nextcloud\Documents\doe-website\background.jpg"),
    ]
)
LOGO_FILE = resolve_existing_path(
    [
        BASE_DIR / "tonga-energy-logo.jpg",
        BASE_DIR.parent / "doe-website" / "public" / "assets" / "tonga-energy-logo.jpg",
        Path(r"c:\Users\halea\Nextcloud\Documents\doe-website\public\assets\tonga-energy-logo.jpg"),
    ]
)
CHART_COLORS = ["#7DD3FC", "#60A5FA", "#A78BFA", "#F59E0B", "#34D399"]
PRICE_FILE = BASE_DIR / "Transformed_for_Analysis.xlsx"


@st.cache_data
def load_data(file_obj, file_mtime=None):
    xls = pd.ExcelFile(file_obj)
    actual = pd.read_excel(file_obj, sheet_name="Actual")
    resupply = pd.read_excel(file_obj, sheet_name="Resupply")
    terminal = pd.read_excel(file_obj, sheet_name="Terminal")

    # Normalize dtypes for reliable filtering and charting.
    if "Date" in actual.columns:
        actual["Date"] = pd.to_datetime(actual["Date"], errors="coerce")
    if "Date" in resupply.columns:
        resupply["Date"] = pd.to_datetime(resupply["Date"], errors="coerce")

    for col in ["Closing Stock", "Offtake", "Tonga Power Offtake"]:
        if col in actual.columns:
            actual[col] = pd.to_numeric(actual[col], errors="coerce")

    if "Quantity" in resupply.columns:
        resupply["Quantity"] = pd.to_numeric(resupply["Quantity"], errors="coerce")

    if "Quantity" in terminal.columns:
        terminal["Quantity"] = pd.to_numeric(terminal["Quantity"], errors="coerce")

    return xls.sheet_names, actual, resupply, terminal


@st.cache_data
def load_price_data(file_path, mtime=None):
    fp = pd.read_excel(file_path, sheet_name="FuelPrice_Long")
    tr = pd.read_excel(file_path, sheet_name="Tariff_Long")
    fp["Date"] = pd.to_datetime(fp["Date"], errors="coerce")
    fp["Price"] = pd.to_numeric(fp["Price"], errors="coerce")
    tr["Value"] = pd.to_numeric(tr["Value"], errors="coerce")
    return fp, tr


def calculate_kpis(actual_df, resupply_df):
    """Calculate KPI metrics."""
    # Total Stock: sum of all latest closing stock values
    latest_actual = actual_df.dropna(subset=["Date", "Closing Stock"])
    if not latest_actual.empty:
        latest_actual = latest_actual.sort_values("Date")
        latest_actual = latest_actual.drop_duplicates(subset=["Company", "Location", "Fuel Type"], keep="last")
        total_stock = latest_actual["Closing Stock"].sum()
    else:
        total_stock = 0

    # Total Offtake excludes Tonga Power so the two KPI cards do not double count.
    offtake_data = actual_df.copy()
    offtake_data["Offtake"] = pd.to_numeric(offtake_data.get("Offtake"), errors="coerce").fillna(0)
    offtake_data["Tonga Power Offtake"] = pd.to_numeric(
        offtake_data.get("Tonga Power Offtake"), errors="coerce"
    ).fillna(0)
    total_offtake = (offtake_data["Offtake"] - offtake_data["Tonga Power Offtake"]).clip(lower=0).sum()
    total_consumption = offtake_data["Offtake"].sum()

    # Upcoming Supply: sum of all resupply quantities
    upcoming_data = resupply_df.dropna(subset=["Quantity"])
    upcoming_supply = upcoming_data["Quantity"].sum()

    return total_stock, total_offtake, upcoming_supply, total_consumption


def calculate_stock_by_fuel(actual_df):
    """Calculate latest stock by fuel type."""
    latest_actual = actual_df.dropna(subset=["Date", "Closing Stock"])
    if not latest_actual.empty:
        latest_actual = latest_actual.sort_values("Date")
        latest_actual = latest_actual.drop_duplicates(subset=["Company", "Location", "Fuel Type"], keep="last")
        stock_by_fuel = latest_actual.groupby("Fuel Type")["Closing Stock"].sum()
        return stock_by_fuel
    return pd.Series()


def calculate_days_of_cover(total_stock, daily_offtake):
    """Calculate days of cover."""
    if daily_offtake > 0:
        return total_stock / daily_offtake
    return 0


def calculate_cover_status(days_of_cover):
    """Classify stock risk status based on days of cover."""
    if days_of_cover >= 45:
        return "Safe"
    if days_of_cover >= 30:
        return "Watch"
    return "Critical"


def to_csv(df):
    """Convert DataFrame to CSV bytes."""
    return df.to_csv(index=False).encode("utf-8")


def checkbox_slicer(container, title, options, key_prefix):
    """Render a checkbox-style slicer and return selected options."""
    container.markdown(f'<div class="filter-title-chip">{title}</div>', unsafe_allow_html=True)

    clear_pending_key = f"{key_prefix}_clear_pending"

    for idx, _ in enumerate(options):
        option_key = f"{key_prefix}_{idx}"
        if option_key not in st.session_state:
            st.session_state[option_key] = True

    # Apply clear action before rendering checkbox widgets to avoid
    # modifying widget-bound keys after instantiation.
    if st.session_state.get(clear_pending_key, False):
        for idx, _ in enumerate(options):
            st.session_state[f"{key_prefix}_{idx}"] = False
        st.session_state[clear_pending_key] = False

    selected = []

    for idx, option in enumerate(options):
        checked = container.checkbox(
            str(option),
            key=f"{key_prefix}_{idx}",
        )
        if checked:
            selected.append(option)

    # Bottom-right action controls for each filter group.
    if container.button("Clear", key=f"{key_prefix}_clear"):
        st.session_state[clear_pending_key] = True
        st.rerun()

    return selected


def checkbox_slicer_horizontal(container, title, options, key_prefix):
    """Render checkbox options in a single horizontal row and return selected items."""
    container.markdown(f'<div class="filter-title-chip">{title}</div>', unsafe_allow_html=True)

    clear_pending_key = f"{key_prefix}_clear_pending"

    for idx, _ in enumerate(options):
        option_key = f"{key_prefix}_{idx}"
        if option_key not in st.session_state:
            st.session_state[option_key] = True

    if st.session_state.get(clear_pending_key, False):
        for idx, _ in enumerate(options):
            st.session_state[f"{key_prefix}_{idx}"] = False
        st.session_state[clear_pending_key] = False

    selected = []
    if options:
        row_cols = container.columns([1] * len(options) + [0.7], gap="small")
        for idx, option in enumerate(options):
            with row_cols[idx]:
                checked = st.checkbox(str(option), key=f"{key_prefix}_{idx}")
            if checked:
                selected.append(option)

        with row_cols[-1]:
            if st.button("Clear", key=f"{key_prefix}_clear"):
                st.session_state[clear_pending_key] = True
                st.rerun()
    else:
        if container.button("Clear", key=f"{key_prefix}_clear"):
            st.session_state[clear_pending_key] = True
            st.rerun()

    return selected


def format_compact(value):
    """Format numbers using compact K/M suffixes for KPI cards."""
    if pd.isna(value):
        return "0"
    abs_value = abs(float(value))
    if abs_value >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"
    if abs_value >= 1_000:
        return f"{value / 1_000:.0f}K"
    return f"{value:.0f}"


def fuel_icon(fuel_name):
    """Return a compact icon per fuel label for KPI readability."""
    fuel_text = str(fuel_name).strip().lower()
    if "diesel" in fuel_text:
        return "🛢️"
    if "petrol" in fuel_text or "gasoline" in fuel_text:
        return "⛽"
    if "kerosene" in fuel_text or "jet" in fuel_text:
        return "✈️"
    return "🔹"


def render_kpi_tile(container, icon, label, value, accent):
    """Render a compact KPI tile with a colored accent for better scanability."""
    container.markdown(
        f"""
        <div class="kpi-tile" style="border-left: 3px solid {accent};">
            <div class="kpi-label">{icon} {label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_category(container, title, items, class_name=""):
    """Render a compact category card with tightly controlled KPI spacing."""
    items_html = "".join(
        [
            f'<div class="kpi-mini" style="border-left: 3px solid {accent};">'
            f'<div class="kpi-mini-line">'
            f'<div class="kpi-mini-label">{icon} {label}</div>'
            f'<div class="kpi-mini-value">{value}</div>'
            f'</div>'
            f'</div>'
            for icon, label, value, accent in items
        ]
    )
    container.markdown(
        (
            f'<div class="kpi-category-card {class_name}">'
            f'<div class="kpi-category-title">{title}</div>'
            f'<div class="kpi-mini-row">{items_html}</div>'
            f'</div>'
        ),
        unsafe_allow_html=True,
    )


def render_kpi_group(container, title, items):
    """Render one grouped KPI block with a single border around title and cards."""
    cards_html = "".join(
        [
            (
                f'<div class="kpi-group-card" style="border-left: 3px solid {accent};">'
                f'<div class="kpi-group-card-label">{icon} {label}</div>'
                f'<div class="kpi-group-card-value">{value}</div>'
                f'</div>'
            )
            for icon, label, value, accent in items
        ]
    )
    container.markdown(
        (
            '<div class="kpi-group-frame">'
            f'<div class="kpi-group-title">{title}</div>'
            f'<div class="kpi-group-row">{cards_html}</div>'
            '</div>'
        ),
        unsafe_allow_html=True,
    )


def render_chart_title(container, title):
    """Render a highlighted chart title bar inside chart containers."""
    container.markdown(f'<div class="chart-title-bar">{title}</div>', unsafe_allow_html=True)

def set_app_background(image_path):
    """Set dashboard background image from a local file."""
    if not image_path.exists():
        return

    encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(8, 13, 20, 0.44), rgba(8, 13, 20, 0.44)), url("data:image/jpeg;base64,{encoded}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}

        /* Dark theme: even stronger overlay for max contrast */
        @media (prefers-color-scheme: dark) {{
            .stApp {{
                background-image: linear-gradient(rgba(5, 9, 14, 0.55), rgba(5, 9, 14, 0.55)), url("data:image/jpeg;base64,{encoded}");
            }}
        }}

        [data-testid="stHeader"] {{
            background: rgba(0, 0, 0, 0);
        }}
        [data-testid="stToolbar"] {{
            right: 0.5rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def apply_chart_theme(fig, height=400, hovermode=None, x_title=None, y_title=None, date_x=False):
    """Apply a consistent high-contrast chart style and readable formatting."""
    fig.update_layout(
        height=height,
        colorway=CHART_COLORS,
        paper_bgcolor="rgba(0, 0, 0, 0)",
        plot_bgcolor="rgba(12, 18, 26, 0.86)",
        font=dict(color="#E6EDF5", size=14),
        title=dict(x=0.02, xanchor="left", font=dict(size=18, color="#F8FBFF")),
        margin=dict(l=16, r=16, t=62, b=68),
        bargap=0.18,
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.16,
            xanchor="left",
            x=0,
            font=dict(color="#E6EDF5", size=13),
        ),
    )
    if hovermode:
        fig.update_layout(hovermode=hovermode)
    fig.update_xaxes(
        showgrid=True,
        gridcolor="rgba(173, 191, 210, 0.22)",
        zeroline=False,
        linecolor="rgba(173, 191, 210, 0.35)",
        tickfont=dict(color="#E6EDF5", size=13),
        title_font=dict(color="#E6EDF5", size=14),
        title_text=x_title,
        tickformat="%d %b" if date_x else None,
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="rgba(173, 191, 210, 0.22)",
        zeroline=False,
        linecolor="rgba(173, 191, 210, 0.35)",
        tickfont=dict(color="#E6EDF5", size=13),
        title_font=dict(color="#E6EDF5", size=14),
        title_text=y_title,
        tickformat=",.0f",
        separatethousands=True,
    )
    fig.update_traces(hoverlabel=dict(font=dict(color="#0B1220")))
    return fig


set_app_background(BACKGROUND_FILE)

logo_html = ""
footer_logo_html = ""
if LOGO_FILE.exists():
    logo_b64 = base64.b64encode(LOGO_FILE.read_bytes()).decode("utf-8")
    logo_html = (
        f'<img src="data:image/jpeg;base64,{logo_b64}" '
        'style="height:78px; width:auto; object-fit:contain; border-radius:8px; margin-right:0.75rem; border:1px solid rgba(173, 191, 210, 0.45);" '
        'alt="Organization Logo" />'
    )
    footer_logo_html = (
        f'<img src="data:image/jpeg;base64,{logo_b64}" '
        'style="height:30px; width:auto; object-fit:contain; border-radius:4px;" '
        'alt="DOE Logo" />'
    )

st.markdown(
    f"""
    <div style="
        padding: 0.14rem 0.62rem;
        margin: 0.01rem 0 0.24rem 0;
        border: 1.6px solid rgba(125, 211, 252, 0.72);
        border-radius: 12px;
        background: linear-gradient(180deg, rgba(30, 64, 175, 0.12), rgba(14, 22, 32, 0.66));
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.2), 0 0 10px rgba(56, 189, 248, 0.16);
        backdrop-filter: blur(2px);
    ">
        <div style="display:flex; align-items:center; gap:0.25rem;">
            {logo_html}
            <div>
                <h1 style="margin: 0; font-size: 2.2rem; line-height: 1.2; font-weight: 700; letter-spacing: 0.2px;">Tonga National Fuel Supply and Consumption Dashboard</h1>
                <p style="margin: 0.18rem 0 0 0; opacity: 0.88; font-size: 1.08rem; line-height: 1.3;">
                    Staged Energy and Fuel Supply Plan Monitoring
                </p>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 0.5rem;
        padding-bottom: 0.55rem;
        margin-left: auto;
        margin-right: auto;
    }

    div[data-testid="stHorizontalBlock"] {
        gap: 0.5rem;
    }

    .stApp {
        color: #F2F6FA;
        font-size: 16px;
    }

    .stMarkdown p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] div {
        font-size: 0.98rem;
    }

    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3 {
        font-size: 1.45rem !important;
        line-height: 1.25;
        font-weight: 700;
        margin-top: 0.08rem;
        margin-bottom: 0.35rem;
        border: 1.8px solid rgba(125, 211, 252, 0.72);
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(30, 64, 175, 0.14), rgba(14, 22, 32, 0.66));
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.2), 0 0 10px rgba(56, 189, 248, 0.16);
        padding: 0.26rem 0.62rem;
        color: #F4FAFF;
    }

    [data-testid="stMarkdownContainer"] h3 {
        font-size: 1.18rem !important;
        line-height: 1.3;
    }

    div[data-testid="stMetric"] {
        padding: 0.1rem 0.2rem;
        color: #F2F6FA;
    }

    div[data-testid="stMetricLabel"],
    div[data-testid="stMetricValue"] {
        color: #F2F6FA;
    }

    div[data-testid="stMetricLabel"] {
        font-size: 0.98rem;
        line-height: 1.2;
    }

    div[data-testid="stMetricValue"] {
        font-size: 1.34rem;
        line-height: 1.15;
        font-weight: 500 !important;
    }

    div[data-testid="stCaptionContainer"] {
        font-size: 1.02rem;
        letter-spacing: 0.2px;
    }

    .kpi-tile {
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(173, 191, 210, 0.18);
        border-radius: 8px;
        padding: 0.45rem 0.5rem;
        min-height: 66px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.45rem;
        margin-bottom: 0.14rem;
    }

    .kpi-group-title {
        border: none;
        border-radius: 8px;
        background: linear-gradient(90deg, rgba(30, 64, 175, 0.42), rgba(14, 116, 144, 0.36));
        padding: 0.24rem 0.42rem;
        font-size: 0.9rem;
        font-weight: 650;
        line-height: 1.2;
        margin-bottom: 0.35rem;
        text-align: center;
        color: #F8FCFF;
        letter-spacing: 0.2px;
        box-shadow: inset 0 0 0 1px rgba(125, 211, 252, 0.22);
    }

    .kpi-group-frame {
        border: 1.6px solid rgba(125, 211, 252, 0.72);
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(14, 116, 144, 0.1), rgba(14, 22, 32, 0.68));
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.2), 0 0 10px rgba(56, 189, 248, 0.16);
        padding: 0.46rem 0.5rem 0.62rem 0.5rem;
    }

    .kpi-group-row {
        display: flex;
        gap: 0.35rem;
        align-items: stretch;
        flex-wrap: wrap;
    }

    .kpi-group-card {
        flex: 1 1 180px;
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(173, 191, 210, 0.18);
        border-radius: 8px;
        padding: 0.45rem 0.5rem;
        min-height: 66px;
        min-width: 0;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.45rem;
    }

    .kpi-group-card-label {
        font-size: 0.9rem;
        line-height: 1.12;
        opacity: 0.96;
        font-weight: 460 !important;
        white-space: normal;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    .kpi-group-card-value {
        font-size: 1.28rem;
        line-height: 1.05;
        font-weight: 520 !important;
        color: #F8FBFF;
        margin-left: auto;
        white-space: nowrap;
    }

    .chart-title-bar {
        border: 1px solid rgba(125, 211, 252, 0.58);
        border-radius: 10px;
        background: linear-gradient(90deg, rgba(30, 64, 175, 0.22), rgba(14, 116, 144, 0.18));
        color: #F4FAFF;
        font-size: 1rem;
        font-weight: 650;
        line-height: 1.2;
        padding: 0.28rem 0.62rem;
        text-align: center;
        margin: 0.02rem 0 0.38rem 0;
        box-shadow: inset 0 0 0 1px rgba(186, 230, 253, 0.1);
    }

    .st-key-filter_company_group,
    .st-key-filter_location_group,
    .st-key-filter_fuel_group,
    .st-key-filter_month_group,
    .st-key-terminal_location_filter_box,
    .st-key-terminal_type_filter_box,
    .st-key-filter_company_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-filter_location_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-filter_fuel_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-filter_month_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-terminal_location_filter_box div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-terminal_type_filter_box div[data-testid="stVerticalBlockBorderWrapper"] {
        border: 1.8px solid rgba(125, 211, 252, 0.72) !important;
        border-radius: 14px !important;
        background: linear-gradient(180deg, rgba(30, 64, 175, 0.12), rgba(14, 22, 32, 0.66)) !important;
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.2), 0 0 10px rgba(56, 189, 248, 0.16);
        padding: 0.3rem 0.45rem 0.35rem 0.45rem !important;
        margin-bottom: 0.35rem;
    }

    .st-key-chart_stock,
    .st-key-chart_offtake,
    .st-key-chart_location,
    .st-key-chart_resupply,
    .st-key-chart_terminal,
    .st-key-chart_fuel_price,
    .st-key-chart_tariff,
    .st-key-chart_dependency,
    .st-key-chart_stock div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_offtake div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_location div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_resupply div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_terminal div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_fuel_price div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_tariff div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-chart_dependency div[data-testid="stVerticalBlockBorderWrapper"] {
        border: 1.8px solid rgba(125, 211, 252, 0.72) !important;
        border-radius: 14px !important;
        background: linear-gradient(180deg, rgba(30, 64, 175, 0.12), rgba(14, 22, 32, 0.66)) !important;
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.2), 0 0 10px rgba(56, 189, 248, 0.16);
    }

    .filter-title-chip {
        border: 1px solid rgba(125, 211, 252, 0.62);
        border-radius: 8px;
        background: linear-gradient(90deg, rgba(30, 64, 175, 0.2), rgba(14, 116, 144, 0.16));
        color: #F4FAFF;
        font-size: 0.9rem;
        font-weight: 650;
        line-height: 1.2;
        padding: 0.2rem 0.42rem;
        margin: 0.06rem 0 0.3rem 0;
        box-shadow: inset 0 0 0 1px rgba(186, 230, 253, 0.08);
        text-align: center;
    }

    .st-key-filter_company_group div[data-testid="stButton"],
    .st-key-filter_location_group div[data-testid="stButton"],
    .st-key-filter_fuel_group div[data-testid="stButton"],
    .st-key-filter_month_group div[data-testid="stButton"] {
        display: flex;
        justify-content: flex-end;
        margin-top: 0.2rem;
    }

    .st-key-filter_company_group div[data-testid="stButton"] > button,
    .st-key-filter_location_group div[data-testid="stButton"] > button,
    .st-key-filter_fuel_group div[data-testid="stButton"] > button,
    .st-key-filter_month_group div[data-testid="stButton"] > button {
        font-size: 0.76rem;
        line-height: 1.1;
        min-height: 1.68rem;
        min-width: 64px;
        white-space: nowrap;
        padding: 0.08rem 0.42rem;
        border-radius: 7px;
        border: 1px solid rgba(125, 211, 252, 0.78);
        background: linear-gradient(180deg, rgba(30, 64, 175, 0.32), rgba(14, 22, 32, 0.74));
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.18), 0 0 8px rgba(56, 189, 248, 0.2);
    }

    .st-key-filter_company_group div[data-testid="stButton"] > button:hover,
    .st-key-filter_location_group div[data-testid="stButton"] > button:hover,
    .st-key-filter_fuel_group div[data-testid="stButton"] > button:hover,
    .st-key-filter_month_group div[data-testid="stButton"] > button:hover {
        border-color: rgba(125, 211, 252, 0.95);
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.26), 0 0 12px rgba(56, 189, 248, 0.28);
    }

    /* Highlight the single group-defining border (subtitle + cards together). */
    .st-key-kpi_supply_group,
    .st-key-kpi_offtake_group,
    .st-key-kpi_fuel_group,
    .st-key-kpi_coverage_group {
        border: 2px solid rgba(125, 211, 252, 0.7) !important;
        border-radius: 14px !important;
        box-shadow: 0 0 0 1px rgba(186, 230, 253, 0.26), 0 0 16px rgba(56, 189, 248, 0.22);
        background: linear-gradient(180deg, rgba(14, 116, 144, 0.16), rgba(14, 22, 32, 0.68)) !important;
        padding: 0.34rem 0.34rem 0.56rem 0.34rem !important;
        overflow: hidden;
    }

    .st-key-kpi_supply_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-kpi_offtake_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-kpi_fuel_group div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-kpi_coverage_group div[data-testid="stVerticalBlockBorderWrapper"] {
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
        padding: 0.48rem 0.48rem 0.98rem 0.48rem !important;
    }

    .kpi-label {
        font-size: 0.9rem;
        line-height: 1.12;
        opacity: 0.96;
        font-weight: 460 !important;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    .kpi-value {
        font-size: 1.28rem;
        line-height: 1.05;
        font-weight: 520 !important;
        color: #F8FBFF;
        margin-left: auto;
        white-space: nowrap;
    }

    .kpi-category-card {
        border: 1px solid rgba(173, 191, 210, 0.38);
        border-radius: 12px;
        background: rgba(14, 22, 32, 0.78);
        backdrop-filter: blur(3px);
        padding: 0.16rem 0.28rem 0.2rem 0.28rem;
    }

    .kpi-category-card.compact-wide {
        max-width: 860px;
        margin-left: auto;
        margin-right: auto;
    }

    .kpi-category-card.compact-wide .kpi-mini-row {
        justify-content: center;
    }

    .kpi-category-card.compact-wide .kpi-mini {
        flex: 0 1 220px;
    }

    .kpi-category-title {
        font-size: 0.98rem;
        font-weight: 500 !important;
        margin: 0 0 0.08rem 0;
        line-height: 1.05;
    }

    .kpi-mini-row {
        display: flex;
        gap: 0.24rem;
        align-items: stretch;
    }

    .kpi-mini {
        flex: 1;
        background: rgba(255, 255, 255, 0.04);
        border-radius: 8px;
        padding: 0.62rem 0.38rem;
        min-height: 100px;
        display: flex;
        align-items: center;
    }

    .kpi-mini-line {
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.5rem;
    }

    .kpi-mini-label {
        font-size: 0.9rem;
        line-height: 1.1;
        opacity: 0.96;
        font-weight: 460 !important;
        white-space: nowrap;
    }

    .kpi-mini-value {
        margin-top: 0;
        font-size: 1.12rem;
        line-height: 1.05;
        font-weight: 500 !important;
        color: #F8FBFF;
        white-space: nowrap;
    }

    /* Strong override for custom KPI blocks */
    .kpi-category-card .kpi-mini .kpi-mini-value,
    .kpi-category-card .kpi-mini .kpi-mini-label,
    .kpi-category-card .kpi-category-title {
        font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
    }

    /* Compact only the top KPI category boxes */
    .st-key-kpi_supply,
    .st-key-kpi_offtake,
    .st-key-kpi_coverage {
        padding-top: 0;
        padding-bottom: 0;
    }

    .st-key-kpi_supply div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-kpi_offtake div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-kpi_coverage div[data-testid="stVerticalBlockBorderWrapper"] {
        padding-top: 0.08rem;
        padding-bottom: 0.2rem;
    }

    .st-key-kpi_supply div[data-testid="stMetric"],
    .st-key-kpi_offtake div[data-testid="stMetric"],
    .st-key-kpi_coverage div[data-testid="stMetric"] {
        padding: 0 0.06rem;
    }

    .st-key-kpi_supply div[data-testid="stMetricValue"],
    .st-key-kpi_offtake div[data-testid="stMetricValue"],
    .st-key-kpi_coverage div[data-testid="stMetricValue"] {
        font-size: 1.12rem;
    }

    .st-key-kpi_supply div[data-testid="stMetricLabel"],
    .st-key-kpi_offtake div[data-testid="stMetricLabel"],
    .st-key-kpi_coverage div[data-testid="stMetricLabel"] {
        font-size: 1rem;
    }

    .st-key-kpi_supply div[data-testid="stCaptionContainer"],
    .st-key-kpi_offtake div[data-testid="stCaptionContainer"],
    .st-key-kpi_coverage div[data-testid="stCaptionContainer"] {
        margin-bottom: 0;
        margin-top: 0;
        line-height: 1.05;
    }

    .st-key-kpi_supply .kpi-tile,
    .st-key-kpi_offtake .kpi-tile,
    .st-key-kpi_coverage .kpi-tile {
        margin-top: -0.05rem;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: rgba(14, 22, 32, 0.78);
        border: 1px solid rgba(173, 191, 210, 0.38);
        border-radius: 12px;
        backdrop-filter: blur(3px);
        padding-top: 0.2rem;
        padding-bottom: 0.2rem;
    }

    div[data-testid="stCaptionContainer"] {
        margin-bottom: 0.1rem;
    }

    div[data-testid="stTabs"] {
        margin-top: 0;
    }

    hr {
        margin-top: 0.3rem !important;
        margin-bottom: 0.3rem !important;
    }

    [data-testid="stSidebar"] {
        background: rgba(9, 15, 23, 0.88);
        border-right: 1px solid rgba(173, 191, 210, 0.32);
    }

    [data-testid="stSidebar"] * {
        color: #F2F6FA !important;
    }

    @media (max-width: 900px) {
        .stApp {
            font-size: 15px;
        }
        [data-testid="stMarkdownContainer"] h2 {
            font-size: 1.28rem !important;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.2rem;
        }
        .kpi-group-frame {
            padding: 0.38rem 0.4rem 0.48rem 0.4rem;
        }
        .kpi-group-row {
            gap: 0.3rem;
        }
        .kpi-group-card {
            flex: 1 1 100%;
            min-height: 58px;
            padding: 0.38rem 0.42rem;
        }
        .kpi-group-card-label {
            font-size: 0.84rem;
            line-height: 1.15;
        }
        .kpi-group-card-value {
            font-size: 1.08rem;
        }
        .kpi-group-title {
            margin-bottom: 0.28rem;
            font-size: 0.86rem;
        }
        .st-key-filter_company_group div[data-testid="stButton"],
        .st-key-filter_location_group div[data-testid="stButton"],
        .st-key-filter_fuel_group div[data-testid="stButton"],
        .st-key-filter_month_group div[data-testid="stButton"] {
            justify-content: stretch;
        }
        .st-key-filter_company_group div[data-testid="stButton"] > button,
        .st-key-filter_location_group div[data-testid="stButton"] > button,
        .st-key-filter_fuel_group div[data-testid="stButton"] > button,
        .st-key-filter_month_group div[data-testid="stButton"] > button {
            width: 100%;
            min-width: 0;
        }
    }

    </style>
    """,
    unsafe_allow_html=True,
)

file_to_use = DEFAULT_FILE

if not DEFAULT_FILE.exists():
    st.error(f"Default file not found at {DEFAULT_FILE}")
    st.stop()

try:
    file_mtime = file_to_use.stat().st_mtime
    sheets, actual_df, resupply_df, terminal_df = load_data(file_to_use, file_mtime=file_mtime)
except Exception as exc:
    st.error(f"Failed to load workbook: {exc}")
    st.stop()

last_sync = pd.to_datetime(file_mtime, unit="s").strftime("%d %b %Y %H:%M")

price_df, tariff_df = None, None
if PRICE_FILE.exists():
    try:
        price_df, tariff_df = load_price_data(str(PRICE_FILE), mtime=PRICE_FILE.stat().st_mtime)
    except Exception:
        pass

# Sidebar filters
st.sidebar.title("🎛️ Filters")

_SECTIONS = ["📊 Fuel Supply", "📦 Terminal Data", "💰 Prices & Tariffs"]
active_section = st.sidebar.radio("Section", _SECTIONS, key="active_section", label_visibility="collapsed")
st.sidebar.divider()

companies = sorted([x for x in actual_df["Company"].dropna().unique()])
locations = sorted([x for x in actual_df["Location"].dropna().unique()])
fuels = sorted([x for x in actual_df["Fuel Type"].dropna().unique()])
months = (
    sorted(actual_df["Date"].dropna().dt.to_period("M").astype(str).unique())
    if "Date" in actual_df.columns
    else []
)

if active_section != "💰 Prices & Tariffs":
    company_filter_box = st.sidebar.container(border=True, key="filter_company_group")
    location_filter_box = st.sidebar.container(border=True, key="filter_location_group")
    fuel_filter_box = st.sidebar.container(border=True, key="filter_fuel_group")
    month_filter_box = st.sidebar.container(border=True, key="filter_month_group")
    company_sel = checkbox_slicer(company_filter_box, "Company", companies, "company")
    location_sel = checkbox_slicer(location_filter_box, "Location", locations, "location")
    fuel_sel = checkbox_slicer(fuel_filter_box, "Fuel Type", fuels, "fuel")
    month_sel = checkbox_slicer(month_filter_box, "Month", months, "month")
else:
    company_sel = companies
    location_sel = locations
    fuel_sel = fuels
    month_sel = months
    # Price & Tariff sidebar filters
    _price_opts = sorted(price_df["Price_Type"].dropna().unique().tolist()) if price_df is not None else []
    _fuel_opts = sorted(price_df["Fuel"].dropna().unique().tolist()) if price_df is not None else []
    _comp_opts = ["Fuel Component", "Non Fuel component", "Total Tariff"]
    price_type_filter_box = st.sidebar.container(border=True, key="filter_price_type_group")
    price_fuel_filter_box = st.sidebar.container(border=True, key="filter_price_fuel_group")
    tariff_comp_filter_box = st.sidebar.container(border=True, key="filter_tariff_comp_group")
    price_type_sel = checkbox_slicer(price_type_filter_box, "Price Type", _price_opts, "fp_price_type")
    fuel_sel_fp = checkbox_slicer(price_fuel_filter_box, "Fuel", _fuel_opts, "fp_fuel")
    comp_sel = checkbox_slicer(tariff_comp_filter_box, "Tariff Component", _comp_opts, "tariff_comp")

actual_df_for_filter = actual_df.copy()
if "Date" in actual_df_for_filter.columns:
    actual_df_for_filter["Month"] = actual_df_for_filter["Date"].dt.to_period("M").astype(str)

resupply_df_for_filter = resupply_df.copy()
if "Date" in resupply_df_for_filter.columns:
    resupply_df_for_filter["Month"] = resupply_df_for_filter["Date"].dt.to_period("M").astype(str)

# Apply filters
filtered_actual = actual_df_for_filter[
    actual_df_for_filter["Company"].isin(company_sel)
    & actual_df_for_filter["Location"].isin(location_sel)
    & actual_df_for_filter["Fuel Type"].isin(fuel_sel)
    & actual_df_for_filter["Month"].isin(month_sel)
].copy()

filtered_resupply = resupply_df_for_filter[
    resupply_df_for_filter["Company"].isin(company_sel)
    & resupply_df_for_filter["Location"].isin(location_sel)
    & resupply_df_for_filter["Fuel Type"].isin(fuel_sel)
    & resupply_df_for_filter["Month"].isin(month_sel)
].copy()

# Display KPIs
st.subheader("Key Performance Indicators")
total_stock, non_power_offtake, upcoming_supply, total_consumption = calculate_kpis(filtered_actual, filtered_resupply)
tonga_power_offtake = (
    filtered_actual["Tonga Power Offtake"].dropna().sum()
    if "Tonga Power Offtake" in filtered_actual.columns
    else 0
)
avg_daily_offtake = total_consumption / len(filtered_actual["Date"].unique()) if len(filtered_actual["Date"].unique()) > 0 else 0
days_of_cover = calculate_days_of_cover(total_stock, avg_daily_offtake)
cover_status = calculate_cover_status(days_of_cover)

status_icon = "🟢" if cover_status == "Safe" else "🟡" if cover_status == "Watch" else "🔴"
status_accent = "#22C55E" if cover_status == "Safe" else "#F59E0B" if cover_status == "Watch" else "#EF4444"

supply_items = [
    ("🔶", "Total Fuel", f"{format_compact(total_stock)}", "#60A5FA"),
    ("📈", "Upcoming Supply", f"{format_compact(upcoming_supply)}", "#34D399"),
]
offtake_items = [
    ("📊", "Total Offtake", f"{format_compact(total_consumption)}", "#7DD3FC"),
    ("⚡", "Tonga Power", f"{format_compact(tonga_power_offtake)}", "#F59E0B"),
    ("📉", "Non-Power", f"{format_compact(non_power_offtake)}", "#A78BFA"),
]
coverage_items = [
    ("⏱️", "Days", f"{days_of_cover:.2f}", "#38BDF8"),
    (status_icon, "Status", cover_status, status_accent),
]

stock_by_fuel = calculate_stock_by_fuel(filtered_actual)

filter_text = {
    "Company": ", ".join([str(x) for x in company_sel]) if company_sel else "All",
    "Location": ", ".join([str(x) for x in location_sel]) if location_sel else "All",
    "Fuel Type": ", ".join([str(x) for x in fuel_sel]) if fuel_sel else "All",
    "Month": ", ".join([str(x) for x in month_sel]) if month_sel else "All",
}

stock_rows = []
if not stock_by_fuel.empty:
    for fuel_name, stock_value in stock_by_fuel.items():
        stock_rows.append((str(fuel_name), float(stock_value)))

latest_prices = pd.DataFrame(columns=["Price_Type", "Fuel", "Price"])
if price_df is not None and not price_df.empty:
    latest_prices = (
        price_df.dropna(subset=["Date", "Price"])
        .sort_values("Date")
        .drop_duplicates(subset=["Price_Type", "Fuel"], keep="last")
    )

latest_tariff_year = ""
latest_tariff_rows = []
if tariff_df is not None and not tariff_df.empty:
    tariff_main = tariff_df[tariff_df["Component"].isin(["Fuel Component", "Non Fuel component", "Total Tariff"])].copy()
    if not tariff_main.empty:
        latest_tariff_year = sorted(tariff_main["Year"].dropna().astype(str).unique().tolist())[-1]
        tariff_avg = (
            tariff_main[tariff_main["Year"].astype(str) == latest_tariff_year]
            .groupby("Component", as_index=False)["Value"]
            .mean()
        )
        latest_tariff_rows = [(str(r.Component), float(r.Value)) for r in tariff_avg.itertuples(index=False)]


def build_summary_pdf_bytes():
    if not REPORTLAB_AVAILABLE:
        return None

    pdf_buffer = io.BytesIO()
    page_w, page_h = landscape(A4)
    pdf = canvas.Canvas(pdf_buffer, pagesize=(page_w, page_h))

    y = page_h - 30
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(24, y, "Tonga Fuel Dashboard - Weekly Summary")
    y -= 18
    pdf.setFont("Helvetica", 9)
    pdf.drawString(24, y, f"Printed: {pd.Timestamp.now().strftime('%d %b %Y %H:%M')}")
    y -= 12
    pdf.drawString(24, y, f"Data Source: {file_to_use.name} | Last Sync: {last_sync}")
    y -= 12
    pdf.drawString(
        24,
        y,
        (
            f"Filters - Company: {filter_text['Company']} | Location: {filter_text['Location']} | "
            f"Fuel: {filter_text['Fuel Type']} | Month: {filter_text['Month']}"
        ),
    )

    y -= 20
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(24, y, "Key Metrics")
    y -= 14
    pdf.setFont("Helvetica", 9)
    pdf.drawString(28, y, f"Total Fuel: {total_stock:,.0f} L")
    pdf.drawString(220, y, f"Upcoming Supply: {upcoming_supply:,.0f} L")
    pdf.drawString(430, y, f"Total Offtake: {total_consumption:,.0f} L")
    pdf.drawString(620, y, f"Days of Cover: {days_of_cover:.2f} ({cover_status})")

    y -= 22
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(24, y, "Stock by Fuel Type")
    y -= 12
    pdf.setFont("Helvetica", 9)
    for fuel_name, stock_value in stock_rows[:8]:
        pdf.drawString(28, y, f"{fuel_name}")
        pdf.drawRightString(200, y, f"{stock_value:,.0f} L")
        y -= 11
    if not stock_rows:
        pdf.drawString(28, y, "No stock data")
        y -= 11

    y -= 8
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(24, y, "Latest Fuel Prices (Retail/Wholesale)")
    y -= 12
    pdf.setFont("Helvetica", 9)
    if not latest_prices.empty:
        for row in latest_prices.itertuples(index=False):
            pdf.drawString(28, y, f"{row.Price_Type} - {row.Fuel}")
            pdf.drawRightString(260, y, f"{float(row.Price):.2f}")
            y -= 11
            if y < 110:
                break
    else:
        pdf.drawString(28, y, "No fuel price data")
        y -= 11

    y -= 8
    pdf.setFont("Helvetica-Bold", 10)
    tariff_title = "Latest Average Tariff Components"
    if latest_tariff_year:
        tariff_title += f" ({latest_tariff_year})"
    pdf.drawString(24, y, tariff_title)
    y -= 12
    pdf.setFont("Helvetica", 9)
    if latest_tariff_rows:
        for component_name, component_value in latest_tariff_rows[:8]:
            pdf.drawString(28, y, component_name)
            pdf.drawRightString(300, y, f"{component_value:.4f} T$/kWh")
            y -= 11
    else:
        pdf.drawString(28, y, "No tariff data")

    pdf.showPage()
    pdf.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


if "summary_pdf_bytes" not in st.session_state:
    st.session_state.summary_pdf_bytes = None

# Row 1: Supply + Offtake (with group titles)
top_supply_col, top_offtake_col = st.columns([2, 3], gap="small")
render_kpi_group(top_supply_col, "Supply", supply_items)
render_kpi_group(top_offtake_col, "Offtake Breakdown", offtake_items)

st.markdown("<div style='height: 0.2rem;'></div>", unsafe_allow_html=True)

# Row 2: Stock by Fuel (left) + Coverage (right)
bottom_left, bottom_right = st.columns([3, 2], gap="small")

with bottom_left:
    if not stock_by_fuel.empty:
        fuel_items = []
        for i, (fuel, stock) in enumerate(stock_by_fuel.items()):
            fuel_items.append((fuel_icon(fuel), str(fuel), f"{format_compact(stock)}", CHART_COLORS[i % len(CHART_COLORS)]))
        render_kpi_group(bottom_left, "Stock by Fuel Type", fuel_items)
    else:
        st.caption("No stock by fuel data")

with bottom_right:
    render_kpi_group(bottom_right, "Coverage", coverage_items)

st.divider()

# Main visualization area
if active_section == "📊 Fuel Supply":
    # Row 1: Key Visualizations
    col1, col2 = st.columns(2)
    
    # Chart 1: Stock Over Time by Fuel Type
    with col1:
        stock_data = filtered_actual.dropna(subset=["Date", "Closing Stock"])
        if not stock_data.empty:
            stock_by_date_fuel = stock_data.groupby(["Date", "Fuel Type"])["Closing Stock"].mean().reset_index()
            fig_stock = px.line(
                stock_by_date_fuel,
                x="Date",
                y="Closing Stock",
                color="Fuel Type",
                title="Stock Level Over Time by Fuel Type",
                markers=True,
                color_discrete_sequence=CHART_COLORS,
            )
            apply_chart_theme(
                fig_stock,
                height=400,
                hovermode="x unified",
                x_title="Date",
                y_title="Closing Stock (L)",
                date_x=True,
            )
            fig_stock.update_layout(title_text="")
            with st.container(border=True, key="chart_stock"):
                render_chart_title(st, "Stock Level Over Time by Fuel Type")
                st.plotly_chart(fig_stock, width='stretch')
        else:
            st.info("No stock data available")
    
    # Chart 2: Fuel Offtake Over Time
    with col2:
        offtake_data = filtered_actual.dropna(subset=["Date", "Offtake"])
        if not offtake_data.empty:
            offtake_by_date_fuel = offtake_data.groupby(["Date", "Fuel Type"])["Offtake"].sum().reset_index()
            fig_offtake = px.bar(
                offtake_by_date_fuel,
                x="Date",
                y="Offtake",
                color="Fuel Type",
                title="Daily Offtake (Take-off) by Fuel Type",
                barmode="stack",
                color_discrete_sequence=CHART_COLORS,
            )

            # Add Tonga Power as an explicit overlay series for direct comparison.
            if "Tonga Power Offtake" in offtake_data.columns:
                tonga_power_by_date = (
                    offtake_data.groupby("Date", as_index=False)["Tonga Power Offtake"]
                    .sum(min_count=1)
                    .fillna(0)
                )
                if not tonga_power_by_date.empty and (tonga_power_by_date["Tonga Power Offtake"] > 0).any():
                    fig_offtake.add_trace(
                        go.Scatter(
                            x=tonga_power_by_date["Date"],
                            y=tonga_power_by_date["Tonga Power Offtake"],
                            mode="lines+markers",
                            name="Tonga Power Offtake",
                            line=dict(color="#FDE047", width=2.5),
                            marker=dict(size=6, color="#FDE047"),
                        )
                    )

            apply_chart_theme(
                fig_offtake,
                height=400,
                hovermode="x unified",
                x_title="Date",
                y_title="Offtake (L)",
                date_x=True,
            )
            fig_offtake.update_layout(title_text="")
            with st.container(border=True, key="chart_offtake"):
                render_chart_title(st, "Daily Offtake (Take-off) by Fuel Type")
                st.plotly_chart(fig_offtake, width='stretch')
        else:
            st.info("No offtake data available")
    
    st.divider()
    
    # Row 2: Comparative Analysis
    col3, col4 = st.columns(2)
    
    # Chart 3: Stock by Location
    with col3:
        stock_by_location = filtered_actual.dropna(subset=["Location", "Closing Stock"])
        if not stock_by_location.empty:
            latest_stock_loc = stock_by_location.sort_values("Date").drop_duplicates(
                subset=["Location", "Fuel Type"], keep="last"
            )
            stock_loc_summary = latest_stock_loc.groupby(["Location", "Fuel Type"])["Closing Stock"].sum().reset_index()
            fig_loc = px.bar(
                stock_loc_summary,
                x="Location",
                y="Closing Stock",
                color="Fuel Type",
                title="Current Stock by Location",
                barmode="group",
                color_discrete_sequence=CHART_COLORS,
            )
            apply_chart_theme(fig_loc, height=400, x_title="Location", y_title="Closing Stock (L)")
            fig_loc.update_layout(title_text="")
            with st.container(border=True, key="chart_location"):
                render_chart_title(st, "Current Stock by Location")
                st.plotly_chart(fig_loc, width='stretch')
        else:
            st.info("No location data available")
    
    # Chart 4: Resupply Schedule
    with col4:
        resupply_data = resupply_df.dropna(subset=["Date", "Quantity"])
        if not resupply_data.empty:
            resupply_summary = resupply_data.groupby(["Date", "Fuel Type"])["Quantity"].sum().reset_index()
            fig_resupply = px.bar(
                resupply_summary,
                x="Date",
                y="Quantity",
                color="Fuel Type",
                title="Scheduled Resupply by Date",
                barmode="stack",
                color_discrete_sequence=CHART_COLORS,
            )
            apply_chart_theme(fig_resupply, height=400, x_title="Resupply Date", y_title="Quantity (L)", date_x=True)
            fig_resupply.update_layout(title_text="")
            with st.container(border=True, key="chart_resupply"):
                render_chart_title(st, "Scheduled Resupply by Date")
                st.plotly_chart(fig_resupply, width='stretch')
        else:
            st.info("No resupply data scheduled")

elif active_section == "📦 Terminal Data":
    st.subheader("Terminal Capacities")
    
    # Terminal data filters
    t_locations = sorted([x for x in terminal_df["Location"].dropna().unique()])
    t_info = sorted([x for x in terminal_df["Terminal Info"].dropna().unique()])

    t_col1, t_col2 = st.columns(2)
    t_loc_box = t_col1.container(border=True, key="terminal_location_filter_box")
    t_info_box = t_col2.container(border=True, key="terminal_type_filter_box")
    t_loc_sel = checkbox_slicer_horizontal(t_loc_box, "Terminal Location", t_locations, "t_loc")
    t_info_sel = checkbox_slicer_horizontal(t_info_box, "Terminal Type", t_info, "t_info")
    
    filtered_terminal = terminal_df[
        terminal_df["Company"].isin(company_sel)
        & terminal_df["Location"].isin(t_loc_sel)
        & terminal_df["Terminal Info"].isin(t_info_sel)
    ].copy()
    
    # Terminal visualization
    valid_terminal = filtered_terminal.dropna(subset=["Quantity"])
    if not valid_terminal.empty:
        fig_terminal = px.bar(
            valid_terminal,
            x="Terminal Info",
            y="Quantity",
            color="Fuel Type",
            facet_col="Location",
            title="Terminal Capacities by Location and Type",
            barmode="group",
            color_discrete_sequence=CHART_COLORS,
        )
        apply_chart_theme(fig_terminal, height=450, x_title="Terminal Category", y_title="Capacity (L)")
        fig_terminal.update_layout(title_text="")
        fig_terminal.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
        with st.container(border=True, key="chart_terminal"):
            render_chart_title(st, "Terminal Capacities by Location and Type")
            st.plotly_chart(fig_terminal, width='stretch')
    else:
        st.info("No terminal data available")

    # Company-level comparison of current stock against available terminal capacity.
    stock_company_base = filtered_actual.dropna(subset=["Company", "Date", "Closing Stock"])
    terminal_company_base = filtered_terminal.dropna(subset=["Company", "Quantity"])

    stock_by_company = pd.DataFrame(columns=["Company", "Current Stock (L)"])
    if not stock_company_base.empty:
        latest_stock_company = stock_company_base.sort_values("Date").drop_duplicates(
            subset=["Company", "Location", "Fuel Type"], keep="last"
        )
        stock_by_company = (
            latest_stock_company.groupby("Company", as_index=False)["Closing Stock"]
            .sum()
            .rename(columns={"Closing Stock": "Current Stock (L)"})
        )

    capacity_by_company = pd.DataFrame(columns=["Company", "Terminal Capacity (L)"])
    if not terminal_company_base.empty:
        capacity_by_company = (
            terminal_company_base.groupby("Company", as_index=False)["Quantity"]
            .sum()
            .rename(columns={"Quantity": "Terminal Capacity (L)"})
        )

    company_compare = pd.merge(stock_by_company, capacity_by_company, on="Company", how="outer").fillna(0)

    if not company_compare.empty:
        company_compare["Utilization (%)"] = np.where(
            company_compare["Terminal Capacity (L)"] > 0,
            (company_compare["Current Stock (L)"] / company_compare["Terminal Capacity (L)"]) * 100,
            np.nan,
        )
        company_compare = company_compare.sort_values("Terminal Capacity (L)", ascending=False)

        # Battery-style view: show how full each company's storage is.
        battery_df = company_compare[["Company", "Utilization (%)"]].copy()
        battery_df["Utilization (%)"] = battery_df["Utilization (%)"].clip(lower=0, upper=100)
        battery_df["Remaining (%)"] = 100 - battery_df["Utilization (%)"]
        battery_df["Status"] = np.select(
            [
                battery_df["Utilization (%)"] < 30,
                battery_df["Utilization (%)"].between(30, 60, inclusive="left"),
            ],
            ["Low", "Medium"],
            default="High",
        )

        fig_battery = go.Figure()
        fig_battery.add_trace(
            go.Bar(
                y=battery_df["Company"],
                x=battery_df["Utilization (%)"],
                orientation="h",
                name="Filled",
                marker=dict(
                    color=np.where(
                        battery_df["Status"] == "Low",
                        "#EF4444",
                        np.where(battery_df["Status"] == "Medium", "#F59E0B", "#22C55E"),
                    ),
                ),
                text=[f"{v:.1f}%" for v in battery_df["Utilization (%)"]],
                textposition="inside",
                insidetextanchor="middle",
                hovertemplate="Company=%{y}<br>Fill=%{x:.1f}%<extra></extra>",
            )
        )
        fig_battery.add_trace(
            go.Bar(
                y=battery_df["Company"],
                x=battery_df["Remaining (%)"],
                orientation="h",
                name="Remaining",
                marker=dict(color="rgba(148, 163, 184, 0.35)"),
                hovertemplate="Company=%{y}<br>Remaining=%{x:.1f}%<extra></extra>",
            )
        )
        apply_chart_theme(fig_battery, height=350, x_title="Terminal Fill Level (%)", y_title="Company")
        fig_battery.update_layout(
            barmode="stack",
            title_text="",
            xaxis=dict(range=[0, 100], ticksuffix="%"),
            legend_title_text="",
        )

        with st.container(border=True, key="chart_battery_stock_capacity"):
            render_chart_title(st, "Terminal Fill Level by Company")
            st.plotly_chart(fig_battery, width="stretch")
    else:
        st.info("No company data available for stock vs capacity comparison")

    # Location-level comparison and battery/discharge view.
    stock_location_base = filtered_actual.dropna(subset=["Location", "Date", "Closing Stock"])
    terminal_location_base = filtered_terminal.dropna(subset=["Location", "Quantity"])

    stock_by_location = pd.DataFrame(columns=["Location", "Current Stock (L)"])
    if not stock_location_base.empty:
        latest_stock_location = stock_location_base.sort_values("Date").drop_duplicates(
            subset=["Company", "Location", "Fuel Type"], keep="last"
        )
        stock_by_location = (
            latest_stock_location.groupby("Location", as_index=False)["Closing Stock"]
            .sum()
            .rename(columns={"Closing Stock": "Current Stock (L)"})
        )

    capacity_by_location = pd.DataFrame(columns=["Location", "Terminal Capacity (L)"])
    if not terminal_location_base.empty:
        capacity_by_location = (
            terminal_location_base.groupby("Location", as_index=False)["Quantity"]
            .sum()
            .rename(columns={"Quantity": "Terminal Capacity (L)"})
        )

    location_compare = pd.merge(stock_by_location, capacity_by_location, on="Location", how="outer").fillna(0)

    if not location_compare.empty:
        location_compare["Utilization (%)"] = np.where(
            location_compare["Terminal Capacity (L)"] > 0,
            (location_compare["Current Stock (L)"] / location_compare["Terminal Capacity (L)"]) * 100,
            np.nan,
        )
        location_compare = location_compare.sort_values("Terminal Capacity (L)", ascending=False)

        battery_loc_df = location_compare[["Location", "Utilization (%)"]].copy()
        battery_loc_df["Utilization (%)"] = battery_loc_df["Utilization (%)"].clip(lower=0, upper=100)
        battery_loc_df["Remaining (%)"] = 100 - battery_loc_df["Utilization (%)"]
        battery_loc_df["Status"] = np.select(
            [
                battery_loc_df["Utilization (%)"] < 30,
                battery_loc_df["Utilization (%)"].between(30, 60, inclusive="left"),
            ],
            ["Low", "Medium"],
            default="High",
        )

        fig_battery_location = go.Figure()
        fig_battery_location.add_trace(
            go.Bar(
                y=battery_loc_df["Location"],
                x=battery_loc_df["Utilization (%)"],
                orientation="h",
                name="Filled",
                marker=dict(
                    color=np.where(
                        battery_loc_df["Status"] == "Low",
                        "#EF4444",
                        np.where(battery_loc_df["Status"] == "Medium", "#F59E0B", "#22C55E"),
                    ),
                ),
                text=[f"{v:.1f}%" for v in battery_loc_df["Utilization (%)"]],
                textposition="inside",
                insidetextanchor="middle",
                hovertemplate="Location=%{y}<br>Fill=%{x:.1f}%<extra></extra>",
            )
        )
        fig_battery_location.add_trace(
            go.Bar(
                y=battery_loc_df["Location"],
                x=battery_loc_df["Remaining (%)"],
                orientation="h",
                name="Remaining",
                marker=dict(color="rgba(148, 163, 184, 0.35)"),
                hovertemplate="Location=%{y}<br>Remaining=%{x:.1f}%<extra></extra>",
            )
        )
        apply_chart_theme(fig_battery_location, height=320, x_title="Terminal Fill Level (%)", y_title="Location")
        fig_battery_location.update_layout(
            barmode="stack",
            title_text="",
            xaxis=dict(range=[0, 100], ticksuffix="%"),
            legend_title_text="",
        )

        with st.container(border=True, key="chart_battery_stock_capacity_location"):
            render_chart_title(st, "Terminal Fill Level by Location")
            st.plotly_chart(fig_battery_location, width="stretch")
    else:
        st.info("No location data available for stock vs capacity comparison")

elif active_section == "💰 Prices & Tariffs":
    if price_df is None or tariff_df is None:
        st.warning(
            "Price & tariff data file not found. "
            "Place Transformed_for_Analysis.xlsx in the project folder."
        )
    else:
        # ── Fuel Prices ───────────────────────────────────────────────
        st.subheader("Fuel Prices")

        latest_price_base = price_df.dropna(subset=["Date", "Price"]).sort_values("Date")
        latest_retail_prices = (
            latest_price_base[latest_price_base["Price_Type"] == "Retail"]
            .drop_duplicates(subset=["Fuel"], keep="last")
        )
        latest_wholesale_prices = (
            latest_price_base[latest_price_base["Price_Type"] == "Wholesale"]
            .drop_duplicates(subset=["Fuel"], keep="last")
        )

        wholesale_col, retail_col = st.columns(2)

        with wholesale_col:
            if not latest_wholesale_prices.empty:
                wholesale_kpi_items = [
                    (
                        fuel_icon(row["Fuel"]),
                        row["Fuel"],
                        f"T${row['Price']:.2f}",
                        CHART_COLORS[i % len(CHART_COLORS)],
                    )
                    for i, (_, row) in enumerate(latest_wholesale_prices.iterrows())
                ]
                render_kpi_group(st, "Latest Wholesale Prices (T$/L)", wholesale_kpi_items)
            else:
                st.info("No wholesale price data available")

        with retail_col:
            if not latest_retail_prices.empty:
                retail_kpi_items = [
                    (
                        fuel_icon(row["Fuel"]),
                        row["Fuel"],
                        f"T${row['Price']:.2f}",
                        CHART_COLORS[i % len(CHART_COLORS)],
                    )
                    for i, (_, row) in enumerate(latest_retail_prices.iterrows())
                ]
                render_kpi_group(st, "Latest Retail Prices (T$/L)", retail_kpi_items)
            else:
                st.info("No retail price data available")

        st.markdown("<div style='height: 0.3rem;'></div>", unsafe_allow_html=True)

        fp_filtered = price_df.copy()
        if price_type_sel:
            fp_filtered = fp_filtered[fp_filtered["Price_Type"].isin(price_type_sel)]
        if fuel_sel_fp:
            fp_filtered = fp_filtered[fp_filtered["Fuel"].isin(fuel_sel_fp)]
        fp_filtered = fp_filtered.dropna(subset=["Date", "Price"]).sort_values("Date")

        if not fp_filtered.empty:
            fig_fp = px.line(
                fp_filtered,
                x="Date",
                y="Price",
                color="Fuel",
                line_dash="Price_Type" if len(price_type_sel) > 1 else None,
                color_discrete_sequence=CHART_COLORS,
            )
            apply_chart_theme(
                fig_fp,
                height=400,
                hovermode="x unified",
                x_title="Date",
                y_title="Price (T$/L)",
                date_x=True,
            )
            fig_fp.update_layout(title_text="")
            with st.container(border=True, key="chart_fuel_price"):
                render_chart_title(st, "Fuel Price Trends Over Time")
                st.plotly_chart(fig_fp, width="stretch")
        else:
            st.info("No fuel price data matches the selected filters")

        st.divider()

        # ── Electricity Tariffs ───────────────────────────────────────
        st.subheader("Electricity Tariffs")

        _MONTH_NUM = {
            "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
        }

        def _parse_tariff_date(row):
            try:
                yr_start = int(str(row["Year"]).split("/")[0])
                m = _MONTH_NUM.get(str(row["Month"]), 1)
                yr = yr_start if m >= 7 else yr_start + 1
                return pd.Timestamp(year=yr, month=m, day=1)
            except Exception:
                return pd.NaT

        tr_work = tariff_df.copy()
        tr_work["Period"] = tr_work.apply(_parse_tariff_date, axis=1)

        main_components = ["Fuel Component", "Non Fuel component", "Total Tariff"]
        tr_main = tr_work[tr_work["Component"].isin(main_components)].dropna(
            subset=["Period", "Value"]
        )

        latest_year = tr_work["Year"].dropna().iloc[-1] if not tr_work.empty else ""
        tr_latest_yr = tr_main[tr_main["Year"] == latest_year]
        if not tr_latest_yr.empty:
            current_tariff_col, average_tariff_col = st.columns(2)

            latest_period = tr_latest_yr["Period"].max()
            tr_current = tr_latest_yr[tr_latest_yr["Period"] == latest_period]
            current_vals = tr_current.groupby("Component")["Value"].mean()
            current_kpi_items = [
                (
                    "⛽" if c == "Fuel Component" else ("🔌" if "non fuel" in c.lower() else "📊"),
                    c,
                    f"T${v:.4f}/kWh",
                    CHART_COLORS[i % len(CHART_COLORS)],
                )
                for i, (c, v) in enumerate(current_vals.items())
            ]
            render_kpi_group(
                current_tariff_col,
                f"Current Tariff — {pd.Timestamp(latest_period).strftime('%b %Y')} (T$/kWh)",
                current_kpi_items,
            )

            avg_vals = tr_latest_yr.groupby("Component")["Value"].mean()
            average_kpi_items = [
                (
                    "⛽" if c == "Fuel Component" else ("🔌" if "non fuel" in c.lower() else "📊"),
                    c,
                    f"T${v:.4f}/kWh",
                    CHART_COLORS[i % len(CHART_COLORS)],
                )
                for i, (c, v) in enumerate(avg_vals.items())
            ]
            render_kpi_group(average_tariff_col, f"Average Tariff — {latest_year} (T$/kWh)", average_kpi_items)
            st.markdown("<div style='height: 0.3rem;'></div>", unsafe_allow_html=True)

        tr_filtered = (
            tr_main[tr_main["Component"].isin(comp_sel)].sort_values("Period")
            if comp_sel
            else tr_main.sort_values("Period")
        )

        if not tr_filtered.empty:
            fig_tr = px.line(
                tr_filtered,
                x="Period",
                y="Value",
                color="Component",
                color_discrete_sequence=CHART_COLORS,
            )
            apply_chart_theme(
                fig_tr,
                height=400,
                hovermode="x unified",
                x_title="Period",
                y_title="Tariff (T$/kWh)",
                date_x=True,
            )
            fig_tr.update_layout(title_text="")
            with st.container(border=True, key="chart_tariff"):
                render_chart_title(st, "Electricity Tariff Components Over Time")
                st.plotly_chart(fig_tr, width="stretch")
        else:
            st.info("No tariff data matches the selected filters")

        fp_heatmap = price_df.copy()
        if price_type_sel:
            fp_heatmap = fp_heatmap[fp_heatmap["Price_Type"].isin(price_type_sel)]
        if fuel_sel_fp:
            fp_heatmap = fp_heatmap[fp_heatmap["Fuel"].isin(fuel_sel_fp)]
        fp_heatmap = fp_heatmap.dropna(subset=["Date", "Price"])
        fp_heatmap["Period"] = fp_heatmap["Date"].dt.to_period("M").dt.to_timestamp()
        fp_heatmap["Series"] = fp_heatmap["Price_Type"] + " " + fp_heatmap["Fuel"]
        fp_matrix = fp_heatmap.pivot_table(
            index="Period",
            columns="Series",
            values="Price",
            aggfunc="mean",
        )

        tr_heatmap = tr_main.copy()
        if comp_sel:
            tr_heatmap = tr_heatmap[tr_heatmap["Component"].isin(comp_sel)]
        tr_matrix = tr_heatmap.pivot_table(
            index="Period",
            columns="Component",
            values="Value",
            aggfunc="mean",
        )

        dependency_frame = fp_matrix.join(tr_matrix, how="inner")
        dependency_corr = dependency_frame.corr(numeric_only=True)

        if not fp_matrix.empty and not tr_matrix.empty and not dependency_corr.empty and len(dependency_frame) >= 3:
            fuel_series = fp_matrix.columns.tolist()
            tariff_series = tr_matrix.columns.tolist()
            dependency_view = dependency_corr.loc[fuel_series, tariff_series]

            fig_dependency = go.Figure(
                data=go.Heatmap(
                    z=dependency_view.values,
                    x=dependency_view.columns.tolist(),
                    y=dependency_view.index.tolist(),
                    colorscale="RdBu",
                    zmin=-1,
                    zmax=1,
                    zmid=0,
                    text=np.round(dependency_view.values, 2),
                    texttemplate="%{text}",
                    textfont={"size": 12},
                    colorbar=dict(title="Correlation"),
                    hovertemplate="Fuel Price: %{y}<br>Tariff: %{x}<br>Correlation: %{z:.2f}<extra></extra>",
                )
            )
            fig_dependency.update_layout(
                height=460,
                paper_bgcolor="rgba(0, 0, 0, 0)",
                plot_bgcolor="rgba(12, 18, 26, 0.86)",
                font=dict(color="#E6EDF5", size=13),
                margin=dict(l=16, r=16, t=24, b=16),
                xaxis=dict(title="Tariff Component", side="bottom"),
                yaxis=dict(title="Fuel Price Series"),
            )
            with st.container(border=True, key="chart_dependency"):
                render_chart_title(st, "Fuel Price vs Tariff Dependence Heatmap")
                st.caption("Correlation across overlapping monthly periods. Values closer to 1 move together more strongly; values near -1 move in opposite directions.")
                st.plotly_chart(fig_dependency, width="stretch")
        else:
            st.info("Not enough overlapping monthly data to calculate a dependency heatmap")



st.divider()
pdf_col1, pdf_col2 = st.columns([1, 5])
with pdf_col1:
    if st.button("⬇️ Generate Summary PDF", key="generate_pdf_btn"):
        st.session_state.summary_pdf_bytes = build_summary_pdf_bytes()
    if st.session_state.summary_pdf_bytes:
        st.download_button(
            "⬇️ Download Summary PDF",
            data=st.session_state.summary_pdf_bytes,
            file_name=f"fuel_dashboard_summary_{pd.Timestamp.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            key="download_summary_pdf",
        )

with pdf_col2:
    if st.session_state.summary_pdf_bytes:
        st.caption("PDF ready — click Download to save.")
    else:
        st.caption("Click Generate to build the summary PDF.")

st.markdown(
    f"""
    <div style="
        margin-top: 0.6rem;
        padding: 0.52rem 0.75rem;
        background: rgba(229, 231, 235, 0.95);
        border: 1px solid rgba(209, 213, 219, 0.95);
        color: #4B5563;
        font-size: 0.88rem;
        line-height: 1.35;
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 0.7rem;
        flex-wrap: wrap;
        text-align: center;
        border-radius: 12px;
    ">
        {footer_logo_html}
        <span>Developed by Department of Energy under Ministry of MEIDECC for Staged Energy and Fuel Supply Plan Monitoring | Data updated: {last_sync}</span>
    </div>
    """,
    unsafe_allow_html=True,
)
