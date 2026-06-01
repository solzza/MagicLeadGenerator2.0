from pathlib import Path
from datetime import datetime
import json
import re

import pandas as pd
import streamlit as st

from lead_cleaner import (
    CleanOptions,
    available_sheets,
    clean_leads,
    dataframe_to_xlsx_bytes,
    has_duplicate_match,
    infer_column_mapping,
    mark_duplicates_against,
    read_table,
    scan_column_quality,
    to_import_columns,
    validate_raw_columns,
)


DOMAIN_REGISTRY_PATH = Path("Domaner.csv")
HISTORY_PATH = Path(".lead_history.json")
HISTORY_FILES_DIR = Path(".lead_history_files")
TABLE_HEIGHT = 650


st.set_page_config(
    page_title="Magic Leads Generator",
    page_icon="",
    layout="wide",
)

st.markdown(
    """
    <style>
    @media (prefers-color-scheme: light) {
    :root {
        --mlg-bg: #f7f8f5;
        --mlg-panel: #ffffff;
        --mlg-ink: #182024;
        --mlg-muted: #63717a;
        --mlg-line: #dfe5df;
        --mlg-accent: #b88a22;
        --mlg-accent-strong: #8f6816;
        --mlg-teal: #2d6f73;
        --mlg-teal-soft: #e8f2ef;
    }

    .stApp {
        background:
            linear-gradient(180deg, rgba(255, 255, 255, 0.9), rgba(247, 248, 245, 0.95) 260px),
            var(--mlg-bg);
        color: var(--mlg-ink);
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    h1 {
        color: var(--mlg-ink);
        font-size: 2.1rem;
        font-weight: 750;
        letter-spacing: 0;
        margin-bottom: 1.25rem;
    }

    section[data-testid="stSidebar"] {
        background: #eef3ef;
        border-right: 1px solid var(--mlg-line);
    }

    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: var(--mlg-ink);
        font-weight: 700;
        letter-spacing: 0;
    }

    div[data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.72);
        border: 1px solid var(--mlg-line);
        border-radius: 8px;
        padding: 0.6rem;
    }

    div[data-testid="stMetric"] {
        background: var(--mlg-panel);
        border: 1px solid var(--mlg-line);
        border-left: 4px solid var(--mlg-teal);
        border-radius: 8px;
        padding: 0.75rem 0.9rem;
        box-shadow: 0 6px 18px rgba(24, 32, 36, 0.05);
    }

    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
        color: var(--mlg-muted);
        font-weight: 650;
    }

    div[data-testid="stMetricValue"] {
        color: var(--mlg-ink);
        font-weight: 760;
    }

    div[data-testid="stTabs"] button[role="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.65rem 1rem;
        color: var(--mlg-muted);
        font-weight: 650;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background: var(--mlg-panel);
        color: var(--mlg-ink);
        border-bottom: 3px solid var(--mlg-accent);
    }

    div[data-testid="stDataFrame"],
    div[data-testid="stDataEditor"] {
        border: 1px solid var(--mlg-line);
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 10px 24px rgba(24, 32, 36, 0.06);
    }

    div[data-testid="stExpander"] {
        background: rgba(255, 255, 255, 0.74);
        border: 1px solid var(--mlg-line);
        border-radius: 8px;
    }

    div[data-testid="stInfo"],
    div[data-testid="stAlert"] {
        border-radius: 8px;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius: 8px;
        border: 1px solid rgba(143, 104, 22, 0.32);
        background: #ffffff;
        color: var(--mlg-ink);
        font-weight: 700;
        white-space: nowrap;
        box-shadow: 0 4px 12px rgba(24, 32, 36, 0.05);
    }

    .stButton > button p,
    .stDownloadButton > button p {
        white-space: nowrap;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: var(--mlg-accent);
        color: var(--mlg-accent-strong);
        background: #fffaf0;
    }

    button[kind="primary"] {
        border-radius: 8px;
        border: 1px solid #245f62;
        background: #2d6f73;
        color: #ffffff;
        font-weight: 750;
        box-shadow: 0 6px 14px rgba(45, 111, 115, 0.18);
    }

    button[kind="primary"]:hover {
        border-color: #1d5053;
        background: #245f62;
        color: #ffffff;
    }

    .stDownloadButton > button {
        background: #fff5dc;
        border-color: rgba(184, 138, 34, 0.55);
        color: #6f5011;
    }

    .stDownloadButton > button:hover {
        background: #ffedbd;
        border-color: var(--mlg-accent);
        color: #5f430b;
    }

    .stSlider [data-baseweb="slider"] > div {
        color: var(--mlg-teal);
    }

    textarea,
    input {
        border-radius: 8px !important;
    }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_default_registry(path: str) -> pd.DataFrame:
    return read_table(path)


def load_history() -> list[dict[str, object]]:
    if not HISTORY_PATH.exists():
        return []
    try:
        data = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return data if isinstance(data, list) else []


def save_history(history: list[dict[str, object]]) -> None:
    HISTORY_PATH.write_text(
        json.dumps(history[:30], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def append_history(entry: dict[str, object]) -> None:
    history = load_history()
    entry = {"time": datetime.now().strftime("%Y-%m-%d %H:%M"), **entry}
    key = entry.get("key")
    if key:
        history = [item for item in history if item.get("key") != key]
    history.insert(0, entry)
    save_history(history)


def reset_review_state() -> None:
    st.session_state.reviewed_rows = set()
    st.session_state.excluded_rows = set()
    st.session_state.review_mark_all = False
    st.session_state.review_editor_version = st.session_state.get("review_editor_version", 0) + 1


def safe_history_filename(name: str) -> str:
    path = Path(name)
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", path.stem).strip("._-") or "rawdata"
    return f"{stem}{path.suffix.lower()}"


def save_uploaded_raw_file(uploaded_file) -> Path:
    HISTORY_FILES_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = HISTORY_FILES_DIR / f"{timestamp}_{safe_history_filename(uploaded_file.name)}"
    output_path.write_bytes(uploaded_file.getvalue())
    return output_path


def render_history() -> None:
    history = load_history()
    with st.sidebar.expander("Historik", expanded=False):
        if not history:
            st.caption("Ingen historik ännu.")
            return
        for index, item in enumerate(history):
            file_name = str(item.get("name", "Okänd fil"))
            item_type = str(item.get("type", ""))
            rows = item.get("rows", "")
            time = str(item.get("time", ""))
            label = f"{item_type}: {file_name}"
            st.caption(f"{time} | {rows} rader")
            if item_type == "Rådata" and item.get("path") and Path(str(item["path"])).exists():
                if st.button(label, key=f"open_history_{index}", use_container_width=True):
                    st.session_state.history_raw_path = str(item["path"])
                    st.session_state.history_raw_name = file_name
                    reset_review_state()
                    st.session_state.clean_input_key = None
            else:
                st.caption(label)


def issue_badge_count(df: pd.DataFrame, issue: str) -> int:
    if df.empty or "Issues" not in df:
        return 0
    return int(df["Issues"].fillna("").str.contains(issue, regex=False).sum())


def highlight_duplicate_rows(row: pd.Series) -> list[str]:
    if not has_duplicate_match(row.get("Duplicate Match", "")):
        return ["" for _ in row]
    return [
        "border-top: 2px solid #c08a00; "
        "border-bottom: 2px solid #c08a00; "
        "box-shadow: inset 0 0 8px rgba(192, 138, 0, 0.45); "
        "background-color: rgba(255, 214, 102, 0.18)"
        for _ in row
    ]


def apply_review_edits(edited_review: pd.DataFrame, original_review_rows: set[int]) -> tuple[pd.DataFrame, set[int]]:
    edited_review_rows = set(edited_review["Rad"].dropna().astype(int).tolist()) if "Rad" in edited_review.columns else set()
    removed_rows = original_review_rows - edited_review_rows

    for row_id in removed_rows:
        st.session_state.excluded_rows.add(row_id)

    edited_not_removed = edited_review[edited_review["Rad"].notna()] if "Rad" in edited_review.columns else edited_review
    saved_working_df = st.session_state.working_df.copy()
    for _, row in edited_not_removed.iterrows():
        row_id = int(row["Rad"])
        values = row.drop(labels=["Klar"], errors="ignore").to_dict()
        row_matches = saved_working_df.index[saved_working_df["Rad"] == row_id].tolist()
        if row_matches:
            row_index = row_matches[0]
            for column, value in values.items():
                if column in saved_working_df.columns:
                    saved_working_df.at[row_index, column] = value
    st.session_state.working_df = saved_working_df
    return edited_not_removed, removed_rows


st.title("Magic Leads Generator")

with st.sidebar:
    st.header("Input")
    raw_file = st.file_uploader("Evaboot-rådata", type=["xlsx", "xlsm", "xls", "csv"])
    if DOMAIN_REGISTRY_PATH.exists():
        registry_file = None
        st.caption("Domänregister: Domaner.csv")
    else:
        registry_file = st.file_uploader(
            "Bolag/domän-register",
            type=["xlsx", "xlsm", "xls", "csv"],
            help="Behövs i deployment om Domaner.csv inte finns med i GitHub-repot.",
        )

    st.header("Inställningar")
    fuzzy_threshold = st.slider("Fuzzy match-gräns", min_value=0.80, max_value=1.00, value=0.92, step=0.01)
    title_keywords_text = st.text_area(
        "Exkludera titlar med ord",
        value="student, assistant, intern, junior, pensionär",
        help="Prospect Position betyder personens jobbtitel/arbetsroll. Kommaseparerad lista. Matchande rader flaggas, men tas inte bort automatiskt.",
    )
    render_history()

history_raw_path = st.session_state.get("history_raw_path")
if raw_file is not None:
    raw_source = raw_file
    raw_source_name = raw_file.name
elif history_raw_path and Path(history_raw_path).exists():
    raw_source = Path(history_raw_path)
    raw_source_name = st.session_state.get("history_raw_name", raw_source.name)
    st.info(f"Öppnad från historik: {raw_source_name}")
else:
    raw_source = None
    raw_source_name = ""

if raw_source is None:
    st.info("Ladda upp en Evaboot-fil för att börja.")
    st.stop()

try:
    raw_sheets = available_sheets(raw_source)
    raw_sheet = st.selectbox("Rådataflik", raw_sheets, index=raw_sheets.index("Rådata") if "Rådata" in raw_sheets else 0) if raw_sheets else 0
    raw_df = read_table(raw_source, raw_sheet)
except Exception as exc:
    st.error(f"Kunde inte läsa rådata: {exc}")
    st.stop()

missing = validate_raw_columns(raw_df)
if missing:
    st.error("Rådata saknar obligatoriska kolumner: " + ", ".join(missing))
    st.stop()

raw_history_key = f"raw:{raw_source_name}:{getattr(raw_file, 'size', '') if raw_file is not None else history_raw_path}"
if raw_file is not None and st.session_state.get("last_raw_history_key") != raw_history_key:
    saved_raw_path = save_uploaded_raw_file(raw_file)
    append_history(
        {
            "key": raw_history_key,
            "type": "Rådata",
            "name": raw_source_name,
            "rows": len(raw_df),
            "sheet": raw_sheet if isinstance(raw_sheet, str) else "",
            "path": str(saved_raw_path),
        }
    )
    st.session_state.last_raw_history_key = raw_history_key

with st.expander("Identifierade kolumner", expanded=False):
    mapping = infer_column_mapping(raw_df)
    mapping_df = scan_column_quality(raw_df, mapping)
    st.dataframe(mapping_df, use_container_width=True, hide_index=True)
    st.caption("Prospect Position = personens titel i sin arbetsroll, till exempel Product Manager eller Head of Sales.")
    warnings = mapping_df[mapping_df["Varning"] != ""]
    if not warnings.empty:
        st.warning("Kontrollera kolumnmappningen: " + "; ".join(warnings["Standardkolumn"] + " - " + warnings["Varning"]))

try:
    if DOMAIN_REGISTRY_PATH.exists():
        registry_df = load_default_registry(str(DOMAIN_REGISTRY_PATH))
        st.caption("Använder Domaner.csv som bolags-/domänregister.")
    elif registry_file is not None:
        registry_sheets = available_sheets(registry_file)
        registry_sheet = st.selectbox(
            "Registerflik",
            registry_sheets,
            index=registry_sheets.index("Domaner") if "Domaner" in registry_sheets else 0,
        ) if registry_sheets else 0
        registry_df = read_table(registry_file, registry_sheet)
        st.caption(f"Använder uppladdat bolags-/domänregister: {registry_file.name}")
    else:
        st.warning("Domaner.csv hittades inte. Ladda upp bolag/domän-registret i sidopanelen.")
        registry_df = pd.DataFrame(columns=["Bolag", "Uppdaterad domän"])
except Exception as exc:
    st.error(f"Kunde inte läsa bolagsregistret: {exc}")
    st.stop()

keywords = tuple(part.strip() for part in title_keywords_text.split(",") if part.strip())
options = CleanOptions(
    fuzzy_threshold=fuzzy_threshold,
    remove_title_keywords=keywords,
)

if "reviewed_rows" not in st.session_state:
    st.session_state.reviewed_rows = set()
if "excluded_rows" not in st.session_state:
    st.session_state.excluded_rows = set()
if "review_mark_all" not in st.session_state:
    st.session_state.review_mark_all = False
if "review_editor_version" not in st.session_state:
    st.session_state.review_editor_version = 0
if "clean_input_key" not in st.session_state:
    st.session_state.clean_input_key = None
if "working_df" not in st.session_state:
    st.session_state.working_df = pd.DataFrame()

registry_key = (
    str(DOMAIN_REGISTRY_PATH.resolve()) if DOMAIN_REGISTRY_PATH.exists()
    else f"{getattr(registry_file, 'name', '')}:{getattr(registry_file, 'size', '')}:{registry_sheet if 'registry_sheet' in locals() else ''}"
)
clean_input_key = (
    raw_source_name,
    getattr(raw_file, "size", "") if raw_file is not None else str(history_raw_path),
    raw_sheet if isinstance(raw_sheet, str) else "",
    registry_key,
    fuzzy_threshold,
    keywords,
)
if st.session_state.clean_input_key != clean_input_key:
    cleaned_df = clean_leads(raw_df, registry_df, options).reset_index(drop=True)
    cleaned_df.insert(0, "Rad", cleaned_df.index + 1)
    st.session_state.working_df = cleaned_df
    st.session_state.clean_input_key = clean_input_key
    reset_review_state()

working_df = st.session_state.working_df.copy()

total = len(working_df)
auto_ready_mask = working_df["Issues"].fillna("") == ""
manual_ready_mask = working_df["Rad"].isin(st.session_state.reviewed_rows)
excluded_mask = working_df["Rad"].isin(st.session_state.excluded_rows)
ready_mask = (auto_ready_mask | manual_ready_mask) & ~excluded_mask
review_mask = ~(auto_ready_mask | manual_ready_mask | excluded_mask)
ready = int(ready_mask.sum())
needs_review = int(review_mask.sum())

summary_cols = st.columns(5)
summary_cols[0].metric("Leads", total)
summary_cols[1].metric("Klara", ready)
summary_cols[2].metric("Behöver granskas", needs_review)
summary_cols[3].metric("Nya bolag", issue_badge_count(working_df, "New company"))
summary_cols[4].metric("E-post saknas", issue_badge_count(working_df, "email_missing"))

quality_cols = st.columns(3)
quality_cols[0].metric("Titel saknas/fel", issue_badge_count(working_df, "title_missing") + issue_badge_count(working_df, "title_invalid"))
quality_cols[1].metric("LinkedIn saknas/fel", issue_badge_count(working_df, "linkedin_missing") + issue_badge_count(working_df, "linkedin_invalid"))
quality_cols[2].metric("Dubletter", int(working_df["Duplicate Key"].fillna("").ne("").sum()))

tab_all, tab_review, tab_ready, tab_excluded, tab_export = st.tabs(["Alla rader", "Granska", "Klara", "Exkluderade", "Export"])

with tab_all:
    edited_all = st.data_editor(
        working_df,
        use_container_width=True,
        hide_index=True,
        height=TABLE_HEIGHT,
        num_rows="dynamic",
        key="all_editor",
    )

with tab_review:
    review_df = working_df[review_mask].copy()
    review_df.insert(0, "Klar", st.session_state.review_mark_all)
    if st.session_state.review_mark_all:
        st.info("Bulkmarkering är aktiv: alla kvarvarande rader i Granska markeras som klara när du klickar Verkställ ändringar.")
    edited_review = st.data_editor(
        review_df,
        use_container_width=True,
        hide_index=True,
        height=TABLE_HEIGHT,
        num_rows="dynamic",
        key=f"review_editor_{st.session_state.review_editor_version}",
    )
    original_review_rows = set(review_df["Rad"].astype(int).tolist())
    edited_not_removed, removed_rows = apply_review_edits(edited_review, original_review_rows)
    if removed_rows:
        st.session_state.review_editor_version += 1
        st.rerun()

    review_button_cols = st.columns([1.8, 1.35, 1.35, 1.5])
    with review_button_cols[0]:
        submitted_review = st.button("Verkställ ändringar", type="primary", use_container_width=True)
    with review_button_cols[1]:
        submitted_mark_all = st.button("Markera alla", use_container_width=True)
    with review_button_cols[2]:
        submitted_clear_all = st.button("Avmarkera alla", use_container_width=True)

    if submitted_review:
        if st.session_state.review_mark_all:
            selected = edited_not_removed
        else:
            selected = edited_not_removed[edited_not_removed["Klar"] == True]
        for _, row in selected.iterrows():
            row_id = int(row["Rad"])
            st.session_state.reviewed_rows.add(row_id)
        st.session_state.review_mark_all = False
        st.session_state.review_editor_version += 1
        st.rerun()
    if submitted_mark_all:
        st.session_state.review_mark_all = True
        st.session_state.review_editor_version += 1
        st.rerun()
    if submitted_clear_all:
        st.session_state.review_mark_all = False
        st.session_state.review_editor_version += 1
        st.rerun()
    st.caption("Granska är ett sidospår. Rader här går inte till Export förrän du bockar i Klar och klickar på knappen.")

with tab_ready:
    ready_df = working_df[ready_mask]
    st.dataframe(ready_df, use_container_width=True, hide_index=True, height=TABLE_HEIGHT)

with tab_excluded:
    excluded_df = working_df[excluded_mask]
    st.dataframe(excluded_df, use_container_width=True, hide_index=True, height=TABLE_HEIGHT)
    if not excluded_df.empty and st.button("Återställ alla exkluderade"):
        st.session_state.excluded_rows = set()
        st.rerun()

with tab_export:
    export_df = working_df[ready_mask]
    import_df = to_import_columns(export_df)

    duplicate_file = st.session_state.get("duplicate_file")
    duplicate_sheets = []
    if duplicate_file is not None:
        try:
            duplicate_sheets = available_sheets(duplicate_file)
            selected_duplicate_sheet = st.session_state.get("duplicate_sheet")
            if selected_duplicate_sheet not in duplicate_sheets:
                selected_duplicate_sheet = duplicate_sheets[0] if duplicate_sheets else 0
            duplicate_sheet = selected_duplicate_sheet
            duplicate_df = read_table(duplicate_file, duplicate_sheet)
            import_df, duplicate_mapping = mark_duplicates_against(import_df, duplicate_df)
            duplicate_count = int(import_df["Duplicate Match"].map(has_duplicate_match).sum())
            st.info(
                f"Hittade {duplicate_count} möjliga dubletter. "
                f"Matchade kolumner: {duplicate_mapping if duplicate_mapping else 'inga'}"
            )
        except Exception as exc:
            st.error(f"Kunde inte köra dublettkontroll: {exc}")

    edited_import_df = st.data_editor(
        import_df,
        use_container_width=True,
        hide_index=True,
        height=TABLE_HEIGHT,
        key="export_editor",
    )

    with st.expander("Dublettkontroll", expanded=False):
        st.caption("Valfritt sista steg. Ladda upp en befintlig lista om du vill markera matchningar på namn och/eller e-post.")
        st.file_uploader(
            "Befintlig lista",
            type=["xlsx", "xlsm", "xls", "csv"],
            key="duplicate_file",
        )
        if duplicate_sheets:
            selected_duplicate_sheet = st.session_state.get("duplicate_sheet")
            if selected_duplicate_sheet not in duplicate_sheets:
                selected_duplicate_sheet = duplicate_sheets[0]
            st.selectbox(
                "Dublettfilens flik",
                duplicate_sheets,
                index=duplicate_sheets.index(selected_duplicate_sheet),
                key="duplicate_sheet",
            )

    excel_bytes = dataframe_to_xlsx_bytes(edited_import_df)
    if st.button("Exportera till Excel", type="primary"):
        output_dir = Path.home() / "Downloads"
        output_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_dir / f"upsales_import_{timestamp}.xlsx"
        output_path.write_bytes(excel_bytes)
        append_history(
            {
                "key": f"export:{output_path}",
                "type": "Export",
                "name": output_path.name,
                "rows": len(edited_import_df),
                "sheet": "Upsales Import",
                "path": str(output_path),
            }
        )
        st.success(f"Sparad: {output_path.resolve()}")
    st.download_button(
        "Ladda ner CSV",
        data=edited_import_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="upsales_import.csv",
        mime="text/csv",
    )
