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
SELECT_COLUMN = "Välj"


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

    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] * {
        color: var(--mlg-ink) !important;
    }

    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: var(--mlg-ink);
        font-weight: 700;
        letter-spacing: 0;
    }

    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: var(--mlg-ink) !important;
    }

    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: var(--mlg-muted) !important;
    }

    div[data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.72);
        border: 1px solid var(--mlg-line);
        border-radius: 8px;
        padding: 0.6rem;
    }

    section[data-testid="stSidebar"] div[data-testid="stFileUploader"],
    section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] {
        background: #ffffff !important;
        border-color: var(--mlg-line) !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] button {
        background: #f8faf7 !important;
        border: 1px solid var(--mlg-line) !important;
        color: var(--mlg-ink) !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] svg {
        color: var(--mlg-ink) !important;
        fill: var(--mlg-ink) !important;
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

    section[data-testid="stSidebar"] [data-baseweb="slider"] div {
        color: var(--mlg-teal) !important;
    }

    section[data-testid="stSidebar"] [data-baseweb="input"],
    section[data-testid="stSidebar"] [data-baseweb="textarea"],
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] input {
        background: #ffffff !important;
        border-color: var(--mlg-line) !important;
        color: var(--mlg-ink) !important;
    }

    section[data-testid="stSidebar"] textarea::placeholder,
    section[data-testid="stSidebar"] input::placeholder {
        color: var(--mlg-muted) !important;
    }

    textarea,
    input {
        border-radius: 8px !important;
    }
    }

    @media (prefers-color-scheme: dark) {
    :root {
        --mlg-dark-bg: #0e1117;
        --mlg-dark-panel: #161b22;
        --mlg-dark-panel-soft: #1b222c;
        --mlg-dark-ink: #fafafa;
        --mlg-dark-muted: #a3aab7;
        --mlg-dark-line: #30363d;
        --mlg-dark-accent: #d8be6a;
        --mlg-dark-teal: #5db9bd;
    }

    .stApp {
        background: var(--mlg-dark-bg);
        color: var(--mlg-dark-ink);
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    h1 {
        color: var(--mlg-dark-ink);
        font-size: 2.1rem;
        font-weight: 750;
        letter-spacing: 0;
        margin-bottom: 1.25rem;
    }

    section[data-testid="stSidebar"] {
        background: #11161d;
        border-right: 1px solid var(--mlg-dark-line);
    }

    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] * {
        color: var(--mlg-dark-ink) !important;
    }

    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: var(--mlg-dark-ink);
        font-weight: 700;
        letter-spacing: 0;
    }

    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: var(--mlg-dark-ink) !important;
    }

    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: var(--mlg-dark-muted) !important;
    }

    div[data-testid="stFileUploader"] {
        background: rgba(22, 27, 34, 0.78);
        border: 1px solid var(--mlg-dark-line);
        border-radius: 8px;
        padding: 0.6rem;
    }

    section[data-testid="stSidebar"] div[data-testid="stFileUploader"],
    section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] {
        background: #161b22 !important;
        border-color: var(--mlg-dark-line) !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] button {
        background: #1b222c !important;
        border: 1px solid var(--mlg-dark-line) !important;
        color: var(--mlg-dark-ink) !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stFileUploaderDropzone"] svg {
        color: var(--mlg-dark-ink) !important;
        fill: var(--mlg-dark-ink) !important;
    }

    div[data-testid="stMetric"] {
        background: var(--mlg-dark-panel);
        border: 1px solid var(--mlg-dark-line);
        border-left: 4px solid var(--mlg-dark-teal);
        border-radius: 8px;
        padding: 0.75rem 0.9rem;
        box-shadow: 0 8px 22px rgba(0, 0, 0, 0.22);
    }

    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
        color: var(--mlg-dark-muted);
        font-weight: 650;
    }

    div[data-testid="stMetricValue"] {
        color: var(--mlg-dark-ink);
        font-weight: 760;
    }

    div[data-testid="stTabs"] button[role="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.65rem 1rem;
        color: var(--mlg-dark-muted);
        font-weight: 650;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background: var(--mlg-dark-panel);
        color: var(--mlg-dark-ink);
        border-bottom: 3px solid var(--mlg-dark-accent);
    }

    div[data-testid="stDataFrame"],
    div[data-testid="stDataEditor"] {
        border: 1px solid var(--mlg-dark-line);
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 10px 26px rgba(0, 0, 0, 0.28);
    }

    div[data-testid="stExpander"] {
        background: rgba(22, 27, 34, 0.8);
        border: 1px solid var(--mlg-dark-line);
        border-radius: 8px;
    }

    div[data-testid="stInfo"],
    div[data-testid="stAlert"] {
        border-radius: 8px;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius: 8px;
        border: 1px solid #30363d;
        background: var(--mlg-dark-panel-soft);
        color: var(--mlg-dark-ink);
        font-weight: 700;
        white-space: nowrap;
        box-shadow: 0 5px 14px rgba(0, 0, 0, 0.22);
    }

    .stButton > button p,
    .stDownloadButton > button p {
        white-space: nowrap;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: var(--mlg-dark-accent);
        color: #fff7cf;
        background: #202733;
    }

    button[kind="primary"] {
        border-radius: 8px;
        border: 1px solid #5db9bd;
        background: #2d6f73;
        color: #ffffff;
        font-weight: 750;
        box-shadow: 0 6px 16px rgba(93, 185, 189, 0.2);
    }

    button[kind="primary"]:hover {
        border-color: #8bd7da;
        background: #245f62;
        color: #ffffff;
    }

    .magic-loader {
        border-color: rgba(216, 190, 106, 0.34);
        background: rgba(22, 27, 34, 0.82);
    }

    .magic-loader-text strong {
        color: var(--mlg-dark-ink);
    }

    .magic-loader-text span {
        color: var(--mlg-dark-muted);
    }

    section[data-testid="stSidebar"] [data-baseweb="slider"] div {
        color: var(--mlg-dark-teal) !important;
    }

    section[data-testid="stSidebar"] [data-baseweb="input"],
    section[data-testid="stSidebar"] [data-baseweb="textarea"],
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] input {
        background: #161b22 !important;
        border-color: var(--mlg-dark-line) !important;
        color: var(--mlg-dark-ink) !important;
    }

    section[data-testid="stSidebar"] textarea::placeholder,
    section[data-testid="stSidebar"] input::placeholder {
        color: var(--mlg-dark-muted) !important;
    }

    textarea,
    input {
        border-radius: 8px !important;
    }
    }

    :root {
        --mlg-forced-bg: #0e1117;
        --mlg-forced-panel: #161b22;
        --mlg-forced-panel-soft: #1b222c;
        --mlg-forced-ink: #fafafa;
        --mlg-forced-muted: #a3aab7;
        --mlg-forced-line: #30363d;
        --mlg-forced-accent: #d8be6a;
        --mlg-forced-teal: #5db9bd;
    }

    .stApp {
        background: var(--mlg-forced-bg) !important;
        color: var(--mlg-forced-ink) !important;
    }

    h1,
    h2,
    h3,
    label,
    p,
    span,
    small,
    [data-testid="stMarkdownContainer"] {
        color: var(--mlg-forced-ink) !important;
    }

    section[data-testid="stSidebar"] {
        background: #11161d !important;
        border-right: 1px solid var(--mlg-forced-line) !important;
    }

    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] * {
        color: var(--mlg-forced-ink) !important;
    }

    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: var(--mlg-forced-muted) !important;
    }

    div[data-testid="stFileUploader"],
    div[data-testid="stFileUploaderDropzone"] {
        background: var(--mlg-forced-panel) !important;
        border-color: var(--mlg-forced-line) !important;
    }

    div[data-testid="stFileUploaderDropzone"] button,
    .stButton > button,
    .stDownloadButton > button {
        background: var(--mlg-forced-panel-soft) !important;
        border: 1px solid var(--mlg-forced-line) !important;
        color: var(--mlg-forced-ink) !important;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: var(--mlg-forced-accent) !important;
        color: #fff7cf !important;
        background: #202733 !important;
    }

    button[kind="primary"] {
        border-color: var(--mlg-forced-teal) !important;
        background: #2d6f73 !important;
        color: #ffffff !important;
    }

    div[data-testid="stMetric"],
    div[data-testid="stExpander"] {
        background: var(--mlg-forced-panel) !important;
        border-color: var(--mlg-forced-line) !important;
    }

    div[data-testid="stDataFrame"],
    div[data-testid="stDataEditor"] {
        border: 1px solid var(--mlg-forced-line) !important;
        box-shadow: 0 10px 26px rgba(0, 0, 0, 0.28) !important;
    }

    [data-baseweb="input"],
    [data-baseweb="textarea"],
    textarea,
    input {
        background: var(--mlg-forced-panel) !important;
        border-color: var(--mlg-forced-line) !important;
        color: var(--mlg-forced-ink) !important;
    }

    textarea::placeholder,
    input::placeholder {
        color: var(--mlg-forced-muted) !important;
    }

    .magic-loader {
        display: flex;
        align-items: center;
        gap: 1rem;
        min-height: 96px;
        margin: 0.75rem 0 1rem;
        padding: 1rem 1.25rem;
        border: 1px solid rgba(184, 138, 34, 0.28);
        border-radius: 8px;
        background: rgba(255, 250, 240, 0.72);
        overflow: hidden;
    }

    .magic-loader-stage {
        position: relative;
        width: 150px;
        height: 88px;
        flex: 0 0 150px;
    }

    .magic-wizard {
        position: absolute;
        left: 8px;
        bottom: 8px;
        width: 56px;
        height: 70px;
        animation: wizard-hop 1.15s ease-in-out infinite;
    }

    .magic-wizard .hat {
        position: absolute;
        left: 12px;
        top: 0;
        width: 0;
        height: 0;
        border-left: 15px solid transparent;
        border-right: 15px solid transparent;
        border-bottom: 36px solid #2d6f73;
        filter: drop-shadow(0 2px 0 rgba(24, 32, 36, 0.22));
        transform: rotate(-8deg);
        z-index: 3;
    }

    .magic-wizard .hat::after {
        content: "";
        position: absolute;
        left: -8px;
        top: 20px;
        width: 9px;
        height: 9px;
        border-radius: 50%;
        background: #d8be6a;
        box-shadow: 0 0 8px rgba(216, 190, 106, 0.9);
    }

    .magic-wizard .brim {
        position: absolute;
        left: 7px;
        top: 29px;
        width: 42px;
        height: 9px;
        border-radius: 999px;
        background: #1f4f52;
        box-shadow: 0 2px 0 rgba(24, 32, 36, 0.18);
        z-index: 4;
    }

    .magic-wizard .face {
        position: absolute;
        left: 15px;
        top: 34px;
        width: 26px;
        height: 22px;
        border-radius: 50%;
        background: #f0cda5;
        z-index: 2;
    }

    .magic-wizard .face::before,
    .magic-wizard .face::after {
        content: "";
        position: absolute;
        top: 8px;
        width: 3px;
        height: 3px;
        border-radius: 50%;
        background: #182024;
    }

    .magic-wizard .face::before {
        left: 8px;
    }

    .magic-wizard .face::after {
        right: 8px;
    }

    .magic-wizard .beard {
        position: absolute;
        left: 12px;
        top: 48px;
        width: 32px;
        height: 23px;
        border-radius: 45% 45% 60% 60%;
        background: #f6f2e8;
        box-shadow: inset 0 -4px 0 rgba(216, 190, 106, 0.22);
        z-index: 1;
    }

    .magic-wizard .robe {
        position: absolute;
        left: 7px;
        bottom: 0;
        width: 42px;
        height: 34px;
        border-radius: 10px 10px 5px 5px;
        background: #1f4f52;
        box-shadow: inset 10px 0 0 rgba(45, 111, 115, 0.9);
    }

    .magic-wizard .sleeve {
        position: absolute;
        left: 39px;
        top: 48px;
        width: 24px;
        height: 12px;
        border-radius: 999px;
        background: #2d6f73;
        transform: rotate(-24deg);
        z-index: 2;
    }

    .magic-wizard .wand {
        position: absolute;
        left: 56px;
        top: 42px;
        width: 42px;
        height: 3px;
        border-radius: 999px;
        background: #6f5011;
        transform-origin: left center;
        transform: rotate(-30deg);
        animation: wand-flick 1.15s ease-in-out infinite;
        z-index: 4;
    }

    .magic-wizard .wand::after {
        content: "";
        position: absolute;
        right: -3px;
        top: -3px;
        width: 9px;
        height: 9px;
        border-radius: 50%;
        background: #fff4a3;
        box-shadow: 0 0 12px #d8be6a, 0 0 20px rgba(45, 111, 115, 0.55);
    }

    .magic-spark {
        position: absolute;
        width: 16px;
        height: 16px;
        background: #ffd34f;
        clip-path: polygon(50% 0%, 62% 34%, 98% 35%, 69% 55%, 79% 91%, 50% 70%, 21% 91%, 31% 55%, 2% 35%, 38% 34%);
        filter:
            drop-shadow(0 0 2px #5f430b)
            drop-shadow(0 0 8px rgba(255, 211, 79, 0.95))
            drop-shadow(0 0 14px rgba(45, 111, 115, 0.5));
        opacity: 0;
        animation: spark-pop 1.15s ease-out infinite;
    }

    .magic-spark.one {
        left: 104px;
        top: 12px;
    }

    .magic-spark.two {
        left: 126px;
        top: 34px;
        animation-delay: 0.16s;
    }

    .magic-spark.three {
        left: 106px;
        top: 58px;
        animation-delay: 0.28s;
    }

    .magic-loader-text strong {
        display: block;
        color: #182024;
        font-weight: 760;
        margin-bottom: 0.2rem;
    }

    .magic-loader-text span {
        color: #63717a;
        font-size: 0.94rem;
    }

    @keyframes wizard-hop {
        0%, 100% { transform: translateX(0) translateY(0) rotate(-2deg); }
        40% { transform: translateX(28px) translateY(-12px) rotate(4deg); }
        70% { transform: translateX(48px) translateY(0) rotate(-3deg); }
    }

    @keyframes wand-flick {
        0%, 100% { transform: rotate(-30deg); }
        45% { transform: rotate(-48deg); }
    }

    @keyframes spark-pop {
        0% { transform: scale(0.4); opacity: 0; }
        35% { transform: scale(1); opacity: 1; }
        100% { transform: translateX(22px) translateY(-12px) scale(0.2); opacity: 0; }
    }

    html,
    body,
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stDecoration"] {
        background-color: #0e1117 !important;
        color: #fafafa !important;
    }

    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div,
    section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        background-color: #11161d !important;
        color: #fafafa !important;
    }

    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] div {
        color: #fafafa !important;
    }

    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] *,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: #a3aab7 !important;
    }

    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:has([data-testid="stFileUploader"]),
    section[data-testid="stSidebar"] [data-testid="stElementContainer"]:has([data-testid="stFileUploader"]),
    section[data-testid="stSidebar"] div:has(> [data-testid="stFileUploader"]) {
        background-color: transparent !important;
        border-color: transparent !important;
        box-shadow: none !important;
    }

    section[data-testid="stSidebar"] [data-testid="stFileUploader"],
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] > div,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] > section,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] label,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"],
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] section {
        background-color: #161b22 !important;
        border-color: #30363d !important;
        color: #fafafa !important;
    }

    section[data-testid="stSidebar"] .stFileUploader,
    section[data-testid="stSidebar"] div.stFileUploader[data-testid="stFileUploader"] {
        background-color: transparent !important;
        border: 0 !important;
        box-shadow: none !important;
        padding: 0 !important;
    }

    section[data-testid="stSidebar"] .stFileUploader > label,
    section[data-testid="stSidebar"] div.stFileUploader[data-testid="stFileUploader"] > label,
    section[data-testid="stSidebar"] .stFileUploader [data-testid="stWidgetLabel"] {
        background-color: transparent !important;
        border: 0 !important;
        box-shadow: none !important;
        color: #fafafa !important;
        padding: 0 0 0.35rem 0 !important;
    }

    section[data-testid="stSidebar"] .stFileUploader [data-testid="stFileUploaderDropzone"],
    section[data-testid="stSidebar"] div.stFileUploader[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzone"] {
        background-color: #161b22 !important;
        border: 1px solid #30363d !important;
        border-radius: 8px !important;
        box-shadow: none !important;
    }

    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
    section[data-testid="stSidebar"] button {
        background-color: #1b222c !important;
        border-color: #30363d !important;
        color: #fafafa !important;
    }

    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] svg,
    section[data-testid="stSidebar"] button svg {
        color: #fafafa !important;
        fill: #fafafa !important;
    }

    section[data-testid="stSidebar"] [data-testid="stTextArea"],
    section[data-testid="stSidebar"] [data-testid="stTextArea"] > div,
    section[data-testid="stSidebar"] [data-testid="stTextAreaRootElement"],
    section[data-testid="stSidebar"] [data-baseweb="textarea"],
    section[data-testid="stSidebar"] [data-baseweb="textarea"] > div,
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] [role="textbox"] {
        background-color: #161b22 !important;
        border-color: #30363d !important;
        color: #fafafa !important;
        caret-color: #fafafa !important;
    }

    section[data-testid="stSidebar"] textarea::placeholder,
    section[data-testid="stSidebar"] input::placeholder {
        color: #a3aab7 !important;
    }

    section[data-testid="stSidebar"] [data-baseweb="slider"] div,
    section[data-testid="stSidebar"] [data-baseweb="slider"] span {
        color: #5db9bd !important;
    }

    .magic-loader {
        background: rgba(22, 27, 34, 0.82) !important;
        border-color: rgba(216, 190, 106, 0.34) !important;
    }

    .magic-loader-text strong {
        color: #fafafa !important;
    }

    .magic-loader-text span {
        color: #a3aab7 !important;
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
    st.session_state.export_output_df = pd.DataFrame()
    st.session_state.export_data_key = None
    st.session_state.review_editor_version = st.session_state.get("review_editor_version", 0) + 1
    st.session_state.export_editor_version = st.session_state.get("export_editor_version", 0) + 1


def render_magic_loader() -> None:
    st.markdown(
        """
        <div class="magic-loader">
            <div class="magic-loader-stage">
                <div class="magic-wizard" aria-hidden="true">
                    <div class="hat"></div>
                    <div class="brim"></div>
                    <div class="face"></div>
                    <div class="beard"></div>
                    <div class="robe"></div>
                    <div class="sleeve"></div>
                    <div class="wand"></div>
                </div>
                <div class="magic-spark one"></div>
                <div class="magic-spark two"></div>
                <div class="magic-spark three"></div>
            </div>
            <div class="magic-loader-text">
                <strong>Bearbetar leads</strong>
                <span>Matchar bolag, domäner och e-post innan raderna delas upp.</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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
        "background-color: rgba(255, 205, 45, 0.42); "
        "color: #fafafa"
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
        values = row.drop(labels=[SELECT_COLUMN], errors="ignore").to_dict()
        row_matches = saved_working_df.index[saved_working_df["Rad"] == row_id].tolist()
        if row_matches:
            row_index = row_matches[0]
            for column, value in values.items():
                if column in saved_working_df.columns:
                    saved_working_df.at[row_index, column] = value
    st.session_state.working_df = saved_working_df
    return edited_not_removed, removed_rows


def apply_export_edits(edited_export: pd.DataFrame, original_export_rows: set[int]) -> tuple[pd.DataFrame, set[int]]:
    if "Rad" not in edited_export.columns:
        return edited_export, set()

    edited_export_rows = set(edited_export["Rad"].dropna().astype(int).tolist())
    removed_rows = original_export_rows - edited_export_rows
    for row_id in removed_rows:
        st.session_state.excluded_rows.add(row_id)

    edited_not_removed = edited_export[edited_export["Rad"].notna()].copy()
    saved_working_df = st.session_state.working_df.copy()
    for _, row in edited_not_removed.iterrows():
        row_id = int(row["Rad"])
        row_matches = saved_working_df.index[saved_working_df["Rad"] == row_id].tolist()
        if not row_matches:
            continue
        row_index = row_matches[0]
        for column, value in row.drop(labels=["Rad", SELECT_COLUMN], errors="ignore").items():
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
if "export_editor_version" not in st.session_state:
    st.session_state.export_editor_version = 0
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "Alla rader"
if "clean_input_key" not in st.session_state:
    st.session_state.clean_input_key = None
if "working_df" not in st.session_state:
    st.session_state.working_df = pd.DataFrame()
if "export_output_df" not in st.session_state:
    st.session_state.export_output_df = pd.DataFrame()
if "export_data_key" not in st.session_state:
    st.session_state.export_data_key = None

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
    loader = st.empty()
    with loader.container():
        render_magic_loader()
    with st.spinner("Bearbetar leads..."):
        cleaned_df = clean_leads(raw_df, registry_df, options).reset_index(drop=True)
    loader.empty()
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

active_tab = st.segmented_control(
    "Flik",
    ["Alla rader", "Granska", "Klara", "Exkluderade", "Export"],
    key="active_tab",
    label_visibility="collapsed",
)

if active_tab == "Alla rader":
    edited_all = st.data_editor(
        working_df,
        use_container_width=True,
        hide_index=True,
        height=TABLE_HEIGHT,
        num_rows="dynamic",
        key="all_editor",
    )

elif active_tab == "Granska":
    review_df = working_df[review_mask].copy()
    review_df.insert(0, SELECT_COLUMN, st.session_state.review_mark_all)
    if st.session_state.review_mark_all:
        st.info("Bulkmarkering är aktiv: alla kvarvarande rader i Granska är markerade tills du avmarkerar.")
    with st.form(f"review_form_{st.session_state.review_editor_version}"):
        edited_review = st.data_editor(
            review_df,
            use_container_width=True,
            hide_index=True,
            height=TABLE_HEIGHT,
            num_rows="dynamic",
            key=f"review_editor_{st.session_state.review_editor_version}",
            column_config={
                SELECT_COLUMN: st.column_config.CheckboxColumn(SELECT_COLUMN, width="small"),
                "Rad": None,
            },
        )
        review_button_cols = st.columns([1.8, 1.35, 1.35, 1.5])
        with review_button_cols[0]:
            submitted_review = st.form_submit_button("Verkställ ändringar", type="primary", use_container_width=True)
        with review_button_cols[1]:
            submitted_exclude = st.form_submit_button("Exkludera valda", use_container_width=True)
        with review_button_cols[2]:
            submitted_mark_all = st.form_submit_button("Markera alla", use_container_width=True)
        with review_button_cols[3]:
            submitted_clear_all = st.form_submit_button("Avmarkera alla", use_container_width=True)

    selected_review_rows = pd.DataFrame()
    if submitted_review or submitted_exclude:
        original_review_rows = set(review_df["Rad"].astype(int).tolist())
        edited_not_removed, removed_rows = apply_review_edits(edited_review, original_review_rows)
        if removed_rows:
            st.session_state.review_editor_version += 1
            st.rerun()
        selected_review_rows = edited_not_removed[edited_not_removed[SELECT_COLUMN] == True]

    if submitted_review:
        if selected_review_rows.empty:
            st.warning("Markera minst en rad först.")
        else:
            for _, row in selected_review_rows.iterrows():
                st.session_state.reviewed_rows.add(int(row["Rad"]))
            st.session_state.review_mark_all = False
            st.session_state.review_editor_version += 1
            st.rerun()
    if submitted_exclude:
        if selected_review_rows.empty:
            st.warning("Markera minst en rad först.")
        else:
            for _, row in selected_review_rows.iterrows():
                st.session_state.excluded_rows.add(int(row["Rad"]))
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
    st.caption("Granska är ett sidospår. Redigera direkt i tabellen, markera rader med Välj och klicka Verkställ ändringar eller Exkludera valda.")

elif active_tab == "Klara":
    ready_df = working_df[ready_mask]
    st.dataframe(ready_df, use_container_width=True, hide_index=True, height=TABLE_HEIGHT)

elif active_tab == "Exkluderade":
    excluded_df = working_df[excluded_mask]
    st.dataframe(excluded_df, use_container_width=True, hide_index=True, height=TABLE_HEIGHT)
    if not excluded_df.empty and st.button("Återställ alla exkluderade"):
        st.session_state.excluded_rows = set()
        st.rerun()

elif active_tab == "Export":
    export_df = working_df[ready_mask]
    import_df = to_import_columns(export_df)
    import_df.insert(0, "Rad", export_df["Rad"].to_numpy())
    import_df.insert(0, SELECT_COLUMN, False)

    duplicate_file = st.session_state.get("duplicate_file")
    duplicate_sheets = []
    duplicate_count = 0
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

    if duplicate_count:
        highlighted_export_df = import_df.drop(columns=[SELECT_COLUMN, "Rad"], errors="ignore")
        st.dataframe(
            highlighted_export_df.style.apply(highlight_duplicate_rows, axis=1),
            use_container_width=True,
            hide_index=True,
            height=min(TABLE_HEIGHT, 360),
        )
        st.caption("Gulmarkerade rader är potentiella dubletter. Redigera exportdatan i tabellen nedan.")

    export_data_key = (
        tuple(import_df["Rad"].dropna().astype(int).tolist()),
        tuple(import_df.get("Duplicate Match", pd.Series("", index=import_df.index)).fillna("").astype(str).tolist()),
        tuple(import_df.get("Issues", pd.Series("", index=import_df.index)).fillna("").astype(str).tolist()),
    )

    with st.form(f"export_form_{st.session_state.export_editor_version}"):
        editor_data = import_df
        if "Duplicate Match" in import_df.columns and import_df["Duplicate Match"].map(has_duplicate_match).any():
            editor_data = import_df.style.apply(highlight_duplicate_rows, axis=1)
        edited_import_df = st.data_editor(
            editor_data,
            use_container_width=True,
            hide_index=True,
            height=TABLE_HEIGHT,
            num_rows="dynamic",
            key=f"export_editor_{st.session_state.export_editor_version}",
            column_config={
                SELECT_COLUMN: st.column_config.CheckboxColumn(SELECT_COLUMN, width="small"),
                "Rad": None,
            },
        )
        export_button_cols = st.columns([1.8, 1.35, 2.85])
        with export_button_cols[0]:
            submitted_export = st.form_submit_button("Verkställ ändringar", type="primary", use_container_width=True)
        with export_button_cols[1]:
            submitted_export_exclude = st.form_submit_button("Exkludera valda", use_container_width=True)

    if submitted_export or submitted_export_exclude:
        original_export_rows = set(import_df["Rad"].astype(int).tolist())
        edited_import_df, removed_export_rows = apply_export_edits(edited_import_df, original_export_rows)
        if removed_export_rows:
            st.session_state.export_editor_version += 1
            st.rerun()
        selected_export_rows = edited_import_df[edited_import_df[SELECT_COLUMN] == True]
        st.session_state.export_output_df = edited_import_df.drop(columns=["Rad", SELECT_COLUMN], errors="ignore")
        st.session_state.export_data_key = export_data_key

        if submitted_export_exclude:
            if not selected_export_rows.empty:
                st.session_state.excluded_rows.update(selected_export_rows["Rad"].dropna().astype(int).tolist())
                st.session_state.export_editor_version += 1
                st.rerun()
            else:
                st.warning("Markera minst en rad först.")
        elif submitted_export:
            st.success("Ändringarna är verkställda.")
    elif st.session_state.export_data_key != export_data_key:
        st.session_state.export_output_df = import_df.drop(columns=["Rad", SELECT_COLUMN], errors="ignore")
        st.session_state.export_data_key = export_data_key

    export_output_df = st.session_state.export_output_df

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

    excel_bytes = dataframe_to_xlsx_bytes(export_output_df)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_file_name = f"upsales_import_{timestamp}.xlsx"
    st.download_button(
        "Exportera till Excel",
        data=excel_bytes,
        file_name=excel_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.download_button(
        "Ladda ner CSV",
        data=export_output_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="upsales_import.csv",
        mime="text/csv",
    )
