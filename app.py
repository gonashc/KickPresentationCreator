from __future__ import annotations

import os
from datetime import date
from pathlib import Path

import streamlit as st

from sow_kickoff.ai_planner import build_deck_plan
from sow_kickoff.docx_reader import extract_sow
from sow_kickoff.pptx_renderer import render_presentation
from sow_kickoff.template_analyzer import analyze_template


st.set_page_config(
    page_title="Kickoff Studio",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .stApp { background: #f4f6fa; }
      .block-container { max-width: 1180px; padding-top: 2.25rem; padding-bottom: 4rem; }
      h1, h2, h3 { color: #17213c; letter-spacing: -0.025em; }
      h1 { font-size: 2.65rem !important; line-height: 1.05 !important; }
      [data-testid="stFileUploader"] {
        background: white; border: 1px solid #e2e6f0; border-radius: 16px;
        padding: 0.5rem 1rem 0.8rem; box-shadow: 0 7px 25px rgba(23,33,60,.05);
      }
      [data-testid="stFileUploaderDropzone"] { background: #fafbff; border-color: #cfd6eb; }
      .hero-kicker { color: #3657d6; font-weight: 750; letter-spacing: .12em; font-size: .77rem; }
      .hero-copy { color: #5e6882; font-size: 1.08rem; max-width: 720px; line-height: 1.6; }
      .step-label { color: #3657d6; font-size: .75rem; letter-spacing: .1em; font-weight: 800; }
      .trust-row { color: #6b748c; font-size: .87rem; margin-top: .2rem; }
      .status-card {
        background: white; border: 1px solid #e2e6f0; border-radius: 16px;
        padding: 1.15rem 1.3rem; min-height: 92px;
      }
      .status-card strong { color: #17213c; display: block; margin-bottom: .28rem; }
      .status-card span { color: #6b748c; font-size: .88rem; }
      div.stButton > button[kind="primary"] {
        min-height: 3.25rem; border-radius: 12px; font-weight: 750; width: 100%;
        box-shadow: 0 8px 20px rgba(54,87,214,.18);
      }
      div[data-testid="stDownloadButton"] > button { width: 100%; min-height: 3rem; border-radius: 12px; }
      .fine-print { color: #7a8399; font-size: .78rem; }
      #MainMenu, footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _dotenv_value(name: str) -> str:
    """Read a simple KEY=value entry from the project-local .env file."""
    env_path = Path(__file__).resolve().with_name(".env")
    try:
        lines = env_path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return ""

    for line in lines:
        entry = line.strip()
        if not entry or entry.startswith("#"):
            continue
        if entry.startswith("export "):
            entry = entry[7:].lstrip()
        key, separator, value = entry.partition("=")
        if not separator or key.strip() != name:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
            value = value[1:-1]
        elif " #" in value:
            value = value.split(" #", 1)[0].rstrip()
        return value
    return ""


def _secret(name: str, default: str = "") -> str:
    try:
        configured = st.secrets.get(name)
        if configured:
            return str(configured)
    except Exception:
        pass
    return os.getenv(name) or _dotenv_value(name) or default


def _friendly_date(value: date) -> str:
    return f"{value.strftime('%B')} {value.day}, {value.year}"


st.markdown('<div class="hero-kicker">SOW → KICKOFF DECK</div>', unsafe_allow_html=True)
st.title("Build a client-ready kickoff deck")
st.markdown(
    '<div class="hero-copy">Bring your approved statement of work and branded PowerPoint. '
    "Kickoff Studio maps only supported facts into the right slides and returns an editable presentation.</div>",
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

with st.expander("Settings", expanded=False):
    model = st.text_input("Model", value=_secret("OPENAI_MODEL", "gpt-5.4-mini"))

st.markdown('<div class="step-label">01 · SOURCE FILES</div>', unsafe_allow_html=True)
left, right = st.columns(2, gap="large")
with left:
    st.subheader("PowerPoint template")
    template_file = st.file_uploader(
        "Upload branded template",
        type=["pptx"],
        label_visibility="collapsed",
        help="The original file is never modified.",
    )
with right:
    st.subheader("Statement of work")
    sow_file = st.file_uploader(
        "Upload signed or approved SoW",
        type=["docx"],
        label_visibility="collapsed",
    )

st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="step-label">02 · PROJECT DETAILS</div>', unsafe_allow_html=True)
details_1, details_2, details_3 = st.columns([1.15, 1.35, 1])
with details_1:
    client_name = st.text_input("Client name", placeholder="Acme Corporation")
with details_2:
    project_name = st.text_input("Project name", placeholder="Implementation kickoff")
with details_3:
    kickoff_date = st.date_input("Kickoff date", value=date.today())

generation_notes = st.text_area(
    "Optional instructions",
    placeholder="Example: Emphasize the phased rollout and executive governance cadence.",
    height=86,
)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="step-label">03 · GENERATE</div>', unsafe_allow_html=True)

ready = template_file is not None and sow_file is not None
button_label = "Generate kickoff presentation" if ready else "Add both source files to continue"
generate = st.button(button_label, type="primary", disabled=not ready, use_container_width=True)
st.markdown(
    '<div class="trust-row">✓ Template-safe editing &nbsp;&nbsp; ✓ Evidence-only content &nbsp;&nbsp; '
    "✓ Missing details stay blank &nbsp;&nbsp; ✓ Editable PowerPoint output</div>",
    unsafe_allow_html=True,
)

if generate:
    resolved_key = _secret("OPENAI_API_KEY")
    if not resolved_key:
        st.error("Set OPENAI_API_KEY in the local .env file or process environment before generating.")
        st.stop()

    template_bytes = template_file.getvalue()
    sow_bytes = sow_file.getvalue()

    try:
        with st.status("Building your kickoff deck…", expanded=True) as status:
            st.write("Reading the statement of work")
            source = extract_sow(sow_bytes)
            st.write(f"Found {source.paragraph_count} paragraphs and {source.table_count} tables")

            st.write("Understanding the PowerPoint layout")
            inventory = analyze_template(template_bytes)
            st.write(f"Mapped {inventory.slide_count} slides and {inventory.editable_target_count} editable targets")

            st.write("Planning the presentation as an account manager")
            plan = build_deck_plan(
                source=source,
                inventory=inventory,
                api_key=resolved_key,
                model=model.strip(),
                client_name=client_name.strip(),
                project_name=project_name.strip(),
                kickoff_date=kickoff_date.isoformat(),
                generation_notes=generation_notes.strip(),
            )

            st.write("Applying approved content to a copy of the template")
            result = render_presentation(
                template_bytes=template_bytes,
                plan=plan,
                metadata={
                    "client_name": client_name.strip(),
                    "project_name": project_name.strip(),
                    "kickoff_date": _friendly_date(kickoff_date),
                },
            )
            status.update(label="Kickoff deck is ready", state="complete", expanded=False)

        st.success("Presentation generated successfully. Review it in PowerPoint before sharing with the client.")

        file_stem = (client_name or project_name or "client").strip().lower().replace(" ", "-")
        output_name = f"{file_stem}-kickoff.pptx"
        download_col, summary_col = st.columns([1, 1.5], gap="large")
        with download_col:
            st.download_button(
                "Download PowerPoint",
                data=result.presentation_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
            )
        with summary_col:
            st.markdown(
                f'<div class="status-card"><strong>{result.applied_count} updates applied</strong>'
                f"<span>{result.skipped_count} invalid or unsafe targets skipped · "
                f"{len(plan.warnings)} review note(s)</span></div>",
                unsafe_allow_html=True,
            )

        with st.expander("Generation audit"):
            st.caption("The audit shows what the AI requested and what the renderer actually changed.")
            st.json(
                {
                    "executive_summary": plan.executive_summary,
                    "applied_updates": result.applied_updates,
                    "skipped_updates": result.skipped_updates,
                    "warnings": plan.warnings,
                }
            )
    except Exception as exc:
        st.error(f"The deck could not be generated: {exc}")
        st.info("Your uploaded files were not modified. Check the file types and API settings, then try again.")

st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    '<div class="fine-print">Kickoff Studio creates a first draft from the uploaded SoW. '
    "A project owner should confirm scope, dates, responsibilities, and commercial commitments before client use.</div>",
    unsafe_allow_html=True,
)

