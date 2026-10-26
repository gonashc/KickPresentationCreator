# Kickoff Studio

Kickoff Studio turns a branded PowerPoint template and a Word statement of work into a client-ready kickoff deck. It edits a copy of the template, so the original slide master, theme, artwork, and untouched content remain intact.

## What the MVP does

- Accepts `.pptx` templates and `.docx` statements of work
- Extracts Word paragraphs and tables with source references
- Inventories editable PowerPoint text shapes and table cells
- Uses a structured OpenAI response to choose what belongs on each slide
- Leaves unsupported or missing information blank
- Applies exact metadata tokens such as `{{CLIENT_NAME}}` deterministically
- Produces a downloadable `.pptx` plus an audit report
- Holds uploaded documents in memory for the active Streamlit session

## Run locally

This repository uses `uv`, which will create and manage the Python environment.

```powershell
uv sync --dev
uv run streamlit run app.py
```

Copy `.env.example` to `.env` and set `OPENAI_API_KEY` there before starting the app. The application reads `.env` automatically; the key is not entered through the web interface. `.env` is excluded from version control.

## Template preparation

The application can infer editable content from slide titles and existing text. Reliability improves when text boxes use descriptive names in PowerPoint's Selection Pane, for example:

- `CLIENT_NAME`
- `PROJECT_NAME`
- `OBJECTIVES`
- `IN_SCOPE`
- `OUT_OF_SCOPE`
- `DELIVERABLES`
- `TIMELINE`
- `ROLES`
- `GOVERNANCE`
- `RISKS`
- `ASSUMPTIONS`
- `NEXT_STEPS`

Supported deterministic tokens are `{{CLIENT_NAME}}`, `{{PROJECT_NAME}}`, `{{KICKOFF_DATE}}`, and equivalent square-bracket forms such as `[Client Name]`.

## Current boundaries

- Standard text boxes, placeholders, and existing table cells are editable.
- SmartArt, embedded workbooks, charts, macros, video, and grouped graphics are preserved but are not edited.
- The MVP does not add new slides. It fills the layouts already present in the template.
- SharePoint selection and save-back are the next integration layer; local upload and download are complete first so generation does not depend on tenant permissions.

## Tests

```powershell
uv run pytest
```

