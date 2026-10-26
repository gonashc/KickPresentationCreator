$ErrorActionPreference = "Stop"
$workspace = $PSScriptRoot
$env:UV_CACHE_DIR = Join-Path $workspace ".uv-cache"
$env:UV_PYTHON_INSTALL_DIR = Join-Path $workspace ".uv-python"
$env:UV_PYTHON_PREFERENCE = "only-managed"

uv python install 3.12
uv sync --dev --python 3.12
& (Join-Path $workspace ".venv\Scripts\python.exe") -m streamlit run (Join-Path $workspace "app.py")
