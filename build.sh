#!/usr/bin/env bash
set -euo pipefail

TASK="${1:-package}"
VENV_PATH="${VENV_PATH:-.venv}"
ENTRY="${ENTRY:-run_app.py}"
NAME="${NAME:-SolicituProd}"
CONSOLE="${CONSOLE:-0}"
ICON="${ICON:-}"
ONEDIR="${ONEDIR:-0}"
NOCLEAN="${NOCLEAN:-0}"
INSTALL_MISSING="${INSTALL_MISSING:-0}"
REQUIREMENTS_FILE="${REQUIREMENTS_FILE:-requirements.txt}"
HOST_PYTHON_PATH="${HOST_PYTHON_PATH:-}"

info() { printf '[INFO] %s\n' "$*"; }
warn() { printf '[WARN] %s\n' "$*" >&2; }
fail() { printf '[ERR] %s\n' "$*" >&2; exit 1; }

resolve_host_python() {
  if [ -n "$HOST_PYTHON_PATH" ] && [ -x "$HOST_PYTHON_PATH" ]; then
    printf '%s\n' "$HOST_PYTHON_PATH"
    return
  fi
  if command -v python3 >/dev/null 2>&1; then
    command -v python3
    return
  fi
  if command -v python >/dev/null 2>&1; then
    command -v python
    return
  fi
  fail "Python no encontrado. Use HOST_PYTHON_PATH o instale Python 3.x"
}

get_python_path() {
  local venv_py="$VENV_PATH/bin/python"
  if [ -x "$venv_py" ]; then
    printf '%s\n' "$venv_py"
    return
  fi
  resolve_host_python
}

new_venv() {
  local py
  py="$(resolve_host_python)"
  if [ ! -d "$VENV_PATH" ]; then
    "$py" -m venv "$VENV_PATH"
  fi
}

install_deps() {
  local py="$1"
  "$py" -m pip install --upgrade pip
  if [ -f "$REQUIREMENTS_FILE" ]; then
    info "Instalando dependencias desde $REQUIREMENTS_FILE"
    "$py" -m pip install -r "$REQUIREMENTS_FILE"
  else
    warn "No se encontro $REQUIREMENTS_FILE; instalando conjunto minimo."
    "$py" -m pip install numpy pandas reportlab pillow openpyxl xlrd pypdf
  fi
}

init_tools_and_deps() {
  local py="$1"
  if [ "$INSTALL_MISSING" = "1" ]; then
    info "Verificando e instalando herramientas (-INSTALL_MISSING=1)"
    "$py" - <<'PY' || "$py" -m pip install -U pyinstaller pyinstaller-hooks-contrib
import importlib.util, sys
sys.exit(0 if importlib.util.find_spec("PyInstaller") else 1)
PY
    "$py" - <<'PY' || "$py" -m pip install -U numpy pandas reportlab pillow openpyxl xlrd pypdf
import importlib.util as u, sys
mods = ["numpy","pandas","reportlab","PIL","openpyxl","xlrd","pypdf","PyPDF2"]
missing = [m for m in mods if u.find_spec(m) is None]
sys.exit(0 if not missing else 1)
PY
  fi

    "$py" - <<'PY' || "$py" -m pip install -U pyinstaller pyinstaller-hooks-contrib
import importlib.util, sys
sys.exit(0 if importlib.util.find_spec("PyInstaller") else 1)
PY
}

new_output_folders() {
  mkdir -p Vales_Historial
}

start_app() {
  local py="$1"
  new_output_folders
  "$py" "$ENTRY"
}

new_app_package() {
  local py="$1"
  init_tools_and_deps "$py"
  [ -f "$ENTRY" ] || fail "No se encontro el entrypoint: $ENTRY"

  local args=()
  args+=("$ENTRY" "--name" "$NAME")
  if [ "$ONEDIR" = "0" ]; then
    args+=("--onefile")
  fi
  if [ "$NOCLEAN" = "0" ]; then
    args+=("--clean")
  fi
  args+=("--noconfirm")
  if [ "$CONSOLE" = "1" ]; then
    args+=("--console")
  else
    args+=("--windowed")
  fi
  if [ -n "$ICON" ] && [ -f "$ICON" ]; then
    args+=("--icon" "$ICON")
  fi

  args+=("--hidden-import=reportlab" "--hidden-import=numpy" "--hidden-import=PIL" "--hidden-import=openpyxl")
  args+=("--hidden-import=xlrd" "--hidden-import=pypdf" "--hidden-import=PyPDF2")
  args+=("--collect-all" "numpy" "--collect-all" "reportlab" "--collect-all" "PIL" "--collect-all" "openpyxl")
  args+=("--collect-all" "pypdf")

  info "Ejecutando: $py -m PyInstaller ${args[*]}"
  "$py" -m PyInstaller "${args[@]}"

  local dist_root="dist"
  local dist_base="$dist_root"
  if [ "$ONEDIR" = "1" ]; then
    dist_base="$dist_root/$NAME"
  fi
  mkdir -p "$dist_base/Vales_Historial"
  if [ -d "Vales_Historial" ]; then
    cp -R "Vales_Historial/"* "$dist_base/Vales_Historial/" 2>/dev/null || true
  fi
  if [ -f "app_settings.json" ]; then
    cp -f "app_settings.json" "$dist_base/app_settings.json"
  fi
  if [ -f "instrucciones.txt" ]; then
    cp -f "instrucciones.txt" "$dist_base/instrucciones.txt"
  fi

  info "Paquete generado en: $dist_base"
}

remove_build_artifacts() {
  rm -rf build dist ./*.spec
}

case "$TASK" in
  setup)
    new_venv
    py="$(get_python_path)"
    install_deps "$py"
    ;;
  run)
    new_venv
    py="$(get_python_path)"
    install_deps "$py"
    start_app "$py"
    ;;
  package)
    new_venv
    py="$(get_python_path)"
    install_deps "$py"
    new_app_package "$py"
    ;;
  clean)
    remove_build_artifacts
    ;;
  all)
    new_venv
    py="$(get_python_path)"
    install_deps "$py"
    new_app_package "$py"
    ;;
  *)
    fail "Tarea desconocida: $TASK (use setup|run|package|clean|all)"
    ;;
esac
