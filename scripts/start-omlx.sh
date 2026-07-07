#!/usr/bin/env bash
set -euo pipefail

: "${OMLX_BIN:=omlx}"
: "${OMLX_MODEL_DIR:=$HOME/AI/models}"
: "${OMLX_BASE_PATH:=$HOME/.omlx}"
: "${OMLX_PORT:=8001}"
: "${OMLX_API_KEY:=}"
: "${OMLX_PRELOAD_MODEL:=Qwen3.6-35B-A3B-4.4bit-msq}"
: "${OMLX_CONFIGURE_MODEL_SETTINGS:=1}"

if ! command -v "$OMLX_BIN" >/dev/null 2>&1; then
  echo "oMLX command not found: $OMLX_BIN" >&2
  echo "Install oMLX or set OMLX_BIN to the full executable path." >&2
  exit 1
fi

if [[ ! -d "$OMLX_MODEL_DIR" ]]; then
  echo "oMLX model directory does not exist: $OMLX_MODEL_DIR" >&2
  exit 1
fi

if [[ -n "$OMLX_PRELOAD_MODEL" && "$OMLX_CONFIGURE_MODEL_SETTINGS" != "0" ]]; then
  if ! command -v python3 >/dev/null 2>&1; then
    echo "python3 command not found; cannot configure oMLX model settings." >&2
    echo "Install python3, or set OMLX_CONFIGURE_MODEL_SETTINGS=0 to skip this step." >&2
    exit 1
  fi

  python3 - "$OMLX_BASE_PATH" "$OMLX_PRELOAD_MODEL" <<'PY'
import json
import sys
from pathlib import Path

base_path = Path(sys.argv[1]).expanduser()
model_id = sys.argv[2]
settings_path = base_path / "model_settings.json"
base_path.mkdir(parents=True, exist_ok=True)

if settings_path.exists():
    try:
        data = json.loads(settings_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SystemExit(f"Invalid oMLX model settings JSON: {settings_path}: {exc}")
else:
    data = {}

data.setdefault("version", 1)
models = data.setdefault("models", {})

for existing_model_id, settings in models.items():
    if existing_model_id != model_id and isinstance(settings, dict):
        settings["is_default"] = False

settings = models.setdefault(model_id, {})
settings["is_pinned"] = True
settings["is_default"] = True

tmp_path = settings_path.with_suffix(".tmp")
tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
tmp_path.replace(settings_path)
print(f"Configured oMLX preload model: {model_id}")
print(f"  settings: {settings_path}")
PY
fi

export OMLX_MODEL_DIR
export OMLX_BASE_PATH
export OMLX_PORT
export OMLX_API_KEY

echo "Starting oMLX local AI service"
echo "  model dir: $OMLX_MODEL_DIR"
echo "  base path: $OMLX_BASE_PATH"
echo "  port:      $OMLX_PORT"
echo "  base URL:  http://127.0.0.1:${OMLX_PORT}/v1"
if [[ -n "$OMLX_API_KEY" ]]; then
  echo "  api key:   configured"
else
  echo "  api key:   not configured"
fi
if [[ -n "$OMLX_PRELOAD_MODEL" ]]; then
  echo "  preload:   $OMLX_PRELOAD_MODEL"
fi

omlx_args=(
  serve
  --model-dir "$OMLX_MODEL_DIR"
  --base-path "$OMLX_BASE_PATH"
  --port "$OMLX_PORT"
)

if [[ -n "$OMLX_API_KEY" ]]; then
  omlx_args+=(--api-key "$OMLX_API_KEY")
fi

exec "$OMLX_BIN" "${omlx_args[@]}"
