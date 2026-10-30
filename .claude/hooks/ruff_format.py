"""PostToolUse hook: run ruff fix + format on edited .py/.ipynb files."""

import json
import subprocess
import sys

data = json.load(sys.stdin)
tool_input = data.get("tool_input", {})
path = tool_input.get("file_path") or tool_input.get("notebook_path") or ""
if path.endswith((".py", ".ipynb")):
    subprocess.run([sys.executable, "-m", "ruff", "check", "--fix", "--quiet", path])
    subprocess.run([sys.executable, "-m", "ruff", "format", "--quiet", path])
