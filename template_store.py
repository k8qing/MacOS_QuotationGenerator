"""
template_store.py -- saves/loads quotation TEMPLATES (a package title + a list of items) in templates.json.

File format (you can also edit it by hand):
    {"Matrice 4 Thermal Package": {"package_title": "MATRICE 4 THERMAL",
                                   "items": [ {item fields...}, {item fields...} ]}}
Each template stores the items exactly as they were when saved (name, quantity, SRP, optional tick,
description, pictures ...), so a template keeps working even if the catalog CSV changes.
"""
import json
import os
from pathlib import Path


def load_templates(path):
    """Return {name: {"package_title": str, "items": [dict, ...]}} (empty if the file is missing/broken)."""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        templates = data.get("templates", {})
        return {str(name): {"package_title": t.get("package_title", ""), "items": list(t.get("items", []))}
                for name, t in templates.items()}
    except (OSError, ValueError, AttributeError):
        return {}


def save_templates(path, templates):
    """Write atomically, so a crash can never leave a half-written file."""
    path = Path(path)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps({"version": 1, "templates": templates}, indent=2, ensure_ascii=False),
                   encoding="utf-8")
    os.replace(tmp, path)
