"""Load and validate YAML annotation files; generate scaffold templates."""
from __future__ import annotations

import textwrap
from pathlib import Path
from typing import Any

import yaml


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_annotations(path: str) -> dict:
    """Load a YAML annotations file and return the parsed dict."""
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        raise ValueError(f"Annotations file {path!r} must be a YAML mapping at the top level.")
    data.setdefault("report", {})
    data.setdefault("slides", [])
    return data


def _common_prefix(names: list[str]) -> str:
    """Return the part of each name up to and including the first underscore, if shared."""
    if not names:
        return ""
    parts = [n.split("_")[0] for n in names]
    return parts[0] if len(set(parts)) == 1 else ""


def _detect_grids(results: list) -> dict[str, list]:
    """Return a mapping of prefix → [results] for battery questions.

    Variables are grouped into a grid when they share a common variable-name
    prefix (text before the first ``_``) AND identical frequency labels.
    Groups with fewer than 3 members are not treated as grids.
    """
    from collections import defaultdict

    # Group first by prefix, then verify shared labels within each prefix group
    by_prefix: dict[str, list] = defaultdict(list)
    for r in results:
        if r.type == "categorical" and r.frequencies and "_" in r.variable:
            prefix = r.variable.split("_")[0]
            by_prefix[prefix].append(r)

    grids: dict[str, list] = {}
    for prefix, group in by_prefix.items():
        if len(group) < 3:
            continue
        # All members must share identical frequency labels
        label_sets = [tuple(f.label for f in r.frequencies) for r in group]
        if len(set(label_sets)) == 1:
            grids[prefix] = group
    return grids


def generate_template(results: list, title: str = "Survey Report") -> dict:
    """Auto-scaffold an annotations dict from a list of QuestionResult objects.

    The returned dict can be serialised to YAML with save_template().
    Variables that form a battery (≥3 sharing identical scale labels and a
    common name prefix) are emitted as a single grid slide.
    """
    from datetime import date

    slides: list[dict] = []

    # Cover
    slides.append({
        "type": "cover",
        "title": title,
        "subtitle": date.today().strftime("%B %Y"),
        "base": f"n={sum(r.n_valid for r in results[:1])} respondents",
    })

    # Detect grid batteries
    grid_map = _detect_grids(results)
    # variable -> prefix (so we can skip individual slides for grid members)
    var_to_grid: dict[str, str] = {}
    for prefix, group in grid_map.items():
        for r in group:
            var_to_grid[r.variable] = prefix
    emitted_grids: set[str] = set()

    for result in results:
        # ── Grid battery ─────────────────────────────────────────────────────
        if result.variable in var_to_grid:
            prefix = var_to_grid[result.variable]
            if prefix in emitted_grids:
                continue  # already emitted
            emitted_grids.add(prefix)
            group = grid_map[prefix]
            freq_labels = [f.label for f in group[0].frequencies]
            n_valid = max(r.n_valid for r in group)
            midpoint = len(freq_labels) // 2
            grid_slide: dict[str, Any] = {
                "type": "grid",
                "heading": prefix.replace("_", " ").title(),
                "question": "(%)",
                "variables": [r.variable for r in group],
                "scale_order": freq_labels,
                "annotations": ["Add commentary here."],
                "base": f"Base: all respondents (n={n_valid})",
            }
            if len(freq_labels) >= 4:
                grid_slide["top_box"] = {
                    "label": "Agree",
                    "values": freq_labels[midpoint:],
                }
                grid_slide["bottom_box"] = {
                    "label": "Disagree",
                    "values": freq_labels[:2],
                }
            slides.append(grid_slide)
            continue

        # ── Individual chart ──────────────────────────────────────────────────
        if result.type == "categorical" and result.frequencies:
            freq_labels = [f.label for f in result.frequencies]
            slide: dict[str, Any] = {
                "type": "chart",
                "variable": result.variable,
                "heading": result.label,
                "question": f"[{result.variable}] (%)",
                "chart_type": "stacked" if len(result.frequencies) >= 4 else "bar",
                "annotations": [
                    "Add subgroup breakdowns here, e.g.:",
                    "– Under 30:  XX%",
                    "– 30–44:     XX%",
                ],
                "base": f"Base: all respondents (n={result.n_valid})",
            }
            if 4 <= len(result.frequencies) <= 6:
                midpoint = len(result.frequencies) // 2
                slide["top_box"] = {
                    "label": "Positive",
                    "values": freq_labels[:midpoint],
                }
                slide["bottom_box"] = {
                    "label": "Negative",
                    "values": freq_labels[midpoint:midpoint + 2],
                }
            slides.append(slide)
        elif result.type == "numeric":
            slides.append({
                "type": "chart",
                "variable": result.variable,
                "heading": result.label,
                "question": f"[{result.variable}] — numeric summary",
                "chart_type": "bar",
                "annotations": ["Add commentary here."],
                "base": f"Base: all respondents (n={result.n_valid})",
            })

    return {
        "report": {
            "title": title,
            "date": date.today().strftime("%B %Y"),
            "base": "All respondents",
        },
        "slides": slides,
    }


def save_template(results: list, output_path: str, title: str = "Survey Report") -> None:
    """Write a scaffold YAML template to *output_path*."""
    data = generate_template(results, title=title)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True, sort_keys=False,
                  default_flow_style=False)
    print(f"Annotations template written to: {output_path}")


def auto_annotations(results: list, title: str = "Survey Report") -> dict:
    """Return a minimal annotations dict (no bullets) for annotation-free runs."""
    return generate_template(results, title=title)
