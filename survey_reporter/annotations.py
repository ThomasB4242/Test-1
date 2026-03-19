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


def generate_template(results: list, title: str = "Survey Report") -> dict:
    """Auto-scaffold an annotations dict from a list of QuestionResult objects.

    The returned dict can be serialised to YAML with save_template().
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

    # One chart slide per question
    for result in results:
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
            # Suggest top_box / bottom_box for 4-5 option questions
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
                "annotations": [
                    "Add commentary here.",
                ],
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
