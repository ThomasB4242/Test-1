"""Generate the WA Energy Survey PPTX report from formatted toplines.

Usage
-----
    python wa_energy_report.py

Reads:
    sample_data/Example - toplines.xlsx  — formatted output toplines
    sample_data/wa_energy_annotations.yaml — slide/annotation spec

Writes:
    sample_data/wa_energy_output.pptx
"""
from __future__ import annotations

from pathlib import Path

from survey_reporter.excel_reader import load_toplines_matrix
from survey_reporter.annotations import load_annotations
from survey_reporter.slide_builder import build_pptx

# ---------------------------------------------------------------------------
# Variable map — tells the parser how to interpret each toplines table
# ---------------------------------------------------------------------------

# Short labels for Q6 (reasons) — original text is ~80 chars
_REASON_LABELS = {
    "Targets create direction and momentum: Provides a clear goal and encourages progress.":
        "Clear goals & direction",
    '"Right thing to do" / general support: Values-based support without a specific reason.':
        '"Right thing to do"',
    "Better environment and health outcomes: Cleaner air, less pollution, and protecting ecosystems.":
        "Better environment & health",
    "Climate change and emissions reduction: Cutting greenhouse gas emissions and addressing climate change.":
        "Climate & emissions reduction",
    "Reduce reliance on fossil fuels: Accelerates the transition away from coal and gas.":
        "Reduce fossil fuel reliance",
    "Accountability and measurement: Enables progress tracking and holds government and industry accountable.":
        "Accountability & measurement",
    "Cost of living and cheaper power over time: Expectation renewables will lower electricity costs.":
        "Cheaper power over time",
    "Future generations: Ensuring a sustainable planet for children and future generations.":
        "Future generations",
    "Conditional support focused on feasibility: Support depends on targets being realistic and affordable.":
        "Conditional / feasibility",
    "Economic opportunity and jobs: Supports new industries, economic growth, and job creation.":
        "Economic opportunity & jobs",
    "Energy security and reliability: Helps ensure stable and reliable energy supply.":
        "Energy security & reliability",
    "Dissatisfaction with politics or delivery: Criticism of government action or speed of delivery.":
        "Dissatisfaction with politics",
    "Certainty for planning and investment: Provides clarity for long-term planning and investment.":
        "Certainty for investment",
    "Unsure": "Unsure",
}

VARIABLE_MAP = {
    # Q1 — Energy types climate impact (grid: one row per energy type)
    "For each of the following, how bad do you think": {
        "type": "grid_rows",
        "n": 1011,
        "question_label": "How bad for the climate?",
        "row_vars": {
            "Coal":                           "energy_climate_coal",
            "Oil, including diesel and petrol": "energy_climate_oil",
            "Gas":                            "energy_climate_gas",
        },
    },

    # Q2 — Main energy concern (single choice, 8 options)
    "Which of the following concerns you most": {
        "type": "single",
        "variable": "energy_concern",
        "n": 1011,
        "question_label": "Which concerns you most about WA's energy future?",
    },

    # Q3 — Renewable energy outcomes likelihood (grid: one row per outcome)
    "And how likely do you think these following outcomes": {
        "type": "grid_rows",
        "n": 1011,
        "question_label": "How likely if WA used more renewable energy?",
        "row_vars": {
            "Providing good jobs":                               "re_outcome_jobs",
            "Maintaining living standards":                      "re_outcome_living_std",
            "A strong economy":                                  "re_outcome_economy",
            "New industries":                                    "re_outcome_industries",
            "Lowering pollution / addressing climate change":    "re_outcome_pollution",
            "Reducing the cost of living":                       "re_outcome_cost_living",
            "Lower power prices":                                "re_outcome_power_prices",
            "Funding public services":                           "re_outcome_public_svcs",
            "Contributing to a national climate resilience fund":"re_outcome_climate_fund",
            "Well-planned renewable energy infrastructure":      "re_outcome_infrastructure",
            "Paying off debt (balancing the federal budget)":    "re_outcome_debt",
            "Better energy system reliability":                  "re_outcome_reliability",
            "More certainty for businesses":                     "re_outcome_businesses",
        },
    },

    # Q4 — Solar consideration (pie: Yes / No / Unsure, base n=473)
    "Are you considering installing rooftop solar": {
        "type": "single",
        "variable": "solar_consider",
        "n": 473,
        "question_label": "Are you considering installing rooftop solar?",
    },

    # Q5 — RE target support (bar: 5-point scale; exclude pre-computed totals)
    "There has been discussion about the WA government setting a target": {
        "type": "single",
        "variable": "re_target_support",
        "n": 1011,
        "question_label": "Support for WA adopting a renewable energy target",
        "exclude_rows": ["TOTAL SUPPORT", "TOTAL OPPOSE"],
    },

    # Q6 — Reasons for supporting target (coded open-end, base n=626)
    "What is the main reason for supporting a target": {
        "type": "single",
        "variable": "re_target_reasons",
        "n": 626,
        "question_label": "Main reason for supporting a renewable energy target",
    },
}


# ---------------------------------------------------------------------------
# Build the report
# ---------------------------------------------------------------------------

def main():
    base_dir   = Path(__file__).parent
    toplines   = base_dir / "sample_data" / "Example - toplines.xlsx"
    annot_file = base_dir / "sample_data" / "wa_energy_annotations.yaml"
    output     = base_dir / "sample_data" / "wa_energy_output.pptx"
    template   = base_dir / "Blank.potm"

    print(f"Reading toplines : {toplines}")
    results = load_toplines_matrix(toplines, VARIABLE_MAP)
    print(f"  → {len(results)} QuestionResult objects")
    for r in results:
        print(f"     {r.variable:35s}  {len(r.frequencies)} options  n={r.n_valid}")

    print(f"Loading annotations: {annot_file}")
    annotations = load_annotations(str(annot_file))

    print(f"Building PPTX → {output}")
    build_pptx(
        results,
        annotations,
        str(output),
        template_path=str(template),
    )
    print(f"Done! → {output}")


if __name__ == "__main__":
    main()
