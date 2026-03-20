"""Generate a synthetic SPSS .sav file for issues-based survey testing.

Simulates a community attitudes survey on an issue of public concern —
the type of research Talbot Mills Research conducts for political and
corporate clients wanting to understand community sentiment and social licence.
"""

import random
import sys
from pathlib import Path

try:
    import pyreadstat
    import pandas as pd
except ImportError:
    print("Install requirements first: pip install -r requirements.txt")
    sys.exit(1)

random.seed(42)
N = 200


def weighted(options, weights):
    return random.choices(options, weights=weights, k=1)[0]


# Build data with realistic skews (not uniform random)
data = {
    "respondent_id": list(range(1, N + 1)),

    # Concern about the issue (skewed toward concerned)
    "issue_concern": [
        weighted([1, 2, 3, 4, 5], [5, 10, 20, 35, 30])
        for _ in range(N)
    ],

    # Agreement: Government is handling this issue well (sceptical skew)
    "govt_handling": [
        weighted([1, 2, 3, 4, 5], [28, 35, 20, 12, 5])
        for _ in range(N)
    ],

    # Grid: Three agreement statements about the issue
    "stmt_community": [
        weighted([1, 2, 3, 4, 5], [5, 10, 18, 35, 32])
        for _ in range(N)
    ],
    "stmt_govt_role": [
        weighted([1, 2, 3, 4, 5], [8, 15, 22, 30, 25])
        for _ in range(N)
    ],
    "stmt_local_action": [
        weighted([1, 2, 3, 4, 5], [6, 12, 20, 33, 29])
        for _ in range(N)
    ],

    # Region (demographic)
    "region": [random.choice([1, 2, 3, 4]) for _ in range(N)],

    # Age group (demographic)
    "age_group": [
        weighted([1, 2, 3, 4, 5], [15, 20, 25, 22, 18])
        for _ in range(N)
    ],

    # Support for action (Yes / No / Unsure) — used for pie chart example
    "support_action": [
        weighted([1, 2, 3], [62, 24, 14])
        for _ in range(N)
    ],
}

df = pd.DataFrame(data)

variable_labels = {
    "respondent_id":    "Respondent ID",
    "support_action":   "Do you support taking action on this issue?",
    "issue_concern":  "How concerned are you about this issue in your community?",
    "govt_handling":  "The government is handling this issue effectively.",
    "stmt_community": "This issue has a direct impact on everyday people in our community.",
    "stmt_govt_role": "Government must take a stronger leadership role on this issue.",
    "stmt_local_action": "Local action can make a meaningful difference on this issue.",
    "region":         "Which region are you based in?",
    "age_group":      "What is your age group?",
}

concern_5 = {
    1: "Not at all concerned",
    2: "Not very concerned",
    3: "Somewhat concerned",
    4: "Very concerned",
    5: "Extremely concerned",
}
agree_5 = {
    1: "Strongly disagree",
    2: "Disagree",
    3: "Neither agree nor disagree",
    4: "Agree",
    5: "Strongly agree",
}

value_labels = {
    "support_action": {1: "Yes", 2: "No", 3: "Unsure"},
    "issue_concern":     concern_5,
    "govt_handling":     agree_5,
    "stmt_community":    agree_5,
    "stmt_govt_role":    agree_5,
    "stmt_local_action": agree_5,
    "region":    {1: "Metro", 2: "Regional", 3: "Rural", 4: "Remote"},
    "age_group": {1: "18–24", 2: "25–34", 3: "35–44", 4: "45–54", 5: "55+"},
}

out = Path(__file__).parent / "sample_survey.sav"
pyreadstat.write_sav(
    df,
    str(out),
    column_labels=variable_labels,
    variable_value_labels=value_labels,
)
print(f"Written: {out}  ({N} rows)")
