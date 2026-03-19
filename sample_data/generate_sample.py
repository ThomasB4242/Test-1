"""Generate a small synthetic SPSS .sav file for testing."""

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

data = {
    "respondent_id": list(range(1, N + 1)),
    "overall_satisfaction": [random.randint(1, 5) for _ in range(N)],
    "recommend_score": [random.randint(0, 10) for _ in range(N)],
    "product_quality": [random.choice([1, 2, 3, 4, 5]) for _ in range(N)],
    "customer_service": [random.choice([1, 2, 3, 4, 5]) for _ in range(N)],
    "value_for_money": [random.choice([1, 2, 3, 4, 5]) for _ in range(N)],
    "region": [random.choice([1, 2, 3, 4]) for _ in range(N)],
    "age_group": [random.choice([1, 2, 3, 4, 5]) for _ in range(N)],
}

df = pd.DataFrame(data)

variable_labels = {
    "respondent_id": "Respondent ID",
    "overall_satisfaction": "Overall, how satisfied are you with our service?",
    "recommend_score": "How likely are you to recommend us? (0-10)",
    "product_quality": "Please rate the quality of our product.",
    "customer_service": "Please rate your experience with customer service.",
    "value_for_money": "Please rate value for money.",
    "region": "Which region are you based in?",
    "age_group": "What is your age group?",
}

likert_5 = {1: "Very dissatisfied", 2: "Dissatisfied", 3: "Neutral", 4: "Satisfied", 5: "Very satisfied"}
quality_5 = {1: "Very poor", 2: "Poor", 3: "Average", 4: "Good", 5: "Excellent"}

value_labels = {
    "overall_satisfaction": likert_5,
    "product_quality": quality_5,
    "customer_service": quality_5,
    "value_for_money": quality_5,
    "region": {1: "North", 2: "South", 3: "East", 4: "West"},
    "age_group": {1: "18-24", 2: "25-34", 3: "35-44", 4: "45-54", 5: "55+"},
}

out = Path(__file__).parent / "sample_survey.sav"
pyreadstat.write_sav(
    df,
    str(out),
    column_labels=variable_labels,
    variable_value_labels=value_labels,
)
print(f"Written: {out}  ({N} rows)")
