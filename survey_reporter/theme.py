# Brand colours
TEAL_DARK  = "#1B5566"   # cover/divider backgrounds, darkest bar segment
TEAL_MID   = "#00B5B8"   # top-box circle, medium bar segment
TEAL_LIGHT = "#7DD8D8"   # lighter bar segment
CORAL      = "#D94438"   # headings, bottom-box circle, strongly-oppose bars
SALMON     = "#F4A89A"   # somewhat-oppose bars
GRAY_MID   = "#BCBEC0"   # unsure bars
GRAY_DARK  = "#58595B"   # body / annotation text
WHITE      = "#FFFFFF"

# Colour sequences for charts
# Stacked bar: index 0 = most positive, last = unsure
LIKERT_COLORS = [TEAL_DARK, TEAL_MID, SALMON, CORAL, GRAY_MID]
BAR_COLOR = TEAL_MID     # simple horizontal bar

# Slide dimensions (inches) — widescreen 13.33 × 7.5
SLIDE_W = 13.33
SLIDE_H = 7.5

# Layout zones as (left, top, width, height) in inches
LOGO_BOX     = (0.27, 0.08, 1.87, 0.42)
PAGE_NUM_BOX = (12.40, 0.10, 0.67, 0.30)
HEADING_BOX  = (0.40, 0.50, 12.53, 0.65)
QUESTION_BOX = (0.40, 1.15, 12.53, 0.38)
CHART_BOX    = (0.40, 1.58, 7.33, 5.30)
ANNOT_BOX    = (8.07, 1.58, 4.87, 5.30)
FOOTER_BOX   = (0.40, 7.10, 12.53, 0.28)

# Stacked-bar circle fallback positions as (cx, cy, r) in inches
# Used when chart_type == "stacked" (dynamic positioning not applicable).
# The bar chart path computes positions dynamically instead.
TOP_CIRCLE_CENTER_R    = (7.33, 2.90, 0.72)   # (cx, cy, radius)
BOTTOM_CIRCLE_CENTER_R = (7.33, 4.80, 0.65)

# Font sizes (points)
FONT_HEADING   = 18
FONT_QUESTION  = 11
FONT_BODY      = 10
FONT_FOOTER    = 8
FONT_CIRCLE    = 20
FONT_CIRCLE_SM = 11
