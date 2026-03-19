# Brand colours — matched to Blank.potm theme
TEAL_DARK  = "#027F7C"   # accent1 — cover/divider backgrounds, darkest bar segment
TEAL_MID   = "#00BDBB"   # accent2 — top-box circle, medium bar segment
TEAL_LIGHT = "#9BEFEF"   # accent3 — lighter bar segment
CORAL      = "#A32015"   # accent6 — headings, bottom-box circle, strongly-oppose bars
SALMON     = "#FA8488"   # accent5 — somewhat-oppose bars
GRAY_MID   = "#BCBEC0"   # unsure bars
GRAY_DARK  = "#44546A"   # dk2 — body / annotation text
WHITE      = "#FFFFFF"

# Template font (DM Sans from Blank.potm theme)
FONT_FACE  = "DM Sans"

# Colour sequences for charts
# Stacked bar: index 0 = most positive, last = unsure
LIKERT_COLORS = [TEAL_DARK, TEAL_MID, SALMON, CORAL, GRAY_MID]
BAR_COLOR = TEAL_MID     # simple horizontal bar

# Slide dimensions (inches) — widescreen 13.33 × 7.5
SLIDE_W = 13.33
SLIDE_H = 7.5

# Layout zones as (left, top, width, height) in inches
# Heading/question sit in the clean white area above the thin rule (y≈1.10)
# Master logo occupies (0.21", 6.85") 0.86"×0.40" — footer stays above it
LOGO_BOX     = (0.27, 0.08, 1.87, 0.42)
PAGE_NUM_BOX = (12.40, 6.80, 0.67, 0.30)
HEADING_BOX  = (0.34, 0.22, 12.65, 0.82)
QUESTION_BOX = (0.34, 1.14, 9.00, 0.38)
CHART_BOX    = (0.34, 1.58, 7.33, 5.10)
ANNOT_BOX    = (8.07, 1.58, 4.87, 5.10)
FOOTER_BOX   = (1.20, 6.82, 11.00, 0.28)

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
