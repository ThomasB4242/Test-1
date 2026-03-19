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

# Slide dimensions (inches) — standard 10 × 7.5
SLIDE_W = 10.0
SLIDE_H = 7.5

# Layout zones as (left, top, width, height) in inches
LOGO_BOX     = (0.20, 0.08, 1.40, 0.42)
PAGE_NUM_BOX = (9.30, 0.10, 0.50, 0.30)
HEADING_BOX  = (0.30, 0.50, 9.40, 0.65)
QUESTION_BOX = (0.30, 1.15, 9.40, 0.38)
CHART_BOX    = (0.30, 1.58, 5.50, 5.30)
ANNOT_BOX    = (6.05, 1.58, 3.65, 5.30)
FOOTER_BOX   = (0.30, 7.10, 9.40, 0.28)

# Circle overlay positions (left, top, width, height) in inches
TOP_CIRCLE    = (4.55, 2.20, 1.50, 1.50)
BOTTOM_CIRCLE = (4.55, 4.00, 1.20, 1.20)

# Font sizes (points)
FONT_HEADING   = 18
FONT_QUESTION  = 11
FONT_BODY      = 10
FONT_FOOTER    = 8
FONT_CIRCLE    = 20
FONT_CIRCLE_SM = 11
