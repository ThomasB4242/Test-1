"""Auto-generated chart script — edit values/labels and run to regenerate the PNG."""
from __future__ import annotations
import sys, os
# Allow running from any directory
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
from survey_reporter import chart_styles, theme

# ── Size — adjust these and re-run to resize the chart ──────────
FIG_W = 12.65   # inches
FIG_H = 4.3   # inches

cluster_labels = ['New jobs', 'New industries', 'Better environment', 'Strong economy', 'Lower power prices', 'More reliable', 'Better public services']
series = [{'label': '18–29', 'values': [79, 75, 78, 70, 60, 72, 66]}, {'label': '30–44', 'values': [76, 73, 74, 68, 58, 70, 64]}, {'label': '45–59', 'values': [70, 68, 69, 63, 52, 65, 59]}, {'label': '60+', 'values': [63, 62, 65, 57, 47, 58, 53]}]
colors = ['#014E4C', '#027F7C', '#019996', '#00BDBB']
legend_spec = [('18–29', '#014E4C', False), ('30–44', '#027F7C', False), ('45–59', '#019996', False), ('60+', '#00BDBB', False)]

buf = chart_styles.render_cluster_column(cluster_labels=cluster_labels, series=series, colors=colors, fig_h=FIG_H, fig_w=FIG_W, legend_spec=legend_spec)
out = '/home/user/Test-1/sample_data/charts/09_cluster_column_Expected_outcomes_of_more_renewable_energy.png'
with open(out, 'wb') as f:
    f.write(buf.read())
print(f'Saved → {out}')
