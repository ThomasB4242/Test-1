"""Auto-generated chart script — edit values/labels and run to regenerate the PNG."""
from __future__ import annotations
import sys, os
# Allow running from any directory
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
from survey_reporter import chart_styles, theme

cluster_labels = ['New jobs', 'New industries', 'Better environment', 'Strong economy', 'Lower power prices', 'More reliable', 'Better public services']
series = [{'label': '18–29', 'values': [79, 75, 78, 70, 60, 72, 66]}, {'label': '30–44', 'values': [76, 73, 74, 68, 58, 70, 64]}, {'label': '45–59', 'values': [70, 68, 69, 63, 52, 65, 59]}, {'label': '60+', 'values': [63, 62, 65, 57, 47, 58, 53]}]
colors = ['#027F7C', '#00BDBB', '#9BEFEF', '#C8F4F3']
fig_h = 4.3
fig_w = 9.0
legend_spec = [('18–29', '#027F7C', False), ('30–44', '#00BDBB', False), ('45–59', '#9BEFEF', False), ('60+', '#C8F4F3', False)]

buf = chart_styles.render_cluster_column(cluster_labels=cluster_labels, series=series, colors=colors, fig_h=fig_h, fig_w=fig_w, legend_spec=legend_spec)
out = '/home/user/Test-1/sample_data/charts/09_cluster_column_Expected_outcomes_of_more_renewable_energy.png'
with open(out, 'wb') as f:
    f.write(buf.read())
print(f'Saved → {out}')
