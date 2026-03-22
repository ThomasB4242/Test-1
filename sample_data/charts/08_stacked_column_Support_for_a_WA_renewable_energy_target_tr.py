"""Auto-generated chart script — edit values/labels and run to regenerate the PNG."""
from __future__ import annotations
import sys, os
# Allow running from any directory
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
from survey_reporter import chart_styles, theme

col_labels = ['Aug 2022', 'Nov 2022', 'Feb 2023', 'May 2023', 'Aug 2023']
segments = [{'label': 'Strongly support', 'values': [28.0, 30.0, 29.0, 31.0, 32.0]}, {'label': 'Somewhat support', 'values': [34.0, 33.0, 35.0, 34.0, 34.0]}, {'label': 'Unsure', 'values': [23.0, 21.0, 20.0, 19.0, 18.0]}, {'label': 'Somewhat oppose', 'values': [8.0, 9.0, 9.0, 9.0, 9.0]}, {'label': 'Strongly oppose', 'values': [7.0, 7.0, 7.0, 7.0, 7.0]}]
colors = ['#027F7C', '#9BEFEF', '#B3B3B3', '#F5C5C5', '#A32015']
fig_h = 4.3
fig_w = 9.0
total_percents = [62.0, 63.0, 64.0, 65.0, 66.0]
circle_color = '#00BDBB'
legend_spec = [('Total support', '#00BDBB', True), ('Strongly support', '#027F7C', False), ('Somewhat support', '#9BEFEF', False), ('Unsure', '#B3B3B3', False), ('Somewhat oppose', '#F5C5C5', False), ('Strongly oppose', '#A32015', False)]

buf = chart_styles.render_stacked_column(col_labels=col_labels, segments=segments, colors=colors, fig_h=fig_h, fig_w=fig_w, total_percents=total_percents, circle_color=circle_color, legend_spec=legend_spec)
out = '/home/user/Test-1/sample_data/charts/08_stacked_column_Support_for_a_WA_renewable_energy_target_tr.png'
with open(out, 'wb') as f:
    f.write(buf.read())
print(f'Saved → {out}')
