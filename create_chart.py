import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Patch
import matplotlib.patches as mpatches

plt.rcParams['font.family'] = 'sans-serif'

mb_planned = 4651.44
mb_achieved = 3842.70
cw_planned = 267.31
cw_achieved = 218.21
total_planned = mb_planned + cw_planned
total_achieved = mb_achieved + cw_achieved

mb_pct = round(mb_achieved/mb_planned*100, 1)
cw_pct = round(cw_achieved/cw_planned*100, 1)
overall_pct = round(total_achieved/total_planned*100, 1)

fig = plt.figure(figsize=(16, 10))
fig.patch.set_facecolor('white')

gs = fig.add_gridspec(3, 3, hspace=0.35, wspace=0.3)

ax_title = fig.add_subplot(gs[0, :])
ax_title.axis('off')
ax_title.text(0.5, 0.7, 'JKD 358 - PRODUCTION BLOCK', fontsize=22, fontweight='bold', color='#1F4E79', ha='center', transform=ax_title.transAxes)
ax_title.text(0.5, 0.35, 'WEEKLY PROGRESS DASHBOARD (21-Feb-2026 to 28-Feb-2026)', fontsize=14, color='#5B5B5B', ha='center', transform=ax_title.transAxes)
ax_title.text(0.5, 0.05, 'Client: JK Defence and Aerospace', fontsize=11, color='#7F7F7F', ha='center', style='italic', transform=ax_title.transAxes)

ax_kpi1 = fig.add_subplot(gs[1, 0])
ax_kpi1.axis('off')
ax_kpi1.add_patch(plt.Rectangle((0.05, 0.1), 0.9, 0.8, fill=True, facecolor='#D6DCE4', edgecolor='#2F5496', linewidth=2, transform=ax_kpi1.transAxes))
ax_kpi1.text(0.5, 0.75, 'TOTAL PLANNED', fontsize=10, color='#5B5B5B', ha='center', transform=ax_kpi1.transAxes, fontweight='bold')
ax_kpi1.text(0.5, 0.45, f'{total_planned:,.0f}', fontsize=26, color='#2F5496', ha='center', transform=ax_kpi1.transAxes, fontweight='bold')
ax_kpi1.text(0.5, 0.15, 'Units', fontsize=9, color='#7F7F7F', ha='center', transform=ax_kpi1.transAxes)

ax_kpi2 = fig.add_subplot(gs[1, 1])
ax_kpi2.axis('off')
ax_kpi2.add_patch(plt.Rectangle((0.05, 0.1), 0.9, 0.8, fill=True, facecolor='#D6DCE4', edgecolor='#2F5496', linewidth=2, transform=ax_kpi2.transAxes))
ax_kpi2.text(0.5, 0.75, 'TOTAL ACHIEVED', fontsize=10, color='#5B5B5B', ha='center', transform=ax_kpi2.transAxes, fontweight='bold')
ax_kpi2.text(0.5, 0.45, f'{total_achieved:,.0f}', fontsize=26, color='#2F5496', ha='center', transform=ax_kpi2.transAxes, fontweight='bold')
ax_kpi2.text(0.5, 0.15, 'Units', fontsize=9, color='#7F7F7F', ha='center', transform=ax_kpi2.transAxes)

ax_kpi3 = fig.add_subplot(gs[1, 2])
ax_kpi3.axis('off')
color = '#70AD47' if overall_pct >= 100 else '#FFC000' if overall_pct >= 80 else '#C00000'
ax_kpi3.add_patch(plt.Rectangle((0.05, 0.1), 0.9, 0.8, fill=True, facecolor=color, edgecolor='#2F5496', linewidth=2, transform=ax_kpi3.transAxes))
ax_kpi3.text(0.5, 0.75, 'OVERALL ACHIEVEMENT', fontsize=10, color='white', ha='center', transform=ax_kpi3.transAxes, fontweight='bold')
ax_kpi3.text(0.5, 0.45, f'{overall_pct}%', fontsize=28, color='white', ha='center', transform=ax_kpi3.transAxes, fontweight='bold')
ax_kpi3.text(0.5, 0.15, 'Completion Rate', fontsize=9, color='white', ha='center', transform=ax_kpi3.transAxes)

ax_bar = fig.add_subplot(gs[2, :2])
categories = ['Main Factory\nBuilding', 'Compound\nWall']
planned = [mb_planned, cw_planned]
achieved = [mb_achieved, cw_achieved]

x = np.arange(len(categories))
width = 0.35

bars1 = ax_bar.bar(x - width/2, planned, width, label='Planned', color='#5B9BD5', edgecolor='white', linewidth=1.5)
bars2 = ax_bar.bar(x + width/2, achieved, width, label='Achieved', color='#70AD47', edgecolor='white', linewidth=1.5)

ax_bar.set_ylabel('Quantity (Units)', fontsize=11, fontweight='bold')
ax_bar.set_title('Planned vs Achieved by Section', fontsize=12, fontweight='bold', color='#1F4E79', pad=10)
ax_bar.set_xticks(x)
ax_bar.set_xticklabels(categories, fontsize=11)
ax_bar.legend(loc='upper right', fontsize=10)
ax_bar.grid(axis='y', alpha=0.3, linestyle='--')
ax_bar.set_ylim(0, max(planned) * 1.15)

for bar in bars1:
    height = bar.get_height()
    ax_bar.annotate(f'{height:,.0f}', xy=(bar.get_x() + bar.get_width()/2, height),
                   xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=9, fontweight='bold')
for bar in bars2:
    height = bar.get_height()
    ax_bar.annotate(f'{height:,.0f}', xy=(bar.get_x() + bar.get_width()/2, height),
                   xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=9, fontweight='bold')

ax_pie = fig.add_subplot(gs[2, 2])
remaining = total_planned - total_achieved
sizes = [total_achieved, remaining]
colors = ['#70AD47', '#D9D9D9']
labels = [f'Achieved\n{total_achieved:,.0f}', f'Pending\n{remaining:,.0f}']
explode = (0.05, 0)

wedges, texts, autotexts = ax_pie.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%',
                                       startangle=90, textprops={'fontsize': 10, 'fontweight': 'bold', 'color': 'white'})
autotexts[0].set_color('white')
autotexts[1].set_color('#5B5B5B')
ax_pie.set_title('Overall Progress', fontsize=12, fontweight='bold', color='#1F4E79', pad=10)

legend_elements = [Patch(facecolor='#5B9BD5', label='Planned'),
                   Patch(facecolor='#70AD47', label='Achieved')]
ax_bar.legend(handles=legend_elements, loc='upper right', fontsize=10)

plt.savefig(r'D:\JKD_Folder\JKD-PROJECT SITE\DPR\Excel_DPR\KPI_Dashboard_Chart.png', dpi=200, bbox_inches='tight', facecolor='white', edgecolor='none')
print("Dashboard chart saved!")
