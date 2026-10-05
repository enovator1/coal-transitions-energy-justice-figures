"""Rebuilds the manuscript figures as vector PDFs (plus 400-dpi PNG previews for the Word files).

Source of every number: Manuscript_Figure_Data_02-10-2026.xlsx (sheet names in the README sheet).
Run:  python3 make_figures.py <path-to-figure-data.xlsx> <output-folder>

Conventions (Elsevier artwork guidance): vector output, fonts embedded (Type 42), no title inside the
image (titles sit in the Word captions), Arial-metric font, colours from the Chart_Colours sheet.
"""
import sys, os
import numpy as np
import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch
from matplotlib.lines import Line2D
from matplotlib.colors import to_rgb

XLSX = sys.argv[1]
OUT = sys.argv[2]
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    'pdf.fonttype': 42, 'ps.fonttype': 42,
    'font.family': 'Liberation Sans',
    'font.size': 9,
    'axes.linewidth': 0.8,
    'axes.edgecolor': '#222222',
    'text.color': '#111111',
    'axes.labelcolor': '#111111',
    'xtick.color': '#111111', 'ytick.color': '#111111',
    'pdf.compression': 6,
})
INFO = {'Author': '', 'Title': '', 'Creator': 'matplotlib', 'Subject': ''}

MAIN = {'DJS': '#C00000', 'PJS': '#548235', 'RJS': '#1F6FB4', 'RSJ': '#7030A0'}
GREY = '#D9D9D9'
W = 6.0  # inches, as displayed in the Word file

wb = openpyxl.load_workbook(XLSX, data_only=True)


def rows(sheet):
    ws = wb[sheet]
    it = ws.iter_rows(values_only=True)
    head = next(it)
    return head, [r for r in it if r[0] is not None]


def blend(main, t):
    """white -> main colour; t = 1 gives the main colour."""
    m = np.array(to_rgb(main))
    return tuple(1 - t * (1 - m))


def save(fig, name):
    fig.savefig(os.path.join(OUT, name + '.pdf'), metadata=INFO)
    fig.savefig(os.path.join(OUT, name + '.png'), dpi=400)
    plt.close(fig)
    print('saved', name)


# ------------------------------------------------------------------ Figure 1
head, d = rows('Fig1_Year_by_Design_STUDY')
years = [str(r[0]) for r in d]
q = np.array([r[1] for r in d]); mx = np.array([r[2] for r in d]); qn = np.array([r[3] for r in d])
cum = np.array([r[5] for r in d])
assert qn.sum() + q.sum() + mx.sum() == 81 and cum[-1] == 81
C1 = {'Qualitative': '#1C2A44', 'Mixed methods': '#64767D', 'Quantitative': '#C8CCD0'}
fig, ax = plt.subplots(figsize=(W, 3.5))
x = np.arange(len(years))
ax.bar(x, q, 0.62, color=C1['Qualitative'], label='Qualitative')
ax.bar(x, mx, 0.62, bottom=q, color=C1['Mixed methods'], label='Mixed methods')
ax.bar(x, qn, 0.62, bottom=q + mx, color=C1['Quantitative'], label='Quantitative')
ax.set_xticks(x, years)
ax.set_ylim(0, 25); ax.set_yticks(range(0, 26, 5))
ax.set_ylabel('Studies published')
for s in ('top',):
    ax.spines[s].set_visible(False)
ax2 = ax.twinx()
RED = '#B03A2E'
ax2.plot(x, cum, color=RED, lw=2, marker='o', ms=4.5, label='Cumulative total')
ax2.set_ylim(0, 100); ax2.set_yticks(range(0, 101, 20))
ax2.set_ylabel('Cumulative studies', color=RED)
ax2.tick_params(axis='y', colors=RED)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_color(RED)
h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, loc='upper left', frameon=False, fontsize=9, handlelength=1.6)
fig.tight_layout()
save(fig, 'Figure_1')

# ------------------------------------------------------------------ Figure 2
head, d = rows('Fig2_Depth_Distribution_ENTRY')
dims = [('DJS', 'Distributive'), ('PJS', 'Procedural'), ('RJS', 'Recognition'), ('RSJ', 'Restorative')]
d = [r for r in d if isinstance(r[1], (int, float))]
pct = []
for r in d:
    v = list(r[1:8]) + [r[8] or 0]
    v = v[:6] + [v[6] + v[7]]            # fold PJS "7" into Comprehensive (same label)
    assert sum(v) == 95
    pct.append([100 * a / 95 for a in v])
T = [0, 0.135, 0.32, 0.49, 0.667, 0.83, 1.0]      # shade ladder, absent..comprehensive
fig, ax = plt.subplots(figsize=(W, 4.2))
for i, (code, name) in enumerate(dims):
    bottom = 0
    for k in range(7):
        h = pct[i][k]
        if h <= 0:
            continue
        col = GREY if k == 0 else blend(MAIN[code], T[k])
        ax.bar(i, h, 0.58, bottom=bottom, color=col, edgecolor='white', linewidth=0.8)
        bottom += h
ax.set_xticks(range(4), [n for _, n in dims], fontsize=10)
ax.set_ylim(0, 100); ax.set_ylabel('% of coded entries (N = 95)')
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
dim_h = [Rectangle((0, 0), 1, 1, color=MAIN[c]) for c, _ in dims]
fig.legend(dim_h, [n for _, n in dims], title='Dimension (colour)', loc='lower center', ncol=4,
           frameon=False, bbox_to_anchor=(0.5, 0.15), fontsize=8.5, title_fontsize=9, columnspacing=1.6)
names = ['Absent', 'Marginal', 'Limited', 'Moderate', 'Substantial', 'Strong', 'Comprehensive']
gs = ['#D9D9D9', '#BFBFBF', '#A6A6A6', '#8C8C8C', '#737373', '#595959', '#404040']
br_h = [Rectangle((0, 0), 1, 1, color=c) for c in gs]
fig.legend(br_h, names, title='Breadth (shade: light \u2192 dark)', loc='lower center', ncol=4, frameon=False,
           bbox_to_anchor=(0.5, -0.01), fontsize=8.5, title_fontsize=9, columnspacing=1.4)
fig.tight_layout(rect=(0, 0.31, 1, 1))
save(fig, 'Figure_2')

# ------------------------------------------------------------------ Figure 3
head, d = rows('Fig3_Indicator_Coverage_ENTRY')
fig, ax = plt.subplots(figsize=(W, 4.5))
n = len(d)
for i, r in enumerate(d):
    code, ind, label, cnt, p = r[:5]
    col = blend(MAIN[code], 0.28 + 0.72 * p / 100)
    ax.barh(i, p, 0.72, color=col)
    ax.text(p + 1.0, i, f'{p:.1f}%', va='center', ha='left', fontsize=7.5, zorder=5, bbox=dict(fc='white', ec='none', pad=0.6))
ax.set_yticks(range(n), [r[1] for r in d], fontsize=8)
ax.invert_yaxis()
ax.set_xlim(0, 112); ax.set_xticks(range(0, 101, 20))
ax.axvline(50, color='#555555', lw=0.8, ls=(0, (4, 3)), zorder=1)
ax.set_xlabel('% of coded entries (N = 95)')
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
dim_h = [Rectangle((0, 0), 1, 1, color=MAIN[c]) for c, _ in dims]
ax.legend(dim_h, [n_ for _, n_ in dims], title='Dimension', loc='lower right', frameon=False,
          fontsize=8, title_fontsize=8.5, bbox_to_anchor=(1.0, 0.0))
fig.tight_layout()
save(fig, 'Figure_3')

# ------------------------------------------------------------------ Figure 4
head, d = rows('Fig4_Country_ENTRY')
sel = [r for r in d if r[1] >= 2]
assert len(sel) == 9
cols = [('DJS', 2, 3), ('PJS', 4, 5), ('RJS', 6, 7), ('RSJ', 8, 9)]
fig, ax = plt.subplots(figsize=(W, 5.3))
for i, r in enumerate(sel):
    for j, (code, pc, nc) in enumerate(cols):
        p, c = r[pc], r[nc]
        t = 0.2 + 0.8 * p / 100
        ax.add_patch(Rectangle((j, i), 1, 1, facecolor=blend(MAIN[code], t), edgecolor='white', linewidth=2))
        ax.text(j + 0.5, i + 0.5, f'{p:.0f}%\n({c})', ha='center', va='center', fontsize=8.5,
                color='white' if p >= 60 else '#1F2937', linespacing=1.15)
ax.set_xlim(0, 4); ax.set_ylim(len(sel), 0)
ax.xaxis.tick_top()
ax.set_xticks([j + 0.5 for j in range(4)], [n for _, n in dims], fontsize=9)
ax.set_yticks([i + 0.5 for i in range(len(sel))], [f'{r[0]} (n={r[1]})' for r in sel], fontsize=9)
ax.tick_params(length=0)
for s in ax.spines.values():
    s.set_visible(False)
fig.tight_layout()
save(fig, 'Figure_4')

# ------------------------------------------------------------------ Figure S1
BLUE = '#2E5A88'
fig = plt.figure(figsize=(6.5, 4.1))
ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, 100); ax.set_ylim(63, 0.5); ax.axis('off')
FS = 7.8


def box(x0, y0, x1, y1):
    ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, facecolor='white', edgecolor=BLUE, linewidth=1.5))


def side(x0, y0, x1, y1, label):
    ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, facecolor=BLUE, edgecolor=BLUE))
    ax.text((x0 + x1) / 2, (y0 + y1) / 2, label, rotation=90, ha='center', va='center', color='white',
            fontsize=FS + 0.6, fontweight='bold')


def arrow(x0, y0, x1, y1):
    ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle='-|>', mutation_scale=9, color=BLUE, lw=1.5,
                                 shrinkA=0, shrinkB=0))


side(2, 2, 10.8, 31, 'Eligibility')
side(2, 44, 10.8, 61.8, 'Included')
box(12.5, 2.2, 50.7, 17.5)
ax.text(14, 6.8, 'Reports assessed for eligibility', fontsize=FS + 0.8, fontweight='bold', va='center')
ax.text(14, 11.9, 'Total not recorded: records excluded at title/abstract\nscreening were not individually logged (Appendix A4).',
        fontsize=FS - 0.8, style='italic', color='#555555', va='center', linespacing=1.3)
box(53.5, 2.2, 98.3, 31)
ax.text(54.8, 6.0, 'Reports excluded (n = 14)', fontsize=FS + 0.8, fontweight='bold', va='center')
items = ['Not focused on coal or lignite phase-out — 6', 'No codeable country-level unit — 2',
         'Systematic review or literature map — 4', 'Outside European (EU-27 + UK) scope — 1',
         'Source document not verifiable — 1']
for k, t in enumerate(items):
    ax.text(54.8, 9.9 + 3.1 * k, '• ' + t, fontsize=FS, va='center')
ax.text(54.8, 27.6, 'Each exclusion is listed, with the criterion invoked, in Table S2.\nA further 3 duplicate records were removed on DOI, '
        'source-PDF\nchecksum and title.', fontsize=FS - 0.8, style='italic', color='#555555', va='center', linespacing=1.3)
box(12.5, 44, 50.7, 51.4)
ax.text(14, 47.0, 'Studies included in the review', fontsize=FS + 0.8, fontweight='bold', va='center')
ax.text(14, 49.4, '(n = 81 empirical studies)', fontsize=FS, va='center')
box(12.5, 54.0, 50.7, 61.8)
ax.text(14, 56.9, 'Country-level entries coded', fontsize=FS + 0.8, fontweight='bold', va='center')
ax.text(14, 59.7, '(n = 95, across 14 countries)', fontsize=FS, va='center')
arrow(50.7, 13.2, 53.4, 13.2)
arrow(31.5, 17.5, 31.5, 43.9)
arrow(31.5, 51.4, 31.5, 53.9)
save(fig, 'Figure_S1')
