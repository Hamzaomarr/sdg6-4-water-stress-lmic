"""Revision analysis: single source of truth for every number quoted in the revised manuscript.

Outputs (revision_outputs/):
  revision_numbers.json            all scalar values used in the text
  supp_S1_country_list.xlsx        full country list, region, data availability
  supp_S2_statistical_tests.xlsx   Kruskal-Wallis, Dunn, Wilcoxon, bootstrap CIs
  supp_S3_threshold_and_sensitivity.xlsx  breakpoint model, coupling check, sensitivity
  table2_revised.xlsx, table3_revised.xlsx, table4_revised.xlsx
"""
import os, re, json, itertools
import numpy as np
import pandas as pd
from scipy import stats

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
DATA_PATH = os.path.join(PROJECT_DIR, 'data', 'AQUASTAT_Project1_Data.xlsx')
OUT = os.path.join(PROJECT_DIR, 'revision_outputs')
os.makedirs(OUT, exist_ok=True)
rng = np.random.default_rng(20262611)

src = open(os.path.join(SCRIPT_DIR, '02_generate_tables.py'), encoding='utf-8').read()
ns = {}
exec(re.search(r"short_names = \{.*?\n\}\n", src, re.S).group(0), ns)
exec(re.search(r"region_map = \{.*?\n\}\n", src, re.S).group(0), ns)
SN, RM = ns['short_names'], ns['region_map']


def fmk(v):
    if pd.isna(v): return 'Not evaluated'
    if v < 500: return 'Absolute scarcity (<500)'
    if v < 1000: return 'Chronic scarcity (500-1,000)'
    if v < 1700: return 'Water vulnerability (1,000-1,700)'
    return 'No scarcity (>1,700)'


def tier(v):
    if v > 100: return 'Critical (>100%)'
    if v >= 50: return 'Chronic (50-100%)'
    if v >= 25: return 'Vulnerable (25-50%)'
    return 'Abundant (<25%)'


def boot_ci(x, stat=np.median, n=5000):
    x = np.asarray(x.dropna() if hasattr(x, 'dropna') else x, float)
    b = [stat(rng.choice(x, len(x))) for _ in range(n)]
    return float(np.percentile(b, 2.5)), float(np.percentile(b, 97.5))


def iqr(x): return float(x.quantile(.75) - x.quantile(.25))


df = pd.read_excel(DATA_PATH, sheet_name='Data')
N = {}
N['records_total'] = int(len(df))
N['records_missing_value'] = int(df.Value.isna().sum())
N['areas_in_file'] = int(df[~df.IsAggregate].Area.nunique())
N['aggregate_rows'] = df[df.IsAggregate].Area.unique().tolist()
c = df[~df.IsAggregate].copy()
p = c.pivot(index=['Area', 'Year'], columns='Variable', values='Value').rename(columns=SN).reset_index()
p['Region'] = p.Area.map(RM)
q = p[p.Year == 2023].copy()
valid = q[q.WS_Total.notna()].copy()
N['n_ws'] = int(valid.shape[0])
N['n_wue'] = int(valid.WUE_Total.notna().sum())
N['n_no_data'] = int(q.WS_Total.isna().sum())
N['no_data_countries'] = q[q.WS_Total.isna()].Area.tolist()
N['n_wue_missing_but_ws'] = valid[valid.WUE_Total.isna()].Area.tolist()
reg = valid.Region.value_counts().to_dict()
N['region_n'] = {k: int(v) for k, v in reg.items()}

# ---------- S1 country list
s1 = q[['Area', 'Region', 'WS_Total', 'WUE_Total', 'WUE_Agri', 'Dependency_Ratio']].copy()
s1['In analysis (stress data)'] = s1.WS_Total.notna()
s1['Observations 2021-2023 with stress data'] = s1.Area.map(p.groupby('Area').WS_Total.count())
NOT_UN_LDC_2023 = ['Cameroon', 'Congo', "Côte d'Ivoire", "Democratic People's Republic of Korea", 'Ghana', 'India', 'Kenya', 'Kyrgyzstan', 'Nicaragua', 'Syrian Arab Republic', 'Tajikistan', 'Uzbekistan', 'Viet Nam', 'Zimbabwe']
LDC_NOTE = {a: 'Never an LDC' for a in NOT_UN_LDC_2023}
LDC_NOTE.update({'Vanuatu': 'Graduated December 2020', 'Bhutan': 'LDC in 2021-2023, graduated December 2023', 'Sao Tome and Principe': 'LDC in 2021-2023, graduated December 2024'})
s1['UN LDC status'] = s1.Area.map(LDC_NOTE).fillna('LDC in 2021-2023')
N['not_un_ldc_n'] = int(s1[s1.WS_Total.notna()].Area.isin(NOT_UN_LDC_2023).sum())
s1['Region'] = s1.Region.fillna('No data (Kiribati/Solomon Is./Tuvalu/Vanuatu)')
s1.columns = ['Country', 'Region', 'Water stress 2023 (%)', 'WUE 2023 (USD/m3)', 'Agricultural WUE 2023 (USD/m3)',
              'Dependency ratio (%)', 'In analysis', 'Years with stress data', 'UN LDC status']
s1.to_excel(os.path.join(OUT, 'supp_S1_country_list.xlsx'), index=False)

# ---------- headline 2023
for k, col in [('ws', 'WS_Total'), ('wue', 'WUE_Total'), ('wue_agri', 'WUE_Agri'), ('wue_ind', 'WUE_Industry'),
               ('wue_serv', 'WUE_Services')]:
    x = valid[col].dropna()
    lo, hi = boot_ci(x)
    N[f'{k}_median'] = float(x.median()); N[f'{k}_q1'] = float(x.quantile(.25)); N[f'{k}_q3'] = float(x.quantile(.75))
    N[f'{k}_min'] = float(x.min()); N[f'{k}_max'] = float(x.max()); N[f'{k}_n'] = int(len(x))
    N[f'{k}_ci_lo'] = lo; N[f'{k}_ci_hi'] = hi
N['ws_min_country'] = valid.loc[valid.WS_Total.idxmin(), 'Area']
N['wue_agri_max_country'] = valid.loc[valid.WUE_Agri.idxmax(), 'Area']
N['wue_agri_n_zero'] = int((valid.WUE_Agri <= 0).sum())
N['wue_agri_top5'] = valid.nlargest(5, 'WUE_Agri')[['Area', 'WUE_Agri']].values.tolist()

# stress tiers
valid['Tier'] = valid.WS_Total.apply(tier)
valid['Falk'] = valid.TRWR_per_capita.apply(fmk)
valid['AgriShare'] = valid.WS_Agri_Contrib / valid.WS_Total * 100
tt = []
for t in ['Critical (>100%)', 'Chronic (50-100%)', 'Vulnerable (25-50%)', 'Abundant (<25%)']:
    g = valid[valid.Tier == t]
    tt.append({'Tier': t, 'N': len(g), 'Share of cohort (%)': round(len(g) / len(valid) * 100, 1),
               'Mean stress (%)': round(g.WS_Total.mean(), 2),
               'Median TRWR per capita (m3/person/yr)': round(g.TRWR_per_capita.median(), 0),
               'Mean TRWR per capita (m3/person/yr)': round(g.TRWR_per_capita.mean(), 0),
               'Countries': '; '.join(g.sort_values('WS_Total', ascending=False).Area)})
t4 = pd.DataFrame(tt); t4.to_excel(os.path.join(OUT, 'table4_revised.xlsx'), index=False)
N['tiers'] = tt
N['ag_share_critical'] = valid[valid.Tier == 'Critical (>100%)'][['Area', 'WS_Agri_Contrib', 'AgriShare', 'WS_Muni_Contrib', 'WS_Ind_Contrib']].round(2).values.tolist()
N['ind_contrib_max'] = valid.nlargest(3, 'WS_Ind_Contrib')[['Area', 'WS_Ind_Contrib']].round(2).values.tolist()
N['ind_contrib_over_5'] = int((valid.WS_Ind_Contrib > 5.2).sum())
N['top12'] = valid.nlargest(12, 'WS_Total')[['Area', 'WS_Total']].round(2).values.tolist()
N['tier_lists'] = {t['Tier']: valid[valid.Tier == t['Tier']].sort_values('WS_Total', ascending=False)[['Area', 'WS_Total']].round(2).values.tolist() for t in tt[:3]}
N['n_over_100'] = int((valid.WS_Total > 100).sum())
N['ag_pct_trwr'] = valid.set_index('Area').Ag_Withdrawal_Pct_TRWR.round(2).to_dict()
N['ag_pct_max'] = float(valid.Ag_Withdrawal_Pct_TRWR.max())

# ---------- Table 2 regional (cohort n and medians with IQR)
rows = []
for r, g in valid.groupby('Region'):
    rows.append({'Region': r, 'N (stress)': len(g), 'N (WUE)': int(g.WUE_Total.notna().sum()),
                 **{f'{n} median': g[col].median() for n, col in [('Agri', 'WUE_Agri'), ('Serv', 'WUE_Services'), ('Ind', 'WUE_Industry'), ('Total', 'WUE_Total')]},
                 **{f'{n} IQR': iqr(g[col].dropna()) for n, col in [('Agri', 'WUE_Agri'), ('Serv', 'WUE_Services'), ('Ind', 'WUE_Industry'), ('Total', 'WUE_Total')]},
                 'Stress median': g.WS_Total.median(), 'Stress IQR': iqr(g.WS_Total)})
g = valid
rows.append({'Region': 'All countries', 'N (stress)': len(g), 'N (WUE)': int(g.WUE_Total.notna().sum()),
             **{f'{n} median': g[col].median() for n, col in [('Agri', 'WUE_Agri'), ('Serv', 'WUE_Services'), ('Ind', 'WUE_Industry'), ('Total', 'WUE_Total')]},
             **{f'{n} IQR': iqr(g[col].dropna()) for n, col in [('Agri', 'WUE_Agri'), ('Serv', 'WUE_Services'), ('Ind', 'WUE_Industry'), ('Total', 'WUE_Total')]},
             'Stress median': g.WS_Total.median(), 'Stress IQR': iqr(g.WS_Total)})
t2 = pd.DataFrame(rows); t2.to_excel(os.path.join(OUT, 'table2_revised.xlsx'), index=False)
N['table2'] = t2.round(2).to_dict('records')

# ---------- Table 3: cohort medians per year (same 57 countries each year)
pw = p[p.WS_Total.notna()]
t3 = []
mets = [('Total water-use efficiency (USD/m3)', 'WUE_Total'), ('Irrigated agriculture efficiency (USD/m3)', 'WUE_Agri'),
        ('Services efficiency (USD/m3)', 'WUE_Services'), ('Industry efficiency (USD/m3)', 'WUE_Industry'),
        ('Total water stress (%)', 'WS_Total'), ('Agricultural contribution to stress (%)', 'WS_Agri_Contrib'),
        ('Municipal contribution to stress (%)', 'WS_Muni_Contrib'), ('Industrial contribution to stress (%)', 'WS_Ind_Contrib')]
for lab, col in mets:
    m = {y: pw[pw.Year == y][col].median() for y in (2021, 2022, 2023)}
    cagr = ((m[2023] / m[2021]) ** 0.5 - 1) * 100
    w = pw.pivot(index='Area', columns='Year', values=col)[[2021, 2023]].dropna()
    nchg = int((w[2021] != w[2023]).sum())
    t3.append({'Indicator': lab, '2021': m[2021], '2022': m[2022], '2023': m[2023], 'Absolute change': m[2023] - m[2021],
               'Change (%)': (m[2023] / m[2021] - 1) * 100, 'Annual growth (%)': cagr, 'n countries': len(w), 'n with changed value': nchg})
t3 = pd.DataFrame(t3); t3.to_excel(os.path.join(OUT, 'table3_revised.xlsx'), index=False)
N['table3'] = t3.round(6).to_dict('records')
N['wue_med'] = {y: float(pw[pw.Year == y].WUE_Total.median()) for y in (2021, 2022, 2023)}
N['ws_med'] = {y: float(pw[pw.Year == y].WS_Total.median()) for y in (2021, 2022, 2023)}
N['wue_cagr'] = float(((N['wue_med'][2023] / N['wue_med'][2021]) ** .5 - 1) * 100)

# ---------- regional annualised growth (as in Fig 4B)
rg = {}
for r in ['Sub-Saharan Africa', 'Central & Southern Asia', 'Eastern & South-Eastern Asia']:
    s = pw[pw.Region == r]
    a = s[s.Year == 2021]; b = s[s.Year == 2023]
    e = (b.WUE_Total.median() - a.WUE_Total.median()) / a.WUE_Total.median() * 100 / 2
    st = (b.WS_Total.median() - a.WS_Total.median()) / a.WS_Total.median() * 100 / 2
    rg[r] = {'eff_growth_pct_per_yr': float(e), 'stress_growth_pct_per_yr': float(st)}
N['regional_growth'] = rg

# ---------- update frequency (data-artefact check)
upd = {}
for col in ['WS_Total', 'WUE_Total', 'WUE_Agri', 'TRWR_per_capita', 'Dependency_Ratio']:
    w = p.pivot(index='Area', columns='Year', values=col).dropna()
    upd[col] = {'n': int(len(w)), 'changed_2021_2022': int((w[2021] != w[2022]).sum()),
                'changed_2022_2023': int((w[2022] != w[2023]).sum()), 'changed_2021_2023': int((w[2021] != w[2023]).sum())}
N['update_counts'] = upd
sym = c[c.Variable == 'SDG 6.4.2. Water Stress'].groupby('Year').Symbol.value_counts().unstack(fill_value=0)
N['symbol_counts_ws'] = sym.to_dict()

# ---------- statistics
stat_rows = []
# Wilcoxon 2021 vs 2023
for lab, col in [('Total WUE', 'WUE_Total'), ('Water stress', 'WS_Total'), ('Agricultural WUE', 'WUE_Agri')]:
    w = pw.pivot(index='Area', columns='Year', values=col)[[2021, 2023]].dropna()
    d = w[2023] - w[2021]
    nz = int((d != 0).sum())
    if nz >= 5:
        r = stats.wilcoxon(w[2021], w[2023])
        z = abs(stats.norm.ppf(r.pvalue / 2)); rbs = z / np.sqrt(nz)
        stat_rows.append({'Test': 'Wilcoxon signed-rank 2021 vs 2023', 'Variable': lab, 'n pairs': len(w), 'n non-zero differences': nz,
                          'Statistic': r.statistic, 'p': r.pvalue, 'Effect size (r)': rbs,
                          'Median change': float(d.median())})
    else:
        stat_rows.append({'Test': 'Wilcoxon signed-rank 2021 vs 2023', 'Variable': lab, 'n pairs': len(w), 'n non-zero differences': nz,
                          'Statistic': np.nan, 'p': np.nan, 'Effect size (r)': np.nan, 'Median change': float(d.median())})
# KW across regions (regions with n>=3)
big = [r for r, n in reg.items() if n >= 3]
for lab, col in [('Agricultural WUE', 'WUE_Agri'), ('Industrial WUE', 'WUE_Industry'), ('Services WUE', 'WUE_Services'),
                 ('Total WUE', 'WUE_Total'), ('Water stress', 'WS_Total'), ('Agricultural withdrawal (% TRWR)', 'Ag_Withdrawal_Pct_TRWR')]:
    gs = [valid[valid.Region == r][col].dropna() for r in big]
    h = stats.kruskal(*gs)
    n = sum(len(g) for g in gs); eps = (h.statistic - len(gs) + 1) / (n - len(gs))
    stat_rows.append({'Test': f'Kruskal-Wallis across {len(big)} regions with n>=3', 'Variable': lab, 'n pairs': n, 'n non-zero differences': np.nan,
                      'Statistic': h.statistic, 'p': h.pvalue, 'Effect size (r)': eps, 'Median change': np.nan})
# Dunn (Bonferroni) for variables with significant KW
dunn_rows = []
for lab, col in [('Agricultural WUE', 'WUE_Agri'), ('Total WUE', 'WUE_Total'), ('Water stress', 'WS_Total')]:
    sub = valid[valid.Region.isin(big)][['Region', col]].dropna()
    sub['rk'] = sub[col].rank()
    nn = len(sub); ties = sub[col].value_counts()
    tc = 1 - (ties ** 3 - ties).sum() / (nn ** 3 - nn)
    pairs = list(itertools.combinations(big, 2))
    for a, b in pairs:
        ga = sub[sub.Region == a]; gb = sub[sub.Region == b]
        se = np.sqrt((nn * (nn + 1) / 12) * tc * (1 / len(ga) + 1 / len(gb)))
        z = (ga.rk.mean() - gb.rk.mean()) / se
        pr = 2 * (1 - stats.norm.cdf(abs(z)))
        dunn_rows.append({'Variable': lab, 'Region A': a, 'Region B': b, 'n A': len(ga), 'n B': len(gb), 'z': z, 'p (unadjusted)': pr,
                          'p (Bonferroni)': min(1, pr * len(pairs))})
# bootstrap CI table
ci_rows = []
for lab, col in [('Total WUE', 'WUE_Total'), ('Agricultural WUE', 'WUE_Agri'), ('Industrial WUE', 'WUE_Industry'),
                 ('Services WUE', 'WUE_Services'), ('Water stress', 'WS_Total')]:
    x = valid[col].dropna(); lo, hi = boot_ci(x)
    ci_rows.append({'Variable': lab, 'n': len(x), 'Median': x.median(), 'IQR low': x.quantile(.25), 'IQR high': x.quantile(.75),
                    '95% bootstrap CI low': lo, '95% bootstrap CI high': hi})
for r in valid.Region.unique():
    g = valid[valid.Region == r]
    for lab, col in [('Agricultural WUE', 'WUE_Agri'), ('Water stress', 'WS_Total')]:
        x = g[col].dropna(); lo, hi = boot_ci(x) if len(x) >= 5 else (np.nan, np.nan)
        ci_rows.append({'Variable': f'{lab} - {r}', 'n': len(x), 'Median': x.median(), 'IQR low': x.quantile(.25), 'IQR high': x.quantile(.75),
                        '95% bootstrap CI low': lo, '95% bootstrap CI high': hi})
with pd.ExcelWriter(os.path.join(OUT, 'supp_S2_statistical_tests.xlsx')) as xw:
    pd.DataFrame(stat_rows).to_excel(xw, sheet_name='Omnibus and paired tests', index=False)
    pd.DataFrame(dunn_rows).to_excel(xw, sheet_name='Dunn pairwise', index=False)
    pd.DataFrame(ci_rows).to_excel(xw, sheet_name='Median CIs', index=False)
N['stat_rows'] = pd.DataFrame(stat_rows).replace({np.nan: None}).to_dict('records')
N['dunn_sig'] = [r for r in dunn_rows if r['p (Bonferroni)'] < 0.05]
N['regions_kw'] = big

# ---------- coupling / identity check
d = valid.dropna(subset=['Ag_Withdrawal_Pct_TRWR', 'WS_Agri_Contrib']).copy()
d = d[d.Ag_Withdrawal_Pct_TRWR > 0]
d['implied_EFR_share'] = 1 - d.Ag_Withdrawal_Pct_TRWR / d.WS_Agri_Contrib
N['coupling'] = {'spearman_ag_pct_vs_ws': float(stats.spearmanr(d.Ag_Withdrawal_Pct_TRWR, d.WS_Total).statistic),
                 'spearman_ag_pct_vs_wsagri': float(stats.spearmanr(d.Ag_Withdrawal_Pct_TRWR, d.WS_Agri_Contrib).statistic),
                 'implied_efr_median': float(d.implied_EFR_share.median()), 'implied_efr_min': float(d.implied_EFR_share.min()),
                 'implied_efr_max': float(d.implied_EFR_share.max()), 'n': int(len(d)),
                 'agri_share_median': float(valid.AgriShare.median()), 'agri_share_min': float(valid.AgriShare.min()),
                 'agri_share_max': float(valid.AgriShare.max()),
                 'agri_share_ge_80': int((valid.AgriShare >= 80).sum())}
# ratio of ws to agri pct: ws = totalW/(TRWR-EFR); agri pct = AgriW/TRWR
d['ratio_ws_to_agpct'] = d.WS_Total / d.Ag_Withdrawal_Pct_TRWR
N['coupling']['implied_efr_negative_n'] = int((d.implied_EFR_share < 0).sum())
N['coupling']['ratio_median'] = float(d.ratio_ws_to_agpct.median())

# ---------- breakpoint (segmented regression on log10-log10 scale; breakpoint searched over the 15th-85th percentile of x)
dd = valid.dropna(subset=['Ag_Withdrawal_Pct_TRWR', 'WS_Total']).copy()
dd = dd[(dd.WS_Total > 0) & (dd.Ag_Withdrawal_Pct_TRWR > 0)]
N['n_dropped_log'] = int(valid.shape[0] - dd.shape[0])
N['dropped_log'] = valid[~valid.Area.isin(dd.Area)].Area.tolist()
x = np.log10(dd.Ag_Withdrawal_Pct_TRWR.values); y = np.log10(dd.WS_Total.values)
n = len(x)


def fit_linear(x, y):
    A = np.c_[np.ones_like(x), x]
    b, *_ = np.linalg.lstsq(A, y, rcond=None); r = y - A @ b
    return b, float(r @ r)


def fit_seg(x, y, bp):
    A = np.c_[np.ones_like(x), x, np.maximum(0, x - bp)]
    b, *_ = np.linalg.lstsq(A, y, rcond=None); r = y - A @ b
    return b, float(r @ r)


def best_bp(x, y, n_grid=120):
    grid = np.linspace(np.percentile(x, 15), np.percentile(x, 85), n_grid)
    best = None
    for bp_ in grid:
        b, s_ = fit_seg(x, y, bp_)
        if best is None or s_ < best[2]: best = (bp_, b, s_)
    return best


bp, bsg, ssr1 = best_bp(x, y)
b0, ssr0 = fit_linear(x, y)
sup_F = ((ssr0 - ssr1) / 2) / (ssr1 / (n - 4))
A0 = np.c_[np.ones_like(x), x]; res0 = y - A0 @ b0; fit0 = A0 @ b0
cnt = 0; B = 1000
for _ in range(B):
    yb = fit0 + rng.choice(res0, n)
    bb0, s0 = fit_linear(x, yb); bb = best_bp(x, yb)
    cnt += (((s0 - bb[2]) / 2) / (bb[2] / (n - 4))) >= sup_F
p_sup = (cnt + 1) / (B + 1)
bps = []
for _ in range(1000):
    ix = rng.integers(0, n, n); bps.append(best_bp(x[ix], y[ix])[0])
bps = 10 ** np.array(bps)
# a-priori 25% breakpoint
b25, ssr25 = fit_seg(x, y, np.log10(25))
F25 = ((ssr0 - ssr25) / 2) / (ssr25 / (n - 4))
p25 = float(1 - stats.f.cdf(F25, 2, n - 4))
r2 = lambda ssr: 1 - ssr / np.sum((y - y.mean()) ** 2)
aic = lambda ssr, k: n * np.log(ssr / n) + 2 * k
slope_lin = float(b0[1])
N['threshold'] = {'n': int(n), 'breakpoint_pct_trwr': float(10 ** bp), 'bp_ci_lo': float(np.percentile(bps, 2.5)), 'bp_ci_hi': float(np.percentile(bps, 97.5)),
                  'slope_below': float(bsg[1]), 'slope_above': float(bsg[1] + bsg[2]), 'slope_linear_loglog': slope_lin,
                  'sup_F': float(sup_F), 'p_boot': float(p_sup),
                  'r2_loglog_linear': float(r2(ssr0)), 'r2_segmented': float(r2(ssr1)), 'r2_fixed25': float(r2(ssr25)),
                  'aic_linear': float(aic(ssr0, 2)), 'aic_segmented': float(aic(ssr1, 4)), 'aic_fixed25': float(aic(ssr25, 4)),
                  'F_fixed25': float(F25), 'p_fixed25': p25,
                  'spearman_rho': float(stats.spearmanr(dd.Ag_Withdrawal_Pct_TRWR, dd.WS_Total).statistic)}
# bootstrap CI of the log-log slope
sl = []
for _ in range(2000):
    ix = rng.integers(0, n, n); sl.append(fit_linear(x[ix], y[ix])[0][1])
N['threshold']['slope_ci_lo'] = float(np.percentile(sl, 2.5)); N['threshold']['slope_ci_hi'] = float(np.percentile(sl, 97.5))
# sensitivity: excluding Yemen; excluding transboundary-dependent (>=70%); alt EFR
sens = []


def sens_row(label, sub):
    sub = sub.dropna(subset=['Ag_Withdrawal_Pct_TRWR', 'WS_Total']); sub = sub[sub.WS_Total > 0]
    sub = sub[sub.Ag_Withdrawal_Pct_TRWR > 0]
    xs = np.log10(sub.Ag_Withdrawal_Pct_TRWR.values); ys = np.log10(sub.WS_Total.values)
    r = best_bp(xs, ys); sl_ = fit_linear(xs, ys)[0][1]
    return {'Scenario': label, 'n': len(sub), 'Median stress (%)': sub.WS_Total.median(), 'Median WUE (USD/m3)': sub.WUE_Total.median(),
            'Median agri WUE (USD/m3)': sub.WUE_Agri.median(),
            'Spearman rho (agri % TRWR vs stress)': stats.spearmanr(xs, sub.WS_Total).statistic,
            'Log-log slope': sl_, 'Best breakpoint (% TRWR)': 10 ** r[0] if r else np.nan, 'Countries above 100% stress': int((sub.WS_Total > 100).sum())}


sens.append(sens_row('Main analysis', valid))
sens.append(sens_row('Excluding Yemen', valid[valid.Area != 'Yemen']))
sens.append(sens_row('Excluding the four countries above 100% stress', valid[valid.WS_Total <= 100]))
sens.append(sens_row('Excluding countries with dependency ratio >= 70%', valid[valid.Dependency_Ratio < 70]))
sens.append(sens_row('Excluding the 14 countries never on the UN LDC list (43 UN LDCs only)', valid[~valid.Area.isin(NOT_UN_LDC_2023)]))
# alt EFR: re-derive stress with EFR scaled by 0.5x and 1.5x of the implied EFR share (uses agri contrib identity)
dd2 = valid.dropna(subset=['Ag_Withdrawal_Pct_TRWR', 'WS_Agri_Contrib', 'WS_Total']).copy()
dd2 = dd2[dd2.Ag_Withdrawal_Pct_TRWR > 0]
dd2['efr'] = 1 - dd2.Ag_Withdrawal_Pct_TRWR / dd2.WS_Agri_Contrib
for f in (0.5, 1.5):
    e2 = (dd2.efr * f).clip(0, 0.95)
    dd2[f'ws_{f}'] = dd2.WS_Total * (1 - dd2.efr) / (1 - e2)
    N[f'efr_x{f}'] = {'median_ws': float(dd2[f'ws_{f}'].median()), 'n_over_100': int((dd2[f'ws_{f}'] > 100).sum()),
                       'n_over_50': int((dd2[f'ws_{f}'] >= 50).sum())}
    sens.append({'Scenario': f'EFR share x{f} (re-derived from reported agricultural contribution)', 'n': len(dd2), 'Median stress (%)': dd2[f'ws_{f}'].median(),
                 'Countries above 100% stress': int((dd2[f'ws_{f}'] > 100).sum())})
N['efr_base'] = {'n_over_100': int((dd2.WS_Total > 100).sum()), 'n_over_50': int((dd2.WS_Total >= 50).sum())}
with pd.ExcelWriter(os.path.join(OUT, 'supp_S3_threshold_and_sensitivity.xlsx')) as xw:
    pd.DataFrame([N['threshold']]).T.reset_index().rename(columns={'index': 'Metric', 0: 'Value'}).to_excel(xw, sheet_name='Breakpoint model', index=False)
    pd.DataFrame(sens).to_excel(xw, sheet_name='Sensitivity', index=False)
    pd.DataFrame([N['coupling']]).T.reset_index().rename(columns={'index': 'Metric', 0: 'Value'}).to_excel(xw, sheet_name='Coupling check', index=False)
N['sensitivity'] = pd.DataFrame(sens).replace({np.nan: None}).round(3).to_dict('records')

# Ag intensity classes vs stress (descriptive support for text)
g10 = valid[valid.Ag_Withdrawal_Pct_TRWR < 10]; g25 = valid[valid.Ag_Withdrawal_Pct_TRWR > 25]; g50 = valid[valid.Ag_Withdrawal_Pct_TRWR > 50]
N['desc_thresholds'] = {'lt10_n': len(g10), 'lt10_max_ws': float(g10.WS_Total.max()), 'gt25_n': len(g25), 'gt25_min_ws': float(g25.WS_Total.min()),
                        'gt25_median_ws': float(g25.WS_Total.median()), 'gt50_n': len(g50), 'gt50_min_ws': float(g50.WS_Total.min()) if len(g50) else None,
                        'gt50_countries': g50.Area.tolist(), 'gt25_countries': g25.sort_values('Ag_Withdrawal_Pct_TRWR', ascending=False)[['Area', 'Ag_Withdrawal_Pct_TRWR', 'WS_Total']].round(2).values.tolist()}

# Table 1 (top-12) values
t1 = valid.sort_values('WS_Total', ascending=False).head(12)
N['ws_by_country'] = {k: round(float(v), 2) for k, v in valid.set_index('Area').WS_Total.items()}
N['table1'] = t1[['Area', 'Region', 'WS_Total', 'WS_Agri_Contrib', 'WS_Muni_Contrib', 'WS_Ind_Contrib', 'TRWR_per_capita', 'Falk', 'Dependency_Ratio']].round(2).to_dict('records')
N['trwr_zero'] = valid.nsmallest(3, 'TRWR_per_capita')[['Area', 'TRWR_per_capita']].round(1).values.tolist()
N['dep_over_70'] = valid[valid.Dependency_Ratio >= 70][['Area', 'Dependency_Ratio', 'WS_Total']].round(1).values.tolist()
N['regions_with_n_lt3'] = [r for r, n_ in reg.items() if n_ < 3]

json.dump(N, open(os.path.join(OUT, 'revision_numbers.json'), 'w'), indent=1, default=lambda o: o if not isinstance(o, (np.generic,)) else o.item())
print(json.dumps({k: N[k] for k in ['n_ws', 'n_wue', 'wue_median', 'ws_median', 'wue_agri_median', 'wue_ind_median', 'wue_serv_median', 'wue_cagr', 'threshold', 'coupling']}, indent=1, default=str))
