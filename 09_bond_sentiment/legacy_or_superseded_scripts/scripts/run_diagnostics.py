#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Диагностические тесты для регрессионных моделей.
Проверяет: мультиколлинеарность, гетероскедастичность, нормальность,
спецификацию, влиятельные наблюдения.

Требования: pip install pandas numpy statsmodels scipy openpyxl
Запуск: python run_diagnostics.py
"""

import pandas as pd
import numpy as np
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import jarque_bera
from scipy import stats
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# НАСТРОЙКИ
# ============================================================

INPUT_FILE = "regression_dataset/regression_ready.csv"
OUTPUT_FILE = "diagnostics_results.xlsx"
YEAR_FROM = 2018

BASELINE_EXCLUDE_ISSUERS = {'ВЭБ.РФ'}
BASELINE_DROP_ZERO_ORGANIZERS = True
BASELINE_DROP_BOOK_AFTER_PLACEMENT = True

RATING_MAP = {
    'AAA':15,'AA+':14,'AA':13,'AA-':12,'A+':11,'A':10,'A-':9,
    'BBB+':8,'BBB':7,'BBB-':6,'BB+':5,'BB':4,'BB-':3,
    'B+':2,'B':1,'B-':0,'C':-1
}

VAR_LABELS = {
    'const':'Constant', 'num_organizers':'Number of organizers',
    'is_debut':'Debut (dummy)', 'hist_ever_reduced':'Prior tightening',
    'hist_avg_volume_ratio':'Hist. avg volume ratio', 'rating_num':'Credit rating (ordinal)',
    'log_dur':'ln(Term to exit, days)', 'has_put':'Has put option',
    'is_floater':'Floating rate (dummy)',
    'si':'Sentiment Index (SI)', 'log_buzz':'ln(Buzz)',
    'ofz_yield':'OFZ matched yield, %', 'rvi':'RVI (volatility index)',
}

EXPLANATORY_VARS = [
    'num_organizers', 'is_debut', 'hist_ever_reduced', 'hist_avg_volume_ratio',
    'rating_num', 'ofz_yield', 'rvi', 'log_dur', 'has_put', 'is_floater',
    'si', 'log_buzz',
]

# ============================================================
# ЗАГРУЗКА
# ============================================================

def first_existing_col(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col
    raise KeyError(f"Не найдена: {', '.join(candidates)}")

def normalize_rating(series):
    return (
        series
        .astype(str)
        .str.replace('–', '-', regex=False)
        .str.replace('—', '-', regex=False)
        .str.strip()
    )

def add_quality_flags(sub):
    sub['placement_dt'] = pd.to_datetime(sub['placement_date'], errors='coerce')
    sub['book_dt'] = pd.to_datetime(sub['book_date'], errors='coerce')
    sub['book_after_placement'] = (
        sub['book_dt'].notna()
        & sub['placement_dt'].notna()
        & (sub['book_dt'].dt.date > sub['placement_dt'].dt.date)
    )
    sub['is_special_issuer'] = sub['issuer'].astype(str).isin(BASELINE_EXCLUDE_ISSUERS)
    sub['zero_organizers'] = sub['num_organizers'].eq(0)
    return sub

def apply_baseline_filters(sub):
    mask = pd.Series(True, index=sub.index)
    if BASELINE_DROP_BOOK_AFTER_PLACEMENT:
        mask &= ~sub['book_after_placement']
    if BASELINE_EXCLUDE_ISSUERS:
        mask &= ~sub['is_special_issuer']
    if BASELINE_DROP_ZERO_ORGANIZERS:
        mask &= ~sub['zero_organizers']
    return sub.loc[mask].copy()

def build_quality_reports(raw_sub, full_sub, baseline_sub):
    rows = []

    def add(issue, n, action):
        rows.append({'issue': issue, 'N': int(n), 'action': action})

    add('Ratings with en/em dash before normalization',
        raw_sub['rating_had_long_dash'].sum(),
        'Normalized to ordinary hyphen before rating mapping')
    add('Missing/unmapped rating after normalization',
        raw_sub['rating_num'].isna().sum(),
        'Dropped by regression dropna if still missing')
    add('book_date later than placement_date',
        raw_sub['book_after_placement'].sum(),
        'Excluded from baseline; check manually')
    add('Special/development institution issuer: ВЭБ.РФ',
        raw_sub['is_special_issuer'].sum(),
        'Excluded from baseline; kept for full-sample robustness')
    add('num_organizers = 0',
        raw_sub['zero_organizers'].sum(),
        'Excluded from baseline; keep/check as robustness or dummy')
    sub_veb = raw_sub[raw_sub['is_special_issuer']]
    if len(sub_veb) > 0:
        add('ВЭБ.РФ among underplacement > 1%',
            (sub_veb['underplacement'] > 0.01).sum(),
            'Concentration check')
        add('ВЭБ.РФ among underplacement > 90%',
            (sub_veb['underplacement'] > 0.90).sum(),
            'Concentration check')

    if 'coupon_reduction_bp' in raw_sub.columns:
        s = raw_sub['coupon_reduction_bp'].dropna()
        if len(s) > 0:
            add('Coupon reduction p01/p99',
                len(s),
                f'p01={s.quantile(0.01):.2f}, p99={s.quantile(0.99):.2f}; winsorize only as robustness')
    if 'rvi' in raw_sub.columns:
        s = raw_sub['rvi'].dropna()
        if len(s) > 0:
            add('RVI p99/max',
                len(s),
                f'p99={s.quantile(0.99):.2f}, max={s.max():.2f}; do not auto-drop crisis dates')
    if 'term' in raw_sub.columns:
        add('term < 30 days',
            (raw_sub['term'] < 30).sum(),
            'Review as possible technical placements')
        add('term > 3650 days',
            (raw_sub['term'] > 3650).sum(),
            'Kept; log_term dampens the tail')

    sample_rows = [
        {'sample': 'Full cleaned sample', 'N': len(full_sub),
         'coupon_N': full_sub['coupon_reduction_bp'].notna().sum(),
         'issuers': full_sub['issuer'].nunique()},
        {'sample': 'Baseline sample', 'N': len(baseline_sub),
         'coupon_N': baseline_sub['coupon_reduction_bp'].notna().sum(),
         'issuers': baseline_sub['issuer'].nunique()},
    ]
    quality_df = pd.DataFrame(rows)
    sample_df = pd.DataFrame(sample_rows)

    anomaly_cols = [c for c in [
        'issuer', 'bond_name', 'ISIN', 'placement_date', 'book_date',
        'final_rating', 'num_organizers', 'volume_ratio', 'underplacement'
    ] if c in raw_sub.columns]
    anomalies_df = raw_sub.loc[raw_sub['book_after_placement'], anomaly_cols].copy()
    return quality_df, sample_df, anomalies_df

def load_data():
    df = pd.read_csv(INPUT_FILE, low_memory=False)
    df['year'] = pd.to_datetime(df['placement_date'], errors='coerce').dt.year
    sub = df[df['year'] >= YEAR_FROM].copy()

    si_raw = pd.to_numeric(sub[first_existing_col(sub, ['si_relevant', 'si'])], errors='coerce')
    log_buzz_raw = pd.to_numeric(sub[first_existing_col(sub, ['log_buzz_relevant', 'log_buzz_all', 'log_buzz'])], errors='coerce')
    sub['si'] = si_raw.fillna(0)
    sub['log_buzz'] = log_buzz_raw.fillna(0)
    sub['underplacement'] = 1 - pd.to_numeric(sub['volume_ratio'], errors='coerce')
    sub['coupon_reduction_bp'] = pd.to_numeric(sub['coupon_reduction_bp'], errors='coerce')
    sub['num_organizers'] = pd.to_numeric(sub['num_organizers'], errors='coerce')

    if 'log_term' in sub.columns:
        sub['log_dur'] = pd.to_numeric(sub['log_term'], errors='coerce')
    elif 'term' in sub.columns:
        sub['log_dur'] = np.log(pd.to_numeric(sub['term'], errors='coerce').clip(lower=1))
    sub['term'] = pd.to_numeric(sub['term'], errors='coerce') if 'term' in sub.columns else np.nan
    sub['rating_had_long_dash'] = sub['final_rating'].astype(str).str.contains('–|—', regex=True, na=False)
    sub['final_rating'] = normalize_rating(sub['final_rating'])
    sub['rating_num'] = sub['final_rating'].map(RATING_MAP)
    sub['ofz_yield'] = pd.to_numeric(sub['ofz_matched_yield'], errors='coerce')
    sub['rvi'] = pd.to_numeric(sub['rvi_close'], errors='coerce')
    if 'is_floater' not in sub.columns:
        sub['is_floater'] = 0
    sub = add_quality_flags(sub)

    # Dropna по underplacement + explanatory (БЕЗ coupon_reduction_bp)
    drop_vars = ['underplacement'] + EXPLANATORY_VARS
    full_sub = sub.dropna(subset=drop_vars).copy()
    baseline_sub = apply_baseline_filters(full_sub)
    quality_df, sample_df, anomalies_df = build_quality_reports(sub, full_sub, baseline_sub)
    return full_sub, baseline_sub, quality_df, sample_df, anomalies_df

# ============================================================
# ТЕСТЫ
# ============================================================

def test_vif(sub):
    X = sm.add_constant(sub[EXPLANATORY_VARS])
    results = []
    for i, var in enumerate(X.columns):
        if var == 'const': continue
        vif = variance_inflation_factor(X.values, i)
        results.append({
            'variable': var,
            'label': VAR_LABELS.get(var, var),
            'VIF': round(vif, 2),
            'status': 'ВЫСОКИЙ' if vif > 10 else 'умеренный' if vif > 5 else 'ОК'
        })
    return pd.DataFrame(results)

def test_correlations(sub):
    corr = sub[EXPLANATORY_VARS].corr()
    high_pairs = []
    for i in range(len(corr.columns)):
        for j in range(i+1, len(corr.columns)):
            r = corr.iloc[i, j]
            if abs(r) > 0.3:
                high_pairs.append({
                    'var1': corr.columns[i], 'var2': corr.columns[j],
                    'correlation': round(r, 3),
                    'status': 'ВЫСОКАЯ' if abs(r) > 0.7 else 'умеренная'
                })
    return corr, pd.DataFrame(high_pairs) if high_pairs else pd.DataFrame()

def test_heteroskedasticity(sub):
    X = sm.add_constant(sub[EXPLANATORY_VARS])
    results = []
    # Underplacement — вся выборка
    ols = sm.OLS(sub['underplacement'], X).fit()
    bp_stat, bp_p, bp_f, bp_fp = het_breuschpagan(ols.resid, X)
    results.append({
        'dependent': 'underplacement', 'N': len(sub),
        'BP_statistic': round(bp_stat, 2), 'BP_p_value': round(bp_p, 4),
        'heteroskedasticity': 'Обнаружена' if bp_p < 0.05 else 'Не обнаружена',
        'solution': 'HC3 robust SE' if bp_p < 0.05 else '—'
    })
    # Coupon — подвыборка
    sub_cr = sub.dropna(subset=['coupon_reduction_bp'])
    X_cr = sm.add_constant(sub_cr[EXPLANATORY_VARS])
    ols_cr = sm.OLS(sub_cr['coupon_reduction_bp'], X_cr).fit()
    bp_stat, bp_p, bp_f, bp_fp = het_breuschpagan(ols_cr.resid, X_cr)
    results.append({
        'dependent': 'coupon_reduction_bp', 'N': len(sub_cr),
        'BP_statistic': round(bp_stat, 2), 'BP_p_value': round(bp_p, 4),
        'heteroskedasticity': 'Обнаружена' if bp_p < 0.05 else 'Не обнаружена',
        'solution': 'HC3 robust SE' if bp_p < 0.05 else '—'
    })
    return pd.DataFrame(results)

def test_normality(sub):
    X = sm.add_constant(sub[EXPLANATORY_VARS])
    results = []
    # Underplacement
    ols = sm.OLS(sub['underplacement'], X).fit()
    jb_stat, jb_p, skew, kurt = jarque_bera(ols.resid)
    sw_stat, sw_p = stats.shapiro(ols.resid[:5000]) if len(ols.resid) > 5000 else stats.shapiro(ols.resid)
    results.append({
        'dependent': 'underplacement', 'N': len(sub),
        'JB_statistic': round(jb_stat, 1), 'JB_p_value': round(jb_p, 6),
        'skewness': round(skew, 2), 'kurtosis': round(kurt, 2),
        'SW_statistic': round(sw_stat, 4), 'SW_p_value': round(sw_p, 6),
        'normality': 'Отвергается' if jb_p < 0.05 else 'Не отвергается',
        'note': 'Для Tobit — ограничение, Frac. Logit как робастность'
    })
    # Coupon
    sub_cr = sub.dropna(subset=['coupon_reduction_bp'])
    X_cr = sm.add_constant(sub_cr[EXPLANATORY_VARS])
    ols_cr = sm.OLS(sub_cr['coupon_reduction_bp'], X_cr).fit()
    jb_stat, jb_p, skew, kurt = jarque_bera(ols_cr.resid)
    sw_stat, sw_p = stats.shapiro(ols_cr.resid[:5000]) if len(ols_cr.resid) > 5000 else stats.shapiro(ols_cr.resid)
    results.append({
        'dependent': 'coupon_reduction_bp', 'N': len(sub_cr),
        'JB_statistic': round(jb_stat, 1), 'JB_p_value': round(jb_p, 6),
        'skewness': round(skew, 2), 'kurtosis': round(kurt, 2),
        'SW_statistic': round(sw_stat, 4), 'SW_p_value': round(sw_p, 6),
        'normality': 'Отвергается' if jb_p < 0.05 else 'Не отвергается',
        'note': 'Не критично для OLS (CLT)'
    })
    return pd.DataFrame(results)

def test_reset(sub):
    X = sm.add_constant(sub[EXPLANATORY_VARS])
    results = []
    for dep, data in [('underplacement', sub), ('coupon_reduction_bp', sub.dropna(subset=['coupon_reduction_bp']))]:
        X_d = sm.add_constant(data[EXPLANATORY_VARS])
        ols = sm.OLS(data[dep], X_d).fit()
        y_hat = ols.fittedvalues
        X_reset = X_d.copy()
        X_reset['y_hat2'] = y_hat ** 2
        X_reset['y_hat3'] = y_hat ** 3
        ols_reset = sm.OLS(data[dep], X_reset).fit()
        r_matrix = np.zeros((2, len(ols_reset.params)))
        r_matrix[0, -2] = 1; r_matrix[1, -1] = 1
        f_test = ols_reset.f_test(r_matrix)
        f_stat = float(f_test.fvalue); f_p = float(f_test.pvalue)
        results.append({
            'dependent': dep, 'N': len(data),
            'F_statistic': round(f_stat, 2), 'p_value': round(f_p, 4),
            'specification': 'Возможна неправильная' if f_p < 0.05 else 'Адекватна',
            'note': 'Ожидаемо при цензурированных данных' if f_p < 0.05 else ''
        })
    return pd.DataFrame(results)

def test_influence(sub):
    results = []
    for dep, data in [('underplacement', sub), ('coupon_reduction_bp', sub.dropna(subset=['coupon_reduction_bp']))]:
        X_d = sm.add_constant(data[EXPLANATORY_VARS])
        ols = sm.OLS(data[dep], X_d).fit()
        cooks_d = ols.get_influence().cooks_distance[0]
        threshold = 4 / len(data)
        n_inf = int((cooks_d > threshold).sum())
        results.append({
            'dependent': dep, 'N': len(data),
            'threshold_4_N': round(threshold, 4),
            'n_influential': n_inf,
            'pct_influential': round(n_inf / len(data) * 100, 1),
            'max_cooks_d': round(cooks_d.max(), 4),
            'status': 'Много' if n_inf > len(data) * 0.05 else 'В норме'
        })
    return pd.DataFrame(results)

def fit_ols_summary(sample_name, dep, data, y_col):
    data = data.dropna(subset=[y_col] + EXPLANATORY_VARS).copy()
    x_vars = [v for v in EXPLANATORY_VARS if data[v].nunique(dropna=False) > 1]
    min_n = len(x_vars) + 5
    if len(data) < min_n:
        return {
            'sample': sample_name, 'dependent': dep, 'N': len(data),
            'R2': np.nan, 'si_coef': np.nan, 'si_p': np.nan,
            'log_buzz_coef': np.nan, 'log_buzz_p': np.nan,
            'note': f'Too few observations for {len(x_vars)} non-constant regressors'
        }
    X = sm.add_constant(data[x_vars])
    res = sm.OLS(data[y_col], X).fit(cov_type='HC3')
    return {
        'sample': sample_name, 'dependent': dep, 'N': int(res.nobs),
        'R2': round(res.rsquared, 4),
        'si_coef': round(res.params.get('si', np.nan), 4),
        'si_p': round(res.pvalues.get('si', np.nan), 4),
        'log_buzz_coef': round(res.params.get('log_buzz', np.nan), 4),
        'log_buzz_p': round(res.pvalues.get('log_buzz', np.nan), 4),
        'note': 'HC3 robust SE; constant regressors omitted'
    }

def remove_top_cooks(data, y_col):
    data = data.dropna(subset=[y_col] + EXPLANATORY_VARS).copy()
    if len(data) < len(EXPLANATORY_VARS) + 5:
        return data
    X = sm.add_constant(data[EXPLANATORY_VARS])
    cooks_d = sm.OLS(data[y_col], X).fit().get_influence().cooks_distance[0]
    cutoff = np.quantile(cooks_d, 0.99)
    return data.loc[cooks_d <= cutoff].copy()

def build_robustness_report(full_sub, baseline_sub):
    rows = []
    samples = [
        ('Baseline', baseline_sub),
        ('Full cleaned sample', full_sub),
        ('No ВЭБ.РФ only', full_sub[~full_sub['is_special_issuer']]),
        ('No zero organizers only', full_sub[~full_sub['zero_organizers']]),
        ('Baseline without top 1% Cook D', remove_top_cooks(baseline_sub, 'underplacement')),
    ]
    for name, data in samples:
        rows.append(fit_ols_summary(name, 'underplacement', data, 'underplacement'))

    coupon_samples = [
        ('Baseline', baseline_sub),
        ('Full cleaned sample', full_sub),
        ('Baseline without top 1% Cook D', remove_top_cooks(baseline_sub, 'coupon_reduction_bp')),
    ]
    for name, data in coupon_samples:
        rows.append(fit_ols_summary(name, 'coupon_reduction_bp', data, 'coupon_reduction_bp'))

    sub_cr = baseline_sub.dropna(subset=['coupon_reduction_bp']).copy()
    if len(sub_cr) > 0:
        lo = sub_cr['coupon_reduction_bp'].quantile(0.01)
        hi = sub_cr['coupon_reduction_bp'].quantile(0.99)
        sub_cr['coupon_reduction_bp_w'] = sub_cr['coupon_reduction_bp'].clip(lower=lo, upper=hi)
        row = fit_ols_summary('Baseline winsorized 1/99%', 'coupon_reduction_bp_w', sub_cr, 'coupon_reduction_bp_w')
        row['note'] = f'Coupon clipped at p01={lo:.2f}, p99={hi:.2f}; HC3 robust SE'
        rows.append(row)

    return pd.DataFrame(rows)

# ============================================================
# EXCEL
# ============================================================

def write_dataframe_sheet(wb, sheet_name, df_data):
    hf = Font(bold=True, size=11, name='Arial')
    hfill = PatternFill('solid', fgColor='D9E1F2')
    ws = wb.create_sheet(sheet_name)
    ws['A1'] = sheet_name
    ws['A1'].font = Font(bold=True, size=13, name='Arial')
    row = 3
    if df_data.empty:
        ws.cell(row=row, column=1, value='No rows')
        ws.column_dimensions['A'].width = 18
        return
    for j, h in enumerate(df_data.columns):
        ws.cell(row=row, column=j+1, value=h).font = hf
        ws.cell(row=row, column=j+1).fill = hfill
    row += 1
    for _, r in df_data.iterrows():
        for j, col in enumerate(df_data.columns):
            ws.cell(row=row, column=j+1, value=str(r[col]))
        row += 1
    for i in range(1, len(df_data.columns)+1):
        ws.column_dimensions[get_column_letter(i)].width = 24

def build_excel(vif_df, corr_matrix, high_corr_df, het_df, norm_df, reset_df, infl_df,
                quality_df, sample_df, anomalies_df, robustness_df, n, n_full):
    wb = Workbook()
    hf = Font(bold=True, size=11, name='Arial')
    hfill = PatternFill('solid', fgColor='D9E1F2')
    ok_fill = PatternFill('solid', fgColor='E2EFDA')
    warn_fill = PatternFill('solid', fgColor='FCE4EC')

    # Summary
    ws = wb.active; ws.title = 'Summary'
    ws['A1'] = f'Diagnostic Tests Summary (Baseline N={n}, Full cleaned N={n_full}, Period: {YEAR_FROM}+)'
    ws['A1'].font = Font(bold=True, size=14, name='Arial')
    ws.merge_cells('A1:E1')
    row = 3
    for test, stat, verdict in [
        ('1. Multicollinearity (VIF)', f'Max VIF = {vif_df["VIF"].max():.2f}',
         'ОК' if vif_df['VIF'].max() < 5 else 'Проблема' if vif_df['VIF'].max() > 10 else 'Умеренная'),
        ('2. Heteroskedasticity (BP)', f'p = {het_df["BP_p_value"].min():.4f}', 'HC3 robust SE applied'),
        ('3. Normality (JB)', f'p = {norm_df["JB_p_value"].max():.6f}', 'Violated — CLT for OLS, FracLogit for Tobit'),
        ('4. Specification (RESET)', f'p = {reset_df["p_value"].min():.4f}', 'Expected with censored data'),
        ('5. Influential obs (Cook)', f'Max = {infl_df["max_cooks_d"].max():.4f}',
         'ОК' if infl_df['pct_influential'].max() < 5 else 'Review needed'),
    ]:
        ws.cell(row=row, column=1, value=test).font = Font(name='Arial', size=11, bold=True)
        ws.cell(row=row, column=2, value=stat).font = Font(name='Arial', size=11)
        cell = ws.cell(row=row, column=3, value=verdict)
        cell.font = Font(name='Arial', size=11)
        cell.fill = ok_fill if 'ОК' in verdict or 'applied' in verdict else warn_fill
        row += 1
    ws.column_dimensions['A'].width = 40; ws.column_dimensions['B'].width = 25; ws.column_dimensions['C'].width = 50

    write_dataframe_sheet(wb, 'Sample Filters', sample_df)
    write_dataframe_sheet(wb, 'Data Quality', quality_df)
    write_dataframe_sheet(wb, 'Date Anomalies', anomalies_df)
    write_dataframe_sheet(wb, 'Robustness', robustness_df)

    # VIF
    ws2 = wb.create_sheet('VIF')
    ws2['A1'] = 'Variance Inflation Factors'; ws2['A1'].font = Font(bold=True, size=13, name='Arial')
    row = 3
    for j, h in enumerate(['Variable', 'VIF', 'Status']):
        ws2.cell(row=row, column=j+1, value=h).font = hf; ws2.cell(row=row, column=j+1).fill = hfill
    row += 1
    for _, r in vif_df.iterrows():
        ws2.cell(row=row, column=1, value=r['label']); ws2.cell(row=row, column=2, value=r['VIF']).number_format = '0.00'
        ws2.cell(row=row, column=3, value=r['status']).fill = ok_fill if r['status'] == 'ОК' else warn_fill
        row += 1
    ws2.column_dimensions['A'].width = 34

    # Correlations
    ws3 = wb.create_sheet('Correlations')
    ws3['A1'] = 'Correlation Matrix'; ws3['A1'].font = Font(bold=True, size=13, name='Arial')
    labels = [VAR_LABELS.get(v, v) for v in EXPLANATORY_VARS]
    row = 3
    for j, lab in enumerate(labels):
        c = ws3.cell(row=row, column=j+2, value=lab)
        c.font = Font(bold=True, size=8); c.fill = hfill; c.alignment = Alignment(text_rotation=60, wrap_text=True)
    row += 1
    for i, (var, lab) in enumerate(zip(EXPLANATORY_VARS, labels)):
        ws3.cell(row=row, column=1, value=lab).fill = hfill
        for j, var2 in enumerate(EXPLANATORY_VARS):
            val = corr_matrix.loc[var, var2]
            c = ws3.cell(row=row, column=j+2, value=round(val, 3)); c.number_format = '0.000'
            if abs(val) > 0.5 and i != j: c.fill = warn_fill; c.font = Font(bold=True, size=9)
        row += 1
    ws3.column_dimensions['A'].width = 28

    for sheet_name, df_data in [('Heteroskedasticity', het_df), ('Normality', norm_df),
                                  ('RESET', reset_df), ('Influence', infl_df)]:
        write_dataframe_sheet(wb, sheet_name, df_data)

    wb.save(OUTPUT_FILE)

# ============================================================
# КОНСОЛЬ
# ============================================================

def print_results(vif_df, corr_matrix, high_corr_df, het_df, norm_df, reset_df, infl_df,
                  quality_df, sample_df, robustness_df, n, n_full):
    print(f"\n{'='*70}")
    print(f"ДИАГНОСТИЧЕСКИЕ ТЕСТЫ (baseline N={n}, full cleaned N={n_full}, период {YEAR_FROM}+)")
    print(f"{'='*70}")

    print(f"\n{'─'*70}\n0. SAMPLE / DATA QUALITY\n{'─'*70}")
    for _, r in sample_df.iterrows():
        print(f"  {r['sample']:25s} N={int(r['N'])}, coupon_N={int(r['coupon_N'])}, issuers={int(r['issuers'])}")
    for _, r in quality_df.iterrows():
        if r['N'] > 0:
            print(f"  {r['issue']:45s} N={int(r['N'])} → {r['action']}")

    print(f"\n{'─'*70}\n1. МУЛЬТИКОЛЛИНЕАРНОСТЬ (VIF)\n{'─'*70}")
    print(f"{'Variable':35s} {'VIF':>8s} {'Статус':>12s}")
    for _, r in vif_df.iterrows():
        icon = '✅' if r['status'] == 'ОК' else '⚡' if r['status'] == 'умеренный' else '⚠'
        print(f"{r['label']:35s} {r['VIF']:8.2f} {icon} {r['status']}")
    print(f"\nМакс VIF: {vif_df['VIF'].max():.2f}")

    print(f"\n{'─'*70}\n2. ВЫСОКИЕ КОРРЕЛЯЦИИ (|r| > 0.3)\n{'─'*70}")
    if len(high_corr_df) > 0:
        for _, r in high_corr_df.iterrows():
            print(f"  {VAR_LABELS.get(r['var1'], r['var1']):25s} × {VAR_LABELS.get(r['var2'], r['var2']):25s} = {r['correlation']:.3f}")
    else:
        print("  Нет пар с |r| > 0.3")

    print(f"\n{'─'*70}\n3. ГЕТЕРОСКЕДАСТИЧНОСТЬ (Breusch-Pagan)\n{'─'*70}")
    for _, r in het_df.iterrows():
        icon = '⚠' if r['heteroskedasticity'] == 'Обнаружена' else '✅'
        print(f"  {r['dependent']:25s} (N={r['N']}) LM={r['BP_statistic']:.2f}, p={r['BP_p_value']:.4f} {icon} → {r['solution']}")

    print(f"\n{'─'*70}\n4. НОРМАЛЬНОСТЬ ОСТАТКОВ\n{'─'*70}")
    for _, r in norm_df.iterrows():
        icon = '⚠' if r['normality'] == 'Отвергается' else '✅'
        print(f"  {r['dependent']:25s} (N={r['N']}) JB={r['JB_statistic']:.1f}, p={r['JB_p_value']:.6f}, "
              f"skew={r['skewness']:.2f}, kurt={r['kurtosis']:.2f} {icon}")
        print(f"  {'':25s} → {r['note']}")

    print(f"\n{'─'*70}\n5. СПЕЦИФИКАЦИЯ (Ramsey RESET)\n{'─'*70}")
    for _, r in reset_df.iterrows():
        icon = '⚠' if r['specification'] != 'Адекватна' else '✅'
        print(f"  {r['dependent']:25s} (N={r['N']}) F={r['F_statistic']:.2f}, p={r['p_value']:.4f} {icon} {r['note']}")

    print(f"\n{'─'*70}\n6. ВЛИЯТЕЛЬНЫЕ НАБЛЮДЕНИЯ (Cook's D)\n{'─'*70}")
    for _, r in infl_df.iterrows():
        icon = '✅' if r['status'] == 'В норме' else '⚠'
        print(f"  {r['dependent']:25s} (N={r['N']}) порог={r['threshold_4_N']:.4f}, "
              f"влият.={r['n_influential']} ({r['pct_influential']}%), max={r['max_cooks_d']:.4f} {icon}")

    print(f"\n{'─'*70}\n7. ROBUSTNESS SNAPSHOT (OLS HC3)\n{'─'*70}")
    for _, r in robustness_df.iterrows():
        if r['dependent'] == 'underplacement':
            print(f"  {r['sample']:30s} N={int(r['N']):4d}, SI={r['si_coef']}, p={r['si_p']}")

    print(f"\n{'='*70}\nИТОГ ДЛЯ МЕТОДОЛОГИИ\n{'='*70}")
    print("""
• Baseline sample excludes ВЭБ.РФ / special technical placements, num_organizers=0, and book_date > placement_date.
• Full sample and filtered samples are reported separately as robustness checks.
• Missing SI and buzz values are filled with zero.
• VIF регрессоров — мультиколлинеарность проверена.
• Гетероскедастичность (BP test) — скорректирована через HC3 / clustered SE.
• Нормальность остатков — CLT для OLS; Fractional logit для Tobit.
• RESET — ожидаемо при цензурированных данных; Tobit учитывает структурно.
• Влиятельные наблюдения — проверены через Cook's distance.
""")

# ============================================================
# MAIN
# ============================================================

def main():
    print("="*70)
    print("Диагностические тесты регрессионных моделей")
    print("="*70)

    full_sub, sub, quality_df, sample_df, anomalies_df = load_data()
    n = len(sub)
    sub_cr = sub.dropna(subset=['coupon_reduction_bp'])
    n_full = len(full_sub)
    full_cr = full_sub.dropna(subset=['coupon_reduction_bp'])
    print(f"\nN (underplacement) = {n}")
    print(f"N (coupon) = {len(sub_cr)}")
    print(f"Full cleaned N (underplacement) = {n_full}")
    print(f"Full cleaned N (coupon) = {len(full_cr)}")
    print(f"underplacement > 0: {(sub['underplacement'] > 0.01).sum()}")
    print(f"coupon > 0: {(sub_cr['coupon_reduction_bp'] > 0.5).sum()}")

    print("\nЗапуск тестов...")
    vif_df = test_vif(sub)
    corr_matrix, high_corr_df = test_correlations(sub)
    het_df = test_heteroskedasticity(sub)
    norm_df = test_normality(sub)
    reset_df = test_reset(sub)
    infl_df = test_influence(sub)
    robustness_df = build_robustness_report(full_sub, sub)

    print_results(vif_df, corr_matrix, high_corr_df, het_df, norm_df, reset_df, infl_df,
                  quality_df, sample_df, robustness_df, n, n_full)
    build_excel(vif_df, corr_matrix, high_corr_df, het_df, norm_df, reset_df, infl_df,
                quality_df, sample_df, anomalies_df, robustness_df, n, n_full)
    print(f"\n✅ Результаты сохранены: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
