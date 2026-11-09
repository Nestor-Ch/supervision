"""
check_sheet.py - validate inputs/dashboard_data/sheet_classification.xlsx against the
questionnaires, the data tables and the dashboard source, and report problems.

Run it every time the sheet (or the data) changes, before building the dashboard:

    python check_sheet.py

Three kinds of finding:
  ERROR  the sheet is inconsistent (unknown word, duplicate card order, chart that cannot
         draw that question kind, a page title that differs from the heading in index.qmd...).
         Fix the sheet.
  DATA   something is wrong in the DATA (a label such as "None ", a doubled "Вчителі:" prefix,
         a table older than its DAF, counts that disagree with the raw export...).
         Fix it at the source (pipeline / questionnaire / export). The dashboard deliberately
         does NOT work around data problems - it would hide them.
  INFO   worth knowing, nothing to fix.

Exit status: 0 = clean, 1 = at least one ERROR or DATA finding.
"""
import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parent
INPUTS = ROOT / 'inputs'
SHEET = INPUTS / 'dashboard_data' / 'sheet_classification.xlsx'

findings = []  # (level, where, message)


def add(level, where, msg):
    findings.append((level, where, msg))


def main():
    sheets = pd.read_excel(SHEET, sheet_name=None, dtype=str)
    for need in ('survey', 'main', 'breakdowns', 'lists'):
        if need not in sheets:
            add('ERROR', 'workbook', f'missing sheet "{need}"')
    if any(f[0] == 'ERROR' for f in findings):
        return report()
    sv, mn, bk, lists = sheets['survey'], sheets['main'], sheets['breakdowns'], sheets['lists']

    # ------------------------------------------------------------------ vocabulary
    vocab = {c: set(lists[c].dropna()) for c in lists.columns}
    vocab_cols = {  # sheet.column -> list name
        ('survey', 'Graph_type'): 'Graph_type', ('survey', 'page_0_graph_type'): 'page_0_graph_type',
        ('survey', 'Page_gen_type_graph'): 'Page_gen_type_graph',
        ('survey', 'Page_graph_disag_chosen'): 'Page_graph_disag_chosen',
        ('survey', 'Orientation_overall'): 'Orientation_overall', ('survey', 'Layout_hint'): 'Layout_hint',
        ('survey', 'Multi_select'): 'yes_no', ('survey', 'Hidden_when_breakdown'): 'yes_no',
        ('survey', 'asked_of'): 'asked_of', ('main', 'page_type'): 'page_type',
        ('main', 'value_toggle'): 'yes_no', ('breakdowns', 'kind'): 'breakdown_kind'}
    frames = {'survey': sv, 'main': mn, 'breakdowns': bk}
    for (sh, col), lst in vocab_cols.items():
        if col not in frames[sh].columns:
            add('ERROR', f'{sh}.{col}', 'column is missing'); continue
        for i, v in frames[sh][col].items():
            if pd.notna(v) and v not in vocab[lst]:
                add('ERROR', f'{sh}.{col} row {i + 2}', f'"{v}" is not in the vocabulary {sorted(vocab[lst])}')

    # ------------------------------------------------------------------ pages / order
    def pages_of(s):
        return [int(x) for x in re.split(r'[;,]', s) if x.strip() != ''] if isinstance(s, str) else []
    page_ids = sorted(int(p) for p in mn['page'])
    sv['_pages'] = sv['pages'].map(pages_of)
    for i, r in sv.iterrows():
        if not r['_pages']:
            add('ERROR', f'survey row {i + 2}', f'{r["name"]}: "pages" is empty or not like "0;2"')
        for p in r['_pages']:
            if p not in page_ids:
                add('ERROR', f'survey row {i + 2}', f'{r["name"]}: page {p} is not in the main sheet')
    for p in page_ids:
        cards = sv[sv['_pages'].map(lambda L: p in L)]
        if p == 0 or mn.loc[mn['page'] == str(p), 'page_type'].iloc[0] == 'text':
            continue
        orders = pd.to_numeric(cards['Card_order'], errors='coerce')
        if orders.isna().any():
            add('ERROR', f'page {p}', 'cards without Card_order: ' + ', '.join(cards.loc[orders.isna(), 'name']))
        o = orders.dropna().astype(int).tolist()
        if len(set(o)) != len(o):
            add('ERROR', f'page {p}', f'duplicate Card_order values {sorted(o)}')
        elif sorted(o) != list(range(1, len(o) + 1)):
            add('ERROR', f'page {p}', f'Card_order must run 1..{len(o)} without gaps, found {sorted(o)}')
        if cards.empty:
            add('ERROR', f'page {p}', 'page has no cards in the survey sheet')

    # ------------------------------------------------------------------ questionnaires
    kobo = pd.read_excel(INPUTS / 'supervision_survey_educators_2.xlsx', sheet_name='survey')
    hk = pd.read_excel(INPUTS / 'supervision_survey_educators_2_harmonised.xlsx', sheet_name='survey')
    ktype = dict(zip(pd.concat([kobo.name, hk.name]), pd.concat([kobo.type, hk.type]).astype(str)))
    for i, r in sv.iterrows():
        n = r['name']
        if n not in ktype:
            add('ERROR', f'survey row {i + 2}', f'{n} is not in either questionnaire'); continue
        kmulti = ktype[n].startswith('select_multiple')
        if (r.get('Multi_select') == 'yes') != kmulti:
            add('ERROR', f'survey {n}', f'Multi_select={r.get("Multi_select")} but the questionnaire says "{ktype[n]}"')
        gen, dis = r.get('Page_gen_type_graph'), r.get('Page_graph_disag_chosen')
        if kmulti and gen in ('donut', 'stacked_bar'):
            add('ERROR', f'survey {n}', f'multi-select question cannot be drawn as {gen} (shares sum above 100 %)')
        if kmulti and dis == 'stacked_bar':
            add('ERROR', f'survey {n}', 'multi-select question cannot be drawn as a stacked bar with a breakdown')
        if (gen == 'kpi_card' or dis == 'mean_bar') and 'scale_1_5' not in ktype[n]:
            add('ERROR', f'survey {n}', f'kpi_card/mean_bar needs a 1-5 rating question, found "{ktype[n]}"')
        if r.get('Hidden_when_breakdown') == 'yes' and not str(r.get('Graph_type')).startswith('static'):
            add('ERROR', f'survey {n}', 'Hidden_when_breakdown=yes requires Graph_type=static')
        if pd.notna(r.get('Orientation_overall')) and gen not in ('bar', 'stacked_bar', 'grouped_bar'):
            add('ERROR', f'survey {n}', f'Orientation_overall is only meaningful for bar charts, chart is {gen}')

    # ------------------------------------------------------------------ main vs dashboard source
    qmd = (ROOT / 'index.qmd').read_text(encoding='utf-8')
    headings, fence = [], False
    for line in qmd.splitlines():
        if line.startswith('```'):
            fence = not fence
        elif not fence and line.startswith('# '):
            headings.append(line[2:].strip())
    nav = (ROOT / 'nav-icons.html').read_text(encoding='utf-8')
    for _, r in mn.iterrows():
        if r['title'] not in headings:
            add('ERROR', f'main page {r["page"]}', f'title "{r["title"]}" is not a "# " heading in index.qmd (headings: {headings})')
        if f'"{r["title"]}"' not in nav:
            add('ERROR', f'main page {r["page"]}', f'title "{r["title"]}" has no entry in nav-icons.html (icon would be lost)')
        if isinstance(r['icon'], str) and f'"{r["icon"]}"' not in nav:
            add('ERROR', f'main page {r["page"]}', f'icon {r["icon"]} differs from nav-icons.html')
        if pd.notna(r.get('breakdowns')):
            for b in r['breakdowns'].split(';'):
                if b not in set(bk['code']):
                    add('ERROR', f'main page {r["page"]}', f'breakdown "{b}" is not in the breakdowns sheet')
        if pd.notna(r.get('min_n')) and not str(r['min_n']).replace('.', '', 1).isdigit():
            add('ERROR', f'main page {r["page"]}', f'min_n "{r["min_n"]}" is not a number')
    for h in headings:
        if h not in set(mn['title']):
            add('ERROR', 'index.qmd', f'heading "{h}" is not a page title in the main sheet')

    # ------------------------------------------------------------------ data tables
    tables = {'clean_df.parquet': pd.read_parquet(INPUTS / 'dashboard_data' / 'clean_df.parquet'),
              'clean_df_harmonized.parquet': pd.read_parquet(INPUTS / 'dashboard_data' / 'clean_df_harmonized.parquet')}
    # respondent-level exports are private (not in the public repo): optional cross-check only
    rf, hf = INPUTS / 'supervision_df_02_11.xlsx', INPUTS / 'supervision_df_02_11_harmonised.xlsx'
    raw = pd.read_excel(rf) if rf.exists() else None
    har = pd.read_excel(hf) if hf.exists() else None

    for _, r in mn.iterrows():                                   # freshness
        if pd.notna(r.get('daf_file')) and pd.notna(r.get('file_used')) and r['file_used'] in tables:
            daf = INPUTS / r['daf_file']; par = INPUTS / 'dashboard_data' / r['file_used']
            if daf.exists() and par.stat().st_mtime < daf.stat().st_mtime:
                add('DATA', f'page {r["page"]}', f'{r["file_used"]} is OLDER than its DAF {r["daf_file"]} - regenerate the table')

    seen = set()
    var_page = {}
    for _, r in sv.iterrows():
        for p in r['_pages']:
            if p >= 1:
                var_page[r['name']] = p
    for i, r in sv.iterrows():
        n = r['name']
        pg = [p for p in r['_pages']]
        tbl = tables['clean_df_harmonized.parquet' if n.startswith('h_') else 'clean_df.parquet']
        d = tbl[tbl.variable_orig == n]
        if d.empty:
            add('ERROR', f'survey {n}', 'variable is not in its data table'); continue
        # label text: one per variable, audience prefix present exactly once
        labels = d.groupby('variable').admin.agg(lambda s: sorted(set(s)))
        if len(labels) > 1:
            add('DATA', f'{n}', 'several different question labels in the table: ' +
                ' | '.join(f'"{k[:60]}" ({",".join(v)})' for k, v in labels.items()))
        lab = d[d.admin == 'Overall'].variable.iloc[0]
        asked = r.get('asked_of')
        want = 'Вчителі: ' if str(asked).startswith('teachers') else 'Адміністрація: ' if str(asked).startswith('leaders') else None
        if want and not lab.startswith(want):
            add('DATA', f'{n}', f'label should start with "{want}" (asked of {asked}): "{lab[:70]}"')
        for pref in ('Вчителі: ', 'Адміністрація: '):
            if lab.count(pref) > 1:
                add('DATA', f'{n}', f'doubled audience prefix in the label: "{lab[:70]}"')
        # option labels
        opts = d.drop_duplicates('option_orig')[['option_orig', 'option']]
        for _, o in opts.iterrows():
            t = str(o['option'])
            if t.strip().lower() in ('none', 'nan', '') or t != t.strip():
                add('DATA', f'{n}', f'bad option label "{t}" for code "{o["option_orig"]}" (questionnaire label missing or untrimmed)')
            elif re.fullmatch(r'\d+\.0', t) and (n, 'float') not in seen:
                seen.add((n, 'float'))
                add('INFO', f'{n}', f'numeric option labels are stored as floats ("{t}"); the dashboard shows them without ".0"')
        # duplicates
        key = ['ID', 'admin', 'admin_category', 'option_orig'] + (['disaggregations_category_1_orig'] if 'disaggregations_category_1_orig' in d.columns else [])
        if d.duplicated(key).any():
            add('DATA', f'{n}', f'{int(d.duplicated(key).sum())} duplicated rows in the table')
        # shares
        ov = d[d.admin == 'Overall']
        grp = ov.groupby('disaggregations_category_1_orig') if 'disaggregations_category_1_orig' in ov.columns and n.startswith('h_') else [(None, ov)]
        for gname, gd in grp:
            s = gd.perc.sum()
            kmulti = ktype.get(n, '').startswith('select_multiple')
            if not kmulti and abs(s - 1) > 0.01:
                add('DATA', f'{n}', f'single-choice shares sum to {s:.3f}' + (f' in group {gname}' if gname else ''))
        # expected_n
        exp = pd.to_numeric(r.get('expected_n'), errors='coerce')
        if pd.notna(exp):
            if n.startswith('h_') and n not in ('h_population',):
                got = int(ov.groupby('disaggregations_category_1_orig').general_count.first().sum()) if n not in ('h_oblast', 'h_experience') else None
            else:
                got = int(ov.general_count.iloc[0]) if n != 'h_population' else int(ov.general_count.iloc[0])
            if n in ('h_oblast', 'h_experience'):
                got = int(ov.groupby('disaggregations_category_1_orig').general_count.first().sum())
            if got is not None and got != int(exp):
                add('DATA', f'{n}', f'table says {got} respondents, the sheet expects {int(exp)}')
            src = har if n.startswith('h_') else raw
            if src is not None and n in src.columns and int(src[n].notna().sum()) != int(exp):
                add('INFO', f'{n}', f'expected_n={int(exp)} but the current raw export has {int(src[n].notna().sum())} answers - update the sheet if the export changed')
    return report()


def report():
    order = {'ERROR': 0, 'DATA': 1, 'INFO': 2}
    counts = {k: sum(1 for f in findings if f[0] == k) for k in order}
    for level, where, msg in sorted(findings, key=lambda f: (order[f[0]], f[1])):
        print(f'[{level}] {where}: {msg}')
    print(f'\n{counts["ERROR"]} sheet error(s), {counts["DATA"]} data issue(s) to fix at the source, {counts["INFO"]} note(s).')
    return 1 if counts['ERROR'] or counts['DATA'] else 0


if __name__ == '__main__':
    sys.exit(main())
