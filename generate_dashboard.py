"""
Command-line script that builds only the "dashboard" output table
(the same table app.py writes as `<name>_dashboard.xlsx`) from a Kobo
tool, a DAF file, and a flat data file.

Unlike app.py, this script:
  - reads the dataframe from .csv or .parquet instead of .xlsx
  - assumes the data is a single flat table (no repeat-group sheets)
  - only produces the dashboard table - no TOC/count tables, no
    analysis-key table, no significance testing, no conditional
    formatting, no join handling (none of these affect the dashboard
    table's contents - see docs/superpowers/specs for the trace-through
    that established that)

Usage:
    python generate_dashboard.py --data data.csv --tool tool.xlsx --daf daf.xlsx
    python generate_dashboard.py --data data.parquet --tool tool.xlsx --daf daf.xlsx \\
        --weight-column weight --sm-delimiter "|" --output my_dashboard.xlsx
    python generate_dashboard.py --data data.parquet --tool tool.xlsx --daf daf.xlsx \\
        --output my_dashboard.parquet
"""

import argparse
import sys
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import pandas as pd
from pandas.api.types import is_numeric_dtype

from www.src.functions import (
    detect_label_column,
    load_tool_choices,
    load_tool_survey,
    check_daf_consistency,
    check_daf_filter,
    disaggregation_creator,
)


def load_data_file(path):
    """
    Load a flat data table from a .csv or .parquet file.

    Parameters
    ----------
    path : str or Path
        Path to the data file. Format is chosen from the file extension
        (case-insensitive): '.csv' or '.parquet'.

    Returns
    -------
    pd.DataFrame

    Raises
    ------
    ValueError
        If the file extension isn't '.csv' or '.parquet'.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == '.csv':
        return pd.read_csv(path)
    elif suffix == '.parquet':
        return pd.read_parquet(path)
    elif suffix == '.xlsx':
        return pd.read_excel(path)
    else:
        raise ValueError(
            f"Unsupported data file format '{suffix}' - expected .csv or .parquet")


def write_dashboard_table(df, path):
    """
    Write the dashboard table out.

    Parameters
    ----------
    df : pd.DataFrame
        The dashboard table.
    path : str or Path
        Output path. Format is chosen from the file extension
        (case-insensitive): '.xlsx' or '.parquet'. Excel caps out at
        1,048,576 rows / 16,384 columns per sheet - use '.parquet' for
        tables that exceed that (it also filters faster downstream).

    Raises
    ------
    ValueError
        If the file extension isn't '.xlsx' or '.parquet'.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == '.xlsx':
        df.to_excel(path, index=False)
    elif suffix == '.parquet':
        if 'option' in df.columns:
            # A DAF variable classified as select_one/select_multiple/freq
            # whose underlying answer column is actually numeric (int/float)
            # leaves raw numbers in 'option' wherever map_names() can't find
            # a matching choice label. Mixed with the string labels other
            # variables produce, that makes 'option' an object column with
            # both str and float/int values, which pyarrow can't infer a
            # single Arrow type for. Coerce to string (preserving true NaNs
            # as nulls) so the parquet write never fails on this.
            df = df.copy()
            df['option'] = df['option'].apply(lambda x: x if pd.isna(x) else str(x))
            df['option_orig'] = df['option_orig'].apply(lambda x: x if pd.isna(x) else str(x))
        df.to_parquet(path, index=False)
    else:
        raise ValueError(
            f"Unsupported output file format '{suffix}' - expected .xlsx or .parquet")


def validate_daf_main(daf):
    """
    Validate the DAF main sheet, the same checks app.py's load_daf_file
    performs (duplicate variable/admin/disaggregation, bad func values,
    duplicate IDs, missing columns).

    Parameters
    ----------
    daf : pd.DataFrame
        The DAF 'main' sheet, as read from Excel.

    Returns
    -------
    list[str]
        Error messages. Empty if the DAF main sheet is valid. Column
        values are stripped of surrounding whitespace in place as a
        side effect, matching app.py's behavior.
    """
    errors = []
    required_columns = {'ID', 'variable', 'variable_label', 'calculation', 'func',
                         'admin', 'disaggregations', 'disaggregations_label', 'join'}
    missing_columns = required_columns.difference(daf.columns)
    if missing_columns:
        errors.append('Missing one or more columns from the DAF file main sheet: '
                       + ', '.join(sorted(missing_columns)))
        return errors

    for column in ['variable', 'admin', 'calculation', 'func', 'disaggregations']:
        daf[column] = daf[column].apply(lambda x: x.strip() if isinstance(x, str) else x)

    if any(daf['variable'] == daf['disaggregations']):
        ids = ', '.join(str(i) for i in daf.loc[daf['variable'] == daf['disaggregations'], 'ID'])
        errors.append(f'Variable and disaggregation are duplicated, problematic IDs: {ids}')

    if any(daf['variable'] == daf['admin']):
        ids = ', '.join(str(i) for i in daf.loc[daf['variable'] == daf['admin'], 'ID'])
        errors.append(f'Variable and admin are duplicated, problematic IDs: {ids}')

    if any(daf['admin'] == daf['disaggregations']):
        ids = ', '.join(str(i) for i in daf.loc[daf['admin'] == daf['disaggregations'], 'ID'])
        errors.append(f'Admin and disaggregation are duplicated, problematic IDs: {ids}')

    wrong_functions = set(daf['func']) - {'mean', 'numeric', 'select_one', 'select_multiple', 'freq'}
    if wrong_functions:
        errors.append(f'Wrong functions entered: {wrong_functions}. Please fix your function entries')

    if daf['ID'].duplicated().any():
        errors.append('Duplicate IDs in the ID column of the DAF')

    return errors


def validate_daf_filter_columns(filter_daf):
    """
    Check the DAF filter sheet has the required columns.

    Parameters
    ----------
    filter_daf : pd.DataFrame
        The DAF 'filter' sheet, as read from Excel.

    Returns
    -------
    list[str]
        Error messages. Empty if the required columns are present.
    """
    required_columns = {'ID', 'variable', 'operation', 'value'}
    missing_columns = required_columns.difference(filter_daf.columns)
    if missing_columns:
        return ['Missing one or more columns from the DAF file filter sheet: '
                + ', '.join(sorted(missing_columns))]
    return []


def merge_daf_with_datasheet(daf, data):
    """
    Tag each DAF row with the data sheet its variable belongs to.

    This script only supports a single flat data table (see the module
    docstring), so every DAF variable either belongs to 'main' or is
    missing from the data entirely - there's no multi-sheet dedup logic
    to run, unlike app.py's equivalent step.

    Parameters
    ----------
    daf : pd.DataFrame
        The DAF main sheet.
    data : dict[str, pd.DataFrame]
        Expected to have exactly one entry, keyed 'main'.

    Returns
    -------
    pd.DataFrame
        A copy of `daf` with a new 'datasheet' column: 'main' for every
        row whose variable exists in `data['main']`, else NA.
    """
    main_columns = set(data['main'].columns)
    daf_merged = daf.copy()
    daf_merged['datasheet'] = daf_merged['variable'].apply(
        lambda v: 'main' if v in main_columns else pd.NA)
    return daf_merged


def build_filter_dict(filter_daf, daf_merged, data):
    """
    Build the filter condition-string dictionary consumed by
    disaggregation_creator, using the same logic as app.py's
    download_data().

    Parameters
    ----------
    filter_daf : pd.DataFrame
        The DAF filter sheet. Non-ID columns are stripped of spaces and
        single quotes in place, matching app.py's behavior.
    daf_merged : pd.DataFrame
        The DAF main sheet, already tagged with a 'datasheet' column
        (see merge_daf_with_datasheet).
    data : dict[str, pd.DataFrame]
        The data dict (see merge_daf_with_datasheet) - referenced by
        name inside the generated condition strings.

    Returns
    -------
    dict[int, str]
        Maps each DAF row ID that has filters to a condition string
        ready to be eval'd by disaggregation_creator. Empty dict if
        `filter_daf` has no rows.
    """
    if filter_daf.shape[0] == 0:
        return {}

    for col in filter_daf.columns:
        if col != 'ID':
            filter_daf[col] = filter_daf[col].apply(
                lambda x: x.replace(' ', '').replace("'", '') if isinstance(x, str) else x)

    filter_daf_full = filter_daf.merge(daf_merged[['ID', 'datasheet']], on='ID', how='left')

    filter_dict = {}
    for _, row in filter_daf_full.iterrows():
        if isinstance(row['value'], str) and row['value'] in data[row['datasheet']].columns:
            condition_str = (f"(data['{row['datasheet']}']['{row['variable']}'] "
                              f"{row['operation']} data['{row['datasheet']}']['{row['value']}'])")
        elif isinstance(row['value'], str) and row['operation'] == '==':
            condition_str = (f"(data['{row['datasheet']}']['{row['variable']}']"
                              f".astype(str).str.contains('{row['value']}', regex=True))")
        elif isinstance(row['value'], str) and row['operation'] == '!=':
            condition_str = (f"(~data['{row['datasheet']}']['{row['variable']}']"
                              f".astype(str).str.contains('{row['value']}', regex=True))")
        else:
            condition_str = (f"(data['{row['datasheet']}']['{row['variable']}'] "
                              f"{row['operation']} {row['value']})")
        filter_dict.setdefault(row['ID'], []).append(condition_str)

    filter_dict = {key: ' & '.join(value) + ']' for key, value in filter_dict.items()}
    return filter_dict


def build_dashboard_table(daf_final, data, filter_dict, tool_choices, tool_survey,
                           label_column, weight_column, sm_delimiter):
    """
    Run disaggregation_creator and reduce its output to the dashboard
    table, the same post-processing app.py applies to produce its
    `<name>_dashboard.xlsx` output: drop the raw count columns, drop
    'Total' rows, and fill blank disaggregation columns with ' Overall'.

    check_significance is always False - the dashboard table is built
    purely from each table's DataFrame (element[0]), never from the
    significance result (element[3]), so computing it would be wasted
    work for this script's one output.

    Parameters
    ----------
    daf_final : pd.DataFrame
        The DAF main sheet, merged with 'datasheet' and 'q.type'.
    data : dict[str, pd.DataFrame]
        The data dict (see merge_daf_with_datasheet).
    filter_dict : dict[int, str]
        See build_filter_dict.
    tool_choices, tool_survey : pd.DataFrame
        See www.src.functions.load_tool_choices / load_tool_survey.
    label_column : str
        The resolved label column name.
    weight_column : str or None
        Name of the weight column in `data['main']`, or None for
        unweighted analysis.
    sm_delimiter : str
        Delimiter used to join multiple selections within a single
        select_multiple answer.

    Returns
    -------
    pd.DataFrame
        The dashboard table.
    """
    disaggregations_full = disaggregation_creator(
        daf_final, data, filter_dict, tool_choices, tool_survey,
        label_colname=label_column, check_significance=False,
        weight_column=weight_column, sm_delimiter=sm_delimiter,
    )

    disaggregations_perc = deepcopy(disaggregations_full)
    for element in disaggregations_perc:
        if isinstance(element[0], pd.DataFrame):
            for column in ['category_count', 'weighted_count', 'unweighted_count']:
                if column in element[0].columns:
                    element[0].drop(columns=column, inplace=True)

    concatenated_df = pd.concat([tpl[0] for tpl in disaggregations_perc], ignore_index=True)
    if 'disaggregations_category_1' in concatenated_df.columns:
        concatenated_df = concatenated_df[
            (concatenated_df['admin'] != 'Total')
            & (concatenated_df['disaggregations_category_1'] != 'Total')]
    else:
        concatenated_df = concatenated_df[concatenated_df['admin'] != 'Total']

    disagg_columns = [col for col in concatenated_df.columns if col.startswith('disaggregations')]
    if disagg_columns:
        concatenated_df.loc[:, disagg_columns] = concatenated_df[disagg_columns].fillna(' Overall')

    return concatenated_df


def run(data_path, tool_path, daf_path, weight_column=None, sm_delimiter=' ', output_path=None):
    """
    Run the full pipeline: load inputs, validate, build the dashboard
    table, write it out.

    Raises
    ------
    ValueError
        On any validation failure (bad DAF, missing variables, wrong
        types, etc.), with all problems found joined into one message.
    """
    data_df = load_data_file(data_path)
    data = {'main': data_df}

    sheets_tool = pd.ExcelFile(tool_path).sheet_names
    if not {'survey', 'choices'}.issubset(set(sheets_tool)):
        raise ValueError('Missing survey or choices sheet in the kobo tool file')
    tool_s = pd.read_excel(tool_path, sheet_name='survey')
    tool_c = pd.read_excel(tool_path, sheet_name='choices')
    settings_sheet = pd.read_excel(tool_path, sheet_name='settings') if 'settings' in sheets_tool else None
    label_column = detect_label_column(tool_s, tool_c, tool_settings=settings_sheet)

    tool_choices = load_tool_choices(tool_path, label_colname=label_column)
    tool_survey = load_tool_survey(tool_path, label_colname=label_column)

    sheets_daf = pd.ExcelFile(daf_path).sheet_names
    if not {'main', 'filter'}.issubset(set(sheets_daf)):
        raise ValueError('Missing one of the following sheets for DAF file: main, filter')
    daf = pd.read_excel(daf_path, sheet_name='main')
    filter_daf = pd.read_excel(daf_path, sheet_name='filter')

    errors = validate_daf_main(daf)
    errors += validate_daf_filter_columns(filter_daf)
    if errors:
        raise ValueError(' '.join(errors))

    if weight_column is not None and weight_column not in data['main'].columns:
        raise ValueError(f"Weight column '{weight_column}' not found in the data file")
    if weight_column is not None and data['main'][weight_column].isnull().any():
        raise ValueError(f"The weight column '{weight_column}' contains NAs, please fix this")

    data['main']['overall'] = ' Overall'
    data['main']['Overall'] = ' Overall'

    daf_merged = merge_daf_with_datasheet(daf, data)
    missing_vars = daf_merged.loc[daf_merged['datasheet'].isna(), 'variable'].unique()
    if len(missing_vars) > 0:
        raise ValueError('Missing the following variables from the dataframe: '
                          + ', '.join(missing_vars.astype(str)))

    daf_merged = check_daf_consistency(daf_merged, data, ['main'], resolve=False)

    numeric_errors = []
    daf_numeric = daf_merged[daf_merged['func'].isin(['numeric', 'mean'])]
    for _, daf_row in daf_numeric.iterrows():
        if not is_numeric_dtype(data[daf_row['datasheet']][daf_row['variable']]):
            numeric_errors.append(
                f"Variable {daf_row['variable']} is not numeric, but you want to "
                f"apply a mean/numeric function to it in your DAF")
    if numeric_errors:
        raise ValueError(' '.join(numeric_errors))

    if filter_daf.shape[0] > 0:
        check_daf_filter(daf=daf_merged, data=data, filter_daf=filter_daf, tool_survey=tool_survey)
    filter_dict = build_filter_dict(filter_daf, daf_merged, data)

    daf_final = daf_merged.merge(tool_survey[['name', 'q.type']], left_on='variable', right_on='name', how='left')
    daf_final['q.type'] = daf_final['q.type'].fillna('select_one')

    dashboard_table = build_dashboard_table(
        daf_final, data, filter_dict, tool_choices, tool_survey,
        label_column, weight_column, sm_delimiter,
    )

    if output_path is None:
        output_path = 'dashboard_' + datetime.today().strftime('%Y_%m_%d') + '.xlsx'
    write_dashboard_table(dashboard_table, output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description='Build the dashboard-input table from a Kobo tool, a DAF, and a flat CSV/parquet data file.')
    parser.add_argument('--data', required=True, help='Path to the data file (.csv or .parquet)')
    parser.add_argument('--tool', required=True, help='Path to the Kobo tool (.xlsx)')
    parser.add_argument('--daf', required=True, help='Path to the DAF file (.xlsx, with main and filter sheets)')
    parser.add_argument('--weight-column', default=None, help='Name of the weight column in the data file (optional)')
    parser.add_argument('--sm-delimiter', default=' ',
                         help="Delimiter joining multiple selections in a select_multiple answer (default: a space)")
    parser.add_argument('--output', default=None,
                         help='Output path, format chosen from the extension: .xlsx or .parquet '
                              '(default: dashboard_<date>.xlsx)')
    args = parser.parse_args()

    try:
        output_path = run(
            data_path=args.data,
            tool_path=args.tool,
            daf_path=args.daf,
            weight_column=args.weight_column,
            sm_delimiter=args.sm_delimiter,
            output_path=args.output,
        )
    except (ValueError, OSError) as e:
        print(f'Error: {e}', file=sys.stderr)
        sys.exit(1)

    print(f'Wrote {output_path}')


if __name__ == '__main__':
    main()
