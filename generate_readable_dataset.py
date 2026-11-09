"""
Command-line script that replaces every Kobo question/choice code in a
dataset with its human-readable label: column headers become question
labels, and select_one/select_multiple cell values become choice labels.

Unlike app.py/generate_dashboard.py, this script needs no DAF file at
all - it just decodes one sheet of a data workbook against a Kobo tool.

Usage:
    python generate_readable_dataset.py --data data.xlsx --tool tool.xlsx
    python generate_readable_dataset.py --data data.xlsx --tool tool.xlsx \\
        --sheet hh_members --sm-delimiter "|" --output readable.xlsx
"""

import argparse
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

from www.src.functions import detect_label_column, load_tool_choices, load_tool_survey

GROUP_OR_REPEAT_MARKER = re.compile(r'((begin)|(end))[_ ](group|repeat)', re.IGNORECASE)


def load_data_file(path, sheet_name=None):
    """
    Load one sheet of an xlsx data file.

    Parameters
    ----------
    path : str or Path
        Path to the data workbook.
    sheet_name : str, optional
        Sheet to load. If None, the workbook's first sheet is used.

    Returns
    -------
    pd.DataFrame
    """
    if sheet_name is None:
        sheet_name = pd.ExcelFile(path).sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet_name)


def resolve_column_labels(tool_survey_raw, label_colname):
    """
    Map every raw Kobo question name found in the tool to its own label,
    covering every question type (not just select_one/select_multiple/
    integer/decimal) since load_tool_survey() filters those out.

    If two or more questions share the exact same label text, each of
    those (and only those) gets disambiguated as "label (raw_name)" so
    the output never has duplicate column headers.

    Parameters
    ----------
    tool_survey_raw : pd.DataFrame
        The UNFILTERED 'survey' sheet, as read directly via
        pd.read_excel(tool_path, sheet_name='survey') - not the output
        of load_tool_survey(), which drops non-eligible question types
        before this function would ever see them.
    label_colname : str
        The resolved label column name (e.g. from detect_label_column).

    Returns
    -------
    dict[str, str]
        Maps each raw question name to its (possibly disambiguated)
        label. Group/repeat marker rows and rows with a blank name or
        label are skipped.
    """
    raw_to_label = {}
    for _, row in tool_survey_raw.iterrows():
        row_type = row.get('type')
        if not isinstance(row_type, str) or GROUP_OR_REPEAT_MARKER.search(row_type):
            continue
        name = row.get('name')
        if not isinstance(name, str) or not name:
            continue
        label = row.get(label_colname)
        if not isinstance(label, str) or not label.strip():
            continue
        raw_to_label[name] = label

    label_counts = Counter(raw_to_label.values())
    return {
        name: (f'{label} ({name})' if label_counts[label] > 1 else label)
        for name, label in raw_to_label.items()
    }


def build_choice_label_map(tool_choices, list_name, label_colname):
    """
    Build a {choice code: label} dict for one choice list.

    Parameters
    ----------
    tool_choices : pd.DataFrame
        See www.src.functions.load_tool_choices.
    list_name : str
        The choice list to build a map for (a select_one/select_multiple
        question's `list_name`).
    label_colname : str
        The resolved label column name.

    Returns
    -------
    dict[str, str]
    """
    subset = tool_choices[tool_choices['list_name'] == list_name]
    return dict(zip(subset['name'], subset[label_colname]))


def decode_select_one_value(value, label_map):
    """Map a single select_one code to its label, leaving NaN and unknown codes as-is."""
    if pd.isna(value):
        return value
    return label_map.get(value, value)


def decode_select_multiple_value(value, label_map, delimiter):
    """
    Map every code in a delimiter-joined select_multiple cell to its
    label, rejoining with the same delimiter. NaN and unknown codes are
    left as-is (an unknown code stays as its own raw text).
    """
    if pd.isna(value):
        return value
    tokens = str(value).split(delimiter)
    return delimiter.join(label_map.get(token, token) for token in tokens)


def decode_dataframe(data, tool_survey, tool_choices, column_labels, label_colname, sm_delimiter):
    """
    Decode select_one/select_multiple values and rename every
    tool-sourced column to its label.

    Parameters
    ----------
    data : pd.DataFrame
        The raw data sheet to decode. Not mutated - a decoded copy is
        returned.
    tool_survey : pd.DataFrame
        load_tool_survey()'s output (eligible types only, with `q.type`
        and `list_name` columns) - used to decide which columns need
        value decoding and which choice list to use.
    tool_choices : pd.DataFrame
        See www.src.functions.load_tool_choices.
    column_labels : dict[str, str]
        See resolve_column_labels - applied to every column, not just
        the ones that get value-decoded.
    label_colname : str
        The resolved label column name.
    sm_delimiter : str
        Delimiter used to join/split multiple selections within a
        single select_multiple answer.

    Returns
    -------
    pd.DataFrame
    """
    data = data.copy()
    survey_by_name = tool_survey.drop_duplicates(subset='name', keep='first').set_index('name')

    for col in data.columns:
        if col not in survey_by_name.index:
            continue
        q_type = survey_by_name.loc[col, 'q.type']
        list_name = survey_by_name.loc[col, 'list_name']
        if not isinstance(list_name, str):
            continue
        label_map = build_choice_label_map(tool_choices, list_name, label_colname)
        if q_type == 'select_one':
            data[col] = data[col].apply(lambda v: decode_select_one_value(v, label_map))
        elif q_type == 'select_multiple':
            data[col] = data[col].apply(lambda v: decode_select_multiple_value(v, label_map, sm_delimiter))

    return data.rename(columns=column_labels)


def run(data_path, tool_path, sheet_name=None, sm_delimiter=' ', output_path=None):
    """
    Run the full pipeline: load the data and tool, decode, write the
    human-readable output.

    Raises
    ------
    ValueError
        If the tool is missing its survey/choices sheet or the label
        column can't be resolved.
    """
    data = load_data_file(data_path, sheet_name=sheet_name)

    sheets_tool = pd.ExcelFile(tool_path).sheet_names
    if not {'survey', 'choices'}.issubset(set(sheets_tool)):
        raise ValueError('Missing survey or choices sheet in the kobo tool file')
    tool_s = pd.read_excel(tool_path, sheet_name='survey')
    tool_c = pd.read_excel(tool_path, sheet_name='choices')
    settings_sheet = pd.read_excel(tool_path, sheet_name='settings') if 'settings' in sheets_tool else None
    label_column = detect_label_column(tool_s, tool_c, tool_settings=settings_sheet)

    tool_choices = load_tool_choices(tool_path, label_colname=label_column)
    tool_survey = load_tool_survey(tool_path, label_colname=label_column)
    column_labels = resolve_column_labels(tool_s, label_column)

    readable = decode_dataframe(data, tool_survey, tool_choices, column_labels, label_column, sm_delimiter)

    if output_path is None:
        stem = Path(data_path).stem
        output_path = f'readable_{stem}_' + datetime.today().strftime('%Y_%m_%d') + '.xlsx'
    readable.to_excel(output_path, index=False)
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description='Replace Kobo question/choice codes in a dataset with their human-readable labels.')
    parser.add_argument('--data', required=True, help='Path to the data file (.xlsx)')
    parser.add_argument('--tool', required=True, help='Path to the Kobo tool (.xlsx)')
    parser.add_argument('--sheet', default=None, help='Sheet name to read from --data (default: the first sheet)')
    parser.add_argument('--sm-delimiter', default=' ',
                         help='Delimiter joining multiple selections in a select_multiple answer, '
                              'used for both reading and re-writing (default: a space)')
    parser.add_argument('--output', default=None,
                         help='Output .xlsx path (default: readable_<data filename>_<date>.xlsx)')
    args = parser.parse_args()

    try:
        output_path = run(
            data_path=args.data,
            tool_path=args.tool,
            sheet_name=args.sheet,
            sm_delimiter=args.sm_delimiter,
            output_path=args.output,
        )
    except (ValueError, OSError) as e:
        print(f'Error: {e}', file=sys.stderr)
        sys.exit(1)

    print(f'Wrote {output_path}')


if __name__ == '__main__':
    main()
