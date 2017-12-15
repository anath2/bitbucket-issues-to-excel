"""
    Script to convert bitbucket issues export to excel
"""
import os
import argparse
import json
import datetime

import pandas as pd
import pandas.io.formats.excel

def parse_json(in_file, out_file):
    """
        Convert the issues json exported from bitbucket
        to excel
        Columns
            1) Title
            2) Description
            3) Date added
            4) Kind
            5) Priority
            6) Status
    """
    with open(in_file) as f:
        data = json.load(f)

    if not 'issues' in data:
        print('No issues')
        return

    issues_df = pd.DataFrame(columns=[
        'Title',
        'Description',
        'Date added',
        'Kind',
        'Priority',
        'Status',
    ])
    for issue in data['issues']:
        ser = pd.Series({
            'Title': issue['title'],
            'Description': issue['content'],
            'Date added': _format_time(issue['created_on']),
            'Kind': issue['kind'],
            'Priority': issue['priority'],
            'Status': issue['status']
        })
        issues_df = issues_df.append(ser, ignore_index=True)

    pandas.io.formats.excel.header_style = None
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
    issues_df.to_excel(writer, sheet_name='Sheet1', index=False)
    _format_and_save_excel(writer)


def _add_date(f_path):
    # Add datetime to output file
    f_name = os.path.basename(f_path)
    datetime_str = datetime.datetime.strftime(
        datetime.datetime.now(),
        '%Y%m%d'
    )
    return os.path.join(
        os.path.dirname(f_path),
        datetime_str + f_name
    )

def _format_time(time_str):
    # Format str looks something like this 2017-12-14T07:10:34.500895+00:00
    date_str = time_str.split('T')[0]
    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')
    formatted_date_str = datetime.datetime.strftime(date_obj, "%A %d %m %Y ")
    return formatted_date_str

def _format_and_save_excel(writer):
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.set_zoom(120)
    header_fmt = workbook.add_format({
        'text_wrap': True,
        'bg_color': '#191970',
        'font_color': '#FFFFFF',
    })
    cell_fmt = workbook.add_format({
        'text_wrap': True,
        'font_color': '#333333',
    })

    text_fmt = workbook.add_format({
        'text_wrap': True,
        'font_color': '#333333',
    })

    green_fmt = workbook.add_format({
        'bg_color': '#22CC11',
        'font_color': '#FFFFFF',
        'text_wrap': True,
        'bold': True
    })
    red_fmt = workbook.add_format({
        'bg_color': '#E74C3C',
        'font_color': '#FFFFFF',
        'text_wrap': True,
        'bold': True
    })
    orange_fmt = workbook.add_format({
        'bg_color': '#F39C12',
        'font_color': '#FFFFFF',
        'text_wrap': True,
    })

    header_fmt.set_align('center')
    header_fmt.set_align('vcenter')
    cell_fmt.set_align('center')
    cell_fmt.set_align('vcenter')
    text_fmt.set_align('left')
    text_fmt.set_align('vcenter')

    worksheet.set_column('A2:B2', 40, text_fmt)
    worksheet.set_column('C2:F2', 40, cell_fmt)
    worksheet.set_row(0, None, header_fmt)


    worksheet.conditional_format('D2:D12', {
        'type':     'text',
        'criteria': 'containing',
        'value':    'enhancement',
        'format':   green_fmt
    })
    worksheet.conditional_format('D2:D12', {
        'type':     'text',
        'criteria': 'containing',
        'value':    'bug',
        'format':   red_fmt
    })

    worksheet.conditional_format('E2:E12', {
        'type':     'text',
        'criteria': 'containing',
        'value':    'major',
        'format':   orange_fmt
    })
    worksheet.conditional_format('E2:E12', {
        'type':     'text',
        'criteria': 'containing',
        'value':    'critical',
        'format':   red_fmt
    })

    worksheet.conditional_format('F2:F12', {
        'type':     'text',
        'criteria': 'containing',
        'value':    'resolved',
        'format':   green_fmt
    })

    writer.save()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate excel output')
    parser.add_argument('-i', '--input', help='Input json file', type=str)
    parser.add_argument('-o', '--output', help='Output excel file', type=str)
    args = parser.parse_args()

    parse_json(args.input, _add_date(args.output))
