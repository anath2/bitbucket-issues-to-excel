"""
    Script to convert bitbucket issues export to excel
"""
import argparse
import json
import pandas as pd

def parse(in_file, out_file):
    """
        Convert the issues json exported from bitbucket
        to excel
        Columns
            1) Issue title
            2) Content
            3) Date added
            4) Type  
            5) Priority 
    """
    with open(in_file) as f:
        data = json.load(f)

    if not 'issues' in data:
        print('No issues')
        return
    
    issues_df = pd.DataFrame()
    for issue in data['issues']:
        ser = pd.Series({
            'Title': issue['title'],
            'Description': issue['content'],
            'Date added': issue['created_on'],
            'Kind': issue['kind'],
            'Priority': issue['priority'],
            'Status': issue['status']
        })
        issues_df = issues_df.append(issue)
    
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
    issues_df.to_excel(writer, sheet_name='Sheet1')

    writer.save()