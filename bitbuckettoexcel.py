"""
    Script to convert bitbucket issues export to excel
"""
import argparse
import json
import pandas as pd

def parse(input_file):
    """
        Convert the issues json exported from bitbucket
        to excel
    """
    with open(input_file) as f:
        data = json.load(f)
