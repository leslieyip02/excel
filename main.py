import argparse
import openpyxl
import pandas as pd


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('filename')

    args = parser.parse_args()
    df = pd.read_csv(args.filename)
    print(df.head())