import numpy as np
import pandas as pd
import json
from nn.layer import *
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from sklearn.model_selection import train_test_split


class Network():
    wb: Workbook
    layers: list[Layer]
    random_state: int

    def __init__(self, csv_path: str, config_path: str, random_state: int) -> None:
        self.wb = Workbook()
        self.wb.active.title = 'Training Data'
        self.wb.create_sheet('Test Data')
        self.random_state = random_state

        self.init_data(csv_path)
        self.init_layers(config_path)

    def init_data(self, csv_path: str):
        df = pd.read_csv(csv_path)
        X = df.iloc[:, :-1].apply(pd.to_numeric)
        y = df.iloc[:, -1].apply(pd.to_numeric)
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.3, random_state=self.random_state)

        train_sheet = self.wb['Training Data']
        train_sheet.append(list(df.columns))
        for i in range(len(X_train)):
            row = list(X_train.iloc[i])
            row.append(y_train.iloc[i])
            train_sheet.append(row)

        train_sheet.append(list(df.columns))
        test_sheet = self.wb['Test Data']
        for i in range(len(X_test)):
            row = list(X_test.iloc[i])
            row.append(y_test.iloc[i])
            test_sheet.append(row)

    def init_layers(self, config_path):
        config = json.load(open(config_path))
        for i, config in enumerate(config['layers']):
            sheet = self.wb.create_sheet(f'layer_{i + 1}')
            layer = Layer(config)
            for row in layer.weights:
                sheet.append(list(row))

    def save(self, title: str) -> None:
        if not title.endswith('.xlsm'):
            title += '.xlsm'
        self.wb.save(title)

        # weird behavior
        # https://stackoverflow.com/questions/59585265/is-there-any-way-to-create-a-xlsm-file-from-scratch-in-python
        wb = load_workbook(title, keep_vba=True)
        wb.save(title)
