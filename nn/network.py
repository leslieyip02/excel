import json
import os
import sys
import pandas as pd
import win32com.client
from nn.layer import *
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from sklearn.model_selection import train_test_split
from tqdm import tqdm


class Network():
    wb: Workbook
    layers: list[Layer]
    random_state: int
    filename: str

    def __init__(
        self,
        csv_path: str,
        config_path: str,
        random_state: int,
    ) -> None:
        self.wb = Workbook()
        self.wb.active.title = 'Training Data'
        self.wb.create_sheet('Test Data')
        self.random_state = random_state
        self.filename = csv_path.split('/')[-1].replace('.csv', '')

        self.init_data(csv_path)
        self.init_layers(config_path)
        self.save()
        self.inject_macros()

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

        test_sheet = self.wb['Test Data']
        test_sheet.append(list(df.columns))
        for i in range(len(X_test)):
            row = list(X_test.iloc[i])
            row.append(y_test.iloc[i])
            test_sheet.append(row)

    def init_layers(self, config_path):
        config = json.load(open(config_path))
        self.layers = []
        for i, config in tqdm(enumerate(config['layers']), desc='Initializing weights'):
            sheet = self.wb.create_sheet(f'layer_{i + 1}')
            layer = Layer(config)
            for row in layer.weights:
                sheet.append(list(row))
            self.layers.append(layer)

    def inject_macros(self):
        _file = os.path.abspath(sys.argv[0])
        path = os.path.join(os.path.dirname(_file), f'{self.filename}.xlsm')
        excel = win32com.client.Dispatch("Excel.Application")

        try:
            excel.Visible = False
            excel.DisplayAlerts = False
            # need to enable access
            # https://stackoverflow.com/questions/25638344/programmatic-access-to-visual-basic-project-is-not-trusted
            wb = excel.Workbooks.Open(Filename=path)

            for layer in tqdm(self.layers, desc='Injecting macros'):
                excelModule = wb.VBProject.VBComponents.Add(1)
                excelModule.CodeModule.AddFromString(layer.macro)
                wb.SaveAs(path)

            excel.Workbooks(1).Close(SaveChanges=1)

        except Exception as e:
            print(f'Unable to load {path}: {e}')

        finally:
            excel.Application.Quit()
            del excel

        self.wb = load_workbook(path, keep_vba=True)

    def save(self) -> None:
        path = f'{self.filename}.xlsm'
        self.wb.save(path)

        # weird behavior
        # https://stackoverflow.com/questions/59585265/is-there-any-way-to-create-a-xlsm-file-from-scratch-in-python
        self.wb = load_workbook(path, keep_vba=True)
        self.wb.save(path)
