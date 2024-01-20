import numpy as np
import pandas as pd
import json
from nn.layer import *
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class Network():
    wb: Workbook
    layers: list[Layer]

    def __init__(self, csv_path: str, config_path: str) -> None:
        self.wb = Workbook()
        self.wb.active.title = 'Data'

        self.init_data(csv_path)
        self.init_layers(config_path)

    def init_data(self, csv_path: str):
        data_sheet = self.wb['Data']
        df = pd.read_csv(csv_path)
        rows = dataframe_to_rows(df, index=False)
        for row in rows:
            data_sheet.append(row)

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
