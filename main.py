import argparse
import pandas as pd
import random
from nn.network import *


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('filename')
    args = parser.parse_args()

    random.seed(42)

    network = Network(args.filename, 'config.json')
    network.save('tmp')
