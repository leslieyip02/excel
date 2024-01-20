import argparse
import pandas as pd
import random
from nn.network import *


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('filename')
    args = parser.parse_args()

    random_state = 42
    random.seed(random_state)

    network = Network(args.filename, 'config.json', random_state)
    network.save('tmp')
