import argparse
import random
from nn.network import *

RANDOM_STATE = 42


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert csv to Excel neural network')
    parser.add_argument('filename', help='path to input csv')
    parser.add_argument('config', help='path to config.json')
    args = parser.parse_args()

    RANDOM_STATE = 42
    random.seed(RANDOM_STATE)

    network = Network(args.filename, args.config, RANDOM_STATE)
    network.save()
