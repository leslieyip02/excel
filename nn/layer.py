import numpy as np


class Layer():
    config: dict[str, int | str]
    weights: np.ndarray[np.float64, np.float64]
    bias: np.float64
    macro: str

    def __init__(self, config) -> None:
        input_size = config['input_size']
        output_size = config['output_size']
        self.weights = (np.random.rand(input_size, output_size) - 0.5) * 2.0
        self.bias = 1.0

        # load a macro
        activation_function = config['activation_function']
        self.macro = ''