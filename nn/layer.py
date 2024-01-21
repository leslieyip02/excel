import numpy as np


class Layer():
    config: dict[str, int | str]
    weights: np.ndarray[np.float64, np.float64]
    bias: np.ndarray[np.float64]
    macro: str

    activation_function_indices = {
        'relu': 1,
        'softmax': 2,
    }

    def __init__(self, config) -> None:
        input_size = config['input_size']
        output_size = config['output_size']
        self.weights = np.random.rand(input_size, output_size) - 0.25
        self.bias = np.ones((input_size,))
        self.macro = Layer.activation_function_indices[config['activation_function']]
