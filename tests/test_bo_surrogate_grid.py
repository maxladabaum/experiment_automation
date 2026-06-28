import numpy as np

from core.bo_session import OPTIMIZER_ORDER, load_bo_config
from gui.tab_bayesian_optimization import BayesianOptimizationTab


class _RecordingGP:
    def __init__(self):
        self.points = None

    def predict(self, points, return_std=False):
        self.points = np.asarray(points, dtype=float)
        means = self.points.sum(axis=1)
        stds = np.full(len(self.points), 0.2)
        return (means, stds) if return_std else means


def test_surrogate_2d_grid_uses_encoded_regular_mesh_and_log_frequency_axis():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._bo_session = None
    tab._config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    tab._config["parameters"]["frequency"]["scale"] = "log"
    gp = _RecordingGP()
    tab._load_surrogate_gp_model = lambda: gp
    tab._surrogate_observations_so_far = lambda: []
    tab._selected_surrogate_observation = lambda: None

    grid = tab._surrogate_2d_prediction_grid(
        [{"best_observed_Q": 0.0}],
        "predicted_mean_Q",
        "frequency",
        "amplitude",
        grid_size=30,
    )

    assert grid is not None
    frequency, amplitude, values, frequency_is_log, amplitude_is_log = grid
    assert values.shape == (30, 30)
    assert frequency_is_log is True
    assert amplitude_is_log is False

    frequency_idx = OPTIMIZER_ORDER.index("frequency")
    amplitude_idx = OPTIMIZER_ORDER.index("amplitude")
    assert np.unique(gp.points[:, frequency_idx]).size == 30
    assert np.unique(gp.points[:, amplitude_idx]).size == 30
    assert gp.points[:, frequency_idx].min() == 0.0
    assert gp.points[:, frequency_idx].max() == 1.0

    ratios = frequency[1:] / frequency[:-1]
    assert np.allclose(ratios, ratios[0])
    assert np.all(np.diff(amplitude) > 0)
