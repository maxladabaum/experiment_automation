import pickle

import numpy as np

from core.bo_session import NumpyGaussianProcessRegressor, _gp_kernel_metadata


def test_numpy_gp_predicts_training_values_and_uncertainty():
    x_train = np.asarray([[0.0], [0.5], [1.0]], dtype=float)
    y_train = np.asarray([1.0, 3.0, 2.0], dtype=float)
    gp = NumpyGaussianProcessRegressor(
        x_train,
        y_train,
        length_scales=[0.25],
        noise_level=1e-8,
    )

    means, stds = gp.predict(x_train, return_std=True)

    assert np.allclose(means, y_train, atol=1e-5)
    assert np.all(stds < 1e-3)
    midpoint_mean, midpoint_std = gp.predict([[0.25]], return_std=True)
    assert np.isfinite(midpoint_mean[0])
    assert midpoint_std[0] > 0.0


def test_numpy_gp_is_pickleable_and_exports_kernel_metadata():
    gp = NumpyGaussianProcessRegressor(
        [[0.0, 0.0], [1.0, 1.0]],
        [2.0, 4.0],
        length_scales=[0.2, 0.4],
        noise_level=1e-4,
    )

    restored = pickle.loads(pickle.dumps(gp))
    original = gp.predict([[0.3, 0.7]], return_std=True)
    reloaded = restored.predict([[0.3, 0.7]], return_std=True)
    metadata = _gp_kernel_metadata(restored)

    assert np.allclose(original[0], reloaded[0])
    assert np.allclose(original[1], reloaded[1])
    assert metadata["gp_matern_nu"] == 2.5
    assert metadata["gp_noise_level"] >= 1e-4
