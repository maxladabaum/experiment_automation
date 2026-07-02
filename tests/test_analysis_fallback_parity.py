import numpy as np
import pytest

from core.analysis_processing import (
    _savgol_numpy,
    fallback_find_peaks,
)


scipy_signal = pytest.importorskip("scipy.signal")


def test_numpy_savgol_matches_scipy_including_edges():
    rng = np.random.default_rng(42)
    signal = np.sin(np.linspace(0.0, 8.0, 101)) + rng.normal(0.0, 0.08, 101)

    expected = scipy_signal.savgol_filter(signal, window_length=15, polyorder=2)
    actual = _savgol_numpy(signal, window_length=15, polyorder=2)

    assert np.allclose(actual, expected, atol=1e-10)


def test_fallback_distance_filter_keeps_taller_peak_like_scipy():
    signal = np.zeros(30, dtype=float)
    signal[8:13] = 4.0
    signal[10] = 5.0
    signal[14] = 4.0

    expected, _ = scipy_signal.find_peaks(signal, prominence=0.5, distance=5)
    actual, _ = fallback_find_peaks(signal, prominence=0.5, distance=5)

    assert actual.tolist() == expected.tolist()
