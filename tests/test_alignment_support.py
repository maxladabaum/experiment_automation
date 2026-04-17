from pathlib import Path

from core.runner import SerialMeasurementRunner


class _FakeConnection:
    def __init__(self, chunks):
        self._chunks = list(chunks)
        self.is_open = True

    @property
    def in_waiting(self):
        return len(self._chunks[0]) if self._chunks else 0

    def read(self, _size):
        if self._chunks:
            return self._chunks.pop(0)
        return b""

    def write(self, _data):
        return None

    def close(self):
        self.is_open = False


def test_runner_parses_alignment_fields():
    runner = SerialMeasurementRunner(Path("methods/library/cv_f2add5.ms"), simulate_measurements=True)
    runner._parse_data_line("Pcg80003E8i;cc8002710i;cd8001388i")

    assert len(runner.data_points) == 1
    point = runner.data_points[0]
    assert point["frequency_hz"] == 1000.0
    assert "z_real_ohm" in point
    assert "z_imag_ohm" in point
    assert "impedance_ohm" in point
    assert "capacitance_nf" in point


def test_runner_simulates_eis_points():
    script = (
        "e\n"
        "var freq\n"
        "var zr\n"
        "var zi\n"
        "meas_loop_eis freq zr zi 10m 10 1000 5i 0\n"
        "\tpck_start\n"
        "\t\tpck_add freq\n"
        "\t\tpck_add zr\n"
        "\t\tpck_add zi\n"
        "\tpck_end\n"
        "endloop\n"
    )

    points = SerialMeasurementRunner._sim_eis_points(script)

    assert len(points) == 5
    assert points[0]["frequency_hz"] > 0
    assert all("impedance_ohm" in point for point in points)
    assert all("capacitance_nf" in point for point in points)


def test_runner_treats_bare_star_as_completion_marker():
    logs = []
    runner = SerialMeasurementRunner(
        Path("methods/library/cv_f2add5.ms"),
        log_callback=logs.append,
        simulate_measurements=True,
    )
    runner.connection = _FakeConnection([
        b"R*\n",
        b"e\n",
        b"M0002\n",
        b"*\n",
    ])

    success = runner.run_script("e\n")

    assert success is True
    assert any("Measurement completed" in entry for entry in logs)
    assert not any("idle timed out" in entry for entry in logs)
