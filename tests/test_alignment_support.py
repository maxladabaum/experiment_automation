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
        "var f\n"
        "var z\n"
        "var i\n"
        "set_max_bandwidth 40\n"
        "set_range ba 100u\n"
        "set_autoranging ba 1n 100u\n"
        "meas_loop_eis f z i 10m 10 1000 5i 0\n"
        "\tpck_start\n"
        "\t\tpck_add f\n"
        "\t\tpck_add z\n"
        "\t\tpck_add i\n"
        "\tpck_end\n"
        "endloop\n"
    )

    points = SerialMeasurementRunner._sim_eis_points(script)

    assert len(points) == 5
    assert points[0]["frequency_hz"] > 0
    assert all("impedance_ohm" in point for point in points)
    assert all("capacitance_nf" in point for point in points)


def test_runner_does_not_treat_bare_star_as_script_completion():
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
        b"Measurement completed\n",
    ])

    success = runner.run_script("e\n")

    assert success is True
    assert any("Measurement completed" in entry for entry in logs)
    assert not any("idle timed out" in entry for entry in logs)


def test_runner_continues_past_first_loop_terminator():
    logs = []
    runner = SerialMeasurementRunner(
        Path("methods/library/swv_a0df5f.ms"),
        log_callback=logs.append,
        simulate_measurements=True,
    )
    runner.connection = _FakeConnection([
        b"R*\n",
        b"e\n",
        b"M0007\n",
        b"Peb0000000i;ab8000000i;ba8000001i\n",
        b"*\n",
        b"M0008\n",
        b"Pda8000001i;ba8000002i;ba8000003i;ba8000004i\n",
        b"*\n",
        b"Measurement completed\n",
    ])

    success = runner.run_script("e\n")

    assert success is True
    assert len(runner.data_points) >= 2
    assert any("current" in point or "current_diff" in point for point in runner.data_points)
