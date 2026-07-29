import json
import tempfile
import unittest
from pathlib import Path

from core.bo_session import BOIntegrationSession


class BOWindowsPathTests(unittest.TestCase):
    def test_nested_analysis_request_stays_below_max_path(self):
        config_path = Path("optimizer/bo_configs/default_swv_bo.json")
        config = json.loads(config_path.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp_dir:
            experiment = (
                Path(temp_dir)
                / "bo_new_machine_20260716_154144"
                / "fluidic_bo_after_patch_20260716_163613"
            )
            session = BOIntegrationSession(config, experiment, config_path=config_path)
            request = (
                session.analysis_dir
                / "group_01_iter_001_buffer_analysis_request.json"
            )
            session._write_json(request, {"folders": []})

            self.assertTrue(request.is_file())
            self.assertLess(len(str(request.resolve())) + 5, 260)
            self.assertLessEqual(len(session.record_dir.name), 16)


if __name__ == "__main__":
    unittest.main()
