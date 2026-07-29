import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class MachineLocalConfigTests(unittest.TestCase):
    def _read_config(self, local_config_path, extra_env=None):
        env = dict(os.environ)
        for key in (
            "EA_DATA_DIR",
            "EA_METHODS_DIR",
            "EA_RECIPE_DIR",
            "EA_PUMP_COM_PORT",
            "EA_PUMP_BAUD",
            "EA_PUMP_DEV",
            "EA_POTENTIOSTAT_PORT",
        ):
            env.pop(key, None)
        env["EA_LOCAL_CONFIG_PATH"] = str(local_config_path)
        env.update(extra_env or {})
        code = (
            "import json, config; "
            "print(json.dumps({"
            "'data': str(config.DATA_DIR), "
            "'methods': str(config.METHODS_DIR), "
            "'recipe': str(config.RECIPE_DIR), "
            "'pump': [config.PUMP_DEFAULT_COM_PORT, config.PUMP_DEFAULT_BAUD, config.PUMP_DEFAULT_DEV], "
            "'potentiostat': config.DEVICE_DEFAULT_PORT"
            "}))"
        )
        completed = subprocess.run(
            [sys.executable, "-c", code],
            cwd=Path.cwd(),
            env=env,
            capture_output=True,
            text=True,
            check=True,
        )
        return json.loads(completed.stdout)

    def test_missing_external_config_uses_current_user_documents(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            values = self._read_config(Path(temp_dir) / "missing.json")
        expected = Path.home() / "Documents" / "Experiment Automation Data"
        self.assertEqual(Path(values["data"]), expected)
        self.assertEqual(Path(values["methods"]), expected / "methods")
        self.assertEqual(Path(values["recipe"]), expected / "recipe_maker")

    def test_external_config_and_environment_precedence(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "local_config.json"
            configured_data = Path(temp_dir) / "configured"
            config_path.write_text(
                json.dumps(
                    {
                        "data_dir": str(configured_data),
                        "pump_com_port": "COM12",
                        "pump_baud": 9600,
                        "pump_dev": 0,
                        "potentiostat_port": "COM13",
                    }
                ),
                encoding="utf-8",
            )
            override = Path(temp_dir) / "environment-override"
            values = self._read_config(
                config_path,
                extra_env={"EA_DATA_DIR": str(override)},
            )
        self.assertEqual(Path(values["data"]), override)
        self.assertEqual(Path(values["methods"]), override / "methods")
        self.assertEqual(values["pump"], [12, 9600, 0])
        self.assertEqual(values["potentiostat"], "COM13")


if __name__ == "__main__":
    unittest.main()
