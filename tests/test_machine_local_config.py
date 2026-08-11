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
            "EA_BO_CONFIG_DIR",
            "EA_BO_DEFAULT_CONFIG_PATH",
            "EA_BO_LOCAL_PATHS_CONFIG",
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
            "'bo_config_dir': str(config.BO_CONFIG_DIR), "
            "'bo_default_config': str(config.BO_DEFAULT_CONFIG_PATH), "
            "'bo_local_paths': str(config.BO_LOCAL_PATHS_CONFIG), "
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
        self.assertEqual(Path(values["bo_config_dir"]), expected / "bo_configs")
        self.assertEqual(
            Path(values["bo_default_config"]),
            expected / "bo_configs" / "default_swv_bo.json",
        )
        self.assertEqual(
            Path(values["bo_local_paths"]),
            expected / "bo_configs" / "local_paths.json",
        )

    def test_external_config_and_environment_precedence(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "local_config.json"
            configured_data = Path(temp_dir) / "configured"
            config_path.write_text(
                json.dumps(
                    {
                        "data_dir": str(configured_data),
                        "bo_config_dir": str(configured_data / "bo"),
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
        self.assertEqual(Path(values["bo_config_dir"]), configured_data / "bo")
        self.assertEqual(
            Path(values["bo_default_config"]),
            configured_data / "bo" / "default_swv_bo.json",
        )
        self.assertEqual(values["pump"], [12, 9600, 0])
        self.assertEqual(values["potentiostat"], "COM13")

    def test_missing_default_bo_config_bootstraps_local_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "bo_configs" / "default_swv_bo.json"
            env = dict(os.environ)
            env["EA_BO_DEFAULT_CONFIG_PATH"] = str(config_path)
            code = (
                "import json; "
                "from pathlib import Path; "
                "from core.bo_session import load_bo_config; "
                "cfg = load_bo_config(); "
                "print(json.dumps({"
                "'exists': Path(%r).exists(), "
                "'channels': cfg.get('channels'), "
                "'has_initial_parameters': bool(cfg.get('initial_parameters'))"
                "}))"
            ) % str(config_path)
            completed = subprocess.run(
                [sys.executable, "-c", code],
                cwd=Path.cwd(),
                env=env,
                capture_output=True,
                text=True,
                check=True,
            )
            values = json.loads(completed.stdout)
        self.assertTrue(values["exists"])
        self.assertTrue(values["channels"])
        self.assertTrue(values["has_initial_parameters"])


if __name__ == "__main__":
    unittest.main()
