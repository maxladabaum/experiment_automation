import tempfile
import unittest
import zipfile
from pathlib import Path

from core.session_archive import archive_session


class SessionArchiveTests(unittest.TestCase):
    def test_archive_preserves_source_and_publishes_complete_zip(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp_dir:
            root = Path(temp_dir)
            session = root / "session_001"
            experiment = session / "experiment_001"
            experiment.mkdir(parents=True)
            data_file = experiment / "measurement.csv"
            data_file.write_text("potential,current\n0.1,2.0\n", encoding="utf-8")

            result = archive_session(session, root / "remote")

            self.assertTrue(data_file.is_file())
            self.assertTrue(result.is_file())
            self.assertFalse(any(result.parent.glob("*.part")))
            with zipfile.ZipFile(result) as archive:
                self.assertEqual(
                    archive.read("session_001/experiment_001/measurement.csv"),
                    data_file.read_bytes(),
                )


if __name__ == "__main__":
    unittest.main()
