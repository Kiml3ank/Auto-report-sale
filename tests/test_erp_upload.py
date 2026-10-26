"""python -m unittest discover -s tests -v"""
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import erp_upload  # noqa: E402


def pdf(path, text="pdf"):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


class Folders(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.signed = Path(self.tmp.name)
        self.uploaded = self.signed / "uploaded"

    def tearDown(self):
        self.tmp.cleanup()

    def test_pending_skips_empty_and_files(self):
        pdf(self.signed / "10-21-2026" / "a.pdf")
        (self.signed / "10-22-2026").mkdir()
        pdf(self.uploaded / "10-20-2026" / "b.pdf")
        pdf(self.signed / "x.pdf")
        self.assertEqual(erp_upload.pending_folders(self.signed), [self.signed / "10-21-2026"])

    def test_pending_sorted_by_date(self):
        for name in ("10-23-2026", "10-9-2026", "misc"):
            pdf(self.signed / name / "a.pdf")
        self.assertEqual([p.name for p in erp_upload.pending_folders(self.signed)], ["10-9-2026", "10-23-2026", "misc"])

    def test_move_done_moves(self):
        pdf(self.signed / "10-21-2026" / "a.pdf")
        to = erp_upload.move_done(self.signed / "10-21-2026", self.uploaded)
        self.assertEqual(to, self.uploaded / "10-21-2026")
        self.assertTrue((to / "a.pdf").exists())
        self.assertFalse((self.signed / "10-21-2026").exists())

    def test_move_done_merges(self):
        pdf(self.uploaded / "10-21-2026" / "a.pdf", "old")
        pdf(self.signed / "10-21-2026" / "a.pdf", "new")
        pdf(self.signed / "10-21-2026" / "b.pdf")
        to = erp_upload.move_done(self.signed / "10-21-2026", self.uploaded)
        self.assertEqual((to / "a.pdf").read_text(), "old")
        self.assertTrue((to / "b.pdf").exists())
        self.assertFalse((self.signed / "10-21-2026").exists())


if __name__ == "__main__":
    unittest.main()
