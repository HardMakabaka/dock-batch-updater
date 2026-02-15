"""Concurrency regression tests for backup naming.

PR1 target: ensure that concurrent backup creation with same-name source
files from different directories never collides, and every backup is
uniquely traceable to its source.
"""

import unittest
import os
import sys
import tempfile
import shutil
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docx import Document
from core.docx_processor import DocxProcessor


def _create_docx(path: str, content: str) -> None:
    """Helper: create a minimal DOCX with distinguishable content."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = Document()
    doc.add_paragraph(content)
    doc.save(path)


def _file_md5(path: str) -> str:
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


class TestBackupNamingUniqueness(unittest.TestCase):
    """Verify backup names are unique even for same-name source files."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="bktest_")
        self.backup_dir = os.path.join(self.root, "backups")
        os.makedirs(self.backup_dir)

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    # ── core scenario: two same-name files, shared backup dir ──

    def test_same_name_different_dirs_sequential(self):
        """Two report.docx from different dirs, sequential backup."""
        dir_a = os.path.join(self.root, "dirA")
        dir_b = os.path.join(self.root, "dirB")
        file_a = os.path.join(dir_a, "report.docx")
        file_b = os.path.join(dir_b, "report.docx")
        _create_docx(file_a, "Content A")
        _create_docx(file_b, "Content B")

        proc_a = DocxProcessor(file_a)
        proc_b = DocxProcessor(file_b)

        bk_a = proc_a.create_backup(self.backup_dir)
        bk_b = proc_b.create_backup(self.backup_dir)

        # Different backup paths
        self.assertNotEqual(bk_a, bk_b)
        # Both exist
        self.assertTrue(os.path.exists(bk_a))
        self.assertTrue(os.path.exists(bk_b))
        # Content matches respective source
        self.assertEqual(_file_md5(file_a), _file_md5(bk_a))
        self.assertEqual(_file_md5(file_b), _file_md5(bk_b))

    def test_same_name_different_dirs_concurrent(self):
        """Two report.docx from different dirs, concurrent backup via thread pool."""
        dir_a = os.path.join(self.root, "部门A")
        dir_b = os.path.join(self.root, "部门B")
        file_a = os.path.join(dir_a, "report.docx")
        file_b = os.path.join(dir_b, "report.docx")
        _create_docx(file_a, "Content A - 部门A")
        _create_docx(file_b, "Content B - 部门B")

        results = {}

        def do_backup(src_path):
            proc = DocxProcessor(src_path)
            return proc.create_backup(self.backup_dir)

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {pool.submit(do_backup, f): f for f in [file_a, file_b]}
            for fut in as_completed(futures):
                src = futures[fut]
                results[src] = fut.result()

        backup_paths = list(results.values())
        # Unique paths
        self.assertEqual(len(set(backup_paths)), 2)
        # Both exist and match source content
        for src, bk in results.items():
            self.assertTrue(os.path.exists(bk), f"Backup missing: {bk}")
            self.assertEqual(_file_md5(src), _file_md5(bk))

    # ── stress: N same-name files, high concurrency ──

    def test_many_same_name_files_concurrent(self):
        """10 same-name files from 10 dirs, 4 workers, shared backup dir."""
        n = 10
        files = []
        for i in range(n):
            d = os.path.join(self.root, f"dir_{i}")
            f = os.path.join(d, "data.docx")
            _create_docx(f, f"Content {i}")
            files.append(f)

        backup_paths = []

        def do_backup(src_path):
            proc = DocxProcessor(src_path)
            return proc.create_backup(self.backup_dir)

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = [pool.submit(do_backup, f) for f in files]
            for fut in as_completed(futures):
                backup_paths.append(fut.result())

        # All unique
        self.assertEqual(len(set(backup_paths)), n,
                         f"Expected {n} unique backups, got {len(set(backup_paths))}")
        # All exist
        for bp in backup_paths:
            self.assertTrue(os.path.exists(bp))

    def test_repeated_backup_same_file(self):
        """Calling create_backup twice on the same source yields two distinct backups."""
        d = os.path.join(self.root, "single")
        f = os.path.join(d, "doc.docx")
        _create_docx(f, "Same file")

        proc = DocxProcessor(f)
        bk1 = proc.create_backup(self.backup_dir)
        bk2 = proc.create_backup(self.backup_dir)

        self.assertNotEqual(bk1, bk2)
        self.assertTrue(os.path.exists(bk1))
        self.assertTrue(os.path.exists(bk2))


class TestBackupNamingTraceability(unittest.TestCase):
    """Verify backup names contain traceable hints."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="bktrace_")
        self.backup_dir = os.path.join(self.root, "backups")
        os.makedirs(self.backup_dir)

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_backup_name_contains_parent_hint(self):
        """Backup filename should contain sanitized parent directory name."""
        d = os.path.join(self.root, "财务部")
        f = os.path.join(d, "report.docx")
        _create_docx(f, "财务报告")

        proc = DocxProcessor(f)
        bk = proc.create_backup(self.backup_dir)
        bk_name = os.path.basename(bk)

        self.assertIn("财务部", bk_name)

    def test_backup_name_contains_hash(self):
        """Backup filename should contain a 6-char hash of the source path."""
        d = os.path.join(self.root, "testdir")
        f = os.path.join(d, "file.docx")
        _create_docx(f, "test")

        expected_hash = hashlib.sha256(
            os.path.abspath(f).encode("utf-8")
        ).hexdigest()[:6]

        proc = DocxProcessor(f)
        bk = proc.create_backup(self.backup_dir)
        bk_name = os.path.basename(bk)

        self.assertIn(expected_hash, bk_name)

    def test_backup_name_sanitizes_special_chars(self):
        """Parent dir with special chars should be sanitized in backup name."""
        d = os.path.join(self.root, "dir (copy) [2]")
        f = os.path.join(d, "file.docx")
        _create_docx(f, "test")

        proc = DocxProcessor(f)
        bk = proc.create_backup(self.backup_dir)
        bk_name = os.path.basename(bk)

        # No parens, brackets, or spaces in the relhint portion
        # The name should still be a valid filename
        self.assertTrue(os.path.exists(bk))
        # Relhint should be sanitized (only alnum, underscore, CJK)
        parts = bk_name.split("_backup_")
        self.assertEqual(len(parts), 2)
        relhint_and_rest = parts[1]  # e.g. "dir_copy_2_a3f2c1_e7b4d9f0.docx"
        # Should not contain original special chars
        for ch in "()[] ":
            self.assertNotIn(ch, relhint_and_rest)

    def test_root_dir_fallback(self):
        """Source file at filesystem root should use _root_ as hint."""
        # Simulate by creating a file whose parent basename is empty
        # We can't easily test actual root, so test the edge case where
        # parent dir name sanitizes to empty string
        d = os.path.join(self.root, "---")  # sanitizes to empty
        f = os.path.join(d, "file.docx")
        _create_docx(f, "test")

        proc = DocxProcessor(f)
        bk = proc.create_backup(self.backup_dir)
        bk_name = os.path.basename(bk)

        self.assertIn("_root_", bk_name)


class TestBackupAtomicity(unittest.TestCase):
    """Verify atomic write behavior: no partial/corrupt backups left behind."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="bkatom_")
        self.backup_dir = os.path.join(self.root, "backups")
        os.makedirs(self.backup_dir)

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_no_tmp_files_left_after_success(self):
        """After successful backup, no .tmp files should remain."""
        d = os.path.join(self.root, "src")
        f = os.path.join(d, "clean.docx")
        _create_docx(f, "clean")

        proc = DocxProcessor(f)
        proc.create_backup(self.backup_dir)

        tmp_files = [x for x in os.listdir(self.backup_dir) if x.endswith(".tmp")]
        self.assertEqual(len(tmp_files), 0, f"Leftover tmp files: {tmp_files}")

    def test_backup_dir_auto_created(self):
        """Backup dir is created automatically if it doesn't exist."""
        new_dir = os.path.join(self.root, "new", "nested", "backups")
        d = os.path.join(self.root, "src")
        f = os.path.join(d, "file.docx")
        _create_docx(f, "test")

        proc = DocxProcessor(f)
        bk = proc.create_backup(new_dir)

        self.assertTrue(os.path.exists(bk))
        self.assertTrue(os.path.isdir(new_dir))

    def test_default_backup_dir_is_source_dir(self):
        """When no backup_dir specified, backup goes to source file's directory."""
        d = os.path.join(self.root, "src")
        f = os.path.join(d, "file.docx")
        _create_docx(f, "test")

        proc = DocxProcessor(f)
        bk = proc.create_backup()  # no backup_dir

        self.assertEqual(os.path.dirname(bk), d)
        self.assertTrue(os.path.exists(bk))


if __name__ == '__main__':
    unittest.main()
