"""
使用判定表/因果图方法测试 DocxProcessor.is_docx_file。

判定表：
条件：
- C1：扩展名为 .docx
- C2：文件存在
- C3：ZIP 可读
- C4：包含必需文件（[Content_Types].xml, _rels/.rels, word/document.xml）
结果：
- R1：返回 True

规则：
- R1：C1=T, C2=T, C3=T, C4=T => True
- R2：C1=F => False
- R3：C1=T, C2=F => False
- R4：C1=T, C2=T, C3=F => False
- R5：C1=T, C2=T, C3=T, C4=F => False

因果图（文字描述）：
- C1、C2、C3、C4 同时为真时，导致 R1 为真；
- 任一条件为假，导致 R1 为假。
"""

import os
import sys
import tempfile
import shutil
import unittest
import zipfile

# 添加 src 到路径以便导入
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

try:
    from docx import Document
    from core.docx_processor import DocxProcessor
except ImportError as e:
    print(f"Import error: {e}")
    print("Please install dependencies: pip install -r requirements.txt")
    sys.exit(1)


class TestDocxValidationDecisionTable(unittest.TestCase):
    """基于判定表的 DOCX 校验测试。"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

    def test_rule_r1_valid_docx(self):
        """R1: 所有条件为真 -> True。"""
        file_path = os.path.join(self.temp_dir, "valid.docx")
        doc = Document()
        doc.add_paragraph("Hello")
        doc.save(file_path)

        self.assertTrue(DocxProcessor.is_docx_file(file_path))

    def test_rule_r2_wrong_extension(self):
        """R2: 扩展名不正确 -> False。"""
        file_path = os.path.join(self.temp_dir, "invalid.txt")
        with open(file_path, "w", encoding="utf-8") as handle:
            handle.write("not a docx")

        self.assertFalse(DocxProcessor.is_docx_file(file_path))

    def test_rule_r3_missing_file(self):
        """R3: 文件不存在 -> False。"""
        file_path = os.path.join(self.temp_dir, "missing.docx")
        self.assertFalse(DocxProcessor.is_docx_file(file_path))

    def test_rule_r4_not_a_zip(self):
        """R4: ZIP 不可读 -> False。"""
        file_path = os.path.join(self.temp_dir, "not_zip.docx")
        with open(file_path, "wb") as handle:
            handle.write(b"not a zip file")

        self.assertFalse(DocxProcessor.is_docx_file(file_path))

    def test_rule_r5_missing_required_files(self):
        """R5: ZIP 可读但缺少必需文件 -> False。"""
        file_path = os.path.join(self.temp_dir, "missing_parts.docx")
        with zipfile.ZipFile(file_path, "w") as zip_ref:
            zip_ref.writestr("[Content_Types].xml", "<Types></Types>")

        self.assertFalse(DocxProcessor.is_docx_file(file_path))


if __name__ == "__main__":
    unittest.main(verbosity=2)
