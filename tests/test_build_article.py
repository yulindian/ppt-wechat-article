import importlib.util
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest

from docx import Document
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("build_article", ROOT / "scripts/build_article.py")
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)


class ArticleTests(unittest.TestCase):
    def args(self, root, **changes):
        values = dict(source=ROOT / "examples/source", metadata=ROOT / "examples/metadata.json",
                      output=root / "result", pdf=None, cover=None, opening_image=None, closing_image=None)
        values.update(changes)
        return SimpleNamespace(**values)

    def test_custom_banner_is_normalized_and_all_pages_are_embedded(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            custom = root / "custom.jpg"
            Image.new("RGB", (900, 300), "#aa3300").save(custom)
            out = builder.build(self.args(root, opening_image=custom, closing_image=custom))
            with Image.open(out / "开头引导图.png") as image:
                self.assertEqual(image.format, "PNG")
                self.assertEqual(image.size, (900, 300))
            doc = Document(out / "公众号文章导入稿.docx")
            self.assertEqual(len(doc.inline_shapes), 5)
            self.assertEqual(len(list((out / "pages").glob("*.jpg"))), 3)
            lead = next(p for p in doc.paragraphs if "下面按照上课顺序" in p.text)
            self.assertTrue(any(r.bold and "配套课件" in r.text for r in lead.runs))
            self.assertIn("以上是最终PPT的完整预览，共3页", "\n".join(p.text for p in doc.paragraphs))
            self.assertTrue((out / "文章正文.md").is_file())

    def test_missing_resource_stops_before_output(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            config = json.loads((ROOT / "examples/metadata.json").read_text(encoding="utf-8"))
            config["resources"][0]["file"] = "missing.pdf"
            metadata = root / "bad.json"
            metadata.write_text(json.dumps(config), encoding="utf-8")
            with self.assertRaises(FileNotFoundError):
                builder.build(self.args(root, metadata=metadata))
            self.assertFalse((root / "result").exists())

    def test_existing_output_is_preserved(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "result").mkdir()
            marker = root / "result/keep.txt"
            marker.write_text("reviewed", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                builder.build(self.args(root))
            self.assertEqual(marker.read_text(), "reviewed")


if __name__ == "__main__":
    unittest.main()
