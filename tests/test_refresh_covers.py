from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'scripts'))
from refresh_covers import prepare, apply


class RefreshCoverTests(unittest.TestCase):
    def fixture(self, base):
        root=base/'articles'; article=root/'lesson'
        (article/'pages').mkdir(parents=True)
        Image.new('RGB',(1600,900),'blue').save(article/'pages/01.jpg')
        Image.new('RGB',(1880,800),'red').save(article/'公众号封面图.png')
        Image.new('RGB',(800,800),'red').save(article/'封面方形裁切检查.png')
        (article/'文章正文.md').write_text('reviewed body',encoding='utf-8')
        return root,article

    def test_apply_backs_up_originals_and_preserves_body(self):
        with TemporaryDirectory() as temp:
            base=Path(temp); root,article=self.fixture(base)
            original=(article/'公众号封面图.png').read_bytes()
            self.assertEqual(prepare(root,base/'review',{}),1)
            self.assertEqual(apply(base/'review',base/'backup'),1)
            self.assertEqual((base/'backup/lesson/公众号封面图.png').read_bytes(),original)
            self.assertEqual((article/'文章正文.md').read_text(encoding='utf-8'),'reviewed body')
            with Image.open(article/'公众号封面图.png') as cover:
                self.assertEqual(cover.size,(1880,800))
                self.assertNotEqual(cover.getpixel((0,0)),(255,0,0))

    def test_changed_source_stops_before_replacing_cover(self):
        with TemporaryDirectory() as temp:
            base=Path(temp); root,article=self.fixture(base)
            original=(article/'公众号封面图.png').read_bytes()
            prepare(root,base/'review',{})
            Image.new('RGB',(1600,900),'green').save(article/'pages/01.jpg')
            with self.assertRaisesRegex(ValueError,'来源已改变'):
                apply(base/'review',base/'backup')
            self.assertEqual((article/'公众号封面图.png').read_bytes(),original)
            self.assertFalse((base/'backup').exists())

    def test_repeated_apply_does_not_replace_existing_backup(self):
        with TemporaryDirectory() as temp:
            base=Path(temp); root,article=self.fixture(base)
            prepare(root,base/'review',{})
            apply(base/'review',base/'backup')
            original_backup=(base/'backup/lesson/公众号封面图.png').read_bytes()
            with self.assertRaises(ValueError): apply(base/'review',base/'backup')
            self.assertEqual((base/'backup/lesson/公众号封面图.png').read_bytes(),original_backup)


if __name__=='__main__': unittest.main()
