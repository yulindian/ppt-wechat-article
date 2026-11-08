"""Prepare or apply landscape covers without rebuilding article bodies.

Prepare into a NEW review directory, inspect images, then --apply that manifest.
The manifest fingerprints sources and existing covers to prevent stale overwrites.
"""
import argparse
import hashlib
import json
from pathlib import Path
from shutil import copy2
from PIL import Image
from build_article import make_cover

NAMES = ('公众号封面图.png', '封面方形裁切检查.png')

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def prepare(root, review, focus_map):
    if review.exists() and any(review.iterdir()):
        raise ValueError('审核目录须为新的或空的目录')
    review.mkdir(parents=True, exist_ok=True)
    records = []
    folders = sorted(p.parent for p in root.rglob(NAMES[0]) if (p.parent/'pages/01.jpg').is_file())
    if not folders: raise ValueError('未发现包含pages/01.jpg的文章目录')
    for n, folder in enumerate(folders, 1):
        rel = folder.relative_to(root).as_posix()
        destination = review/f'{n:03}'
        destination.mkdir()
        y = float(focus_map.get(rel, 0.0))
        with Image.open(folder/'pages/01.jpg') as im:
            make_cover(im.convert('RGB'), destination/NAMES[0], None, y)
        records.append(dict(id=n, relative_folder=rel, focus_y=y,
            source_sha=digest(folder/'pages/01.jpg'),
            old_sha={name:digest(folder/name) if (folder/name).exists() else None for name in NAMES},
            new_sha={name:digest(destination/name) for name in NAMES},
            body_sha={name:digest(folder/name) for name in ('文章正文.md','公众号文章导入稿.docx') if (folder/name).exists()}))
    manifest = dict(root=str(root), records=records)
    (review/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    return len(records)

def apply(review, backup):
    manifest=json.loads((review/'manifest.json').read_text(encoding='utf-8'))
    root=Path(manifest['root']).resolve()
    if backup.exists() and any(backup.iterdir()): raise ValueError('备份目录须为新的或空的目录')
    # Validate the entire batch before the first replacement.
    for rec in manifest['records']:
        folder=(root/rec['relative_folder']).resolve()
        if not folder.is_relative_to(root): raise ValueError('目录超出文章根路径')
        if digest(folder/'pages/01.jpg') != rec['source_sha']: raise ValueError('来源已改变，请重新准备')
        for name, expected in rec['old_sha'].items():
            actual=digest(folder/name) if (folder/name).exists() else None
            if actual != expected: raise ValueError('原封面已改变，请重新准备')
        for name, expected in rec['new_sha'].items():
            path=review/f"{rec['id']:03}"/name
            if digest(path) != expected: raise ValueError('审核封面已改变，请重新准备')
            with Image.open(path) as im:
                if im.size != ((1880,800) if name == NAMES[0] else (800,800)):
                    raise ValueError('封面尺寸不符')
    backup.mkdir(parents=True,exist_ok=True)
    for rec in manifest['records']:
        folder=root/rec['relative_folder']
        old=backup/rec['relative_folder']
        old.mkdir(parents=True)
        for name in NAMES:
            if (folder/name).exists(): copy2(folder/name,old/name)
            copy2(review/f"{rec['id']:03}"/name,folder/name)
        for name, sha in rec['body_sha'].items():
            if digest(folder/name) != sha: raise ValueError('正文变化，请检查')
        rec['applied']=True
    (backup/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    return len(manifest['records'])

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',type=Path)
    parser.add_argument('--review',type=Path,required=True)
    parser.add_argument('--focus-map',type=Path)
    parser.add_argument('--apply',action='store_true')
    parser.add_argument('--backup',type=Path)
    args=parser.parse_args()
    if args.apply:
        if args.backup is None: parser.error('--apply须指定--backup')
        count=apply(args.review.resolve(),args.backup.resolve())
    else:
        if args.root is None: parser.error('准备封面须指定--root')
        focus=json.loads(args.focus_map.read_text(encoding='utf-8')) if args.focus_map else {}
        count=prepare(args.root.resolve(),args.review.resolve(),focus)
    print(f'{count} covers '+('applied; originals backed up' if args.apply else 'prepared for visual review'))
