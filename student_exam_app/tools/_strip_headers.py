"""移除 docx 的页眉/页脚（引用、部件、关系、内容类型），用于清理被比较文档污染的素材。"""
import os
import re
import shutil
import sys
import zipfile


def main():
    if len(sys.argv) < 2:
        print('用法: python tools/_strip_headers.py 文件.docx')
        return 1
    path = sys.argv[1]
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as src:
        parts = {i.filename: src.read(i.filename) for i in src.infolist()}
        infos = src.infolist()

    doc = parts['word/document.xml'].decode('utf-8')
    doc = re.sub(r'<w:(header|footer)Reference [^>]*/>', '', doc)
    parts['word/document.xml'] = doc.encode('utf-8')

    refs = re.findall(r'Target="(header\d+\.xml|footer\d+\.xml)"',
                      parts.get('word/_rels/document.xml.rels', b'').decode('utf-8'))
    rels = parts.get('word/_rels/document.xml.rels', b'').decode('utf-8')
    rels = re.sub(r'<Relationship [^>]*Target="(?:header|footer)\d+\.xml"[^>]*/>', '', rels)
    parts['word/_rels/document.xml.rels'] = rels.encode('utf-8')

    drop = set()
    for r in refs:
        drop.add('word/' + r)
        drop.add('word/_rels/' + r + '.rels')
    ct = parts.get('[Content_Types].xml', b'').decode('utf-8')
    if drop:
        for name in drop:
            parts.pop(name, None)
        ct = re.sub(r'<Override PartName="/word/(?:header|footer)\d+\.xml"[^>]*/>', '', ct)
        parts['[Content_Types].xml'] = ct.encode('utf-8')

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as dst:
        for info in infos:
            if info.filename in drop:
                continue
            dst.writestr(info, parts[info.filename])
    shutil.move(tmp, path)
    print('已移除:', sorted(drop) or '无引用')
    return 0


if __name__ == '__main__':
    sys.exit(main())
