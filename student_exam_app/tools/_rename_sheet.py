"""把 xlsx 里的工作表改名（同步 definedName 引用），用于修正答案文件。"""
import os
import re
import shutil
import sys
import zipfile


def main():
    if len(sys.argv) < 4:
        print('用法: python tools/_rename_sheet.py 文件.xlsx 旧名 新名')
        return 1
    path, old, new = sys.argv[1], sys.argv[2], sys.argv[3]
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as src:
        parts = {i.filename: src.read(i.filename) for i in src.infolist()}
    wb = parts['xl/workbook.xml'].decode('utf-8')
    if 'name="%s"' % new in wb:
        print('已存在同名工作表，未修改')
        return 1
    if 'name="%s"' % old not in wb:
        print('未找到工作表 %s' % old)
        return 1
    wb = wb.replace('name="%s"' % old, 'name="%s"' % new, 1)
    wb = wb.replace('%s!' % old, '%s!' % new)
    parts['xl/workbook.xml'] = wb.encode('utf-8')
    for name, data in list(parts.items()):
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            text = data.decode('utf-8')
            if old + '!' in text:
                parts[name] = text.replace(old + '!', new + '!').encode('utf-8')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as dst:
        for name, data in parts.items():
            dst.writestr(name, data)
    shutil.move(tmp, path)
    print('已把 %s 改名为 %s' % (old, new))
    return 0


if __name__ == '__main__':
    sys.exit(main())
