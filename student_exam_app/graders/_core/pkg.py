"""Office 包（docx/xlsx/pptx）的通用解包与关系解析。"""
import posixpath
import zipfile
import xml.etree.ElementTree as ET

PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_DOC = '/officeDocument'


class Package:
    def __init__(self, path):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self.names = self.zf.namelist()
        self._names = set(self.names)
        self._xml = {}
        self._rels = {}

    def close(self):
        self.zf.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def has(self, name):
        return name in self._names

    def read(self, name):
        return self.zf.read(name).decode('utf-8', 'replace')

    def root(self, name):
        if name not in self._xml:
            self._xml[name] = ET.fromstring(self.zf.read(name))
        return self._xml[name]

    def parts(self, prefix='', suffix='.xml', exclude=('_rels/',)):
        out = []
        for n in self.names:
            if not n.startswith(prefix) or not n.endswith(suffix):
                continue
            if any(x in n for x in exclude):
                continue
            out.append(n)
        return sorted(out)

    def rels(self, part):
        """返回 {rId: (Type, 包内绝对路径)}，Target 已按源部件目录归一。"""
        if part in self._rels:
            return self._rels[part]
        base = posixpath.dirname(part)
        rels_name = posixpath.join(base, '_rels', posixpath.basename(part) + '.rels')
        table = {}
        if rels_name in self._names:
            for rel in self.root(rels_name).iter('{%s}Relationship' % PKG):
                target = rel.get('Target', '')
                if rel.get('TargetMode') == 'External':
                    resolved = target
                elif target.startswith('/'):
                    resolved = target.lstrip('/')
                else:
                    resolved = posixpath.normpath(posixpath.join(base, target))
                table[rel.get('Id', '')] = (rel.get('Type', ''), resolved)
        self._rels[part] = table
        return table

    def rel_target(self, part, rid):
        entry = self.rels(part).get(rid)
        return entry[1] if entry else ''

    def related_of_type(self, part, type_suffix):
        out = []
        for rtype, target in self.rels(part).values():
            if rtype.endswith(type_suffix) and self.has(target):
                out.append(target)
        return out

    def main_part(self, default, type_suffix=OFFICE_DOC):
        for rtype, target in self.rels('').values():
            if rtype.endswith(type_suffix) and self.has(target):
                return target
        return default
