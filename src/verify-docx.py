"""
verify-docx.py - 验证 docx 文件结构完整性

检查:
1. 能否作为 zip 打开
2. 所有 XML 能否被 lxml 解析
3. document.xml 中所有 r:embed/r:id 是否在 rels 中有定义
4. 所有 image relationship 的 Target 文件是否存在
5. [Content_Types].xml 是否包含图片扩展名

用法:
    python verify-docx.py <文件.docx>
"""

import sys
import os
import zipfile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

def qn(tag):
    if ':' not in tag:
        return f'{{{W}}}{tag}'
    prefix, local = tag.split(':')
    ns = {'w': W, 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': R_NS}
    return f'{{{ns[prefix]}}}{local}'


def verify(docx_path):
    errors = []
    warnings = []

    # 1. 能否作为 zip 打开
    try:
        z = zipfile.ZipFile(docx_path, 'r')
    except Exception as e:
        print(f'FAIL: 无法作为 zip 打开: {e}')
        return False

    names = set(z.namelist())

    # 2. 检查所有 XML 能否解析
    xml_files = [n for n in names if n.endswith('.xml') or n.endswith('.rels')]
    for xml_name in xml_files:
        try:
            data = z.read(xml_name)
            etree.fromstring(data)
        except Exception as e:
            errors.append(f'XML 解析失败: {xml_name}: {e}')

    # 3. 检查 document.xml 中的 rId 引用
    if 'word/document.xml' not in names:
        errors.append('缺少 word/document.xml')
        z.close()
        return len(errors) == 0

    doc_xml = z.read('word/document.xml')
    doc_tree = etree.fromstring(doc_xml)

    # 收集所有引用的 rId
    used_rids = set()
    for elem in doc_tree.iter():
        for attr in [qn('r:embed'), qn('r:id'), qn('r:link')]:
            val = elem.get(attr)
            if val:
                used_rids.add(val)

    # 读取 rels
    rels_path = 'word/_rels/document.xml.rels'
    rels_map = {}  # rId -> Target
    image_targets = []

    if rels_path in names:
        rels_xml = z.read(rels_path)
        rels_tree = etree.fromstring(rels_xml)
        REL_NS = f'{{{PKG_REL}}}'
        for rel in rels_tree.findall(f'{REL_NS}Relationship'):
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            rel_type = rel.get('Type', '')
            rels_map[rid] = target
            if 'image' in rel_type:
                image_targets.append((rid, target))

        # 检查引用的 rId 是否都有定义
        for rid in used_rids:
            if rid not in rels_map:
                errors.append(f'rId 引用未定义: {rid} (在 document.xml 中引用但 rels 中无此 ID)')
    else:
        errors.append(f'缺少 {rels_path}')

    # 4. 检查图片文件是否存在
    for rid, target in image_targets:
        if target.startswith('/'):
            file_path = target.lstrip('/')
        else:
            file_path = f'word/{target}'
        if file_path not in names:
            errors.append(f'图片文件缺失: {file_path} (rId={rid}, Target={target})')

    # 5. 检查 [Content_Types].xml
    if '[Content_Types].xml' in names:
        ct_xml = z.read('[Content_Types].xml')
        ct_tree = etree.fromstring(ct_xml)
        ct_exts = set()
        for default in ct_tree.findall(f'{{{CT_NS}}}Default'):
            ct_exts.add(default.get('Extension', '').lower())

        img_exts_needed = set()
        for rid, target in image_targets:
            ext = os.path.splitext(target)[1].lstrip('.').lower()
            if ext:
                img_exts_needed.add(ext)

        for ext in img_exts_needed:
            if ext not in ct_exts:
                warnings.append(f'[Content_Types].xml 缺少图片扩展名: {ext}')
    else:
        warnings.append('缺少 [Content_Types].xml')

    z.close()

    # 输出结果
    if errors:
        print(f'ERRORS ({len(errors)}):')
        for e in errors:
            print(f'  - {e}')

    if warnings:
        print(f'WARNINGS ({len(warnings)}):')
        for w in warnings:
            print(f'  - {w}')

    if not errors and not warnings:
        print('PASS: 所有检查通过')
    elif not errors:
        print('PASS: 有警告但无错误')

    return len(errors) == 0


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python verify-docx.py <文件.docx>')
        sys.exit(1)
    ok = verify(sys.argv[1])
    sys.exit(0 if ok else 1)
