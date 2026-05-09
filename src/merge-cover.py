"""
merge-cover.py - 合并封面模板和正文内容

从模板 docx 提取封面页，替换正文内容生成新报告。
安全处理：关系复制、rId 映射、Content_Types 补全、deepcopy sectPr。

用法:
    python merge-cover.py <模板.docx> <正文.docx> <输出.docx>
"""

import sys
import os
import copy
from lxml import etree

# ── Word XML 命名空间 ──
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

def qn(tag):
    """将 w:tag 转换为 {namespace}tag"""
    if ':' not in tag:
        return f'{{{W}}}{tag}'
    prefix, local = tag.split(':')
    return f'{{{NS[prefix]}}}{local}'


def _extract_last_section(body_elem):
    """从 body 元素中提取最后一个 section 的内容（不含 sectPr）。"""
    first_sectpr_para = None
    first_sectpr_direct = None

    for child in body_elem:
        if child.tag == qn('w:sectPr'):
            first_sectpr_direct = child
            break
        if child.tag == qn('w:p'):
            pPr = child.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                first_sectpr_para = child
                break

    if first_sectpr_para is None and first_sectpr_direct is None:
        children = list(body_elem)
        last_sectpr = body_elem.find(qn('w:sectPr'))
        if last_sectpr is not None:
            children.remove(last_sectpr)
        return children

    if first_sectpr_direct is not None:
        children = list(body_elem)
        children.remove(first_sectpr_direct)
        return children

    result = []
    found = False
    for child in list(body_elem):
        if not found:
            if child is first_sectpr_para:
                found = True
                pPr = child.find(qn('w:pPr'))
                sectPr = pPr.find(qn('w:sectPr'))
                pPr.remove(sectPr)
                result.append(child)
            continue
        if child.tag == qn('w:sectPr'):
            continue
        result.append(child)

    return result


def _collect_used_rids(body_children):
    """扫描正文 XML 中实际引用的所有 rId。"""
    used_rids = set()
    for elem in body_children:
        # r:embed (图片、OLE 等)
        for node in elem.iter():
            for attr_name in [qn('r:embed'), qn('r:id'), qn('r:link')]:
                val = node.get(attr_name)
                if val:
                    used_rids.add(val)
    return used_rids


def _get_max_rid(tpl_rels_root):
    """从模板 rels 中找到最大 rId 编号。"""
    max_rid = 0
    REL_NS = f'{{{PKG_REL}}}'
    for rel in tpl_rels_root.findall(f'{REL_NS}Relationship'):
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                num = int(rid[3:])
                max_rid = max(max_rid, num)
            except ValueError:
                pass
    return max_rid


def _ensure_content_types(template_dir, extensions):
    """确保 [Content_Types].xml 包含所需图片扩展名的 ContentType。"""
    ct_path = os.path.join(template_dir, '[Content_Types].xml')
    if not os.path.exists(ct_path):
        return

    EXT_TO_CT = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'tiff': 'image/tiff',
        'tif': 'image/tiff',
        'wmf': 'image/x-wmf',
        'emf': 'image/x-emf',
    }

    tree = etree.parse(ct_path)
    root = tree.getroot()
    existing_exts = set()
    for default in root.findall(f'{{{CT_NS}}}Default'):
        existing_exts.add(default.get('Extension', '').lower())

    modified = False
    for ext in extensions:
        ext_lower = ext.lower()
        if ext_lower not in existing_exts and ext_lower in EXT_TO_CT:
            new_default = etree.SubElement(root, f'{{{CT_NS}}}Default')
            new_default.set('Extension', ext_lower)
            new_default.set('ContentType', EXT_TO_CT[ext_lower])
            modified = True

    if modified:
        tree.write(ct_path, xml_declaration=True, encoding='UTF-8')


def main():
    if len(sys.argv) < 4:
        print('用法: python merge-cover.py <模板.docx> <正文.docx> <输出.docx>')
        sys.exit(1)

    template_path = sys.argv[1]
    body_path = sys.argv[2]
    output_path = sys.argv[3]

    for p in [template_path, body_path]:
        if not os.path.exists(p):
            print(f'错误: 文件不存在: {p}')
            sys.exit(1)

    import zipfile
    import tempfile
    import shutil

    template_dir = tempfile.mkdtemp(prefix='docx_tpl_')
    with zipfile.ZipFile(template_path, 'r') as z:
        z.extractall(template_dir)

    body_dir = tempfile.mkdtemp(prefix='docx_body_')
    with zipfile.ZipFile(body_path, 'r') as z:
        z.extractall(body_dir)

    tpl_doc_path = os.path.join(template_dir, 'word', 'document.xml')
    body_doc_path = os.path.join(body_dir, 'word', 'document.xml')

    tpl_tree = etree.parse(tpl_doc_path)
    tpl_root = tpl_tree.getroot()
    tpl_body = tpl_root.find(qn('w:body'))

    body_tree = etree.parse(body_doc_path)
    body_root = body_tree.getroot()
    body_body = body_root.find(qn('w:body'))

    # ── 找到模板中的封面分节符 ──
    sect_break_para = None
    for para in tpl_body.findall(qn('w:p')):
        pPr = para.find(qn('w:pPr'))
        if pPr is not None:
            sectPr = pPr.find(qn('w:sectPr'))
            if sectPr is not None:
                sect_break_para = para
                break

    if sect_break_para is None:
        print('错误: 模板中未找到封面分节符')
        shutil.rmtree(template_dir, ignore_errors=True)
        shutil.rmtree(body_dir, ignore_errors=True)
        sys.exit(1)

    # ── 清除模板正文（保留封面和分节符）──
    last_sectPr = tpl_body.find(qn('w:sectPr'))

    to_remove = []
    found_break = False
    for child in list(tpl_body):
        if child is sect_break_para:
            found_break = True
            continue
        if found_break and child is not last_sectPr:
            to_remove.append(child)

    for elem in to_remove:
        tpl_body.remove(elem)

    # ── 从正文文档中提取 body section 的内容 ──
    body_children = _extract_last_section(body_body)

    # ── 深拷贝正文 sectPr 的页面属性到模板 last_sectPr ──
    body_sectPr = body_body.find(qn('w:sectPr'))
    if body_sectPr is not None and last_sectPr is not None:
        body_pgMar = body_sectPr.find(qn('w:pgMar'))
        if body_pgMar is not None:
            tpl_pgMar = last_sectPr.find(qn('w:pgMar'))
            if tpl_pgMar is not None:
                for attr in ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']:
                    val = body_pgMar.get(qn(f'w:{attr}'))
                    if val is not None:
                        tpl_pgMar.set(qn(f'w:{attr}'), val)
            else:
                last_sectPr.append(copy.deepcopy(body_pgMar))
        body_pgSz = body_sectPr.find(qn('w:pgSz'))
        if body_pgSz is not None:
            tpl_pgSz = last_sectPr.find(qn('w:pgSz'))
            if tpl_pgSz is not None:
                for attr in ['w', 'h', 'orient']:
                    val = body_pgSz.get(qn(f'w:{attr}'))
                    if val is not None:
                        tpl_pgSz.set(qn(f'w:{attr}'), val)
            else:
                last_sectPr.append(copy.deepcopy(body_pgSz))

    # ── 处理关系：正文的所有关系都分配新 rId，图片复制为唯一文件名 ──
    tpl_rels_path = os.path.join(template_dir, 'word', '_rels', 'document.xml.rels')
    body_rels_path = os.path.join(body_dir, 'word', '_rels', 'document.xml.rels')
    REL_NS = f'{{{PKG_REL}}}'

    rid_map = {}  # body_rid -> new_tpl_rid
    used_extensions = set()

    # 扫描正文 XML 中实际引用的 rId
    used_rids = _collect_used_rids(body_children)

    if os.path.exists(body_rels_path):
        body_rels_tree = etree.parse(body_rels_path)
        body_rels_root = body_rels_tree.getroot()

        if os.path.exists(tpl_rels_path):
            tpl_rels_tree = etree.parse(tpl_rels_path)
            tpl_rels_root = tpl_rels_tree.getroot()
        else:
            tpl_rels_root = etree.Element(f'{REL_NS}Relationships',
                nsmap={'': PKG_REL})

        max_rid = _get_max_rid(tpl_rels_root)

        tpl_media_dir = os.path.join(template_dir, 'word', 'media')
        os.makedirs(tpl_media_dir, exist_ok=True)

        # 收集模板已有的 Target 路径（用于避免文件名冲突）
        existing_targets = set()
        for rel in tpl_rels_root.findall(f'{REL_NS}Relationship'):
            t = rel.get('Target', '')
            if t:
                existing_targets.add(t)

        body_media_dir = os.path.join(body_dir, 'word', 'media')
        img_counter = 0

        for rel in body_rels_root.findall(f'{REL_NS}Relationship'):
            body_rid = rel.get('Id', '')
            if body_rid not in used_rids:
                continue  # 跳过未被正文内容引用的关系

            rel_type = rel.get('Type', '')
            target = rel.get('Target', '')
            is_image = 'image' in rel_type

            max_rid += 1
            new_rid = f'rId{max_rid}'
            rid_map[body_rid] = new_rid

            if is_image:
                # 图片：复制为唯一文件名，避免与模板图片冲突
                img_counter += 1
                orig_ext = os.path.splitext(target)[1] or '.png'
                new_filename = f'body_img_{img_counter}{orig_ext}'
                new_target = f'media/{new_filename}'

                src_path = os.path.join(body_dir, 'word', target)
                dst_path = os.path.join(template_dir, 'word', new_target)
                if os.path.exists(src_path):
                    shutil.copy2(src_path, dst_path)

                used_extensions.add(orig_ext.lstrip('.'))

                new_rel = etree.SubElement(tpl_rels_root, f'{REL_NS}Relationship')
                new_rel.set('Id', new_rid)
                new_rel.set('Type', rel_type)
                new_rel.set('Target', new_target)
            else:
                # 非图片关系：直接复制（hyperlink、oleObject 等）
                new_rel = etree.SubElement(tpl_rels_root, f'{REL_NS}Relationship')
                new_rel.set('Id', new_rid)
                new_rel.set('Type', rel_type)
                new_rel.set('Target', target)
                # 如果 Target 是相对路径，复制对应文件
                if not target.startswith('http') and not target.startswith('#'):
                    src_path = os.path.join(body_dir, 'word', target)
                    dst_path = os.path.join(template_dir, 'word', target)
                    if os.path.exists(src_path) and not os.path.exists(dst_path):
                        os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                        shutil.copy2(src_path, dst_path)

        # 保存更新后的 rels
        tpl_rels_tree = etree.ElementTree(tpl_rels_root)
        tpl_rels_tree.write(tpl_rels_path, xml_declaration=True, encoding='UTF-8')

    # ── 先更新正文 XML 中的 rId 引用，再插入到模板 ──
    for elem in body_children:
        for node in elem.iter():
            for attr_name in [qn('r:embed'), qn('r:id'), qn('r:link')]:
                old_val = node.get(attr_name)
                if old_val and old_val in rid_map:
                    node.set(attr_name, rid_map[old_val])

    # ── 将正文内容插入到模板中 ──
    insert_idx = list(tpl_body).index(last_sectPr)
    for i, child in enumerate(body_children):
        tpl_body.insert(insert_idx + i, child)

    # ── 更新 [Content_Types].xml ──
    if used_extensions:
        _ensure_content_types(template_dir, used_extensions)

    # ── 保存结果 ──
    tpl_tree = etree.ElementTree(tpl_root)
    tpl_tree.write(tpl_doc_path, xml_declaration=True, encoding='UTF-8')

    # 重新打包为 docx
    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(template_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, template_dir)
                zout.write(file_path, arcname)

    shutil.rmtree(template_dir, ignore_errors=True)
    shutil.rmtree(body_dir, ignore_errors=True)

    print(f'封面模板合并完成: {output_path}')


if __name__ == '__main__':
    main()
