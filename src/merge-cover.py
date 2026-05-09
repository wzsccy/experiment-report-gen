"""
merge-cover.py - 合并封面模板和正文内容

从模板 docx 提取封面页，替换正文内容生成新报告。

用法:
    python merge-cover.py <模板.docx> <正文.docx> <输出.docx>
"""

import sys
import os
from lxml import etree

# ── Word XML 命名空间 ──
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

def qn(tag):
    """将 w:tag 转换为 {namespace}tag"""
    if ':' not in tag:
        # 默认使用 Word 命名空间
        return f'{{{W}}}{tag}'
    prefix, local = tag.split(':')
    return f'{{{NS[prefix]}}}{local}'


def _extract_last_section(body_elem):
    """从 body 元素中提取最后一个 section 的内容（不含 sectPr）。

    如果文档有多个 section（封面+正文），只取最后一个 section。
    """
    # 查找所有 sectPr：直接子元素 或 嵌套在 w:p/w:pPr 中
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

    # 如果没有找到 section break，返回所有内容（单 section 文档）
    if first_sectpr_para is None and first_sectpr_direct is None:
        children = list(body_elem)
        last_sectpr = body_elem.find(qn('w:sectPr'))
        if last_sectpr is not None:
            children.remove(last_sectpr)
        return children

    # 单 section 文档：sectPr 是 w:body 的最后一个直接子元素
    # 所有内容都在 sectPr 之前，直接返回（去掉 sectPr）
    if first_sectpr_direct is not None:
        children = list(body_elem)
        children.remove(first_sectpr_direct)
        return children

    # 多 section 文档：sectPr 嵌套在段落中（封面结束标志）
    # 取该段落之后的内容
    result = []
    found = False
    for child in list(body_elem):
        if not found:
            if child is first_sectpr_para:
                found = True
                # 保留段落文本但去掉 sectPr
                pPr = child.find(qn('w:pPr'))
                sectPr = pPr.find(qn('w:sectPr'))
                pPr.remove(sectPr)
                result.append(child)
            continue
        if child.tag == qn('w:sectPr'):
            continue  # 跳过末尾的 sectPr
        result.append(child)

    return result


def main():
    if len(sys.argv) < 4:
        print('用法: python merge-cover.py <模板.docx> <正文.docx> <输出.docx>')
        sys.exit(1)

    template_path = sys.argv[1]
    body_path = sys.argv[2]
    output_path = sys.argv[3]

    # 验证文件存在
    for p in [template_path, body_path]:
        if not os.path.exists(p):
            print(f'错误: 文件不存在: {p}')
            sys.exit(1)

    # docx 是 zip 文件，直接操作 XML
    import zipfile
    import tempfile
    import shutil

    # 解压模板
    template_dir = tempfile.mkdtemp(prefix='docx_tpl_')
    with zipfile.ZipFile(template_path, 'r') as z:
        z.extractall(template_dir)

    # 解压正文
    body_dir = tempfile.mkdtemp(prefix='docx_body_')
    with zipfile.ZipFile(body_path, 'r') as z:
        z.extractall(body_dir)

    # 解析模板的 document.xml
    tpl_doc_path = os.path.join(template_dir, 'word', 'document.xml')
    body_doc_path = os.path.join(body_dir, 'word', 'document.xml')

    tpl_tree = etree.parse(tpl_doc_path)
    tpl_root = tpl_tree.getroot()
    tpl_body = tpl_root.find(qn('w:body'))

    body_tree = etree.parse(body_doc_path)
    body_root = body_tree.getroot()
    body_body = body_root.find(qn('w:body'))

    # ── 找到模板中的封面分节符 ──
    # 封面结束标志：段落中包含 w:sectPr（section break）
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
        # 清理
        shutil.rmtree(template_dir, ignore_errors=True)
        shutil.rmtree(body_dir, ignore_errors=True)
        sys.exit(1)

    # ── 清除模板正文（保留封面和分节符）──
    # 删除分节符之后、last_sectPr 之前的所有元素
    # 保留分节符段落，它是封面和正文的分界线
    last_sectPr = tpl_body.find(qn('w:sectPr'))

    to_remove = []
    found_break = False
    for child in list(tpl_body):
        if child is sect_break_para:
            found_break = True
            continue  # 保留分节符段落
        if found_break and child is not last_sectPr:
            to_remove.append(child)

    for elem in to_remove:
        tpl_body.remove(elem)

    # ── 从正文文档中提取 body section 的内容 ──
    # 如果正文有多个 section（封面+正文），只取最后一个 section 的内容
    body_children = _extract_last_section(body_body)

    # ── 从正文文档中提取 sectPr 的页面属性（页边距、页眉页脚距离等）──
    # 正文文档的 sectPr 包含 generateReport() 设置的自定义页边距，
    # 需要合并到模板的 last_sectPr 中，否则模板的默认边距会覆盖自定义设置
    body_sectPr = body_body.find(qn('w:sectPr'))
    if body_sectPr is not None:
        # 提取正文 sectPr 中的 pgMar（页边距）
        body_pgMar = body_sectPr.find(qn('w:pgMar'))
        if body_pgMar is not None:
            tpl_pgMar = last_sectPr.find(qn('w:pgMar'))
            if tpl_pgMar is not None:
                # 用正文的页边距覆盖模板的页边距
                for attr in ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']:
                    val = body_pgMar.get(qn(f'w:{attr}'))
                    if val is not None:
                        tpl_pgMar.set(qn(f'w:{attr}'), val)
            else:
                # 模板没有 pgMar，直接添加正文的
                last_sectPr.append(body_pgMar)
        # 提取正文 sectPr 中的 pgSz（页面尺寸）
        body_pgSz = body_sectPr.find(qn('w:pgSz'))
        if body_pgSz is not None:
            tpl_pgSz = last_sectPr.find(qn('w:pgSz'))
            if tpl_pgSz is not None:
                for attr in ['w', 'h', 'orient']:
                    val = body_pgSz.get(qn(f'w:{attr}'))
                    if val is not None:
                        tpl_pgSz.set(qn(f'w:{attr}'), val)

    # ── 将正文内容插入到模板中 ──
    # 在 last_sectPr 之前插入
    insert_idx = list(tpl_body).index(last_sectPr)
    for i, child in enumerate(body_children):
        tpl_body.insert(insert_idx + i, child)

    # ── 处理图片（_rels）──
    tpl_rels_path = os.path.join(template_dir, 'word', '_rels', 'document.xml.rels')
    body_rels_path = os.path.join(body_dir, 'word', '_rels', 'document.xml.rels')

    # 关系文件使用 PKG_REL 命名空间
    REL_NS = f'{{{PKG_REL}}}'

    if os.path.exists(body_rels_path):
        body_rels_tree = etree.parse(body_rels_path)
        body_rels_root = body_rels_tree.getroot()

        if os.path.exists(tpl_rels_path):
            tpl_rels_tree = etree.parse(tpl_rels_path)
            tpl_rels_root = tpl_rels_tree.getroot()
        else:
            tpl_rels_root = etree.Element(f'{REL_NS}Relationships',
                nsmap={'': PKG_REL})

        # 收集模板中已有的 rId
        existing_rids = set()
        for rel in tpl_rels_root.findall(f'{REL_NS}Relationship'):
            rid = rel.get('Id')
            if rid:
                existing_rids.add(rid)

        # 找到最大的 rId 编号
        max_rid = 0
        for rid in existing_rids:
            if rid.startswith('rId'):
                try:
                    num = int(rid[3:])
                    max_rid = max(max_rid, num)
                except ValueError:
                    pass

        # 复制正文中的图片关系和文件
        body_media_dir = os.path.join(body_dir, 'word', 'media')
        tpl_media_dir = os.path.join(template_dir, 'word', 'media')
        os.makedirs(tpl_media_dir, exist_ok=True)

        rid_map = {}  # body_rid -> tpl_rid

        for rel in body_rels_root.findall(f'{REL_NS}Relationship'):
            rel_type = rel.get('Type', '')
            if 'image' in rel_type:
                target = rel.get('Target', '')
                body_rid = rel.get('Id')

                # 检查是否已存在相同的图片
                found_existing = False
                for existing_rel in tpl_rels_root.findall(f'{REL_NS}Relationship'):
                    if existing_rel.get('Target') == target:
                        rid_map[body_rid] = existing_rel.get('Id')
                        found_existing = True
                        break

                if not found_existing:
                    max_rid += 1
                    new_rid = f'rId{max_rid}'
                    rid_map[body_rid] = new_rid

                    # 添加关系
                    new_rel = etree.SubElement(tpl_rels_root, f'{REL_NS}Relationship')
                    new_rel.set('Id', new_rid)
                    new_rel.set('Type', rel_type)
                    new_rel.set('Target', target)

                    # 复制图片文件
                    src_img = os.path.join(body_dir, 'word', target)
                    dst_img = os.path.join(template_dir, 'word', target)
                    if os.path.exists(src_img) and not os.path.exists(dst_img):
                        shutil.copy2(src_img, dst_img)

        # 保存关系文件
        tpl_rels_tree = etree.ElementTree(tpl_rels_root)
        tpl_rels_tree.write(tpl_rels_path, xml_declaration=True, encoding='UTF-8')

        # ── 更新 document.xml 中的图片引用 ──
        # 替换正文内容中的 rId
        for elem in body_children:
            for drawing in elem.iter(qn('w:drawing')):
                for blip in drawing.iter(qn('a:blip')):
                    embed = blip.get(qn('r:embed'))
                    if embed and embed in rid_map:
                        blip.set(qn('r:embed'), rid_map[embed])

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

    # 清理临时目录
    shutil.rmtree(template_dir, ignore_errors=True)
    shutil.rmtree(body_dir, ignore_errors=True)

    print(f'封面模板合并完成: {output_path}')


if __name__ == '__main__':
    main()
