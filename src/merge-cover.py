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

    # ── 清除模板正文（保留封面）──
    # 删除分节符之后的所有元素（保留 body 末尾的 sectPr）
    last_sectPr = tpl_body.find(qn('w:sectPr'))

    to_remove = []
    found_break = False
    for child in list(tpl_body):
        if child is sect_break_para:
            found_break = True
            # 分节符段落本身也要删除（封面不需要它）
            to_remove.append(child)
            continue
        if found_break and child is not last_sectPr:
            to_remove.append(child)

    for elem in to_remove:
        tpl_body.remove(elem)

    # ── 从正文文档中提取所有内容 ──
    # 正文文档的第一个 section 的内容（标题、日期、表格等）
    body_children = list(body_body)

    # 移除正文的 sectPr（我们用模板的）
    body_sectPr = body_body.find(qn('w:sectPr'))
    if body_sectPr is not None:
        body_children.remove(body_sectPr)

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
        tpl_rels_tree.write(tpl_rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

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
    tpl_tree.write(tpl_doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)

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
