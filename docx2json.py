#!/usr/bin/env python3
"""
python docx2json.py template.docx
docx2json.py — 将任意 .docx 转为一份完整、自描述的 JSON 结构。
按页面对象块（段落/表格/图片）从上到下记录位置、样式、内容。

输出两份文件:
  1) full.json   — 完整格式（机器用 + restore 脚本用）
  2) brief.json  — 精简版（LLM 用，只保留语义关键信息）

JSON Schema 设计原则:
  - 每个块有唯一的 index（body-children 顺序）
  - 样式分三层: docDefaults → style 定义 → run 显式覆盖
  - 属性名直译 OOXML，不做"友好化"转换（避免歧义）
  - 缺省值用 null，明确区分"未设置"和"设置为默认值"
"""

import zipfile, json, os, sys
from lxml import etree

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_DRAW = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
NS_PIC = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_RELS = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def _a(el, attr, ns=NS):
    """Get attribute value or None."""
    v = el.get(f'{ns}{attr}')
    return v if v else None

def _int(v):
    return int(v) if v else None

def _bool(v):
    """OOXML boolean: attribute missing means True (element exists = ON).
       Only explicit val="0"/"false" means False."""
    # v is the attribute value; None means <w:b/> without w:val — that's ON
    return v is None or v not in ('0', 'false', 'False')


# ─── Style extraction ─────────────────────────────────────

def extract_style_rpr(rpr_elem):
    """Extract run properties from a <w:rPr> element."""
    if rpr_elem is None:
        return None
    props = {}
    # Font — 记录字体名和 hint 属性
    rf = rpr_elem.find(f'{NS}rFonts')
    if rf is not None:
        f = {}
        for k in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            v = _a(rf, k)
            if v:
                f[k] = v
        hint = _a(rf, 'hint')
        if hint:
            f['hint'] = hint
        if f:
            props['font'] = f
    # Size
    sz = rpr_elem.find(f'{NS}sz')
    if sz is not None:
        props['sz_hp'] = _int(_a(sz, 'val'))      # half-points (e.g. 22 = 11pt)
    szCs = rpr_elem.find(f'{NS}szCs')
    if szCs is not None:
        props['szCs_hp'] = _int(_a(szCs, 'val'))
    # Color
    color = rpr_elem.find(f'{NS}color')
    if color is not None:
        props['color'] = _a(color, 'val')          # RGB hex without #
        tc = _a(color, 'themeColor')
        if tc:
            props['color_theme'] = tc
    # Bold
    b = rpr_elem.find(f'{NS}b')
    if b is not None:
        props['bold'] = _bool(_a(b, 'val'))
    bCs = rpr_elem.find(f'{NS}bCs')
    if bCs is not None:
        props['bold_cjk'] = _bool(_a(bCs, 'val'))
    # Italic
    i = rpr_elem.find(f'{NS}i')
    if i is not None:
        props['italic'] = _bool(_a(i, 'val'))
    iCs = rpr_elem.find(f'{NS}iCs')
    if iCs is not None:
        props['italic_cjk'] = _bool(_a(iCs, 'val'))
    # Underline
    u = rpr_elem.find(f'{NS}u')
    if u is not None:
        props['underline'] = _a(u, 'val') or 'single'
    # Highlight
    hl = rpr_elem.find(f'{NS}highlight')
    if hl is not None:
        props['highlight'] = _a(hl, 'val')
    # Vertical alignment (superscript/subscript)
    va = rpr_elem.find(f'{NS}vertAlign')
    if va is not None:
        props['vertAlign'] = _a(va, 'val')
    # Small caps, caps, strike, etc.
    for tag, key in [('smallCaps', 'smallCaps'), ('caps', 'caps'), ('strike', 'strike'),
                      ('dstrike', 'dstrike'), ('shadow', 'shadow'), ('outline', 'outline')]:
        if rpr_elem.find(f'{NS}{tag}') is not None:
            props[key] = True
    return props if props else None


def extract_style_ppr(ppr_elem):
    """Extract paragraph properties from a <w:pPr> element."""
    if ppr_elem is None:
        return None
    props = {}
    # Spacing
    sp = ppr_elem.find(f'{NS}spacing')
    if sp is not None:
        s = {}
        for k in ('before', 'after', 'line', 'lineRule'):
            v = _a(sp, k)
            if v:
                s[k] = _int(v) if k in ('before', 'after') else v
        if s:
            props['spacing'] = s
    # Alignment
    jc = ppr_elem.find(f'{NS}jc')
    if jc is not None:
        props['alignment'] = _a(jc, 'val')
    # Indent
    ind = ppr_elem.find(f'{NS}ind')
    if ind is not None:
        i_props = {}
        for k in ('left', 'right', 'firstLine', 'hanging', 'firstLineChars'):
            v = _a(ind, k)
            if v:
                i_props[k] = _int(v)
        if i_props:
            props['indent'] = i_props
    # Outline level
    ol = ppr_elem.find(f'{NS}outlineLvl')
    if ol is not None:
        props['outlineLvl'] = _int(_a(ol, 'val'))
    # Keep
    for tag in ('keepNext', 'keepLines', 'pageBreakBefore', 'widowControl',
                'contextualSpacing', 'autoSpaceDE', 'autoSpaceDN'):
        if ppr_elem.find(f'{NS}{tag}') is not None:
            props[tag] = True
    # Numbering
    numPr = ppr_elem.find(f'{NS}numPr')
    if numPr is not None:
        np = {}
        ilvl = numPr.find(f'{NS}ilvl')
        numId = numPr.find(f'{NS}numId')
        if ilvl is not None:
            np['level'] = _int(_a(ilvl, 'val'))
        if numId is not None:
            np['numId'] = _a(numId, 'val')
        if np:
            props['numbering'] = np
    # Paragraph-level rPr (default run formatting for runs inside this paragraph)
    p_rPr = ppr_elem.find(f'{NS}rPr')
    p_rpr_props = extract_style_rpr(p_rPr)
    if p_rpr_props:
        props['defaultRunFmt'] = p_rpr_props
    return props if props else None


def extract_style_tblpr(tblpr_elem):
    """Extract table properties from <w:tblPr>."""
    if tblpr_elem is None:
        return None
    props = {}
    # Borders
    tblBorders = tblpr_elem.find(f'{NS}tblBorders')
    if tblBorders is not None:
        borders = {}
        for b in tblBorders:
            bname = b.tag.replace(NS, 'w:')
            bd = {}
            for k in ('val', 'sz', 'space', 'color'):
                v = _a(b, k)
                if v:
                    bd[k] = _int(v) if k in ('sz', 'space') else v
            if bd:
                borders[bname] = bd
        if borders:
            props['borders'] = borders
    # Cell margins
    tblCellMar = tblpr_elem.find(f'{NS}tblCellMar')
    if tblCellMar is not None:
        mar = {}
        for m in tblCellMar:
            mname = m.tag.replace(NS, 'w:')
            mar[mname] = _int(_a(m, 'w')) or 0
        if mar:
            props['cellMargins'] = mar
    # Width
    tblW = tblpr_elem.find(f'{NS}tblW')
    if tblW is not None:
        props['width_dxa'] = _int(_a(tblW, 'w'))
        props['width_type'] = _a(tblW, 'type')
    # Indent
    tblInd = tblpr_elem.find(f'{NS}tblInd')
    if tblInd is not None:
        props['indent_dxa'] = _int(_a(tblInd, 'w')) or 0
    # Layout
    tblLayout = tblpr_elem.find(f'{NS}tblLayout')
    if tblLayout is not None:
        props['layout'] = _a(tblLayout, 'type')
    # Alignment
    tblJc = tblpr_elem.find(f'{NS}jc')
    if tblJc is not None:
        props['alignment'] = _a(tblJc, 'val')
    # Style
    tblStyle = tblpr_elem.find(f'{NS}tblStyle')
    if tblStyle is not None:
        props['style'] = _a(tblStyle, 'val')
    # Look
    tblLook = tblpr_elem.find(f'{NS}tblLook')
    if tblLook is not None:
        look = {}
        for k in ('firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand'):
            v = _a(tblLook, k)
            if v:
                look[k] = _int(v)
        if look:
            props['look'] = look
    return props if props else None


def extract_style_tcpr(tcpr_elem):
    """Extract cell properties from <w:tcPr>."""
    if tcpr_elem is None:
        return None
    props = {}
    tcW = tcpr_elem.find(f'{NS}tcW')
    if tcW is not None:
        props['width_dxa'] = _int(_a(tcW, 'w'))
    shd = tcpr_elem.find(f'{NS}shd')
    if shd is not None:
        props['fill'] = _a(shd, 'fill')
        props['fill_color'] = _a(shd, 'color')
        props['fill_val'] = _a(shd, 'val')
    vAlign = tcpr_elem.find(f'{NS}vAlign')
    if vAlign is not None:
        props['vAlign'] = _a(vAlign, 'val')
    tcMar = tcpr_elem.find(f'{NS}tcMar')
    if tcMar is not None:
        mar = {}
        for m in tcMar:
            mname = m.tag.replace(NS, 'w:')
            mar[mname] = _int(_a(m, 'w')) or 0
        if mar:
            props['margins'] = mar
    tcBorders = tcpr_elem.find(f'{NS}tcBorders')
    if tcBorders is not None:
        borders = {}
        for b in tcBorders:
            bname = b.tag.replace(NS, 'w:')
            bd = {}
            for k in ('val', 'sz', 'space', 'color'):
                v = _a(b, k)
                if v:
                    bd[k] = _int(v) if k in ('sz', 'space') else v
            if bd:
                borders[bname] = bd
        if borders:
            props['borders'] = borders
    return props if props else None


# ─── Block extraction ─────────────────────────────────────

def _extract_header_footer_content(docx_path, target):
    """从 docx 文件中提取页眉或页脚的内容"""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            header_footer_path = f'word/{target}'
            content_xml = zf.read(header_footer_path)
            elem = etree.fromstring(content_xml)

            blocks = []
            for idx, child in enumerate(elem):
                tag = child.tag.replace(NS, 'w:')
                if tag == 'w:p':
                    block = extract_paragraph(child)
                    block['index'] = idx
                    blocks.append(block)
                    has_fields = any(r.get('field') for r in block.get('runs', []))
                    has_styles = any(r.get('fmt') for r in block.get('runs', []))
                    print(f"[extract_hf] 提取段落 #{idx}, 字段数={len(block.get('runs', []))}, 包含域代码={has_fields}, 包含样式={has_styles}")
                elif tag == 'w:tbl':
                    block = extract_table(child)
                    block['index'] = idx
                    blocks.append(block)

            print(f"[extract_hf] {target} 共提取 {len(blocks)} 个块")
            return blocks if blocks else None
    except (KeyError, Exception) as e:
        print(f"[extract_hf] 提取 {target} 失败: {e}")
        return None


def extract_run(r_elem):
    """Extract a single <w:r> into {text, format, field}."""
    text = ''.join(r_elem.xpath('.//w:t/text()',
        namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
    rPr = r_elem.find(f'{NS}rPr')
    fmt = extract_style_rpr(rPr)
    
    field = None
    fldChar = r_elem.find(f'{NS}fldChar')
    if fldChar is not None:
        fldCharType = _a(fldChar, 'fldCharType')
        if fldCharType:
            field = {'type': fldCharType}
    
    instrText = r_elem.find(f'{NS}instrText')
    if instrText is not None:
        instr_val = instrText.text or ''
        field = {'type': 'instruction', 'value': instr_val.strip()}
    
    if field or fmt:
        print(f"[extract_run] 文本='{text[:20]}', 样式={fmt}, 域代码={field}")
    
    result = {'text': text}
    if fmt:
        result['fmt'] = fmt
    if field:
        result['field'] = field
    return result


def extract_paragraph(p_elem):
    """Extract a <w:p> block."""
    pPr = p_elem.find(f'{NS}pPr')
    para_fmt = extract_style_ppr(pPr)
    style_id = None
    if pPr is not None:
        ps = pPr.find(f'{NS}pStyle')
        if ps is not None:
            style_id = _a(ps, 'val')

    runs = []
    for r in p_elem.findall(f'{NS}r'):
        has_text = bool(''.join(r.xpath('.//w:t/text()',
            namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})))
        has_fldChar = r.find(f'{NS}fldChar') is not None
        has_instrText = r.find(f'{NS}instrText') is not None
        if has_text or has_fldChar or has_instrText:
            runs.append(extract_run(r))

    # Check for image
    drawing = p_elem.find(f'{NS}r/{NS}drawing')
    has_image = drawing is not None
    image_info = None
    if has_image:
        inline = drawing.find(f'{NS_DRAW}inline')
        if inline is not None:
            extent = inline.find(f'{NS_DRAW}extent')
            blip = drawing.find(f'.//{NS_A}blip')
            image_info = {
                'width_emu': _int(extent.get('cx')) if extent is not None else None,
                'height_emu': _int(extent.get('cy')) if extent is not None else None,
                'rId': _a(blip, 'embed', NS_R) if blip is not None else None,
            }

    result = {
        'type': 'paragraph',
        'style': style_id,
        'runs': runs,
    }
    if para_fmt:
        result['paraFmt'] = para_fmt
    if has_image:
        result['image'] = image_info
    return result


def extract_table(tbl_elem):
    """Extract a <w:tbl> block."""
    tblPr = tbl_elem.find(f'{NS}tblPr')
    tbl_fmt = extract_style_tblpr(tblPr)

    # Grid
    tblGrid = tbl_elem.find(f'{NS}tblGrid')
    columns = []
    if tblGrid is not None:
        for gc in tblGrid.findall(f'{NS}gridCol'):
            w = _a(gc, 'w')
            columns.append({'width_dxa': _int(w) if w else None})

    rows = []
    for tr in tbl_elem.findall(f'{NS}tr'):
        cells = []
        for tc in tr.findall(f'{NS}tc'):
            tcPr = tc.find(f'{NS}tcPr')
            cell_fmt = extract_style_tcpr(tcPr)

            # Paragraphs inside cell
            paras = []
            for p in tc.findall(f'{NS}p'):
                paras.append(extract_paragraph(p))

            cell = {'paragraphs': paras}
            if cell_fmt:
                cell['cellFmt'] = cell_fmt
            cells.append(cell)
        rows.append({'cells': cells})

    result = {
        'type': 'table',
        'columns': columns,
        'numRows': len(rows),
        'numCols': len(columns),
        'rows': rows,
    }
    if tbl_fmt:
        result['tableFmt'] = tbl_fmt
    return result


# ─── Main conversion ──────────────────────────────────────

def docx_to_json(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as zf:
        doc_xml = zf.read('word/document.xml')
        styles_xml = zf.read('word/styles.xml')

        # 读取关系文件，用于解析页眉页脚引用
        try:
            rels_xml = zf.read('word/_rels/document.xml.rels')
        except KeyError:
            rels_xml = None

    doc = etree.fromstring(doc_xml)
    styles_doc = etree.fromstring(styles_xml)
    body = doc.find(f'{NS}body')

    # ── 解析关系映射 ──
    rels_map = {}
    if rels_xml:
        rels_doc = etree.fromstring(rels_xml)
        for rel in rels_doc.findall(f'{NS_RELS}Relationship'):
            rid = rel.get('Id')
            target = rel.get('Target')
            if rid and target:
                rels_map[rid] = target

    # ── Page setup ──
    sectPr = body.find(f'{NS}sectPr')
    page = {}
    headers = {}
    footers = {}

    if sectPr is not None:
        pgSz = sectPr.find(f'{NS}pgSz')
        if pgSz is not None:
            page['width_twips'] = _int(pgSz.get(f'{NS}w'))   # OOXML 单位是 twips
            page['height_twips'] = _int(pgSz.get(f'{NS}h'))
        pgMar = sectPr.find(f'{NS}pgMar')
        if pgMar is not None:
            page['margins'] = {
                'top': _int(pgMar.get(f'{NS}top')),
                'right': _int(pgMar.get(f'{NS}right')),
                'bottom': _int(pgMar.get(f'{NS}bottom')),
                'left': _int(pgMar.get(f'{NS}left')),
                'header': _int(pgMar.get(f'{NS}header')),
                'footer': _int(pgMar.get(f'{NS}footer')),
            }

        # ── 提取页眉页脚引用 ──
        header_refs = sectPr.findall(f'{NS}headerReference')
        footer_refs = sectPr.findall(f'{NS}footerReference')

        # 解析页眉
        for href in header_refs:
            htype = _a(href, 'type')  # default/first/even
            rid = _a(href, 'id', NS_R)
            if htype and rid and rid in rels_map:
                target = rels_map[rid]
                # 提取页眉内容
                header_content = _extract_header_footer_content(docx_path, target)
                if header_content:
                    headers[htype] = header_content

        # 解析页脚
        for fref in footer_refs:
            ftype = _a(fref, 'type')  # default/first/even
            rid = _a(fref, 'id', NS_R)
            if ftype and rid and rid in rels_map:
                target = rels_map[rid]
                # 提取页脚内容
                footer_content = _extract_header_footer_content(docx_path, target)
                if footer_content:
                    footers[ftype] = footer_content

    # ── Document defaults ──
    defaults = {}
    dd = styles_doc.find(f'{NS}docDefaults')
    if dd is not None:
        rPrDefault = dd.find(f'{NS}rPrDefault/{NS}rPr')
        if rPrDefault is not None:
            defaults['run'] = extract_style_rpr(rPrDefault)
        pPrDefault = dd.find(f'{NS}pPrDefault/{NS}pPr')
        if pPrDefault is not None:
            defaults['paragraph'] = extract_style_ppr(pPrDefault)

    # ── Styles ──
    styles = {}
    for st in styles_doc.findall(f'{NS}style'):
        sid = st.get(f'{NS}styleId')
        sname = st.find(f'{NS}name')
        name = sname.get(f'{NS}val') if sname is not None else sid
        stype = st.get(f'{NS}type')  # paragraph / character / table / numbering
        basedOn = st.find(f'{NS}basedOn')
        based = basedOn.get(f'{NS}val') if basedOn is not None else None
        nextStyle = st.find(f'{NS}next')
        next_s = nextStyle.get(f'{NS}val') if nextStyle is not None else None

        entry = {
            'name': name,
            'type': stype,
            'id': sid,
        }
        if based:
            entry['basedOn'] = based
        if next_s:
            entry['nextStyle'] = next_s

        rPr = st.find(f'{NS}rPr')
        rpr_data = extract_style_rpr(rPr)
        if rpr_data:
            entry['run'] = rpr_data

        pPr = st.find(f'{NS}pPr')
        ppr_data = extract_style_ppr(pPr)
        if ppr_data:
            entry['paragraph'] = ppr_data

        tblPr = st.find(f'{NS}tblPr')
        tblpr_data = extract_style_tblpr(tblPr)
        if tblpr_data:
            entry['table'] = tblpr_data

        tcPr = st.find(f'{NS}tcPr')
        tcpr_data = extract_style_tcpr(tcPr)
        if tcpr_data:
            entry['cell'] = tcpr_data

        styles[sid] = entry

    # ── Blocks (body children in order) ──
    blocks = []
    for idx, child in enumerate(body):
        tag = child.tag.replace(NS, 'w:')
        if tag == 'w:p':
            block = extract_paragraph(child)
            block['index'] = idx
            blocks.append(block)
        elif tag == 'w:tbl':
            block = extract_table(child)
            block['index'] = idx
            blocks.append(block)
        elif tag == 'w:sectPr':
            break
        # Skip other elements (bookmarks, etc.)

    return {
        'page': page,
        'defaults': defaults,
        'styles': styles,
        'blocks': blocks,
        'headers': headers,
        'footers': footers,
    }


# ─── CLI ──────────────────────────────────────────────────

if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser(description='Convert .docx to structured JSON')
    p.add_argument('docx', help='Path to .docx file')
    p.add_argument('--out', default=None, help='Output directory (default: same as docx)')
    args = p.parse_args()

    src = args.docx
    out_dir = args.out or os.path.dirname(os.path.abspath(src))
    base = os.path.splitext(os.path.basename(src))[0]

    full = docx_to_json(src)

    # Full JSON
    full_path = os.path.join(out_dir, f'{base}_full.json')
    with open(full_path, 'w', encoding='utf-8') as f:
        json.dump(full, f, ensure_ascii=False, indent=2)
    print(f'[full] {full_path}  ({len(full["blocks"])} blocks, {len(full["styles"])} styles)')


