"""Word document operations replacing the legacy selection/clipboard routines."""
from copy import deepcopy
from pathlib import Path
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

from .engine import paragraphs, chart_image, bullet_numbering

TAG = re.compile(r'@#\$_[A-Za-z0-9_]+', re.I)


class DocumentError(ValueError):
    pass


def text_nodes(paragraph):
    # Descendant text includes hyperlinks/content controls but not field instructions.
    return list(paragraph._p.iter(qn('w:t')))


def text_of(paragraph):
    return ''.join(n.text or '' for n in text_nodes(paragraph))


def replace_literal(paragraph, before, after, tag=False):
    if not before:
        raise DocumentError('An empty search string is not allowed')
    nodes = text_nodes(paragraph)
    text = ''.join(n.text or '' for n in nodes)
    pattern = re.escape(before) + (r'(?![A-Za-z0-9_])' if tag else '')
    matches = list(re.finditer(pattern, text, re.I))
    for match in reversed(matches):
        offset = 0
        for node in nodes:
            value = node.text or ''
            end = offset + len(value)
            if offset < match.end() and end > match.start():
                left, right = max(0, match.start()-offset), min(len(value), match.end()-offset)
                insert = str(after) if offset <= match.start() < end else ''
                node.text = value[:left] + insert + value[right:]
                node.set(qn('xml:space'), 'preserve')
            offset = end
    return len(matches)


def set_property(parent, name, attributes):
    node = parent.find(qn(name))
    if node is None:
        node = OxmlElement(name)
        parent.append(node)
    for key, value in attributes.items():
        node.set(qn(key), str(value))
    return node


def apply_cell_style(cell, style):
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        alignment = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'right': WD_ALIGN_PARAGRAPH.RIGHT,
                     'center': WD_ALIGN_PARAGRAPH.CENTER, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}
        if style.get('horizontal') in alignment:
            paragraph.alignment = alignment[style['horizontal']]
        for run in paragraph.runs:
            if style.get('font'):
                run.font.name = style['font']
            if style.get('font_size'):
                run.font.size = Pt(style['font_size'])
            run.bold, run.italic = style.get('bold', False), style.get('italic', False)
            if style.get('font_color'):
                run.font.color.rgb = RGBColor.from_string(style['font_color'])
    if style.get('vertical') in ('top', 'center', 'bottom'):
        cell.vertical_alignment = {'top': WD_CELL_VERTICAL_ALIGNMENT.TOP,
                                   'center': WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                                   'bottom': WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}[style['vertical']]
    props = cell._tc.get_or_add_tcPr()
    if style.get('fill'):
        set_property(props, 'w:shd', {'w:fill': style['fill'], 'w:val': 'clear'})
    if 'borders' in style:
        borders = set_property(props, 'w:tcBorders', {})
        mapping = {'thin': ('single', 4), 'medium': ('single', 8), 'thick': ('single', 12),
                   'hair': ('single', 2), 'dotted': ('dotted', 4), 'dashed': ('dashed', 4),
                   'double': ('double', 4), 'dashDot': ('dotDash', 4), 'dashDotDot': ('dotDotDash', 4),
                   'mediumDashed': ('dashed', 8), 'mediumDashDot': ('dotDash', 8),
                   'mediumDashDotDot': ('dotDotDash', 8), 'slantDashDot': ('dotDash', 8)}
        for side, spec in style['borders'].items():
            if spec.get('style') and spec['style'] not in mapping:
                raise DocumentError(f'Unsupported Excel border style: {spec["style"]}')
            line, width = mapping.get(spec.get('style'), ('nil', 0))
            set_property(borders, 'w:'+side, {'w:val':line,'w:sz':width,'w:color':spec.get('color') or 'auto'})


def table_properties(table, element):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER if element.get('align') == 'center' else WD_TABLE_ALIGNMENT.LEFT
    if element.get('fit'):
        set_property(table._tbl.tblPr,'w:tblW',{'w:type':'pct','w:w':5000})
        table.autofit=True
    elif element.get('widths_inches'):
        table.autofit=False
        for column,width in zip(table.columns,element['widths_inches']):
            column.width=Inches(width)
        for row in table.rows:
            for cell,width in zip(row.cells,element['widths_inches']):
                cell.width=Inches(width)
    for i,row in enumerate(table.rows):
        props=row._tr.get_or_add_trPr()
        old=props.find(qn('w:tblHeader'))
        if old is not None:
            props.remove(old)
        if i<element.get('header_rows',0):
            set_property(props,'w:tblHeader',{'w:val':'true'})
        if i<len(element.get('row_heights_points',[])) and element['row_heights_points'][i]:
            row.height=Pt(element['row_heights_points'][i])
    set_property(table._tbl.tblPr,'w:tblCaption',{'w:val':element.get('title','')})
    set_property(table._tbl.tblPr,'w:tblDescription',{'w:val':element.get('alt_text','')})
    if element.get('padding_px') is not None and element.get('padding_sides'):
        margins=set_property(table._tbl.tblPr,'w:tblCellMar',{})
        for letter,side in [('l','left'),('r','right'),('t','top'),('b','bottom')]:
            if letter in element['padding_sides'].lower():
                set_property(margins,'w:'+side,{'w:w':round(float(element['padding_px'])*15),'w:type':'dxa'})


def containing_table(paragraph):
    node=paragraph._p.getparent()
    while node is not None and node.tag != qn('w:tbl'):
        node=node.getparent()
    if node is None:
        raise DocumentError('Existing-table tag must be inside a Word table')
    return Table(node,paragraph._parent)


def available_width(document, paragraph):
    """Return the containing cell/page width in inches for a chart anchor."""
    node=paragraph._p
    cell_width=None
    while node.getparent() is not document.element.body:
        if node.tag==qn('w:tc'):
            width=node.find('./'+qn('w:tcPr')+'/'+qn('w:tcW'))
            if width is not None and width.get(qn('w:type'))=='dxa':
                cell_width=int(width.get(qn('w:w')))/1440
        node=node.getparent()
        if node is None:
            break
    blocks=list(document.element.body)
    position=blocks.index(node) if node in blocks else 0
    section=document.sections[-1]
    for candidate in document.sections:
        end=candidate._sectPr.getparent()
        if end.tag==qn('w:pPr'):
            end=end.getparent()
            if end in blocks and blocks.index(end)>=position:
                section=candidate
                break
    page=(section.page_width-section.left_margin-section.right_margin-(section.gutter or 0))/914400
    return min(page,cell_width) if cell_width else page


def fill_shell(paragraph, tag, element):
    table=containing_table(paragraph)
    start=None
    for r,row in enumerate(table.rows):
        for c,cell in enumerate(row.cells):
            if cell._tc is paragraph._p.getparent():
                start=(r,c)
                break
        if start:
            break
    if start is None:
        raise DocumentError('Cannot locate shell-table tag cell')
    r1,c1=start
    values=element['values']
    if not values:
        replace_literal(paragraph,tag,'',tag=True)
        return
    if any(c1+len(row)>len(table.columns) for row in values):
        raise DocumentError(f'Table shell is too narrow for {tag}')
    required=r1+len(values)
    if required>len(table.rows):
        if not element.get('variable_rows'):
            raise DocumentError(f'Fixed table shell has {len(table.rows)-r1} data rows but {tag} needs {len(values)}')
        prototype=table.rows[r1]._tr
        while len(table.rows)<required:
            table._tbl.append(deepcopy(prototype))
    for i,row in enumerate(values):
        for j,value in enumerate(row):
            cell=table.cell(r1+i,c1+j)
            # Retain each shell cell's paragraph/run style while replacing its data.
            p=cell.paragraphs[0]
            first=deepcopy(p.runs[0]._r.rPr) if p.runs and p.runs[0]._r.rPr is not None else None
            p.clear()
            run=p.add_run(str(value))
            if first is not None:
                run._r.insert(0,first)
            for extra in list(cell.paragraphs[1:]):
                cell._tc.remove(extra._p)
    # Shell formatting is authoritative; apply only metadata, not Excel cell styles.
    for key,xml in [('title','w:tblCaption'),('alt_text','w:tblDescription')]:
        if key in element:
            set_property(table._tbl.tblPr,xml,{'w:val':element[key]})


def insert_legacy_table(document, paragraph, tag, element):
    if element.get('mode')=='shell':
        fill_shell(paragraph,tag,element)
        return
    values=element['values']
    if not values or not values[0] or any(len(row)!=len(values[0]) for row in values):
        raise DocumentError(f'Invalid table dimensions for {tag}')
    table=document.add_table(rows=len(values),cols=len(values[0]))
    for i,row in enumerate(values):
        for j,value in enumerate(row):
            cell=table.cell(i,j)
            style=element['cell_styles'][i][j] if element.get('cell_styles') else {}
            # Wingdings lowercase l is the legacy bullet glyph, not report data.
            glyph=str(value)=='l' and (style.get('font') or '').casefold()=='wingdings'
            cell.text='\u2022' if glyph else str(value)
            if element.get('cell_styles'):
                apply_cell_style(cell,dict(style,font='Arial') if glyph else style)
    table_properties(table,element)
    for r1,c1,r2,c2 in element.get('merges',[]):
        merged=table.cell(r1,c1).merge(table.cell(r2,c2))
        merged.text=str(values[r1][c1])
        if element.get('cell_styles'):
            apply_cell_style(merged,element['cell_styles'][r1][c1])
    # A block table may replace an inline tag; retain prefix/suffix paragraphs.
    split_anchor(paragraph,tag)
    paragraph._p.addnext(table._tbl)


def split_anchor(paragraph,tag):
    """Split a paragraph at a block placeholder, keeping surrounding content."""
    match=re.search(re.escape(tag)+r'(?![A-Za-z0-9_])',text_of(paragraph),re.I)
    if not match:
        raise DocumentError(f'Cannot find block tag {tag}')
    text=text_of(paragraph)
    if text[match.end():]:
        suffix=deepcopy(paragraph._p)
        paragraph._p.addnext(suffix)
        other=Paragraph(suffix,paragraph._parent)
        # Keep text-node styles; trim matching character interval directly.
        trim_interval(other,0,match.end())
    trim_interval(paragraph,match.start(),len(text))


def trim_interval(paragraph,start,end):
    offset=0
    for node in text_nodes(paragraph):
        value=node.text or '';last=offset+len(value)
        if offset<end and last>start:
            node.text=value[:max(0,start-offset)]+value[min(len(value),end-offset):]
        offset=last


def cut_sections(document,sections):
    events=[]
    for spec in sections:
        body_paragraphs=[p for p in paragraphs(document) if p._p.getparent() is not None and p._p in document.element.body.iter(qn('w:p'))]
        start=next((p for p in body_paragraphs if spec['start'].casefold() in text_of(p).casefold()),None)
        end=next((p for p in body_paragraphs if spec['end'].casefold() in text_of(p).casefold()),None)
        if start is None and end is None:
            events.append({'section':spec['start'],'action':'absent'})
            continue
        if start is None or end is None:
            raise DocumentError(f'Unpaired SmartSections markers: {spec["start"]} / {spec["end"]}')
        if spec['include']:
            replace_literal(start,spec['start'],'')
            replace_literal(end,spec['end'],'')
            events.append({'section':spec['start'],'action':'kept'})
            continue
        if start._p is end._p:
            text=text_of(start);a=text.casefold().index(spec['start'].casefold());b=text.casefold().index(spec['end'].casefold())+len(spec['end'])
            trim_interval(start,a,b)
            start.add_run(spec['reason'])
        else:
            body=document.element.body
            if start._p.getparent() is not body or end._p.getparent() is not body:
                raise DocumentError('SmartSections spanning paragraphs must use body-level marker paragraphs, not markers inside table cells')
            blocks=list(body);a,b=blocks.index(start._p),blocks.index(end._p)
            if b<a:
                raise DocumentError('SmartSections end precedes start')
            prefix=text_of(start).casefold().index(spec['start'].casefold())
            endpos=text_of(end).casefold().index(spec['end'].casefold())+len(spec['end'])
            trim_interval(start,prefix,len(text_of(start)))
            trim_interval(end,0,endpos)
            for node in blocks[a+1:b]:
                body.remove(node)
            start.add_run(spec['reason'])
            if not text_of(end) and end._p.pPr is None:
                body.remove(end._p)
        if spec.get('bullet_mode')==2:
            start.alignment=WD_ALIGN_PARAGRAPH.LEFT
        if spec.get('bullet_mode')==1 and not spec['reason'] and not text_of(start):
            # Original backspace behavior joins the cut with the previous paragraph.
            if start._p.pPr is None or start._p.pPr.sectPr is None:
                start._p.getparent().remove(start._p)
        events.append({'section':spec['start'],'action':'removed','reason':spec['reason']})
    return events


def insert_documents(document,specs):
    events=[]
    from docxcompose.composer import Composer
    composer=Composer(document)
    for spec in specs:
        matches=[p for p in paragraphs(document) if re.search(re.escape(spec['tag'])+r'(?![A-Za-z0-9_])',text_of(p),re.I)]
        if not matches:
            continue
        path=Path(spec['path'])
        for p in matches:
            if not path.is_file():
                replace_literal(p,spec['tag'],f'Section missing: File {spec["filename"]} does not exist.',tag=True)
                events.append({'tag':spec['tag'],'action':'missing','file':str(path)})
                continue
            if p._p.getparent() is not document.element.body:
                raise DocumentError('Additional document insertion requires a body paragraph tag')
            split_anchor(p,spec['tag'])
            index=list(document.element.body).index(p._p)+1
            composer.insert(index,Document(path),remove_property_fields=False)
            events.append({'tag':spec['tag'],'action':'inserted','file':str(path)})
    return events


def change_headers(document,specs):
    align={'Left':WD_ALIGN_PARAGRAPH.LEFT,'Center':WD_ALIGN_PARAGRAPH.CENTER,'Right':WD_ALIGN_PARAGRAPH.RIGHT}
    for spec in specs:
        for kind in ('header','footer'):
            value=spec.get(kind)
            if value in (None,''):
                continue
            if spec['section']>len(document.sections):
                raise DocumentError(f'Header/footer configuration references missing section {spec["section"]}')
            story=getattr(document.sections[spec['section']-1],kind)
            story.is_linked_to_previous=False
            p=story.paragraphs[0]
            p.text=str(value)
            for node in list(story._element):
                if node is not p._p:
                    story._element.remove(node)
            p.alignment=align.get(spec.get(kind+'_align'),WD_ALIGN_PARAGRAPH.LEFT)


def populate(document,elements):
    folded={key.casefold():value for key,value in elements.items()}
    if len(folded)!=len(elements):
        raise DocumentError('Duplicate case-insensitive element tags')
    populated=set()
    for p in list(paragraphs(document)):
        # A table-shell fill can remove several original paragraph nodes.
        if p._p.getparent() is None:
            continue
        for tag in list(dict.fromkeys(TAG.findall(text_of(p)))):
            if tag.casefold() not in folded:
                continue  # Another pipeline stage may own this tag.
            element=folded[tag.casefold()]
            kind=element['type']
            if kind=='text':
                replace_literal(p,tag,str(element['value']),tag=True)
            elif kind=='legacy_table':
                insert_legacy_table(document,p,tag,element)
            elif kind=='chart':
                # An image can live inline; split runs after the tag without disturbing prefix.
                split_anchor(p,tag)
                original_width=element.get('width_inches',5.5)
                width=min(original_width,available_width(document,p))
                height=element.get('height_inches')
                shape=p.add_run().add_picture(chart_image(element),width=Inches(width),
                                              height=Inches(height*width/original_width) if height else None)
                shape._inline.docPr.set('descr',element.get('alt_text',''))
            else:
                raise DocumentError(f'Unsupported workbook element type: {kind}')
            populated.add(tag.casefold())
    return sorted(populated)


def generate(template,destination,elements,mass_elements=None,sections=None,postprocess=None):
    document=Document(template)
    events=cut_sections(document,sections or [])
    used=populate(document,mass_elements or {})
    used+=populate(document,elements)
    settings=postprocess or {}
    events+=insert_documents(document,settings.get('documents',[]))
    for before,after in settings.get('replacements',[]):
        for paragraph in paragraphs(document):
            replace_literal(paragraph,before,after)
    change_headers(document,settings.get('headers',[]))
    remaining={tag for p in paragraphs(document) for tag in TAG.findall(text_of(p))}
    if remaining:
        raise DocumentError('Unresolved template tags: '+', '.join(sorted(remaining)))
    set_property(document.settings.element,'w:updateFields',{'w:val':'false'})
    from .files import save_document
    save_document(document,destination)
    return {'output':str(destination),'tags_populated':sorted(set(used)),'document_events':events}
