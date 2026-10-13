"""Rebuild heading-based TOC cached results from the renderer's actual PDF layout.

The supported contract is the original templates' TOC \\o plus \\h/\\z/\\u.
Other field switches and mismatched PDF headings fail instead of guessing pages.
"""
from copy import deepcopy
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader

from .document import text_of,paragraphs,set_property


class TocError(ValueError):
    pass


def pdf_outline(path):
    reader=PdfReader(path)
    result=[]
    def visit(items,level=1):
        for item in items:
            if isinstance(item,list):
                visit(item,level+1)
            else:
                page=reader.get_destination_page_number(item)
                result.append({'level':level,'title':str(item.title),'page':page+1,
                               'page_label':reader.page_labels[page]})
    visit(reader.outline)
    return result,len(reader.pages)


def heading_level(paragraph):
    node=paragraph._p.pPr
    if node is not None and node.outlineLvl is not None:
        return node.outlineLvl.val+1
    style=paragraph.style
    seen=set()
    while style and style.style_id not in seen:
        seen.add(style.style_id)
        props=style.element.pPr
        if props is not None:
            level=props.find(qn('w:outlineLvl'))
            if level is not None:
                return int(level.get(qn('w:val')))+1
        match=re.fullmatch(r'Heading\s*(\d+)',style.name,re.I)
        if match:
            return int(match[1])
        style=style.base_style
    return None


def toc_field(document):
    instructions=[n for n in document.element.iter(qn('w:instrText')) if n.text and re.match(r'^\s*TOC\b',n.text,re.I)]
    if len(instructions)!=1:
        raise TocError(f'Expected one heading-based table of contents; found {len(instructions)}')
    instruction=instructions[0]
    text=instruction.text
    unknown=set(re.findall(r'\\([A-Za-z])',text))-set('ohzu')
    if unknown:
        raise TocError('Unsupported TOC switches: '+', '.join(sorted(unknown)))
    limits=re.search(r'\\o\s+"(\d+)-(\d+)"',text)
    if not limits:
        raise TocError('TOC must specify a heading range with \\o "1-3"')
    first=instruction
    while first.tag!=qn('w:p'):
        first=first.getparent()
    parent=first.getparent()
    blocks=list(parent);start=blocks.index(first);depth=0;end=None
    for i in range(start,len(blocks)):
        for field in blocks[i].iter(qn('w:fldChar')):
            kind=field.get(qn('w:fldCharType'))
            if kind=='begin':
                depth+=1
            elif kind=='end':
                depth-=1
                if depth==0:
                    end=i
                    break
        if end is not None:
            break
    if end is None:
        raise TocError('TOC field end was not found')
    return parent,start,end,tuple(map(int,limits.groups())),text


def run_node(text=None,field=None,instruction=None,prototype=None):
    run=OxmlElement('w:r')
    if prototype is not None:
        run.append(deepcopy(prototype))
    if text is not None:
        node=OxmlElement('w:t');node.text=str(text);node.set(qn('xml:space'),'preserve');run.append(node)
    if field:
        node=OxmlElement('w:fldChar');node.set(qn('w:fldCharType'),field);node.set(qn('w:dirty'),'false');run.append(node)
    if instruction:
        node=OxmlElement('w:instrText');node.text=instruction;node.set(qn('xml:space'),'preserve');run.append(node)
    return run


def rebuild_toc(docx,pdf,destination):
    document=Document(docx)
    parent,start,end,limits,instruction=toc_field(document)
    outline,pages=pdf_outline(pdf)
    entries=[e for e in outline if limits[0]<=e['level']<=limits[1]]
    # Match actual PDF outline entries to live body headings, including numbering.
    headings=[p for p in paragraphs(document) if heading_level(p) is not None]
    normalized=lambda text:re.sub(r'\s+',' ',text).strip()
    cursor=0
    matching=[]
    for entry in entries:
        title=normalized(entry['title'])
        found=None
        for i in range(cursor,len(headings)):
            p=headings[i];value=normalized(text_of(p))
            if heading_level(p)==entry['level'] and (title==value or title.endswith(' '+value)):
                found=p;cursor=i+1;break
        if found is None:
            raise TocError(f'Cannot match rendered PDF heading to template: {entry["title"]}')
        matching.append(found)
    if not entries:
        raise TocError('Renderer exported no matching heading bookmarks; cannot calculate TOC page numbers')
    old=list(parent)[start:end+1]
    templates={}
    run_props=None
    for p in old:
        props=p.find(qn('w:pPr'))
        if props is not None:
            style=props.find(qn('w:pStyle'))
            if style is not None:
                match=re.fullmatch(r'TOC(\d+)',style.get(qn('w:val'),''),re.I)
                if match:
                    templates.setdefault(int(match[1]),props)
        if run_props is None:
            run_props=next(p.iter(qn('w:rPr')),None)
    ids=[int(n.get(qn('w:id'))) for n in document.element.iter(qn('w:bookmarkStart'))]
    next_id=max(ids,default=0)+1
    # Reuse stable names on repeated layout passes; avoid adding duplicate bookmarks.
    anchors=[]
    for i,p in enumerate(matching):
        name=f'_TRWContents{i+1}'
        existing=next((n for n in p._p.iter(qn('w:bookmarkStart')) if n.get(qn('w:name'))==name),None)
        if existing is None:
            begin=OxmlElement('w:bookmarkStart');begin.set(qn('w:id'),str(next_id));begin.set(qn('w:name'),name)
            finish=OxmlElement('w:bookmarkEnd');finish.set(qn('w:id'),str(next_id));next_id+=1
            p._p.insert(1 if p._p.pPr is not None else 0,begin);p._p.append(finish)
        anchors.append(name)
    for p in old:
        parent.remove(p)
    for i,(entry,anchor) in enumerate(zip(entries,anchors)):
        p=OxmlElement('w:p')
        props=deepcopy(templates.get(entry['level'],next(iter(templates.values()),None)))
        if props is None:
            props=OxmlElement('w:pPr')
        set_property(props,'w:pStyle',{'w:val':f'TOC{entry["level"]}'})
        p.append(props)
        if i==0:
            p.append(run_node(field='begin'));p.append(run_node(instruction=instruction));p.append(run_node(field='separate'))
        link=OxmlElement('w:hyperlink');link.set(qn('w:anchor'),anchor);link.set(qn('w:history'),'1')
        link.append(run_node(entry['title'],prototype=run_props))
        tab=OxmlElement('w:r');tab.append(OxmlElement('w:tab'));link.append(tab)
        link.append(run_node(field='begin'));link.append(run_node(instruction=f' PAGEREF {anchor} \\h '));link.append(run_node(field='separate'))
        link.append(run_node(entry['page_label'],prototype=run_props));link.append(run_node(field='end'))
        p.append(link)
        if i==len(entries)-1:
            p.append(run_node(field='end'))
        parent.insert(start+i,p)
    set_property(document.settings.element,'w:updateFields',{'w:val':'false'})
    document.save(destination)
    return entries,pages
