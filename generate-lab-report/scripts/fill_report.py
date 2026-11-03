#!/usr/bin/env python3
"""Fill five labeled regions without modifying fixed template information."""
import argparse
import json
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell

LABELS = {'purpose':'实验目的：','principle':'实验原理：','process':'实验内容及过程：','results':'实验结果及分析：','summary':'实验小结：'}

def locate(doc):
    candidates=[]
    for table in doc.tables:
        found={}
        for row in table.rows:
            for key,label in LABELS.items():
                if row.cells[0].text.lstrip().startswith(label):
                    if key in found:
                        raise ValueError('模板包含多套实验区域；请先显式选择并提取单个实验模板')
                    found[key]=row.cells[0]
        if len(found)==5: candidates.append((table,found))
    if len(candidates)!=1:
        raise ValueError('需要且只能有一个包含五个目标标题的表格')
    return candidates[0]

def fill(doc,cell,blocks,heading=True):
    paragraphs=cell.paragraphs
    for node in list(cell._tc):
        if node.tag==qn('w:tcPr'): continue
        if heading and node is paragraphs[0]._p: continue
        cell._tc.remove(node)
    for block in blocks:
        if isinstance(block,str): block={'type':'text','text':block}
        kind=block.get('type','text')
        if kind not in ('text','heading','label','code','command','screenshot'):
            raise ValueError('未知内容类型')
        p=cell.add_paragraph()
        grid=OxmlElement('w:snapToGrid'); grid.set(qn('w:val'),'0'); p._p.get_or_add_pPr().append(grid)
        fmt=p.paragraph_format
        fmt.space_before=Pt(0); fmt.space_after=Pt(0)
        fmt.line_spacing=1.0 if kind in ('code','command') else 1.5
        fmt.first_line_indent=Pt(24) if kind=='text' else Pt(0)
        fmt.left_indent=Pt(24) if kind=='command' else Pt(0)
        fmt.keep_with_next=kind in ('heading','label')
        fmt.widow_control=True
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER if kind=='screenshot' else (WD_ALIGN_PARAGRAPH.JUSTIFY if kind=='text' else WD_ALIGN_PARAGRAPH.LEFT)
        def add(text,bold=False):
            r=p.add_run(text)
            r.font.name='Consolas' if kind in ('code','command') else 'Times New Roman'
            r.font.size=Pt(10 if kind=='code' else 12)
            r.font.bold=bold
            r.font.color.rgb=RGBColor(0,0,0)
            r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
            return r
        if kind=='screenshot':
            if not block.get('name') or not block.get('comment'): raise ValueError('截图需要名称及批注步骤')
            run=add('【截图：'+block['name']+'】')
            doc.add_comment(run,text=block['comment'],author='截图操作指引',initials='图')
        elif kind in ('heading','label'):
            add(block['text'],True)
        else:
            text=block['text']
            prefixes=('任务要求：','操作步骤：','参数分析：','实验结果：','任务分析：','代码分析：','分析：')
            prefix=next((v for v in prefixes if text.startswith(v)),None) if kind=='text' else None
            if prefix:
                add(prefix,True); add(text[len(prefix):])
            else: add(text)
    if not cell.paragraphs: cell.add_paragraph()

def generate(data,out,template):
    for key in ('purpose','principle','results','summary','process_pages'):
        if not isinstance(data.get(key),list) or not data[key]: raise ValueError('缺少非空数组：'+key)
    if any(not isinstance(p,list) or not p for p in data['process_pages']): raise ValueError('每个过程分页必须非空')
    doc=Document(template)
    if not hasattr(doc,'add_comment'): raise RuntimeError('需要python-docx>=1.2以写入真实Word批注')
    table,cells=locate(doc)
    start=cells['process']._tc.getparent(); end=cells['results']._tc.getparent()
    middle=[]; node=start.getnext()
    while node is not None and node is not end:
        if ''.join(node.itertext()).strip():
            # itertext includes XML text only; allow only empty process continuation cells.
            texts=node.xpath('.//w:t')
            if any(t.text and t.text.strip() for t in texts): raise ValueError('过程续页含有内容，请先确认模板')
        middle.append(node); node=node.getnext()
    if node is None: raise ValueError('过程与结果区域顺序异常')
    # Keep exactly one process cell; Word/WPS flows the row across pages.
    for row in middle: table._tbl.remove(row)
    pr=start.find(qn('w:trPr'))
    if pr is not None:
        for node in list(pr):
            if node.tag==qn('w:cantSplit'): pr.remove(node)
            elif node.tag==qn('w:trHeight') and node.get(qn('w:hRule'))=='exact':
                node.set(qn('w:hRule'),'atLeast')
    blocks=[block for group in data['process_pages'] for block in group]
    fill(doc,cells['process'],blocks)
    for key in ('purpose','principle','results','summary'): fill(doc,cells[key],data[key])
    Path(out).parent.mkdir(parents=True,exist_ok=True)
    doc.save(out)
    return out

if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('content');ap.add_argument('output')
    ap.add_argument('--template',default=str(Path(__file__).resolve().parents[1]/'assets/template.docx'))
    args=ap.parse_args()
    print(generate(json.loads(Path(args.content).read_text(encoding='utf-8')),args.output,args.template))
