#!/usr/bin/env python3
"""Verify fixed XML and rendered PDF heading co-location. Exit nonzero on mismatch."""
import argparse
from docx import Document
from lxml import etree
import fitz
from fill_report import locate

def canon(node):
    return etree.tostring(node,method='c14n')

def fixed(doc):
    table,cells=locate(doc)
    # Locate directly by labels because lxml element wrappers may be recreated.
    labels=('实验目的：','实验原理：','实验内容及过程：','实验结果及分析：','实验小结：')
    rows=[]
    for row in table.rows:
        text=row.cells[0].text.lstrip()
        if text.startswith(labels) or not text.strip(): continue
        if text.startswith(('第','学生学习态度：','教师批阅：')): rows.append(canon(row._tr))
    outside=[canon(n) for n in doc._element.body if n is not table._tbl]
    titles={k:canon(c.paragraphs[0]._p) for k,c in cells.items()}
    return rows,outside,titles,canon(table._tbl.tblPr)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('template');p.add_argument('report');p.add_argument('--pdf')
    a=p.parse_args(); errors=[]
    if fixed(Document(a.template))!=fixed(Document(a.report)): errors.append('固定区、栏目标题或表格属性发生变化')
    if a.pdf:
        pdf=fitz.open(a.pdf)
        def pages(label): return [i+1 for i,page in enumerate(pdf) if label in ''.join(page.get_text().split())]
        groups=[['实验目的：','实验原理：','实验项目名称：'],['实验结果及分析：','实验小结：','学生学习态度：','安全清洁及其他：','教师批阅：','教师签名']]
        for group in groups:
            mapping={label:pages(label) for label in group}
            if any(len(v)!=1 for v in mapping.values()) or len({tuple(v) for v in mapping.values()})!=1: errors.append(str(mapping))
    if errors:
        print('\n'.join(errors));raise SystemExit(1)
    print('固定XML检查通过。' + ('PDF栏目同页检查通过。' if a.pdf else '未提供PDF，尚未验证分页。') + '仍需检查每页PNG有无裁切与溢出。')
