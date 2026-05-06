# -*- coding: utf-8 -*-
import os
import re
import uuid
import json
from typing import List, Optional, Union
from urllib.parse import urlparse, quote
from datetime import datetime, timedelta

import requests
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 8000))
BASE_URL = os.environ.get("BASE_URL", f"http://127.0.0.1:{PORT}")

file_store = {}

app = FastAPI(title="制式文档生成API", description="Word模板生成服务", version="1.3.0-FINAL")

@app.exception_handler(Exception)
async def global_err(request: Request, exc: Exception):
    return JSONResponse(status_code=500, content={"detail": f"错误：{str(exc)}"})

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class GenerateRequest(BaseModel):
    template_file_url: str = "https://raw.githubusercontent.com/3293203203411/docx-generator/main/gaoqilixiangmoban.docx"
    text_keys: Union[List[str], str]
    text_values: Union[List[str], str]
    filename: Optional[str] = None


def parse_json_param(param):
    if param is None:
        return []
    if isinstance(param, list):
        return [str(item) for item in param]
    if isinstance(param, str):
        try:
            parsed = json.loads(param)
            if isinstance(parsed, list):
                return [str(i) for i in parsed]
            return [str(parsed)]
        except json.JSONDecodeError:
            return [param]
    return [str(param)]


def download_file(url: str) -> bytes:
    try:
        r = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=30)
        r.raise_for_status()
        return r.content
    except:
        raise HTTPException(status_code=400, detail="模板文件下载失败")


def merge_runs_in_paragraph(para):
    txt = "".join(r.text for r in para.runs)
    if not para.runs:
        return {"text": txt, "fmt": None}
    r = para.runs[0]
    return {
        "text": txt,
        "fmt": {
            "bold": r.bold,
            "italic": r.italic,
            "underline": r.underline,
            "font.name": r.font.name,
            "font.size": r.font.size,
        }
    }


def replace_placeholders_in_text(text, keys, values):
    for k, v in zip(keys, values):
        text = re.sub(r'\{\{\s*' + re.escape(k) + r'\s*\}\}', str(v), text)
    return text


def has_color_tag(s):
    return '<red>' in s and '</red>' in s


def parse_colored_segments(v):
    segs = []
    pat = re.compile(r'<red>(.*?)</red>', re.DOTALL)
    last = 0
    for m in pat.finditer(v):
        if m.start() > last:
            segs.append((v[last:m.start()], False))
        segs.append((m.group(1), True))
        last = m.end()
    if last < len(v):
        segs.append((v[last:], False))
    return segs if segs else [(v, False)]


def set_cell_border(cell):
    tc = cell._tc
    pr = tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for s in ['top', 'left', 'bottom', 'right']:
        e = OxmlElement(f'w:{s}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:color'), '000000')
        b.append(e)
    pr.append(b)


def set_cell_bg(cell, color):
    pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    pr.append(shd)


def set_cell_text(cell, text, bold=False, sz=10.5):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(sz)
    r.font.name = '宋体'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')


def parse_pipe_table(text):
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    return [[c.strip() for c in l.split('|')] for l in lines]


def replace_placeholder_with_auto_table(doc, placeholder, table_text):
    target = None
    for elem in doc.element.body:
        if isinstance(elem, CT_P):
            p = Paragraph(elem, doc)
            if placeholder in ''.join(r.text for r in p.runs):
                target = p
                break
    if not target:
        return False

    data = parse_pipe_table(table_text)
    if not data:
        return False

    table = doc.add_table(rows=len(data), cols=max(len(r) for r in data))
    table.style = 'Table Grid'

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            if j >= len(table.rows[i].cells):
                continue
            c = table.cell(i, j)
            set_cell_border(c)
            if i == 0:
                set_cell_bg(c, 'E6E6E6')
                set_cell_text(c, val, bold=True)
            else:
                if i % 2 == 1:
                    set_cell_bg(c, 'F5F5F5')
                set_cell_text(c, val)

    target._element.addprevious(table._tbl)
    target._element.getparent().remove(target._element)
    return True


def process_paragraph(para, keys, values):
    if not para.runs:
        return
    info = merge_runs_in_paragraph(para)
    txt = info['text']
    if not any(re.search(r'\{\{\s*' + re.escape(k) + r'\s*\}\}', txt) for k in keys):
        return
    new_txt = replace_placeholders_in_text(txt, keys, values)
    if new_txt == txt:
        return

    fmt = info['fmt']
    pe = para._p
    for r in para.runs:
        pe.remove(r._r)

    lines = new_txt.split('\n')
    for line_idx, line in enumerate(lines):
        segs = parse_colored_segments(line)
        for seg, red in segs:
            if not seg:
                continue
            r = OxmlElement('w:r')
            pr = OxmlElement('w:rPr')

            if fmt and fmt.get('font.name'):
                f = OxmlElement('w:rFonts')
                f.set(qn('w:eastAsia'), fmt['font.name'])
                pr.append(f)
            if fmt and fmt.get('bold'):
                pr.append(OxmlElement('w:b'))
            if fmt and fmt.get('font.size'):
                try:
                    sz = str(int(fmt['font.size'].pt * 2))
                    pr.append(OxmlElement('w:sz')).set(qn('w:val'), sz)
                except:
                    pass

            clr = OxmlElement('w:color')
            clr.set(qn('w:val'), 'FF0000' if red else '000000')
            pr.append(clr)
            r.append(pr)

            t = OxmlElement('w:t')
            t.text = seg
            if seg.startswith(' ') or seg.endswith(' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.append(t)
            pe.append(r)

        if line_idx < len(lines) - 1:
            br = OxmlElement('w:r')
            br.append(OxmlElement('w:br'))
            pe.append(br)


def process_table(table, keys, values):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                process_paragraph(p, keys, values)


def process_document(doc, keys, values):
    atk = "性能指标表格"
    idx = next((i for i, k in enumerate(keys) if k == atk), None)
    if idx is not None:
        replace_placeholder_with_auto_table(doc, "{{"+atk+"}}", values[idx])
        keys.pop(idx)
        values.pop(idx)

    for elem in doc.element.body:
        if isinstance(elem, CT_P):
            process_paragraph(Paragraph(elem, doc), keys, values)
    for t in doc.tables:
        process_table(t, keys, values)


def gen_fn(custom=None):
    if custom:
        return custom if custom.endswith('.docx') else custom + '.docx'
    return f"output_{uuid.uuid4().hex[:8]}.docx"


@app.get("/")
def root():
    return {"status": "running", "version": "1.3.0-FINAL"}


@app.post("/generate")
def api_gen(req: GenerateRequest):
    keys = parse_json_param(req.text_keys)
    vals = parse_json_param(req.text_values)
    if len(keys) != len(vals):
        raise HTTPException(status_code=400, detail="键值数量不匹配")

    content = download_file(req.template_file_url)
    with open('/tmp/tmp.docx', 'wb') as f:
        f.write(content)
    doc = Document('/tmp/tmp.docx')
    process_document(doc, keys, vals)

    import io
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    data = buf.getvalue()

    fn = gen_fn(req.filename)
    token = uuid.uuid4().hex
    file_store[token] = {
        "data": data,
        "name": fn,
        "exp": datetime.now() + timedelta(hours=2)
    }

    return {
        "success": True,
        "filename": fn,
        "full_download_url": f"{BASE_URL}/dl/{token}"
    }


@app.get("/dl/{token}")
def dl(token: str):
    if token not in file_store:
        raise HTTPException(status_code=404, detail="文件不存在")
    item = file_store[token]
    if datetime.now() > item["exp"]:
        del file_store[token]
        raise HTTPException(status_code=404, detail="文件已过期")
    return Response(
        content=item["data"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename*=utf-8\'\'{quote(item["name"])}'}
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=HOST, port=PORT)
