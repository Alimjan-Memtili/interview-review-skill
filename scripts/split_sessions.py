#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
split_sessions.py — 把总复盘录按场次切分成独立小 Word

思路：复制总档案，用 python-docx 删除不属于该场的段落/表格。
锚点：场次标题段落（sz=36, 蓝色, 文本以"第 N 场面试"开头）是章节边界。
优点：python-docx 正确处理命名空间，表格不丢失，内容与总档案逐字一致。

用法：
  python split_sessions.py --in 面试复盘录.docx --outdir 复盘档案
"""
import argparse
import os
import re
import sys
import shutil

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
except ImportError:
    sys.stderr.write("[split] python-docx required\n")
    raise

FONT = "微软雅黑"


def get_para_text_and_size(p):
    """返回 (文本, 最大字号或None, 是否蓝色标题)。"""
    text = ''.join(node.text or '' for node in p._element.findall('.//' + qn('w:t')))
    sz = None
    sizes = p._element.findall('.//' + qn('w:sz'))
    if sizes:
        vals = [int(s.get(qn('w:val'))) for s in sizes if s.get(qn('w:val'))]
        if vals:
            sz = max(vals)
    color_match = False
    colors = p._element.findall('.//' + qn('w:color'))
    for c in colors:
        if c.get(qn('w:val'), '').upper() == '2932E1':
            color_match = True
            break
    return text, sz, color_match


def is_session_title(p):
    """判断段落是否是场次标题（sz=36, 蓝色, 文本以'第 N 场面试'开头）。"""
    text, sz, color = get_para_text_and_size(p)
    return bool(re.match(r'^\s*第\s*\d+\s*场面试', text)) and sz == 36 and color


def get_session_round(p):
    """从场次标题提取轮次号（int）。非标题返回 None。"""
    if not is_session_title(p):
        return None
    text, _, _ = get_para_text_and_size(p)
    m = re.match(r'^\s*第\s*(\d+)\s*场面试', text)
    return int(m.group(1)) if m else None


def get_session_title_text(p):
    if is_session_title(p):
        text, _, _ = get_para_text_and_size(p)
        return text.strip()
    return None


def split_document(infile, outdir):
    os.makedirs(outdir, exist_ok=True)

    # 先扫描总档案，定位每个场次的标题段落 + 它在 body 中的元素索引范围
    doc = Document(infile)
    body = doc.element.body
    # body 的直接子元素：段落 <w:p> 和表格 <w:tbl>（还有 sectPr）
    children = list(body)
    P_TAG = qn('w:p')       # 完整命名空间 URI 形式
    TBL_TAG = qn('w:tbl')
    SECT_TAG = qn('w:sectPr')

    # 找所有场次标题的元素索引
    session_starts = []  # [(child_index, round_n, title_text)]
    for idx, child in enumerate(children):
        if child.tag == P_TAG:
            # 包装成 Paragraph 来复用函数
            from docx.text.paragraph import Paragraph
            p = Paragraph(child, doc)
            rnd = get_session_round(p)
            if rnd is not None:
                title = get_session_title_text(p)
                session_starts.append((idx, rnd, title))

    if not session_starts:
        sys.stderr.write("[split] 未找到任何场次标题\n")
        return 1

    # 每场的元素范围：[start_idx, next_session_start_idx 或 末尾(不含sectPr)]
    # 注意最后一个 body 子元素通常是 <w:sectPr>，要排除
    last_content_idx = len(children) - 1
    if children and children[-1].tag == SECT_TAG:
        last_content_idx = len(children) - 2

    ranges = []
    for i, (start_idx, rnd, title) in enumerate(session_starts):
        end_idx = session_starts[i + 1][0] if i + 1 < len(session_starts) else last_content_idx + 1
        ranges.append((rnd, title, start_idx, end_idx))

    sys.stderr.write(f"[split] 识别到 {len(ranges)} 场:\n")
    for rnd, title, s, e in ranges:
        sys.stderr.write(f"  - 第{rnd}场 [{s}:{e}]: {title[:40]}\n")

    # 为每场生成独立 docx：复制总档案 → 删除非该场元素
    for rnd, title, start_idx, end_idx in ranges:
        # 复制总档案作为基础
        # 解析文件名
        m = re.match(r'第\s*\d+\s*场面试\s*·\s*(.+?)\s*·\s*(\d{4}-\d{2}-\d{2})', title)
        if m:
            company_pos = m.group(1).strip()
            date = m.group(2)
            safe = re.sub(r'[·\s/\\:*?"<>|（）()]+', '_', company_pos).strip('_')
            fname = f"第{rnd}场_{safe}_{date}.docx"
        else:
            safe = re.sub(r'[·\s/\\:*?"<>|]+', '_', title).strip('_')
            fname = f"第{rnd}场_{safe}.docx"
        out_path = os.path.join(outdir, fname)

        # 复制源文件
        shutil.copy2(infile, out_path)
        d = Document(out_path)
        b = d.element.body
        kids = list(b)

        # 计算要保留的元素索引（在新文档里重新定位，因为复制后元素引用变了）
        # 重新扫描新文档找场次边界
        new_children = list(b)
        new_starts = []
        for idx, child in enumerate(new_children):
            if child.tag == P_TAG:
                from docx.text.paragraph import Paragraph
                p = Paragraph(child, d)
                r = get_session_round(p)
                if r is not None:
                    new_starts.append((idx, r))

        # 找本场在的新边界
        new_last = len(new_children) - 1
        if new_children and new_children[-1].tag == SECT_TAG:
            new_last = len(new_children) - 2

        my_start = None
        my_end = None
        for i, (idx, r) in enumerate(new_starts):
            if r == rnd:
                my_start = idx
                my_end = new_starts[i + 1][0] if i + 1 < len(new_starts) else new_last + 1
                break

        if my_start is None:
            sys.stderr.write(f"[split] 警告: 第{rnd}场在新文档中未定位到\n")
            continue

        # 删除：所有不在 [my_start, my_end) 的内容元素（保留 sectPr）
        # 从后往前删，避免索引错乱
        to_remove = []
        for idx, child in enumerate(new_children):
            if child.tag == SECT_TAG:
                continue  # 保留页面设置
            # 删除封面区域（第一个场次标题之前的所有内容）和非本场的
            keep = (my_start <= idx < my_end)
            if not keep:
                to_remove.append(child)
        for child in to_remove:
            b.remove(child)

        d.save(out_path)
        # 统计表格数
        n_tbl = len([c for c in list(b) if c.tag == TBL_TAG])
        sys.stderr.write(f"[split] 写出 {fname} (表格数={n_tbl})\n")

    return 0


def main():
    ap = argparse.ArgumentParser(description="按场次切分总复盘录为独立小word")
    ap.add_argument("--in", dest="infile", required=True, help="总复盘录 docx")
    ap.add_argument("--outdir", required=True, help="输出目录")
    args = ap.parse_args()
    return split_document(args.infile, args.outdir)


if __name__ == "__main__":
    sys.exit(main())
