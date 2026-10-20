# -*- coding: utf-8 -*-
"""sql_report_make —— 由 Markdown 实验报告生成排版统一的 Word 实验报告。

用法：
    python make_sql_report.py <配置文件.json>

配置文件示例见 examples/config-exp1.json。脚本做五件事：
  1. 解析 Markdown（标题 / 段落 / 列表 / 引用 / 表格 / 代码块 / 分隔线 / <form> 报告表）；
  2. 按统一规范排版成 docx：全文黑色、统一固定行距、代码块等宽字体；
  3. 全角空格统一换成两个半角空格；章节标题可按 label_map 映射成报告表栏目名；
  4. 【可选】把 <form> 块渲染成母本尾部那张 6 列报告表，并插入表格单元格；
  5. 保存并自检：正文颜色、实际 run 颜色、使用样板式颜色、行距取值、段落/表格数。
     另可选生成“目录”表（PAGEREF 域），默认关闭。

依赖：python-docx（必需）、Pillow（可选，仅在需要插入截图时使用）。
"""
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

BLACK = RGBColor(0, 0, 0)
CENTER = WD_ALIGN_PARAGRAPH.CENTER
LEFT = WD_ALIGN_PARAGRAPH.LEFT

# 行距取值（磅，固定值）：正文 / 代码 / 表格 三档，全篇只用这三种，避免行距忽大忽小
LINE_BODY = 18
LINE_CODE = 12
LINE_TABLE = 14


# ==========================================================================
# 一、配置
# ==========================================================================
def load_config(path):
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    base = os.path.dirname(os.path.abspath(path))

    def resolve(p):
        if not p:
            return None
        return p if os.path.isabs(p) else os.path.normpath(os.path.join(base, p))

    if not cfg.get("markdown"):
        raise SystemExit("配置缺少 markdown 字段")
    cfg["markdown"] = resolve(cfg["markdown"])
    cfg["out"] = resolve(cfg.get("out") or "实验报告.docx")
    cfg["template"] = resolve(cfg.get("template"))
    if not os.path.exists(cfg["markdown"]):
        raise SystemExit(f"Markdown 不存在：{cfg['markdown']}")
    cfg.setdefault("fonts", {})
    cfg.setdefault("sizes", {})
    return cfg


def opt(cfg, key, default):
    return cfg.get(key, default)


# ==========================================================================
# 二、Markdown 解析
# ==========================================================================
def parse_markdown(text, counter=None):
    """把 Markdown 解析成块列表。

    counter：代码块编号用的可变计数器（list，counter[0] 为下一个编号）。
    文档按顺序解析，因此 <form> 内的代码块与正文代码块共用同一条编号序列——
    这与“配置里的 fence_to_sql 按代码块出现顺序”的约定一致。
    行内加粗/等宽标记保留在文本里，由渲染层解析。
    """
    if counter is None:
        counter = [0]
    lines = text.replace("\r\n", "\n").split("\n")
    blocks = []
    i, n = 0, len(lines)

    def is_table_sep(s):
        return bool(re.match(r"^\s*\|?[\s:\-|]+\|[\s:\-|]*$", s)) and "-" in s

    while i < n:
        line = lines[i]
        s = line.strip()

        # <form> ... </form>：教师报告表的内容块
        if s == "<form>":
            i += 1
            inside = {"title": "", "info_rows": [], "sections": []}
            cur = None
            while i < n and lines[i].strip() != "</form>":
                cur_line = lines[i]
                st = cur_line.strip()
                if st.startswith("<title>"):
                    inside["title"] = re.sub(r"</?title>", "", st).strip()
                elif st.startswith("<row>"):
                    row = re.sub(r"</?row>", "", st).strip()
                    cells = [c.strip() for c in row.split("|")]
                    one = {}
                    for k in range(0, len(cells) - 1, 2):
                        one[cells[k]] = cells[k + 1]
                    inside["info_rows"].append(one)
                elif st.startswith("<section"):
                    m = re.match(r'<section\s+name="([^"]*)"', st)
                    cur = {"name": m.group(1) if m else "", "lines": []}
                    inside["sections"].append(cur)
                elif st == "</section>":
                    cur = None
                elif cur is not None:
                    cur["lines"].append(cur_line)
                i += 1
            i += 1
            # 每段内容用小解析器再解析一次（共用编号计数器），拿到表格/代码块等块级结构
            for sec in inside["sections"]:
                sec["blocks"] = parse_markdown("\n".join(sec["lines"]), counter)
            blocks.append({"type": "form", "form": inside})
            continue

        # 代码块
        if s.startswith("```"):
            i += 1
            buf = []
            while i < n and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            blocks.append({"type": "code", "lines": buf, "fence_no": counter[0]})
            counter[0] += 1
            continue

        # 分隔线
        if re.match(r"^\s*(-{3,}|\*{3,}|_{3,})\s*$", line):
            blocks.append({"type": "hr"})
            i += 1
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            blocks.append({"type": "heading", "level": len(m.group(1)),
                           "text": m.group(2).strip()})
            i += 1
            continue

        # 表格：当前行含 |，下一行是分隔行
        if "|" in line and i + 1 < n and is_table_sep(lines[i + 1]):
            header = [c.strip() for c in line.strip().strip("|").split("|")]
            i += 2
            rows = []
            while i < n and "|" in lines[i] and lines[i].strip():
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(cells)
                i += 1
            blocks.append({"type": "table", "header": header, "rows": rows})
            continue

        # 引用
        if s.startswith(">"):
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append({"type": "quote", "lines": buf})
            continue

        # 无序列表
        if re.match(r"^\s*[-*+]\s+", line):
            items = []
            while i < n and re.match(r"^\s*[-*+]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*+]\s+", "", lines[i]).strip())
                i += 1
            blocks.append({"type": "ul", "items": items})
            continue

        # 有序列表
        if re.match(r"^\s*\d+[.)]\s+", line):
            items = []
            while i < n and re.match(r"^\s*\d+[.)]\s+", lines[i]):
                items.append(re.sub(r"^\s*\d+[.)]\s+", "", lines[i]).strip())
                i += 1
            blocks.append({"type": "ol", "items": items})
            continue

        # 空行
        if not s:
            i += 1
            continue

        # 普通段落
        buf = [s]
        i += 1
        while i < n and lines[i].strip() and not re.match(
                r"^\s*(#{1,6}\s|[-*+]\s|\d+[.)]\s|>|```|\|)", lines[i]) \
                and not re.match(r"^\s*(-{3,}|\*{3,}|_{3,})\s*$", lines[i]):
            buf.append(lines[i].strip())
            i += 1
        blocks.append({"type": "para", "text": " ".join(buf)})
    return blocks


def strip_inline(text):
    """去掉行内标记（**加粗**、`代码`、*斜体*、链接），只保留文字。"""
    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = text.replace("**", "").replace("__", "")
    text = re.sub(r"(?<!\w)\*(?!\s)([^*]+)\*", r"\1", text)
    text = text.replace("`", "")
    return text.strip()


def parse_inline(text):
    """把一行拆成 [(文字, 是否加粗, 是否等宽), ...]，用于**加粗**与`等宽`。"""
    out = []
    pattern = re.compile(r"\*\*(.+?)\*\*|`([^`]+)`")
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False, False))
        if m.group(1) is not None:
            out.append((m.group(1), True, False))
        else:
            out.append((m.group(2), False, True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False, False))
    return out or [(text, False, False)]


# ==========================================================================
# 三、docx 写入工具
# ==========================================================================
def set_line(par, line_pt, before=0, after=0, snap_off=True):
    """统一设置行距（固定值）+ 段前段后；并关闭“对齐到文档网格”。"""
    pf = par.paragraph_format
    pf.line_spacing = Pt(line_pt)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if snap_off:
        ppr = par._element.get_or_add_pPr()
        snap = ppr.find(qn("w:snapToGrid"))
        if snap is None:
            snap = OxmlElement("w:snapToGrid")
            ppr.insert(0, snap)
        snap.set(qn("w:val"), "0")


def style_run(run, font, size_pt, bold=False, color=BLACK, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color          # 统一黑色
    run.font.name = font
    run.font.size = Pt(size_pt)
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), font)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)


class Builder:
    def __init__(self, cfg, doc, toc_enabled=False):
        self.cfg = cfg
        self.doc = doc
        self.toc_enabled = toc_enabled
        self.font_body = opt(cfg["fonts"], "body", "宋体")
        self.font_heading = opt(cfg["fonts"], "heading", "黑体")
        self.font_code = opt(cfg["fonts"], "code", "Consolas")
        self.sz_body = opt(cfg["sizes"], "body_pt", 12)
        self.sz_code = opt(cfg["sizes"], "code_pt", 9)
        self.sz_table = opt(cfg["sizes"], "table_pt", 10.5)
        self.headings = []              # [(序号, 级别, 文本)]
        # 全角空格统一换成两个半角空格；标题可映射成报告表里的栏目名
        self.space_fix = opt(cfg, "space_fix", True)
        self.label_map = opt(cfg, "label_map", {}) or {}

    def normalize(self, text):
        """去行内标记；按需把全角空格 U+3000 换成两个半角空格。"""
        text = strip_inline(text)
        if self.space_fix:
            text = text.replace("\u3000", "  ").replace("\\u3000", "  ")
        return text

    def label(self, text):
        return self.label_map.get(text, text)

    # ---------- 基本段落 ----------
    def para(self, text="", font=None, size=None, bold=False, align=None,
             line=LINE_BODY, before=0, after=0, indent_chars=0):
        p = self.doc.add_paragraph()
        if align is not None:
            p.alignment = align
        # 空段落（表格后的间隔）用正文字号的行距，避免出现“细行”
        use_line = line if text else max(line, LINE_BODY)
        set_line(p, use_line, before, after)
        if indent_chars:
            p.paragraph_format.first_line_indent = Pt(
                indent_chars * (size or self.sz_body))
        if text:
            run = p.add_run(text)
            style_run(run, font or self.font_body, size or self.sz_body, bold)
        return p

    def heading(self, level, text):
        """标题：按 label_map 换成报告表栏目名；只有开启目录时才写书签。"""
        text = self.label(text)
        if level == 1:
            size, font, align, before, after = 16, self.font_heading, CENTER, 0, 10
        elif level == 2:
            size, font, align, before, after = 14, self.font_heading, None, 10, 6
        else:
            size, font, align, before, after = 12, self.font_heading, None, 8, 4
        p = self.para(text, font=font, size=size, bold=True, align=align,
                      line=LINE_BODY, before=before, after=after)
        bid = len(self.headings) + 1
        self.headings.append((bid, level, text))
        if self.toc_enabled:
            self._add_bookmark(p, f"_H{bid}", bid)
        return p

    @staticmethod
    def _add_bookmark(par, name, bid):
        """书签 id 用确定性的序号，避免 hash() 随机化导致每次生成结果不同。"""
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(bid))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(bid))
        par._element.insert(0, start)
        par._element.append(end)

    def code_block(self, lines):
        """把代码块写成普通等宽段落（不参与 .sql 替换的场景，例如报告表内）。"""
        for raw in lines:
            text = raw.rstrip()
            if self.space_fix:
                text = text.replace("\u3000", "  ").replace("\\u3000", "  ")
            # 空行用零宽空格占位：纯空格的段落会被 WPS 保存时丢弃
            p = self.para(text if text.strip() else "\u200b",
                          font=self.font_code, size=self.sz_code,
                          line=LINE_CODE, before=0, after=0)
            p.paragraph_format.first_line_indent = Pt(0)

    def quote(self, lines):
        for line in lines:
            p = self.para(self.normalize(line), size=self.sz_body - 0.5,
                          line=LINE_BODY, before=0, after=0)
            p.paragraph_format.left_indent = Pt(24)

    def list_block(self, items, ordered):
        for k, item in enumerate(items, 1):
            mark = f"{k}. " if ordered else "· "
            p = self.para(mark + self.normalize(item), line=LINE_BODY,
                          before=0, after=0)
            p.paragraph_format.left_indent = Pt(24)
            p.paragraph_format.first_line_indent = Pt(-12)

    def hr(self):
        p = self.doc.add_paragraph()
        set_line(p, LINE_BODY, 4, 4)
        ppr = p._element.get_or_add_pPr()
        borders = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:color"), "000000")
        borders.append(bottom)
        ppr.append(borders)

    # ---------- 表格 ----------
    def table(self, header, rows, widths=None):
        ncol = len(header)
        t = self.doc.add_table(rows=1 + len(rows), cols=ncol)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = True

        for j, cell_text in enumerate(header):
            self._fill_cell(t.cell(0, j), self.normalize(cell_text), bold=True,
                            align=CENTER)
        for ri, row in enumerate(rows, 1):
            for j in range(ncol):
                txt = self.normalize(row[j]) if j < len(row) else ""
                align = LEFT
                if re.fullmatch(r"\d+(\.\d+)?", txt):
                    align = CENTER
                self._fill_cell(t.cell(ri, j), txt, bold=False, align=align)
        self.para("", line=LINE_BODY, after=0)
        return t

    def _fill_cell(self, cell, text, bold=False, align=LEFT):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align
        set_line(p, LINE_TABLE, 0, 0)
        p.paragraph_format.first_line_indent = Pt(0)
        run = p.add_run(text)
        style_run(run, self.font_body, self.sz_table, bold)

    # ---------- 教师报告表 ----------
    def form_table(self, form, code_renderer=None):
        """按母本尾部那张 6 列报告表排版：标题行 / 信息行 / 栏目内容行。

        code_renderer(lines) 用于把报告表里的代码块按统一规范逐行着色；
        传 None 时用普通等宽段落。
        """
        title = form.get("title") or ""
        info_rows = form.get("info_rows") or ([form.get("info")] if form.get("info") else [])
        info_rows = [r for r in info_rows if r]
        sections = form.get("sections") or []

        t = self.doc.add_table(rows=2 + len(info_rows), cols=6)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = True

        # 第 1 行：标题（6 列合并）
        self._fill_cell(t.cell(0, 0), self.normalize(title), bold=True,
                        align=CENTER)
        for j in range(1, 6):
            self._fill_cell(t.cell(0, j), "", align=CENTER)
        self._merge_row(t.rows[0])

        # 信息行：题目/姓名/日期 等（键居中、值左对齐）
        for ri, info in enumerate(info_rows, 1):
            pairs = []
            for k, v in info.items():
                pairs.extend([self.normalize(k), self.normalize(v)])
            for j in range(6):
                txt = pairs[j] if j < len(pairs) else ""
                self._fill_cell(t.cell(ri, j), txt,
                                align=CENTER if j % 2 == 0 else LEFT)

        # 最后一行：5 个栏目（加粗栏目名 + 内容）
        cell = t.cell(len(info_rows) + 1, 0)
        cell.text = ""
        first = True

        def add_cell_par():
            nonlocal first
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            p.alignment = LEFT
            set_line(p, LINE_TABLE, 0, 0)
            p.paragraph_format.first_line_indent = Pt(0)
            return p

        for sec in sections:
            # 栏目名
            lp = add_cell_par()
            run = lp.add_run(sec.get("name", ""))
            style_run(run, self.font_heading, self.sz_body, bold=True)
            for blk in sec.get("blocks", []):
                bt = blk.get("type")
                if bt == "code":
                    if code_renderer is not None:
                        code_renderer(blk["lines"], cell, blk.get("fence_no"))
                    continue
                if bt == "table":
                    self._cell_table(cell, blk["header"], blk["rows"])
                    continue
                if bt in ("ul", "ol"):
                    for k, item in enumerate(blk["items"], 1):
                        mark = f"{k}. " if bt == "ol" else "· "
                        p = add_cell_par()
                        self._cell_runs(p, mark + item)
                    continue
                if bt == "hr":
                    p = add_cell_par()
                    ppr = p._element.get_or_add_pPr()
                    bd = OxmlElement("w:pBdr")
                    bottom = OxmlElement("w:bottom")
                    bottom.set(qn("w:val"), "single")
                    bottom.set(qn("w:sz"), "6")
                    bottom.set(qn("w:color"), "000000")
                    bd.append(bottom)
                    ppr.append(bd)
                    continue
                if bt == "heading":
                    p = add_cell_par()
                    r = p.add_run(self.normalize(blk["text"]))
                    style_run(r, self.font_heading, self.sz_body, bold=True)
                    continue
                # 普通段落
                text = blk.get("text", "")
                if not text.strip():
                    continue
                p = add_cell_par()
                self._cell_runs(p, text)

        self._merge_row(t.rows[len(info_rows) + 1])
        self.para("", line=LINE_BODY, after=0)
        return t

    def _cell_runs(self, p, text):
        """把一个段落写进单元格：识别 **加粗** 与 `等宽`。

        注意：必须先 parse_inline() 再对每段单独去标记——反过来会把 ** 和 ` 删掉，
        导致表格里的加粗/等宽全部丢失。
        """
        for seg, bold, mono in parse_inline(text):
            seg = self.normalize(seg)
            if not seg:
                continue
            r = p.add_run(seg)
            style_run(r, self.font_code if mono else self.font_body,
                      self.sz_code if mono else self.sz_table, bold=bold)

    def _cell_table(self, cell, header, rows):
        """在报告表单元格里嵌一张小表（多行数据表用）。"""
        ncol = len(header)
        t = cell.add_table(rows=1 + len(rows), cols=ncol)
        t.style = "Table Grid"
        for j, txt in enumerate(header):
            self._fill_cell(t.cell(0, j), self.normalize(txt), bold=True,
                            align=CENTER)
        for ri, row in enumerate(rows, 1):
            for j in range(ncol):
                txt = self.normalize(row[j]) if j < len(row) else ""
                self._fill_cell(t.cell(ri, j), txt, align=LEFT)

    @staticmethod
    def _merge_row(row):
        """把一行的所有单元格横向合并（row.cells 的底层 tc 保留第一个）。"""
        tcs = row._tr.findall(qn("w:tc"))
        if len(tcs) < 2:
            return
        first = tcs[0]
        tcpr = first.find(qn("w:tcPr"))
        if tcpr is None:
            tcpr = OxmlElement("w:tcPr")
            first.insert(0, tcpr)
        span = OxmlElement("w:gridSpan")
        span.set(qn("w:val"), str(len(tcs)))
        tcpr.append(span)
        for tc in tcs[1:]:
            row._tr.remove(tc)

    # ---------- 目录 ----------
    def toc(self, title):
        """生成目录表：左列标题名，右列 PAGEREF 域（打开文档时自动算出页码）。"""
        self.para(title, font=self.font_heading, size=14, bold=True,
                  line=LINE_BODY, before=0, after=6)
        items = [(bid, lvl, txt) for bid, lvl, txt in self.headings if lvl <= 3]
        t = self.doc.add_table(rows=1 + len(items), cols=2)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._fill_cell(t.cell(0, 0), "章　节", bold=True, align=CENTER)
        self._fill_cell(t.cell(0, 1), "页码", bold=True, align=CENTER)
        for ri, (bid, lvl, txt) in enumerate(items, 1):
            name = ("　" * (lvl - 2)) + txt if lvl >= 2 else txt
            self._fill_cell(t.cell(ri, 0), name, align=LEFT)
            cell = t.cell(ri, 1)
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = CENTER
            set_line(p, LINE_TABLE, 0, 0)
            run = p.add_run()
            style_run(run, self.font_body, self.sz_table)
            self._add_pageref(run, f"_H{bid}")
        self.para("", line=LINE_BODY, after=0)

    @staticmethod
    def _add_pageref(run, bookmark):
        """插入 PAGEREF 域（含结果占位文本），打开文档时自动显示页码。"""
        parent = run._element.getparent()
        idx = list(parent).index(run._element)
        begin = OxmlElement("w:r")
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        begin.append(fld_begin)
        instr_r = OxmlElement("w:r")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f" PAGEREF {bookmark} \\h "
        instr_r.append(instr)
        result = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "000000")
        rpr.append(color)
        result.append(rpr)
        t = OxmlElement("w:t")
        t.text = "1"
        result.append(t)
        end = OxmlElement("w:r")
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        end.append(fld_end)
        for k, node in enumerate([begin, instr_r, result, end]):
            parent.insert(idx + k, node)


# ==========================================================================
# 四、代码块着色（可选）
# ==========================================================================
def build_highlight_rules(spec):
    """把配置里的 {pattern: color} 编译成 [(regex, color), ...]，长模式优先。"""
    rules = []
    for pattern, color in (spec or {}).items():
        rules.append((re.compile(pattern, re.I), color))
    rules.sort(key=lambda r: -len(r[0].pattern))
    return rules


def highlight(par, code, rules, font, size):
    """按规则把一行代码切成多段着色；未匹配部分用黑色。

    rules: [(regex, color), (regex, color), ...]，按顺序在当前位置尝试匹配，
    取“匹配最长”的那条（保证字符串/注释优先于关键字）。
    """
    pos, n = 0, len(code)
    while pos < n:
        best_m, best_color = None, None
        for rx, color in rules:
            m = rx.match(code, pos)
            if m and m.end() > pos and (best_m is None
                                        or m.end() > best_m.end()):
                best_m, best_color = m, color
        if best_m is None:
            run = par.add_run(code[pos:])
            style_run(run, font, size, color=BLACK)
            break
        run = par.add_run(best_m.group(0))
        style_run(run, font, size, color=RGBColor.from_string(best_color))
        pos = best_m.end()


# ==========================================================================
# 五、主流程
# ==========================================================================
def force_black_styles(doc):
    """把所有样式的字体颜色改成黑色，防止样式里带彩色（如语法高亮的 Source Code）。"""
    changed = 0
    for style in doc.styles:
        try:
            rpr = style.element.find(qn("w:rPr"))
        except Exception:
            continue
        if rpr is None:
            continue
        for color in rpr.findall(qn("w:color")):
            if color.get(qn("w:val")) not in ("auto", "000000"):
                color.set(qn("w:val"), "000000")
                changed += 1
    return changed


def force_black_theme(path):
    """把主题里的配色方案全部改成黑色：主题强调色一旦被样式引用就会带颜色。"""
    import shutil
    import zipfile

    if not os.path.exists(path):
        return 0
    tmp = path + ".tmp"
    changed = 0
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("word/theme/") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                new = re.sub(r'(<a:(?:srgbClr|sysClr)[^>]*?val=")[0-9A-Fa-f]{6}(")',
                             r"\g<1>000000\g<2>", text)
                new = re.sub(r'(<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
                             r"\g<1>000000\g<2>", new)
                changed += len(re.findall(r'val="000000"', new)) - \
                    len(re.findall(r'val="000000"', text))
                data = new.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return changed


def set_page_setup(doc, cfg):
    """按母本设置页边距，并关闭文档网格（否则行距会被网格重排）。"""
    sec = doc.sections[0]
    sec.page_width = Pt(595.3)          # A4 宽
    sec.page_height = Pt(841.9)         # A4 高
    sec.top_margin = Pt(72)
    sec.bottom_margin = Pt(72)
    sec.left_margin = Pt(90)
    sec.right_margin = Pt(90)
    sectPr = sec._sectPr
    for grid in sectPr.findall(qn("w:docGrid")):
        grid.set(qn("w:type"), "default")


def check(path):
    """自检：正文颜色、行距取值、段落数、表格数。

    说明：Word 内置样式（styles.xml 的 latentStyles 与未使用的表格样式）里
    本身就带颜色定义（如 FFFFFF、1F497D），它们不参与实际渲染，因此只检查
    document.xml 的颜色和“文档实际用到的样式”的颜色。
    """
    from collections import Counter
    import zipfile
    with zipfile.ZipFile(path) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")
    d = Document(path)

    doc_colors = Counter(re.findall(r'<w:color w:val="([^"]+)"', doc_xml))
    run_colors = Counter()
    for p in d.paragraphs:
        for r in p.runs:
            if r.font.color and r.font.color.rgb:
                run_colors[str(r.font.color.rgb)] += 1
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        if r.font.color and r.font.color.rgb:
                            run_colors[str(r.font.color.rgb)] += 1
    used = {p.style.name for p in d.paragraphs}
    used_style_colors = []
    for s in d.styles:
        try:
            if s.name not in used:
                continue
        except Exception:
            continue
        rpr = s.element.find(qn("w:rPr"))
        if rpr is None:
            continue
        for c in rpr.findall(qn("w:color")):
            v = c.get(qn("w:val"))
            if v not in ("auto", "000000"):
                used_style_colors.append((s.name, v))

    # 递归统计（报告表里的数据表是嵌套表格，不递归会漏掉）
    def walk(container):
        for p in container.paragraphs:
            yield p
        for tb in container.tables:
            for row in tb.rows:
                seen = set()
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    yield from walk(c)

    all_pars = list(walk(d))
    n_tables = doc_xml.count("<w:tbl>")

    lines = Counter(re.findall(r'w:line="(\d+)" w:lineRule="(\w+)"', doc_xml))
    print()
    print("已生成：", path)
    print("  顶层段落数：", len(d.paragraphs),
          "；含表格内共：", len(all_pars))
    print("  表格数（含嵌套）：", n_tables)
    print("  正文颜色：", dict(doc_colors))
    print("  实际 run 颜色：", dict(run_colors))
    print("  实际使用样式：", sorted(used))
    print("  使用样式中的非黑颜色：",
          used_style_colors if used_style_colors else "无")
    print("  行距取值：", {f"{k[0]}twips/{k[1]}": v for k, v in lines.items()})
    print("  目录域/书签：", doc_xml.count("PAGEREF"), "/",
          doc_xml.count("bookmarkStart"))
    bad = [c for c in list(doc_colors) + list(run_colors)
           if c.upper() not in ("000000", "AUTO")]
    ok = not bad and not used_style_colors
    print("  结论：" + ("全文黑色、行距统一为三档固定值" if ok
                      else f"仍有非黑色：{bad or used_style_colors}"))
    return ok


def main():
    if len(sys.argv) < 2:
        raise SystemExit("用法：python make_sql_report.py <配置文件.json>")
    cfg = load_config(sys.argv[1])

    with open(cfg["markdown"], encoding="utf-8") as f:
        md = f.read()
    fence_counter = [0]
    blocks = parse_markdown(md, fence_counter)
    print(f"解析到 {len(blocks)} 个内容块，代码块 {fence_counter[0]} 个（含报告表内）")

    # 母本 <序号>_command.docx（工作区根目录）只用来取页面设置（纸张与页边距），报告结构由 Markdown 决定
    tpl = cfg.get("template")
    if tpl and os.path.exists(tpl):
        doc = Document(tpl)
        for p in list(doc.paragraphs):
            p._element.getparent().remove(p._element)
        for t in list(doc.tables):
            t._element.getparent().remove(t._element)
        print(f"  页面设置取自母本：{os.path.basename(tpl)}")
    else:
        doc = Document()
        set_page_setup(doc, cfg)
        if tpl:
            print(f"  [warn] 母本不存在，改用默认页面设置：{tpl}")
    # 默认样式先改成黑色，避免样式自带彩色
    normal = doc.styles["Normal"]
    normal.font.name = opt(cfg["fonts"], "body", "宋体")
    normal.font.size = Pt(opt(cfg["sizes"], "body_pt", 12))
    normal.font.color.rgb = BLACK
    normal._element.get_or_add_rPr().get_or_add_rFonts().set(
        qn("w:eastAsia"), opt(cfg["fonts"], "body", "宋体"))
    force_black_styles(doc)
    set_page_setup(doc, cfg)

    toc_cfg = cfg.get("toc") or {}
    b = Builder(cfg, doc, toc_enabled=bool(toc_cfg.get("enabled")))
    # 先扫描标题，保证标题序号在前（目录需要）
    for blk in blocks:
        if blk["type"] == "heading":
            b.headings.append((0, blk["level"], b.label(b.normalize(blk["text"]))))
    b.headings = [(i + 1, lvl, txt)
                  for i, (_, lvl, txt) in enumerate(b.headings)]
    hl_rules = build_highlight_rules(cfg.get("highlight"))
    sql_files = cfg.get("sql_files") or {}
    fence_targets = {int(k): v for k, v in (cfg.get("fence_to_sql") or {}).items()}
    applied = []                     # 实际命中的 fence_to_sql 编号
    total_fences = [0]               # 解析阶段记录到的代码块总数

    def render_code(lines, fence_no=None, cell=None):
        """渲染一个代码块：可替换为 .sql 源文件，并逐行着色。

        fence_no：该代码块在全文中的出现序号（含报告表内部），由解析阶段记录。
        cell 不为 None 时写进该单元格（报告表内容格），否则写在正文。
        """
        if fence_no is not None:
            total_fences[0] = max(total_fences[0], fence_no + 1)
            if fence_no in fence_targets and fence_no not in applied:
                target = fence_targets[fence_no]
                if target in sql_files:
                    p = sql_files[target]
                    if not os.path.isabs(p):
                        p = os.path.normpath(
                            os.path.join(os.path.dirname(cfg["markdown"]), p))
                    with open(p, encoding="utf-8") as f:
                        lines = f.read().rstrip("\n").split("\n")
                    applied.append(fence_no)
                    print(f"  代码块{fence_no} 改为引用源文件：{target}")
                else:
                    print(f"  [warn] fence_to_sql[{fence_no}] 指向的 {target} "
                          f"不在 sql_files 里，已忽略")
        for raw in lines:
            text = raw.rstrip() or "\u200b"
            if cell is not None:
                p = cell.add_paragraph()
            else:
                p = doc.add_paragraph()
            set_line(p, LINE_CODE, 0, 0)
            p.paragraph_format.first_line_indent = Pt(0)
            if hl_rules:
                highlight(p, text, hl_rules, b.font_code, b.sz_code)
            else:
                r = p.add_run(text)
                style_run(r, b.font_code, b.sz_code)
        return lines

    for blk in blocks:
        t = blk["type"]
        if t == "heading":
            b.heading(blk["level"], b.normalize(blk["text"]))
        elif t == "para":
            b.para(b.normalize(blk["text"]), line=LINE_BODY, after=0,
                   indent_chars=2)
        elif t == "ul":
            b.list_block(blk["items"], ordered=False)
        elif t == "ol":
            b.list_block(blk["items"], ordered=True)
        elif t == "quote":
            b.quote(blk["lines"])
        elif t == "hr":
            b.hr()
        elif t == "table":
            b.table(blk["header"], blk["rows"])
        elif t == "form":
            # <form> 内的代码块与正文共用编号（解析阶段已记录 fence_no）
            form = blk["form"]

            def form_code_renderer(lines, cell, fence_no):
                render_code(lines, fence_no=fence_no, cell=cell)

            b.form_table(form, code_renderer=form_code_renderer)
        elif t == "code":
            render_code(blk["lines"], blk.get("fence_no"))

    # fence_to_sql 的命中情况提示：配了映射却一个都没用到，通常意味着编号对不上
    if fence_targets:
        missed = sorted(set(fence_targets) - set(applied))
        print(f"  fence_to_sql：命中 {len(applied)}/{len(fence_targets)}"
              + (f"，未命中 {missed}（全文共 {total_fences[0]} 个代码块，"
                 f"请核对编号）" if missed else ""))

    # 附加：实验小结表（可选）
    sum_cfg = cfg.get("summary_table") or {}
    if sum_cfg.get("enabled"):
        b.para(sum_cfg.get("title", "附：实验小结"), font=b.font_heading,
               size=14, bold=True, line=LINE_BODY, before=10, after=6)
        b.table(sum_cfg.get("columns", []), sum_cfg.get("rows", []))

    # 目录：把刚生成的目录表（最后一个表格）及其标题段落移到正文最前面
    if toc_cfg.get("enabled"):
        b.toc(toc_cfg.get("title", "目　录"))
        body = doc.element.body
        tbls = [el for el in body.iter() if el.tag == qn("w:tbl")]
        if tbls:
            toc_tbl = tbls[-1]
            prev = toc_tbl.getprevious()
            nodes = [prev, toc_tbl] if prev is not None else [toc_tbl]
            for node in nodes:
                body.remove(node)
            for k, node in enumerate(nodes):
                body.insert(k, node)

    os.makedirs(os.path.dirname(cfg["out"]) or ".", exist_ok=True)
    doc.save(cfg["out"])

    # 二次检查：把样式里残留的彩色再抹一遍，并把主题配色也改成黑色
    d2 = Document(cfg["out"])
    changed = force_black_styles(d2)
    if changed:
        print(f"  清理样式彩色 {changed} 处")
        d2.save(cfg["out"])
    themed = force_black_theme(cfg["out"])
    if themed:
        print(f"  主题配色改为黑色 {themed} 处")
    ok = check(cfg["out"])
    print("  结论：" + ("全部为黑色，行距统一" if ok else "仍有非黑色，请检查"))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
