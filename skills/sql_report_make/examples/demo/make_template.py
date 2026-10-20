# -*- coding: utf-8 -*-
"""生成 demo 用的空白母本 template.docx（A4、页边距 72/90 磅、宋体 12 磅）。

真实使用时把配置里的 `template` 指向你自己的 `<序号>_command.docx`（学校发的实验指导书）即可，
本脚本只是为了让仓库里的示例在没有学校文件的情况下也能直接跑通——生成器只从母本取
页面设置与样式。

用法：
    python examples/demo/make_template.py
"""
import os

from docx import Document
from docx.shared import Pt


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "template.docx")

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Pt(595.3)          # A4 宽
    sec.page_height = Pt(841.9)         # A4 高
    sec.top_margin = Pt(72)
    sec.bottom_margin = Pt(72)
    sec.left_margin = Pt(90)
    sec.right_margin = Pt(90)

    # 母本的 Normal 样式：宋体 12 磅（生成器会在此基础上统一改黑）
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(12)

    doc.save(out)
    print("已生成母本：", out)


if __name__ == "__main__":
    main()
