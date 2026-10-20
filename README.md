# sql_skills · SQL 实验报告排版技能

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![test](https://img.shields.io/badge/tests-smoke%20passed-brightgreen.svg)](tests/test_demo.py)

把 **Markdown 写的数据库 / SQL 上机实验报告**，一键排成学校要的 Word 文档：
正文装进那张 6 列报告表（`1．实验目的：` ～ `实验总结：`），**全文黑色、三档固定行距、代码块等宽字体**，
报告里的代码块还能自动替换成真正执行过的 `SQL/*.sql` 源文件，保证「报告里的代码 = 跑过的代码」。

面向 AI 编程助手（Claude Code / DSH 等）的 **skill**：说一句「用 sql_report_make 排报告」就行；
也可以当普通命令行工具用。

## 解决什么问题

| 常见抱怨 | 原因 | 本技能的做法 |
| --- | --- | --- |
| 行距时大时小、太挤 | 原文没设行距，只能由字体默认行高决定；代码块与正文字号不同，交界处跳变 | 全篇只设三种**固定值**行距：正文 18 磅、代码 12 磅、表格 14 磅；每个段落关闭「对齐到文档网格」 |
| 颜色太多、花花绿绿 | 转换工具带来的 `source code` 语法高亮样式与主题配色 | 正文 run、样式定义、`theme1.xml` **三处一起改**成黑色 |
| 全角空格与半角空格混用 | 源文件里混着 `U+3000` | `space_fix` 统一换成两个半角空格 |
| 正文没装进学校的报告表 | 6 列表格，手抄容易错位 | Markdown 里用 `<form>` 块描述，生成器渲染成报告表并合并单元格 |
| 不要目录页 | 报告不需要目录 | `toc.enabled=false`，不生成目录表与 `PAGEREF` 域 |
| 代码块与脚本不一致 | 报告里是手抄的 | `fence_to_sql` 把指定代码块替换为 `SQL/*.sql` 源文件内容（⚠ 编号按全文顺序，改 md 后必须重新核对） |
| 代码里一堆装饰性横幅注释 | 脚本开头写 `/* ==== 实验1 … 任务N：… ==== */` | 脚本只保留 `--` 逐行说明注释 |

## 生成的报告长什么样

```
┌──────────────────────── 实验报告（6 列合并）────────────────────────┐
│ 题目 │ 实验1  数据库定义实验 │ 姓名 │ 张三 │ 学号 │ 2025001          │
│ 日期 │ 2026-10-08          │ 实验环境 │ …    │ 实验数据 │ …           │
├─────────────────────────────────────────────────────────────────────┤
│ 1．实验目的：                                                        │
│    …正文 12 磅、18 磅固定行距、首行缩进 2 字符…                        │
│ 实验内容和要求：                                                     │
│    ├ 有序/无序列表                                                   │
│    ├ 嵌套数据表（Table Grid 边框）                                    │
│ 实验重点和难点：                                                     │
│ 实验程序：                                                           │
│    ```sql 代码块（9 磅 Consolas、12 磅固定行距、逐行注释）```          │
│ 实验总结：                                                           │
└─────────────────────────────────────────────────────────────────────┘
## 附：实验脚本清单（报告表之外的小表）
```

## 安装

技能本体就是 `skills/sql_report_make/` 这一个文件夹（`SKILL.md` + `scripts/` + `examples/` + `reference/`），
把它复制到助手的 skills 目录即可：

```bash
# Claude Code
git clone <本仓库地址> && cp -r <本仓库>/skills/sql_report_make ~/.claude/skills/

# DSH（Windows 下为 %USERPROFILE%\.dsh\skills）
cp -r <本仓库>/skills/sql_report_make ~/.dsh/skills/
```

Windows PowerShell：

```powershell
Copy-Item -Recurse .\skills\sql_report_make "$env:USERPROFILE\.claude\skills\"
```

也可以只当命令行工具用，不必装进助手：

```bash
pip install python-docx
python skills/sql_report_make/scripts/make_sql_report.py 你的配置.json
```

> 只在 Python 3.9+ 上测过，唯一必需依赖是 `python-docx`（Windows / macOS / Linux 都行，产物是标准 `.docx`）。

## 快速开始（跑仓库自带的示例）

示例在 `skills/sql_report_make/examples/demo/`，自带一份空白母本和两段 SQL，**不需要学校的指导书**：

```bash
pip install python-docx
python skills/sql_report_make/scripts/make_sql_report.py skills/sql_report_make/examples/demo/config-example.json
# 产物：skills/sql_report_make/examples/demo/out/实验报告.docx
```

生成器会打印自检结果：

```
解析到 4 个内容块，代码块 2 个（含报告表内）
  页面设置取自母本：template.docx
  代码块0 改为引用源文件：01_create_database.sql
  代码块1 改为引用源文件：02_create_tables.sql
  fence_to_sql：命中 2/2
  主题配色改为黑色 10 处
  正文颜色： {'000000': 97}
  行距取值： {'360twips/exact': 4, '280twips/exact': 48, '240twips/exact': 31}
  结论：全文黑色、行距统一为三档固定值
```

## 在真实作业里怎么用

1. 把 `examples/demo/config-example.json` 复制成 `<序号>/config-exp<序号>.json`，和 `SQL/`、`实验报告.md` 放在同一个序号文件夹；
2. 改配置里的 `template`（指向学校发的 `<序号>_command.docx`）、`markdown`、`out`、`sql_files`、`fence_to_sql`；
3. 跑一遍生成器，打开 docx 核对报告表栏目、嵌套表格、每个代码块的内容。

`template` 指向的母本**只用来取页面设置（纸张、页边距）和样式**，报告结构完全由 Markdown 决定——
不要理解成「往母本的表格里填字」。

## Markdown 怎么写

正文列在 `<form>` 块里，栏目名照学校报告表的原样写（实验一 5 个；实验二 4 个，没有「实验重点和难点」）：

```markdown
# 实验报告

<form>
<title>实验报告</title>
<row>题目 | 实验1  数据库定义实验 | 姓名 | 张三 | 学号 | 2025001</row>
<row>日期 | 2026-10-08 | 实验环境 | MySQL 8.0 | 实验数据 | 学生选课库</row>

<section name="1．实验目的：">
掌握 DDL 的基本用法。
</section>

<section name="实验内容和要求：">
1. 创建数据库 S；
2. 创建表 S、P，设置主键与外键。
</section>

<section name="实验重点和难点："> ... </section>
<section name="实验程序：">

```sql
CREATE DATABASE S;
```

</section>
<section name="实验总结："> ... </section>
</form>

## 附：实验脚本清单
| 文件 | 对应任务 | 说明 |
```

| 标签 / 语法 | 作用 |
| --- | --- |
| `<form>` … `</form>` | 报告表区域，里面的内容渲染进 6 列表格 |
| `<title>…</title>` | 报告表第 1 行标题（6 列合并） |
| `<row>键 | 值 | 键 | 值 | 键 | 值</row>` | 一行信息（题目/姓名/学号 或 日期/实验环境/实验数据），3 对一组填满 6 列 |
| `<section name="栏目名：">` … `</section>` | 一个栏目；栏内支持段落、列表、引用、表格、代码块、`---` 分隔线、`**加粗**`、`` `等宽` `` |
| `<form>` 之外的 Markdown | 附录章节（如脚本清单表），排在报告表之后 |

## 目录结构

```
sql_skills/
├─ README.md                        本文件（给使用者看）
├─ CHANGELOG.md                     版本记录
├─ LICENSE                          MIT
├─ .github/workflows/test.yml       CI：Ubuntu / Windows 上跑冒烟测试
├─ tests/test_demo.py               冒烟测试：跑 demo 并校验排版约定
└─ skills/sql_report_make/          技能本体（复制这个文件夹即可安装）
   ├─ SKILL.md                      技能主文档（给 AI 助手看）
   ├─ scripts/make_sql_report.py    生成器：Markdown → docx，约 900 行
   ├─ examples/demo/                可直接运行的示例
   │  ├─ config-example.json        示例配置（全部相对路径）
   │  ├─ 实验报告.md                 示例报告
   │  ├─ template.docx              空白母本（由 make_template.py 生成）
   │  ├─ make_template.py           重新生成母本
   │  └─ SQL/*.sql                  被引用进报告的脚本
   └─ reference/排版规范.md          排版规范、统一黑色的原理、踩坑记录
```

## 配置字段

| 字段 | 说明 |
| --- | --- |
| `template` | 母本 `<序号>_command.docx`：只取页面设置与样式；不存在时退回默认 A4（上下 72pt、左右 90pt） |
| `markdown` | 报告 Markdown 路径（正文在 `<form>` 块里） |
| `out` | 输出 docx 路径（目录不存在会自动创建） |
| `fonts` | `body` 正文字体（宋体）、`heading` 栏目标题字体（黑体）、`code` 代码字体（Consolas） |
| `sizes` | `body_pt` / `code_pt` / `table_pt` 字号 |
| `toc` | 目录：`enabled`（默认 `false`）、`title` |
| `space_fix` | `true` 时把全角空格 `U+3000` 换成两个半角空格 |
| `label_map` | 可选：把 `## 一、实验目的` 这类标题映射成栏目名（用 `<form>` 时不需要） |
| `highlight` | 代码块着色规则 `{正则: 颜色}`；全部填 `000000` 即为纯黑 |
| `fence_to_sql` | `{代码块序号: 文件名}`：把该序号的代码块换成 `sql_files` 里的源文件内容。**序号按全文代码块出现顺序从 0 数，包含报告表内部的代码块**；生成器会打印「命中 x/y」 |
| `sql_files` | 逻辑名 → `.sql` 路径；相对路径按 **`markdown` 所在目录**解析 |
| `summary_table` | 可选：文末再追加一张小结表 `enabled/title/columns/rows` |

`template`/`markdown`/`out` 的相对路径按**配置文件所在目录**解析，`sql_files` 的相对路径按 **markdown 所在目录**解析，
也可以全用绝对路径——仓库里所有示例都用相对路径，换机器不用改。

## 排版规范速查

| 内容 | 行距（固定值） | 字号 | 颜色 |
| --- | --- | --- | --- |
| 报告标题（`#`） | 18 磅 | 16 磅加粗居中 | 黑 |
| 一级/二级标题（`##`/`###`） | 18 磅 | 14 / 12 磅加粗 | 黑 |
| 正文、列表、引用 | 18 磅 | 12 磅，首行缩进 2 字符 | 黑 |
| 报告表栏目名 | 14 磅 | 12 磅加粗（黑体） | 黑 |
| 报告表内容、表格 | 14 磅 | 10.5 磅 | 黑 |
| 代码、命令输出 | 12 磅 | 9 磅 Consolas | 黑 |

## 常见问题

- **改了 Markdown 后代码块内容串了？** `fence_to_sql` 按全文代码块顺序编号，增删表格/代码块都会让编号整体错位，改完必须逐个核对。
- **报告里出现蓝色/红色？** 检查是否只改了正文 run；样式定义和 `theme1.xml` 也要一起改（生成器已处理）。
- **行距忽大忽小？** 必须用 `exact` 固定行距并 `snapToGrid=0`，同时把 `sectPr` 的 `docGrid` 设成 `default`。
- **代码块空行被 WPS 吃掉？** 空行用零宽空格 `\u200b` 占位（生成器已处理）。
- **docx 能直接改吗？** 不建议，docx 是生成物；改内容改 `实验报告.md`，重新生成。
- 更多踩坑记录见 [`reference/排版规范.md`](skills/sql_report_make/reference/排版规范.md) 与 [`SKILL.md`](skills/sql_report_make/SKILL.md)。

## 开发与测试

```bash
pip install python-docx
python tests/test_demo.py     # 生成 demo 并校验：黑色 / 三档行距 / 报告表结构 / 全角空格 / 可重复生成
```

测试覆盖：`fence_to_sql` 全部命中、`gridSpan=6` 的两次合并、无 `PAGEREF`/书签、`document.xml` 与实际 run 全黑、
主题配色全黑、栏目齐全、嵌套表格存在、两次生成 `document.xml` 的 sha256 相同。
CI 在 Ubuntu 与 Windows、Python 3.10 / 3.13 上跑同一个测试（见 `.github/workflows/test.yml`）。

改动生成器后请连跑两遍比对 `document.xml` 的 sha256——书签 id 之类必须用确定性序号，不要用 `hash()`。

## License

[MIT](LICENSE)
