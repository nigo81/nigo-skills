---
文件: manifest_template.md
层级: 参考文档
用途: manifest.json 结构契约模板 + 4种报告格式样例（合体docx/分体xlsx/分体docx/PDF）
---

# manifest.json 结构契约模板

manifest.json 是四表提取的结构契约，声明所有格式差异（文件角色/列布局/sheet分段/项目名别名）。Claude 在 Step 2 理解结构后写 manifest，apply_manifest 按 manifest 纯机械抽数（零猜测、新格式改 JSON 不改代码）。

## Schema

```json
{
  "report_format": "merged_docx | split_xlsx | split_docx | split_pdf",
  "company": "编制单位",
  "period": "会计期间",
  "files": {
    "文件名": {"role": "statements|notes|opinion|cover", "loader": "xlsx|docx|pdf|doc"}
  },
  "statements_map": {
    "合并资产负债表": {
      "source": "文件名",
      "loader": "xlsx|docx|pdf",
      "sheets": ["sheet名"],            // xlsx 用，docx/pdf 留空
      "source_table": "表名",            // docx/pdf 用，按此在 extracted_tables 定位（留空时按内容匹配）
      "kind": "linear | matrix",         // linear=BS/IS/CF, matrix=权益变动表
      "item_col": 0, "current_col": 2, "prior_col": 3,
      "field_aliases": {"原项目名": "标准项目名"},
      "merge_into": null                 // 分段合并目标表名（BS主表+续表列布局不同时用）
    }
  }
}
```

## 字段说明（每个对应一个踩过的坑）

| 字段 | 解决的问题 | 示例 |
|---|---|---|
| files.role+loader | 分体式文件角色识别 | 天邑/Macko 4文件（报表/附注/正文/封面） |
| item_col/current_col/prior_col | 列布局不同 | 天邑(B/E/F) vs Macko(A/C/D)，看 dump 确定 |
| sheets+merge_into | BS续表/分段合并 | Macko母公司BS主表(col1)+续表(col0)列布局不同→分两段merge |
| source_table | docx/pdf 四表定位 | 巨东嵌套表 name 为空→按内容匹配"资产总计"+"流动资产合计"=BS |
| field_aliases | 项目名标准化 | "股东权益合计"→"所有者权益合计" |
| kind: matrix | 权益变动表 | sheets_main(本期)+sheets_prior(续表上年)+value_col |

## 4 种格式样例

### 格式1: 合体 docx（巨东，嵌套表+sdt）
```json
{
  "report_format": "merged_docx",
  "files": {"巨东2025.docx": {"role": "statements+notes", "loader": "docx"}},
  "statements_map": {
    "合并资产负债表": {"source": "巨东2025.docx", "loader": "docx", "source_table": "", "kind": "linear", "item_col": 0, "current_col": 2, "prior_col": 3, "field_aliases": {"负债和所有者权益合计": "负债和所有者权益总计"}}
  }
}
```
注意：巨东四表在嵌套表里（parse_report 已修复），extracted_tables 里 name 常为空，source_table 留空→apply_manifest 按内容匹配（含"资产总计"+"流动资产合计"=BS）。

### 格式2: 分体 xlsx（Macko，列布局A/C/D + 续表分段）
```json
{
  "report_format": "split_xlsx",
  "files": {
    "3：报表.xlsx": {"role": "statements", "loader": "xlsx"},
    "4：附注.docx": {"role": "notes", "loader": "docx"},
    "2：正文.docx": {"role": "opinion", "loader": "docx"},
    "1：cover.doc": {"role": "cover", "loader": "doc"}
  },
  "statements_map": {
    "合并资产负债表": {"source": "3：报表.xlsx", "sheets": ["1合并资产负债表 ", "2合并资产负债表(续)"], "kind": "linear", "item_col": 0, "current_col": 2, "prior_col": 3, "field_aliases": {"负债和股东权益总计": "负债和所有者权益总计", "股东权益合计": "所有者权益合计"}},
    "母公司资产负债表": {"source": "3：报表.xlsx", "sheets": ["4母公司资产负债表（续）"], "kind": "linear", "item_col": 0, "current_col": 2, "prior_col": 3},
    "母公司资产负债表_资产段": {"source": "3：报表.xlsx", "sheets": ["3母公司资产负债表"], "kind": "linear", "item_col": 1, "current_col": 3, "prior_col": 4, "merge_into": "母公司资产负债表"},
    "合并所有者权益变动表": {"source": "3：报表.xlsx", "sheets_main": ["9合并股东权益变动表"], "sheets_prior": ["10合并股东权益变动表 (续)"], "kind": "matrix", "value_col": 2, "field_aliases": {"期末余额": "本年年末余额", "期初余额": "上年年末余额"}}
  }
}
```
注意：母公司BS主表(项目在col1)和续表(col0)列布局不同→分两段，资产段 merge_into 母公司资产负债表。

### 格式3: 分体 xlsm（天邑，列布局B/E/F）
```json
{
  "report_format": "split_xlsx",
  "files": {"2：报表.xlsm": {"role": "statements", "loader": "xlsx"}, "3：附注.docx": {"role": "notes", "loader": "docx"}},
  "statements_map": {
    "合并资产负债表": {"source": "2：报表.xlsm", "sheets": ["合并资产负债表", "合并资产负债表（续）"], "kind": "linear", "item_col": 1, "current_col": 4, "prior_col": 5}
  }
}
```
注意：天邑 col B(1)=项目名、col E(4)=期末、col F(5)=期初。col A 是"应收账款净值"，col B 是标准名"应收账款"，优先用 col B。

### 格式4: PDF 报表（文本型 pdfplumber / 扫描型 mineru）
```json
{
  "report_format": "split_pdf",
  "files": {"年报.pdf": {"role": "statements+notes", "loader": "pdf"}},
  "statements_map": {
    "合并资产负债表": {"source": "年报.pdf", "loader": "pdf", "source_table": "", "kind": "linear", "item_col": 0, "current_col": 2, "prior_col": 3}
  }
}
```
注意：PDF 报表先经 parse_report.py（pdfplumber 文本型 / mineru 扫描型）提取表格进 extracted_tables.json，apply_manifest 按 source_table 定位（name 为空时按内容匹配）。PDF 的四表勾稽与 xlsx/docx 走同一套算术。

## 写 manifest 的方法

1. **xlsx**：`openpyxl load_workbook` 看 sheet 名，dump 每个四表 sheet 的表头+前3行（找"项目"列+"期末/期初"金额列），确定 item_col/current_col/prior_col
2. **docx/pdf**：看 extracted_tables.json 的四表（name 或按内容定位），dump headers+前3行确定列
3. **权益变动表**：看矩阵结构（行=项目，列=权益成分），确定 value_col（合计列）+ sheets_main/sheets_prior
4. **项目名别名**：对照 `calculator_rules.py` 的必需字段名，报表原文非标准名写进 field_aliases

## 验证 manifest

```bash
python3 scripts/apply_manifest.py <manifest.json> -o statements.json
```
检查：资产总计=负债+权益（平衡）、关键值（货币资金/应收账款/营业收入）与报表一致。
