# 动态表格功能总览与详细方案

## 功能总览

| 功能 | 状态 | 默认行为 |
|---|---|---|
| 长文本行高自适应 | 已实现 | `wrap`，根据实际列宽、字体和中英文字符估算行高 |
| 空数据处理 | 已实现 | 外层数组为空或 `rows=[]` 时不生成标题和表格 |
| 生成前布局预检 | 已实现 | 只拒绝无法靠换页解决的布局；正常内容继续自动换页 |
| 多动态表格 block | 已实现 | 同一模板页可以包含多个 `item + block` |
| block 独立动态区域 | 已实现 | 根据 block 位置和下方相交 block 自动推导区域 |
| 列宽策略 | 已实现 | `template`、`content`、`equal` |
| 最小列宽与拆列 | 已实现 | 低于最小宽度时按连续列组拆表 |
| 数字、日期格式 | 已实现 | 列级 `format`；ISO 日期字符串支持 `strftime` |
| 空值显示 | 已实现 | 表级默认值，列级可覆盖 |
| 文本溢出 | 已实现 | `wrap`、`shrink`、`truncate`、`error` |
| 续页标题 | 已实现 | 模板通过 `_page` 系统上下文控制 |
| Shape 名称唯一性 | 已实现 | 名称包含 block、数据、列分片和行分片序号 |

## 数据契约

`columns` 与 `rows` 仍是唯一必填字段。原有字符串列定义保持兼容：

```json
{
  "columns": ["区域", "数量"],
  "rows": [["华东", 12]]
}
```

需要格式化时，列可以写成对象：

```json
{
  "columns": [
    {"key": "amount", "label": "金额", "format": ",.2f", "nullValue": "-"},
    {"key": "rate", "label": "占比", "format": ".1%"},
    {"key": "createdAt", "label": "日期", "format": "%Y-%m-%d"}
  ],
  "rows": [
    {"amount": 1234.5, "rate": 0.256, "createdAt": "2026-09-20"}
  ]
}
```

- 数字格式使用 Python 标准格式说明符，例如 `,.2f`、`.1%`。
- 日期格式包含 `%` 时，支持 `date`、`datetime` 和 ISO 8601 字符串。
- 对象行缺失字段与显式 `null` 使用相同的空值策略。
- 非法格式会返回明确的模板数据错误，不会静默使用错误文本。

## 表级选项

表项可包含保留对象 `_table`：

```json
{
  "columns": ["编号", "说明"],
  "rows": [[1, "长文本"]],
  "_table": {
    "columnWidth": "content",
    "minColumnWidth": 0.7,
    "repeatLeadingColumns": 1,
    "overflow": "wrap",
    "maxRowHeight": 1.2,
    "minFontSize": 8,
    "nullValue": "-"
  }
}
```

| 选项 | 默认值 | 说明 |
|---|---:|---|
| `columnWidth` | `template` | `template` 保留模板比例；`content` 按内容权重；`equal` 等宽 |
| `minColumnWidth` | `0.65` | 最小列宽，单位英寸 |
| `repeatLeadingColumns` | `0` | 拆列后重复前 N 列，适合重复编号或名称列 |
| `overflow` | `wrap` | 长文本处理策略 |
| `maxRowHeight` | `1.4` | `shrink`、`truncate`、`error` 的触发高度，单位英寸 |
| `minFontSize` | `8` | `shrink` 最小字号，单位磅 |
| `nullValue` | 空字符串 | 表级空值显示 |

## 长文本与溢出

系统使用单元格有效宽度、模板字号、内边距和中英文字符宽度估算换行数。行高取该行所有单元格所需高度的最大值，并且不低于模板原型行高。

- `wrap`：保留字号并扩高。若一行在完整区域仍放不下，布局预检报错。
- `shrink`：缩小到 `minFontSize`；仍超出时按可用行数省略。
- `truncate`：保持字号，将超出内容替换为省略号结尾的文本。
- `error`：预计高度超过 `maxRowHeight` 时拒绝生成。

PowerPoint 没有稳定的公开排版测量 API，因此行高属于保守估算。模板字体、字号或内边距变化后应执行一次真实渲染回归。

## 空数据

- `item` 对应数组为空：删除该 block 的全部占位对象，不生成空表。
- 某个表项的 `rows=[]`：跳过这个表项，同时不生成它的标题和说明对象。
- 其他非动态企业页眉、页脚和背景保持不变。

## 多 block 与独立区域

同一模板页可以出现多个 block：

```text
<foreach item=primaryTables block=primary>{{title}}</foreach>
<foreach item=primaryTables block=primary>{{name}}</foreach>

<foreach item=secondaryTables block=secondary>{{title}}</foreach>
<foreach item=secondaryTables block=secondary>{{name}}</foreach>
```

区域推导规则：

1. block 顶部为该 block 所有对象的最上边。
2. 查找下方与该 block 水平方向相交的最近 block。
3. 找到时，以其顶部减 block 间距作为当前区域底部。
4. 找不到时，区域延伸到页面底边距。
5. 左右并排的 block 各自向下排版；实际矩形重叠会在生成前报错。

每个 block 独立分页。续页复制企业固定元素，只生成当前 block 的续页内容。

## 列宽和拆列

系统先根据 `minColumnWidth` 计算单个表格片段最多可容纳的列数。全部列可读时不拆分；否则按原顺序拆成多个列片段。

设置 `repeatLeadingColumns=1` 后，每个后续列片段都会重复第一列。重复列数量必须小于单个片段可容纳的列数，否则报错。

每个列片段仍遵守原有纵向分页规则。列片段和行片段组合后，每个原生 PowerPoint table 都有唯一名称。

## 续页上下文

标题或说明文本可以使用：

```text
{{title}}{{_page.continuationSuffix}}
```

| 字段 | 含义 |
|---|---|
| `_page.isContinuation` | 是否为该表的后续行片段或列片段 |
| `_page.continuationSuffix` | 首页为空，续页为 `（续）` |
| `_page.fragmentIndex` | 当前行分页序号，从 1 开始 |
| `_page.fragmentCount` | 当前列片段内的行分页总数 |
| `_page.columnFragmentIndex` | 当前列片段序号 |
| `_page.columnFragmentCount` | 列片段总数 |

## 内部布局预检

预检不改变正常分页行为。当前页剩余空间不足时仍整体换页；单表超过新页时仍拆行。只有以下无法通过自动换页解决的情况会提前失败：

- 表头和最少数据行在完整 block 区域也放不下。
- `wrap` 后的单行高度超过完整区域。
- 表格宽度小于最小列宽。
- 重复列占满整个列片段，无法容纳新列。
- 多个 block 的实际矩形互相覆盖。
