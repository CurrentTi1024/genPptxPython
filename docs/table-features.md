# 表格功能（v0.4）

## 取值与数据

独立原生 table 名称：

```text
<foreach item=reportTables></foreach>
```

reportTables 可以是单表对象或表对象数组。组内表格选择器：

```text
<foreach item=reports block=report>{{table1}}</foreach>
<foreach item=reports block=report>{{table2}}</foreach>
```

table1/table2 均可为单表或表数组。字段缺失或 rows 为空只隐藏该表格，组内其他元素仍保留。

```json
{
  "columns": [{"loc": "区域"}, {"count": "数量"}],
  "rows": [{"loc": "华中", "count": 12}]
}
```

columns 省略时只从第一条对象行推导字段和顺序，不扫描后续行。仍兼容字符串列和矩阵行；矩阵行必须提供 columns。只使用 rows，不支持 raws。列映射只有单键简写，不支持 key/label 元数据形式。

## 格式与列宽

```json
{
  "columns": [{"amount": "金额"}, {"day": "日期"}, {"note": "备注"}],
  "rows": [{"amount": 1234.5, "day": "2026-10-12", "note": null}],
  "_table": {
    "formats": {"amount": ",.2f", "day": "%Y/%m/%d"},
    "nullValues": {"note": "-"},
    "nullValue": "",
    "columnWidth": "content",
    "minColumnWidth": 0.65,
    "repeatLeadingColumns": 0,
    "overflow": "wrap",
    "maxRowHeight": 1.4,
    "minFontSize": 8
  }
}
```

数字格式使用 Python 格式说明符，日期支持 ISO 字符串和 date/datetime。非法格式、非有限数字和过大的格式宽度拒绝生成。

columnWidth 支持 template/content/equal。宽表按最小列宽拆成连续列片段，可重复前导列。表头、数据行、交替行、段落及 run 样式继承模板。

## 文本与分页

wrap 增高；shrink 缩小到最小字号后必要时省略；truncate 省略；error 超出最大行高时报错。高度按字体和有效宽度估算。

下一完整组放不下时新页开始。超长组内的表格按行拆页、重复表头，表格结束后再输出下方图片。组标题只生成一次，无自动续字。续页从统一顶部边距开始，母版不变。

模板合并单元格不支持。表头加一条数据行、或单段落在完整续页仍放不下时明确报错。

旧 _page 续页标题语法不再自动提供，因为组标题不再复制。需要文字“（续）”时使用固定文本，但它只在原组标题出现一次。
