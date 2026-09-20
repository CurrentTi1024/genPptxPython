# 模板与数据协议

## 模板标记

标题或说明文本框：

```text
<foreach item=tables block=reasonTable>{{title}}</foreach>
```

原生表格：

```text
<foreach item=tables block=reasonTable>{{tableName}}</foreach>
```

`item` 必须与输入 JSON 的数组路径一致。`block` 由模板作者命名；同一 `item + block` 的 shape 作为一个整体生成。不使用 `role`，系统通过 shape 类型识别文本和表格。

每个 block 必须包含一个原生表格，同一模板页可以包含多个 block。纯文本或图片的 `foreach` 不由本模块接管，可继续交给原有模板引擎。

## 数据协议

只有 `columns` 和 `rows` 是保留字段，其他字段均由用户定义：

```json
{
  "tables": [
    {
      "title": "资料不完整",
      "tableName": "missing_documents",
      "columns": ["区域", "数量", "金额"],
      "rows": [
        ["华东区", 12, 36000],
        ["华南区", 8, 19000]
      ]
    }
  ]
}
```

`rows` 也支持对象数组，此时 `columns` 的值同时作为取值键。二维数组每一行的值数量必须与 `columns` 相同；对象数组缺失字段时补空字符串。

支持的单元格值包括字符串、整数、有限浮点数、布尔值、`Decimal`、日期、日期时间和 `null`。不接受 NaN、Infinity、嵌套对象或数组。

列级格式、空值、列宽、拆列、长文本策略和续页变量详见[动态表格功能与详细协议](table-features.md)。

## 表格结构和样式

```text
目标列数 = len(columns)
目标行数 = 1 + len(rows)
```

模板行列过多时删除，过少时增加。第一行继承模板表头样式；数据行循环继承模板数据行样式；首列、中间列、末列分别继承对应位置样式。表格总宽度保持模板宽度。列宽低于最小可读宽度时自动拆列。

## 模板页

只需要一张模板页。续页复制同一页的版式和固定对象，并重新生成动态 block。模板输出路径必须与输入模板路径不同。
