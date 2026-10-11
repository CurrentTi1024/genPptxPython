# 统一入口与自动路由（v0.4）

完整契约见[语法与功能设计](solution-and-implementation.md)。

```python
from genppt import GeneratorConfig, generate_presentation

result = generate_presentation(
    template_path="/srv/templates/report.pptx",
    data=payload,
    output_path="/srv/output/report.pptx",
    config=GeneratorConfig(
        continuation_top_margin_inches=0.55,
        allowed_image_roots=("/srv/uploads",),
    ),
)
```

## 页级计划

一次加载模板，先快照全部原始页并校验语法。每页产生独立 PagePlan，包含当前页的 block、Shape 原型和原始页序号。生成续页不会再次扫描。

| 元素 | 处理 |
|---|---|
| 无名称指令的普通文本 | 根 data 的插值、if、段落 foreach |
| 普通图片名称 {{path}} | 根 data 的单图片替换 |
| 带 foreach 的文本 | 当前外层对象的文本模板 |
| 带 foreach 的 table | 当前表项或选择器指定的表格对象/数组 |
| 带 foreach 的图片 | 当前路径或选择器指定的路径/数组 |
| 无动态内容 | 保留原页元素 |

相同 item + block 仅在同页关联。不带 block 的每个 Shape 独立。同页可以混排文字、图片和多个表格。

## 执行顺序

1. 校验路径、根数据和页数。
2. 编译所有原页及其文本语法。
3. 处理未被循环接管的静态单元格、组合对象内文字和单图片。
4. 按原页顺序展开 block，渲染文本、规范化表格、校验图片。
5. 统一计算组内横排与纵向流动，生成原页内容和同母版空白续页。
6. 一次原子保存。

插值后产生的业务文本不再解析。替换值继承占位符第一个左花括号的样式。占位符从原布局继承的有效样式先实体化，防止续页改用空白布局后样式变化。

## 失败保护

语法错误、非法数据、无法容纳的单行/段落、资源超限均返回 TemplateError。临时输出和最终目标在同一目录，成功后原子替换；失败保留既有文件。

旧 allow_additional_template_slides 及 repeat_non_table_shapes_on_continuation 配置保留调用兼容，但不再改变页面行为。所有原页始终处理，组标题不在续页重复。
