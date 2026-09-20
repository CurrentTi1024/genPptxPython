# genPptxPython

基于 `python-pptx` 的模板驱动 PowerPoint 生成服务。调用方只需提供 PPTX 模板本地路径和数据对象，生成器会自动识别普通文字、图片、图片循环和动态表格，并尽量保留模板中的企业样式。

## 已实现能力

- 普通文字与静态表格单元格中的 `{{field}}` 插值，支持点路径并保留 run 样式。
- 单图片替换，支持默认、`(x)`、`(y)`、`(x,y)` 四种尺寸模式。
- 图片数组循环、图片与文字逻辑组循环，可选 `rows`、`cols`，自动网格和分页。
- 原生 PowerPoint 表格循环，按照 `columns`、`rows` 自动增删行列并继承模板样式。
- 多张短表优先排在同一页，下一张完整表放不下时整体换页，长表按行分页。
- 长文本行高、列宽策略、宽表拆列、数字日期格式、空值和溢出策略。
- 多个动态表格 block、独立动态区域、续页标题和唯一 Shape 名称。
- 空数组、空表和无效图片自动删除对应占位元素。
- 模板预检、输入规模限制、本地图片访问限制及原子输出。

## 安装

```bash
python -m pip install -e .
```

要求 Python 3.10 及以上。主要依赖为 `python-pptx>=1.0,<2.0` 和 `Pillow>=10,<13`。

## Python 入口

```python
from genppt import GeneratorConfig, generate_presentation

result = generate_presentation(
    template_path="/srv/templates/report.pptx",
    data=payload,
    output_path="/srv/output/report.pptx",
    config=GeneratorConfig(
        allowed_image_roots=("/srv/uploads", "/srv/assets"),
    ),
)
```

统一入口只加载一次 PPTX。调用方无需声明模板中有哪些元素，生成器会根据 Shape 类型、Shape 名称和 foreach block 自动路由。

## CLI

```bash
genppt TEMPLATE.pptx DATA.json OUTPUT.pptx
```

未安装包时：

```bash
PYTHONPATH=src python -m genppt.cli \
  tests/fixtures/test2.pptx \
  examples/test_data.json \
  output/test2-generated.pptx
```

图片示例：

```bash
PYTHONPATH=src python -m genppt.cli \
  tests/fixtures/test-img.pptx \
  examples/image_data.json \
  output/test-img-generated.pptx
```

## 最小模板协议

普通文字写在 Shape 文本中：

```text
报告名称：{{report.title}}
```

单图片写在图片 Shape 名称中：

```text
{{imagePath}}(x)
```

动态表格标题和表格 Shape 使用相同的 `item + block`：

```text
<foreach item=tables block=reasonTable>{{title}}</foreach>
<foreach item=tables block=reasonTable>{{name}}</foreach>
```

第二个 Shape 必须是原生 PowerPoint table。每个表项只有 `columns` 和 `rows` 是固定字段：

```json
{
  "tables": [
    {
      "title": "资料不完整",
      "columns": ["区域", "数量"],
      "rows": [["华东", 12]]
    }
  ]
}
```

图片逻辑组同样通过相同的 `item + block` 关联；`rows`、`cols` 均为可选字段：

```text
<foreach item=imageGroups block=equipment cols=3>{{title}}</foreach>
<foreach item=imageGroups block=equipment>{{imagePath}}</foreach>(x,y)
```

完整语法、数据协议、分页行为和生产约束见[完整方案与实现](docs/solution-and-implementation.md)。

## 自动路由与限制

| 模板元素 | 自动处理器 |
|---|---|
| 文本或静态表格单元格中的 `{{...}}` | 普通文字插值 |
| 图片 Shape 名称中的 `{{...}}` | 单图片替换 |
| foreach block 包含图片且不包含表格 | 图片循环 |
| foreach block 包含原生表格 | 动态表格 |
| 没有动态标记 | 原样复制 |

同一模板页不能同时包含图片循环和表格循环，因为两者都有独立分页器；应放在不同模板页。普通单图片可以与动态表格同页。历史项目的 `if` 和纯文本 foreach 处理器未包含在当前仓库中，入口会明确拒绝未注册语法。

## 项目结构

```text
src/genppt/
├── cli.py              # CLI 入口
├── config.py           # 策略与资源限制
├── data.py             # 表格数据校验与格式化
├── directives.py       # 模板指令解析
├── image_renderer.py   # 图片替换、网格与分页
├── planner.py          # 模板扫描与自动路由
├── pptx_utils.py       # PPTX 克隆兼容层
├── service.py          # 统一生成入口与表格分页
├── table_renderer.py   # 表格布局与样式复制
└── text_renderer.py    # 普通文字插值

tests/                  # 自动化测试和 PPTX fixture
examples/               # 示例数据
docs/                   # 方案、协议、运维和审查文档
```

## 测试

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
```

当前共有 55 项自动化测试，覆盖统一入口、普通文字、表格、图片、循环、分页、失败保护和资源限制。

## 文档

- [完整方案与实现](docs/solution-and-implementation.md)
- [统一入口与自动路由](docs/entry-routing.md)
- [动态表格功能](docs/table-features.md)
- [动态图片功能](docs/image-features.md)
- [模板协议](docs/template-contract.md)
- [系统架构](docs/architecture.md)
- [生产运行](docs/operations.md)
- [代码审查](docs/code-review.md)
