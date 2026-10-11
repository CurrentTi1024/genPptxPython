# genPptxPython

基于 `python-pptx` 的模板驱动 PowerPoint 生成服务。调用方只需提供 PPTX 模板本地路径和数据对象，生成器会自动识别普通文字、图片、图片循环和动态表格，并尽量保留模板中的企业样式。

## 已实现能力

- 普通文字与静态表格单元格支持插值、行内 if 和整段 foreach；替换值保留第一个 `{` 的 run 样式。
- 单图片替换，支持默认、`(x)`、`(y)`、`(x,y)` 四种尺寸模式。
- 图片数组循环、图片与文字逻辑组循环，可选 `rows`、`cols`，自动网格和分页。
- 原生 PowerPoint 表格循环，按照 `columns`、`rows` 自动增删行列并继承模板样式。
- 多张短表优先排在同一页，下一张完整表放不下时整体换页，长表按行分页。
- 长文本行高、列宽策略、宽表拆列、数字日期格式、空值和溢出策略。
- 多个动态 block、独立动态区域、同母版空白续页和唯一 Shape 名称。
- 所有原始模板页按页序独立展开，支持文字、多个表格和图片的混合组；续页使用同母版空白布局。
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
<foreach item=tables block=reasonTable>分类标题</foreach>
<foreach item=tables block=reasonTable>{{table1}}</foreach>
```

第二个 Shape 必须是原生 PowerPoint table；第一个 Shape 的实际文本写 `分类：{{title}}`。表格字段只有 `rows` 必需，`columns` 可省略并从第一行推导：

```json
{
  "tables": [
    {
      "title": "资料不完整",
      "table1": {
        "columns": [{"loc": "区域"}, {"count": "数量"}],
        "rows": [{"loc": "华东", "count": 12}]
      }
    }
  ]
}
```

图片逻辑组同样通过相同的 `item + block` 关联；`rows`、`cols` 均为可选字段：

```text
<foreach item=imageGroups block=equipment>设备标题</foreach>
<foreach item=imageGroups block=equipment cols=3>{{images}}</foreach>（x，y）
```

完整语法、数据协议、分页行为和生产约束见[完整方案与实现](docs/solution-and-implementation.md)。

## 自动路由与限制

| 模板元素 | 自动处理器 |
|---|---|
| 文本或静态表格单元格中的文本语法 | 插值、if、段落 foreach |
| 图片 Shape 名称中的 `{{...}}` | 单图片替换 |
| Shape foreach | 按类型展开文本、表格或图片 |
| 相同 item + block | 同页混合组，统一布局和分页 |
| 没有动态标记 | 原样复制 |

同页支持混合组和多个独立循环。文本内容中的 foreach 必须包住整段，内部数组只能是对象数组。Shape 名称中的文本内容不解析表达式。使用 `page=group` 可以切换每组新起一页。循环缺失时保留原页，只删除动态元素。

## 项目结构

```text
src/genppt/
├── cli.py              # CLI 入口
├── config.py           # 策略与资源限制
├── data.py             # 表格数据校验与格式化
├── directives.py       # 模板指令解析
├── flow_renderer.py    # 组内流式布局和统一分页
├── image_renderer.py   # 图片验证、缓存与替换
├── planner.py          # 模板扫描与自动路由
├── pptx_utils.py       # PPTX 克隆兼容层
├── service.py          # 统一生成入口与原子输出
├── table_renderer.py   # 表格布局与样式复制
└── text_renderer.py    # 带样式的插值、if 与段落循环

tests/                  # 自动化测试和 PPTX fixture
examples/               # 示例数据
docs/                   # 方案、协议、运维和审查文档
```

## 测试

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
```

自动化测试覆盖统一入口、多模板页、混合组、文字语法、表格、图片、分页、失败保护和累计资源限制。端到端示例：

```bash
PYTHONPATH=src python examples/run_multipage_demo.py
```

结果生成到 `output/multipage-demo/generated.pptx`。用户原始模板不被修改。

## 文档

- [完整方案与实现](docs/solution-and-implementation.md)
- [实现计划与执行记录](docs/implementation-plan.md)
- [统一入口与自动路由](docs/entry-routing.md)
- [动态表格功能](docs/table-features.md)
- [动态图片功能](docs/image-features.md)
- [模板协议](docs/template-contract.md)
- [系统架构](docs/architecture.md)
- [生产运行](docs/operations.md)
- [代码审查](docs/code-review.md)
