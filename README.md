# genPptxPython

基于 `python-pptx` 的模板驱动动态表格生成库，适合包装为 HTTP 服务并接入 Dify 工作流。

## 功能

- 使用 PowerPoint 原生表格作为样式原型。
- 通过 `<foreach item=... block=...>` 将标题、说明和表格组成逻辑块。
- 自动按照 `columns`、`rows` 增减表格行列。
- 多张完整表优先排列在同一页。
- 下一张完整表放不下时整体换页；单表超过整页时才拆分。
- 续页重复表头，生成结果保持为可编辑原生表格。
- 支持长文本行高、列宽策略、宽表拆列和数字日期格式。
- 支持同一模板页的多个独立动态表格 block。
- 空表自动隐藏，并提供续页标题上下文和唯一 Shape 名称。
- 提供资源上限、模板校验和原子文件写入，便于 Web 服务调用。

## 项目结构

```text
src/genppt/
├── cli.py              # 命令行入口
├── config.py           # 策略与资源限制
├── data.py             # columns / rows 校验与规范化
├── directives.py       # foreach 和 {{ }} 解析
├── errors.py           # 公共异常
├── models.py           # 领域模型
├── pptx_utils.py       # 受控的 PPTX 克隆兼容层
├── service.py          # 生成服务与分页布局
└── table_renderer.py   # 原生表格创建和样式复制

tests/
├── fixtures/           # PPTX 测试模板
└── test_generator.py   # 回归测试

docs/
├── architecture.md
├── template-contract.md
├── operations.md
├── requirements.md
├── table-features.md
└── code-review.md
```

## 安装与运行

```bash
python -m pip install -e .
genppt tests/fixtures/test2.pptx examples/test_data.json output/test2-generated.pptx
```

不安装包时也可以运行：

```bash
PYTHONPATH=src python -m genppt.cli \
  tests/fixtures/test2.pptx \
  examples/test_data.json \
  output/test2-generated.pptx
```

## 测试

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
```

## 文档

- [需求方案](docs/requirements.md)
- [动态表格功能与详细协议](docs/table-features.md)
- [模板与数据协议](docs/template-contract.md)
- [系统架构](docs/architecture.md)
- [运行与生产集成](docs/operations.md)
- [代码审查](docs/code-review.md)
