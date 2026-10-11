# genPptxPython

模板驱动的 PowerPoint HTTP 后端，基于 FastAPI 和 python-pptx。两个业务参数：模板路径或二进制文件流、data JSON 字符串。返回生成 PPTX 的绝对本地路径。

## 安装和启动

Python 3.10+：

```bash
python -m pip install -e '.[dev]'
GENPPT_TEMPLATE_ROOTS=/srv/templates \
GENPPT_IMAGE_ROOTS=/srv/assets \
GENPPT_OUTPUT_DIR=/srv/generated \
uvicorn genppt.main:create_app --factory --host 127.0.0.1 --port 8000
```

交互式接口文档 `/docs`；健康检查 `/health`。生产鉴权、TLS、请求超时和限流由网关提供，服务默认仅监听本机。

## 两参数 Python 入口

```python
from genppt import generate_ppt

path = generate_ppt("/srv/templates/report.pptx", '{"tables": []}')
with open("/srv/templates/report.pptx", "rb") as template:
    path = generate_ppt(template, '{"tables": []}')
```

也支持 `Path`、`bytes` 和 `BytesIO`。同步可 seek 的流从开头读取，之后恢复原位置；不关闭调用方的流。非 seek 流从当前位置消费。不支持异步 read 方法，FastAPI 的 UploadFile 使用其 `.file`。

输出目录等配置不属于业务参数。需要程序化配置时：

```python
from pathlib import Path
from genppt import GeneratorConfig, PresentationService
from genppt.core.settings import Settings

service = PresentationService(Settings(
    output_dir=Path("/srv/generated"),
    generator=GeneratorConfig(allowed_image_roots=("/srv/assets",)),
))
path = service.generate(template, data_json_string)
```

每次生成使用独立 UUID 文件名，成功后原子保存，不覆盖模板。旧 `generate_presentation(template_path, data_object, output_path, config=...)` 和 CLI 继续兼容。

## HTTP 接口

`POST /api/v1/presentations`，只接受 `template`、`data` 两个字段。响应是 JSON 字符串，例如 `"/srv/generated/0123....pptx"`，不是文件下载或客户端路径。

路径模式，Content-Type 为 application/json：

```bash
curl --noproxy '*' http://127.0.0.1:8000/api/v1/presentations \
  -H 'Content-Type: application/json' \
  -d '{"template":"/srv/templates/report.pptx","data":"{\"tables\":[]}"}'
```

上传模式，Content-Type 为 multipart/form-data：

```bash
curl --noproxy '*' http://127.0.0.1:8000/api/v1/presentations \
  -F 'template=@/local/report.pptx' \
  -F 'data={"tables":[]}'
```

HTTP 路径必须在 `GENPPT_TEMPLATE_ROOTS` 内。文件上传不用服务端模板路径，上传文件名不参与输出路径。`data` 必须是 JSON 对象的字符串，不接受对象、数组或额外 output 参数。

## 模板功能

- 所有原始模板页依次独立展开，静态页保留，续页紧跟原页。
- 文本内容：`{{field}}`、行内 `<if field>...</if>`、对象数组的整段 `<foreach items=details>...</foreach>`。
- 每个插值 value 继承第一个 `{` 的完整样式，不二次解释业务数据。
- Shape 名称：`<foreach item=reports block=report>名称或选择器</foreach>`；不填 block 时独立，显式同页相同 item + block 时组成混合组。
- 表格 `{{table1}}` 选择单表或表数组；rows 固定，columns 可从第一行推导；原生表格自动调整行列，继承模板样式。
- 图片 `{{images}}` 选择单路径或路径数组；rows/cols 各自可选；x/y/xy 后缀支持中文括号、逗号。
- 尽量同页放多组，完整下一组不足时新页；`page=group` 每项新页；长表按行拆页重复表头，后方图片在表格结束后输出。
- 同母版空白续页，不重复普通标题或组标题。空/缺失成员只删除自身，全部循环为空保留原页固定元素。
- 行高、列宽、拆列、格式化、空值与溢出策略保留。

完整语法和限制见[语法与功能文档](docs/solution-and-implementation.md)。

## 项目结构

```text
src/genppt/
├── main.py              # ASGI 工厂与错误响应
├── api/                 # 路由、请求 schema、并发准入
├── core/                # 配置、部署 settings、异常
├── services/            # 两参数应用服务、JSON 校验
├── engine/              # PPTX 包、模板解析、布局、原生渲染
├── cli.py               # CLI
└── config/errors/service/generator.py  # 旧公开入口兼容转发
tests/
├── unit/                # 服务、语法、布局、样式回归
├── integration/         # HTTP 与真实模板端到端
└── fixtures/            # 模板与测试图片
examples/                # 数据及 Python/HTTP 示例
docs/                    # 语法、接口、架构、运维、审查
var/generated/           # 默认运行产物，Git 忽略
```

## 验证

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
ruff check src tests examples
ruff format --check src tests examples
PYTHONPATH=src python examples/run_multipage_demo.py
```

启动服务后，可运行 `python examples/run_http_demo.py` 验证路径与上传两种输入。示例使用 tests/fixtures/test-multipage.pptx，需要将模板和图片根目录配置到对应 fixtures 目录。代理环境本机测试需设置 `NO_PROXY=127.0.0.1,localhost`。

## 文档

- [完整语法与功能](docs/solution-and-implementation.md)
- [两参数入口和 HTTP 协议](docs/entry-routing.md)
- [架构与目录](docs/architecture.md)
- [生产运行与边界](docs/operations.md)
- [实现计划与执行记录](docs/implementation-plan.md)
- [代码审查](docs/code-review.md)
- [表格功能](docs/table-features.md)、[图片功能](docs/image-features.md)、[模板约定](docs/template-contract.md)
