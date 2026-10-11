# 两参数入口与 HTTP 协议（v0.5）

## 业务契约

```python
from genppt import generate_ppt

output_path: str = generate_ppt(template, data)
```

只有两个参数：

1. template：本地 str/Path，或 bytes/同步二进制流。可 seek 流从开头读并恢复位置；非 seek 流从当前位置读；不关闭 caller stream。
2. data：JSON 字符串，解码后的根必须为对象。禁止重复字段、NaN/Infinity、非法 XML 字符、超深/过大数据。

返回已有生成文件的绝对服务器路径 str。输出根目录和资源限制由 Settings 决定；每次 UUID 文件名，失败不返回路径、不留下半成品。数据格式错误在读取模板前失败。

`PresentationService(settings).generate(template, data)` 用于依赖注入。`generate_ppt` 使用环境变量创建 Settings。旧三参数 generate_presentation 保留用于指定输出的 CLI/历史库调用。

## HTTP

`POST /api/v1/presentations` 使用同一路径支持两种 content type：

| 模式 | template | data |
|---|---|---|
| application/json | 服务端路径字符串，必须在模板允许目录内 | JSON 字符串，不是 JSON 对象 |
| multipart/form-data | 单个上传文件 | 表单字符串，内容为 JSON 对象 |

严格只接收这两个字段。响应 200 为 JSON 字符串路径，例如 `"/srv/generated/UUID.pptx"`。不提供文件下载地址，不把服务器路径当作客户端可访问文件。

状态码：400 格式损坏的 multipart 或表单超限；403 模板目录越界；413 请求体/上传文件超限；415 content type 不支持；422 模板/数据/参数非法；503 并发已满（Retry-After）；500 内部存储/未知故障，响应不暴露堆栈。

输入解析和请求大小限制采用受维护的框架实现。文件使用 multipart，解析后异步路由在线程池调用同步生成服务；上传上下文在工作完成或报错后关闭，上传名称不参与输出路径。参考 [FastAPI 文件接收](https://fastapi.tiangolo.com/tutorial/request-files/) 和 [Starlette 请求限制](https://www.starlette.io/requests/)。

## 自动识别

| 模板元素 | 处理 |
|---|---|
| 普通文本/静态表格单元格的内容标签 | 根 data 插值、if、整段 foreach |
| 图片名称 {{path}} | 根 data 单图替换 |
| 文本 Shape foreach | 当前对象的实际文本内容；名称正文按字面量 |
| 原生 table foreach | 当前表项或字段选择器的对象/数组 |
| 图片 foreach | 当前路径或字段选择器的路径/数组 |
| 显式同页相同 item + block | 同项混合布局、统一分页 |
| 无动态标记 | 原页保留 |

一次加载，编译所有原页，处理静态内容，然后逐页展开、一次原子保存。缺失表格字段隐藏自身，不误取无关外层 rows。兼容旧“当前表项 + 名称字段”形式：当前项含 rows 且名称字段为标量时，读取当前表项；新模板建议明确使用 {{table1}}。

详细模板语法见[完整设计](solution-and-implementation.md)。
