# HTTP 生产运行与边界（v0.5）

## 启动

```bash
python -m pip install -e .
GENPPT_TEMPLATE_ROOTS=/srv/templates \
GENPPT_IMAGE_ROOTS=/srv/assets \
GENPPT_OUTPUT_DIR=/srv/generated \
uvicorn genppt.main:create_app --factory --host 127.0.0.1 --port 8000
```

API `/api/v1/presentations`，文档 `/docs`，健康检查 `/health`。CLI genppt TEMPLATE DATA.json OUTPUT 继续可用。

## 部署配置

| 环境变量 | 默认 | 说明 |
|---|---|---|
| GENPPT_OUTPUT_DIR | var/generated | 绝对化后的输出根 |
| GENPPT_TEMPLATE_ROOTS | templates | HTTP 路径允许目录，按 os.pathsep 分隔 |
| GENPPT_IMAGE_ROOTS | assets | 图片允许目录，按 os.pathsep 分隔 |
| GENPPT_MAX_DATA_BYTES | 20971520 | data JSON UTF-8 字节上限 |
| GENPPT_MAX_CONCURRENT_JOBS | 2 | 每个 worker 的同时请求数，含上传过程 |

模板根为空时拒绝所有 HTTP 路径输入，上传仍可用。图片根为空时部署配置立即报错，避免静默解除访问限制；底层空 allowed_image_roots 只用于显式可信库调用。

程序化 Settings 可配置 JSON 深度（64）、节点数（500,000）、GeneratorConfig 的页数（500）、模板大小（压缩 50 MiB/解压 250 MiB/10,000 部件）、表格/图片/文字等累计资源上限。整体请求体上限为 max(模板上限 + data 上限, 6 × data 上限) + 1 MiB，用于兼容 JSON 字符转义与 multipart 边界；模板和 data 仍分别限制。

布局参数保留续页上下边距、组间距、小碎片策略、最小列宽和溢出配置。旧两个页面行为布尔参数只兼容调用，不控制新布局。

## 安全、并发和存储

- 模板/素材目录应只读且受控；路径解析防止普通遍历和指向允许目录外的符号链接。不允许不可信本地进程同时改写这些目录，文件系统权限仍是安全边界。
- 设置网关鉴权、TLS、速率和连接/上传超时。服务自身不实现用户认证，不应直接暴露互联网。
- 并发满立即 503，无无界任务队列；总 worker 并发是 worker 数 × 上限。线程池用于不阻塞 HTTP loop，CPU 重任务用进程级部署。
- 线程运行中的同步任务不能硬取消。需要强制 CPU/内存/时长隔离时，使用进程 worker 或容器配额；网关超时不等于后台任务停止。
- UUID 输出文件属于调用方，需要自行下载/归档和按业务保留期清理。项目不添加定时删除，以免误删交付物。上传临时流与失败的保存临时文件自动关闭/清理。
- 原子保存与唯一名称防止失败损坏文件及不同请求互相覆盖；磁盘满、输出目录权限等返回通用 500，服务端记录堆栈。
- 锁定已验证的 FastAPI/Starlette 小版本范围；升级需运行全部测试，特别是请求体限制、multipart 失败清理和并发准入。

## 验证

安装 .[dev] 后运行 unittest discover、Ruff 和真实模板示例。启动 HTTP 服务后运行 examples/run_http_demo.py，先请求路径再上传模板，两者返回路径；代理环境设置 NO_PROXY。

发布 wheel 使用干净 checkout/build 目录。setuptools 可能保留迁移前 build/lib 中的旧文件，应检查包内 api/core/services/engine 模块，并在源码目录之外安装、验证 ASGI 工厂和两参数生成入口。本轮已执行干净 wheel 构建与安装回归。

模板、字体和 python-pptx 升级后需包结构、几何与逐页渲染验证。字体测量为估算，本项目没有证明所有 PowerPoint 客户端像素一致。
