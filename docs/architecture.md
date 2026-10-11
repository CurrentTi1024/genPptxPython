# 系统架构与目录（v0.5）

## 分层

| 层 | 目录/入口 | 职责 |
|---|---|---|
| HTTP | main.py、api/ | ASGI 工厂、路由、协议校验、路径访问、错误响应、并发准入 |
| 配置 | core/ | GeneratorConfig、部署 Settings、业务异常 |
| 应用 | services/presentation.py | 两参数入口、严格 JSON 校验、唯一输出路径 |
| 引擎 | engine/ | 有界模板加载、原页计划、流式布局、文字/图片/表格原生渲染和原子保存 |
| 适配 | cli.py、旧公开入口 | 保留已有调用方式，转发到同一引擎 |

引擎不依赖 FastAPI，API 不实现 PPT 排版。services 和 engine 均不 import HTTP 框架，可在后台任务/现有 Dify HTTP 适配器中复用。没有数据库需求，不创建空 repository/DAO 目录。

## 引擎职责

- package.py：路径/bytes/流的有界读取和 ZIP/PPTX 校验，恢复 caller stream 位置。
- generation.py：一次加载、统筹生成和原子保存。
- planner.py/directives.py：原页快照、语法、同页 block、区域。
- flow_renderer.py：组内 Unit/Band、整组保持、Placement、统一分页和累计资源。
- text_renderer.py：样式插值、if、段落循环、文字测量。
- data.py/table_renderer.py：表结构、格式化、列宽、行高、原生表格。
- image_renderer.py：受控图片读取、字节缓存、EXIF 与替换。
- pptx_utils.py/models.py：克隆兼容层及结构类型。

## 请求流程

```text
HTTP 准入及请求大小检查
    → 路径/上传 + data 字符串
    → PresentationService.generate(template, data)
    → JSON 校验 + UUID 输出名
    → 有界加载 + PagePlan[]
    → Band[] / Unit[] → Placement[]
    → 原生 PPTX → 原子保存 → 绝对路径字符串
```

普通文字、图片、表格和循环全部由现有模板语法识别。生成页不参与第二次扫描。所有缓存和预算在单次请求内创建，不共享可变 Presentation 或图片对象。

HTTP 生成任务在线程池执行以避免阻塞事件循环，准入上限在读取请求体前生效。CPU 高负载部署多 worker/进程并限制总体并发；不使用“异步包装”冒充可取消的 CPU 任务。

## 兼容与测试

保留 genppt 顶层公开 API、genppt.config、genppt.errors、genppt.service 和 genppt.generator 转发。旧内部模块搬到 genppt.engine；直接使用内部导入的调用方应更新导入路径。unit/ 与 integration/ 都通过 unittest discover 或 pytest 发现，fixtures 路径与测试模块位置解耦。

python-pptx 私有接口集中用于原生格式复制，升级需完整包/几何/逐页渲染回归。
