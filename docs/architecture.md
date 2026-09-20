# 系统架构

## 分层

```text
CLI / HTTP Adapter
        ↓
Generation Service
        ├── Template plan and automatic routing
        ├── Template directives
        ├── Data normalization
        ├── Flow pagination
        ├── Image validation and grid layout
        └── Atomic output
                ↓
PPTX compatibility layer
        ├── Slide and shape cloning
        ├── Relationship remapping
        └── Native table style copying
```

`generate_presentation()` 是服务层入口。函数无共享可变状态，每次调用独立加载模板，适合 Web 请求并发执行。对于 CPU 和内存隔离要求较高的部署，建议通过进程池执行生成任务。

## 模块职责

| 模块 | 职责 |
|---|---|
| `config.py` | 页面策略和防止资源耗尽的硬限制 |
| `directives.py` | 严格解析模板指令，拒绝未知属性和重复属性 |
| `planner.py` | 扫描模板并区分普通文字、图片循环和动态表格 |
| `text_renderer.py` | 保留 run 样式的普通文字与静态表格单元格插值 |
| `data.py` | 校验 `columns`、`rows`，转换安全的显示文本 |
| `pptx_utils.py` | 集中管理 `python-pptx` 私有接口，避免业务层散落 XML 操作 |
| `table_renderer.py` | 行高测量、列宽分配、拆列、溢出处理和原生表格样式复制 |
| `image_renderer.py` | 图片路径校验、比例替换、图片循环、逻辑组网格和分页 |
| `service.py` | 多 block 区域规划、整表保持、分页、续页复制和原子保存 |

统一入口只加载一次 PPTX，按照“模板规划、普通静态文字、图片、表格、原子保存”的顺序执行。详细规则见[统一生成入口与自动路由](entry-routing.md)。

## 分页决策

```text
完整表格适合当前剩余空间
    → 当前页

当前页放不下，但完整新页放得下
    → 整体移动到新页

表格自身超过完整新页
    → 新页开始，按数据行拆分并重复表头
```

分页前先执行不可行布局检查；可以通过换页解决的空间不足不属于错误。

## 私有 API 边界

`python-pptx` 没有公开的幻灯片复制、完整表格样式复制和动态行列克隆 API。项目只在 `pptx_utils.py` 与 `table_renderer.py` 中使用受控私有接口，并通过包完整性、布局检查和真实渲染回归降低升级风险。依赖版本限制在 `python-pptx>=1.0,<2.0`。
