# 运行与生产集成

## Python 调用

```python
from pathlib import Path

from genppt import GeneratorConfig, generate_presentation

output = generate_presentation(
    Path("template.pptx"),
    payload,
    Path("output.pptx"),
    config=GeneratorConfig(),
)
```

服务每次调用都创建独立的 `Presentation`，没有进程内共享模板状态。输出先写入同目录临时文件，再使用原子替换，生成失败不会留下半成品输出。

## 默认资源限制

| 限制 | 默认值 |
|---|---:|
| 表格数量 | 200 |
| 单表列数 | 50 |
| 单表数据行 | 10,000 |
| 总单元格 | 200,000 |
| 输出页数 | 500 |
| 单元格字符数 | 10,000 |
| 默认最小列宽 | 0.65 英寸 |
| 默认最大行高 | 1.4 英寸 |
| 缩小策略最小字号 | 8 磅 |

HTTP 层应将 `TemplateError` 映射为 4xx 参数或模板错误；文件系统和未知异常映射为 5xx。不要把内部堆栈直接返回给调用方。

## 并发和性能

- 主要开销与总单元格数量和输出页数近似线性相关。
- 每个请求会在内存中持有一份 PPTX 对象，不建议在线程中处理超大文件。
- Web 服务建议将生成任务放入受限进程池，并设置请求体、执行时间和输出文件大小限制。
- 模板可在业务层按文件哈希缓存验证结果，但不要复用可变的 `Presentation` 实例。

## 发布前检查

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
PYTHONPATH=src python -m compileall -q src tests
git diff --check
```

对关键模板还应生成真实 PPTX，运行包结构与布局检查，并至少使用 LibreOffice 或 PowerPoint 渲染全部页面进行视觉复核。
