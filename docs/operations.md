# 生产运行与集成（v0.4）

## 调用

Python 3.10+：

```bash
python -m pip install -e .
genppt TEMPLATE.pptx DATA.json OUTPUT.pptx
```

HTTP 适配器调用 generate_presentation，模板路径、图片目录和输出目录由服务端控制。每个任务使用独立目录和输出名称。TemplateError 映射为可读的 4xx，未知异常记录日志后返回 5xx。

## 配置

- continuation_top_margin_inches：续页内容顶部，默认 0.55。
- bottom_margin_inches：页脚预留，默认 0.55。
- block_gap_points：不同组的间距，默认 12。
- minimum_fragment_rows：长表尽量避免小碎片，默认 2；完整续页容量不足两行时仍允许一行。
- allowed_image_roots：允许的本地图片目录，生产环境应配置。
- max_tables/max_total_cells/max_output_slides：跨全文件累计。
- max_images/max_image_bytes/max_total_image_bytes/max_image_pixels：图片资源限制。
- max_loop_items/max_text_paragraphs/max_text_characters：循环及文字资源限制。
- max_template_bytes/max_template_uncompressed_bytes/max_template_parts：默认 50 MiB / 250 MiB / 10,000 个部件。

旧 allow_additional_template_slides、repeat_non_table_shapes_on_continuation 参数只保留调用兼容，不再控制页面行为。

## 并发和文件

一次请求只加载一次 PPTX。高负载使用进程池，限制并发和任务超时。原子保存可防止失败破坏已有输出，但不能解决多个请求写同一文件的业务竞态，应使用唯一输出路径。

不接收任意客户端服务器文件路径；上传目录及图片素材需要受控。HTTP 层仍应限制请求体与上传大小；库内部另有模板压缩/解压与部件上限。累计单元格上限同时约束输入数据和含重复表头的生成结果。

## 回归与示例

```bash
PYTHONPATH=src python -m unittest discover -s tests -v
PYTHONPATH=src python examples/run_multipage_demo.py
```

示例读取 tests/fixtures/test-multipage.pptx，输出到 output/multipage-demo。使用 --prepare --source PATH 可从三页源模板制作副本，原件不覆盖。数据和结果在输出目录生成。

文字高度估算不能替代渲染检查。模板、字体或依赖升级时应执行包检查、几何检查和真实渲染。
