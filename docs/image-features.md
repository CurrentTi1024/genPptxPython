# 动态图片替换与循环生成方案

## 功能总览

| 能力 | 模板写法 | 行为 |
|---|---|---|
| 单图片替换 | `{{imagePath}}` | 默认保持占位图宽度，按原图比例计算高度 |
| 宽度锁定 | `{{imagePath}}(x)` | 宽度不变，高度按原图比例缩放 |
| 高度锁定 | `{{imagePath}}(y)` | 高度不变，宽度按原图比例缩放 |
| 宽高锁定 | `{{imagePath}}(x,y)` | 使用占位图宽高，原图可能发生拉伸 |
| 图片数组循环 | `<foreach item=imgs></foreach>(x)` | 数组元素本身是图片路径 |
| 对象数组循环 | `<foreach item=imgs>{{imagePath}}</foreach>(x)` | 从每个对象读取图片路径 |
| 逻辑组循环 | 相同 `item + block` | 图片、文字和其他 Shape 作为一个单元循环 |
| 自动网格 | 不写 `rows`、`cols` | 根据原型尺寸和页面剩余区域计算行列 |
| 指定网格 | 可选 `rows`、`cols` | 可分别指定；必要时等比缩小逻辑组 |
| 自动分页 | 数据超过单页容量 | 复制企业固定元素并继续填充 |
| 空值和坏图 | 空路径、文件不存在、图片损坏 | 删除普通占位图；循环组跳过整个数据项 |

## 普通图片替换

将 PowerPoint 图片 Shape 的名称设置为：

```text
{{eqp_img_path}}
{{eqp_img_path}}(x)
{{eqp_img_path}}(y)
{{eqp_img_path}}(x,y)
```

没有后缀时等价于 `(x)`。引擎只替换图片关系和几何尺寸，保留原 Shape 的位置、旋转、阴影、边框等 PowerPoint 样式，并清除旧图片裁剪。

当变量不存在、值为空、路径不是文件、文件超过限制、不是可解码图片或图片损坏时，直接删除占位图片，不保留模板示例图。

## 图片数组循环

最简单的数据是路径字符串数组：

```json
{
  "imgs": [
    "/data/images/a.png",
    "/data/images/b.jpg"
  ]
}
```

图片 Shape 名称：

```text
<foreach item=imgs></foreach>(x)
```

空 body 表示数组元素本身就是路径。也支持对象数组：

```json
{
  "imgs": [
    {"imagePath": "/data/images/a.png"},
    {"imagePath": "/data/images/b.jpg"}
  ]
}
```

```text
<foreach item=imgs>{{imagePath}}</foreach>(x)
```

## 图片与其他元素的逻辑组循环

数据：

```json
{
  "imageGroups": [
    {"title": "设备 A", "imagePath": "/data/images/a.png"},
    {"title": "设备 B", "imagePath": "/data/images/b.png"}
  ]
}
```

标题 Shape 名称：

```text
<foreach item=imageGroups block=equipment>{{title}}</foreach>
```

图片 Shape 名称：

```text
<foreach item=imageGroups block=equipment>{{imagePath}}</foreach>(x,y)
```

相同 `item + block` 的所有 Shape 构成一个逻辑组，不要求在 PowerPoint 中执行物理组合。文字内容中的 `{{title}}` 等表达式按照当前数组对象替换。

一个逻辑组可包含多个图片。任一必需图片为空或不可读取时，跳过整个数据项，避免生成孤立标题或说明。

## rows 与 cols

`rows`、`cols` 均为可选正整数，也可以只指定其中一个：

```text
<foreach item=imgs rows=2></foreach>(x)
<foreach item=imgs cols=3></foreach>(x)
<foreach item=imgs rows=2 cols=3></foreach>(x)
```

| 配置 | 排版规则 |
|---|---|
| 都不设置 | 根据原型逻辑组宽高和页面剩余空间自动计算 |
| 只设置 `cols` | 固定列数，必要时等比缩小，再自动计算行数 |
| 只设置 `rows` | 固定行数，必要时等比缩小，再自动计算列数 |
| 同时设置 | 使用指定网格，必要时等比缩小整个逻辑组 |

同行同列使用统一单元格尺寸，按照先行后列的顺序填充。逻辑组相对位置和比例保持不变；缩小时，显式设置字号的文本也同比缩小。

同一 block 的多个 Shape 可以只在其中一个 Shape 上填写 `rows` 或 `cols`。如果多个 Shape 都填写，值必须一致。

## 分页和模板元素

每个图片循环模板页当前支持一个图片循环 block。数据超过 `rows × cols` 或自动网格容量时，引擎复制该页的固定 Shape、页眉、页脚和背景关系，然后继续填充剩余数据。

循环原型当前不支持旋转 Shape。旋转后的边界框会影响网格推导，因此系统在生成前明确报错；普通非循环图片替换可以保留原 Shape 旋转。

一个 PPTX 可以包含多个图片循环模板页。例如 `test-img.pptx` 第一页用于图片数组循环，第二页用于图片和标题的逻辑组循环；每一页独立扩展，生成的续页紧跟其模板页。

## 安全与资源限制

默认限制：

| 限制 | 默认值 |
|---|---:|
| 图片数量 | 1,000 |
| 单图片文件大小 | 25 MB |
| 全部唯一图片总大小 | 250 MB |
| 单图片解码像素 | 40,000,000 |
| 输出幻灯片 | 500 |

仅支持服务器本地、Pillow 可解码的栅格图片，不下载 HTTP/HTTPS URL。

HTTP 服务应配置 `allowed_image_roots`，只允许读取任务上传目录或受控素材目录：

```python
GeneratorConfig(
    allowed_image_roots=("/srv/genppt/uploads", "/srv/genppt/assets"),
)
```

路径会先解析符号链接再检查根目录，从而避免通过相对路径或符号链接越权读取其他服务器文件。

## Shape 名称

循环生成的 Shape 名称包含 block、原始数据项序号和组内 Shape 序号：

```text
设备A__equipment_i1_s1
image__equipment_i1_s2
```

即使无效数据项被跳过，名称仍使用原始数组序号，因此在日志和源数据之间可以稳定定位。

## `test-img.pptx` 验证范围

- 第一页：路径字符串数组、自动网格、默认 `(x)`、空路径和不存在文件跳过、自动分页。
- 第二页：图片和标题逻辑组、`(x,y)`、无效图片跳过整个组、模板文字替换。
- 独立测试模板：普通图片的默认、`(x)`、`(y)`、`(x,y)`；只设置 rows、只设置 cols、同时设置；资源限制和路径白名单。
