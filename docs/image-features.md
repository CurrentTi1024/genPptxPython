# 图片功能（v0.4）

## 普通图片

名称为 {{imagePath}}，可附加 (x)、(y)、(x,y)。默认 x。中文括号和逗号均支持，如 （x，y）。

- x 固定宽度，按图片比例计算高度。
- y 固定高度，按图片比例计算宽度。
- x,y 固定原型长宽，允许拉伸。
- 空路径、缺失字段、坏图或超出单图片限制：删除该图片。
- 保留原 Shape 位置、边框、阴影等，清除旧裁剪。

## 独立循环

```text
<foreach item=images rows=2 cols=3></foreach>（x，y）
<foreach item=devices>{{imagePath}}</foreach>(x)
```

images 可以是单路径或路径数组。devices 为对象数组。rows、cols 可独立省略，系统自动计算缺失维度。超过容量自动续页。图片网格缩小后使用统一缩放比例。

## 混合组与组内数组

```text
<foreach item=reports block=report>标题</foreach>
<foreach item=reports block=report>{{table1}}</foreach>
<foreach item=reports block=report cols=2>{{images}}</foreach>（x，y）
```

每组先完成自己的标题、表格和图片，然后处理下一组。images 支持单路径或路径数组。组内坏图只隐藏对应图片，其他成员保留。横向并排原型保持相对位置，整排放不下则移到续页。

## 续页和安全

每张原页独立扩展。续页只复用同一母版的空白布局，原页普通标题和介绍不复制。支持多个模板页及同页多个独立 block。

循环原型不支持旋转，普通图片可以旋转。本地路径必须位于 allowed_image_roots（配置后）。规范路径缓存已验证的字节，保证本次请求的图片一致。EXIF 方向会实际应用，不受 python-pptx 对 EXIF 支持差异影响；非原生格式如 WEBP 转为 PNG。

默认上限：1,000 张、单图片 25 MiB、唯一图片总量 250 MiB、40,000,000 像素、500 页。资源计数跨所有模板页累计。
