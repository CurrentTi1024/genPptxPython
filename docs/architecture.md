# 系统架构（v0.4）

## 分层与职责

| 模块 | 职责 |
|---|---|
| service.py | 一次加载、调用流程、原子保存 |
| planner.py | 原页快照、语法校验、同页 block 关联及区域检查 |
| directives.py | Shape foreach、图片尺寸后缀和数据路径 |
| text_renderer.py | 插值、if、段落 foreach、run 样式和文本高度 |
| flow_renderer.py | 组内成员展开、横排、整组保持、跨页布局和累计资源 |
| data.py | 简写列映射、第一行推导、格式化和输入限制 |
| table_renderer.py | 列宽、行高、拆列、原生表格样式 |
| image_renderer.py | 本地路径校验、图片快照、EXIF、缓存和替换 |
| pptx_utils.py | Shape/关系复制、占位符继承实体化、同母版空白页 |
| config.py | 布局配置和资源上限 |

图片和表格不再分别拥有分页器。flow_renderer 统一管理每组的元素顺序和分页，因此可以让长表结束后继续输出图片。

## 数据流

```text
CLI / HTTP adapter
    → generate_presentation
    → 原页 PagePlan[]
    → 每页 PageBlock[]
    → 每项 Band[] / Unit[]
    → Placement[]（页序号、位置、表格行片段或文本段落）
    → 原生 PPTX Shape
    → 原子保存
```

原页排序、逻辑组数据和渲染副作用分离。横排中的不同成员共享页序号，整排可容纳时整体换页。多 block 在原页使用独立区域，各自的续页连续插入原页之后。

所有请求状态局部创建。图片读取一次后按规范路径缓存字节快照，后续生成不重复解码或读取可能已变化的文件。共享模板不被修改。

## 私有接口边界

python-pptx 缺少公开的完整幻灯片克隆和格式继承 API。受控底层操作集中在 pptx_utils、text_renderer、table_renderer，须结合包完整性和视觉回归。项目依赖 python-pptx>=1.0,<2.0。
