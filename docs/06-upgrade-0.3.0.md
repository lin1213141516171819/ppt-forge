# PPT Forge 0.3.0

## 模板内容安全

显式 `fieldMap` 模式默认清空当前样板页上所有未绑定文本框，避免参考稿原文泄露。保留固定文字时，在页面上写：

```json
{
  "type": "content",
  "preserveText": [{"name":"CompanyLogo"}, {"name":"FooterText"}],
  "title": "新标题",
  "bullets": ["新要点"]
}
```

`preserveText` 只保留文本，不代表该对象可被内容修改。自动识别模式仍按原有规则清理样板文本。

## 文字和槽位

`content` 支持 `title`、`bullets`、`text`、`image`、`images`、`caption`、`callout`；`image-text` 还支持 `captions`。多图使用多个图片选择器：

```json
"fieldMap": {
  "content": {
    "title": {"name":"Title"},
    "text": {"name":"Body"},
    "image": {"name":"HeroImage"},
    "caption": {"name":"Caption"},
    "callout": {"name":"Conclusion"}
  }
}
```

文本缩放以文本框实际宽度和模板继承字号为依据。不能在最低字号内放下时，质量报告保留容量警告；应缩短文字、拆页或改样板。

## 成稿编辑

```json
{
  "action": "edit",
  "path": "deck.pptx",
  "output": "deck-edited.pptx",
  "operations": {
    "deletePages": [4, 7],
    "addPages": [{"after": 3, "templatePage": 5, "replace": {"Title": "新增页面"}}],
    "pageOrder": [1, 2, 3, 4, 5],
    "replace": [{"page": 2, "selector": {"name":"Body"}, "text":"替换后的正文"}],
    "syncToc": {"page": 2, "selector": {"name":"TOC"}},
    "pageNumbers": {"selector": {"name":"PageNumber"}, "format":"{page:02d}/{total:02d}"}
  }
}
```

编辑写入临时 PPTX 后原子替换目标文件。页面选择器不唯一或页码不合法时整体失败。`pageOrder` 要包含编辑后的每一页且每页只能出现一次。新增页面复制同一 deck 中的样板页，所以新页可继续使用样板的图形和版式。

## 质量和审查

质量检查按估算文字区域判断相交，`quality.ignoreOverlaps` 可使用形状名称或 ID：

```json
"quality": {"ignoreOverlaps": [{"name":"IntentionalCallout"}]}
```

视觉审查：

```json
{"action":"review","path":"deck.pptx","review":{"reviewType":"visual","reviewToken":"...","findings":[]}}
```

内容审查：

```json
{"action":"review","path":"deck.pptx","review":{"reviewType":"content","reviewToken":"...","findings":[{"page":5,"severity":"error","message":"结论缺少来源支撑"}]}}
```

`reviewToken` 必须来自当前版本的完整渲染。视觉审查检查溢出、遮挡、字体和版式；内容审查检查标题正文关系、事实支撑、离题和结论依据。插件只保存审查结论，不替代模型判断。

## 目录同步

自由排版 outline 可以设置 `syncToc: true`，在生成前用每页标题重写目录。模板目录建议使用 `edit.syncToc`，因为它能按实际成稿页序和页数同步。页码对象需要在模板中预先命名，`edit.pageNumbers` 会逐页更新其文字。
