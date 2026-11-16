# PPT Forge 0.2.0

保留 DSH 插件及 `ppt_forge` 入口，增加明确的对象绑定、生成前校验和可追踪的验收流程。未引入额外模型提供商或在线服务。

## 检查模板

```json
{"action":"inspect","path":"template.pptx"}
```

返回每页对象的 `id`、`name`、`kind`、`boundsIn`、文本和估算容量。`page` 和 `samplePages` 从 1 开始。形状 ID 只在页面内有效；重名时使用 ID，无法唯一匹配会报错。组内对象标记为 `group-local` 坐标。

## 明确绑定后生成

```json
{
  "action": "build",
  "template": "template.pptx",
  "samplePages": {"content": 3},
  "fieldMap": {
    "content": {
      "title": {"name": "Headline"},
      "bullets": [{"id": 5}, {"id": 6}]
    }
  },
  "outline": {
    "title": "项目汇报",
    "slides": [{"type":"content","title":"项目进展","bullets":["完成需求分析","完成原型","开始联调"]}]
  },
  "render": true
}
```

正文超过目标框数量时，剩余条目放进最后一个框的多个段落，报告提示估算容量风险。`outline.slides[].fieldMap` 可覆盖单页映射。

指定映射后，该页所有有内容的字段都必须有绑定。绑定但没有新内容的槽位会清空，其他对象保留，适用于 Logo、页脚和固定品牌文案。未指定映射的模板页继续自动识别并清理旧文本；正式模板建议使用显式映射。

| 页面类型 | 可映射字段 |
| --- | --- |
| cover | title、subtitle、tag、image、eyebrow、badge |
| toc | title、items |
| section | title、number、subtitle、image |
| content | title、bullets |
| two-column | title、leftTitle、leftBullets、rightTitle、rightBullets |
| table | title、table（输入仍为 headers、rows） |
| image | title、image、caption |
| image-text | title、image、caption、bullets |
| gallery | title、images、captions |
| stats | title、items（将数值、标签、说明写入文本槽位） |
| end | title、subtitle |

`table` 绑定现有表格。超出行列容量或写入被合并遮蔽的单元格时拒绝生成，不会截断数据。`image` 绑定图片或占位符，替换时保留位置、层次、名称和 ID，并居中裁剪；图片必须存在。多对象列表适用于 bullets、items、images、captions 等列表字段。

模板对象的位置由模板决定；自由排版的 `layout` 和配色选项不负责重新排列模板。

## 读取质量报告

```json
{"action":"check","path":"deck.pptx"}
```

报告与 PPT 同目录，例如 `deck.pptx` 对应 `deck.quality.json`。

| 字段 | 含义 |
| --- | --- |
| static.findings | 包结构错误、容量估算、文本框相交、文本/表格越界等，包含页码和对象 ID |
| render.status | not-requested、partial、failed 或 complete |
| render.missingFonts | 渲染器报告的缺失字体 |
| visualReview.status | pending、failed 或 passed |
| readyForDelivery | 当前文件完成渲染和视觉审查且无阻断错误时为 true |

默认 `qualityMode: "warn"` 保留静态警告以便查看预览。`qualityMode: "strict"` 遇到任何静态发现（含估算警告）即拒绝替换目标文件；这仍不代替视觉审查。坏的包内关系始终阻止发布，图片缺失导致占位符时也会记录警告。

容量受字体和段落影响，只是估算；相交文本框也可能是刻意设计。组内对象不做跨坐标空间的边界比较。缺字、颜色、图表是否易读等需实际查看预览后判断。

## 查看预览并提交审查

```json
{"action":"render","path":"deck.pptx"}
```

每次渲染产生新的 `preview/<文件名>-<内容指纹>-<随机标记>/` 目录，返回图片路径、缺失字体和 `reviewToken`。部分页面渲染不允许提交整份通过结论，需要完整渲染。旧目录保留用于对照，不与当前结果混用。

Agent 或用户实际查看每张预览后提交发现：

```json
{
  "action": "review",
  "path": "deck.pptx",
  "review": {
    "reviewToken": "最新完整渲染返回的值",
    "reviewer": "host-agent",
    "findings": [{"page":2,"severity":"error","message":"标题被裁切，需要缩短"}]
  }
}
```

`error` 阻止通过，`warning` 表示已审查但需留意。只有实际检查后未发现问题才用 `findings: []`。发现错误后修改 outline 或映射，重新生成、完整渲染并再次审查。更换 PPT 或修改预览后旧 token 不再适用。审查记录是具名结论，不会自动调用外部视觉模型。

渲染依赖 DSH 的 LibreOffice Kit；不可用时返回 `previewError`，保留未通过状态。

## 开发验证

```powershell
python -m pip install -r python/requirements.txt
npm test
npm run test:integration
```

真实 PPTX 测试覆盖对象清单、文本/表格/图片替换、背景与备注、11 种自由版式及失败时保留旧文件。`test:integration` 缺少依赖会明确失败。附带 Windows/Linux CI 配置，安装依赖后运行全部测试；本地跳过的测试不能视为通过。

设计参考：[pptx-template](https://github.com/m3dev/pptx-template) 的字段绑定、[pptx-automizer](https://github.com/singerla/pptx-automizer) 的对象选择和 [PPTAgent](https://github.com/icip-cas/PPTAgent) 的渲染审查流程。本次代码为本项目的独立实现。
