# PPT 工坊（@local/ppt-forge）

DSH Web GUI 插件：读取现有 `.pptx` 的文字与结构，学习样本风格，按模板或内置版式生成新的 16:9 演示文稿。

## 安装（作为 DSH 插件）

前置条件：DSH 桌面版（带 `dsh` 命令）与 Python 3；DSH 自带运行时已包含 `python-pptx` 1.0.2。

**方式 A：从 Release 安装（推荐）**

1. 从 [Releases](https://github.com/lin1213141516171819/ppt-forge/releases) 下载 `ppt-forge-0.1.3.zip` 并解压；
2. 把解压出的 `ppt-forge` 目录加入你的 profile：

   ```powershell
   dsh plugin --profile <profile> add <解压目录>\ppt-forge
   ```

3. 重启 DSH 桌面应用，插件加载后即可在对话中使用 `ppt_forge` 工具。

**方式 B：直接从仓库安装**

```powershell
dsh plugin --profile <profile> add https://github.com/lin1213141516171819/ppt-forge
```

要求 pnpm 能从 git 地址安装；失败时改用方式 A。插件配置见文末「配置」一节。

## 0.1.3 修复说明

- 模板模式保留全部正文要点和目录项；超过文本框数量的条目合并为最后一个框内的多个段落。请渲染检查长文本是否溢出。
- 封面、章节和结束页会更新副标题及章节编号，清理未使用的旧文案。新页备注独立写入，不沿用样本备注。
- 复制模板页面时保留页面背景，并重映射背景图片关系。
- 模板模式支持 `cover`、`toc`、`section`、`content`、`end`。其他类型（如 `table`、`two-column`、`stats`）会明确报错；需要这些版式时请省略 `template` / `templateName` 使用自由排版。
- 模板文件不存在、没有页面或缺少所需文本框时明确报错；请修正路径或通过 `samplePages` 选择合适页面。
- 历史接口支持默认项目子目录，并兼容旧版直接输出的 `.pptx` 文件。
- 图片裁剪缓存按图片内容更新，修复同路径替换图片后仍使用旧图的问题。
- 修复 Python 命令行 `--theme teal` 参数解析。

开发验证：

```powershell
python -m pip install -r python/requirements.txt
npm test
```

`npm test` 包含 JavaScript 历史接口测试、Python 函数回归测试以及真实 PPTX 集成测试。缺少 `python-pptx` 时，PPTX 集成测试会标记为跳过；这不代表完整生成与渲染验收通过。渲染仍需 DSH 的 LibreOffice Kit。

DSH Web GUI 插件：读取现有 `.pptx` 的文字与结构，学习样本风格，按模板或内置版式生成新的 16:9 演示文稿。

## 组成

| 部分 | 说明 |
| --- | --- |
| `ppt_forge` 工具 | 对话中由 agent 调用，是主要入口 |
| `client.js` | 早期设置面板源码；0.1.2 起 `dsh.client` 声明已移除，不再加载（无独立设置页） |
| `python/extract.py` | 解析 pptx：每页标题、段落（含层级）、表格、演讲备注、图片数量 |
| `python/build.py` | 按结构化大纲渲染新 pptx（自由版式） |
| `python/template_engine.py` | 复制模板页面并改写文字（模板模式） |
| `python/learn.py` | 批量分析样本风格 |
| `python/extract_assets.py` | 提取背景/装饰素材 |

## 工具用法

```
ppt_forge { action: "extract", path: "<源文件.pptx>" }
→ 返回每页结构化内容 JSON

ppt_forge { action: "learn", path: "<样本文件夹>", name: "<档案名>" }
→ 批量分析样本结构/版式/语言/配色，生成 <工作目录>/ppt-forge-library/<档案名>.json + .md

ppt_forge { action: "assets", path: "<样本文件夹>", name: "<素材名>" }
→ 提取可复用的背景图与装饰素材到 ppt-forge-library/<素材名>/assets/，附 manifest.json

ppt_forge { action: "template-add", path: "<模板.pptx>", name: "<模板名>" }
→ 导入模板到 ppt-forge-library/templates/（registry.json 记录）

ppt_forge { action: "template-list" }
→ 列出模板库

ppt_forge { action: "build", outline: { ... }, template | templateName, samplePages, theme, style, render, output }
→ 生成 pptx；不传 output 时输出项目文件夹 <工作目录>/ppt-forge-output/<标题>-<时间戳>/deck.pptx

ppt_forge { action: "render", path: "<deck.pptx>", pages: "1,3" }
→ 渲染为 <项目目录>/preview/*.png，返回图片清单与缺失字体
```

## 模板模式

传入 `template`（路径）或 `templateName`（库内名称）即启用。生成时会：

1. 分析模板页面类型（封面/目录/章节/内容/结尾）与文字块；
2. 按大纲页面类型复制对应样板页（保留全部图片与装饰）；
3. 改写标题与正文文字（保留原格式，超长时按比例缩字号）；
4. 删除模板原有页面，只保留生成结果。

`samplePages` 可用 1-based 页码覆盖样板页，例如 `{cover: 1, section: 4, content: 6, end: 16}`。自动分类是启发式的，复杂模板建议手动指定。

## 版式

自由版式（不传模板时）支持这些幻灯片类型：

| 类型 | 字段 |
| --- | --- |
| `cover` | title, subtitle, tag, image |
| `toc` | title, items[] |
| `section` | number, title, subtitle, image |
| `content` | title, bullets[], notes |
| `two-column` | title, leftTitle, leftBullets, rightTitle, rightBullets |
| `table` | title, headers[], rows[][] |
| `image` | title, image, caption |
| `image-text` | title, image, caption, layout(left/right), bullets[] |
| `gallery` | title, images[], captions[]（1–4 张自动网格） |
| `stats` | title, items[{value,label,note}]（数据卡片） |
| `end` | title, subtitle |

- 顶层 `background`：整本 PPT 的底纹图（从素材库选，或自己的图）。
- 图片路径支持绝对路径或相对工作目录；文件缺失时自动画「图片占位」虚框，方便事后替换。

## 素材库

`assets` 动作会扫描样本 pptx 内的图片，按哈希去重后分类：

- `backgrounds/`：接近 16:9 的大图（底纹、背景图）
- `decorations/`：小尺寸装饰元素

`manifest.json` 记录每张图的尺寸、出现次数与来源文件。生成时把素材路径传给 `background` 或幻灯片 `image` 字段即可复用模板视觉。

## 风格档案（learn）

学习几十份不同主题的样本后，档案里记录：

- **结构习惯**：平均页数、目录页比例、结尾页比例、备注/图片/表格使用率、每页要点数
- **版式**：页面尺寸、标题/正文字号范围、常见字体、标题对齐倾向
- **语言**：标题平均长度、常见要点开头（如「现状：」「问题：」）
- **配色**：高频使用的颜色

生成新 PPT 时，把档案映射成 `style` 参数传给 `build`，例如：

```json
{
  "style": {
    "titleAlign": "center",
    "showPageNumber": false,
    "showFooter": false,
    "bulletChar": "▪",
    "titleSize": 30,
    "bodySize": 18,
    "palette": { "primary": "C0392B", "accent": "E67E22", "dark": "1A1A1E" }
  }
}
```

## 大纲格式

```json
{
  "title": "演示标题",
  "subtitle": "副标题",
  "theme": "navy",
  "slides": [
    { "type": "cover", "title": "...", "subtitle": "...", "tag": "2026.11" },
    { "type": "toc", "title": "目录", "items": ["一", "二"] },
    { "type": "section", "number": "01", "title": "章节名", "subtitle": "可选" },
    { "type": "content", "title": "页标题", "bullets": ["要点", { "text": "子要点", "level": 1 }], "notes": "备注" },
    { "type": "two-column", "title": "...", "leftTitle": "左", "leftBullets": [], "rightTitle": "右", "rightBullets": [] },
    { "type": "table", "title": "...", "headers": [], "rows": [[]] },
    { "type": "end", "title": "谢谢观看", "subtitle": "可选" }
  ]
}
```

## 主题

`navy`（海军蓝）、`teal`（松石绿）、`wine`（酒红）、`graphite`（石墨）。

## PPT 制作经验

### 模板

1. 下载来的模板，视觉设计几乎全在**页面形状**里（图片、色块、艺术字），母版和版式通常只有占位符。想复用就复制页面再改字——插件的模板模式正是这样做的，"继承母版"行不通。
2. **字体是设计资产**。模板缺字体（思源黑体、演示镇魂行楷等）会被系统字体替换，气质直接打折。用模板前先装模板包里的字体。
3. 红色党政风的视觉公式：红白高对比 + 大面积留白 + 粗大标题 + 图片压边 + 底部色条。学会公式，自由版式也能做出七八分神似。

### 自动套模板的边界

4. 结构越简单越能自动替换：封面、章节页、结尾页（只有标题+副标题）成功率高；四卡片页、雷达图页、时间轴页每个文本框有特定语义，自动替换必然错位——这类页面手动改。
5. **文字容量是硬约束**。模板给一个文本框多大地方就只能放多少字，超出会遮挡、溢出、压到别的元素。内容必须适配版式，不能反过来。
6. **渲染验收不能省**。阴影、溢出、字体替换、图片比例，只有渲染成图才看得见；生成完直接交付等于交盲盒。

### 参考材料

7. 参考文档的最大价值是**结构**（章节顺序、详略分配、语气），视觉是次要的——结构可以套用到任何视觉模板。
8. 固定赛道（职业规划、述职、答辩）有评委期待的结构，偏离结构写得好也不讨好。

### 使用插件时的要点

9. 先估算内容量（每页字数、条目数）再选版式，而不是生成完才发现溢出。
10. 复杂模板尽量用 `samplePages` 手动指定样板页，自动分类只适合结构规整的模板。
11. 横图放进竖图区时，等比缩放会显得很小，需要铺满裁切才饱满；封面/章节页配图优先用横版。

## 配置（profile 的 cordis.patch.yml）

```yaml
- insert:
    - id: ppt-forge
      name: '@local/ppt-forge'
      config:
        pythonPath: ''      # 留空自动探测；也可用环境变量 DSH_PYTHON
        defaultTheme: navy
```

## 依赖

- Python 3，内置运行时自带 `python-pptx` 1.0.2；
- Windows 下自动探测顺序：配置 `pythonPath` → `DSH_PYTHON` → DSH 自带运行时 → PATH 中的 `python`。

## 卸载

在插件管理器中移除 `@local/ppt-forge`，或：
`plugin_manager { action: "remove_bundle", target: "@local/ppt-forge" }`
