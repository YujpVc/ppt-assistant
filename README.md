# PPT 助手

DeepSeek Harness 的独立 PPT 助手预设：需求分析、大纲、可编辑 PowerPoint 生成和修改、自主 GPT 生图、满幅/渐变蒙版图片排版、结构质检、可选的联网配图与渲染审查。安装只插入 `ppt-assistant` 预设，不修改标准模式、模型路由、权限或现有会话。

## 安装

两个通道，内容同源；内网优先用私库，公网用 GitHub。

- **公网（GitHub）**

  ```bash
  dsh plugin --profile web add https://github.com/YujpVc/ppt-assistant/releases/download/v1.5.0/ppt-assistant-1.5.0.tgz
  ```

装完新建会话并选择 **PPT 助手**。包名为 `@yujp_vc/ppt-assistant`，通过 GitHub 与内网 GitHub 双通道分发，尚未发布到 npm。社区市场收录申请见 [PR #10056](https://github.com/awesome-dsh-plugin/awesome-dsh-plugin/pull/10056)，审核合并前无法在市场目录搜索到。

此包保留原有 PPT 助手提示词和工具组合，同时附带 Python 生成引擎和通用主题资源；不复制用户配置、API 密钥、机构模板、机构 Logo、历史会话或生成的文件。提示词沿用原来的引擎发现规则：`deepseek-harness/ppt-assistant`、`~/ppt-assistant` 或安装后的 `node_modules/@yujp_vc/ppt-assistant`。

## 运行环境

使用 Python 3.10+。在引擎目录运行 `bash run_py.sh --help` 可确认解释器；`run_py.sh` 会挑选真正装有 python-pptx/Pillow 的 Python（DSH 非登录 shell 里的 `python3` 可能只是运行时）。依赖安装：`python3 -m pip install -r requirements.txt`，不会在插件安装时自动执行。`run_pptfast.sh` 可选使用已安装的 `@liustack/pptfast@0.20.0`，缺少时使用内置 `build_deck.py`；本包不自动下载 pptfast。

生图客户端 `gen_image.py` 只用标准库，不需要额外依赖。它的配置按顺序从 `--config`、`$PPT_IMAGE_CONFIG`、引擎目录 `image-config.json`、`~/.config/ppt-assistant/image.json` 读取，格式为 `{"api_base": "https://uuapi.io/v1", "api_key": "..."}`；也接受 `UUAPI_API_KEY` / `UUAPI_BASE_URL`。密钥不写入本仓库、不写进成图、不打印。

渲染需要 LibreOffice 和 `pdftoppm`，扫描文档 OCR 需要 Tesseract。联网取图可通过 `PEXELS_API_KEY` 设置个人密钥，视觉审查需要 `DEEPSEEK_API_KEY` 和可用的视觉模型权限。旧提示词提到的 dsh-office、mindmap 等补充插件需另行安装，本包不自动安装或宣称提供这些插件的工具。

## 使用

```bash
bash run_py.sh build_deck.py spec.json output.pptx
bash run_py.sh quality_check.py output.pptx
bash run_py.sh build_deck.py template-spec.json output.pptx --template template.pptx
bash run_py.sh quality_check.py output.pptx --template template.pptx
```

内置默认生成器支持原生可编辑文本、表格与图表。`quality_check.py` 会拒绝「整页图片、没有可编辑文字」的稿——一键 AI 工具导出的图片 PPTX 不能作为交付。模板、Logo 由使用者显式提供，本包提供通用默认模板；旧版 `generate_ppt.py` 可用 `--template` 指定用户模板，不包含本机机构素材。生成质量取决于输入素材和逐页检查；没有运行视觉检查时不得声称已通过。

## 自主生图

```bash
bash run_py.sh gen_image.py --prompt-file hero.txt --out assets/hero.png \
  --size landscape --model gpt-image-2.5-flare --retries 1
```

- 尺寸：`1024x1024`、`1536x1024`（landscape，满幅 16:9 用这个再裁）、`1024x1536`。
- `--ref a.png --ref b.png` 走参考图编辑，最多 4 张；`--dry-run` 先校验配置与端点。
- 每次生成会写同名 `.json` 记录提示词、模型、耗时与哈希；密钥不在记录里。
- 单张约 30 秒到 4 分钟，慢是上游排队、不是卡死。整套 deck 建议 3–6 张，同一组光线/色温词保证风格统一。
- `--transparent` 生成透明底 PNG，用于 `cutout_statement` 的破格主体。
- 生成后跑 `image_check.py <图> --intent "..." --min-score 7`，有文字/水印/主体不符就重生成；透明抠图加 `--transparent`（预览会压成白底，别误判）。
- 出口被挡时可走 SSH 隧道：`ssh -D 1081 <host>`，再设 `PPT_IMAGE_PROXY=socks5h://127.0.0.1:1081`。

## 图片版式

`build_deck.py` 在原有版式外新增三种图片关系，全部由引擎完成裁切、渐变和阴影，不需要手工修图：

| 版式 | 用途 | 关键字段 |
| --- | --- | --- |
| `image_hero` | 封面/章节/收尾：满幅出血 + 渐变蒙版 + 大标题 | `image` `kicker` `title` `body` `scrim` `ink` |
| `image_overlay` | 满幅图 + 单侧渐变面板，放正文、要点、数字 | `image` `side` `title` `body` `bullets` `stats` |
| `image_wall` | 2–4 张瀑布流错位，圆角 + 细阴影 | `items[{image,title,body}]` |
| `image_triptych` | 三图一主两辅：左大图 + 右双叠 | `items[3]{image,title,body}` |
| `image_panel` | 一侧整幅图 + 一侧实心纸面：标题/正文/3 要点/2 指标 | `image` `side` `title` `body` `bullets` `stats` `source` |

`image_split` / `image_grid` 也升级为满格裁切 + 圆角 + 软阴影。所有图片一律 cover 裁切、不留白边，蒙版颜色由图片暗部采样，文字区按页型自动避让。

## 结构版式

| 版式 | 用途 | 关键字段 |
| --- | --- | --- |
| `architecture` | 3 层总体路线：深色主张带 + 大编号 + 细网格 + 回写环 + 验证跑道 | `items[3]{title,body}` `band` `loop` `steps` `bottom_title` |
| `sequence` | 之字形主轴：3 步上下交替，适合机制/流程 | `items[3]{title,body,foot}` `bottom` |
| `pillars` | 3 根深色柱头 + 纸面正文，适合并列机制 | `items[3]{title,body,foot}` `bottom` |

学术/基金 deck 的总体路线优先用它重画，不把论文总图整张贴进来。

## 图件体检

```bash
bash run_py.sh figure_audit.py output.pptx
bash run_py.sh figure_audit.py output.pptx --strict   # 有 DENSE 就 exit 2
```

它用 OCR 估计图内文字的实际字号，并结合图片占版面积，标出投影上读不了的密集图。经验线：图内文字低于 14pt 就不该直接上版；4800px 论文总图塞进 7in，30px 字只剩约 3pt。`--no-ocr` 可跳过 OCR，只按像素密度判断。

## 炫技版式

| 版式 | 用途 | 关键字段 |
| --- | --- | --- |
| `section_number` | 章节页：巨型数字 + 网格底 + 进度点 | `number` `title` `items` `sections` `section_no` `art` |
| `image_statement` | 满幅图 + 巨型字艺 + 结论标题 | `image` `art_text` `art` `title` |
| `split_diagonal` | 两张图斜切对照 | `images:[a,b]` `seam` `labels` |
| `glass_overlay` | 满幅图 + 毛玻璃面板 + 数字 | `image` `glass_opacity` `stats` |
| `heat_bars` | 甘特时间轴：阶段按时长累加，颜色映射火力 | `items[{title,body,fire,level,minutes,time}]` |
| `quote_stage` | 金句页：大引号 + 网格底 | `quote` `by` |
| `kpi_showcase` | 大数字墙 + 进度条 | `stats[{value,label,note,share}]` |
| `cutout_statement` | 透明抠图主体压住版面/标题 | `image` `side` `height` `flip` |
| `image_annotated` | 原生圆点 + 引线 + 标签 | `points[{x,y,text}]` |
| `image_arch` | 拱形裁切 + 侧栏文字 | `image` `items` |

- `art` 可选 `gradient`（主色渐变）、`outline`（空心描边）、`image`（用图片填字，需配 `image`，自动加主色细描边）；字艺是装饰位图，正文仍是真文本。
- `**关键词**` 在主色里高亮；`meta.transition: fade|push|morph` 加页面切换；`treatment: mono|duotone|halftone|glitch` 做印刷/故障质感。
- 一套 deck 最多 2–3 页炫技，相邻页不连续堆特效。
- 路线选择：默认输出可编辑 PPTX；已有母版走模板模式；明确要 Google Slides 时说明本环境无原生导出；浏览器动画/交互是核心时才走单文件 HTML（需自备 layouts/styles，HTML 不在 quality_check 覆盖范围）。

## 视觉总监与风格配方

- 先明确受众、要促成的判断、内容类型、场景、情绪；再写七行风格简报：三个形容词 / 一个设计张力 / 密度 / 配色逻辑 / 字体逻辑 / 图片行为 / 空间节奏。
- 从 `<engine-root>/style-recipes.md` 选一套主配方：`consulting` `formal` `grant` `civic` `artfair` `editorial` `history`；需要时最多加一套辅助，声明主/辅。把 `meta.recipe` 记进 spec。
- 配方 tokens 落进 `meta`：`theme` `accent` `colors` `font` `titleFont`。`titleFont` 可让标题单独用衬线（如 `Noto Serif CJK SC`），正文仍用 `font`。
- 评审顺序：先 `contact_sheet.py` 出缩略图总览，看节奏/密度/重复，再抽封面、最密页、特殊媒体页、结束页全尺寸看。

```bash
bash run_py.sh contact_sheet.py <工作区>/.render --out contact.png --cols 4 --labels
```

## 页计划与风格

- 动笔前先写四行 style brief：用途与受众、气质、主色（只选一个）、图像风格。主色写进 `meta.accent`；需要换底色、面板或文字色时用 `meta.colors` 覆盖 `bg/panel/text/muted/accent`（6 位 hex）。
- 每个正文页的标题必须是一句话结论（「执行率 88.5%，预算仍有余量」），不是名词短语。写 spec 前先列页计划：`页 / 结论句 / 证据 / 视觉类型 / 来源`。
- 配色按 HSL 想：邻近色相做区分，同色相拉明度做层级，降饱和+提亮=莫兰迪；一页只留一个高饱和重点色。

## 改稿合同（scoped edit）

改现有 PPT 时不要直接动手，先立基线：

```bash
bash run_py.sh deck_baseline.py snapshot 原稿.pptx baseline.json
# 修改工作副本 → 重建
bash run_py.sh deck_baseline.py compare baseline.json 新稿.pptx \
  --protect P01,P02,P05 --expect P03 --out ledger.md
```

- 稳定页名 `P01…` 写在 slide name 里，换序也不丢身份；台账按页列出标题、文字、数字、图片的变化。
- 保护页被改、越权修改、缺页都会让命令以 exit 2 结束；交付时附上 `ledger.md`。
- 数字逐位核对，不为了排版改精度；截图与交付文件版本必须一致。

## 打包

```bash
npm pack          # 生成 @yujp_vc/ppt-assistant-1.5.0.tgz
```

插件不包含模型服务或账号，遵循会话已有权限。安装需要 DSH `0.2.1-rc.3` 或兼容的 `0.2.x`，已在 Linux Desktop 0.2.1-rc.3 的插件管理入口验证安装；激活验证结果以发行说明为准。

引擎与预设提取自本机 DeepSeek Harness PPT 助手，保留 MIT 许可及上游署名。
