# PPT 助手「密度优先」配方（博采众长版）

来源：CyberPPT（高密度咨询式）、pptfast（DSH deck 引擎）、Awesome-PPT-Design-Skills（日系/未来科技编辑风）、ppt-master（工作流纪律）、dsh-office、dsh-wanghong-handwritten-ppt，以及本机多轮 DeepSeek 视觉模型的实测迭代。

核心结论一句话：**pptfast 的默认是「编辑风留白」，要「饱满充实」必须显式走密度优先配方——dense 节奏 + 满幅背景 + 富内容 + 多版式 + 禁止大白卡 + 不对称构图。**

## 1. 引擎与主题
- 引擎：pptfast（语义 IR → 原生可编辑 PPTX + `audit` 自动几何质检）。
- 主题按用途选：
  - 实用/指南/汇报 → `consulting`（最密）或 `enterprise`；
  - 情感/叙事 → `bloom`（暖玫瑰）；
  - 科技 → `tech`；财务 → `insight`。
- **narrative.pacing 一律用 `dense`**（balanced/spacious 会大量留白，用户明确拒绝「空荡荡」）。
- 已知坑：`deck.spec.json` 的 `theme` 只接受**字符串**（不能写 `{id, style}` 覆盖主题色）；要改主题色用 `render --style <style.json>`（如把 consulting 的 muted `#786961` 加深到 `#5F5148` 级别，修掉暖纸渐变上的 0.06 对比度差）。

## 2. 背景（消灭大白板的关键）
- 每页必设 `background`，**禁止纯白**：
  - 满幅照片：`{"kind":"asset","asset_id":<图>,"overlay":{"color":"#FFFFFF","opacity":0.85~0.88}}`（照片可见 + 文字可读，实测 0.88 是「饱满 vs 对比度」最优）。
  - 或暖纸渐变（CyberPPT 8 色之一，如 `#F3F4EF`→`#F4F1EA`），用于表格/对比页（表格要最干净底）。
- 封面/结束页用 `background: asset` + 深色 overlay（`#0B1220` 0.55）+ 白字。

## 3. 排版尺度（数值化，来自编辑风 style-system）
- 封面标题 56–76px 加粗；页标题 36–48px；正文 18–24px；说明/标签 11–13px。
- 版心安全边距 ≥72px；12 列网格 + 24px 沟槽。
- 行高：大标题 1.08，正文 1.35–1.55。
- 线条 1–2px、直角 0–4px，**不用软阴影/霓虹/3D**。

## 4. 内容密度（每页至少 4 个元素，忌纯 bullet）
- 每页 content 至少 3–4 个组件：`paragraph` + `bullets` + `verdict_banner` + `image`/`image-split`；数字用 `kpi_cards`、对照用 `comparison`、流程用 `steps`。
- **避免「密集 bullet 列表」**（编辑风明令：bullet 墙 = 单调）：把要点转成「标签块/步骤/时间线/对比模块/图标卡」。
- **不用「大面积默认白卡片」**（CyberPPT 明令：默认白卡/默认圆角/默认阴影 = 失败）——卡片底色交给背景，不要叠白卡。

## 5. 版式变换（每套 ≥5 种，相邻页不同）
- 显式 pin 常用：`image-split`(image_side left/right)、`image-top`、`quote-stage`(金句)、`two-column`、`quiet-frame`、`narrow-column`。
- **不对称构图**（编辑风核心）：一个主导块 + 一个辅助簇，别每页都居中对称；留一侧「安静区」形成节奏。
- 不 pin 则交给引擎按 `seed` 自动选型（相邻页不重复）。
- **避开 `image-lead-split`**：它右侧带一个空「手机占位框」，只有 `device_mockup` 组件能填；普通图文别用它。

## 6. 图片（图文并茂 + 切题 + 默认国内）
- **默认走必应国内图片搜索**（cn.bing.com，直连、无需 key、国内站点优先：苏宁/知乎/新浪/阿里/京东图床），国内题材贴切、不违和；国内找不到 → Pexels（外网、走代理）。
- 关键词用标准说法（如「电动自行车」而非「两轮电瓶车」）命中更准。
- 抓图后**逐张 vision 核验「图对不对题」**，不对就换关键词重抓。
- 图片放 `assets/`，PIL 归一化成真 JPEG（源图常是 PNG 却命名 .jpg）。
- 禁止真人/伤口/AI插画（学术体制内类）；情感类可用牵手/剪影（无脸）。

## 7. 容量上限（实测，写死避免试错）
- `bullets` 最多 **6 条**（dense）。
- `comparison` 最多 **5 行**。
- `kpi_cards` 最多 **3 张**（4 张窄卡会被省略号截断文字）。
- `steps` 最多 **5 步**（再多会溢出、压到底部 verdict）。
- `image-split` 半宽侧塞不下 `paragraph + kpi(3) + verdict`；只放 `image + 3 张卡 + verdict`。
- 超了就拆页/换组件，别硬塞。

## 7.5 调研/分析类内容结构（学 TraeWork/WorkBuddy，天然密度高）
封面/目录 → **多维度对比总表**（每个维度打分 + 综合第一标绿/色块突出）→ **逐对象深度档案**（每家/每项一页，末尾一段优劣势小结）→ 关键差异点表格+色块 → 结论页 **3 条可执行建议** → 最后一页「**一页速览**」总结卡片（可直接截图转发）。这种结构比「标题+要点」密度高得多，也更有用。

## 8. 质检（闭环）
1. `pptfast audit <deck项目>` → 必须 exit 0（溢出/越界/低对比/重叠/截断全查）。
2. `python3 vision_check.py <render/pg-xx.png> "<问题>"`（DeepSeek `deepseek-v4-flash-vision-exp`）逐页评饱满度，修到 ≥8。
3. 渲染：`soffice --headless --convert-to pdf` → `pdftoppm -png`。
4. **查「单字成行」**：读 `pptx_read` 输出，看有没有「E / 照」「6–7 / 折」「…三 / 点」这类数字/字母后跟单字量词被断行的——有就重排/缩短标题要点。
5. **以 audit 为准**：consulting 主题的 banner-motif 黄色装饰条会被 vision 误判成「文字被遮挡」，但确定性 audit 是 0 问题——几何结论信 audit，vision 只做「留白/断行/单调」这类主观反馈。

## 8. 边界页（已知限制，写进预期）
- pptfast 的 `cover`/`chapter`/`ending` 是「边界页」：`ending` 运行时只保留 `heading`（subheading/background/decor 全被丢弃），天然简洁——把它当「谢谢页」，别指望它「饱满」；收尾的丰富度放在它前面的 content 页（quote-stage 满幅金句）。
