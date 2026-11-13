# 本机历史稿里真正好用的版式原型

这些不是固定模板，而是“页面要义 + 元素关系”。复用原型时换内容、换图、换宽窄，不要照抄坐标。

| 原型 | 历史稿 | 结构 | 适合 |
|---|---|---|---|
| 双磨砂章节页 | cover_frost_test / v13 plus P1 | 左数字 `number_frost=0.50`；右侧 `panel_opacity=0.28` 圆角磨砂；同一左基准，不要分割线 | 章节过渡/封面，左右都磨砂 |
| 侧图清单页 | v12 plus P5/P7 / v11 plus P2/P6 | 一侧满高出图 45–50%；另一侧编号行 + 标题 + 一句机制 + 右端标签；底部一条入口/出口/回写带 | 能力、恢复、问题定位 |
| 四道闸页 | v11 plus P7 / v10 P6 | 四列等宽卡，卡头浅底 + 卡内“查什么/不通过怎么办”；底部处置胶囊链 + 边界带 | 约束、门、检查 |
| 图像证据网格 | v12 plus P4 / v10 P2 | 满幅图或图带 + 3–4 个编号标注 + 底部三栏证据；图上文字必须有磨砂或蒙版 | 现场、场景、机制证据 |
| 玻璃双栏 | v13 plus P6 | 满幅图 + 左宽右窄两块玻璃；左三行机制，右坐标/边界 | 记忆、恢复、指标体系 |
| 学术层级表 | v13 plus P7 / v10 P8 | 表题 + 两级表头 + 01–04 层级号 + 左轨 + 行带 + 末列浅底 + 表下口径 | 验证、评价、真二维矩阵 |
| 总路线 | v11 plus P4 / v10 P4 | 顶部三段主链；中部三列机制；底部四级验证跑道；一条回写旁路 | 总体路线、系统关系 |
| 两栏判断 | v12 plus P9 / v9 plus P7 | 满幅图 + 标题压图 + 下方两栏结论；或左文右图 + 两行判断 | 本章落点、结论页 |

## 引擎内置的机制关系页

以下四种由 `build_deck.py` 直接渲染，几何固定、文案来自 spec；不要再在项目里写一次性 builder。四种都保留 stock 标题带、页码和页脚来源，正文最低 8 pt。

| 版式 | 页面要义 | 必填字段 | 可选字段（缺省时用结构标签） |
|---|---|---|---|
| `state_overview` | 一条共享状态 bus（S0–S4）压在上方，下面三块**不等宽**区域：规划 / 技能 / 安全；S4 虚线回写 S1/S2 | `title` `body` `steps[3–5]` `step_fields[]`（与 steps 等长）`state_reads[3]` `items[3]{title,body}` `writeback` `source` | `bus_label` `bus_sub` `bus_note` `step_codes` `planning_tags[2]` `task_graph_label` `chain[4]` `chain_note` `planning_rule` `planning_feedback` `planning_bullets[2–3]` `planning_note` `image` `image_caption` `gauges[3]` `vla_strip[3]` `vla_strip_widths[3]` `vla_loop_label` `edges{check,yes,fix,retry_check,retry,stop,note}` `stop_label` `normal_label` `overrun_label` `safety_note` |
| `task_swimlane` | 任务 / 设备 / 机器人三泳道；前三列是任务节点 + D/R 支撑块，第 4 列是 `gate=true` 的判定菱形与 D4/R4 | `title` `body` `lanes[3]{label,en,sub,fill}` `items[4]{code,title,device,robot}`（第 4 个标 `gate`）`dep_labels[2]` `source` | `legend_main` `legend_dep` `link_labels[2]` `gate_link_label` `device_note` `robot_note` `bottom_note` |
| `safety_mechanism` | 左侧三条约束量表（上限/下限方向），右侧上三分阶段导纳与两级阈值曲线、下完整处置分支 | `title` `body` `items[3]{title,body,note}` `steps[3]{title,body,param}` `branches[2]{title}` `source` | `constraint_label` `constraint_tag` `scale_labels[3]` `bound` `image` `image_caption` `phase_label` `phase_sub` `warning_label` `stop_label` `curve[]` `graze` `graze_label` `stage_note` `branch_label` `branch_sub` `decision_check` `retry_check` `yes_label` `continue_label` `retry_branch_label` `branch.short` `branch_footer` |
| `ledger_rails` | 左侧四行状态账本（执行 / 偏差 / 分支处置 / 回写），右侧 A/B 两条同构指标轨 | `title` `body` `items[4]{code,title,body}` `rails[2]{title,metrics[3+]{name,criterion,direction,unit}}` `branches[2]{title,body}` `ledger_fields[4+]` `rewrite_chip` `source` | `ledger_label` `rails_label` `field_label` `metric_headers[4]` `rail_note` `rail_legend` `item.fields` `branch.short` |

用法要点：

- 这四种是**关系图**，不是四张卡：`state_overview` 的三区宽度和 `task_swimlane` 的三层高度由引擎固定，换内容时不要改坐标；`safety_mechanism` 的曲线和阈值线、`ledger_rails` 的账本轨和指标轨也是一套结构。
- `state_overview` 的 `chain` 是规划区的小任务链（4 个短标签）；`edges` 是安全区的判定分支文案，用 `{check,yes,fix,retry_check,retry,stop,note}` 覆盖默认标签。
- `safety_mechanism` 的 `items[i].bound="lower"`（或 note 里出现“下限”）会把危险端画在左侧；`scale_labels` 可换三段刻度的文字。
- `ledger_rails` 的 `branches` 建议同时给 `short`，否则用 `body` 作第二行；`metric_headers` 必须是 4 个短标签。
- 四种版式都走 stock 的 `source_line` 与 `ensure_page_number`，`source` 不能省；`density_check.py` 会检查关键字段、模块数和 8 pt 下的标签长度。

## 组合纪律

- 一页只选一个原型作为骨架；其余元素只做辅助，不要把所有元素都塞进一页。
- 全章至少两种非 `blocks` 原型；不要 8 页全部 blocks。
- 侧图清单和四道闸可以各有 1–2 页；超过就换原型。
- 图片必须承担语义（场景/证据/对象），不能只是背景。
- 大字标题和正文分工：标题给结论，正文只给机制/判据，不写段落。
- 每个原型都保留页脚、页码和来源。
