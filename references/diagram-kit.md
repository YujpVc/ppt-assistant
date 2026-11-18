# 流程图基础元素库（`diagram_kit.py`）

用于 custom builder，全部是原生可编辑形状/连接线。**一条关系只调一次 `edge()`。**

## 节点

`node(slide, x, y, w, h, text, kind=..., sub=...)`

| kind | 用途 |
|---|---|
| `box` | 普通步骤/实体 |
| `round` / `pill` / `start` / `end` | 起止、入口出口、状态 |
| `diamond` | 判断 |
| `hexagon` | 判定/约束 |
| `circle` | 事件、交汇点 |
| `cylinder` | 账本、数据库、记忆 |
| `document` | 记录、证据、报告 |
| `subprocess` | 子流程/编排 |
| `manual` | 人工介入 |
| `delay` | 等待/窗口 |
| `note` | 说明、口径、豁免 |

`port(shape, side)` 取 `left/right/top/bottom/center` 锚点。

## 连线

`edge(slide, a, b, from_side=..., to_side=..., label=..., dash=..., color=...)`

- 自动走正交折线，只在终点一个箭头；
- 支持 `label`、`dash`、`arrow_start`；
- 不要用多个 STRAIGHT 连接线拼折角。

## 辅助元素

- `lane()`：泳道/容器底；
- `badge()`：小状态标签；
- `legend()`：实线/虚线/颜色图例；
- `bracket()`：等权节点的分组括号；
- `axis()`：时间轴/刻度；
- `note()`：口径/边界说明；
- `text()`：无框文本。

## 纪律

- 一页最多两个中性面 + 一个主强调色；
- 语义色只给状态，不能兼作普通卡底；
- 图注必须与图一致；通用场景只能是“场景示意”，不占证据位；
- 四柱等权必须有分组/阶梯/对照隔断。
