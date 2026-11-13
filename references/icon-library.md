# 语义图标库（内置）

路径：`<engine-root>/assets/icons/`

- `robot.png` / `robot_white.png`：机器人本机、执行端
- `gripper.png` / `gripper_white.png`：夹爪、抓取、交接
- `brain.png` / `brain_white.png`：任务规划、推理、记忆
- `bulb.png` / `bulb_white.png`：洞察、方案、判断
- `gear.png` / `gear_white.png`：设备、参数、编排
- `shield.png` / `shield_white.png`：安全约束、边界
- `database.png` / `database_white.png`：状态账本、证据、记忆
- `clock.png` / `clock_white.png`：时间窗、时序、等待

## 用法

- 一页 1–3 个，尺寸 0.32–0.62in；深色图用 `*_white.png`，纸面用深色版。
- 图标必须指向机制（如“抓取/交接”旁边放 gripper，“安全”旁边放 shield），不能当角落装饰。
- 用 `blocks.image` 或 custom builder 的 `artwork()` 直接放；不要用 emoji 代替图标。
- 同一页图标线宽/尺寸一致；不要 8 个图标铺成图标墙。
