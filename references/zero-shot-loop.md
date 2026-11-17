# zero shot 的 loop 执行单（lead 负责）

zero shot 只产出 draft。除非以下全部完成，不说“完成”、不发布：

1. **构建**：spec + custom builder → deck.pptx → PDF/contact。
2. **第一轮独立审查**：另开只读子 agent；报告必须带 `review-meta`，reviewer ≠ builder。
3. **逐条修 must_fix**：回写 spec/builder，重建、重渲染；不得只改 PPTX。
4. **第二轮独立审查**：新开 reviewer，不能复用主 agent 自评；最多 3 轮。
5. **gate**：`review_gate.py <output>` 必须 exit 0；没有 pass 就只能叫 draft。
6. **交付报告**：写清轮次、已修/豁免、剩余风险；真实失败就写失败。

禁止：
- 自写 review 报告冒充独立审查；
- 只跑一轮审查就交付；
- 把“用户接受”写成“review_gate PASS”；
- 审查说 must-fix 时用“次级问题”绕过。
