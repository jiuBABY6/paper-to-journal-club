---
name: audit-journal-club-deck
description: Review an academic journal-club deck for evidence traceability, scientific wording, editable-content coverage, and presentation readability.
---

# Audit Journal Club Deck

使用 `audit_journal_club_deck` 作为结构与证据门禁，然后再检查 PowerPoint 新鲜导出的预览图。

未传入 `required_sections` 时，审核器默认强制要求背景、创新点、方法、实验数据、局限性和未来方向六个模块。下列任一情况均为硬失败，不能进入 PPT 生成：

- 缺少任一必备 `section` 幻灯片；
- 必备页没有可追溯的 claim、section 或 figure 来源；
- 必备页为 `missing` / `needs-review` 状态；
- 实验数据页没有结果 claim，或引用不存在的 claim / section / figure id；
- 未来方向是汇报者建议，却没有明确标为 `presenter-discussion`，或未回链到作者结果/局限性。

候选 claim 仍为 `needs-review` 时至少给出科学核对警告。文本裁切、过多项目符号、整页/整图截图遮蔽可编辑对象也必须报告。每条发现应包含页码、对象或 section、证据、修正动作和可验证的验收条件。审核阶段不要直接修改；将发现返回给内容或演示文稿作者。
