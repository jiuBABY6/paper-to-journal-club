---
name: audit-journal-club-deck
description: Review an academic-paper journal-club deck for evidence traceability, scientific wording, editable-content coverage, and presentation readability. Use only for paper-to-journal-club workflows, not for reference-image recreation or standalone scientific-figure design.
---

# Audit Journal Club Deck

使用 `audit_journal_club_deck` 作为结构与证据门禁。默认根据返回的结构质量结果复核；只有用户明确要求视觉复核时，才检查 PowerPoint 新鲜导出的 PNG 预览图。

未传入 `required_sections` 时，审核器默认强制要求背景、创新点、方法、实验数据、局限性和未来方向六个模块。下列任一情况均为硬失败，不能进入 PPT 生成：

- 缺少任一必备 `section` 幻灯片；
- 必备页没有可追溯的 claim、section 或 figure 来源；
- 必备页为 `missing` / `needs-review` 状态；
- 实验数据页没有结果 claim，或引用不存在的 claim / section / figure id；
- 未来方向是汇报者建议，却没有明确标为 `presenter-discussion`，或未回链到作者结果/局限性。

还要检查叙事与可读性：非封面标题必须是简短陈述句，不能含问号或冒号；背景、创新、方法、实验、局限、未来和结论必须按因果顺序出现，结论页必须最后。系统结构页若插入图，必须有至少两条来自论文方法或图注的输入、模块、输出或验证解释。实验图表页必须同时有可回链的结果 claim，以及“比较、解释、限定条件”三层分析；原文没有明确比较或统计依据时，只能写“论文报告的观察”，不能自行判断效果好坏。每页最多三个支持点，过长文字应当作为硬性可读性问题返回。

同时审核视觉素材范围与链路。默认可纳入的仅是方法/系统结构、主实验结果和消融实验；未被用户明确要求的案例分析、失败示例或错误案例应标记为不必要并建议移除。每张插入图或表都必须能回链到完整 Figure/Table 编号、来源页和对应原文/图注；若使用 PDF 页面渲染或裁剪，还必须记录用户已确认的裁剪区域。说明文字必须由方法段、结果原文或图注支持，不能由截图内容补写。

候选 claim 仍为 `needs-review` 时至少给出科学核对警告。文本裁切、过多项目符号、整页/整图截图遮蔽可编辑对象也必须报告。每条发现应包含页码、对象或 section、证据、修正动作和可验证的验收条件。审核阶段不要直接修改；将发现返回给内容或演示文稿作者。
