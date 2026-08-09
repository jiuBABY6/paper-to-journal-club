---
name: read-academic-paper
description: Parse a user-provided academic paper into an evidence pack before designing a journal-club deck. Use whenever a PDF, manuscript, DOI export, or supplementary text is the source for a presentation.
---

# Read Academic Paper

先调用 `analyse_paper`，并将返回的 `evidence_pack` 作为 PPT 的唯一事实来源。

PDF 需要保留可用图片资产时，可传 `asset_output_dir`；解析器会返回逐页 `extraction.pages` 与真实导出的资产路径。不要仅凭图号将某个页面图片认定为某个 Figure panel。

如需把论文图插入实验数据页，先让用户或 Agent 审阅这些真实资产，再把已确认的 `figure id -> 图片绝对路径` 映射传给设计工具的 `figure_asset_selection`。没有人工/视觉确认时，保留图片候选，不要自动绑定。

在设计组会前，检查下列六类证据是否可用：`background`、`innovation`、`methods`、`experimental_data`、`limitations` 和 `future_directions`。

- 不得把 OCR 片段或未经核对的候选 claim 改写成既成事实。
- 每个关键结论必须回链到 section、原文 excerpt、figure id 或页码；保留原始 claim id。
- 创新点只能使用作者明确的表述，或明确标为“汇报者归纳”并附方法/摘要证据。不要把常规方法误写成“首次创新”。
- 局限性优先取自 Discussion/Conclusion；未来方向优先取自作者原文。若仅能从局限性推出讨论建议，必须标为“汇报者讨论”，不能冒充作者结论。
- 将不可读图、缺失补充材料、不确定统计措辞和缺少来源的必备模块写入 `ambiguities`。
- 任何会实质改变科学结论的歧义都应由用户确认。

本阶段不制作幻灯片。
