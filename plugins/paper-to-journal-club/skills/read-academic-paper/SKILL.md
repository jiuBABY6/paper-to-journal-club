---
name: read-academic-paper
description: Parse a user-provided academic paper into an evidence pack before designing a journal-club deck. Use only for an academic-paper-to-journal-club workflow, not for reference-image recreation or standalone scientific-figure design.
---

# Read Academic Paper

先调用 `analyse_paper`，并将返回的 `evidence_pack` 作为 PPT 的唯一事实来源。

PDF 解析会返回逐页 `extraction.pages`、真实导出的图片资产，以及 Figure/Table 的候选来源。组会视觉素材默认只为三类内容取用：方法/系统结构、主实验结果和消融实验。案例分析、失败示例和错误案例默认只保留为可追溯证据，不选择、不渲染、不插入；只有用户明确要求时才进入后续流程。始终按完整 Figure/Table 编号匹配，`Fig. 1` 不得匹配 `Fig. 10`。

若同一页**恰好只有一个已识别 Figure 和一张合规位图**，服务会生成 `automatic_binding`，明确记录 `fig-1 -> page-03-image-01`、来源页和图片路径；这是唯一允许自动关联的情形。自动关联仍只能用于上述三类默认视觉内容。

多图页、多个候选图片、矢量图未能导出、或图号与页面无法可靠关联时，保留候选但不要猜测插图。此时可让用户确认后，再把 `figure id -> 图片绝对路径` 映射传给设计工具的 `figure_asset_selection`。没有可用或已确认的论文图时，PPT 不插图，也不生成伪图占位。

对要使用的每个视觉素材，遵守固定顺序：按完整 Figure/Table 编号定位 → 核对 evidence pack 中的来源页和原文/图注 → 对矢量图、表格或需要指定 panel 的页面调用 `render_paper_visual` → 由用户确认归一化裁剪区域 → 将返回的 `image_path` 映射到 `figure_asset_selection` → 在后续 PPT 中插入。表格与页面渲染图绝不自动插入，也不能从截图猜测数值、统计显著性或指标方向。插入后的系统说明、实验比较和消融结论必须依据对应方法段、结果原文和图注生成。

在设计组会前，检查下列六类证据是否可用：`background`、`innovation`、`methods`、`experimental_data`、`limitations` 和 `future_directions`。

- 不得把 OCR 片段或未经核对的候选 claim 改写成既成事实。
- 每个关键结论必须回链到 section、原文 excerpt、figure id 或页码；保留原始 claim id。
- 创新点只能使用作者明确的表述，或明确标为“汇报者归纳”并附方法/摘要证据。不要把常规方法误写成“首次创新”。
- 局限性优先取自 Discussion/Conclusion；未来方向优先取自作者原文。若仅能从局限性推出讨论建议，必须标为“汇报者讨论”，不能冒充作者结论。
- 将不可读图、缺失补充材料、不确定统计措辞和缺少来源的必备模块写入 `ambiguities`。
- 结果比较只在原文明确写出优于、未改善、无显著差异等措辞时才标注方向；单独的增加或减少不代表效果好坏。
- 当前解析器只能可靠导出 PDF 内嵌的 PNG/JPEG。矢量系统图、表格和整页截图尚不能自动裁剪；需要保留原始页码和图表上下文，并在后续由用户确认后再插入。
- 任何会实质改变科学结论的歧义都应由用户确认。

本阶段不制作幻灯片。
