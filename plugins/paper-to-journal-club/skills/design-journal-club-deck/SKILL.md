---
name: design-journal-club-deck
description: Design an evidence-backed journal-club narrative and slide-by-slide deck specification from an academic-paper evidence pack. Use only for paper-to-journal-club workflows, not for reference-image recreation or standalone scientific-figure design.
---

# Design Journal Club Deck

仅在 evidence pack 已检查后调用 `design_journal_club_deck`。

默认将 `deck_spec` 保留在工具返回结果和当前对话中；除非用户明确要求长期保存 JSON 文件，否则不要传 `output_path`。

默认 `required_sections` 必须为：

- `background`：研究背景与知识缺口；
- `innovation`：创新点与贡献；
- `methods`：研究方法与实验设计；
- `experimental_data`：实验数据与核心发现；
- `limitations`：局限性与批判性评价；
- `future_directions`：未来研究方向。

用户明确要求精简汇报时，才传入较小的 `required_sections` 数组；该数组不能为空，且只能使用上述稳定 ID。不要用“默认省略”来绕过任何必备模块。

按听众和时长组织叙事，不要机械地一章对应一页。默认顺序固定为“研究问题与核心结论 → 背景缺口 → 创新贡献 → 研究设计与系统 → 逐项实验依据 → 局限 → 下一步 → 结论”；不要把实验结果放在方法之前，也不要把结论放在中间。每个幻灯片必须有明确的 `section`、`source_claim_ids` / `source_section_ids` / `source_figure_ids`、`evidence_status` 和 `content_mode`。

非封面标题必须是简短陈述句，不使用问号、`标题：解释` 或“实验数据 1”式空标题。每页最多三个支持点；标题、结论和要点应服从服务端的可读性长度上限，而不是靠缩小字号塞入全文。

视觉素材默认只用于方法/系统结构、主实验结果和消融实验。不要为了“图多”而加入案例分析、失败示例或错误案例；它们默认跳过。用户明确要求使用这类素材时，必须将其标为补充讨论或局限性，并维持同样的来源、图注和人工确认要求。

方法/系统页若有可靠结构图，必须标记为 `visual_role=system-architecture`，并给出 2–4 条能回链到方法段或图注的 `explanation_points`，依次解释输入、关键模块、输出和验证。主实验和消融实验页必须至少引用一个结果 claim，并生成 `result_analysis` 中的比较、解释和局限；只有原文存在明确比较性证据时，才能写“更优”“未改善”或“无明确差异”。

每张要插入的 Figure/Table 都遵循固定流程：先按完整编号定位 evidence pack 实体及来源页；若 `Fig. 1` 所在页恰好只有一个已识别 Figure 和一张合规导出位图，才可用 `automatic_binding` 将 `fig-1` 精确绑定到 `page-03-image-01`。多图、多图片、表格、矢量图或需要指定 panel 时，先调用 `render_paper_visual` 渲染候选页，再由用户确认裁剪区域，最后把已确认的 `image_path` 传给 `figure_asset_selection`。保留 slide 的 `source_asset_id`，使 PPT 内图片可回溯到原始解析资产。必须按完整编号匹配，`Fig. 1` 绝不能命中 `Fig. 10`；结论所在页有多个图、多个候选或关联不明确时，不允许按页码猜测绑定；没有可靠图片时不插图，也不保留视觉占位卡。

确认后的 Figure/Table 图片只能配合原文和图注生成说明：系统页写入 `explanation_points`；主实验或消融页写入 `result_analysis` 的比较、解释和限定条件。表格映射进入结果页后使用 `visual_role=result-table`，但结论仍只能来自关联的结果 claim、原文证据和图注，不能从图像像素推断。

必须区分作者直接证据、作者解释和汇报者讨论问题。没有证据的模块只能生成待核对占位页，随后由审核器阻断生成；不得补写实验结果、创新点、局限性或未来方向。若受众、语言、时长或科学重点会实质改变叙事，应先让用户确认。
