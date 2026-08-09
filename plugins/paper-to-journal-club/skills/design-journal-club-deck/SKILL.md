---
name: design-journal-club-deck
description: Design an evidence-backed journal-club narrative and a slide-by-slide deck specification from an academic-paper evidence pack.
---

# Design Journal Club Deck

仅在 evidence pack 已检查后调用 `design_journal_club_deck`。

默认 `required_sections` 必须为：

- `background`：研究背景与知识缺口；
- `innovation`：创新点与贡献；
- `methods`：研究方法与实验设计；
- `experimental_data`：实验数据与核心发现；
- `limitations`：局限性与批判性评价；
- `future_directions`：未来研究方向。

用户明确要求精简汇报时，才传入较小的 `required_sections` 数组；该数组不能为空，且只能使用上述稳定 ID。不要用“默认省略”来绕过任何必备模块。

按听众和时长组织叙事，不要机械地一章对应一页。每个幻灯片必须有明确的 `section`、`source_claim_ids` / `source_section_ids` / `source_figure_ids`、`evidence_status` 和 `content_mode`。每张实验数据页必须至少引用一个结果 claim；有对应图片时还应引用 figure id。

若要插入从 PDF 导出的真实图片，只能通过 `figure_asset_selection` 显式传递经审阅的 `figure id -> 图片路径` 映射。服务会验证文件存在、格式和 figure 追溯关系；当结论所在页有多个图时，不允许按页码猜测绑定。

必须区分作者直接证据、作者解释和汇报者讨论问题。没有证据的模块只能生成待核对占位页，随后由审核器阻断生成；不得补写实验结果、创新点、局限性或未来方向。若受众、语言、时长或科学重点会实质改变叙事，应先让用户确认。
