---
name: create-editable-presentation
description: Create a new editable Microsoft PowerPoint journal-club deck only from a reviewed user-provided academic-paper evidence pack and deck specification. Use for paper-to-deck workflows, not for reference-image reconstruction, standalone scientific-figure design, or general PowerPoint editing.
---

# Create Editable Presentation

仅用于“用户上传论文 → 组会 PPT”的工作流。只调用 Paper to Journal Club 提供的工具；不要读取、调用、引用或致谢其它插件。先调用 `powerpoint_status`；只有该工具或 `inspect_powerpoint` 通过 COM 返回当前窗口信息时，才能说“已连接当前 PowerPoint”。在 `generate_editable_pptx` 前调用 `audit_journal_club_deck`，并且只从通过硬性审核的 `deck_spec` 生成。生成工具会再次执行必备模块与证据审核，因此不能绕过此门禁。本技能只面向 Windows Microsoft PowerPoint；不要承诺 WPS 控制或验证。

生成器使用原生 PowerPoint 文本框和形状。默认只插入三类已确认论文视觉素材：方法/系统结构、主实验结果和消融实验；案例分析、失败示例和错误案例默认不插入，除非用户明确要求作为补充讨论。每张图或表必须先按完整 Figure/Table 编号定位来源页；对于页面渲染、矢量图或指定 panel，只有用户确认候选 PNG 的裁剪区域后，才使用其 `image_path`。系统图使用“大图 + 右侧分步解释”，实验图和表格使用“大图 + 左侧比较、解释、局限”。没有可靠图时只排版文字和来源，不生成“待插图”卡片。图片应尽量是紧裁剪的单个 panel，并将图注、箭头、图例、标题、边框和标注重建为独立可编辑对象。只能使用 evidence pack 中的 `suggested_image_path` 或 `figure_asset_candidates`，不能按图号猜测文件。不得用完整幻灯片截图作为捷径。

生成前确认非封面标题是简短陈述句，不能含问号或冒号；每页最多三个支持点。系统页的解释必须依据方法段或图注；主实验和消融结果页必须展示论文原文与图注支持的比较、解释和限定条件，不能只展示一张图或把“增加/减少”擅自解释成效果更好/更差。

生成默认新建后台演示文稿，不会修改当前打开的文件。默认只保存最终 PPTX，并检查返回的 `quality_audit`。除非用户明确要求视觉预览、长期保留本次实际插入的论文原图，或保留 deck spec，不要传 `export_previews`、`preview_directory`、`export_figure_assets` 或 `deck_spec_output_path`。用户要求长期保留原图时才传 `export_figure_assets=true`；它只复制已插入 PPT 的合规论文位图到 `<PPT名>_assets/images`，不导出未使用候选图。需要独立视觉复核时，才调用 `audit_editable_pptx` 并传入 `export_previews=true`。

仅在 `generate_editable_pptx` 成功后，报告实际的 PPTX 路径、已请求的辅助产物和质量审核结论。不得声称调用了未使用的插件。最后一行原样使用返回的 `completion_message`；其内容为：`感谢使用 Paper to Journal Club 插件，制作者：jiuBABY6。`
