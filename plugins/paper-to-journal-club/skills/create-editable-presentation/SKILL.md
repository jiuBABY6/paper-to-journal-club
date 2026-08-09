---
name: create-editable-presentation
description: Directly control Windows Microsoft PowerPoint to generate an editable journal-club deck from a reviewed deck specification and preserve a clear raster-versus-native-object contract.
---

# Create Editable Presentation

先调用 `powerpoint_status`；只有该工具或 `inspect_powerpoint` 通过 COM 返回当前窗口信息时，才能说“已连接当前 PowerPoint”。在 `generate_editable_pptx` 前调用 `audit_journal_club_deck`，并且只从通过硬性审核的 `deck_spec` 生成。生成工具会再次执行必备模块与证据审核，因此不能绕过此门禁。本技能只面向 Windows Microsoft PowerPoint；不要承诺 WPS 控制或验证。

生成器使用原生 PowerPoint 文本框和形状。需要使用论文图时，只插入紧裁剪的单个图片 panel，并将图注、箭头、图例、标题、边框和标注重建为独立可编辑对象。若 evidence pack 已提供 `suggested_image_path` 或 `figure_asset_candidates`，只能按其中明确给出的候选路径使用，不能按图号猜测文件。不得用完整幻灯片截图作为捷径。

生成默认新建后台演示文稿，不会修改当前打开的文件。生成后必须检查返回的 `quality_audit` 与 `preview_paths`；需要独立复核时，调用 `audit_editable_pptx` 以只读方式重新打开并导出原生预览。
