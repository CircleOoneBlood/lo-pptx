---
name: visual-collab
description: >
  人类与 Agent 共同编辑视觉内容的工作流。PPTX 是载体，最终产物是图片（小红书配图、海报等）。
  当用户提到 PPT/演示文稿/海报/小红书配图/视觉内容，或需要在画布上协作时，务必激活此 skill。
  核心原则：每次收到用户消息先运行 lo diff，不覆盖人类改动，意图不确定就问。
---

# Visual Collab

人类与 Agent 共同编辑视觉内容的工作流。

PPTX 是载体，不是目的。它是一块**人类可以直接操作的视觉画布**——
人类拖拽、改色、改文字；agent 生成初稿、读懂改动、同步意图。
最终产物是图片（小红书配图、海报等），不是演示文稿。

核心价值：**视觉细节用操作表达，不用语言描述。**
直接把元素拖到想要的位置，比跟 agent 描述快一个数量级。

## 协作模型

```
Agent 生成初稿（PPTX）
       ↓
人类直接编辑（拖拽 / 改色 / 改文字）
       ↓
人类发消息（= "我改完了，看一下"）
       ↓
Agent 运行 lo diff（快照对比），报告变化，自动更新 baseline
       ↓
Agent 理解意图，决定是否需要跟进
       ↓
Agent 热重载或继续操作
```

对话框"发送"是天然的"编辑告一段落"信号。
**每次收到用户消息，先运行 lo diff，再回应。**

## 核心命令

所有命令在 `python3 -m lo` 下运行（lo 是项目内的 CLI 工具）。

### lo init
从零生成新的 PPTX 模板（新项目时用一次）。
```bash
python3 -m lo init -o output.pptx
```

### lo diff
对比 current PPTX vs baseline 快照，报告所有变化，并自动更新 baseline。
```bash
python3 -m lo diff
```
**每次收到用户消息后第一时间运行此命令。**

输出格式：
- 每个 slide 的变化列表
- 文本、位置、颜色、字号等属性变化
- 标注"需重新生成图片"的 slide

### lo shape
操作单个 shape 的属性（set-text、set-fill、move、resize、set-text-color、get-info）。
```bash
python3 -m lo shape set-text --name cover_title_main --text "新标题"
python3 -m lo shape set-fill --name cover_accent_bar --color "#FF3B5C"
python3 -m lo shape move --name cover_title_main --x 60 --y 200
python3 -m lo shape get-info --name cover_title_main
```

### lo export
导出 slide 为 PNG/PDF。**每次导出自动覆盖同名文件**。
```bash
python3 -m lo export png --slide 1 --out slides/slide1.png
```
修改 PPTX 后，导出对应 slide 的 PNG 来确认效果。

### lo reload
热重载：在 LibreOffice 中重新打开 PPTX（实时看到修改效果）。
```bash
python3 -m lo reload --open
```

## 原则

1. **diff 优先**：先运行 lo diff 看改了什么，再决定如何回应
2. **自动快照**：lo diff 完成后自动将 current 保存为新 baseline，无需手动确认
3. **不覆盖人类改动**：重新生成时，已接受的人类编辑要保留进代码
4. **意图不确定就问**：lo diff 能精准定位变化，但为什么改还是要人类说明
5. **改完必导出**：修改 PPTX 后，用 `lo export` 生成图片确认视觉效果
6. **图片自动覆盖**：`lo export` 每张 slide 生成一张 PNG，后续导出自动覆盖，无版本管理

## 何时使用此 Skill

- 用户提到需要做 PPT、演示文稿、海报、小红书配图
- 用户在 pptx 相关目录工作
- 用户请求生成或编辑视觉内容
- 用户说"我改完了"或类似表达到达编辑告一段落的信号

## 参考文档

- `references/lo-commands.md` — lo CLI 命令详细参考
- `references/workflow-principles.md` — 协作原则详解

## 项目结构

```
pptx/
├── lo/                      # lo CLI 工具（Python）
│   ├── __main__.py         # 入口
│   ├── commands/          # 子命令
│   │   ├── init.py        # 生成模板
│   │   ├── diff.py        # 快照对比
│   │   ├── shape.py       # 形状操作
│   │   ├── reload.py      # 热重载
│   │   └── export.py      # 导出
│   └── core/              # 核心逻辑
│       ├── pptx_ops.py
│       ├── shape_finder.py
│       └── config.py
├── pptx-workflow.json     # 工作流配置
└── *.pptx                 # 当前 PPTX 文件
```
