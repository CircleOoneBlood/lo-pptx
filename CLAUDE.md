# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 这是什么

**人类与 agent 共同编辑视觉内容的工作流。**

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

```bash
lo init        # 从零生成新的 PPTX 模板（新项目时用一次）
lo diff        # 对比 current PPTX vs baseline 快照，报告所有变化，并自动更新 baseline
lo snapshot    # 手动保存当前 PPTX 为新的 baseline（通常不需要手动调用）
```

## 原则

1. **diff 优先**：先运行 lo diff 看改了什么，再决定如何回应
2. **自动快照**：lo diff 完成后自动将 current 保存为新 baseline，无需手动确认
3. **不覆盖人类改动**：重新生成时，已接受的人类编辑要保留进代码
4. **意图不确定就问**：lo diff 能精准定位变化，但为什么改还是要人类说明
5. **技术可迭代，目标不变**：具体工具链（库、API、热重载方式）随时换，协作模型不变

## 当前阶段

已稳定。协作机制已提炼为 `visual-collab` skill，可全局使用。

此目录是 `lo` CLI 工具的家，也是当前工作区。
Skill 指令见：`~/.claude/skills/visual-collab/SKILL.md`
