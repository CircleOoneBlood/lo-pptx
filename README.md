# lo

Agent-friendly PPTX surgical editor via snapshot diff.

**核心用途**：人类在 LibreOffice 里直接拖拽/改色/改文字，Agent 通过快照对比自动读懂改动意图。

## 安装

```bash
pip install -e .
```

安装后全局可用：

```bash
lo init -o output.pptx    # 生成新模板
lo diff                   # 对比快照
lo shape set-text --name title --text "新标题"
lo export png --slide 1 --out slide1.png
lo reload --open          # 热重载预览
```

## 依赖

- Python 3.10+
- python-pptx

## 适用场景

- 小红书配图（竖版 4:5）
- 社交媒体海报
- 需要精细调整的视觉内容

## 工作流

```
Agent 生成初稿（PPTX）
       ↓
人类直接编辑（拖拽 / 改色 / 改文字）
       ↓
人类发消息（= "我改完了，看一下"）
       ↓
Agent 运行 lo diff，报告变化，自动更新 baseline
       ↓
Agent 理解意图，决定是否需要跟进
       ↓
Agent 热重载或继续操作
```

## 作为 Python 模块

```python
from lo.commands.diff import diff
from lo.commands.shape import shape_group

# diff: 对比当前 PPTX 与 baseline 快照
# shape: set-text / set-fill / move / resize / set-text-color / get-info
```
