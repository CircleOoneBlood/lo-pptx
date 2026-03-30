# lo CLI 命令参考

## 概览

```
python3 -m lo <command> [options]
```

## init — 生成新模板

从零生成一个新的 PPTX 模板，包含封面 + 5 张图文页。

```bash
python3 -m lo init -o output.pptx
```

- 输出路径可省略，默认从 `pptx-workflow.json` 读取
- 自动创建对应的 `.baseline.pptx` 快照文件

## diff — 快照对比（核心命令）

对比当前 PPTX 与 baseline 快照，报告所有变化，**并自动更新 baseline**。

```bash
python3 -m lo diff                      # 使用默认路径
python3 -m lo diff --current my.pptx    # 指定当前文件
python3 -m lo diff --baseline base.pptx # 指定 baseline 文件
python3 -m lo diff --no-snapshot        # 仅报告变化，不更新 baseline
```

### diff 检测的变化类型

| 属性 | 说明 |
|------|------|
| `text` | 文本内容变化 |
| `font_size` | 字号变化（单位 pt） |
| `font_color` | 字体颜色（#RRGGBB） |
| `bold` | 粗体开关 |
| `italic` | 斜体开关 |
| `position` | 位置变化 (x, y) |
| `size` | 尺寸变化 (w×h) |
| `fill_color` | 填充色变化 |
| `shape_added` | 新增形状 |
| `shape_deleted` | 删除形状 |

### 输出示例

```
## PPTX Changes (3 change(s), 1 slide(s))

### Slide 1
**`cover_title_main`**
  - 文本: `初稿` → `终稿`
  - 位置: `(60,180)` → `(120,180)` ⚠️ 需重新生成图片

**`cover_accent_bar`**
  - 填充色: `#FF3B5C` → `#4ECDC4` ⚠️ 需重新生成图片

⚠️ **需要运行 `python3 -m lo export png` 生成新图片**

✓ Baseline auto-updated → cover.baseline.pptx
```

## shape — 形状操作

操作单个 shape 的属性。所有操作直接修改 PPTX 文件。

### set-text — 设置文本

```bash
python3 -m lo shape set-text --name cover_title_main --text "新标题"
```

### set-fill — 设置填充色

```bash
python3 -m lo shape set-fill --name cover_accent_bar --color "#FF3B5C"
```

颜色格式：`#RRGGBB`（如 `#FF3B5C`）

### move — 移动位置

```bash
python3 -m lo shape move --name cover_title_main --x 60 --y 200
```

位置单位：像素（px），基于 1080×1350 画布。

### resize — 调整尺寸

```bash
python3 -m lo shape resize --name cover_title_main --w 960 --h 120
```

尺寸单位：像素（px）。

### set-text-color — 设置字体颜色

```bash
python3 -m lo shape set-text-color --name cover_title_main --color "#FFFFFF"
```

### get-info — 查看形状信息

```bash
python3 -m lo shape get-info --name cover_title_main
```

输出形状的当前属性（位置、尺寸、填充、文本等），**不保存文件**。

## export — 导出

将 slide 导出为图片或 PDF。**每次导出自动覆盖同名文件**，无需手动删除旧文件。

### 导出 PNG

```bash
python3 -m lo export png --slide 1 --out slides/slide1.png
python3 -m lo export png --slide 1 --out slides/slide1.png --file my.pptx
```

### 导出 PDF

```bash
python3 -m lo export pdf --out output.pdf
```

## reload — 热重载

在 LibreOffice 中重新打开 PPTX，用于实时预览修改效果。

```bash
python3 -m lo reload --open
```

不加 `--open` 只重启，加 `--open` 会打开 GUI。

## 配置文件

`pptx-workflow.json` 定义工作流路径：

```json
{
  "pptx": "visual-collab.pptx"
}
```

lo 各命令会从此文件读取默认 PPTX 路径。
