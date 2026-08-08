<div align="center">

# SciDiagram PPTX｜科研图示原生复现

**把科研框架图、机制图和流程图，重建为真正可编辑的 PowerPoint 形状、文字与连接线。**

研究框架 · 技术路线 · 机制模型 · 算法流程 · 学术图示

[![validation](https://img.shields.io/github/actions/workflow/status/SciToolsmith/sci-diagram-pptx/validate.yml?branch=main&style=flat-square&label=validation)](https://github.com/SciToolsmith/sci-diagram-pptx/actions/workflows/validate.yml)
[![license](https://img.shields.io/badge/license-Apache--2.0-4C6FFF?style=flat-square)](LICENSE)
[![Codex Skill](https://img.shields.io/badge/Codex-Agent%20Skill-111827?style=flat-square)](skills/sci-diagram-pptx/SKILL.md)

[案例](#gallery) · [安装](#install) · [使用](#usage) · [适用边界](#scope) · [English](#english)

</div>

保留文字、公式、层级和拓扑；不以整页图片伪装“可编辑”。

<a id="gallery"></a>

## 真实案例

预览均由可编辑重建页实际渲染。下载 PPTX 后，可以逐个修改文字、形状和连接线。

<div align="center">

### 全球创新产业链研究框架

<a href="docs/cases/global-innovation-chain-framework.png">
  <img src="docs/cases/global-innovation-chain-framework.png" width="430" alt="全球创新产业链研究框架：纵向研究阶段、中央多层研究内容与右侧研究方法组成的可编辑科研框架图">
</a>

[查看大图](docs/cases/global-innovation-chain-framework.png) · [下载可编辑 PPTX](https://github.com/SciToolsmith/sci-diagram-pptx/raw/main/examples/cases/global-innovation-chain-framework-editable.pptx)

</div>

<table>
  <tr>
    <th width="50%">载荷循环机制图</th>
    <th width="50%">自适应 FIR 模态分解流程图</th>
  </tr>
  <tr>
    <td align="center" valign="top">
      <a href="docs/cases/load-cycle-model.png">
        <img src="docs/cases/load-cycle-model.png" width="100%" alt="载荷循环机制图：两个循环阶段之间通过实线与虚线连接的可编辑数学机制图">
      </a>
    </td>
    <td align="center" valign="middle">
      <a href="docs/cases/adaptive-fir-mode-decomposition.png">
        <img src="docs/cases/adaptive-fir-mode-decomposition.png" width="76%" alt="自适应 FIR 模态分解流程图：包含判断分支、反馈回路与模式选择模块的可编辑算法流程图">
      </a>
    </td>
  </tr>
  <tr>
    <td align="center">
      <a href="docs/cases/load-cycle-model.png">查看大图</a> ·
      <a href="https://github.com/SciToolsmith/sci-diagram-pptx/raw/main/examples/cases/load-cycle-model-editable.pptx">下载可编辑 PPTX</a>
    </td>
    <td align="center">
      <a href="docs/cases/adaptive-fir-mode-decomposition.png">查看大图</a> ·
      <a href="https://github.com/SciToolsmith/sci-diagram-pptx/raw/main/examples/cases/adaptive-fir-mode-decomposition-flowchart-editable.pptx">下载可编辑 PPTX</a>
    </td>
  </tr>
</table>

> 公开案例仅包含可编辑重建页，不包含任务源图、参考页或本机备注。

<a id="install"></a>

## 安装

在 Codex 中输入：

```text
请使用 $skill-installer 安装：
https://github.com/SciToolsmith/sci-diagram-pptx/tree/main/skills/sci-diagram-pptx
```

安装后可用 `$sci-diagram-pptx` 显式调用。如果没有立即出现，请重启 Codex。

<details>
<summary>手动安装</summary>

```bash
git clone https://github.com/SciToolsmith/sci-diagram-pptx.git
mkdir -p ~/.codex/skills
cp -R sci-diagram-pptx/skills/sci-diagram-pptx ~/.codex/skills/
```

</details>

<a id="usage"></a>

## 使用

上传参考图后，直接提出目标：

```text
$sci-diagram-pptx 把这张科研流程图忠实重建为原生可编辑 PPTX，
保留文字、公式、层级和连线关系，不要重新设计。
```

也可以要求它修复已有 PPTX、只做结构检查，或明确重建混合图中的某个示意图面板。

<a id="scope"></a>

## 适用边界

| 适合 | 不适合 |
| --- | --- |
| 研究框架图、理论框架与技术路线 | 柱状图、折线图、散点图、热图等数据图 |
| 科研流程图、机制图与算法流程图 | 显微照片、实验照片或通用 OCR |
| 科研系统结构图、概念模型 | 普通商业组织图、海报或自由创意设计 |
| 含数学符号或公式的学术图示 | 以坐标轴、尺度和数据几何承载证据的图件 |
| 已有科研图示 PPTX 的结构修复 | 翻译、论文写作或统计分析本身 |

一个简单判断：含义主要由**节点、标签与连接关系**承载，用本 Skill；含义主要由**坐标轴、尺度与数据**承载，用 [SciPlot](https://github.com/SciToolsmith/sci-plot)。

## 交付原则

常规重建默认交付两页：

- **第 1 页**：原生可编辑重建图。
- **第 2 页**：源图对照页。

公开案例为保护输入素材，仅发布第 1 页。

```text
确认目标 → 梳理结构 → 原生重建 → 渲染检查 → 必要时集中修复一次 → 交付
```

- **原生可编辑**：优先使用 PowerPoint 形状、文本框和连接线，不用整页截图冒充编辑性。
- **科学含义优先**：保护文字、公式、方向和拓扑；歧义可能改变含义时先询问用户。
- **聚焦验证**：完成一次全图渲染和一次结构检查；只为明显阻断项集中修复，避免无休止追逐像素微差。
- **跨平台务实**：默认交付一份便携 PPTX，优先稳定的标准形状、独立文本框、明确字体和留白；保障可读与可编辑，不承诺 Windows 与 macOS 像素完全一致。

[查看完整 Skill 工作流与质量规则](skills/sci-diagram-pptx/SKILL.md)

<details>
<summary><strong>开发与验证</strong></summary>

仓库 CI 检查 Skill 元数据、脚本和确定性结构规则；它不能替代科学含义核验，也不声称不同 PowerPoint 渲染器完全等价。

```bash
python3 -m pip install -r requirements-ci.txt
python3 tests/test_panel_crop.py
python3 tests/test_check.py
```

欢迎通过 Issues 或 Pull Requests 贡献修复。请勿提交来源不明的论文截图、第三方素材或带本机隐私信息的文件。

</details>

## 许可与声明

代码以 [Apache License 2.0](LICENSE) 发布。用户应确认输入素材的使用权，并在交付前核验文字、公式、箭头方向与科学含义。

SciDiagram PPTX 是独立社区项目，与 OpenAI、Microsoft、Nature、Springer Nature 或任何期刊不存在官方隶属关系。

<a id="english"></a>

<details>
<summary><strong>English quick start</strong></summary>

### What it does

SciDiagram PPTX reconstructs scientific frameworks, mechanism diagrams, and research flowcharts as native editable PowerPoint shapes, text, and connectors. It preserves content and topology instead of redesigning the source or hiding it behind a full-slide bitmap.

For charts whose evidence is encoded by axes, scales, and data-driven geometry, use [SciPlot](https://github.com/SciToolsmith/sci-plot) instead.

### Install

```text
Use $skill-installer to install:
https://github.com/SciToolsmith/sci-diagram-pptx/tree/main/skills/sci-diagram-pptx
```

### Invoke

```text
$sci-diagram-pptx Reconstruct this scientific flowchart as a native editable PPTX.
Preserve all labels, formulas, hierarchy, directions, and connections. Do not redesign it.
```

### Editable examples

- [Global innovation-chain framework](https://github.com/SciToolsmith/sci-diagram-pptx/raw/main/examples/cases/global-innovation-chain-framework-editable.pptx)
- [Load-cycle mechanism](https://github.com/SciToolsmith/sci-diagram-pptx/raw/main/examples/cases/load-cycle-model-editable.pptx)
- [Adaptive FIR decomposition flowchart](https://github.com/SciToolsmith/sci-diagram-pptx/raw/main/examples/cases/adaptive-fir-mode-decomposition-flowchart-editable.pptx)

Public examples contain only the editable reconstruction slide; task source images and local notes are excluded.

</details>
