# 会话记录 - 2026-03-30

## 今日目标

围绕 `D:\sourcePro\frontPro\myAgentTools` 这个新目录，从零开始搭一个本地 agent 工具项目。

本次对话重点是：

- 明确技术路线
- 先把大的架构搭起来
- 优先完成 PPT 主链路
- 再做一个可直接本地打开的桌面界面

## 今日关键决策

### 1. 主技术路线

最终采用：

- `Python` 作为 agent / tool 主语言
- `python-pptx` 作为 PPT 生成后端
- 桌面 GUI 使用 `Tkinter`

没有采用：

- Node/TS 作为主线
- 浏览器 Web UI 作为当前主界面
- C# WinForms/WPF 作为当前实现路径

原因：

- Python 更适合未来扩展多个本地工具
- 当前最重要的是先做稳定的 PPT 生成器
- 本地桌面界面可先用 Tkinter 快速成型
- 当前环境没有 `.NET SDK`，不适合立即做 C# GUI 编译

### 2. PPT 当前不接 LLM

这次已经明确：

- 当前本地项目的 PPT 后端 **没有接入 LLM**
- 只是一个 `DeckSpec -> PPTX` 的确定性引擎

这意味着：

- 当前还不支持“用户直接输入一句话，AI 自动生成 PPT”
- 当前界面主要是围绕结构化 spec 和预览来工作

### 3. 后续产品方向

已经确认的长期方向：

- 用户在 GUI 中直接输入需求
- agent 根据输入内容生成 PPT 大纲
- 用户确认或修改内容
- 再生成 PPT
- 后续还会支持搜索增强和其他来源

## 今日已完成的工程内容

### 1. 项目骨架

已创建：

- `pyproject.toml`
- `README.md`
- `docs/architecture.md`
- `examples/sample_deck.json`

### 2. 数据结构层

已完成：

- `src/my_agent_tools/specs.py`

主要定义：

- `DeckSpec`
- `SlideSpec`
- `BlockSpec`
- `MetricSpec`

### 3. PPT 渲染器

已完成：

- `src/my_agent_tools/tools/ppt.py`

当前可用能力：

- 从 JSON spec 生成 `.pptx`
- 支持基础页面类型
- 支持指标卡片页
- 支持表格页
- 支持双栏页
- 支持讲者备注
- 支持比例切换
- 支持主题配置入口

### 4. CLI

已完成：

- `src/my_agent_tools/cli.py`

可用命令：

- `render-ppt`
- `inspect-spec`

### 5. 示例与输出

已完成：

- `examples/sample_deck.json`
- `out/sample_deck.pptx`

### 6. 桌面 GUI

已完成：

- `src/my_agent_tools/desktop_app.py`
- `launch_gui.pyw`
- `launch_gui.bat`

当前 GUI 已实现：

- 中文界面
- 左侧页面列表
- 右侧页面预览
- JSON 编辑
- 打开 / 另存规格文件
- 校验规格
- 生成 PPT
- 打开输出目录

### 7. 测试

已完成：

- `tests/test_specs.py`
- `tests/test_ppt.py`

并已验证通过。

## 今日讨论后的核心认知

如果要把 PPT 真正做“完善”，不能只靠现在这个渲染器，还至少需要补下面这些层：

### 1. 需求理解层

用户自然语言输入 -> 结构化任务描述

### 2. 内容规划层

结构化任务描述 -> DeckSpec / 大纲 / 每页结构

### 3. 人工确认层

用户先确认大纲、资料和页面结构，再正式生成

### 4. 渲染增强层

包括：

- 模板映射
- 文本溢出处理
- 表格分页
- 图表支持
- 品牌样式控制

### 5. 校验修正层

包括：

- 结构校验
- 内容缺失校验
- 布局风险校验
- 自动重试或重新生成

### 6. 数据来源层

包括未来的：

- 用户直接输入
- 搜索增强
- 本地资料
- 其他文档来源

## 明天最推荐做什么

如果明天继续，建议先不要急着接搜索，也不要扩 MD/HTML。

建议优先推进：

### 第一优先级

继续完善 PPT 路线：

- 模板系统
- 页面布局策略
- 溢出和分页
- 图表支持
- 生成质量

### 第二优先级

把 GUI 从“编辑 JSON”升级到“面向普通用户输入需求”的形态：

- 需求输入区
- 大纲区
- 页面编辑区
- 生成区

### 第三优先级

再做 AI 规划器：

- 自然语言 -> 大纲
- 大纲 -> DeckSpec

## 明天继续时给助手的建议提示词

如果开新对话，建议直接发送：

```text
继续 D:\sourcePro\frontPro\myAgentTools 这个项目。
请先读取 docs/project-status.md 和 docs/session-notes-2026-03-30.md，
然后读取 docs/architecture.md、src/my_agent_tools/desktop_app.py、
src/my_agent_tools/tools/ppt.py、src/my_agent_tools/specs.py。
今天继续优先推进 PPT 完善路线，不要先做 MD/HTML。
```

## 当前风险和注意事项

### 1. GUI 只是第一版

它目前更适合开发者和调试，不适合最终普通用户直接使用。

### 2. 还没有 LLM 集成

项目当前不具备 AI 自动规划 PPT 的能力。

### 3. 还没有模板映射

当前 PPT 主要依赖代码布局，还没有真正接企业模板。

### 4. 还没有搜索增强

与联网检索相关的能力还没有进入本地项目实现。
