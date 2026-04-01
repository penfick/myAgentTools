# 项目状态文档

更新时间：2026-03-31

## 1. 项目目标

这是一个本地 agent 工具项目，目标是围绕“内容生成”建立一套可扩展的本地能力。

当前阶段的主目标：

- 先把 `PPT` 生成功能做扎实
- 提供本地桌面 GUI
- 后续再扩展到：
  - Markdown 生成
  - HTML / 前端设计代码生成
  - 搜索增强
  - 多来源资料整理后生成 PPT

## 2. 当前技术路线

目前已确定的技术路线如下：

- 主语言：`Python`
- 当前 PPT 渲染后端：`python-pptx`
- AI 规划后端：`OpenAI API + openai Python SDK`
- GUI：`Tkinter` 本地桌面界面
- 数据组织方式：统一的结构化规格模型

当前已经接入 LLM 作为“规划层”，但没有接入搜索增强。

这意味着当前 PPT 生成链路是：

`自然语言需求 -> AI 生成大纲 -> 用户确认/补充 -> AI 生成 DeckSpec -> python-pptx 渲染 -> .pptx 文件`

并且明确不是：

`自然语言 -> AI 不经确认直接生成最终 PPT`

## 3. 当前已完成内容

### 3.1 项目骨架

已建立基础项目结构，核心文件包括：

- `pyproject.toml`
- `README.md`
- `docs/architecture.md`
- `examples/sample_deck.json`

### 3.2 统一数据结构

核心数据结构在：

- `src/my_agent_tools/specs.py`

当前主要模型：

- `DeckSpec`
- `SlideSpec`
- `BlockSpec`
- `MetricSpec`

当前支持的 slide kind：

- `title`
- `section`
- `content`
- `two_column`
- `table`
- `image`
- `metrics`
- `chart`
- `closing`

### 3.3 PPT 渲染器

PPT 主渲染器在：

- `src/my_agent_tools/tools/ppt.py`

当前能力包括：

- 根据 `DeckSpec` 生成 `.pptx`
- 支持主题
- 支持 `16:9 / 4:3`
- 支持讲者备注
- 支持基础页面类型渲染
- 支持基础表格页、双栏页、指标页、结尾页
- 支持图表页
- 支持长文本自动续页
- 支持表格自动分页
- 支持页码 footer
- 支持模板配置文件驱动的布局选择
- 支持基于占位符索引的标题、正文、图片、页脚映射

### 3.4 命令行入口

CLI 在：

- `src/my_agent_tools/cli.py`

当前已支持：

- `render-ppt`
- `inspect-spec`
- `render-md`（预留）
- `render-html`（预留）
- `--template-config` 参数

### 3.5 桌面 GUI

桌面界面入口：

- `launch_gui.bat`
- `launch_gui.pyw`

桌面主程序：

- `src/my_agent_tools/desktop_app.py`

当前 GUI 已支持：

- 输入自然语言需求
- 配置模型、API Base URL 和 API Key
- 支持接口预设与本地保存模型配置
- 先生成大纲
- 当信息不足时向用户继续提问
- 用户补充信息后重生成大纲
- 用户确认大纲后再生成规格
- 在界面中显示待确认问题，并可一键回填答复模板
- 加载示例规格
- 打开规格文件
- 另存规格文件
- 校验规格
- 左侧页面列表
- 右侧页面预览
- JSON 编辑
- 一键生成 PPT
- 打开输出目录
- 支持 OpenAI 官方接口与 OpenAI 兼容第三方接口
- 优先走 Responses API，失败后自动回退到 Chat Completions

界面文案已做中文化。

### 3.6 示例输出

示例规格文件：

- `examples/sample_deck.json`
- `examples/default_template_config.json`

当前已生成的示例 PPT：

- `out/sample_deck.pptx`
- `out/sample_deck_v3.pptx`

### 3.7 测试与验证

当前已有测试：

- `tests/test_specs.py`
- `tests/test_ppt.py`

已验证通过：

- spec 可正常校验
- PPT 文件可正常写出
- 图表页可正常写出
- 分页逻辑可正常展开额外页面
- 模板配置可驱动默认 PowerPoint 布局和占位符映射
- GUI 模块可导入
- 代码可编译

## 4. 当前未完成但已预留的部分

当前以下能力只是预留接口，还未真正实现：

- `Markdown` 生成
- `HTML` 生成
- 搜索增强
- 自动校验与重试
- 图片 fit / crop 策略
- PDF 导出校验
- 企业品牌模板适配工具
- 基于真实文本测量的高精度布局拟合

## 5. 当前最重要的结论

所以项目当前状态是：

- 已有一个确定性的 PPT 引擎
- 已有一个确认式 AI 规划层
- 还没有搜索增强和自动研究能力

## 6. 推荐的后续开发顺序

建议严格按下面顺序推进：

### 阶段 A：把 PPT 渲染器做强

目标：

- 让不依赖 AI 的 PPT 渲染器先足够稳定和可用

重点内容：

- 模板系统
- 企业模板映射
- 文本溢出处理
- 表格分页
- 图表支持
- 页码 / Logo / 页脚
- 输出前校验

### 阶段 B：做 AI 大纲生成

目标：

- 用户输入自然语言需求
- LLM 只负责先生成 PPT 大纲和页面结构

重点内容：

- 需求解析
- 大纲生成
- 页数和页面类型规划
- 用户确认后再继续

### 阶段 C：做 AI 内容填充

目标：

- 大纲确定后，自动补每页内容

重点内容：

- 每页 bullet
- 每页摘要
- 图表建议
- 备注生成

### 阶段 D：做搜索增强

目标：

- 支持根据用户主题先检索资料，再生成 PPT

重点内容：

- 搜索结果汇总
- 用户勾选和修正
- 引用内容进入 PPT 规划链路

### 阶段 E：再扩展 MD / HTML

目标：

- 在 PPT 主链路稳定后，把相同思路扩展到其他输出格式

## 7. 明天继续时建议优先读取的文件

如果明天继续开发，建议优先读取这几个文件：

- `docs/project-status.md`
- `docs/session-notes-2026-03-30.md`
- `docs/architecture.md`
- `src/my_agent_tools/desktop_app.py`
- `src/my_agent_tools/tools/ppt.py`
- `src/my_agent_tools/specs.py`

## 8. 明天继续时的建议开场语

可以直接对助手说：

```text
继续 D:\sourcePro\frontPro\myAgentTools 这个项目。
请先读取 docs/project-status.md 和 docs/session-notes-2026-03-30.md，
再读取 docs/architecture.md、src/my_agent_tools/desktop_app.py、src/my_agent_tools/tools/ppt.py。
今天先接着推进 PPT 完善路线，不要先做 MD/HTML。
```
