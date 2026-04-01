# My Agent Tools

本项目是一个以 `PPT` 生成为核心的本地 Agent 工具链。当前路线很明确：

- 先把 `PPT` 生成做扎实
- 用本地可控工具完成渲染
- 用 `LLM` 负责规划与确认，不直接生成最终文件
- 后续再扩展 `MD / HTML / 搜索增强`

## Current Focus

当前已经可用的主链路：

`用户需求 -> AI 生成大纲 -> 用户补充/确认 -> AI 生成 DeckSpec -> 本地渲染 PPT`

当前能力包括：

- `python-pptx` 本地渲染 `.pptx`
- `DeckSpec` / `SlideSpec` / `BlockSpec` 结构化模型
- 图表、表格、双栏、指标、封面、章节、结尾等页面类型
- 文本续页、表格分页、模板映射
- 本地桌面 GUI
- 第三方 OpenAI 兼容 `API Base URL / API Key`
- `Responses API` 优先，失败后自动回退 `Chat Completions`
- 确认式 AI 工作流，不允许一键盲生成最终 PPT

## Project Layout

```text
myAgentTools/
  docs/                 项目文档与会话记录
  examples/             示例 JSON 与模板配置
  src/my_agent_tools/   主代码
    tools/ppt.py        PPT 渲染器
    desktop_app.py      本地桌面界面
    openai_planner.py   LLM 规划层
    specs.py            共享数据结构
  tests/                pytest 测试
  launch_gui.pyw        GUI 启动入口
```

## Quick Start

安装依赖：

```powershell
pip install -e .[dev]
```

运行测试：

```powershell
python -m pytest
```

生成示例 PPT：

```powershell
python -m my_agent_tools.cli render-ppt --spec .\examples\sample_deck.json --output .\out\sample_deck.pptx
```

启动桌面界面：

```powershell
pythonw .\launch_gui.pyw
```

或者直接双击 `launch_gui.bat`。

## Desktop Workflow

桌面界面当前推荐这样用：

1. 在 `AI 规划` 页填写“用户需求”
2. 配置 `Model`、`API Base URL`、`API Key`
3. 点击“生成大纲”
4. 如果 AI 提出追问，在“待确认问题”查看，并把答复写入“补充说明 / 用户修正”
5. 点击“结合补充重生成大纲”
6. 确认大纲后，点击“确认大纲并生成规格”
7. 检查 `DeckSpec JSON`
8. 最后再点击“生成 PPT”

## Architecture Rules

- `PPT` 是当前最高优先级，不先扩散到 `MD / HTML`
- `LLM` 只负责规划，不直接生成最终 PPT 文件
- AI 输出必须先过结构化校验，再进入渲染器
- 布局逻辑必须在本地代码里，不放在 prompt 里赌结果
- GUI 流程必须保持“先确认，后生成”

## Tests

当前测试覆盖：

- schema 校验
- PPT 渲染
- planner 回退逻辑
- desktop settings 持久化

运行：

```powershell
python -m pytest
```

## Notes

- 本地敏感配置在 `config/desktop_settings.json`
- 如果勾选“记住 API Key”，该文件会以本地明文方式保存 key
- 生成产物默认输出到 `out/`

## Next

下一阶段优先做：

- 更好的确认式 AI 交互
- 更稳定的 PPT 排版质量
- 更强的模板适配能力
