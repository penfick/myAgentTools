# 公司 OA 知识库建设方案

如果是给公司 OA 接一个知识库，我建议默认按“内网知识库 + 权限过滤 + 审计日志”来设计，而不是先上一个公网聊天机器人。

## 推荐架构

- 文档和知识库放内网：制度、流程、合同模板、项目文档、FAQ 都先入内网知识库。
- OA 做统一入口：员工在 OA 里提问，调用内部 AI 服务，不直接碰外部平台页面。
- 检索先行：先从知识库里找相关片段，再把最少必要上下文交给模型。
- 权限前置：回答前先按用户的 OA 身份、部门、角色做文档过滤。
- 答案带来源：返回答案时附原文出处、文档名、更新时间。

## 三种部署方式

### 1. 全本地

文档解析、向量库、Embedding、Rerank、LLM 全在公司内网。
适合财务、人事、合同、客户资料、源码这类高敏数据。

### 2. 混合部署

知识库、向量库、权限、日志在内网；模型可接受控云 API 或私有模型服务。
这是最现实的企业方案，安全和成本比较平衡。

### 3. 云端为主

知识库和模型都交给外部平台。
上线快，但泄露风险、合规压力、供应商锁定都更高，不建议直接用于 OA 核心知识。

## 优先建议

- 如果文档大多是制度、手册、FAQ：`Dify` 或 `FastGPT` 自部署，先跑起来。
- 如果有大量扫描件、表格、合同、复杂 PDF：底层 RAG 更适合用 `RAGFlow`。
- 一个比较稳的组合是：
  - `OA`
  - `内部 AI 网关`
  - `Dify/FastGPT` 作为应用层
  - `Qdrant` 或 `pgvector` 作为向量库
  - `本地 Embedding + 本地 Rerank`
  - `本地 LLM` 或受控模型服务

## 安全上真正要控的点

- 文档原文不能随便出内网
- Embedding 接口不能把敏感片段发到外面
- Prompt 日志、聊天记录、缓存要可控
- 权限必须和 OA 账号体系打通
- 删除文档后，索引和缓存也要同步删除
- 敏感字段要支持脱敏和审计

## 实施顺序

1. 先选 50-100 份典型文档做 PoC。
2. 先验证命中率、答案准确率、引用是否正确。
3. 再接 OA 单点登录和权限体系。
4. 最后再决定大模型是不是必须全本地。

## 一句话结论

- 不是一定要“整套模型都本地”。
- 但对公司 OA 来说，知识库、权限、日志、检索链路最好先内网化。
- 如果内容敏感级别高，再把大模型也切到本地或私有化。

## 参考

截至 2026-03-30 的官方资料显示，`Dify` 和 `FastGPT` 都支持自部署，`FastGPT` 也有本地模型接入文档，`RAGFlow` 官方强调企业级 RAG、自托管和混合检索能力。

- Dify Docs: https://docs.dify.ai/en/introduction
- FastGPT Self-Host: https://doc.fastgpt.io/en/docs/self-host
- FastGPT Local Models: https://doc.fastgpt.io/en/docs/introduction/development/custom-models/ollama
- RAGFlow 官方站: https://ragflow.io/
- RAGFlow GitHub README: https://github.com/infiniflow/ragflow
