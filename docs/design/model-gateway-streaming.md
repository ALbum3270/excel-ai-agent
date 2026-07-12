# 专项设计：Model Gateway（流式/非流式统一与调用双模）

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 评审通过（见 [`design-review-record.md`](./design-review-record.md)） |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §5.8 |
| 现有 | `AiGateway`（非流式）、`HttpStreamService`（流式+部分 MCP/ReAct）、`ReasoningRequestHelper`、`DefaultConversationRuntime`、`LLMUtil`、翻译等旁路 HTTP |
| 关联 | Harness Planner/Loop、Safety 超时、AppLogger 脱敏、smoke-ai-gateway-provider |

---

## 1. 目标与非目标

### 1.1 目标

1. **一个 Provider 适配层** 同时服务：流式聊天、非流式内部调用、Agent 规划/修复。  
2. 统一：鉴权、超时、重试、错误码、Reasoning、Anthropic/OpenAI 差异、日志脱敏。  
3. 支持两种 Agent 动作协议：  
   - **JSON Plan / JSON tool 指令**（现状主路径）  
   - **原生 tool_calls / function calling**（可选增强）  
4. 调用方只依赖 `IModelClient`，不再直接 `New HttpClient` 拼 body。  

### 1.2 非目标

- 不实现自建模型推理服务。  
- v1 不强制所有翻译/UDF 立刻迁完（分批），但禁止新增旁路。  
- 不做多模态视觉主路径（预留 messages 内容块）。  

### 1.3 原则

| 原则 | 说明 |
|---|---|
| 适配器模式 | Provider 差异关在 Gateway 内 |
| 流式是传输形态 | 与「是否 tool-calling」正交 |
| 密钥永不入日志 | Redact |
| 可取消 | CancellationToken 贯穿 |
| 可测 | BuildRequest 无网络单测（已有雏形） |

---

## 2. 现状与差距

| 组件 | 能力 | 问题 |
|---|---|---|
| AiGateway | OpenAI-compatible + Anthropic 非流式 | 无 stream；HttpClient 单例；调用面窄 |
| HttpStreamService | SSE 流 + UI 刷 + MCP/ReAct 循环 | Provider/业务/UI 耦合；与 Gateway 重复 |
| LLMUtil | 同步/异步工具 HTTP | 历史包袱 |
| 翻译等 | 自建 body | 行为不一致 |
| ReasoningRequestHelper | 推理参数 | 需两处共用 |

---

## 3. 目标架构

```text
                    ┌─────────────────────┐
   Planner/Loop     │   IModelClient      │
   Chat Surface     │  ChatAsync          │
   Memory Extract   │  ChatStreamAsync    │
   Intent/AutoComp  │  (cancel, headers)  │
                    └─────────┬───────────┘
                              │
                    ┌─────────▼───────────┐
                    │  ModelGateway       │
                    │  - build request    │
                    │  - auth / timeout   │
                    │  - retry / map err  │
                    │  - redact log       │
                    └─────────┬───────────┘
                              │
           ┌──────────────────┼──────────────────┐
           ▼                  ▼                  ▼
    OpenAICompatible    AnthropicMessages    (Future)
    /chat/completions   /v1/messages
    stream & non-stream
```

**HttpStreamService 演进**：降为「Stream UI Pump」——只消费 `IAsyncEnumerable<ModelStreamEvent>` 刷 WebView，不再拼 Provider body。

---

## 4. IModelClient 契约

### 4.1 请求

```json
{
  "requestId": "req_...",
  "purpose": "chat | plan | repair | intent | memory_extract | autocomplete | translate",
  "stream": false,
  "model": "gpt-x",
  "platform": "openai|anthropic|custom",
  "apiUrl": "https://...",
  "apiKey": "***",
  "reasoningMode": "default|enabled|disabled",
  "temperature": 0.2,
  "maxTokens": 4096,
  "timeoutSeconds": 120,
  "messages": [
    { "role": "system|user|assistant|tool", "content": "...", "name": null, "toolCallId": null }
  ],
  "tools": null,
  "toolChoice": "auto|none|required|{name}",
  "responseFormat": null,
  "metadata": { "turnId": "...", "runId": "..." }
}
```

| purpose | 默认 timeout | 默认 temperature 建议 |
|---|---|---|
| chat | 120 | 用户配置 |
| plan / repair | 90 | 0–0.3 |
| intent | 30 | 0 |
| memory_extract | 60 | 0 |
| autocomplete | 15 | 0 |
| translate | 120 | 0.3 |

### 4.2 非流式响应

```json
{
  "success": true,
  "content": "助手文本",
  "toolCalls": [
    { "id": "call_1", "name": "FormatText", "argumentsJson": "{...}" }
  ],
  "finishReason": "stop|tool_calls|length|error",
  "usage": { "promptTokens": 0, "completionTokens": 0 },
  "errorCode": "",
  "errorMessage": "",
  "raw": null,
  "statusCode": 200
}
```

`errorCode` 映射：`NETWORK_ERROR` / `TIMEOUT` / `AUTH_FAILED` / `RATE_LIMITED` / `PROVIDER_ERROR` / `PARSE_ERROR`。

### 4.3 流式事件

```text
ModelStreamEvent:
  type: content_delta | tool_call_delta | tool_call_finish | message_finish | error | usage
  textDelta?: string
  toolCall?: { id, name, argumentsDelta }
  errorCode?: string
  errorMessage?: string
```

UI Pump：

- content_delta → 刷 Markdown 缓冲（沿用 min chars / interval）  
- tool_call_* → 聚合完整 arguments → 交 Harness/旧 ReAct 桥  
- error → 用户可见文案  

---

## 5. Provider 适配

### 5.1 OpenAI-compatible

- 非流式：`POST chat/completions` `stream:false`  
- 流式：`stream:true` + SSE `data:` 解析  
- tools：`tools` / `tool_choice` 标准字段  
- Reasoning：`ReasoningRequestHelper.Apply*`  

### 5.2 Anthropic Messages

- URL 识别保持 `IsAnthropicEndpoint`  
- Header：`x-api-key` + `anthropic-version`  
- system 独立字段；messages 转换  
- 流式：`content_block_delta` 等事件映射到统一 ModelStreamEvent  
- tools：Anthropic tool 格式转换（实现阶段对齐官方）  

### 5.3 未知 Provider

- 默认按 OpenAI-compatible 尝试  
- 失败 error 可诊断  

---

## 6. 双模 Agent 动作协议

### 6.1 模式 A：JSON Plan（默认，兼容现状）

```text
模型输出文本/JSON
  → Loop 解析 toolId+params 或 plan steps
  → ToolBroker 执行
```

适用：任意兼容 chat 模型，无原生 tools。

### 6.2 模式 B：原生 Tool Calling（可选）

```text
请求带 tools schema（VisibleTools）
  → 响应 tool_calls
  → 直接转 ToolCall[]
  → 执行后 messages 追加 tool role
  → 再请求直到 stop
```

### 6.3 选择策略

| 条件 | 模式 |
|---|---|
| 配置 `model.toolCalling=off` | A |
| 配置 `auto` 且能力画像 supportsTools | B |
| Provider 不支持或失败 | 回退 A |
| purpose=plan 且需要结构化 plan | A 或 response_format json |

**Capability 画像（配置/探测）**

```json
{
  "supportsTools": true,
  "supportsStream": true,
  "supportsVision": false,
  "maxContextTokens": 128000
}
```

---

## 7. 可靠性

### 7.1 超时

- 请求级 `timeoutSeconds`  
- 流式：空闲超时（如 60s 无 delta）→ `TIMEOUT`  

### 7.2 重试

| 错误 | 重试 |
|---|---|
| 429 / 5xx | 指数退避，最多 2 次 |
| 超时 | 1 次（purpose≠autocomplete） |
| 4xx 鉴权 | 不重试 |
| 解析失败 | 不重试 body，可记 PARSE_ERROR |

### 7.3 取消

- `StopStream` / Harness Cancel → Cancel CTS  
- 保证 HttpClient 发送取消  

### 7.4 HttpClient

- 使用 `HttpClientPool` 或静态客户端 per base address  
- 禁止每次 `New HttpClient` 泄漏套接字  

---

## 8. 安全与日志

| 项 | 规则 |
|---|---|
| 日志 | method, url host, status, latency, purpose, turnId |
| 禁止 | apiKey、Authorization、完整 messages 默认 |
| Debug 开关 | `gateway.logMessages=true` 才记截断 messages |
| Redact | AppLogger.Redact |

---

## 9. 调用方迁移矩阵

| 调用方 | 现状 | 目标 |
|---|---|---|
| IntentRecognition | AiGateway | IModelClient.ChatAsync purpose=intent |
| LlmMemoryExtractor | AiGateway | purpose=memory_extract |
| Autocomplete | AiGateway | purpose=autocomplete |
| Agent plan/repair | 多路径 | purpose=plan/repair |
| 主聊天流 | HttpStreamService | ChatStreamAsync + UI Pump |
| 翻译 | 自建 | 分批 ChatAsync purpose=translate |
| ExcelDna UDF | LLMUtil Sync | 保留 Sync 桥，内部调 Gateway |
| 新代码 | — | **禁止**旁路 |

---

## 10. 与 Harness 的集成

```text
Planner:
  client.ChatAsync(purpose=plan, messages, tools?=visible)

Repair:
  client.ChatAsync(purpose=repair, ...)

Chat mode:
  client.ChatStreamAsync(purpose=chat, ...)
    → UI pump
    → (optional) legacy tool_call loop 移出到 Harness 统一
```

**重要**：MCP/ReAct 循环应从 HttpStreamService **上移到 Harness/Loop**，Gateway 只做模型 IO。否则流式层继续膨胀。

---

## 11. 配置项建议

| 键 | 默认 | 说明 |
|---|---|---|
| gateway.toolCalling | auto | off/auto/on |
| gateway.maxRetries | 2 | |
| gateway.streamIdleTimeoutSec | 60 | |
| gateway.defaultTimeoutSec | 120 | |
| gateway.logMessages | false | |

模型 URL/Key/Name 仍来自 `ConfigSettings`。

---

## 12. 测试

| 层 | 内容 |
|---|---|
| 单测 | BuildProviderRequest OpenAI/Anthropic（扩展现有 smoke） |
| 单测 | 流事件解析夹具 |
| 单测 | 错误码映射 |
| 集成 | 可选真网 purpose=intent 短调用 |
| 回归 | 聊天流式 UI 不卡顿；取消有效 |

---

## 13. 验收标准

1. 新调用方 100% 走 IModelClient。  
2. 同一套 Reasoning 在 stream/non-stream 行为一致。  
3. 日志无 sk-/Bearer。  
4. 取消流式请求后不再刷 UI delta。  
5. toolCalling=off 时 Agent 仍可用 JSON Plan。  
6. Anthropic 非流式与流式均可测（有 key 环境）。  

---

## 14. 开放问题

| # | 问题 | 建议默认 |
|---|---|---|
| Q1 | 是否保留 AiGateway 类名？ | 可作 IModelClient 静态门面 |
| Q2 | 流式 tool_calls 与 UI 混排？ | 工具轮不展示原始 JSON，显示「正在调用 XX」 |
| Q3 | 多模态图片消息？ | v2 |
| Q4 | 本地代理/抓包？ | 尊重系统代理 |

---

## 15. 落地顺序

1. 抽出 `IModelClient` + 请求/响应 DTO  
2. 用 AiGateway 实现非流式  
3. 实现流式解析器 + HttpStreamService 改为 Pump  
4. Plan/Repair 迁 purpose  
5. 翻译分批迁  
6. 原生 tool calling 可选开关  
7. 删除重复 body 构建  

---

## 16. 决策摘要（评审勾选）

- [x] 同意统一 IModelClient  
- [x] 同意 HttpStreamService 降级为 UI Pump（**D10**）  
- [x] 同意 MCP/ReAct 上移 Harness（**D10**）  
- [x] 同意默认 JSON Plan，tool calling 为 auto（**D11**）  
- [x] 同意分批迁移翻译/UDF，禁止新增旁路  

开放问题冻结：Q1 可保留 AiGateway 作门面；Q2 工具轮 UI 显示「正在调用」；Q3 视觉 v2；Q4 尊重系统代理。

---

*Gateway 是「大脑的网线」：统一后，Harness 才能稳定做 plan/repair，而不被各处 HTTP 细节拖垮。*
