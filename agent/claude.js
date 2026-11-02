// 最小化的 Claude Messages API 客户端（无第三方依赖，使用 Node 内置 fetch）
const API_URL = process.env.ANTHROPIC_BASE_URL || "https://api.anthropic.com";
const API_VERSION = "2023-06-01";

async function createMessage({
  apiKey,
  model,
  system,
  messages,
  tools,
  maxTokens = 2048,
  fetchImpl = fetch,
}) {
  const res = await fetchImpl(`${API_URL}/v1/messages`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": API_VERSION,
    },
    body: JSON.stringify({
      model,
      max_tokens: maxTokens,
      system,
      messages,
      tools,
    }),
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const detail = data?.error?.message || res.statusText;
    throw new Error(`Anthropic API 请求失败 (${res.status}): ${detail}`);
  }
  return data;
}

module.exports = { createMessage };
