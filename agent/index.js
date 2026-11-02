#!/usr/bin/env node
// 命令行入口：node agent/index.js "你的研究问题"
const { createMessage } = require("./claude");
const { makeTools } = require("./tools");
const { runAgent } = require("./loop");

const SYSTEM_PROMPT = `你是学术文献调研助手。
- 收到研究问题后，使用 search_works 在 OpenAlex 上检索，可以换关键词多次检索。
- 只依据工具返回的内容作答，不要编造文献。
- 引用文献时给出标题、年份和 DOI。
- 如果检索结果不足以回答，请直接说明。
- 最终回答使用中文。`;

async function main() {
  const question = process.argv.slice(2).join(" ").trim();
  if (!question) {
    console.error('用法: node agent/index.js "你的研究问题"');
    process.exit(1);
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.error("请先设置环境变量 ANTHROPIC_API_KEY");
    process.exit(1);
  }

  const model = process.env.AGENT_MODEL || "claude-sonnet-5-5";
  const tools = makeTools();

  const result = await runAgent({
    question,
    system: SYSTEM_PROMPT,
    tools,
    maxSteps: 6,
    callModel: (args) => createMessage({ apiKey, model, ...args }),
    log: (msg) => console.error(`[agent] ${msg}`),
  });

  console.log(result.answer);
  console.error(`[agent] 完成，共 ${result.steps} 步`);
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
