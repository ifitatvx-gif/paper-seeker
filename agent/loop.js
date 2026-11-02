// 智能体主循环：让模型思考 → 调用工具 → 把结果喂回去 → 直到模型给出最终回答。

async function runAgent({
  question,
  callModel,   // async ({ system, messages, tools }) => Anthropic 响应
  tools,       // { definitions, run(name, input) }
  system,
  maxSteps = 6,
  log = () => {},
}) {
  const messages = [{ role: "user", content: question }];

  for (let step = 1; step <= maxSteps; step++) {
    const response = await callModel({
      system,
      messages,
      tools: tools.definitions,
    });

    messages.push({ role: "assistant", content: response.content });

    const toolUses = response.content.filter((b) => b.type === "tool_use");

    // 模型没有再请求工具：这就是最终回答
    if (response.stop_reason !== "tool_use" || toolUses.length === 0) {
      const answer = response.content
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("\n")
        .trim();
      return { answer, steps: step, messages };
    }

    // 执行本轮所有工具调用；出错时把错误信息告诉模型，而不是直接崩溃
    const results = [];
    for (const use of toolUses) {
      log(`第 ${step} 步 调用工具 ${use.name} ${JSON.stringify(use.input)}`);
      try {
        const output = await tools.run(use.name, use.input);
        results.push({
          type: "tool_result",
          tool_use_id: use.id,
          content: JSON.stringify(output),
        });
      } catch (err) {
        results.push({
          type: "tool_result",
          tool_use_id: use.id,
          content: `错误: ${err.message}`,
          is_error: true,
        });
      }
    }

    messages.push({ role: "user", content: results });
  }

  throw new Error(`智能体在 ${maxSteps} 步内未能完成任务`);
}

module.exports = { runAgent };
