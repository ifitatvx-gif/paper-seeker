// 离线测试：不访问网络，全部使用模拟的模型和 HTTP 响应
const { test } = require("node:test");
const assert = require("node:assert/strict");
const { runAgent } = require("./loop");
const { makeTools, reconstructAbstract } = require("./tools");

// 按顺序返回预设的模型响应
function scriptedModel(responses) {
  let i = 0;
  return async () => {
    if (i >= responses.length) throw new Error("模型被调用次数超出预设");
    return responses[i++];
  };
}

const toolUse = (id, name, input) => ({
  type: "tool_use",
  id,
  name,
  input,
});

test("循环：先调用工具，再给出最终回答", async () => {
  const calls = [];
  const tools = {
    definitions: [],
    run: async (name, input) => {
      calls.push({ name, input });
      return [{ title: "Paper A", year: 2020 }];
    },
  };

  const result = await runAgent({
    question: "测试问题",
    system: "sys",
    tools,
    callModel: scriptedModel([
      {
        stop_reason: "tool_use",
        content: [toolUse("t1", "search_works", { query: "graph" })],
      },
      {
        stop_reason: "end_turn",
        content: [{ type: "text", text: "找到了 Paper A。" }],
      },
    ]),
  });

  assert.equal(result.answer, "找到了 Paper A。");
  assert.equal(result.steps, 2);
  assert.deepEqual(calls, [{ name: "search_works", input: { query: "graph" } }]);

  // 第二次调用模型时，消息里应包含工具结果
  const toolResultMsg = result.messages[2];
  assert.equal(toolResultMsg.role, "user");
  assert.equal(toolResultMsg.content[0].type, "tool_result");
  assert.equal(toolResultMsg.content[0].tool_use_id, "t1");
});

test("工具报错时，错误作为 tool_result 返回给模型，而不是崩溃", async () => {
  const tools = {
    definitions: [],
    run: async () => {
      throw new Error("网络超时");
    },
  };

  const result = await runAgent({
    question: "q",
    system: "sys",
    tools,
    callModel: scriptedModel([
      {
        stop_reason: "tool_use",
        content: [toolUse("t1", "search_works", { query: "x" })],
      },
      {
        stop_reason: "end_turn",
        content: [{ type: "text", text: "检索失败，无法回答。" }],
      },
    ]),
  });

  const errResult = result.messages[2].content[0];
  assert.equal(errResult.is_error, true);
  assert.match(errResult.content, /网络超时/);
  assert.equal(result.answer, "检索失败，无法回答。");
});

test("超过最大步数时抛出错误", async () => {
  const alwaysTool = async () => ({
    stop_reason: "tool_use",
    content: [toolUse("t", "search_works", { query: "x" })],
  });

  await assert.rejects(
    runAgent({
      question: "q",
      system: "sys",
      tools: { definitions: [], run: async () => [] },
      callModel: alwaysTool,
      maxSteps: 3,
    }),
    /3 步内未能完成/
  );
});

test("search_works 解析 OpenAlex 响应并还原摘要", async () => {
  const fakeFetch = async (url) => {
    assert.match(url, /search=graph%20neural/);
    assert.match(url, /per-page=2/);
    return {
      ok: true,
      json: async () => ({
        results: [
          {
            display_name: "GNN Survey",
            publication_year: 2021,
            cited_by_count: 900,
            doi: "https://doi.org/10.1/abc",
            authorships: [{ author: { display_name: "Alice" } }],
            abstract_inverted_index: { Graphs: [0], are: [1], useful: [2] },
          },
        ],
      }),
    };
  };

  const tools = makeTools({ fetchImpl: fakeFetch });
  const out = await tools.run("search_works", { query: "graph neural", limit: 2 });

  assert.equal(out.length, 1);
  assert.equal(out[0].title, "GNN Survey");
  assert.equal(out[0].abstract, "Graphs are useful");
  assert.deepEqual(out[0].authors, ["Alice"]);
});

test("reconstructAbstract 能处理空值与乱序位置", () => {
  assert.equal(reconstructAbstract(undefined), "");
  assert.equal(reconstructAbstract({ b: [1], a: [0] }), "a b");
});

test("未知工具名会被拒绝", async () => {
  const tools = makeTools({ fetchImpl: async () => ({}) });
  await assert.rejects(tools.run("delete_all", {}), /未知工具/);
});
