// 智能体可调用的工具。目前只有一个：在 OpenAlex 上检索文献。
const OA = "https://api.openalex.org";
const MAILTO = "paper-query@example.com"; // OpenAlex 礼貌池标识，与 server.js 保持一致

// OpenAlex 的摘要是倒排索引 {word: [position, ...]}，还原为普通文本
function reconstructAbstract(invertedIndex) {
  if (!invertedIndex) return "";
  const words = [];
  for (const [word, positions] of Object.entries(invertedIndex)) {
    for (const pos of positions) words[pos] = word;
  }
  return words.filter(Boolean).join(" ");
}

const toolDefinitions = [
  {
    name: "search_works",
    description:
      "在 OpenAlex 学术数据库中按关键词检索文献（论文、综述等），按被引量降序返回。" +
      "每条结果包含标题、年份、被引次数、DOI、前几位作者和摘要片段。关键词请使用英文。",
    input_schema: {
      type: "object",
      properties: {
        query: { type: "string", description: "检索关键词，英文更有效" },
        limit: {
          type: "integer",
          minimum: 1,
          maximum: 10,
          description: "返回条数，默认 5",
        },
      },
      required: ["query"],
    },
  },
];

function makeTools({ fetchImpl = fetch } = {}) {
  return {
    definitions: toolDefinitions,

    async run(name, input = {}) {
      if (name !== "search_works") {
        throw new Error(`未知工具: ${name}`);
      }

      const query = String(input.query ?? "").trim();
      if (!query) throw new Error("query 不能为空");

      const parsed = parseInt(input.limit ?? 5, 10);
      const limit = Math.min(Math.max(Number.isNaN(parsed) ? 5 : parsed, 1), 10);

      const url =
        `${OA}/works?search=${encodeURIComponent(query)}` +
        `&per-page=${limit}&sort=cited_by_count:desc&mailto=${MAILTO}`;

      const res = await fetchImpl(url);
      if (!res.ok) throw new Error(`OpenAlex 请求失败 (${res.status})`);

      const data = await res.json();
      return (data.results || []).map((w) => ({
        title: w.display_name,
        year: w.publication_year,
        cited_by: w.cited_by_count,
        doi: w.doi,
        authors: (w.authorships || [])
          .slice(0, 5)
          .map((a) => a.author?.display_name)
          .filter(Boolean),
        abstract: reconstructAbstract(w.abstract_inverted_index).slice(0, 600),
      }));
    },
  };
}

module.exports = { makeTools, reconstructAbstract };
