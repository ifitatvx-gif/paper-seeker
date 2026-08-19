const express = require("express");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  ExternalHyperlink,
  AlignmentType,
} = require("docx");

const PORT = process.env.PORT || 3000;
const OA = "https://api.openalex.org";
const MAILTO = "paper-query@example.com"; // OpenAlex 礼貌池标识

// ---------------- 常见中文领域 -> 英文概念名 ----------------
// 用于解决 OpenAlex 概念检索对中文支持较差的问题；同时前端会据此生成下拉建议。
const FIELD_DICT = {
  // 计算机 / 人工智能
  "人工智能": "Artificial intelligence",
  "机器学习": "Machine learning",
  "深度学习": "Deep learning",
  "计算机视觉": "Computer vision",
  "自然语言处理": "Natural language processing",
  "数据挖掘": "Data mining",
  "计算机科学": "Computer science",
  "机器人": "Robotics",
  "强化学习": "Reinforcement learning",
  "大数据": "Big data",
  "神经网络": "Artificial neural network",
  "计算机图形学": "Computer graphics",
  "人机交互": "Human-computer interaction",
  "信息安全": "Computer security",
  "软件工程": "Software engineering",
  "网络": "Computer network",
  // 生物 / 医学
  "生物信息学": "Bioinformatics",
  "生物化学": "Biochemistry",
  "分子生物学": "Molecular biology",
  "基因组学": "Genomics",
  "细胞生物学": "Cell biology",
  "遗传学": "Genetics",
  "微生物学": "Microbiology",
  "免疫学": "Immunology",
  "神经科学": "Neuroscience",
  "药理学": "Pharmacology",
  "肿瘤学": "Oncology",
  "心血管": "Cardiology",
  "公共卫生": "Public health",
  "流行病学": "Epidemiology",
  "临床医学": "Medicine",
  "医学影像": "Medical imaging",
  // 理化 / 材料
  "材料科学": "Materials science",
  "纳米技术": "Nanotechnology",
  "化学": "Chemistry",
  "有机化学": "Organic chemistry",
  "无机化学": "Inorganic chemistry",
  "物理化学": "Physical chemistry",
  "物理": "Physics",
  "量子力学": "Quantum mechanics",
  "凝聚态物理": "Condensed matter physics",
  "光学": "Optics",
  // 数学 / 统计 / 经管
  "数学": "Mathematics",
  "统计学": "Statistics",
  "经济学": "Economics",
  "金融": "Finance",
  "管理学": "Management",
  "心理学": "Psychology",
  "社会学": "Sociology",
  "教育学": "Education",
  // 工程 / 环境 / 地学
  "环境科学": "Environmental science",
  "生态学": "Ecology",
  "地理学": "Geography",
  "地球科学": "Earth science",
  "大气科学": "Atmospheric science",
  "遥感": "Remote sensing",
  "土木工程": "Civil engineering",
  "机械工程": "Mechanical engineering",
  "电气工程": "Electrical engineering",
  "电子工程": "Electronic engineering",
  "化学工程": "Chemical engineering",
  "生物医学工程": "Biomedical engineering",
  "航空航天": "Aerospace engineering",
  "能源": "Energy",
  "量子计算": "Quantum computing",
  "自动驾驶": "Autonomous driving",
  "推荐系统": "Recommender system",
  "知识图谱": "Knowledge graph",
  "语音识别": "Speech recognition",
  "图像处理": "Image processing",
  "区块链": "Blockchain",
  "物联网": "Internet of things",
  "密码学": "Cryptography",
  "半导体": "Semiconductor",
  "操作系统": "Operating system",
  "数据库": "Database",
  "通信": "Telecommunications",
  "控制工程": "Control engineering",
  "信号处理": "Signal processing",
  "优化": "Mathematical optimization",
  "计算生物学": "Computational biology",
  "药物研发": "Drug discovery",
  "海洋科学": "Oceanography",
  "地质学": "Geology",
  "天文学": "Astronomy",
};

// 英文简称 -> 中文领域（用于简称/模糊检索）
const FIELD_ALIASES = {
  ai: "人工智能",
  ml: "机器学习",
  dl: "深度学习",
  nlp: "自然语言处理",
  cv: "计算机视觉",
  rl: "强化学习",
  dm: "数据挖掘",
  cs: "计算机科学",
  hci: "人机交互",
  iot: "物联网",
  qc: "量子计算",
};

const REVIEW_TYPES = new Set(["review"]);
const PAPER_TYPES = new Set([
  "article",
  "conference-paper",
  "proceedings-article",
  "preprint",
  "book-chapter",
  "letter",
  "editorial",
  "report",
  "dissertation",
]);

function typeLabel(t) {
  return (
    {
      review: "综述",
      article: "期刊论文",
      "conference-paper": "会议论文",
      "proceedings-article": "会议论文",
      preprint: "预印本",
      "book-chapter": "书章节",
      letter: "快报",
      editorial: "社论",
      report: "报告",
      dissertation: "学位论文",
    }[t] || "论文"
  );
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// ---------------- OpenAlex 请求封装 ----------------
async function oaFetch(pathname, retries = 2) {
  let lastErr = null;
  for (let attempt = 0; attempt <= retries; attempt++) {
    const url =
      OA +
      pathname +
      (pathname.includes("?") ? "&" : "?") +
      "mailto=" +
      encodeURIComponent(MAILTO);
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 30000);
    try {
      const res = await fetch(url, {
        headers: {
          "User-Agent": "Academic-Paper-Search/1.0 (mailto:" + MAILTO + ")",
        },
        signal: ctrl.signal,
      });
      if (res.status === 429) {
        // 限流：退避后重试
        await sleep(800 * (attempt + 1));
        continue;
      }
      if (!res.ok) throw new Error("OpenAlex HTTP " + res.status);
      return await res.json();
    } catch (e) {
      lastErr = e;
      if (attempt < retries) await sleep(500 * (attempt + 1));
    } finally {
      clearTimeout(timer);
    }
  }
  throw lastErr || new Error("OpenAlex 请求失败");
}

function shortId(id) {
  return id ? id.split("/").pop() : null;
}

// ---------------- 大学解析 ----------------
function hasChinese(s) {
  return /[\u4e00-\u9fff]/.test(s || "");
}
function pickZhName(list) {
  if (!Array.isArray(list)) return null;
  for (const a of list) {
    if (a && hasChinese(a)) return a;
  }
  return null;
}

async function resolveInstitution(q) {
  const j = await oaFetch(
    "/institutions?search=" + encodeURIComponent(q) + "&per-page=5"
  );
  const r = j.results && j.results[0];
  if (!r) return null;
  return {
    id: r.id,
    shortId: shortId(r.id),
    display_name: r.display_name,
    display_name_zh: pickZhName(r.display_name_alternatives),
    country_code: r.country_code,
  };
}

// ---------------- 领域解析 ----------------
async function resolveConcept(q) {
  const j = await oaFetch(
    "/concepts?search=" + encodeURIComponent(q) + "&per-page=8"
  );
  const results = j.results || [];
  if (!results.length) return null;
  const exact = results.find(
    (c) => (c.display_name || "").toLowerCase() === q.toLowerCase()
  );
  const c = exact || results[0];
  return {
    id: c.id,
    shortId: shortId(c.id),
    display_name: c.display_name,
    level: c.level,
    works_count: c.works_count,
  };
}

function norm(s) {
  return String(s == null ? "" : s)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[·,，。.、;；:：\-—()（）[\]【】]/g, " ")
    .trim();
}

function lev(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  let prev = [];
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    const curr = [i];
    for (let j = 1; j <= n; j++) {
      curr[j] = Math.min(
        prev[j] + 1,
        curr[j - 1] + 1,
        prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
    }
    prev = curr;
  }
  return prev[n];
}

function fieldCandidates() {
  const cands = [];
  for (const k of Object.keys(FIELD_DICT)) cands.push({ zh: k, en: FIELD_DICT[k], alias: "" });
  for (const k of Object.keys(FIELD_ALIASES)) {
    const zh = FIELD_ALIASES[k];
    cands.push({ zh, en: FIELD_DICT[zh] || zh, alias: k });
  }
  for (const c of cands) {
    c.zhN = norm(c.zh);
    c.enN = norm(c.en);
    c.aliasN = norm(c.alias);
  }
  return cands;
}

function findFieldMatch(input) {
  const term = norm(input);
  if (!term) return null;
  const cands = fieldCandidates();
  let best = null;

  // 1. 精确匹配（中/英/简称）
  for (const c of cands) {
    if (c.zhN === term || c.enN === term || (c.aliasN && c.aliasN === term)) {
      return { zh: c.zh, en: c.en };
    }
  }
  // 2a. 前缀匹配（输入是词典条目的前缀）—— 按词典顺序取第一个（常见词在前）
  for (const c of cands) {
    if ((c.zhN && c.zhN.indexOf(term) === 0) || (c.enN && c.enN.indexOf(term) === 0)) {
      return { zh: c.zh, en: c.en };
    }
  }
  // 2b. 包含匹配（输入是词典条目的子串）—— 优先更短更精确
  best = null;
  for (const c of cands) {
    if (c.zhN && c.zhN.includes(term) && (!best || c.zhN.length < best.len)) best = { zh: c.zh, en: c.en, len: c.zhN.length };
    if (c.enN && c.enN.includes(term) && (!best || c.enN.length < best.len)) best = { zh: c.zh, en: c.en, len: c.enN.length };
  }
  if (best) return best;
  // 3. 反向包含（输入中嵌有词典条目）
  best = null;
  for (const c of cands) {
    if (c.zhN && term.includes(c.zhN) && (!best || c.zhN.length > best.len)) best = { zh: c.zh, en: c.en, len: c.zhN.length };
    if (c.enN && term.includes(c.enN) && (!best || c.enN.length > best.len)) best = { zh: c.zh, en: c.en, len: c.enN.length };
  }
  if (best) return best;
  // 4. 英文模糊（Levenshtein 距离，容忍拼写错误/字母换位；较长词放宽容差）
  if (/^[a-z0-9 .-]+$/.test(term) && term.length >= 4) {
    const tol = term.length >= 12 ? 3 : 2;
    best = null;
    for (const c of cands) {
      const d = lev(term, c.enN);
      if (d <= tol && (!best || d < best.d)) best = { zh: c.zh, en: c.en, d };
    }
    if (best) return best;
  }
  return null;
}

async function resolveField(input) {
  const term = (input || "").trim();
  if (!term) return { error: "请输入专业 / 领域" };
  const m = findFieldMatch(term);
  if (m && m.en) {
    const concept = await resolveConcept(m.en);
    if (concept) {
      const label = m.zh ? `${m.zh}（${m.en}）` : m.en;
      return { concept, label };
    }
  }
  // 未命中内置词典：英文直接交给 OpenAlex（其相关性/拼写容错较好）
  if (/^[\x20-\x7E]+$/.test(term)) {
    const concept2 = await resolveConcept(term);
    if (concept2) return { concept: concept2, label: concept2.display_name };
  }
  return { error: `未能识别领域「${term}」，请尝试更通用的中文或英文术语，或从下拉建议中选择` };
}

// ---------------- 作品抓取与整理 ----------------
async function fetchWorks(filter, perPage) {
  const j = await oaFetch(
    "/works?filter=" +
      encodeURIComponent(filter) +
      "&sort=cited_by_count:desc&per-page=" +
      perPage
  );
  return j.results || [];
}

function extractWork(x) {
  const doi = x.doi ? x.doi.replace("https://doi.org/", "") : null;
  const bestOa =
    (x.best_oa_location && x.best_oa_location.pdf_url) ||
    (x.primary_location && x.primary_location.pdf_url) ||
    null;
  const oaLanding =
    (x.best_oa_location && x.best_oa_location.landing_page_url) || null;
  const landing = x.primary_location && x.primary_location.landing_page_url;
  const source =
    x.primary_location &&
    x.primary_location.source &&
    x.primary_location.source.display_name
      ? x.primary_location.source.display_name
      : null;
  const authors = (x.authorships || [])
    .map((a) => a.author && a.author.display_name)
    .filter(Boolean)
    .slice(0, 8);
  return {
    id: x.id,
    title: x.title || "（无标题）",
    year: x.publication_year,
    date: x.publication_date,
    cited: x.cited_by_count || 0,
    doi: doi,
    doiUrl: doi ? "https://doi.org/" + doi : null,
    pdf: bestOa,
    oaUrl: oaLanding,
    url: x.id,
    landing: landing,
    source: source,
    type: x.type,
    typeLabel: typeLabel(x.type),
    authors: authors,
  };
}

async function mapLimit(items, limit, fn) {
  const out = new Array(items.length);
  let idx = 0;
  async function worker() {
    while (idx < items.length) {
      const i = idx++;
      out[i] = await fn(items[i], i);
    }
  }
  const workers = [];
  for (let k = 0; k < Math.min(limit, items.length); k++) workers.push(worker());
  await Promise.all(workers);
  return out;
}

// ---------------- 主查询流程 ----------------
async function runSearch(field, university) {
  const inst = await resolveInstitution(university);
  if (!inst)
    return { ok: false, error: `未找到大学「${university}」，请检查名称或尝试英文名` };

  const fr = await resolveField(field);
  if (fr.error) return { ok: false, error: fr.error };
  const concept = fr.concept;
  const instShort = inst.shortId;
  const conShort = concept.shortId;

  // 1) 抓取该机构在该领域被引最高的作品，并单独抓取综述与中文文献，用于识别主要研究人员
  const [works, reviewWorks, cnWorks] = await Promise.all([
    fetchWorks(`institutions.id:${instShort},concepts.id:${conShort}`, 100),
    fetchWorks(
      `institutions.id:${instShort},concepts.id:${conShort},type:review`,
      60
    ),
    fetchWorks(`institutions.id:${instShort},language:zh`, 30),
  ]);

  // 该领域代表性综述（独立于教授分组，直接呈现给用户）
  const topReviews = reviewWorks.slice(0, 10).map(extractWork);
  // 该机构的中文文献（仅保留标题为中文的）
  const topChinese = cnWorks
    .filter((w) => hasChinese(w.title))
    .slice(0, 15)
    .map(extractWork);

  // 2) 聚合作者（论文与综述分别统计）
  const authors = {};
  function addAuthor(a, kind, cited) {
    const insts = (a.institutions || [])
      .map((i) => shortId(i.id))
      .filter(Boolean);
    if (!insts.includes(instShort)) return;
    const aid = a.author && a.author.id;
    if (!aid) return;
    if (!authors[aid]) {
      authors[aid] = {
        id: aid,
        shortId: shortId(aid),
        name: a.author.display_name,
        orcid: a.author.orcid || null,
        affiliation: (a.raw_affiliation_strings && a.raw_affiliation_strings[0]) || "",
        papers: 0,
        reviews: 0,
        cited: 0,
      };
    }
    authors[aid][kind]++;
    authors[aid].cited += cited;
  }
  for (const w of works) {
    if (REVIEW_TYPES.has(w.type) || !PAPER_TYPES.has(w.type)) continue;
    for (const a of w.authorships || [])
      addAuthor(a, "papers", w.cited_by_count || 0);
  }
  for (const w of reviewWorks) {
    for (const a of w.authorships || [])
      addAuthor(a, "reviews", w.cited_by_count || 0);
  }

  // 综述数量少且被用户明确需要，故对综述作者加权，确保其进入结果
  const top = Object.values(authors)
    .sort(
      (x, y) =>
        y.papers + y.reviews * 2 - (x.papers + x.reviews * 2) ||
        y.cited - x.cited
    )
    .slice(0, 12);

  if (!top.length)
    return {
      ok: false,
      error:
        "未找到相关研究人员，请尝试更宽泛的领域，或核对大学名称是否正确",
    };

  // 3) 为每位研究人员抓取其在该领域的论文与综述
  const professors = await mapLimit(top, 4, async (a) => {
    await sleep(120); // 控制请求速率
    // 论文与综述分别查询：综述用 type:review 精确过滤，避免被高被引论文淹没
    const [paperWs, reviewWs, cnWs] = await Promise.all([
      fetchWorks(`author.id:${a.shortId},concepts.id:${conShort}`, 30),
      fetchWorks(`author.id:${a.shortId},concepts.id:${conShort},type:review`, 15),
      fetchWorks(`author.id:${a.shortId},language:zh`, 10),
    ]);
    const papers = [];
    const reviews = [];
    const chinese = cnWs
      .filter((w) => hasChinese(w.title))
      .slice(0, 8)
      .map(extractWork);
    for (const w of paperWs) {
      if (REVIEW_TYPES.has(w.type)) continue;
      if (PAPER_TYPES.has(w.type)) {
        if (papers.length < 15) papers.push(extractWork(w));
      }
    }
    for (const w of reviewWs) {
      if (REVIEW_TYPES.has(w.type)) {
        if (reviews.length < 12) reviews.push(extractWork(w));
      }
    }
    return {
      id: a.shortId,
      name: a.name,
      orcid: a.orcid,
      affiliation: a.affiliation,
      paperCount: papers.length,
      reviewCount: reviews.length,
      cited: a.cited,
      papers,
      reviews,
      chinese,
    };
  });

  return {
    ok: true,
    institution: inst,
    concept: concept,
    fieldLabel: fr.label,
    topReviews,
    topChinese,
    professors,
  };
}

// ---------------- 简单内存缓存 ----------------
const cache = new Map(); // key -> {ts, data}
function cacheGet(key) {
  const c = cache.get(key);
  if (c && Date.now() - c.ts < 10 * 60 * 1000) return c.data;
  return null;
}
function cacheSet(key, data) {
  if (cache.size > 200) cache.clear();
  cache.set(key, { ts: Date.now(), data });
}

// ---------------- Express 应用 ----------------
const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/fields", (req, res) => {
  const list = Object.keys(FIELD_DICT).map((zh) => ({
    zh,
    en: FIELD_DICT[zh],
  }));
  res.json({ ok: true, fields: list });
});

app.get("/api/search", async (req, res) => {
  const field = (req.query.field || "").trim();
  const university = (req.query.university || "").trim();
  if (!field || !university)
    return res.json({ ok: false, error: "请同时填写「大学名称」与「专业/领域」" });

  const key = field + "||" + university;
  const hit = cacheGet(key);
  if (hit) return res.json(hit);

  try {
    const data = await runSearch(field, university);
    if (data.ok) cacheSet(key, data);
    res.json(data);
  } catch (e) {
    res.json({ ok: false, error: "查询失败：" + e.message });
  }
});

// ---------------- Word 导出 ----------------
function tr(text, opts = {}) {
  return new TextRun(Object.assign({ text: String(text), font: "Microsoft YaHei" }, opts));
}

function workBlocks(w, idx) {
  const blocks = [];
  const titleText = `[${idx}] ${w.title}`;
  if (w.doiUrl) {
    blocks.push(
      new Paragraph({
        children: [
          new ExternalHyperlink({
            children: [
              tr(titleText, { color: "0563C1", underline: {}, size: 20 }),
            ],
            link: w.doiUrl,
          }),
        ],
      })
    );
  } else if (w.url) {
    blocks.push(
      new Paragraph({
        children: [
          new ExternalHyperlink({
            children: [
              tr(titleText, { color: "0563C1", underline: {}, size: 20 }),
            ],
            link: w.url,
          }),
        ],
      })
    );
  } else {
    blocks.push(new Paragraph({ children: [tr(titleText, { size: 20 })] }));
  }

  const meta = [];
  if (w.year) meta.push(String(w.year));
  if (w.source) meta.push(w.source);
  if (w.cited) meta.push("被引 " + w.cited);
  if (w.authors && w.authors.length) meta.push("作者: " + w.authors.join(", "));
  blocks.push(
    new Paragraph({
      children: [tr(meta.join(" · "), { size: 16, color: "595959" })],
    })
  );

  const links = [];
  if (w.doiUrl) links.push(w.doiUrl);
  if (w.pdf) links.push(w.pdf);
  if (links.length) {
    blocks.push(
      new Paragraph({
        children: [
          tr("链接: " + links.join("   "), { size: 16, color: "0563C1" }),
        ],
      })
    );
  }
  blocks.push(new Paragraph({ children: [tr("")] }));
  return blocks;
}

function buildDoc({ university, concept, fieldLabel, professors, topReviews, topChinese }) {
  const children = [];

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [tr("论文seeker · 论文查询报告", { bold: true, size: 40 })],
    })
  );
  children.push(
    new Paragraph({
      children: [
        tr(
          `大学：${university.display_name}    领域：${fieldLabel || concept.display_name}`,
          { size: 22, bold: true }
        ),
      ],
    })
  );
  children.push(
    new Paragraph({
      spacing: { after: 300 },
      children: [
        tr(
          `生成时间：${new Date().toLocaleString("zh-CN")}    研究人员：${professors.length} 位`,
          { size: 18, color: "595959" }
        ),
      ],
    })
  );

  if (topReviews && topReviews.length) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [tr("领域代表性综述")],
      })
    );
    topReviews.forEach((w, i) => children.push(...workBlocks(w, i + 1)));
  }

  if (topChinese && topChinese.length) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [tr("中文文献")],
      })
    );
    topChinese.forEach((w, i) => children.push(...workBlocks(w, i + 1)));
  }

  for (const p of professors) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [tr(p.name)],
      })
    );
    if (p.affiliation) {
      children.push(
        new Paragraph({
          children: [
            tr("所属机构/院系：" + p.affiliation, {
              italics: true,
              size: 18,
              color: "595959",
            }),
          ],
        })
      );
    }
    children.push(
      new Paragraph({
        spacing: { after: 200 },
        children: [
          tr(`论文 ${p.papers.length} 篇 · 综述 ${p.reviews.length} 篇`, {
            size: 18,
            color: "595959",
          }),
        ],
      })
    );

    if (p.papers.length) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [tr("论文")],
        })
      );
      p.papers.forEach((w, i) => children.push(...workBlocks(w, i + 1)));
    }
    if (p.reviews.length) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [tr("综述")],
        })
      );
      p.reviews.forEach((w, i) => children.push(...workBlocks(w, i + 1)));
    }
    if (p.chinese && p.chinese.length) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [tr("中文论文")],
        })
      );
      p.chinese.forEach((w, i) => children.push(...workBlocks(w, i + 1)));
    }
  }

  return new Document({ sections: [{ children }] });
}

app.post("/api/export-word", async (req, res) => {
  try {
    const { university, concept, fieldLabel, professors, topReviews, topChinese } = req.body || {};
    if (!university || !professors || !professors.length) {
      return res.status(400).json({ ok: false, error: "没有可导出的数据" });
    }
    const doc = buildDoc({ university, concept, fieldLabel, professors, topReviews, topChinese });
    const buf = await Packer.toBuffer(doc);
    const fname =
      "论文查询报告_" + (university.display_name || "结果") + ".docx";
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="report.docx"; filename*=UTF-8''${encodeURIComponent(fname)}`
    );
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(500).json({ ok: false, error: "导出失败：" + e.message });
  }
});

// 前端路由兜底
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`论文seeker 已启动: http://localhost:${PORT}`);
});
