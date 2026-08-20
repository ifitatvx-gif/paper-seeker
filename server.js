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
  "云计算": "Cloud computing",
  "边缘计算": "Edge computing",
  "分布式系统": "Distributed computing",
  "并行计算": "Parallel computing",
  "编程语言": "Programming language",
  "编译器": "Compiler",
  "计算机体系结构": "Computer architecture",
  "虚拟现实": "Virtual reality",
  "增强现实": "Augmented reality",
  "多媒体": "Multimedia",
  "嵌入式系统": "Embedded system",
  "集成电路": "Integrated circuit",
  "微电子": "Microelectronics",
  "光电子": "Optoelectronics",
  "电磁场": "Electromagnetic field",
  "微波": "Microwave",
  "雷达": "Radar",
  "天线": "Antenna (electronics)",
  "模式识别": "Pattern recognition",
  "智能控制": "Intelligent control",
  "车辆工程": "Automotive engineering",
  "机械制造": "Manufacturing engineering",
  "精密仪器": "Instrumentation",
  "金属材料": "Metallurgy",
  "高分子材料": "Polymer science",
  "复合材料": "Composite material",
  "热能工程": "Thermal engineering",
  "动力工程": "Power engineering",
  "核能": "Nuclear power",
  "新能源": "Renewable energy",
  "电力系统": "Electric power system",
  "电机": "Electric motor",
  "高电压": "High voltage",
  "建筑学": "Architecture",
  "城乡规划": "Urban planning",
  "结构工程": "Structural engineering",
  "岩土工程": "Geotechnical engineering",
  "桥梁工程": "Bridge engineering",
  "水利工程": "Hydraulic engineering",
  "环境工程": "Environmental engineering",
  "水处理": "Water treatment",
  "大气污染": "Air pollution",
  "交通运输": "Transportation engineering",
  "轨道交通": "Rail transport",
  "船舶与海洋工程": "Naval architecture",
  "制药工程": "Pharmaceutical engineering",
  "食品科学": "Food science",
  "食品工程": "Food engineering",
  "安全工程": "Safety engineering",
  "矿业工程": "Mining engineering",
  "石油工程": "Petroleum engineering",
  "地质工程": "Engineering geology",
  "应用数学": "Applied mathematics",
  "计算数学": "Computational mathematics",
  "概率论": "Probability theory",
  "运筹学": "Operations research",
  "理论物理": "Theoretical physics",
  "粒子物理": "Particle physics",
  "原子分子物理": "Atomic physics",
  "声学": "Acoustics",
  "等离子体": "Plasma (physics)",
  "分析化学": "Analytical chemistry",
  "高分子化学": "Polymer chemistry",
  "植物学": "Botany",
  "动物学": "Zoology",
  "生理学": "Physiology",
  "发育生物学": "Developmental biology",
  "进化生物学": "Evolutionary biology",
  "生物物理": "Biophysics",
  "自然地理": "Physical geography",
  "人文地理": "Human geography",
  "地理信息系统": "Geographic information system",
  "地球物理": "Geophysics",
  "土壤学": "Soil science",
  "农学": "Agronomy",
  "植物保护": "Plant protection",
  "园艺": "Horticulture",
  "动物科学": "Animal science",
  "兽医学": "Veterinary medicine",
  "林学": "Forestry",
  "水产": "Aquaculture",
  "农业工程": "Agricultural engineering",
  "基础医学": "Biomedicine",
  "口腔医学": "Dentistry",
  "中医学": "Traditional Chinese medicine",
  "药学": "Pharmacy",
  "护理学": "Nursing",
  "预防医学": "Preventive medicine",
  "内科学": "Internal medicine",
  "外科学": "Surgery",
  "儿科学": "Pediatrics",
  "妇产科学": "Obstetrics and gynaecology",
  "眼科学": "Ophthalmology",
  "耳鼻喉": "Otolaryngology",
  "皮肤病": "Dermatology",
  "神经病学": "Neurology",
  "精神病学": "Psychiatry",
  "麻醉": "Anesthesiology",
  "病理学": "Pathology",
  "康复医学": "Rehabilitation medicine",
  "会计学": "Accounting",
  "市场营销": "Marketing",
  "工商管理": "Business administration",
  "人力资源": "Human resource management",
  "国际贸易": "International trade",
  "产业经济": "Industrial organization",
  "区域经济": "Regional economics",
  "数量经济": "Econometrics",
  "财政学": "Public finance",
  "保险学": "Insurance",
  "旅游管理": "Tourism management",
  "物流管理": "Logistics",
  "电子商务": "E-commerce",
  "信息管理": "Information management",
  "法学": "Law",
  "政治学": "Political science",
  "马克思主义": "Marxism",
  "历史学": "History",
  "考古学": "Archaeology",
  "哲学": "Philosophy",
  "逻辑学": "Logic",
  "伦理学": "Ethics",
  "新闻传播": "Communication",
  "语言学": "Linguistics",
  "中国语言文学": "Chinese literature",
  "外国语言文学": "Foreign language",
  "图书情报": "Library science",
  "艺术学": "Art",
  "音乐": "Music",
  "美术": "Fine art",
  "设计学": "Design",
  "戏剧": "Theatre",
  "影视": "Film",
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
  gis: "地理信息系统",
  bigdata: "大数据",
  econ: "经济学",
  se: "软件工程",
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

// 常见高校：中文全称/简称/英文缩写 -> 英文全称（用于学校模糊检索）
const UNIV_ALIASES = {
  "清华大学": "Tsinghua University", "清华": "Tsinghua University", "tsinghua": "Tsinghua University", "thu": "Tsinghua University",
  "北京大学": "Peking University", "北大": "Peking University", "pku": "Peking University",
  "复旦大学": "Fudan University", "复旦": "Fudan University", "fudan": "Fudan University",
  "上海交通大学": "Shanghai Jiao Tong University", "上海交大": "Shanghai Jiao Tong University", "上交": "Shanghai Jiao Tong University", "sjtu": "Shanghai Jiao Tong University",
  "浙江大学": "Zhejiang University", "浙大": "Zhejiang University", "zju": "Zhejiang University",
  "南京大学": "Nanjing University", "南大": "Nanjing University", "nju": "Nanjing University",
  "中国科学技术大学": "University of Science and Technology of China", "中科大": "University of Science and Technology of China", "中国科技大学": "University of Science and Technology of China", "ustc": "University of Science and Technology of China",
  "哈尔滨工业大学": "Harbin Institute of Technology", "哈工大": "Harbin Institute of Technology", "hit": "Harbin Institute of Technology",
  "西安交通大学": "Xi'an Jiaotong University", "西安交大": "Xi'an Jiaotong University", "西交": "Xi'an Jiaotong University", "xjtu": "Xi'an Jiaotong University",
  "武汉大学": "Wuhan University", "武大": "Wuhan University", "whu": "Wuhan University",
  "华中科技大学": "Huazhong University of Science and Technology", "华科": "Huazhong University of Science and Technology", "hust": "Huazhong University of Science and Technology",
  "中山大学": "Sun Yat-sen University", "中山": "Sun Yat-sen University", "sysu": "Sun Yat-sen University",
  "同济大学": "Tongji University", "同济": "Tongji University",
  "东南大学": "Southeast University", "东南": "Southeast University", "seu": "Southeast University",
  "北京航空航天大学": "Beihang University", "北航": "Beihang University", "buaa": "Beihang University",
  "北京理工大学": "Beijing Institute of Technology", "北理工": "Beijing Institute of Technology", "bit": "Beijing Institute of Technology",
  "四川大学": "Sichuan University", "川大": "Sichuan University", "scu": "Sichuan University",
  "天津大学": "Tianjin University", "天大": "Tianjin University", "tju": "Tianjin University",
  "南开大学": "Nankai University", "南开": "Nankai University",
  "山东大学": "Shandong University", "山大": "Shandong University", "sdu": "Shandong University",
  "厦门大学": "Xiamen University", "厦大": "Xiamen University", "xmu": "Xiamen University",
  "吉林大学": "Jilin University", "吉大": "Jilin University", "jlu": "Jilin University",
  "中南大学": "Central South University", "中南": "Central South University", "csu": "Central South University",
  "湖南大学": "Hunan University", "湖大": "Hunan University", "hnu": "Hunan University",
  "重庆大学": "Chongqing University", "重大": "Chongqing University", "cqu": "Chongqing University",
  "电子科技大学": "University of Electronic Science and Technology of China", "电子科大": "University of Electronic Science and Technology of China", "uestc": "University of Electronic Science and Technology of China",
  "西安电子科技大学": "Xidian University", "西电": "Xidian University",
  "北京邮电大学": "Beijing University of Posts and Telecommunications", "北邮": "Beijing University of Posts and Telecommunications", "bupt": "Beijing University of Posts and Telecommunications",
  "南京航空航天大学": "Nanjing University of Aeronautics and Astronautics", "南航": "Nanjing University of Aeronautics and Astronautics", "nuaa": "Nanjing University of Aeronautics and Astronautics",
  "南京理工大学": "Nanjing University of Science and Technology", "南理工": "Nanjing University of Science and Technology",
  "北京交通大学": "Beijing Jiaotong University", "北交": "Beijing Jiaotong University",
  "西南交通大学": "Southwest Jiaotong University", "西南交大": "Southwest Jiaotong University",
  "北京师范大学": "Beijing Normal University", "北师大": "Beijing Normal University", "bnu": "Beijing Normal University",
  "华东师范大学": "East China Normal University", "华东师大": "East China Normal University", "华师大": "East China Normal University", "ecnu": "East China Normal University",
  "华中师范大学": "Central China Normal University", "华中师大": "Central China Normal University",
  "东北师范大学": "Northeast Normal University", "东北师大": "Northeast Normal University",
  "中国人民大学": "Renmin University of China", "人大": "Renmin University of China", "ruc": "Renmin University of China",
  "上海财经大学": "Shanghai University of Finance and Economics", "上财": "Shanghai University of Finance and Economics",
  "中央财经大学": "Central University of Finance and Economics", "央财": "Central University of Finance and Economics",
  "对外经济贸易大学": "University of International Business and Economics", "对外经贸": "University of International Business and Economics", "uibe": "University of International Business and Economics",
  "中国政法大学": "China University of Political Science and Law", "中国政法": "China University of Political Science and Law", "法大": "China University of Political Science and Law",
  "中国传媒大学": "Communication University of China", "中传": "Communication University of China",
  "中国农业大学": "China Agricultural University", "中农": "China Agricultural University", "中国农大": "China Agricultural University",
  "华中农业大学": "Huazhong Agricultural University", "华农": "Huazhong Agricultural University",
  "南京农业大学": "Nanjing Agricultural University", "南农": "Nanjing Agricultural University",
  "中国海洋大学": "Ocean University of China", "海大": "Ocean University of China",
  "中国石油大学": "China University of Petroleum", "中石油": "China University of Petroleum",
  "中国矿业大学": "China University of Mining and Technology", "中国矿业": "China University of Mining and Technology",
  "中国地质大学": "China University of Geosciences", "中国地质": "China University of Geosciences",
  "河海大学": "Hohai University", "河海": "Hohai University",
  "西北工业大学": "Northwestern Polytechnical University", "西工大": "Northwestern Polytechnical University", "nwpu": "Northwestern Polytechnical University",
  "兰州大学": "Lanzhou University", "兰大": "Lanzhou University", "lzu": "Lanzhou University",
  "东北大学": "Northeastern University, China",
  "大连理工大学": "Dalian University of Technology", "大连理工": "Dalian University of Technology", "大工": "Dalian University of Technology", "dlut": "Dalian University of Technology",
  "华南理工大学": "South China University of Technology", "华南理工": "South China University of Technology", "华工": "South China University of Technology", "scut": "South China University of Technology",
  "西南财经大学": "Southwestern University of Finance and Economics", "西财": "Southwestern University of Finance and Economics",
  "国防科技大学": "National University of Defense Technology", "国防科大": "National University of Defense Technology",
  "苏州大学": "Soochow University", "苏大": "Soochow University",
  "郑州大学": "Zhengzhou University", "郑大": "Zhengzhou University",
  "云南大学": "Yunnan University", "云大": "Yunnan University",
  "上海大学": "Shanghai University", "上大": "Shanghai University",
  "深圳大学": "Shenzhen University", "深大": "Shenzhen University",
  "南方科技大学": "Southern University of Science and Technology", "南科大": "Southern University of Science and Technology",
  "华东理工大学": "East China University of Science and Technology", "华理": "East China University of Science and Technology",
  "北京科技大学": "University of Science and Technology Beijing", "北科大": "University of Science and Technology Beijing",
  "北京工业大学": "Beijing University of Technology", "北工大": "Beijing University of Technology",
  "上海科技大学": "ShanghaiTech University", "上科大": "ShanghaiTech University",
  "首都医科大学": "Capital Medical University", "首医": "Capital Medical University",
  "南京医科大学": "Nanjing Medical University", "南医": "Nanjing Medical University",
  "麻省理工": "Massachusetts Institute of Technology", "mit": "Massachusetts Institute of Technology",
  "哈佛": "Harvard University", "harvard": "Harvard University",
  "斯坦福": "Stanford University", "stanford": "Stanford University",
  "剑桥": "University of Cambridge", "cambridge": "University of Cambridge",
  "牛津": "University of Oxford", "oxford": "University of Oxford",
  "伯克利": "University of California, Berkeley", "uc berkeley": "University of California, Berkeley", "berkeley": "University of California, Berkeley",
  "加州理工": "California Institute of Technology", "caltech": "California Institute of Technology",
  "普林斯顿": "Princeton University", "princeton": "Princeton University",
  "耶鲁": "Yale University", "yale": "Yale University",
  "哥伦比亚": "Columbia University", "columbia": "Columbia University",
  "康奈尔": "Cornell University", "cornell": "Cornell University",
  "卡内基梅隆": "Carnegie Mellon University", "cmu": "Carnegie Mellon University",
  "苏黎世联邦理工": "ETH Zurich", "eth": "ETH Zurich",
  "帝国理工": "Imperial College London", "imperial": "Imperial College London",
  "伦敦大学学院": "University College London", "ucl": "University College London",
  "多伦多": "University of Toronto", "toronto": "University of Toronto",
  "新加坡国立": "National University of Singapore", "nus": "National University of Singapore",
  "南洋理工": "Nanyang Technological University", "ntu": "Nanyang Technological University",
  "香港大学": "University of Hong Kong", "港大": "University of Hong Kong", "hku": "University of Hong Kong",
  "香港中文大学": "Chinese University of Hong Kong", "港中文": "Chinese University of Hong Kong", "cuhk": "Chinese University of Hong Kong",
  "香港科技大学": "Hong Kong University of Science and Technology", "港科大": "Hong Kong University of Science and Technology", "hkust": "Hong Kong University of Science and Technology",
  "东京大学": "University of Tokyo", "utokyo": "University of Tokyo",
  "首尔大学": "Seoul National University", "snu": "Seoul National University",
  "kaist": "KAIST",
};

function findUnivName(input) {
  const term = norm(input);
  if (!term) return null;
  // 1. 精确匹配
  for (const k of Object.keys(UNIV_ALIASES)) {
    if (norm(k) === term) return UNIV_ALIASES[k];
  }
  // 2. 输入是词典键的前缀（如「清华」→「清华大学」），取最短键
  let best = null;
  let bestLen = Infinity;
  for (const k of Object.keys(UNIV_ALIASES)) {
    const kn = norm(k);
    if (kn && kn.startsWith(term) && kn.length < bestLen) {
      bestLen = kn.length;
      best = UNIV_ALIASES[k];
    }
  }
  if (best) return best;
  // 3. 词典键是输入的子串（如「清华大学（北京）」）
  best = null;
  bestLen = 0;
  for (const k of Object.keys(UNIV_ALIASES)) {
    const kn = norm(k);
    if (kn && term.includes(kn) && kn.length > bestLen) {
      bestLen = kn.length;
      best = UNIV_ALIASES[k];
    }
  }
  return best;
}

async function resolveInstitution(q) {
  const query = findUnivName(q) || q;
  const j = await oaFetch(
    "/institutions?search=" + encodeURIComponent(query) + "&per-page=5"
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
