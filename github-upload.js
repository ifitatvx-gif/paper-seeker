// 论文seeker —— GitHub 上传脚本（通过 GitHub REST API，无需安装 Git）
// 用法（环境变量传参，避免把 Token 写进文件）：
//   $env:GH_USER="你的用户名"; $env:GH_TOKEN="ghp_xxx"; $env:REPO="paper-seeker"; $env:VISIBILITY="public"
//   node github-upload.js
const fs = require("fs");
const path = require("path");

const GH_USER = process.env.GH_USER;
const GH_TOKEN = process.env.GH_TOKEN;
const REPO = process.env.REPO;
const VISIBILITY = process.env.VISIBILITY || "public";

if (!GH_USER || !GH_TOKEN || !REPO) {
  console.error("缺少参数：请设置环境变量 GH_USER、GH_TOKEN、REPO（可选 VISIBILITY）");
  process.exit(1);
}

const API = "https://api.github.com";
const SKIP = new Set(["node_modules", ".npm-cache", ".git"]);

async function api(method, pathname, body) {
  const headers = {
    Authorization: "Bearer " + GH_TOKEN,
    Accept: "application/vnd.github+json",
    "User-Agent": "paper-seeker-upload",
  };
  if (body) headers["Content-Type"] = "application/json";
  const res = await fetch(API + pathname, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });
  const txt = await res.text();
  let data = txt;
  try { data = JSON.parse(txt); } catch (e) {}
  return { status: res.status, data };
}

function walk(dir, base, out) {
  for (const name of fs.readdirSync(dir)) {
    if (SKIP.has(name)) continue;
    const full = path.join(dir, name);
    const rel = path.relative(base, full).replace(/\\/g, "/");
    const st = fs.statSync(full);
    if (st.isDirectory()) walk(full, base, out);
    else if (st.isFile()) out.push(rel);
  }
  return out;
}

(async () => {
  // 1. 检查/创建仓库
  let r = await api("GET", `/repos/${GH_USER}/${REPO}`);
  if (r.status === 404) {
    r = await api("POST", "/user/repos", {
      name: REPO,
      private: VISIBILITY === "private",
      description: "论文seeker —— 按大学 + 领域查询教授论文/综述的面试准备工具",
      auto_init: false,
    });
    if (r.status >= 400) {
      console.error("创建仓库失败:", r.status, JSON.stringify(r.data).slice(0, 400));
      process.exit(1);
    }
    console.log("✅ 已创建仓库:", r.data.html_url);
  } else if (r.status === 200) {
    console.log("ℹ️  仓库已存在:", r.data.html_url);
  } else {
    console.error("检查仓库失败:", r.status, JSON.stringify(r.data).slice(0, 400));
    process.exit(1);
  }

  // 2. 上传文件
  const base = process.cwd();
  const files = walk(base, base, []);
  let ok = 0, fail = 0;
  for (const rel of files) {
    const content = fs.readFileSync(path.join(base, rel)).toString("base64");
    const u = await api("PUT", `/repos/${GH_USER}/${REPO}/contents/${rel}`, {
      message: "上传 " + rel,
      content,
      branch: "main",
    });
    if (u.status === 201 || u.status === 200) {
      console.log("  ✓", rel);
      ok++;
    } else {
      console.log("  ✗", rel, u.status, JSON.stringify(u.data).slice(0, 160));
      fail++;
    }
  }

  console.log(`\n完成：成功 ${ok} 个，失败 ${fail} 个`);
  console.log(`仓库地址：https://github.com/${GH_USER}/${REPO}`);
})();
