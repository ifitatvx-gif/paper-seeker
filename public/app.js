(function () {
  "use strict";

  const $ = (sel) => document.querySelector(sel);
  const universityInput = $("#university");
  const fieldInput = $("#field");
  const searchBtn = $("#search-btn");
  const statusBox = $("#status");
  const resultsBox = $("#results");
  const datalist = $("#field-suggestions");

  let currentResult = null;

  // ---------- 安全转义 ----------
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // ---------- 领域建议 ----------
  async function loadFields() {
    try {
      const r = await fetch("/api/fields");
      const j = await r.json();
      const opts = [];
      (j.fields || []).forEach((f) => {
        opts.push(`<option value="${esc(f.zh)} · ${esc(f.en)}"></option>`);
      });
      datalist.innerHTML = opts.join("");
    } catch (e) {
      /* 忽略 */
    }
  }

  // ---------- 状态提示 ----------
  function setStatus(type, html) {
    if (!html) {
      statusBox.className = "status hidden";
      statusBox.innerHTML = "";
      return;
    }
    statusBox.className = "status " + type;
    statusBox.innerHTML = html;
  }

  // ---------- 查询 ----------
  async function doSearch() {
    const university = universityInput.value.trim();
    const field = fieldInput.value.trim();
    if (!university || !field) {
      setStatus("error", "⚠️ 请同时填写「大学名称」和「专业/领域」");
      return;
    }
    searchBtn.disabled = true;
    resultsBox.innerHTML = "";
    setStatus(
      "info",
      '<span class="spinner"></span>正在查询 OpenAlex 学术数据库，识别研究人员并整理其论文与综述，约需 5–20 秒…'
    );

    try {
      const url =
        "/api/search?field=" +
        encodeURIComponent(field) +
        "&university=" +
        encodeURIComponent(university);
      const r = await fetch(url);
      const data = await r.json();

      if (!data.ok) {
        setStatus("error", "⚠️ " + esc(data.error || "查询失败"));
        currentResult = null;
        return;
      }
      currentResult = data;
      setStatus(null);
      renderResults(data);
    } catch (e) {
      setStatus("error", "⚠️ 网络或服务器错误：" + esc(e.message));
      currentResult = null;
    } finally {
      searchBtn.disabled = false;
    }
  }

  // ---------- 渲染结果 ----------
  function workItemHtml(w, idx) {
    const links = [];
    if (w.doiUrl)
      links.push(`<a class="link-btn doi" href="${esc(w.doiUrl)}" target="_blank" rel="noopener">DOI</a>`);
    if (w.pdf)
      links.push(`<a class="link-btn pdf" href="${esc(w.pdf)}" target="_blank" rel="noopener">PDF</a>`);
    if (w.oaUrl && w.oaUrl !== w.pdf)
      links.push(`<a class="link-btn oa" href="${esc(w.oaUrl)}" target="_blank" rel="noopener">开放获取页</a>`);
    if (!links.length && w.url)
      links.push(`<a class="link-btn oa" href="${esc(w.url)}" target="_blank" rel="noopener">OpenAlex</a>`);

    const meta = [];
    if (w.year) meta.push("📅 " + w.year);
    if (w.source) meta.push("📖 " + w.source);
    if (w.cited) meta.push("🔥 被引 " + w.cited);
    if (w.authors && w.authors.length) meta.push("✍ " + w.authors.join(", "));

    const titleLink = w.doiUrl || w.url || "#";
    return `
      <div class="work-item">
        <p class="work-title" data-title="${esc(w.title)}"><a href="${esc(titleLink)}" target="_blank" rel="noopener">${idx}. ${esc(w.title)}</a></p>
        <div class="work-meta">${meta.map(esc).join(" · ")}</div>
        <div class="work-links">${links.join("")}</div>
      </div>`;
  }

  function professorCardHtml(p, i) {
    const orcid = p.orcid
      ? `<span class="orcid">ORCID: ${esc(p.orcid)}</span>`
      : "";
    const papers = p.papers || [];
    const reviews = p.reviews || [];
    const chinese = p.chinese || [];
    const papersHtml = papers.length
      ? papers.map((w, k) => workItemHtml(w, k + 1)).join("")
      : '<div class="empty-note">暂无论文记录</div>';
    const reviewsHtml = reviews.length
      ? reviews.map((w, k) => workItemHtml(w, k + 1)).join("")
      : '<div class="empty-note">暂无综述记录</div>';
    const chineseHtml = chinese.length
      ? chinese.map((w, k) => workItemHtml(w, k + 1)).join("")
      : '<div class="empty-note">暂无中文文献记录</div>';

    return `
      <article class="professor-card" id="prof-${i}">
        <div class="prof-head">
          <div>
            <h2 class="prof-name">${esc(p.name)}${orcid}</h2>
            ${p.affiliation ? `<div class="prof-affil">${esc(p.affiliation)}</div>` : ""}
          </div>
          <button class="export-one-btn" data-export-one="${i}">📄 导出此教授 Word</button>
        </div>
        <div class="prof-stats">
          <span class="stat-pill">论文 <b>${papers.length}</b> 篇</span>
          <span class="stat-pill">综述 <b>${reviews.length}</b> 篇</span>
          <span class="stat-pill">中文论文 <b>${chinese.length}</b> 篇</span>
          <span class="stat-pill">代表作品累计被引 <b>${p.cited}</b></span>
        </div>
        <div class="tab-bar">
          <div class="tab active" data-tab="papers">📄 论文 <span class="count">(${papers.length})</span></div>
          <div class="tab" data-tab="reviews">📑 综述 <span class="count">(${reviews.length})</span></div>
          <div class="tab" data-tab="chinese">🇨🇳 中文论文 <span class="count">(${chinese.length})</span></div>
        </div>
        <div class="work-list" data-pane="papers">${papersHtml}</div>
        <div class="work-list hidden" data-pane="reviews">${reviewsHtml}</div>
        <div class="work-list hidden" data-pane="chinese">${chineseHtml}</div>
      </article>`;
  }

  function renderResults(data) {
    const profs = data.professors || [];
    const inst = data.institution || {};
    const concept = data.concept || {};

    let html = `
      <div class="summary">
        <div class="meta">
          <strong>🏛 ${esc(inst.display_name)}${inst.display_name_zh ? "（" + esc(inst.display_name_zh) + "）" : ""}</strong>
          <span class="tag">领域：${esc(data.fieldLabel || concept.display_name)}</span><br/>
          <span style="font-size:13px;color:var(--muted)">共找到 ${profs.length} 位研究人员（按领域内发文量与被引排序）</span>
        </div>
        <div class="summary-actions">
          <button class="translate-btn" id="translate-btn">🌐 标题翻译成中文</button>
          <button class="export-all-btn" id="export-all">📥 导出全部 Word 文档</button>
        </div>
      </div>`;

    const topReviews = data.topReviews || [];
    if (topReviews.length) {
      html += `
        <section class="review-panel">
          <h2 class="panel-title">📑 领域代表性综述（${esc(inst.display_name)}${inst.display_name_zh ? "（" + esc(inst.display_name_zh) + "）" : ""} · ${esc(data.fieldLabel || concept.display_name)}）</h2>
          <div class="work-list">${topReviews.map((w, i) => workItemHtml(w, i + 1)).join("")}</div>
        </section>`;
    }

    const topChinese = data.topChinese || [];
    if (topChinese.length) {
      html += `
        <section class="review-panel cn-panel">
          <h2 class="panel-title">🇨🇳 中文文献（${esc(inst.display_name)}${inst.display_name_zh ? "（" + esc(inst.display_name_zh) + "）" : ""}，按被引排序）</h2>
          <div class="work-list">${topChinese.map((w, i) => workItemHtml(w, i + 1)).join("")}</div>
        </section>`;
    }

    html += profs.map(professorCardHtml).join("");
    resultsBox.innerHTML = html;

    // 事件绑定
    document.getElementById("export-all").addEventListener("click", () => {
      exportWord(currentResult);
    });
    document.getElementById("translate-btn").addEventListener("click", translateAllTitles);

    resultsBox.querySelectorAll(".professor-card").forEach((card) => {
      card.querySelectorAll(".tab").forEach((tab) => {
        tab.addEventListener("click", () => {
          const target = tab.getAttribute("data-tab");
          card.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
          tab.classList.add("active");
          card.querySelectorAll("[data-pane]").forEach((pane) => {
            pane.classList.toggle("hidden", pane.getAttribute("data-pane") !== target);
          });
        });
      });
      card.querySelectorAll("[data-export-one]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const idx = parseInt(btn.getAttribute("data-export-one"), 10);
          const p = currentResult.professors[idx];
          exportWord({
            university: currentResult.institution,
            concept: currentResult.concept,
            fieldLabel: currentResult.fieldLabel,
            professors: [p],
          });
        });
      });
    });

    resultsBox.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // ---------- 导出 Word ----------
  async function exportWord(payload) {
    if (!payload || !payload.professors || !payload.professors.length) return;
    const btn = document.querySelector(".export-all-btn");
    try {
      const res = await fetch("/api/export-word", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const j = await res.json().catch(() => ({}));
        alert("导出失败：" + (j.error || res.status));
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download =
        "论文查询报告_" + (payload.university.display_name || "结果") + ".docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (e) {
      alert("导出失败：" + e.message);
    }
  }

  // ---------- 标题翻译 ----------
  const mtCache = {};
  let mtBusy = false;

  async function translateTitle(text) {
    if (mtCache[text] !== undefined) return mtCache[text];
    try {
      const url = "https://api.mymemory.translated.net/get?q=" + encodeURIComponent(text) + "&langpair=en|zh-CN";
      const res = await fetch(url);
      const j = await res.json();
      let zh = null;
      if (j && j.responseStatus === 200 && j.responseData && j.responseData.translatedText) {
        zh = j.responseData.translatedText.trim();
        if (zh.toLowerCase() === text.toLowerCase()) zh = null;
      }
      if (j && j.quotaFinished) mtCache["__quota__"] = true;
      mtCache[text] = zh;
      return zh;
    } catch (e) {
      mtCache[text] = null;
      return null;
    }
  }

  async function translateAllTitles() {
    if (mtBusy) return;
    const btn = document.getElementById("translate-btn");
    const items = Array.from(resultsBox.querySelectorAll(".work-title[data-title]"));
    const todo = items.filter((el) => el.getAttribute("data-title") && !el.querySelector(".zh-title"));
    if (!todo.length) { alert("没有需要翻译的标题（可能已全部翻译）"); return; }
    mtBusy = true;
    let done = 0, failed = 0, quotaHit = false, idx = 0;
    const CONC = 3;
    async function worker() {
      while (idx < todo.length && !quotaHit) {
        const el = todo[idx++];
        const en = el.getAttribute("data-title");
        const zh = await translateTitle(en);
        if (zh) {
          const div = document.createElement("div");
          div.className = "zh-title";
          div.textContent = "中文：" + zh;
          el.appendChild(div);
          done++;
        } else {
          failed++;
          if (mtCache["__quota__"]) quotaHit = true;
        }
        btn.textContent = "🌐 翻译中 " + (done + failed) + "/" + todo.length;
      }
    }
    const workers = [];
    for (let k = 0; k < CONC; k++) workers.push(worker());
    await Promise.all(workers);
    mtBusy = false;
    btn.textContent = "🌐 标题翻译成中文";
    let msg = "已翻译 " + done + " 条标题";
    if (quotaHit) msg += "（免费翻译额度已用完，其余标题暂未翻译）";
    else if (failed) msg += "，" + failed + " 条翻译失败";
    setStatus("info", msg);
  }

  // ---------- 初始化 ----------
  searchBtn.addEventListener("click", doSearch);
  [universityInput, fieldInput].forEach((el) => {
    el.addEventListener("keydown", (e) => {
      if (e.key === "Enter") doSearch();
    });
  });
  loadFields();
})();
