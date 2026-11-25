// onboarding.js — 首次启动使用引导（欢迎卡 + 分步高亮 coach marks）
// 纯前端控制 + 极薄服务端：状态落 config/onboarding.json，GET/POST /api/onboarding。
// 引导只做概念级主线条导引，不演示操作细节；非阻塞、可跳过、可重看（重看不改 done）。

(function () {
  const overlay = document.getElementById("onboarding-overlay");
  const TOTAL = 4;

  // 步骤顺序（用户拍板）：欢迎 → 设置·配置模型服务 → 新建/导入模板 → 文档排版
  // 文案 2026-08-28 用户逐屏定稿，与《首次启动使用引导-设计方案.md》§五一致。
  const STEPS = [
    { kind: "welcome" },
    {
      kind: "nav", sel: "#nav-settings",
      title: "设置 · 配置模型服务",
      text: "配置 LLM 模型服务，启用 AI 增强能力。输入模型服务名称（自定义）、URL 地址、API Key、模型名称。",
    },
    {
      kind: "nav", sel: "#nav-design",
      title: "新建 / 导入模板",
      text: "模板决定输出规范。可从零新建模板，或从已有 Word 文档一键导入生成模板，复用既有版式。",
    },
    {
      kind: "nav", sel: "#nav-convert",
      title: "文档排版",
      text: "把本地或在线文档按模板重新排版，一键输出格式规范的 Word 文档；排版在线文档需先扫码登录。",
    },
  ];

  let obForce = false; // 设置页重看 = true：关闭后不改 done
  let obStep = 0;

  function el(tag, cls) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    return e;
  }

  function openOb(opts) {
    opts = opts || {};
    obForce = !!opts.force;
    obStep = 0;
    overlay.classList.remove("hidden");
    overlay.innerHTML = "";
    render();
    window.addEventListener("resize", onResize);
  }

  async function maybeShowOnboarding() {
    try {
      const r = await fetch("/api/onboarding");
      const d = await r.json();
      if (d && d.done === false) openOb({ force: false });
    } catch (e) {
      /* 忽略：获取失败不阻断主流程 */
    }
  }

  function onResize() {
    if (!overlay.classList.contains("hidden")) render();
  }

  function render() {
    overlay.innerHTML = "";
    const step = STEPS[obStep];
    if (step.kind === "welcome") renderWelcome();
    else renderNav(step);
  }

  function renderWelcome() {
    const blocker = el("div", "ob-blocker");
    const card = el("div", "ob-welcome");
    card.innerHTML =
      '<div class="ob-welcome-title">欢迎使用文档审核</div>' +
      '<div class="ob-welcome-desc">把在线或本地文档，按规范模板重新排版，三步得到格式规范的 Word 文档：</div>' +
      '<ul class="ob-lines">' +
      '<li><div class="ob-line-head"><span class="ob-no">1</span><b>设置 · 配置模型服务</b><i>必选</i></div>' +
      '<div class="ob-line-text">配置 LLM 大模型服务后方可启用 AI 增强能力，进行模板导入、文档排版。</div></li>' +
      '<li><div class="ob-line-head"><span class="ob-no">2</span><b>新建 / 导入模板</b></div>' +
      '<div class="ob-line-text">模板决定输出规范。可以从零新建，也可以从已有 Word 文档导入版式。</div></li>' +
      '<li><div class="ob-line-head"><span class="ob-no">3</span><b>文档排版</b><i>核心能力</i></div>' +
      '<div class="ob-line-text">输入本地或在线文档链接，选择模板和格式选项，输出格式规范的 Word 文档。</div></li>' +
      "</ul>" +
      '<div class="ob-actions">' +
      '<button class="small" id="ob-skip">跳过</button>' +
      '<button class="big primary" id="ob-next">开始引导</button>' +
      "</div>";
    overlay.appendChild(blocker);
    overlay.appendChild(card);
    card.querySelector("#ob-next").onclick = () => { obStep = 1; render(); };
    card.querySelector("#ob-skip").onclick = finish;
  }

  function renderNav(step) {
    const target = document.querySelector(step.sel);
    const blocker = el("div", "ob-blocker");
    const hole = el("div", "ob-hole");
    const pop = el("div", "ob-pop");
    const isLast = obStep === STEPS.length - 1;
    const prevBtn = obStep > 1 ? '<button class="small" id="ob-prev">上一步</button>' : "";
    const nextBtn = isLast
      ? '<button class="small primary" id="ob-done">完成</button>'
      : '<button class="small primary" id="ob-next2">下一步</button>';
    pop.innerHTML =
      '<div class="ob-pop-title">' + step.title + "</div>" +
      '<div class="ob-pop-text">' + step.text + "</div>" +
      '<div class="ob-pop-foot">' +
      '<span class="ob-step">' + (obStep + 1) + " / " + STEPS.length + "</span>" +
      '<span class="ob-btns">' + prevBtn +
      '<button class="small" id="ob-skip2">跳过</button>' + nextBtn +
      "</span></div>";
    overlay.appendChild(blocker);
    overlay.appendChild(hole);
    overlay.appendChild(pop);

    if (target) {
      const r = target.getBoundingClientRect();
      const PAD = 6;
      hole.style.top = (r.top - PAD) + "px";
      hole.style.left = (r.left - PAD) + "px";
      hole.style.width = (r.width + PAD * 2) + "px";
      hole.style.height = (r.height + PAD * 2) + "px";

      const popW = pop.offsetWidth;
      const popH = pop.offsetHeight;
      let left = r.right + 16;
      if (left + popW > window.innerWidth - 12) left = r.left - popW - 16;
      if (left < 12) left = 12;
      let top = r.top;
      if (top + popH > window.innerHeight - 12) top = Math.max(12, window.innerHeight - popH - 12);
      pop.style.left = left + "px";
      pop.style.top = top + "px";
    }

    const prev = pop.querySelector("#ob-prev");
    if (prev) prev.onclick = () => { obStep--; render(); };
    const nx = pop.querySelector("#ob-next2");
    if (nx) nx.onclick = () => { obStep++; render(); };
    const done = pop.querySelector("#ob-done");
    if (done) done.onclick = finish;
    pop.querySelector("#ob-skip2").onclick = finish;
  }

  function finish() {
    window.removeEventListener("resize", onResize);
    overlay.classList.add("hidden");
    overlay.innerHTML = "";
    if (!obForce) {
      fetch("/api/onboarding", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ done: true }),
      }).catch(() => {});
    }
  }

  // 暴露给 app.js：init 末尾自动判定；设置页「查看使用引导」按钮调用（force=true 重看）
  window.maybeShowOnboarding = maybeShowOnboarding;
  window.showOnboarding = openOb;
})();
