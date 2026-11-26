/* 文档审核（文档格式化输出）· 模板设计器逻辑
   概念：模板集（TemplateSet）包含多个模板（Template）；模板设计对象是「模板集 / 模板」。
   属性区为 Word 风格「样式」面板：样式列表（名称+摘要+启用开关）→ 点击展开格式编辑。 */

"use strict";

// ═══════════════════════════════════════════
// 刷新快捷键（Launcher WebView2 无菜单/无 DevTools）
//   F5 / Ctrl+R      普通刷新（server.js 已对静态文件返回 no-store，reload 即拉最新）
//   Ctrl+F5          强制刷新（URL 加时间戳，彻底绕过一切缓存）
// ═══════════════════════════════════════════
document.addEventListener("keydown", e => {
  const hard = e.ctrlKey && e.key === "F5";
  const soft = e.key === "F5" || (e.ctrlKey && (e.key === "r" || e.key === "R"));
  if (!hard && !soft) return;
  e.preventDefault();
  if (hard) {
    const u = new URL(location.href);
    u.searchParams.set("_t", Date.now());
    location.href = u.toString();
  } else {
    location.reload();
  }
});

// ═══════════════════════════════════════════
// 全局状态
// ═══════════════════════════════════════════
const state = {
  templateSet: null,          // 当前编辑的模板集 {name, templates:[...]}
  currentFile: "",            // 模板集文件名（空=未保存）
  currentTemplateKey: null,   // 当前编辑的模板 key
  selectedElem: null,         // 当前展开编辑表单的样式 key
  paraSelected: false,        // （已弃用：原覆盖模式标志，现统一用 selectedParaKey）
  selectedParaKey: null,      // 选中段落当前生效的样式 key（用于样式面板高亮）
  selectedParaIdx: null,      // 选中段落对应在 sampleBlocks 里的索引（精确到段，避免同文本多段一起被恢复为选中）
  paraOverrides: [],          // 段落级样式覆盖（预览试样式用，存 localStorage，不落模板）
  collapsed: {},              // 样式分组折叠状态（key：text / table / page / "table:<二级组名>"）
  pendingTableGroup: false,   // 表格「新建分组」占位：渲染一个空分组并进入改名（无弹窗）
  sampleBlocks: [],           // 当前样品文档块清单（数据驱动预览，页面零硬编码）
  sampleFile: "",             // 当前样品文件名（用于工具条显示）
  sampleRaw: "",              // 当前模板样本原文（编辑弹窗预填用；无样本为空串）
  dirty: false,
  polling: null,
  nodeUseEnvProxy: false,     // Node 进程是否带 NODE_USE_ENV_PROXY 开关（进程启动时决定）
  systemProxyEnabled: false,  // Windows 系统代理是否启用（实时读注册表）
};

// ── 样式库访问（数据驱动，替代硬编码 ELEM_KEYS/ELEM_LABELS） ──
// v1.30 解耦：样式定义全部随模板自包含（templates[].styles[]），模板集不再持有任何样式数组。
// 种子（config/style-seeds.json）仅作「新建模板/旧文件兜底」时的一次性深拷贝来源。
let SEED_STYLE_CACHE = null;   // 种子定义缓存（模块级，懒加载）
async function ensureSeedStyles() {
  if (SEED_STYLE_CACHE) return SEED_STYLE_CACHE;
  // 种子里的字体是字族语义名，落值要靠映射表 ⇒ 先备好映射数据再取种子
  await ensureFontMapping();
  try {
    const seed = await api("/api/style-seeds");
    SEED_STYLE_CACHE = (seed && Array.isArray(seed.styles)) ? JSON.parse(JSON.stringify(seed.styles)) : [];
  } catch (e) { SEED_STYLE_CACHE = []; }
  return SEED_STYLE_CACHE;
}

/** 字体映射表 + 应用级开关（幂等）：设置页展示与「种子 → 模板」写入路径共用同一份数据。 */
let _fontMappingPromise = null;
async function ensureFontMapping() {
  if (!_fontMappingPromise) {
    _fontMappingPromise = (async () => {
      try {
        const r = await api("/api/font-families");
        fontFamilies = Array.isArray(r.families) ? JSON.parse(JSON.stringify(r.families)) : [];
      } catch (e) { fontFamilies = []; }
      try { appSettings = await api("/api/settings"); }
      catch (e) { appSettings = { forceMappedFont: true }; }   // 读不到 → 与服务端默认一致（开）
    })();
  }
  return _fontMappingPromise;
}

/** 字体落值（「种子 → 模板」的唯一收口）：`font` 走映射，字族语义名 → 当前开关指定的具体字体名。
 *  ★ 字体值分两层：**种子存国标字族语义名**（小标宋 / 黑体 / 楷体 / 仿宋 / 宋体），
 *    **模板存具体字体名**（导出 `w:rFonts` 要真名，如「方正小标宋简体」）。
 *    落值只发生在「种子 → 模板」的写入时刻 —— 写入点：本函数（及调用它的 seedToTemplateDef）
 *    见于新建模板 / 从默认样式中选 / 重置样式 / 导入结果页勾选补入 / 补建引导段；
 *    后端对应 engineSeedDefs()（模板集/模板补缺、表格样式兜底）。导入侧走 normalizeFontForTemplate()。
 *  ⚠ **别往种子 config/style-seeds.json 里写具体字体名** —— 写了「使用方正字库」开关在新建模板
 *    这条主路上就失效（开关只切映射取值列，不回溯存量模板、不影响转换）。
 *  判据：字族是国标规定，具体字体名是本机实现细节，只有模板这一层该承载实现值。
 *  非字族名（用户自装字体）与映射表未就绪时原样返回，不改写用户给的值。
 *  `pageFont` 不走映射：页码国标 7.5 定死「4 号半角宋体」，取系统自带宋体才各机器通用。 */
function materializeSeedDefaults(defaults) {
  const d = JSON.parse(JSON.stringify(defaults || {}));
  if (typeof d.font === "string" && d.font) d.font = resolveFontForWrite(d.font);
  return d;
}

/** 种子样式条目 → 模板样式条目（defaults 内字体已落值）。 */
function seedToTemplateDef(seed) {
  const copy = JSON.parse(JSON.stringify(seed));
  copy.defaults = materializeSeedDefaults(copy.defaults);
  return copy;
}
/** 种子「页面」样式的默认值（同步读已就绪的种子缓存；未就绪返回 {}）。
 *  新建/导入模板的页面默认值统一从这里取——硬编码会与种子漂移（改种子不生效）。 */
function seedPageDefaults() {
  return ((SEED_STYLE_CACHE || []).find(s => s.category === "page") || {}).defaults || {};
}

// 模板缺 styles[]（异常/损坏文件）时用种子补齐定义。
// 2026-09-02 方案A：样式值唯一真相 = styles[].defaults（顶层扁平字段废弃）；
// styles 补齐后执行一次性迁移（旧文件扁平值 → defaults 合并后删除）。
async function ensureTplStyles(tpl) {
  if (!tpl) return;
  if (!(Array.isArray(tpl.styles) && tpl.styles.length)) {
    const seeds = await ensureSeedStyles();
    tpl.styles = seeds.map(seedToTemplateDef);
  }
  migrateFlatToStyleDefaults(tpl);
}
// 方案A（2026-09-02）一次性迁移：顶层扁平样式字段（tpl[样式key]）→ 合并进 styles[].defaults 后删除。
// 扁平值覆盖 defaults 同键（旧语义「扁平优先」）；迁移在内存态完成，保存落盘即新结构。
function migrateFlatToStyleDefaults(tpl) {
  if (!tpl || !Array.isArray(tpl.styles)) return;
  for (const s of tpl.styles) {
    const flat = tpl[s.key];
    if (flat && typeof flat === "object" && !Array.isArray(flat)) {
      s.defaults = Object.assign({}, s.defaults || {}, flat);
    }
  }
  for (const s of tpl.styles) delete tpl[s.key];
}
// ⚠ 无「已废弃样式」清理（2026-09-14 起）：styles[] 条目按 key 存在即合法，与后端 schema 重建同口径
//   （convert.js 顶部 schema 表）。出厂删过的样式在存量模板里与用户自建样式形状一致、无法区分，
//   摘除等于静默丢弃用户文件内容 ⇒ 一律保留，去留交用户在模板设计页决定。
// 当前模板的自包含样式定义（唯一来源；无模板时为空）
function templateStyles() {
  const tpl = currentTemplate();
  return (tpl && Array.isArray(tpl.styles)) ? tpl.styles : [];
}
function allStyleDefs() {
  return templateStyles();
}
function styleKeys() {
  const keys = allStyleDefs().map(s => s.key);
  if (!keys.includes("footer")) keys.push("footer");
  return keys;
}
// 预览 include 兜底 key 集：与服务端 defsForTemplate 同源（引擎种子 ∪ 模板自包含 styles[] ∪ footer）。
// 仅用于 include 为空时的显示门控降级——include 为空时全量展示样本段落，
// 否则样本段落会被 display:none 全部隐藏 → 预览空白（2026-08-30）。SEED_STYLE_CACHE 未就绪时
// 退化为模板样式 ∪ footer（仍优于全隐藏）。v1.3：服务端已删全量占位，此处仅作显示门控。
function previewIncludeFallbackKeys() {
  const ks = new Set((SEED_STYLE_CACHE || []).map(s => s.key));
  for (const s of templateStyles()) ks.add(s.key);
  ks.add("footer");
  return [...ks];
}
function styleName(key) {
  const s = styleDef(key);
  return (s && s.name) ? s.name : (key === "footer" ? "页码" : key);   // v5.4 一条链：无 name 时样式名即 key
}
function styleDef(key) {
  return templateStyles().find(x => x.key === key) || null;
}
// 是否种子固定样式（判断基于种子 key，缓存未加载时视为自定义）。
// 固定 ≠ 不可删：可从当前模板移除（移除后按正文类样式降级渲染），随时从「从默认样式中选」加回。
function isFixedStyle(key) {
  return !!(SEED_STYLE_CACHE && SEED_STYLE_CACHE.some(s => s.key === key));
}

const FONTS = [
  // 中文 Windows 自带字体（小标宋不在其列：系统无此字体，须装方正字库，故不进下拉）
  "黑体", "楷体", "仿宋", "宋体",
  // 方正字库（办公环境常见，映射表首选）
  "方正小标宋简体", "方正黑体简体", "方正仿宋简体", "方正楷体简体", "方正书宋简体",
  "仿宋_GB2312", "楷体_GB2312",
  "微软雅黑",
  "Times New Roman", "Arial", "Calibri",
];
const ALIGNS = [
  { v: "left",    label: "左对齐" },
  { v: "center",  label: "居中" },
  { v: "right",   label: "右对齐" },
  { v: "justify", label: "两端对齐" },
];
// 中文字号 → pt
const CN_SIZES = [
  { label: "初号", pt: 42 }, { label: "小初", pt: 36 },
  { label: "一号", pt: 26 }, { label: "小一", pt: 24 },
  { label: "二号", pt: 22 }, { label: "小二", pt: 18 },
  { label: "三号", pt: 16 }, { label: "小三", pt: 15 },
  { label: "四号", pt: 14 }, { label: "小四", pt: 12 },
  { label: "五号", pt: 10.5 }, { label: "小五", pt: 9 },
  { label: "六号", pt: 7.5 }, { label: "小六", pt: 6.5 },
  { label: "七号", pt: 5.5 }, { label: "八号", pt: 5 },
];
// 单位换算
const DXA_PER_CM = 566.93;   // 1cm ≈ 566.93 DXA
function cmToDxa(cm) { return Math.round(parseFloat(cm || 0) * DXA_PER_CM); }
// cm 展示统一保留 1 位小数（0.1cm 粒度；2026-09-02 用户定调：样式面板页边距步进 0.1，显示/输入/回显同精度）
function dxaToCm(dxa) { return (Math.round((parseFloat(dxa || 0) / DXA_PER_CM) * 10) / 10).toFixed(1); }
function ptToCnSize(pt) {
  const p = parseFloat(pt);
  const exact = CN_SIZES.find(s => Math.abs(s.pt - p) < 0.01);
  return exact ? exact.label : (p ? `${p}pt` : "");
}

// ═══════════════════════════════════════════
// 小工具
// ═══════════════════════════════════════════
async function api(path, opts = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...opts,
  });
  let data = null;
  try { data = await res.json(); } catch (e) { /* 非 JSON */ }
  if (!res.ok) throw new Error((data && data.error) || `请求失败(${res.status})`);
  return data;
}

function toast(msg, type = "") {
  const el = document.getElementById("toast");
  document.getElementById("toast-msg").textContent = msg;
  el.className = "toast " + type;
  el.classList.remove("hidden");
  clearTimeout(toast._t);
  // 按时间自动消失：err 6s（常含「去设置页处理」等指引，给足阅读/复制时间），ok/warn 3s；
  // 也可点右上角 × 立即关闭。正文文字可选中复制，点击/选中文字不会关闭（关闭只走 × 按钮）。
  toast._t = setTimeout(() => el.classList.add("hidden"), type === "err" ? 6000 : 3000);
}

// LLM 配置前置检查（导入模板 / 文档排版共用）：
//   —— 判据与设置页一致：只认"当前启用的那一个"（自配中被启用且挂了模型）；
//   一个都没启用 → false，由调用方提示用户到「设置」启用，不发起任务、不产生占位。
// 查询接口失败时不拦截（返回 true），由后端执行时报错兜底，避免网络抖动误伤正常流程。
async function checkLlmConfigured() {
  try {
    // 拉一次最新状态（顺带预取内置凭证），再走内存态闸门 —— 全站只有一个判据 llmConfigured()，
    // 不在这里另写一套"有服务即算配置"的逻辑（那条老口径会把已停用的服务算进来）。
    await loadLlmList();
    return llmConfigured();
  } catch (e) {
    return true;
  }
}

// toast 关闭只走右上角 × 按钮；正文区域留给用户选中复制，不再「点哪都关」
document.getElementById("toast-close").addEventListener("click", () => {
  document.getElementById("toast").classList.add("hidden");
});

// 登录状态行已移除：状态/错误统一走 toast 提示。
function setStatus(html, cls) {
  toast(html, cls);
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// 过程行状态图标：全站内联 SVG（docs/界面样式规范.md §6：禁止字符/emoji 充当状态图标）；
// 唯一的例外是「进行中 ◌」——它走 klUI 的 CSS 转圈 `.kl-spinner`（全站转圈只此一个形状）。
// 入参沿用历史字符键（✓ 完成 / ✗ 失败 / ◌ 进行中 / ◆ 里程碑 / ⏳ 等待），调用点零改动；
// 历史任务的 procHtml 是序列化 HTML 回显，里面仍是纯字符或旧 SVG → CSS 侧保留字符配色与旧 SVG 旋转兜底，三者共存。
// ⚠ `.conv-line` 及其状态类名（ok / err / cur / wait / info）不可改名、不可删除：历史任务的过程面板是
//    序列化 HTML 快照（落盘在任务目录，回显时整段 innerHTML 回灌），改类名 = 所有历史任务过程样式全丢。
//    新增类一律向后兼容，旧类名保留。
const PROC_ICONS = {
  "✓": '<i class="fas fa-check"></i>',
  "✗": '<i class="fas fa-xmark"></i>',
  "◌": '<span class="kl-spinner"></span>',
  "◆": '<i class="fas fa-diamond"></i>',
  "⏳": '<i class="fas fa-hourglass-half"></i>',
  "■": '<i class="fas fa-square"></i>',
};
function procIcon(ch) { return PROC_ICONS[ch] || esc(ch); }
// 图标字符 → 行样式类：ok/err/cur 与历史 procHtml 同源；◆ 是中性里程碑（不染色），⏳ 等待 / ■ 已停止走 warn
function procLineCls(ch) {
  return ch === "✓" ? " ok" : ch === "✗" ? " err" : ch === "◌" ? " cur" : (ch === "⏳" || ch === "■") ? " wait" : "";
}

// 字节数格式化（B / KB / MB），用于转换过程与结果卡片显示文件大小
function formatBytes(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
  return (n / 1024 / 1024).toFixed(2) + " MB";
}

function currentTemplate() {
  if (!state.templateSet) return null;
  return state.templateSet.templates.find(t => t.key === state.currentTemplateKey) || null;
}

// ═══════════════════════════════════════════
// 自定义模态框（替换 prompt/confirm）
// ═══════════════════════════════════════════
let _modalResolve = null;

function openModal(opts) {
  return new Promise(resolve => {
    _modalResolve = resolve;
    const modalEl = document.querySelector("#modal-overlay .modal");
    // 复位：避免编辑器宽模态 / 信息模态 / 名单编辑模态残留到其他弹窗
    if (modalEl) modalEl.classList.remove("sample-editor-modal", "info-modal", "roster-editor-modal");
    if (opts.modalClass && modalEl) modalEl.classList.add(opts.modalClass);
    // 复位底部操作区：上一轮可能是「自建按钮」模式（hideButtons），那时额外挂了带 data-modal-custom
    // 标记的按钮 —— 这里只把那些按钮摘掉，**不动原生 ok/cancel 节点**。
    // ⚠ 不可写成 `actionsReset.innerHTML = '原生按钮…'` 去重建：设置 innerHTML 会销毁原节点，
    //   而 bindModalEvents() 的 click 监听是一次性绑在这些静态节点上的，节点一换监听就丢 ——
    //   症状是「点确定 / 删除毫无反应」，且第一次打开弹窗后所有原生按钮永久失效。
    // ★ 推论（后来人注意）：#modal-ok / #modal-cancel 是常驻节点，任何「专用弹窗接管」都必须用
    //   `okBtn.onclick = fn`（属性赋值天然覆盖、幂等），不能用 addEventListener —— 那会随每次打开
    //   弹窗累积，且「取消」关窗不解绑，下次点保存就并发跑多个 handler（重复提交、重复提示）。
    const actionsReset = document.querySelector("#modal-overlay .modal-actions");
    if (actionsReset) {
      actionsReset.querySelectorAll("[data-modal-custom]").forEach(n => n.remove());
      actionsReset.style.display = "";
    }
    disarmCancelBtn();   // 「取消」的二次确认态（红底 / 「再次点击确认放弃」）不残留到下一个弹窗
    if (opts.titleHtml !== undefined) document.getElementById("modal-title").innerHTML = opts.titleHtml;
    else document.getElementById("modal-title").textContent = opts.title || "";
    const body = document.getElementById("modal-body");
    if (opts.body !== undefined) {
      body.innerHTML = opts.body;              // 自定义 HTML 内容
    } else if (opts.input) {
      body.innerHTML = `${opts.message ? `<div class="modal-msg">${esc(opts.message)}</div>` : ""}
        <input type="text" id="modal-input" value="${esc(opts.value || "")}" placeholder="${esc(opts.placeholder || "")}" autofocus>`;
    } else if (opts.warn) {
      // 危险/不可逆确认：统一 .modal-warn 警告块（规范 §2）
      body.innerHTML = `<div class="modal-warn"><span class="mw-ic">${ICONS.warn}</span><span>${opts.message || ""}</span></div>`;
    } else {
      body.innerHTML = `<div class="modal-msg">${opts.message || ""}</div>`;
    }
    document.getElementById("modal-ok").textContent = opts.okText || "确定";
    const okBtn = document.getElementById("modal-ok");
    const cancelBtn = document.getElementById("modal-cancel");
    // 配套 actions 容器：custom 按钮模式下整体隐藏，避免 .modal-actions 的 border-top 露出空分隔线
    const actionsEl = document.querySelector("#modal-overlay .modal-actions");
    // 清除上一轮「专用弹窗」留下的自定义处理器（如 LLM 表单 / 样式编辑接管的保存逻辑）。
    // 两种模式都清：hideButtons 弹窗里虽点不到该按钮，但残留的处理器会活到下一个弹窗，越早清越稳。
    okBtn.onclick = null;
    if (opts.hideButtons) {
      okBtn.style.display = "none";
      cancelBtn.style.display = "none";
      if (actionsEl) actionsEl.style.display = "none";
    } else {
      okBtn.style.display = "";
      // 信息型弹窗只需「关闭」：隐藏取消，避免出现「取消 / 关闭」双按钮
      cancelBtn.style.display = opts.hideCancel ? "none" : "";
      if (actionsEl) actionsEl.style.display = "";
    }
    document.getElementById("modal-overlay").classList.remove("hidden");
    if (opts.input) {
      const inp = document.getElementById("modal-input");
      setTimeout(() => { inp.focus(); inp.select(); }, 50);
    }
  });
}

function closeModal(result) {
  document.getElementById("modal-overlay").classList.add("hidden");
  // 自建按钮模式（hideButtons + 追加按钮）的收尾：摘掉自建按钮（[data-modal-custom]）。
  // 原生 ok/cancel 节点始终留着 —— 它们的显隐由 openModal 复位。
  const acts = document.querySelector("#modal-overlay .modal-actions");
  if (acts) acts.querySelectorAll("[data-modal-custom]").forEach(n => n.remove());
  _modalTextarea = false;   // 多行输入态复位（Enter 拦截只在弹窗存活期间生效）
  if (_modalResolve) { _modalResolve(result); _modalResolve = null; }
}

/** 输入型模态框：返回输入值（取消返回 null） */
function modalPrompt(title, opts = {}) {
  return openModal({ title, input: true, value: opts.value || "", placeholder: opts.placeholder || "", message: opts.message || "" });
}

/** 多行输入型模态框（textarea）：返回输入值（取消返回 null）。
 *  与 modalPrompt 的差异：Enter 用于换行而非提交（_modalTextarea 守卫），仅「确定」按钮提交。 */
let _modalTextarea = false;
function modalPromptArea(title, opts = {}) {
  const p = openModal({
    title,
    body: `${opts.message ? `<div class="modal-msg">${esc(opts.message)}</div>` : ""}
      <textarea id="modal-input" class="modal-textarea" rows="${opts.rows || 3}" maxlength="${opts.maxlength || 200}"
        placeholder="${esc(opts.placeholder || "")}">${esc(opts.value || "")}</textarea>`
  });
  _modalTextarea = true;
  setTimeout(() => {
    const el = document.getElementById("modal-input");
    if (!el) return;
    el.focus();
    const n = el.value.length;
    try { el.setSelectionRange(n, n); } catch (e) { /* 忽略 */ }
  }, 50);
  return p;
}

/** 确认型模态框：确定返回 true，取消返回 null；危险/不可逆确认统一带 .modal-warn 警告块 */
function modalConfirm(title, message, okText = "确定") {
  return openModal({ title, message, okText, warn: true });
}

/** 信息型弹窗（关于本软件 / 支持格式）：仅「关闭」一个按钮，内容区超高可滚动。
 *  titleHtml 只在确需给标题加行内标记时传（如「关于」标题右侧的版本号），
 *  传了走 innerHTML、没传仍走 textContent —— 标题里可能拼用户数据，默认不开这口子。 */
function infoModal(title, bodyHtml, titleHtml) {
  return openModal({ title, titleHtml, body: bodyHtml, okText: "关闭", hideCancel: true, modalClass: "info-modal" });
}

// ── 设置页页脚：关于本软件 / 支持格式（使用引导复用 onboarding，见 bindEvents）──

function aboutBody() {
  const b = (k, v) => `<div class="info-block"><div class="info-k">${k}</div><div class="info-v">${v}</div></div>`;
  return b("用途", "按模板自动规范排版，输出标准 Word 文档。")
    + b("输入来源", "在线文档、WPS 文档，本地 Word / TXT / Markdown。")
    + b("智能校对", "可调用大模型校对内容（错别字、标点、表述）。")
    + b("模板设计", "导入 Word 文档自动识别样式生成模板。");
}

function formatsBody() {
  const b = (k, items) => `<div class="info-block"><div class="info-k">${k}</div><ul class="info-list">${items.map(t => `<li>${t}</li>`).join("")}</ul></div>`;
  return b("在线链接", [
    "在线文档、WPS 文档",
    "云盘文件：Word（docx / doc）",
  ])
    + b("本地文件", [
      "docx · doc · txt · md",
      "json：在线文档源快照（转换时自动存进任务目录，可离线重转对比）",
    ])
    + b("输出结果", [
      "docx（按模板样式重排）",
    ]);
}

// 更新日志 markdown → 折叠分节 HTML（## 版本标题为一节；首节=当前版本默认展开，历史版本折叠）
// 条目兼容 - 与 * 两种条目符；「新增 / 优化 / 修复 …」类型词加粗；# 总标题跳过（弹窗标题已表达）
function changelogMdToHtml(md) {
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const sections = [];
  let cur = null;
  for (const raw of String(md || "").split(/\r?\n/)) {
    const line = raw.trimEnd();
    if (/^##\s+/.test(line)) {
      cur = { title: esc(line.replace(/^##\s+/, "")), items: [] };
      sections.push(cur);
    } else if (cur && /^[-*]\s+/.test(line)) {
      const item = esc(line.replace(/^[-*]\s+/, "").trim());
      if (!item) continue;   // 空条目行跳过，不产生空列表项
      const m = item.match(/^(新增|优化|修复|变更|移除|安全)\s+(.*)$/);
      cur.items.push(`<li>${m ? `<b>${m[1]}</b> ${m[2]}` : item}</li>`);
    }
  }
  return sections.map((s, i) => {
    const open = i === 0 ? " open" : "";
    const unreleased = /未发布/.test(s.title) ? " cl-unreleased" : "";
    const body = s.items.length
      ? `<ul class="info-list cl-items">${s.items.join("")}</ul>`
      : `<div class="info-note">本版本暂无条目。</div>`;
    return `<details class="cl-sec"${open}><summary class="cl-ver${unreleased}">${s.title}</summary>${body}</details>`;
  }).join("");
}

async function openChangelogModal() {
  let bodyHtml;
  try {
    const d = await api("/api/changelog");
    bodyHtml = changelogMdToHtml(d && d.markdown);
    if (!bodyHtml) bodyHtml = `<div class="info-note">暂无更新日志。</div>`;
  } catch (e) {
    bodyHtml = `<div class="info-note">更新日志加载失败，请重试。</div>`;
  }
  infoModal("更新日志", bodyHtml);
}

// ═══════════════════════════════════════════
// 自动更新（协议：dist 仓 docs/更新协议.md schema 1）
//   前端的职责窄到只有三件事：展示新版说明、跟着下载进度、触发重启。
//   真正铺文件的是 launcher（下次启动、拉起 node 服务之前）—— 正在运行的程序覆盖不了自己，
//   前端永远不去动安装目录里的任何一个文件，也不决定「有没有新版」（那是服务端拉清单的结论）。
// ═══════════════════════════════════════════
const updater = {
  channel: "stable",
  current: "",
  latest: null,
  // 已下载待重启的更新（/api/update/state 的 pending）；有值时菜单显示「重启以完成更新」
  pending: null,
  checking: false,
  downloading: false,
};

function fmtBytes(n) {
  const b = Number(n) || 0;
  if (!b) return "未知";
  return b >= 1024 * 1024 ? (b / 1024 / 1024).toFixed(1) + "MB" : Math.round(b / 1024) + "KB";
}

/**
 * 取更新状态（当前版本 / 渠道 / 是否有已下载待重启的更新）。
 * 账号菜单里显示当前版本、以及判断要不要提示「重启即可完成更新」都靠它。
 * 取不到就返回 null —— 状态拿不到不该连累菜单里其它项。
 */
async function refreshUpdateState() {
  try {
    const st = await api("/api/update/state");
    updater.current = String((st && st.current) || "");
    updater.channel = String((st && st.channel) || "stable");
    updater.pending = (st && st.pending) || null;
    return st;
  } catch (e) {
    return null;
  }
}

/**
 * 检查更新。
 * @param {boolean} manual true = 用户在账号菜单里点的（任何结果都要给明确反馈）；
 *                         自动检查只在「真有新版本」时才打扰用户。
 *
 * 入口在侧栏底部的账号菜单（见 openAccountMenu）。点击后菜单已关闭，所以反馈走
 * toast 与更新弹窗，不再依赖某个常驻按钮的文案。
 */
async function checkForUpdate(manual) {
  if (updater.checking || updater.downloading) return;
  updater.checking = true;
  try {
    const r = await api("/api/update/check");
    if (r && r.hasUpdate && r.latest) {
      updater.latest = r.latest;
      openUpdateModal(r.latest);
      return;
    }
    if (!manual) return;
    if (r && !r.ok) { toast("检查更新失败，请稍后再试。", "err"); return; }
    const reason = (r && r.reason) || "";
    if (reason === "LOCAL_VERSION_UNKNOWN") {
      toast("当前为开发态版本（无版本号），不支持自动更新。", "warn");
    } else if (reason === "CLIENT_TOO_OLD") {
      toast("本渠道要求客户端 ≥ " + ((r && r.requiresClient) || "") + "，请下载完整安装包。", "warn");
    } else if (reason === "NO_ASSET_FOR_PLATFORM") {
      toast("清单里没有适用于本平台的安装包。", "err");
    } else {
      toast("已是最新版本" + (updater.current ? "（" + updater.current + "）" : "") + "。", "ok");
    }
  } catch (e) {
    if (manual) toast("检查更新失败：" + e.message, "err");
  } finally {
    updater.checking = false;
  }
}

function openUpdateModal(info) {
  const head = '<div class="up-dlg-head">'
    + '<div class="info-block"><div class="info-k">新版本</div><div class="info-v">' + esc(info.version)
    + (info.mandatory ? "（必须更新后才能继续使用）" : "") + '</div></div>'
    + '<div class="info-block"><div class="info-k">发布时间</div><div class="info-v">'
    + esc((info.publishedAt || "").replace("T", " ").slice(0, 16)) + '</div></div>'
    + '<div class="info-block"><div class="info-k">下载大小</div><div class="info-v">' + fmtBytes(info.size) + '</div></div>'
    + '</div>';
  const notes = changelogMdToHtml(info.notes)
    || '<div class="info-note">未能取得本次更新的说明（不影响下载安装）。</div>';
  openModal({ title: "发现新版本", body: head + notes, hideButtons: true, modalClass: "info-modal" });
  const actionsEl = document.querySelector("#modal-overlay .modal-actions");
  if (actionsEl) {
    actionsEl.style.display = "";
    // ⚠ 只追加自建按钮（带 data-modal-custom，由 openModal/closeModal 摘除）——不能碰原生 ok/cancel 节点
    actionsEl.insertAdjacentHTML("beforeend",
      (info.mandatory ? "" : '<button type="button" class="small" data-modal-custom id="upd-later">稍后提醒我</button>')
      + '<button type="button" class="small primary" data-modal-custom id="upd-now">立即更新</button>');
  }
  const btnLater = document.getElementById("upd-later");
  if (btnLater) btnLater.addEventListener("click", () => {
    api("/api/update/postpone", { method: "POST", body: JSON.stringify({ version: info.version }) }).catch(() => {});
    closeModal();
  });
  document.getElementById("upd-now").addEventListener("click", () => runUpdateDownload(info));
}

/** 下载并在弹窗里跟进度（SSS：download 阶段 → done/error） */
async function runUpdateDownload(info) {
  updater.downloading = true;
  const body = document.getElementById("modal-body");
  if (body) {
    body.innerHTML = '<div class="up-dlg-head"><div class="info-block"><div class="info-k">正在更新到</div>'
      + '<div class="info-v">' + esc(info.version) + '</div></div></div>'
      + '<div class="update-progress"><div class="up-bar"><div class="up-bar-in" id="upd-bar-in"></div></div>'
      + '<div class="up-stage" id="upd-stage">正在准备…</div></div>';
  }
  const actionsEl = document.querySelector("#modal-overlay .modal-actions");
  if (actionsEl) actionsEl.style.display = "none";
  const barIn = document.getElementById("upd-bar-in");
  const stage = document.getElementById("upd-stage");
  let result = null, errText = "";
  try {
    const resp = await fetch("/api/update/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ version: info.version, channel: updater.channel }),
    });
    if (!resp.ok || !resp.body) throw new Error("无法开始下载（HTTP " + resp.status + "）");
    await sseReadStream(resp, (ev, data) => {
      if (ev === "progress" && data) {
        if (barIn) barIn.style.width = Math.min(100, Number(data.percent) || 0) + "%";
        if (stage) {
          stage.textContent = (data.stage === "verify" ? "正在校验文件…" : "正在下载 " + (data.file || ""))
            + " " + (Number(data.percent) || 0) + "%（" + fmtBytes(data.received) + " / " + fmtBytes(data.total) + "）";
        }
      } else if (ev === "done") {
        result = data;
      } else if (ev === "error") {
        errText = (data && data.message) || "更新失败";
      }
    });
  } catch (e) {
    errText = e.message;
  } finally {
    updater.downloading = false;
  }
  if (result && result.ok) openUpdateRestartModal(info);
  else openUpdateFailModal(errText || "更新失败，请稍后重试。", info);
}

function openUpdateFailModal(msg, info) {
  openModal({ title: "更新未完成", body: '<div class="up-dlg-head"><div class="modal-msg">' + esc(msg)
    + '</div></div><div class="info-note">当前版本未做任何改动，仍可正常使用。</div>', hideButtons: true });
  const actionsEl = document.querySelector("#modal-overlay .modal-actions");
  if (actionsEl) {
    actionsEl.style.display = "";
    actionsEl.insertAdjacentHTML("beforeend",
      '<button type="button" class="small" data-modal-custom id="upd-close">关闭</button>'
      + '<button type="button" class="small primary" data-modal-custom id="upd-retry">重试</button>');
  }
  const c = document.getElementById("upd-close");
  if (c) c.addEventListener("click", () => closeModal());
  document.getElementById("upd-retry").addEventListener("click", () => runUpdateDownload(info));
}

function openUpdateRestartModal(info) {
  openModal({
    title: "更新已就绪",
    body: '<div class="up-dlg-head"><div class="modal-msg">新版本已下载并通过校验，重启应用后自动完成安装。</div></div>'
      + '<div class="info-note">模板、人员名单、任务记录都保存在用户数据目录中，安装新版本不会改动它们。</div>'
      + '<div class="up-stage" id="upd-stage2"></div>',
    hideButtons: true,
  });
  const actionsEl = document.querySelector("#modal-overlay .modal-actions");
  if (actionsEl) {
    actionsEl.style.display = "";
    actionsEl.insertAdjacentHTML("beforeend",
      '<button type="button" class="small" data-modal-custom id="upd-later2">稍后自行重启</button>'
      + '<button type="button" class="small primary" data-modal-custom id="upd-restart">立即重启</button>');
  }
  const l = document.getElementById("upd-later2");
  if (l) l.addEventListener("click", () => { closeModal(); refreshUpdateState(); });
  document.getElementById("upd-restart").addEventListener("click", async () => {
    const b = document.getElementById("upd-restart");
    if (b) { b.disabled = true; b.textContent = "正在重启…"; }
    try {
      const r = await api("/api/app/restart", { method: "POST", body: "{}" });
      if (r && r.error) throw new Error(r.error);
    } catch (e) {
      if (b) { b.disabled = false; b.textContent = "立即重启"; }
      const s = document.getElementById("upd-stage2");
      if (s) s.textContent = e.message;
      // 重启失败不影响已下载的更新：下次开机 launcher 会自己铺进去
    }
  });
}

// 页脚版本/年份：版本由服务端 /api/status 下发（APP_VERSION），取不到则留空不显示 v
let appVersion = "";
let appIsDev = false;       // 服务端报「未经 esbuild 打包」时为真 —— 页脚据此显示「开发版」标签
let _footerVerLoaded = false;

/**
 * 版本号的展示文案：正常是裸版本号（不带 v），拿不到或服务端明说缺失时返回空串。
 *
 * 服务端两条路都能给出真实版本：打包态取 esbuild 注入的常量，开发态用 fs 读 package.json
 * （见 server/server.js 的 readPkgVersion）。"dev" 只剩「两条路都失败」这一种可能 ——
 * 那种情况显示 `dev` 会像个真版本号，所以直接不显示，由「开发版」标签表达状态。
 */
function versionText(v) {
  if (!v || v === "dev") return "";
  return v;
}

/** 把 /api/status 的版本字段落到页脚状态（接口不可用时保持原样，不影响其余内容） */
function applyStatusVersion(st) {
  if (st && st.version) appVersion = String(st.version);
  if (st) appIsDev = !!st.dev;
}

async function initSettingsFooter() {
  const y = document.getElementById("sf-year");
  if (y) y.textContent = String(new Date().getFullYear());
  if (_footerVerLoaded) { renderFooterVersion(); return; }
  try {
    applyStatusVersion(await api("/api/status"));
  } catch (e) { /* 状态接口不可用：版本留空，不影响页脚其余内容 */ }
  _footerVerLoaded = true;
  renderFooterVersion();
}
function renderFooterVersion() {
  const el = document.getElementById("sf-version");
  if (el) el.textContent = versionText(appVersion);
  const dev = document.getElementById("sf-dev");
  if (dev) dev.hidden = !appIsDev;
}

/** 「关于」弹窗兜底取版本：页脚初始化只在进过设置页且接口返回后才填 appVersion，
 *  点到「关于」时可能还是空（或接口当次失败被 _footerVerLoaded 锁住），这里单独补一次。 */
async function ensureVersion() {
  if (appVersion) return;
  try {
    applyStatusVersion(await api("/api/status"));
  } catch (e) { /* 状态接口不可用：标题只显示软件名，不显示版本 */ }
}

/** 「取消」按钮的二次确认态：武装 = 红底白字 + 文案「再次点击确认放弃」；解除 = 复原「取消」。
 *  节点是常驻的（全站共用一个 modal-cancel），文案/样式不会自己还原 ⇒ 只有这两个出口，
 *  且 openModal 每次打开都会 disarm 一次，避免「武装后被保存等其他路径关窗」把红底文案残留到下一个弹窗。 */
function armCancelBtn() {
  const cb = document.getElementById("modal-cancel");
  if (!cb) return;
  cb.textContent = "再次点击确认放弃";
  cb.classList.add("danger");
}
function disarmCancelBtn() {
  const cb = document.getElementById("modal-cancel");
  if (!cb) return;
  cb.textContent = "取消";
  cb.classList.remove("danger");
}

function bindModalEvents() {
  document.getElementById("modal-ok").addEventListener("click", () => {
    // 样式编辑等自定义流程：由 editStyleModal 自行关闭并清理 handler
    if (_styleModalState) return;
    const inp = document.getElementById("modal-input");
    closeModal(inp ? inp.value : true);
  });
  document.getElementById("modal-cancel").addEventListener("click", () => {
    // LLM 表单有未保存改动 → 二次确认：第一次点击把按钮切成「再次点击确认放弃」，第二次点击才真正放弃。
    // ★ 武装态**不设超时自动复位**：用户看按钮、想清楚再点第二次，间隔往往超过几秒；自动复位会让
    //   第二次点击只是「重新武装」，弹窗永远关不掉。武装态只由「再点一次」或 openModal 打开新弹窗解除。
    if (_llmFormState && _llmFormState.dirty) {
      if (!_llmFormState.discardArmed) {
        _llmFormState.discardArmed = true;
        armCancelBtn();
        return;
      }
      _llmFormState.discardArmed = false;
    }
    _llmFormState = null;
    _styleModalState = null;
    closeModal(null);
  });
  // 遮罩点击：**不关闭弹窗**（全站统一，无例外）。这个空监听故意留着 —— 它承载「为什么不关」的说明，
  // 删掉之后后来人容易顺手把「点窗外关闭」加回来。
  document.getElementById("modal-overlay").addEventListener("click", e => {
    if (e.target.id !== "modal-overlay") return;
    // ★ 点窗外是随手动作（甩鼠标、点空处失焦），把它当作「放弃当前操作」＝ 让用户在毫无察觉的情况下
    //   丢掉刚改的东西，是反人性的设计（曾经全站如此，已按用户反馈改掉）。
    //   任何弹窗都不靠点窗外退出，一律点按钮。不弹 toast 提醒 —— 什么也没发生比弹提示更不打扰。
    return;
  });
  document.addEventListener("keydown", e => {
    if (document.getElementById("modal-overlay").classList.contains("hidden")) return;
    const inSampleEditor = !!document.querySelector(".sample-editor-modal");
    const inRosterEditor = !!document.querySelector("#modal-overlay .modal.roster-editor-modal");
    if (e.key === "Enter") {
      // 样本编辑弹窗 / 名单编辑弹窗 / 多行输入弹窗（模板说明）：回车用于 textarea 内换行，不关闭弹窗
      if (_styleModalState || inSampleEditor || inRosterEditor || _modalTextarea) return;
      const inp = document.getElementById("modal-input"); closeModal(inp ? inp.value : true);
    }
    if (e.key === "Escape") {
      // ★ 只有「纯告知 / 纯确认」弹窗（无任何输入控件）才允许 Esc 关闭 —— 它关掉不丢任何东西。
      //   凡带选择 / 编辑类控件的弹窗一律不响应 Esc，必须点「取消」/「删除」等按钮明确退出：
      //   Esc 与点窗外同理，都是一个键就让用户丢掉刚做的全部改动。
      //   判定用**有无输入控件**（input / textarea / select），比逐个记类名可靠 ——
      //   以后新增表单弹窗自动被这条规则覆盖，不用再往白名单里加名字。
      const box = document.querySelector("#modal-overlay .modal");
      const hasInput = !!(box && box.querySelector("input:not([type=hidden]), textarea, select"));
      if (hasInput) return;
      if (_llmFormState && _llmFormState.dirty) {
        toast("有未保存的改动，请点「取消」确认放弃", "err");
        return;
      }
      _llmFormState = null; _styleModalState = null; closeModal(null);
    }
  });
}

// ═══════════════════════════════════════════
// 视图切换
// ═══════════════════════════════════════════
function switchView(name, clearConvert) {
  const prevView = currentView;
  currentView = name;
  // 离开转换视图前把共享显示区刷进聚焦任务快照（保险：显示区 DOM 可能比快照新）
  if (prevView === "convert") snapshotConvState();
  document.querySelectorAll(".nav-item").forEach(n => n.classList.toggle("active", n.dataset.view === name));
  document.getElementById("view-design").classList.toggle("hidden", name !== "design");
  document.getElementById("view-convert").classList.toggle("hidden", name !== "convert");
  document.getElementById("view-settings").classList.toggle("hidden", name !== "settings");
  const vl = document.getElementById("view-llmdemo");
  if (vl) vl.classList.toggle("hidden", name !== "llmdemo");
  const vi = document.getElementById("view-import");
  if (vi) vi.classList.toggle("hidden", name !== "import");
  if (name === "convert") {
    refreshConvertView(null, clearConvert);
    // 仅「从其他视图切回」时恢复最新任务的过程/结果（并行模型：不中断旧流）。
    // 已在 convert 视图内的切换（如菜单「文档排版」新建任务）不恢复，保持新建态空显示区。
    if (prevView !== "convert" && convTasks.length) restoreConvProcess();
  }
  if (name === "settings") { loadLlmList(); loadSeedStyles(); loadFontFamilies(); loadRosterSettings(); initSettingsFooter(); }
  if (name === "llmdemo") loadLlmList().then(llmDemoRefreshGate);
  // 视图切换后，依据新的激活态重渲染两侧树，确保选中高亮不跨区残留
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
  renderTasks();
  // 进入模板设计视图时按当前画布宽度刷新侧栏折叠态（ResizeObserver 通常已覆盖，此处兜底）
  if (name === "design") {
    sidebarAutoApply();
    // v1.50 多页预览：分页装箱依赖真实测量（offsetTop/offsetHeight），视图 hidden 时测量全 0 →
    // 启动期 renderAll 产生的装箱无效；切入可见态时重跑 renderPreview 重新装箱
    renderPreview();
  }
}

// ── 侧边栏选中态统一收口：同一时刻只允许一个叶子菜单处于选中（current/active） ──
// 顶层导航（文档排版 / 模板设计 / 设置）按视图切换保留 active；
// 其下的任务、模板集、模板三者互斥，切换任一处即清除其余处选中。
function clearTaskSelection() {
  activeTaskId = null;
  renderTasks();
}
function clearTemplateSelection() {
  state.currentTemplateKey = null;
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
}

// 点击左侧「文档排版」= 新建任务：清空链接 / 模板集 / 模板选择
function newConvertTask() {
  clearPendingConvert();   // 用户显式新建任务 = 放弃未完成的授权续跑
  // 并行模型：新建任务不中断进行中的转换（旧流继续跑、完成时保存任务）；
  // 仅让共享显示区脱离任务流（回到新建表单态），不再显示任何任务的过程
  focusedRunId = null;
  _historyViewTask = null;   // 脱离历史查看
  _convEchoTask = false;     // 回到新建任务态
  updateReconvertBtn();      // 右上角「重新转换」隐藏（无重跑对象）
  convertLocalFiles = []; renderUploadList();
  _urlHitHint = null;
  currentTsFile = "";
  currentTsName = "";
  window.__tplPickChecked = {};
  window.__tplPickCheckedNames = {};
  activeTaskId = null;
  // 新建任务时展开输入/模板/继承卡片
  const blockInput = document.getElementById("block-input");
  const blockTemplate = document.getElementById("block-template");
  const blockInherit = document.getElementById("block-inherit");
  if (blockInput) { blockInput.classList.remove("collapsed"); blockInput.classList.remove("link-only", "local-only"); }
  if (blockTemplate) blockTemplate.classList.remove("collapsed");
  if (blockInherit) blockInherit.classList.remove("collapsed");
  // 清空过程区
  const rb = document.getElementById("result-block");
  const procEl = document.getElementById("conv-process");
  const rl = document.getElementById("result-list");
  const actionsRow = document.getElementById("actions-row");
  if (rb) rb.style.display = "none";
  if (procEl) procEl.innerHTML = "";
  if (rl) rl.innerHTML = "";
  const resultPhase = document.getElementById("result-phase");
  if (resultPhase) resultPhase.style.display = "none";   // 新建任务清空 → ⑤下载结果隐藏
  // 重置按钮位置：第④步重新可见，按钮 DOM 移回原位
  const btn = document.getElementById("btn-convert");
  if (actionsRow) actionsRow.style.display = "";
  const stepAction = btn ? btn.closest(".step-action") : null;
  if (btn && stepAction) stepAction.appendChild(btn);  // appendChild 会自动从原 parent 移走
  if (btn) { btn.disabled = false; btn.textContent = "开始转换"; }   // 新表单：按钮可用（进行中的旧任务不中断）
  switchView("convert", true);
  // 切换视图后重渲染两侧树：文档排版视图下不显示模板侧选中，模板侧也不显示任务选中
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
  renderTasks();
  // 新建任务从表单顶部开始：清空结果区让内容变短，浏览器会把上一页留下的 scrollTop
  // 夹到「新内容的底部」——不复位就是「点『文档排版』新页面停在底部、看不到地址/模板卡」。
  resetConvScroll();
}

// ── 图标（Font Awesome 6，class 方式引用；currentColor 继承文字色，可主题化） ──
const ICONS = {
  refresh:   `<i class="fas fa-rotate-right"></i>`,
  design:    `<i class="fas fa-table-cells"></i>`,
  convert:   `<i class="fas fa-arrows-rotate"></i>`,
  set:       `<i class="fas fa-folder"></i>`,
  // Doc格式转换：Word 文档图标
  docWord:   `<i class="fas fa-file-word"></i>`,
  template:  `<i class="fas fa-file"></i>`,
  newTpl:    `<i class="fas fa-circle-plus"></i>`,
  importWord:`<i class="fas fa-file-import"></i>`,
  exportSet: `<i class="fas fa-file-export"></i>`,
  rename:    `<i class="fas fa-pen-to-square"></i>`,
  del:       `<i class="fas fa-trash-can"></i>`,
  copy:      `<i class="fas fa-copy"></i>`,
  newSet:    `<i class="fas fa-folder-plus"></i>`,
  reset:     `<i class="fas fa-rotate-left"></i>`,
  save:      `<i class="fas fa-floppy-disk"></i>`,
  caretDown: `<i class="fas fa-chevron-down"></i>`,
  caretRight:`<i class="fas fa-chevron-right"></i>`,
  login:     `<i class="fas fa-circle-user"></i>`,
  logout:    `<i class="fas fa-right-from-bracket"></i>`,
  settings:  `<i class="fas fa-gear"></i>`,
  llmdemo:   `<i class="fas fa-robot"></i>`,
  // 规范 §6：状态/过程图标全部图标字体（禁止 emoji）
  file:      `<i class="fas fa-file"></i>`,
  robot:     `<i class="fas fa-robot"></i>`,
  clock:     `<i class="fas fa-clock"></i>`,
  warn:      `<i class="fas fa-triangle-exclamation"></i>`,
  // 内置固定样式标记（样式行尾小锁）
  lock:      `<i class="fas fa-lock"></i>`,
  // 标题引导段落（lead）：段落列表图标
  lead:      `<i class="fas fa-bars"></i>`,
};

// 左侧菜单展开状态（按文件记录哪些模板集已展开）
// focusSet：设计视图内左侧树高亮跟随的模板集（单击模板集行时切换；null = 跟随当前加载集 state.currentFile）。
// 模板集是纯容器（无可编辑内容），点击集行只做「折叠/展开 + 高亮焦点切换」，不加载、右侧设计页不动。
const uiState = { expandedSets: new Set(), focusSet: null };

// ── 浮动子菜单（⋯） ──
let _submenuEl = null, _submenuDocHandler = null, _submenuKeyHandler = null, _submenuAnchor = null, _submenuReposHandler = null;
function closeSubmenu() {
  if (_submenuEl) { _submenuEl.remove(); _submenuEl = null; }
  if (_submenuAnchor) { _submenuAnchor.classList.remove("dots-active"); _submenuAnchor = null; }
  if (_submenuDocHandler) { document.removeEventListener("click", _submenuDocHandler, true); _submenuDocHandler = null; }
  if (_submenuKeyHandler) { document.removeEventListener("keydown", _submenuKeyHandler, true); _submenuKeyHandler = null; }
  if (_submenuReposHandler) { window.removeEventListener("resize", _submenuReposHandler); window.removeEventListener("scroll", _submenuReposHandler, true); _submenuReposHandler = null; }
}

// 子菜单定位常量：菜单与锚点/视口边缘的留白
const SUBMENU_GAP = 4, SUBMENU_MARGIN = 8;
const clampNum = (v, min, max) => Math.min(Math.max(v, min), Math.max(min, max));

// 边界检测 + 位置纠正：保证菜单完全落在可视区域内（页面滚动/窗口缩放时由监听重算）。
// opts.align："right"（默认，菜单右缘对齐锚点右缘，左侧放不下翻转为左对齐）
//            / "topleft"（菜单左上角贴锚点，右键菜单风格，空间不足时翻转）
function positionSubmenu(menu, anchor, opts = {}) {
  const mw = menu.offsetWidth, mh = menu.offsetHeight;
  const r = anchor.getBoundingClientRect();
  const vw = window.innerWidth, vh = window.innerHeight;
  // 水平：默认右对齐；左侧放不下 → 翻转为左对齐；两侧都不足 → 贴视口边缘（不溢出）
  let left = opts.align === "topleft" ? r.left : r.right - mw;
  if (left < SUBMENU_MARGIN) left = opts.align === "topleft" ? r.right - mw : r.left;
  left = clampNum(left, SUBMENU_MARGIN, vw - mw - SUBMENU_MARGIN);
  // 垂直：优先锚点下方；下方放不下 → 翻到锚点上方；上下都不足 → 贴视口边缘（不溢出）
  let top = r.bottom + SUBMENU_GAP;
  if (top + mh > vh - SUBMENU_MARGIN) top = r.top - mh - SUBMENU_GAP;
  top = clampNum(top, SUBMENU_MARGIN, vh - mh - SUBMENU_MARGIN);
  menu.style.left = left + "px";
  menu.style.top = top + "px";
  fitSubmenuDirections(menu);
}

// 二级菜单（悬停展开）方向自适应：默认向右/向下展开，视口余量不足时翻转到左/上，防溢出屏幕
function fitSubmenuDirections(menu) {
  menu.querySelectorAll(".submenu-item.has-sub").forEach(item => {
    item.classList.remove("sub-open-left", "sub-open-up");
    const sub = item.querySelector(":scope > .submenu-sub");
    if (!sub) return;
    const r = item.getBoundingClientRect();
    // 量真实尺寸：二级菜单默认 display:none，先临时显示（不可见）测量再还原
    const prevD = sub.style.display, prevV = sub.style.visibility;
    sub.style.display = "block"; sub.style.visibility = "hidden";
    const sw = sub.offsetWidth, sh = sub.offsetHeight;
    sub.style.display = prevD; sub.style.visibility = prevV;
    if (r.right + sw > window.innerWidth - SUBMENU_MARGIN) item.classList.add("sub-open-left");
    if (r.bottom + sh > window.innerHeight - SUBMENU_MARGIN) item.classList.add("sub-open-up");
  });
}

function openSubmenu(anchor, items, opts = {}) {
  closeSubmenu();
  _submenuAnchor = anchor;
  anchor.classList.add("dots-active");
  const menu = document.createElement("div");
  menu.className = "submenu";
  menu.innerHTML = (opts.header || "") + items.map((it, i) => `
    <div class="submenu-item${it.danger ? " danger" : ""}${it.disabled ? " disabled" : ""}${it.sub ? " has-sub" : ""}" data-i="${i}">
      ${it.icon ? `<span class="submenu-icon">${it.icon}</span>` : ""}
      <span class="submenu-label">${esc(it.label)}</span>
      ${it.hint ? `<span class="submenu-hint">${esc(it.hint)}</span>` : ""}
      ${it.sub ? `<span class="submenu-arrow">▸</span>` : ""}
    </div>`).join("");
  menu.querySelectorAll(".submenu-item").forEach(elm => {
    const it = items[+elm.dataset.i];
    elm.addEventListener("click", e => {
      e.stopPropagation();
      closeSubmenu();
      if (it.disabled) return;
      if (it.onClick) it.onClick();
    });
    // 二级菜单（悬停展开）：复用同一套 submenu 样式，内嵌在父 item 下
    if (it.sub && it.sub.length) {
      const sub = document.createElement("div");
      sub.className = "submenu submenu-sub";
      sub.innerHTML = it.sub.map((s, j) => `
        <div class="submenu-item${s.danger ? " danger" : ""}${s.disabled ? " disabled" : ""}" data-j="${j}">
          ${s.icon ? `<span class="submenu-icon">${s.icon}</span>` : ""}
          <span class="submenu-label">${esc(s.label)}</span>
        </div>`).join("");
      sub.querySelectorAll(".submenu-item").forEach(subElm => {
        subElm.addEventListener("click", e => {
          e.stopPropagation();
          closeSubmenu();
          const s = it.sub[+subElm.dataset.j];
          if (s.disabled) return;
          if (s.onClick) s.onClick();
        });
      });
      elm.appendChild(sub);
    }
  });
  document.body.appendChild(menu);
  _submenuEl = menu;
  positionSubmenu(menu, anchor, opts);
  // 页面滚动 / 窗口缩放（Ctrl±、拖拽边框）时锚点位置会变 → 重定位，保证菜单始终不溢出屏幕
  _submenuReposHandler = () => { if (_submenuEl && _submenuAnchor) positionSubmenu(_submenuEl, _submenuAnchor, opts); };
  window.addEventListener("resize", _submenuReposHandler);
  window.addEventListener("scroll", _submenuReposHandler, true);
  _submenuDocHandler = e => { if (_submenuEl && !_submenuEl.contains(e.target)) closeSubmenu(); };
  _submenuKeyHandler = e => { if (e.key === "Escape") closeSubmenu(); };
  document.addEventListener("click", _submenuDocHandler, true);
  document.addEventListener("keydown", _submenuKeyHandler, true);
}

// ── 模板集导出（前端取 JSON 直接下载，无需后端改动） ──
async function exportTemplateSet(file) {
  try {
    const r = await api("/api/templateset?file=" + encodeURIComponent(file));
    const data = r.data;
    const name = (data.name || "模板集").replace(/[\\/:*?"<>|]/g, "_");
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name + ".json";
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
      } catch (e) {
    toast("导出失败：" + e.message, "err");
  }
}

// 单独导出某个模板（含其样式，导出为可被当作模板集重新导入的 {name, templates:[tpl]} 结构）
async function exportTemplate(file, tplName) {
  try {
    const r = await api("/api/templateset?file=" + encodeURIComponent(file));
    const tpl = (r.data.templates || []).find(t => t.name === tplName);
    if (!tpl) { toast("模板未找到", "err"); return; }
    const payload = { name: tpl.name, templates: [JSON.parse(JSON.stringify(tpl))] };
    const name = (tpl.name || "模板").replace(/[\\/:*?"<>|]/g, "_");
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name + ".json";
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    toast("模板导出失败：" + e.message, "err");
  }
}

// 重命名模板集：同时改名文件，保证显示名与文件名一致；重名自动加序号（服务端处理）
async function renameSetFile(file, newName) {
  try {
    const r = await api("/api/templateset/rename", { method: "POST", body: JSON.stringify({ file, newName }) });
    if (state.currentFile === file) {
      state.currentFile = r.file;
      if (state.templateSet) state.templateSet.name = r.name;
      state.dirty = false;
    }
    if (uiState.focusSet === file) uiState.focusSet = r.file;   // 改名后焦点跟随新文件名
    const st = await refreshStatus();
    renderSetTree(st);
    renderAll();
  } catch (e) {
    toast("重命名失败：" + e.message, "err");
    renderSetTree(_lastTreeStatus);
  }
}

// 模板跨集拖动：把模板从源集整体移到目标集；携带其 include 引用、但目标集没有的样式定义，避免样式丢失
async function moveTemplateToSet(srcFile, tplName, targetFile) {
  if (srcFile === targetFile) return;
  try {
    if (state.templateSet) await flushSave();   // 先落盘当前集，保证磁盘最新
    const src = await api("/api/templateset?file=" + encodeURIComponent(srcFile));
    const srcData = src.data;
    const idx = (srcData.templates || []).findIndex(t => t.name === tplName);
    if (idx < 0) { toast("源模板未找到", "err"); return; }
    const tpl = srcData.templates[idx];
    srcData.templates.splice(idx, 1);

    const target = await api("/api/templateset?file=" + encodeURIComponent(targetFile));
    const targetData = target.data;
    // v1.30 解耦：模板自包含 styles[]（定义+值随模板走），跨集移动无需 carry 补定义
    targetData.templates = targetData.templates || [];
    targetData.templates.push(tpl);

    await api("/api/templateset", { method: "POST", body: JSON.stringify({ file: targetFile, data: targetData }) });
    await api("/api/templateset", { method: "POST", body: JSON.stringify({ file: srcFile, data: srcData }) });

    // 同步内存（若源/目标是当前编辑集）
    if (state.currentFile === srcFile && state.templateSet) {
      const i = state.templateSet.templates.findIndex(t => t.name === tplName);
      if (i >= 0) state.templateSet.templates.splice(i, 1);
      if (state.currentTemplateKey && !state.templateSet.templates.some(t => t.key === state.currentTemplateKey)) {
        state.currentTemplateKey = state.templateSet.templates.length ? state.templateSet.templates[0].key : null;
        state.selectedElem = null;
      }
    }
    if (state.currentFile === targetFile && state.templateSet) {
      state.templateSet.templates.push(JSON.parse(JSON.stringify(tpl)));
    }
    const st = await refreshStatus();
    renderSetTree(st);
    renderAll();
    toast(`已将「${tplName}」移动到「${targetData.name || targetFile}」`);
  } catch (e) {
    toast("移动失败：" + e.message, "err");
  }
}

// ── 导入 Word 文档生成模板：从哪个模板集的 ⋯ 进入，就钉死保存到哪个集 ──
// importEntrySet 在点击 ⋯ 的瞬间确定，后续全程不再读 state.currentFile，
// 因此「对话框开着时手贱切了别的集」也不会改变落点。
let importEntrySet = null;
// 模板跨集拖动：dragstart 时记录 {srcFile, tplName}，drop 到目标集时整体移过去
let dragTpl = null;
function startImport(file) {
  importEntrySet = file;   // 入口集钉死
  // 顺手把当前集切到入口集，让占位/结果在树上挂在该集下（已是则免 reload）
  if (state.currentFile !== file) {
    loadTemplateSet(file).then(() => document.getElementById("file-input").click());
  } else {
    document.getElementById("file-input").click();
  }
}

// ═══════════════════════════════════════════
// 模板集树
// ═══════════════════════════════════════════
async function refreshStatus() {
  const st = await api("/api/status");
  // 全局样式库已移除：样式全部随模板集走（见 docs/样式架构-角色与模板集.md）
  return st;
}

// 就地重命名期间的状态：_inlineEditing 阻止并发的 renderSetTree 覆盖输入框；
// _lastTreeStatus 缓存最近一次状态，供取消重命名时还原行。
let _inlineEditing = false;
let _lastTreeStatus = null;
let styleRenaming = false;   // 样式行就地重命名期间，阻止 renderStyleList 重建覆盖输入框
// 收口当前编辑态：用于开始新编辑前强制完成旧的输入（保证同一时刻只有一个就地编辑行）
let _inlineEditFinish = null;

// 强制完成当前就地编辑（若有）。用于切换行 / 切换视图 / 开始新编辑前，
// 避免出现"上一行输入框还在、又开了新一行"的并存状态。
function finishInlineEditing(save) {
  if (_inlineEditing && _inlineEditFinish) _inlineEditFinish(!!save);
}

function renderSetTree(st) {
  if (_inlineEditing) return; // 正在就地编辑某个行，不重建树，避免覆盖输入框
  _lastTreeStatus = st;
  const tree = document.getElementById("sets-tree");
  if (!st.templatesets.length) {
    tree.innerHTML = `<div class="hint">暂无模板集，点「模板设计 ⋯」→ 新建模板集开始。</div>`;
    return;
  }
  let html = `<div class="tree">`;
  const curTpl = currentTemplate();
  // 仅当「模板设计」导航处于激活（即当前在模板设计视图）时，才高亮模板集/模板；
  // 在「文档排版」视图下不显示模板侧选中，避免两类菜单同时处于选中。
  const designActive = document.getElementById("nav-design") &&
    document.getElementById("nav-design").classList.contains("active");
  for (const ts of st.templatesets) {
    const isCurrent = state.currentFile === ts.file;
    const expanded = uiState.expandedSets.has(ts.file);
    // 高亮跟随的集：单击集行时 focusSet 指到被点的集；未点击过（null）则跟随当前加载集
    const focusFile = uiState.focusSet || state.currentFile;
    // 模板列表 = 该集的导入占位（全局 Map，切走模板集也不丢）+ 集内模板
    // （当前集用内存列表，含尚未落盘的新建模板；其余集用磁盘列表）
    const items = [];
    for (const ph of _importPlaceholders.values()) {
      if (ph._tsFile === ts.file || (!ph._tsFile && isCurrent && !state.currentFile)) {
        items.push({ name: ph.name, ph: true, st: ph._status });
      }
    }
    const memList = (isCurrent && state.templateSet)
      ? state.templateSet.templates.map(t => ({ name: t.name, ph: false, st: "" }))
      : null;
    const phNames = new Set(items.filter(i => i.ph).map(i => i.name));
    for (const t of (memList || (ts.templates || []).map(n => ({ name: n, ph: false, st: "" })))) {
      if (phNames.has(t.name)) continue; // 占位已单独渲染，跳过同名
      items.push(t);
    }
    const tplList = items;
    const tplNames = tplList.map(t => t.name);
    // 同一时刻仅一个叶子选中：有模板被选中时，模板集整行不再标 current
    const tplSelected = !!(isCurrent && curTpl && tplNames.includes(curTpl.name));
    const setCur = designActive && focusFile === ts.file && !tplSelected;
    // 当前模板名（用于在下面的 map 里逐行匹配高亮）。注意不能写成 `curTpl.name === t`：
    // t 是下面 map 回调里的形参，此处尚未声明，访问会抛 ReferenceError: t is not defined。
    // 模板行高亮仅当焦点集为空或等于本集——点击集行把焦点指到别处时，当前加载集的模板行不高亮。
    const curTplName = (designActive && isCurrent && (!uiState.focusSet || uiState.focusSet === ts.file) && curTpl) ? curTpl.name : null;
    // 出厂默认模板集（default.json）加 tooltip：该集会随版本升级被出厂文件覆盖，
    // 用户想保留自己的改动须另存新集（2026-09-14 拍板）。判定按文件名而非 source——
    // 它在用户态也有一份（升级覆盖），source 会是 "user"。
    const setTitle = ts.file === "default.json"
      ? ' title="系统自带模板，用户自定义模板请新建模板集"'
      : "";
    html += `
      <div class="tree-node">
        <div class="tree-set ${setCur ? "current" : ""}" data-file="${esc(ts.file)}"${setTitle}>
          <span class="tree-icon">${ICONS.set}</span>
          <span class="tree-set-name">${esc(ts.name)}</span>
          <button class="tree-dots" data-dots-set="${esc(ts.file)}" aria-label="模板集操作">⋯</button>
        </div>
        <div class="tree-templates ${expanded ? "" : "collapsed"}">
          ${tplList.map(t => {
            // 占位模板（导入中/待保存/失败）特殊显示：灰色斜体 + 状态图标；普通模板正常高亮
            const cls = t.ph
              ? (t.st === "error" ? "tree-tpl import-error"
                : t.st === "ready" ? "tree-tpl import-ready"
                : "tree-tpl importing")
              : `tree-tpl ${(curTplName !== null && t.name === curTplName) ? "current" : ""}`;
            const icon = t.ph
              ? (t.st === "error" ? ICONS.del
                : t.st === "ready" ? ICONS.save
                : ICONS.clock)
              : ICONS.template;
            return `<div class="${cls}" data-file="${esc(ts.file)}" data-tpl="${esc(t.name)}" draggable="true">
              <span class="tree-icon">${icon}</span>
              <span class="tree-tpl-name" title="${esc(t.name)}">${esc(t.name)}</span>
              <button class="tree-dots" data-dots-tpl="${esc(ts.file)}|${esc(t.name)}" aria-label="模板操作">⋯</button>
            </div>`;
          }).join("")}
        </div>
      </div>`;
  }
  html += `</div>`;
  tree.innerHTML = html;

  // 模板集整行：单击 = 折叠/展开其模板列表 + 高亮焦点切到该集（模板集是纯容器，无内容可编辑，
  // 不加载、右侧设计页不动；编辑模板需点击其下的模板行）；双击进入行内改名。
  tree.querySelectorAll(".tree-set").forEach(el => {
    const file = el.dataset.file;
    el.addEventListener("click", e => {
      if (e.target.closest(".tree-dots")) return;
      e.stopPropagation();
      const node = el.closest(".tree-node");
      const child = node.querySelector(".tree-templates");
      const willExpand = child.classList.contains("collapsed");
      child.classList.toggle("collapsed", !willExpand);
      if (willExpand) uiState.expandedSets.add(file); else uiState.expandedSets.delete(file);
      // 焦点跟随：高亮切到被点击的集。手动同步类名（不重建树，避免打断双击重命名的 dblclick 判定），
      // 规则与 renderSetTree 渲染一致：集行高亮 = 设计视图激活 && 焦点集匹配 && 无模板行选中优先。
      const designActive = document.getElementById("nav-design") &&
        document.getElementById("nav-design").classList.contains("active");
      if (!designActive) return;   // 非设计视图：不高亮（渲染规则同款门控）
      uiState.focusSet = file;
      const curTpl = currentTemplate();
      const tplSelected = !!(state.currentFile === file && curTpl && child &&
        Array.from(child.querySelectorAll(".tree-tpl")).some(x => x.dataset.tpl === curTpl.name));
      document.querySelectorAll(".tree-set.current").forEach(n => n.classList.remove("current"));
      if (!tplSelected) el.classList.add("current");
      // 焦点移到别的集时，当前加载集的模板行高亮熄灭（渲染规则同款抑制）
      if (uiState.focusSet !== state.currentFile) {
        document.querySelectorAll(".tree-tpl.current").forEach(n => n.classList.remove("current"));
      }
    });
    el.addEventListener("dblclick", e => {
      if (e.target.closest(".tree-dots")) return;
      // 单击的 loadTemplateSet 可能已重建树，闭包 el 已离屏；重新定位当前实际节点再改名
      const row = document.querySelector(`.tree-set[data-file="${CSS.escape(file)}"]`);
      if (row) beginInlineRenameSet(row);
    });
    // ── 模板跨集拖动：本集作为放置目标 ──
    el.addEventListener("dragover", e => {
      if (dragTpl && dragTpl.srcFile !== file) {
        e.preventDefault();                       // 允许放置
        e.dataTransfer.dropEffect = "move";
        el.classList.add("drag-over");
      }
    });
    el.addEventListener("dragleave", e => {
      if (e.target === el) el.classList.remove("drag-over");
    });
    el.addEventListener("drop", e => {
      el.classList.remove("drag-over");
      if (!dragTpl || dragTpl.srcFile === file) return;
      e.preventDefault();
      const src = dragTpl; dragTpl = null;
      moveTemplateToSet(src.srcFile, src.tplName, file);
    });
  });
  // 模板集 ⋯ 子菜单（出厂集只读：不提供重命名/删除；后端同样拦截）
  tree.querySelectorAll(".tree-dots[data-dots-set]").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      const file = btn.dataset.dotsSet;
      const factory = _lastTreeStatus.templatesets.some(t => t.file === file && t.source === "factory");
      const items = [
        { icon: ICONS.newTpl, label: "新建模板", onClick: async () => {
            // 新建 = 在目标模板集里加一个模板并立即打开；全程兜底，失败 toast，绝不静默停在中间态
            try {
              clearTaskSelection();
              switchView("design");
              if (state.currentFile !== file || !state.templateSet) await loadTemplateSet(file);
              uiState.focusSet = file;   // 新建模板 = 焦点跟随目标集
              await addTemplate();
            } catch (e) {
              toast("新建模板失败：" + e.message, "err");
            }
          } },
        { icon: ICONS.importWord, label: "导入docx生成模板", onClick: () => startImport(file) },
      ];
      if (!factory) items.push(
        { icon: ICONS.rename, label: "重命名模板集", onClick: () => {
            // 行内改名入口（与双击集名同行为，避免弹窗打断）；如行不在 DOM（视图已切走）则静默 no-op
            const row = document.querySelector(`.tree-set[data-file="${CSS.escape(file)}"]`);
            if (row) beginInlineRenameSet(row);
          } });
      items.push({ icon: ICONS.exportSet, label: "导出模板集", onClick: () => exportTemplateSet(file) });
      if (!factory) items.push({ icon: ICONS.del, label: "删除模板集", danger: true, onClick: () => deleteTemplateSet(file) });
      openSubmenu(btn, items);
    });
  });
  // 模板子节点：单击切换，双击重命名
  tree.querySelectorAll(".tree-tpl[data-tpl]").forEach(el => {
    const file = el.dataset.file, tplName = el.dataset.tpl;
    // 按名查全局导入占位（占位在 _importPlaceholders，不在内存模板集里）
    const findPlaceholder = () => {
      for (const p of _importPlaceholders.values()) if (p.name === tplName) return p;
      return null;
    };
    el.addEventListener("click", async e => {
      if (e.target.closest(".tree-dots")) return;
      e.stopPropagation(); // 阻止冒泡到 tree-set，避免触发折叠/展开
      clearTaskSelection(); // 切到模板设计区，清除「文档排版」任务选中
      switchView("design");
      // 占位模板：点击切到导入视图查看进度/错误/结果，不加载模板集、不进设计器
      const ph = findPlaceholder();
      if (ph) {
        if (ph._status === "ready") {
          // 待保存：从缓存恢复结果页继续修改/保存（缓存与 ready 占位一一对应）
          if (!await restoreImportView()) switchView("import");
        } else {
          switchView("import");
        }
        if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
        return;
      }
      if (file !== state.currentFile || !state.templateSet) await loadTemplateSet(file);
      uiState.focusSet = file;   // 点击模板行 = 进入该集编辑，焦点跟随
      const t = state.templateSet.templates.find(x => x.name === tplName);
      if (t) { state.currentTemplateKey = t.key; state.selectedElem = null; renderAll(); }
      // 同步左侧树：确保仅被点中的模板高亮（loadTemplateSet 可能先把首个模板标 current）
      if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
    });
    el.addEventListener("dblclick", e => {
      if (e.target.closest(".tree-dots")) return;
      if (findPlaceholder()) return; // 占位不可重命名
      beginInlineRenameTpl(el, file, tplName);
    });
  });
  // 模板 ⋯ 子菜单
  tree.querySelectorAll(".tree-dots[data-dots-tpl]").forEach(btn => {
    btn.addEventListener("click", async e => {
      e.stopPropagation();
      const [file, tplName] = btn.dataset.dotsTpl.split("|");
      // 占位模板：仅提供取消/删除（导入中取消会丢弃返回结果；失败态直接删除）
      let ph = null;
      for (const p of _importPlaceholders.values()) { if (p.name === tplName) { ph = p; break; } }
      if (ph) {
        openSubmenu(btn, [
          {
            icon: ICONS.del,
            label: ph._status === "importing" ? "取消导入" : (ph._status === "ready" ? "放弃导入（不保存）" : "删除占位"),
            danger: true,
            onClick: () => {
              if (ph._status === "ready") clearImportViewCache();  // ready 占位持有对应结果缓存，一并清理（否则重启会恢复出占位）
              removeImportPlaceholder(ph.key);
              if (currentView === "import") switchView("convert"); // 放弃后离开导入视图，回到干净态
            },
          },
        ]);
        return;
      }
      if (file !== state.currentFile) await loadTemplateSet(file);
      const t = state.templateSet.templates.find(x => x.name === tplName);
      if (!t) return;
      openSubmenu(btn, [
        { icon: ICONS.rename, label: "模板改名", onClick: () => renameTemplate(t.key) },
        { icon: ICONS.copy, label: "模板复制", onClick: () => duplicateTemplate(t.key) },
        { icon: ICONS.exportSet, label: "模板导出", onClick: () => exportTemplate(file, tplName) },
        { icon: ICONS.del, label: "模板删除", danger: true, onClick: () => removeTemplateByKey(t.key) },
      ]);
    });
  });
  // 模板拖拽：dragstart 记录源集 + 模板名，供 drop 到目标集时使用
  tree.querySelectorAll(".tree-tpl[draggable='true']").forEach(el => {
    el.addEventListener("dragstart", e => {
      dragTpl = { srcFile: el.dataset.file, tplName: el.dataset.tpl };
      try { e.dataTransfer.setData("text/plain", el.dataset.tpl); } catch (_) { /* 某些浏览器需同步设值 */ }
      e.dataTransfer.effectAllowed = "move";
    });
    el.addEventListener("dragend", () => {
      dragTpl = null;
      document.querySelectorAll(".tree-set.drag-over").forEach(n => n.classList.remove("drag-over"));
    });
  });
}

// 把当前内存中的模板列表同步回 _lastTreeStatus（左侧树的数据源），
// 使新增/改名/删除模板后，侧边栏能立即反映，而不必等保存+刷新状态。
function syncCurrentSetInStatus() {
  if (!_lastTreeStatus || !state.templateSet) return;
  const ts = _lastTreeStatus.templatesets.find(t => t.file === state.currentFile);
  if (ts) ts.templates = state.templateSet.templates.map(t => t.name);
}

// ── 就地重命名：模板集（双击整行 → 行内输入框） ──
function beginInlineRenameSet(el) {
  const file = el.dataset.file;
  if (_lastTreeStatus && _lastTreeStatus.templatesets.some(t => t.file === file && t.source === "factory")) {
    toast("出厂自带模板集不可重命名", "err");
    return;
  }
  // 收口旧编辑态：同一时刻只允许一个就地编辑行
  finishInlineEditing(true);
  const nameSpan = el.querySelector(".tree-set-name");
  const oldName = nameSpan ? nameSpan.textContent : "";
  el.classList.add("editing");
  _inlineEditing = true;
  el.innerHTML = `<input class="tree-rename-input" type="text" value="${esc(oldName)}" />`;
  const input = el.querySelector("input");
  input.focus();
  input.select();
  let done = false;
  const finish = async (save) => {
    if (done) return;
    done = true;
    _inlineEditFinish = null;
    _inlineEditing = false;
    const v = input.value.trim();
    if (!save || !v || v === oldName) { renderSetTree(_lastTreeStatus); return; }
    await renameSetFile(file, v);
  };
  _inlineEditFinish = finish;
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); finish(true); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); finish(false); }
  });
  input.addEventListener("blur", () => finish(true));
}

// ── 就地重命名：模板（双击整行 → 行内输入框） ──
async function beginInlineRenameTpl(el, file, tplName) {
  // 收口旧编辑态：同一时刻只允许一个就地编辑行
  finishInlineEditing(true);
  // 若目标集尚未加载，先加载（await 期间 click/loadTemplateSet 可能重建树，闭包里的 el 会失效）
  if (file !== state.currentFile || !state.templateSet) await loadTemplateSet(file);
  // 重新定位当前行：双击前两次 click 已触发 renderSetTree 重建树，旧 el 已脱离文档；
  // 输入框必须挂在当前 DOM 行上，否则界面不显示改名输入框
  const row = document.querySelector(`.tree-tpl[data-file="${CSS.escape(file)}"][data-tpl="${CSS.escape(tplName)}"]`);
  if (!row) return;
  el = row;
  // 先同步取出旧名并进入编辑态（_inlineEditing 必须在任何 await 之前置位，
  // 否则并发的 renderSetTree 会重建树、把当前行 DOM 拆掉）。
  const nameSpan = el.querySelector(".tree-tpl-name");
  const oldName = nameSpan ? nameSpan.textContent : tplName;
  el.classList.add("editing");
  _inlineEditing = true;
  el.innerHTML = `<input class="tree-rename-input" type="text" value="${esc(oldName)}" />`;
  const input = el.querySelector("input");
  input.focus();
  input.select();
  const tpl = state.templateSet.templates.find(x => x.name === tplName);
  if (!tpl) { _inlineEditing = false; _inlineEditFinish = null; renderSetTree(_lastTreeStatus); return; }
  let done = false;
  const finish = async (save) => {
    if (done) return;
    done = true;
    _inlineEditFinish = null;
    _inlineEditing = false;
    const v = input.value.trim();
    if (!save || !v || v === oldName) { renderSetTree(_lastTreeStatus); return; }
    tpl.name = v;
    markDirty();
    // 立即落盘 + 用最新磁盘状态刷新树，与弹窗改名行为一致
    try { await persistCurrentSet(); } catch (e) { /* 已 toast */ }
    let st = null;
    try { st = await refreshStatus(); } catch (e) { toast("刷新模板集列表失败：" + e.message, "err"); }
    syncCurrentSetInStatus();
    renderSetTree(st || _lastTreeStatus);
    renderAll();
      };
  _inlineEditFinish = finish;
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); finish(true); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); finish(false); }
  });
  input.addEventListener("blur", () => finish(true));
}

// ── 模板集重命名（双击，改名后落盘） ──
async function renameTemplateSet(file) {
  try {
    const r = await api("/api/templateset?file=" + encodeURIComponent(file));
    const newName = await modalPrompt("重命名模板集", { value: r.data.name || "", placeholder: "模板集名称" });
    if (!newName || newName === r.data.name) return;
    await renameSetFile(file, newName);
  } catch (e) {
    toast("重命名失败：" + e.message, "err");
  }
}

// ── 模板集删除 ──
async function deleteTemplateSet(file) {
  try {
    const st = await api("/api/status");
    const ts = st.templatesets.find(t => t.file === file);
    if (ts && ts.source === "factory") {
      toast("出厂自带模板集不可删除", "err");
      return;
    }
    const name = ts ? ts.name : file;
    const ok = await modalConfirm("删除模板集", `确定删除模板集「${name}」？<br>（删除后不可恢复，文件将被永久移除）`, "删除");
    if (!ok) return;
    await api("/api/templateset?file=" + encodeURIComponent(file), { method: "DELETE" });
    if (state.currentFile === file) {
      state.templateSet = null;
      state.currentFile = "";
      state.currentTemplateKey = null;
      state.selectedElem = null;
      state.dirty = false;
      renderAll();
    }
    if (uiState.focusSet === file) uiState.focusSet = null;   // 被删集的焦点失效，回落跟随
    const st2 = await refreshStatus();
    renderSetTree(st2);
      } catch (e) {
    toast("删除失败：" + e.message, "err");
  }
}

// ── 模板重命名（弹窗改名，立即落盘 + 刷新左侧树） ──
async function renameTemplate(key) {
  const tpl = state.templateSet.templates.find(t => t.key === key);
  if (!tpl) return;
  const newName = await modalPrompt("重命名模板", { value: tpl.name || "", placeholder: "模板名称" });
  if (!newName || newName === tpl.name) return;
  const oldName = tpl.name;
  tpl.name = newName;
  markDirty();
  // 立即落盘，避免防抖窗口内状态丢失；随后用最新磁盘状态刷新左侧树，
  // 保证改名结果立即可见（旧缓存 _lastTreeStatus 渲染可能不反映改名）。
  try { await persistCurrentSet(); } catch (e) { /* 已 toast */ }
  let st = null;
  try { st = await refreshStatus(); } catch (e) { toast("刷新模板集列表失败：" + e.message, "err"); }
  syncCurrentSetInStatus();
  renderSetTree(st || _lastTreeStatus);
  renderAll();
  toast(`已重命名模板为「${newName}」`);
}

// ── 模板复制（同集内深拷贝一份，命名加序号，插在源模板之后并切到新模板） ──
async function duplicateTemplate(key) {
  if (!state.templateSet) return;
  const idx = state.templateSet.templates.findIndex(t => t.key === key);
  if (idx < 0) return;
  const src = state.templateSet.templates[idx];
  // key 唯一：删除过模板后 length+1 可能撞已有 key（同 addTemplate 防重）
  const usedKeys = new Set(state.templateSet.templates.map(t => t.key));
  const usedNames = new Set(state.templateSet.templates.map(t => t.name));
  let n = state.templateSet.templates.length + 1;
  while (usedKeys.has("tpl" + n)) n++;
  // 命名加序号：先脱掉尾部既有 _N，避免「_2_2」式叠加；仍撞名则递增
  const base = String(src.name || "模板").replace(/_\d+$/, "");
  let i = 2, name = `${base}_${i}`;
  while (usedNames.has(name)) { i++; name = `${base}_${i}`; }
  const copy = JSON.parse(JSON.stringify(src));   // 深拷贝：styles/include/page/footer 全随模板走
  copy.key = "tpl" + n;
  copy.name = name;
  await ensureTplStyles(copy);   // 兼容旧文件缺 styles[]：补齐为自包含
  state.templateSet.templates.splice(idx + 1, 0, copy);
  state.currentTemplateKey = copy.key;
  state.selectedElem = null;
  markDirty();
  // 立即落盘 + 刷新左侧树，复制结果立即可见（与改名/删除一致）
  try { await persistCurrentSet(); } catch (e) { toast("复制模板保存失败：" + e.message, "err"); }
  if (state.currentFile) uiState.expandedSets.add(state.currentFile);
  syncCurrentSetInStatus();
  let st = null;
  try { st = await refreshStatus(); } catch (e) { toast("刷新模板集列表失败：" + e.message, "err"); }
  renderSetTree(st || _lastTreeStatus);
  renderAll();
  toast(`已复制模板为「${name}」`);
}

async function removeTemplateByKey(key) {
  if (!state.templateSet) return;
  finishInlineEditing(false);   // 强制收口残留的就地重命名编辑态，避免 renderSetTree 被 _inlineEditing 闸门跳过
  const idx = state.templateSet.templates.findIndex(t => t.key === key);
  if (idx < 0) return;
  const name = state.templateSet.templates[idx].name;
  const ok = await modalConfirm("删除模板", `确定删除模板「${name}」？`, "删除");
  if (!ok) return;
  state.templateSet.templates.splice(idx, 1);
  if (state.currentTemplateKey === key) {
    state.currentTemplateKey = state.templateSet.templates.length ? state.templateSet.templates[0].key : null;
    state.selectedElem = null;
  }
  markDirty();
  // 立即落盘，避免防抖窗口内删除状态丢失；随后用最新磁盘状态刷新左侧树，
  // 保证删除结果立即可见（旧缓存 _lastTreeStatus 渲染可能不反映删除）。
  try { await persistCurrentSet(); } catch (e) { /* 已 toast */ }
  let st = null;
  try { st = await refreshStatus(); } catch (e) { toast("刷新模板集列表失败：" + e.message, "err"); }
  syncCurrentSetInStatus();
  renderSetTree(st || _lastTreeStatus);
  renderAll();
  toast(`已删除模板「${name}」`);
}

async function loadTemplateSet(file) {
  await flushSave();   // 切换前先把当前集待保存改动落盘（防抖窗口内也不丢）
  const r = await api("/api/templateset?file=" + encodeURIComponent(file));
  state.templateSet = r.data;
  // v1.30 解耦：模板自包含 styles[]；旧文件缺 styles[] 时用种子兜底补齐（加载即自包含化）
  for (const t of (state.templateSet.templates || [])) { await ensureTplStyles(t); }
  state.currentFile = file;
  state.currentTemplateKey = r.data.templates.length ? r.data.templates[0].key : null;
  state.selectedElem = null;
  uiState.focusSet = null;   // 加载后焦点跟随当前集（点击集行才会把焦点指到别处）
  clearParaSelection();   // 切换模板集时清除段落选中态（避免残留虚线框+selectedParaKey脏状态）
  state.dirty = false;
  renderAll();
  // 列表刷新失败不阻断加载（避免中断「新建模板」链路）
  let st = null;
  try { st = await refreshStatus(); } catch (e) { /* 忽略：树稍后由调用方兜底渲染 */ }
  if (st) renderSetTree(st);
}

async function newTemplateSet() {
  await flushSave();   // 先落盘当前集（若有未保存改动）
  // v1.30 解耦：模板集为纯容器，不持任何样式；种子仅供新建模板时深拷贝（预热缓存）
  await ensureSeedStyles();
  // 显示名与文件名保持一致：先算唯一基名（重名自动加序号），name 直接等于该基名，
  // 后续 persistCurrentSet 按 name 派生文件名时二者天然相等。
  const _st0 = await refreshStatus();
  const _used0 = new Set((_st0.templatesets || []).map(t => t.file));
  let _base = "新模板集", _nf = _base + ".json", _ni = 1;
  while (_used0.has(_nf)) { _ni++; _nf = `${_base}_${_ni}.json`; }
  const _setName = _nf.replace(/\.json$/, "");
  state.templateSet = { name: _setName, templates: [] };
  state.currentFile = "";
  state.currentTemplateKey = null;
  state.selectedElem = null;
  clearParaSelection();   // 新建模板集时清除段落选中态
  // 立即落盘（派生文件名）+ 刷新左侧树，新模板集立即可见（不依赖防抖自动保存）
  try { await persistCurrentSet(); } catch (e) { toast("新建模板集保存失败：" + e.message, "err"); }
  uiState.focusSet = null;   // 新建后焦点跟随新集
  let st = null;
  try { st = await refreshStatus(); } catch (e) { toast("刷新模板集列表失败：" + e.message, "err"); }
  if (state.currentFile) uiState.expandedSets.add(state.currentFile);
  renderSetTree(st || _lastTreeStatus);
  // 新建模板集后进入设计器（renderSetTree 已把 _lastTreeStatus 更新为最新状态，
  // switchView 内部重绘树时不会丢失新集）；菜单项点击不再切视图，故此处显式切换
  switchView("design");
  renderAll();
}

// 显式「保存模板集」= 立即落盘（替代原保存按钮；自动保存下用于手动立即保存）
async function saveTemplateSet() {
  return flushSave();
}

// ═══════════════════════════════════════════
// 模板管理（模板集内的模板）
// ═══════════════════════════════════════════
// 模板集写串行化：同一时刻只允许一个写请求在途（防抖自动保存 × 导入保存 × 新建/删除/重命名模板集
// 并发落盘时，「后完成的写」会覆盖「先完成的写」→ 其他模板/新导入模板丢失（丢更新）。
// 所有模板集 POST 经此队列排队执行；入队项在执行时才序列化 state.templateSet（执行时快照），
// 故队尾写入的必然是最新状态，串行化后天然「最后一次写=全量最新」。
let _setWriteChain = Promise.resolve();
function enqueueSetWrite(fn) {
  const p = _setWriteChain.then(fn);
  _setWriteChain = p.catch(() => {});   // 单次写失败不中断队列
  return p;
}

// 立即把当前内存模板集落盘（无弹窗、无确认）；currentFile 为空时按集名派生唯一文件名
async function persistCurrentSet() {
  if (!state.templateSet) return;
  let file = state.currentFile;
  if (!file) {
    const st = await refreshStatus();
    const used = new Set(st.templatesets.map(t => t.file));
    const base = (state.templateSet.name || "模板集").replace(/[\\/:*?"<>|]/g, "_").trim() || "模板集";
    file = base + ".json";
    let i = 1;
    while (used.has(file)) { i++; file = `${base}_${i}.json`; }
  }
  try {
    await enqueueSetWrite(() => api("/api/templateset", { method: "POST", body: JSON.stringify({ file, data: state.templateSet }) }));
    state.currentFile = file;
    state.dirty = false;
    setSaveStatus("已保存");
  } catch (e) {
    setSaveStatus("保存失败");
    toast("保存失败：" + e.message, "err");
  }
}

async function addTemplate() {
  if (!state.templateSet) return;
  // key/name 唯一：删除过模板后 length+1 可能撞已有 key（如删 tpl2 后新 tpl2），
  // 而 currentTemplate() 按 key 查第一个匹配，会导致打开的不是新建的模板
  const usedKeys = new Set(state.templateSet.templates.map(t => t.key));
  const usedNames = new Set(state.templateSet.templates.map(t => t.name));
  let n = state.templateSet.templates.length + 1;
  while (usedKeys.has("tpl" + n) || usedNames.has("模板" + n)) n++;
  const key = "tpl" + n;
  // v1.30 解耦：新模板从种子深拷贝自包含 styles[]（定义+值随模板走）
  // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults（种子深拷贝自带），不再写顶层扁平字段
  const seeds = await ensureSeedStyles();
  const seedKeys = seeds.map(s => s.key);
  // 页面默认值同源于种子「页面」样式（GB/T 9704：22 行网格 linePitch 579；页码锚点 numberGap 7mm）。
  // 硬编码会与种子漂移（改种子不生效），故从种子取；取不到才落常量兜底。
  const seedPage = (seeds.find(s => s.category === "page") || {}).defaults || {};
  state.templateSet.templates.push({
    key, name: "模板" + n,
    styles: seeds.map(seedToTemplateDef),
    include: [...seedKeys],
    page: {
      width: 11906, height: 16838,
      marginTop: 2098, marginRight: 1474, marginBottom: 1984, marginLeft: 1587,
      linePitch: seedPage.linePitch != null ? seedPage.linePitch : 579,
      numberGap: seedPage.numberGap != null ? seedPage.numberGap : DEFAULT_NUMBER_GAP,
      pageNumber: { format: "dash", align: "center", first: true },
    },
    footer: {},
  });
  state.currentTemplateKey = key;
  state.selectedElem = null;
  try {
    await persistCurrentSet();
  } catch (e) {
    toast("新建模板保存失败：" + e.message, "err");
  }
  // 落盘后展开当前集 + 同步内存模板列表到树数据源（与改名/删除一致），新模板立即可见
  if (state.currentFile) uiState.expandedSets.add(state.currentFile);
  syncCurrentSetInStatus();
  let st = null;
  try {
    st = await refreshStatus();
  } catch (e) {
    toast("刷新模板集列表失败：" + e.message, "err");
  }
  renderSetTree(st || _lastTreeStatus);
  renderAll();
}

// 预览区标题（模板设计：模板集名 · 模板名；圆点分隔符独立 span 便于样式定制）
function renderTemplateTitle() {
  const title = document.getElementById("doc-title");
  const sep = document.getElementById("doc-sep");
  const meta = document.getElementById("doc-meta");
  if (!state.templateSet) {
    title.textContent = "模板设计：未选择模板集";
    sep.textContent = "";
    meta.textContent = "";
    return;
  }
  title.textContent = `模板设计：${state.templateSet.name}`;
  const tpl = currentTemplate();
  if (tpl) {
    sep.textContent = "·";
    meta.textContent = tpl.name;
  } else {
    sep.textContent = "";
    meta.textContent = "";
  }
  // 「模板说明」按钮的 hover 提示显示当前说明（无需打开弹窗即可查看；无说明时为功能说明）
  const descBtn = document.getElementById("btn-tpl-desc");
  if (descBtn) {
    const d = ((tpl && tpl.description) || "").trim();
    descBtn.title = d ? `当前说明：${d}` : "编辑当前模板的说明（悬停模板名时可见）";
  }
}

// 模板说明编辑（模板设计页工具条「模板说明」）：写入 templates[].description 并落盘。
// 说明的转换侧出口：fillTplPick 在「文档排版 → 选择模板」把 description 挂到模板名 title（tooltip）。
// 留空保存 = 清除说明（delete 字段，不留空串）。
async function editTemplateDesc() {
  const tpl = currentTemplate();
  if (!tpl) { toast("请先在左侧选择一个模板", "err"); return; }
  const v = await modalPromptArea(`模板说明 · ${tpl.name || tpl.key}`, {
    value: tpl.description || "",
    placeholder: "一句话说明这个模板的用途或适用场景",
    message: "保存后，在「文档排版」选择模板时，鼠标悬停模板名称即可看到（留空则清除说明）",
    maxlength: 200,
  });
  if (v === null) return;   // 取消
  const desc = (v || "").trim();
  if (desc) tpl.description = desc; else delete tpl.description;
  await persistCurrentSet();
  renderTemplateTitle();     // 刷新按钮 hover 提示（显示新说明）
  toast(desc ? "已保存模板说明" : "已清除模板说明", "ok");
}

// ═══════════════════════════════════════════
// 元素样式面板（Word 风格：样式列表 + 展开编辑 + 启用开关）
// ═══════════════════════════════════════════
function fmtStyleSummary(spec) {
  if (!spec) return "";
  const parts = [];
  if (spec.font) parts.push(spec.font);
  if (spec.size) parts.push(ptToCnSize(spec.size));
  return parts.join(" ") || "默认";
}

// ── 样式分类（category）分组：段落 / 表格（按 group 二级分组）/ 页面 ──
const STYLE_CATS = [
  { cat: "text",  label: "段落样式", canAdd: true },
  { cat: "char",  label: "字符样式", canAdd: true },
  { cat: "table", label: "表格样式", canAdd: true },
  { cat: "page",  label: "页面样式", canAdd: false },
];
const TABLE_GROUP_DEFAULT = "数据表格";   // 表格样式缺省二级组（新建默认落此组）
let styleFilterText = "";                 // 样式搜索关键词（纯 UI 态，不落盘）

function styleCategory(def) {
  const c = def && def.category;
  return (c === "table" || c === "page" || c === "char") ? c : "text";
}
function tableGroupOf(def) {
  return (def && def.group) || TABLE_GROUP_DEFAULT;
}

// ── 段落选中 ↔ 右侧面板联动 ──
// 标题引导段落的内部样式 key（mode:"lead"）；在面板里作为固定行呈现，
// 不进段落样式列表（B 方案：不生成 Word 段落样式）。
const LEAD_STYLE_KEY = "标题引导段落";
// 兜底定义（种子未加载、或模板里没有引导段样式时，面板固定行与补建都用它；值取自种子默认）
const FALLBACK_LEAD_DEF = {
  key: LEAD_STYLE_KEY, name: LEAD_STYLE_KEY, category: "text",
  matchType: "none", match: "", mode: "lead", next: "正文段落",
  defaults: { font: "楷体", size: 16, align: "justify", line: 28, bold: true, indentChars: 2, spaceBefore: 0, spaceAfter: 0 },
};

function isLeadKey(key) {
  const d = styleDef(key);
  return !!(d && d.mode === "lead");
}
// 手动设引导段时，若当前模板没有引导段样式定义 → 按种子/兜底补建（并纳入 include）
// 老模板（v1.2x 之前创建）普遍缺这一条，不补建的话手动标记无处落地、预览也不会分轨。
async function ensureLeadStyleDef() {
  const tpl = currentTemplate();
  if (!tpl) return null;
  let changed = false;
  const has = key => allStyleDefs().some(d => d.key === key);
  const seeds = await ensureSeedStyles();
  // 补建清单：仅引导段样式（标题句字体/加粗直接由该样式定义承载，2026-09-01 起无独立字符样式）
  const need = [
    { key: LEAD_STYLE_KEY, fallback: FALLBACK_LEAD_DEF, pick: s => s.mode === "lead" },
  ];
  let leadDef = null;
  if (!Array.isArray(tpl.styles)) tpl.styles = [];
  for (const n of need) {
    if (has(n.key)) { if (n.key === LEAD_STYLE_KEY) leadDef = allStyleDefs().find(d => d.key === LEAD_STYLE_KEY); continue; }
    const seed = seeds.find(n.pick) || n.fallback;
    const def = seedToTemplateDef(seed);   // 字体按当前开关落值（种子存的是字族名）
    tpl.styles.push(def);
    ensureStyleOrderComplete(tpl);   // 2026-09-01：补建样式同步进 styleOrder，防残缺甩尾
    if (!Array.isArray(tpl.include)) tpl.include = [];
    if (!tpl.include.includes(def.key)) tpl.include.push(def.key);
    // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults（补建样式自带），不再写顶层扁平字段
    if (n.key === LEAD_STYLE_KEY) leadDef = def;
    changed = true;
  }
  if (changed) markDirty();
  return leadDef;
}
// 样式 key 所在的面板分组 id（一级组：text/char/table/page；表格为二级组 table:<组名>）
function styleGroupIdOf(key) {
  if (key === "footer") return "page";
  const d = styleDef(key);
  if (!d) return null;
  const cat = styleCategory(d);
  return cat === "table" ? "table:" + tableGroupOf(d) : cat;
}
// 展开样式所在分组（折叠态下样式行高度为 0，不展开则滚动定位等于没滚）
function expandStyleGroupFor(key) {
  const gid = styleGroupIdOf(key);
  if (!gid) return;
  if (!state.collapsed) state.collapsed = {};
  state.collapsed[gid] = false;
}
// 段落生效样式 → 面板联动对象：
// 引导段（lead）样式行就在面板里（2026-08-30 起按普通段落样式行渲染），联动指回它自身。
function panelLinkKeyOf(effKey) {
  return effKey;
}
// 段落选中/套用样式后，右侧面板：展开所在组 + 展开该样式编辑表单（联动深度 = D1）
function syncPanelToPara(effKey) {
  const linkKey = panelLinkKeyOf(effKey);
  state.selectedElem = linkKey || null;
  expandStyleGroupFor(linkKey);
  return linkKey;
}
// 滚动样式行到可见区（列表重建后调用）
function scrollStyleRowIntoView(key) {
  if (!key) return;
  const row = document.querySelector(`#style-list .style-item[data-elem="${CSS.escape(key)}"]`);
  if (row) row.scrollIntoView({ block: "nearest", behavior: "smooth" });
}
// 行摘要：spec（模板自定义值）叠在 def.defaults 之上取有效值
function styleSummaryOf(def, key, tpl) {
  if (key === "footer") return fmtStyleSummary(tpl.footer || {});
  if (styleCategory(def) === "page") {
    const p = tpl.page || {};
    const mm = v => Math.round((parseFloat(v) || 0) / 1440 * 25.4);
    return `${mm(p.width || 11906)}×${mm(p.height || 16838)}mm`;
  }
  // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults，标签直接读条目 defaults
  const eff = (def && def.defaults) || {};
  if (styleCategory(def) === "char") {
    const parts = [];
    parts.push(eff.font ? eff.font : "跟随段落");
    if (eff.bold) parts.push("加粗");
    if (eff.size) parts.push(ptToCnSize(eff.size));   // 字号空 = 跟随段落（不显示）
    return parts.join(" ") || "默认";
  }
  return fmtStyleSummary(eff);
}

function renderStyleList() {
  if (styleRenaming) return; // 样式行就地重命名期间，不重建列表，避免覆盖输入框
  const tpl = currentTemplate();
  const el = document.getElementById("style-list");
  if (!tpl) { el.innerHTML = ""; return; }
  const include = tpl.include || [];

  // 样式全部随模板集走：本模板集「固定角色 + 自定义样式」按 category 分组渲染
  const defs = allStyleDefs();
  // 面板排序（2026-09-01 修复）：组内顺序 = 用户拖出的 styleOrder（模板字段，随模板落盘）优先；
  // 不在 styleOrder 的 key 不再一律甩尾（旧实现 1000+seedOrder → 残缺 styleOrder 时标题 2/正文等
  // 后加样式全被排到尾部，是面板乱序根因），改为按种子权威序插入「seed 序前驱在 styleOrder 中的位置」之后，
  // 同锚点按 seed 序稳定 —— 存量残缺 styleOrder 模板立即恢复公序，且不依赖用户重拖。
  // 识别链不受影响：matchers 仍按模板 styles[] 顺序（convert.js），styleOrder 仅面板展示序。
  // 种子基准用 SEED_STYLE_CACHE（懒加载缓存，全应用共用）优先，seedStyles（设置页专用）兜底：
  // 模板设计页未进过设置页时 seedStyles 为空，锚点插入会失效（缺失 key 仍被甩尾）。
  if (!SEED_STYLE_CACHE) {
    ensureSeedStyles().then(() => {
      if (document.getElementById("style-list")) renderStyleList();
    });
  }
  const seedOrder = new Map((SEED_STYLE_CACHE || seedStyles || []).map((s, i) => [s.key, i]));
  const orderIdx = new Map(((tpl && tpl.styleOrder) || []).map((k, i) => [k, i]));
  const rankOf = (key) => {
    if (orderIdx.has(key)) return orderIdx.get(key);
    const si = seedOrder.get(key);
    if (si === undefined) return 1000 + 9999;      // 自定义样式且未拖过：尾部兜底
    let anchor = -1;
    for (const [k, i] of orderIdx) {
      const sk = seedOrder.get(k);
      if (sk !== undefined && sk < si && i > anchor) anchor = i;
    }
    return anchor + 0.5;
  };
  const byPanelOrder = (a, b) => {
    const ra = rankOf(a.key), rb = rankOf(b.key);
    if (ra !== rb) return ra - rb;
    return (seedOrder.get(a.key) ?? 9999) - (seedOrder.get(b.key) ?? 9999);
  };

  let html = "";

  if (!state.templateSet) {
    el.innerHTML = `<div class="style-group-empty">未加载模板集，样式暂不可用</div>`;
    return;
  }

  for (const c of STYLE_CATS) {
    // 段落样式组：不过滤 lead（2026-08-30 用户定调：标题引导段落不再特殊处理，按普通段落样式行渲染）
    const items = defs.filter(d => styleCategory(d) === c.cat).sort(byPanelOrder);
    const col = !!state.collapsed[c.cat];
    html += `<div class="style-group${col ? " collapsed" : ""}" data-group="${c.cat}">
      <div class="style-group-h">
        <span class="group-h-tog">${col ? "▸" : "▾"}</span>
        <span class="group-h-title">${c.label}</span>
        ${c.canAdd ? `<button class="group-h-add" data-add-cat="${c.cat}" title="${c.cat === "table" ? "新建表格分组" : "新建样式"}" aria-label="${c.cat === "table" ? "新建表格分组" : "新建样式"}">＋</button>` : ""}
      </div>
      <div class="style-group-body">`;
    if (c.cat === "table") {
      // 表格样式二级分组（group：数据表格 / 版记 / …）
      // 分组来源：样式推导组 ∪ 模板集 tableGroups（新建的空分组，可持久化）
      const groups = [];
      for (const d of items) { const g = tableGroupOf(d); if (!groups.includes(g)) groups.push(g); }
      const extraGroups = ((state.templateSet && state.templateSet.tableGroups) || []).filter(g => !groups.includes(g));
      if (!groups.length && !extraGroups.length) html += `<div class="style-group-empty">暂无表格样式，点组头 ＋ 新建</div>`;
      // 新建分组占位：无弹窗，直接在表格组最上方插入空分组并进入改名
      if (state.pendingTableGroup) {
        html += `<div class="style-subgroup table-group-new" data-group="table:">
          <div class="style-subgroup-h">
            <span class="group-h-tog">▾</span>
            <input class="group-rename-input" data-new-table-group placeholder="输入分组名，回车确认（Esc 取消）">
          </div>
          <div class="style-group-body"><div class="style-group-empty">暂无样式</div></div>
        </div>`;
      }
      const renderSub = (g) => {
        const gid = "table:" + g;
        const gcol = !!state.collapsed[gid];
        const gitems = items.filter(d => tableGroupOf(d) === g).sort(byPanelOrder);
        html += `<div class="style-subgroup${gcol ? " collapsed" : ""}" data-group="${esc(gid)}">
          <div class="style-subgroup-h">
            <span class="group-h-tog">${gcol ? "▸" : "▾"}</span>
            <span class="group-h-title">${esc(g)}</span>
            <button class="group-h-add" data-add-cat="table" data-add-group="${esc(g)}" title="在该组新建样式" aria-label="在该组新建样式">＋</button>
          </div>
          <div class="style-group-body">`;
        if (!gitems.length) html += `<div class="style-group-empty">暂无样式，点组头 ＋ 新建</div>`;
        for (const d of gitems) html += styleRowHtml(d, "set", include, tpl);
        html += `</div></div>`;
      };
      for (const g of groups) renderSub(g);
      for (const g of extraGroups) renderSub(g);
    } else if (c.cat === "page") {
      // 页面样式：三子组（纸张 / 页边距 / 页码），非样式行（不参与 include/tpl[key]/styleOrder）
      html += pageSubgroupsHtml(tpl);
    } else {
      if (!items.length) html += `<div class="style-group-empty">${c.cat === "text" ? "暂无段落样式，点组头 ＋ 新建" : "暂无页面样式"}</div>`;
      for (const d of items) html += styleRowHtml(d, "set", include, tpl);
    }
    html += `</div></div>`;
  }

  el.innerHTML = html;

  // 新建表格分组占位输入：回车确认建组（空名/Esc 取消），提交后落盘
  const newGroupInput = el.querySelector("input[data-new-table-group]");
  if (newGroupInput) {
    let done = false;
    const finish = (commit) => {
      if (done) return; done = true;
      const name = (newGroupInput.value || "").trim();
      if (commit && name && state.templateSet) {
        state.templateSet.tableGroups = Array.isArray(state.templateSet.tableGroups)
          ? state.templateSet.tableGroups : [];
        if (!state.templateSet.tableGroups.includes(name)) {
          state.templateSet.tableGroups.push(name);
          state.collapsed["table:" + name] = false;
          markDirty();
        }
      }
      state.pendingTableGroup = false;
      renderStyleList();
    };
    newGroupInput.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); finish(true); }
      else if (e.key === "Escape") { e.preventDefault(); finish(false); }
    });
    newGroupInput.addEventListener("blur", () => finish(true));
    newGroupInput.focus();
    newGroupInput.select();
  }

  // 组头 ＋：一级组（段落/字符）→ 下拉菜单（新建空样式 / 从默认样式中选）；表格一级组 → 新建表格分组；表格二级组 → 在该组新建样式
  //   data-add-cat=text/char/table/page → 一级组；cat=table + data-add-group=xxx → 表格二级分组
  el.querySelectorAll(".group-h-add[data-add-cat]").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();          // 不触发组折叠
      const cat = btn.dataset.addCat;
      const group = btn.dataset.addGroup || "";
      if (group) createStyleInline(cat, group);      // 表格二级组：在本组新建样式
      else if (cat === "table") createTableGroup();  // 表格一级组：新建表格分组
      else openAddStyleMenu(btn, cat);               // 段落/字符一级组：下拉菜单
    });
  });

  // 组头点击 → 折叠/展开（一级组与表格二级组同逻辑；点组头按钮/链接不折叠）
  el.querySelectorAll(".style-group-h, .style-subgroup-h").forEach(h => {
    h.addEventListener("click", e => {
      if (e.target.closest(".link-btn") || e.target.closest(".group-h-add")) return;
      const group = h.closest(".style-subgroup") || h.closest(".style-group");
      const g = group.dataset.group;
      if (!state.collapsed) state.collapsed = {};
      state.collapsed[g] = !state.collapsed[g];
      group.classList.toggle("collapsed", state.collapsed[g]);
      const tog = h.querySelector(".group-h-tog");
      if (tog) tog.textContent = state.collapsed[g] ? "▸" : "▾";
    });
  });

  // 点击样式项 → 三分支（Word 语义）：
  //   ① 有段落选中 + 点该段当前生效样式 → 展开编辑表单
  //   ② 有段落选中 + 点其他样式 → 把该样式应用到选中段落（刷漆）
  //   ③ 无段落选中 → 展开编辑表单
  //   ＋ 点「标题引导段落」固定行 → 对选中段落开/关引导段（无选中则落到「标题引导段落」编辑表单）
  // 双击=就地改名
  el.querySelectorAll(".style-item[data-elem]").forEach(item => {
    item.addEventListener("click", e => {
      if (e.target.closest(".style-dots")) return;   // ⋯ 点击不展开
      const selPara = document.querySelector(".para.selected");
      const clicked = item.dataset.elem;
      if (selPara) {
        const srcKey = selPara.dataset.key;
        // 用 paraTextOf 取正文原文（不含段首样式标签文本），否则标签名被带入覆盖匹配永远落空
        const effKey = effKeyOf(paraTextOf(selPara)) || srcKey;
        if (clicked === effKey) toggleStyleAccordion(clicked);
        else applyStyleToSelectedPara(clicked);
      } else {
        toggleStyleAccordion(clicked);
      }
    });
    item.addEventListener("dblclick", e => {
      if (e.target.closest(".style-dots")) return;
      beginInlineRenameStyle(item, item.dataset.elem);
    });
  });
  // ⋯ 样式操作子菜单（按桶区分：全局 / 定制）
  el.querySelectorAll(".style-dots[data-dots-style]").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      openStyleMenu(btn.dataset.dotsStyle, btn.dataset.scope, btn);
    });
  });
  // 拖拽：可拖拽行 dragstart（固定样式也可拖，组内排序）
  el.querySelectorAll(".style-item.draggable").forEach(row => {
    row.addEventListener("dragstart", e => {
      e.dataTransfer.setData("text/plain", JSON.stringify({ key: row.dataset.elem, scope: row.dataset.scope }));
      e.dataTransfer.effectAllowed = "move";
      row.classList.add("dragging");
    });
    row.addEventListener("dragend", () => row.classList.remove("dragging"));
    // 行间精确插入排序：dragover 到某行 → 高亮插入位；drop → 在该行前插入（组内重排）
    row.addEventListener("dragover", e => {
      if (!e.dataTransfer.types || !e.dataTransfer.types.includes("text/plain")) return;
      e.preventDefault();
      e.stopPropagation();
      el.querySelectorAll(".drop-before").forEach(r => r.classList.remove("drop-before"));
      row.classList.add("drop-before");
    });
    row.addEventListener("dragleave", e => {
      if (!row.contains(e.relatedTarget)) row.classList.remove("drop-before");
    });
    row.addEventListener("drop", e => {
      e.preventDefault();
      e.stopPropagation();
      row.classList.remove("drop-before");
      let data; try { data = JSON.parse(e.dataTransfer.getData("text/plain") || "{}"); } catch (_) { data = {}; }
      if (!data.key || data.key === row.dataset.elem) return;
      reorderStyle(data.key, row.dataset.elem);
    });
  });
  // 拖拽：分组/二级组作为放置区（跨组拖 = 改样式分类；同组拖 = 移到组末尾）
  el.querySelectorAll(".style-group[data-group], .style-subgroup[data-group]").forEach(group => {
    group.addEventListener("dragover", e => {
      e.preventDefault();
      e.stopPropagation();
      e.dataTransfer.dropEffect = "move";
      el.querySelectorAll(".drag-over").forEach(g => g.classList.remove("drag-over"));
      group.classList.add("drag-over");
    });
    group.addEventListener("dragleave", e => {
      if (!group.contains(e.relatedTarget)) group.classList.remove("drag-over");
    });
    group.addEventListener("drop", e => {
      e.preventDefault();
      e.stopPropagation();       // 二级组优先，避免冒泡到一级组重复处理
      group.classList.remove("drag-over");
      let data; try { data = JSON.parse(e.dataTransfer.getData("text/plain") || "{}"); } catch (_) { data = {}; }
      if (!data.key) return;
      moveStyleToGroup(data.key, group.dataset.group);
    });
  });
  // 编辑表单事件绑定（选中项）；确保 spec 为对象，避免写入报错（方案A：写入 styles 条目 defaults）
  if (state.selectedElem) {
    const selDef = styleDef(state.selectedElem);
    if (selDef && !selDef.defaults) selDef.defaults = {};
    const editEl = el.querySelector(".style-edit");
    if (editEl && selDef) bindElemForm(state.selectedElem, selDef.defaults);
  }
  bindPageSubgroupForms();   // 页面三子组表单（常驻，不依赖选中项）
  applyStyleFilter();   // 保持搜索过滤态（列表重建后重新应用）
}

// 样式搜索过滤：不重建 DOM（避免输入框失焦），按名称显示/隐藏行与空组
function applyStyleFilter() {
  const el = document.getElementById("style-list");
  if (!el) return;
  const q = (styleFilterText || "").trim().toLowerCase();
  el.classList.toggle("filtering", !!q);
  el.querySelectorAll(".style-item").forEach(it => {
    const name = (it.dataset.name || "").toLowerCase();
    it.classList.toggle("filtered-out", !!q && !name.includes(q));
  });
  el.querySelectorAll(".style-subgroup, .style-group").forEach(g => {
    const rows = Array.from(g.querySelectorAll(".style-item"));
    const visible = rows.some(r => !r.classList.contains("filtered-out"));
    g.classList.toggle("filtered-out", !!q && !visible);
    if (q) {
      g.classList.remove("collapsed");           // 搜索期间强制展开命中组
    } else {
      const gid = g.dataset.group;
      g.classList.toggle("collapsed", !!(gid && state.collapsed[gid]));
    }
    const tog = g.querySelector(":scope > .style-group-h > .group-h-tog, :scope > .style-subgroup-h > .group-h-tog");
    if (tog) tog.textContent = g.classList.contains("collapsed") ? "▸" : "▾";
  });
}

// 渲染单个样式行（名称 + 格式摘要 + ⋯；仅自定义样式可拖拽换组）
function styleRowHtml(def, scope, include, tpl) {
  const k = def.key;
  // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults，表单直接读条目 defaults（所见即所得）
  const spec = (def.defaults) || {};
  const isEdit = state.selectedElem === k;                              // 展开编辑表单
  const isParaHi = !!state.selectedParaKey && state.selectedParaKey === k; // 段落选中关联高亮（不展开表单）
  const active = isEdit ? " active" : "";
  const paraHi = isParaHi ? " para-hi" : "";
  const fixed = isFixedStyle(k);
  // 拖拽：固定样式也可拖动（组内排序，2026-08-29 用户定调）；仅内置非样式行（引导段/页码）不可拖。
  // 固定样式跨组移动被 moveStyleToGroup 拦截（分类不可变）。
  const draggable = (scope !== "builtin") ? "true" : "false";
  const dragCls = draggable === "true" ? " draggable" : "";
  const name = def.name || def.key;
  const dots = `<button class="style-dots" data-dots-style="${esc(k)}" data-scope="${esc(scope)}" aria-label="样式操作">⋯</button>`;
  return `<div class="style-item${active}${paraHi}${dragCls}" data-elem="${esc(k)}" data-scope="${esc(scope)}" data-name="${esc(name)}" draggable="${draggable}">
    <span class="style-name">${esc(name)}</span>
    <span class="style-summary">${esc(styleSummaryOf(def, k, tpl))}</span>
    ${dots}
  </div>` + (isEdit ? `<div class="style-edit">${buildElemFormHTML(k, spec)}</div>` : "");
}

// ── 页面样式三子组（纸张 / 页边距 / 页码）──
// 2026-08-30 重构：不再走「页面」样式行 + 「页码」footer 内置行，改三个常驻二级子组，
// 各自展开编辑表单；不参与 include/tpl[key]/styleOrder/新建/改名/删除。
function pageSubgroupsHtml(tpl) {
  const sub = (gid, label, formHtml) => {
    const col = !!state.collapsed[gid];
    return `<div class="style-subgroup${col ? " collapsed" : ""}" data-group="${esc(gid)}">
      <div class="style-subgroup-h">
        <span class="group-h-tog">${col ? "▸" : "▾"}</span>
        <span class="group-h-title">${label}</span>
      </div>
      <div class="style-group-body">${formHtml}</div>
    </div>`;
  };
  return sub("page:纸张", "纸张", pagePaperFormHtml())
    + sub("page:页边距", "页边距", pageMarginFormHtml())
    + sub("page:页码", "页码", pageNumberFormHtml());
}

// 常用纸张规格表（扩展点：新增纸张在此追加一行即可，面板与落盘自动生效）
// w/h：厘米（下拉文案展示）；wTwip/hTwip：DXA（1/1440 英寸，tpl.page.width/height 的真相单位）
const PAPER_SIZES = [
  { id: "a3",     label: "A3",     w: 29.7,  h: 42.0,   wTwip: 16838, hTwip: 23811 },
  { id: "a4",     label: "A4",     w: 21.0,  h: 29.7,   wTwip: 11906, hTwip: 16838 },
  { id: "a5",     label: "A5",     w: 14.8,  h: 21.0,   wTwip:  8391, hTwip: 11906 },
  { id: "b5",     label: "B5",     w: 18.2,  h: 25.7,   wTwip: 10318, hTwip: 14570 },
  { id: "16k",    label: "16开",   w: 18.4,  h: 26.0,   wTwip: 10432, hTwip: 14740 },
  { id: "letter", label: "Letter", w: 21.59, h: 27.94,  wTwip: 12240, hTwip: 15840 },
];
// 反查当前纸张（twip 容差 2 内视为命中）；未命中 = 自定义尺寸
function paperSizeOf(width, height) {
  return PAPER_SIZES.find(s =>
    Math.abs(s.wTwip - (width || 0)) <= 2 && Math.abs(s.hTwip - (height || 0)) <= 2) || null;
}

// 纸张：下拉框选常用规格 → 自动映射 width/height（twip）到 tpl.page
function pagePaperFormHtml() {
  const tpl = currentTemplate() || {};
  const p = tpl.page || (tpl.page = {});
  const w = p.width || 11906, h = p.height || 16838;
  const hit = paperSizeOf(w, h);
  const opts = PAPER_SIZES.map(s =>
    `<option value="${s.id}"${hit && hit.id === s.id ? " selected" : ""}>${s.label}（${s.w} × ${s.h} 厘米）</option>`).join("");
  // 当前尺寸不在预设表 → 追加只读「自定义」占位（提示现状，选中不会覆盖用户数据）
  const custom = hit ? "" :
    `<option value="__custom" selected disabled>自定义（${dxaToCm(w)} × ${dxaToCm(h)} 厘米）</option>`;
  return `<div class="page-form">
    ${buildField("纸张", `<select id="pg-size" class="pg-size">${custom}${opts}</select>`)}
    <div class="pg-size-hint">选择常用纸张尺寸，将自动应用到当前模板</div>
  </div>`;
}

// 页边距：十字四边（上／左-右／下）
function pageMarginFormHtml() {
  const tpl = currentTemplate() || {};
  const p = tpl.page || (tpl.page = {});
  const num = (id, val) =>
    `<input type="number" id="${id}" min="0" max="10" step="0.1" value="${val}"><span class="unit-suffix">cm</span>`;
  return `<div class="page-form">
    <div class="margin-cross">
      <div class="mc-row mc-top">
        <label>上</label>
        ${num("p-top", dxaToCm(p.marginTop || 2098))}
      </div>
      <div class="mc-row mc-mid">
        <div class="mc-side">
          <label>左</label>
          ${num("p-left", dxaToCm(p.marginLeft || 1587))}
        </div>
        <div class="mc-center">页边距</div>
        <div class="mc-side">
          <label>右</label>
          ${num("p-right", dxaToCm(p.marginRight || 1474))}
        </div>
      </div>
      <div class="mc-row mc-bottom">
        <label>下</label>
        ${num("p-bottom", dxaToCm(p.marginBottom || 1984))}
      </div>
    </div>
  </div>`;
}

// 页码：样式全局统一（一份材料一种样式）+ 首页/奇偶开关 + 对齐按钮组（随奇偶开关切换选项集）。
// 数据模型（2026-08-31 用户拍板简化）：pageNumber = { format, align, first }；page.oddEven 决定对齐选项集。
//   首页页码 = 首页显示与否（first 布尔，默认开 = 首页显示页码）；
//   奇偶页不同 = 外侧能力开关（开 = 可用外侧，奇数右偶数左镜像；关 = 所有页同位置）。
//   位置按钮固定 4 个（左/中/右/外侧），不适用项置灰：奇偶关 → 外侧无意义；奇偶开 → 左/右统一无意义。
// 兼容旧三态结构（读取防御）：{ default:{format,align}, first:{...}, even:{...} } 映射到全局值。
const PAGE_NUM_FORMATS = [
  { v: "plain",  label: "1",            demo: "1" },
  { v: "pageOf", label: "第x页共x页",    demo: "第 1 页 共 9 页" },
  { v: "dash",   label: "— 1 —",        demo: "— 1 —" },
  { v: "wave",   label: "～ 1 ～",       demo: "～ 1 ～" },
  // 无 = 全篇无页码（2026-09-01 用户澄清，不是"首页无页码"）；选中后隐藏首页/奇偶/位置控制
  { v: "none",   label: "无",            demo: "无" },
];
// 页码位置：固定 4 个图标，不适用项置灰（由奇偶开关控制生效）
const PN_ALIGNS = [
  { v: "left",    label: "左对齐" },
  { v: "center",  label: "居中" },
  { v: "right",   label: "右对齐" },
  { v: "outside", label: "外侧" },
];
function pageNumberFormHtml() {
  const tpl = currentTemplate() || {};
  const p = tpl.page || (tpl.page = {});
  const pn = p.pageNumber || (p.pageNumber = {});
  // 兼容旧三态：顶层缺 format/align/first 时从 default 态取
  const legacy = pn.default || null;
  const fmt = pn.format !== undefined ? pn.format : (legacy && legacy.format);
  const align = pn.align !== undefined ? pn.align : (legacy && legacy.align);
  // 首页页码开关：默认开 = 首页显示页码（与导出默认一致；存量无字段模板按开显示）
  const first = pn.first !== undefined ? !!pn.first : true;
  const oddEven = !!p.oddEven;
  const fmtOpts = PAGE_NUM_FORMATS.map(f =>
    `<option value="${f.v}"${(fmt || "dash") === f.v ? " selected" : ""}>${esc(f.demo)}</option>`).join("");
  // 不适用项置灰：奇偶关 → 外侧无意义；奇偶开 → 左/右统一无意义
  const disabled = oddEven ? ["left", "right"] : ["outside"];
  // 当前对齐值不在 4 个里（如 justify/旧值）→ 视觉归居中，数据保留（用户点按钮才改写）
  const curAlign = PN_ALIGNS.some(a => a.v === align) ? align : "center";
  // 「无」= 首页不显示页码、后续从 1 开始：隐藏首页/奇偶/位置控制（避免与无语义冲突的调节项）
  const isNone = fmt === "none";
  return `<div class="page-form">
    <div class="pg-row"><span class="pg-row-k">页码样式</span><select id="pg-fmt" title="${isNone ? "无 = 全篇不显示页码" : ""}">${fmtOpts}</select></div>
    ${isNone ? "" : `
    <div class="pg-row"><span class="pg-row-k">首页页码</span>${switchField("pg-first", first, first ? "首页显示页码" : "首页不显示页码")}</div>
    <div class="pg-row"><span class="pg-row-k">奇偶页不同</span>${switchField("pg-oddeven", oddEven, oddEven ? "奇偶页页码位置分别设置" : "奇偶页页码位置一致")}</div>
    <div class="pg-row"><span class="pg-row-k">页码位置</span>${alignButtons("pg-align", curAlign, PN_ALIGNS, disabled)}</div>
    <div class="pg-row"><span class="pg-row-k">距版心下边缘</span><span class="pg-num"><input type="number" id="pg-gap" min="0" max="60" step="0.5" value="${p.numberGap != null ? p.numberGap : DEFAULT_NUMBER_GAP}"><span class="unit-suffix">mm</span></span></div>`}
  </div>`;
}

// 奇偶开关 → 页码对齐联动（2026-09-02 用户拍板 D1=A/D2=A）：
//   开关关（oddEven=false）且页码正处外侧镜像 → 统一右对齐（取奇数页外侧位，视觉连续）；
//   开关开（oddEven=true）且页码是统一左/右 → 切外侧镜像（奇右偶左）。center 不动。
//   返回新 align；null = 无需迁移。保证合法组合 only outside+oddEven=true（与导入归一 convert.js 同口径），
//   否则 toggle 会制造 outside+oddEven=false 残留——预览/导出只认 align==="outside"，关了开关页码仍镜像（假"没联动"）。
function pageNumberAlignForSwitch(curAlign, oddEvenOn) {
  if (oddEvenOn) {
    if (curAlign === "left" || curAlign === "right") return "outside";
  } else if (curAlign === "outside") return "right";
  return null;
}
// 页面子组表单绑定（renderStyleList 末尾调用；表单常驻，不依赖 selectedElem）
function bindPageSubgroupForms() {
  const tpl = currentTemplate();
  if (!tpl) return;
  const p = tpl.page || (tpl.page = {});
  const mark = () => { markDirty(); renderPreview(); };
  // 纸张：下拉选规格 → 映射 width/height（twip 落盘）；自定义项不可选，天然保持用户数据
  const sizeSel = document.getElementById("pg-size");
  if (sizeSel) {
    sizeSel.addEventListener("change", () => {
      const s = PAPER_SIZES.find(x => x.id === sizeSel.value);
      if (!s) return;
      p.width = s.wTwip;
      p.height = s.hTwip;
      mark();
      renderStyleList();
    });
  }
  // 页边距（输入先规整 0.1cm 再落 twip，保证回显 dxaToCm 1 位一致，2026-09-02）
  const bdMargin = (id, field) => {
    const el = document.getElementById(id);
    if (!el) return;
    const cm1 = v => Math.round((parseFloat(v) || 0) * 10) / 10;
    el.addEventListener("input", () => { p[field] = cmToDxa(cm1(el.value)); mark(); });
    el.addEventListener("blur", () => renderStyleList());
  };
  bdMargin("p-top", "marginTop");
  bdMargin("p-left", "marginLeft");
  bdMargin("p-right", "marginRight");
  bdMargin("p-bottom", "marginBottom");
  // 页码：样式全局下拉 + 首页开关（页码有无）+ 奇偶开关（对齐分档）+ 对齐按钮组
  // 2026-08-31 全局化（用户拍板）：pageNumber = { format, align, first }；对齐选项集随奇偶开关切换
  const pn = p.pageNumber || (p.pageNumber = {});
  const fmtSel = document.getElementById("pg-fmt");
  if (fmtSel) fmtSel.addEventListener("change", () => {
    pn.format = fmtSel.value;
    mark(); renderStyleList();
  });
  // 页码垂直锚点（GB/T 9704 7.5）：page.numberGap（mm），页脚位置的唯一口径。
  // 空/非法 → 回落国标缺省 7mm（不再"清除字段回页脚距边老口径"——那条路径已废弃）。
  const gapEl = document.getElementById("pg-gap");
  if (gapEl) {
    gapEl.addEventListener("input", () => {
      const raw = String(gapEl.value).trim();
      if (raw === "") return;                    // 中间态（清空重输）：等 blur 再规范，避免输入过程中跳值
      const v = parseFloat(raw);
      if (!isFinite(v)) return;
      p.numberGap = Math.max(0, Math.round(v * 10) / 10);
      mark();
    });
    gapEl.addEventListener("blur", () => {
      const v = parseFloat(String(gapEl.value).trim());
      p.numberGap = isFinite(v) ? Math.max(0, Math.round(v * 10) / 10) : DEFAULT_NUMBER_GAP;
      renderStyleList();                         // 重渲染把规范后的值写回输入框
    });
  }
  const alGroup = document.getElementById("pg-align");
  if (alGroup) alGroup.querySelectorAll(".align-btn").forEach(b => {
    b.addEventListener("click", () => {
      pn.align = b.dataset.align;
      alGroup.querySelectorAll(".align-btn").forEach(x => x.classList.toggle("active", x === b));
      mark(); renderStyleList();
    });
  });
  // 首页开关 → pageNumber.first（首页页码有无）；奇偶开关 → page.oddEven（对齐选项集联动重渲染）
  const bindChk = (id, set) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("change", () => { set(el.checked); mark(); renderStyleList(); });
  };
  bindChk("pg-first", v => { pn.first = v; });
  // 奇偶开关联动页码对齐（2026-09-02 拍板）：toggle 时同步迁移 align——
  //   否则开关只改 oddEven、预览/导出的镜像判定只认 align==="outside"，开关对页码零联动
  bindChk("pg-oddeven", v => {
    p.oddEven = v;
    const legacy = pn.default || null;
    const curAlign = pn.align !== undefined ? pn.align : (legacy && legacy.align);
    const next = pageNumberAlignForSwitch(curAlign, v);
    if (next) pn.align = next;
  });
}

// ── 新建表格分组（表格一级组头＋）：无弹窗，直接在表格组最上方插入空分组并进入改名 ──
// 空分组持久化于模板集 tableGroups 字段（组名不再依赖样式 group 推导）
function createTableGroup() {
  if (!state.templateSet) { toast("请先加载模板集", "err"); return; }
  state.pendingTableGroup = true;
  renderStyleList();
  const inp = document.querySelector("#style-list input[data-new-table-group]");
  if (inp) { inp.focus(); inp.select(); }
}

// ── 行内新建样式（无弹窗）：插入到该分类末尾 → 立即进入重命名态 ──
// category：text（段落）/ char（字符）/ table（表格）
// group（仅表格）：表格二级分组名（如「数据表格」「版记」），由二级组头＋触发时传入
async function createStyleInline(cat, group) {
  if (!state.templateSet) { toast("请先加载模板集", "err"); return; }
  const isTable = cat === "table";
  const isChar = cat === "char";
  const base = isTable ? "表格样式" : isChar ? "字符样式" : "新样式";
  const used = new Set(allStyleDefs().map(s => s.key));
  let key = base, i = 1;
  while (used.has(key)) { i++; key = `${base} ${i}`; }
  // 新建样式的默认字体：系统替用户填值 → 按「字体映射」开关解析（开=模板映射的方正字体，关=系统自带字体）。
  // 用户随后手选的字一律原样保存，不再被改写（映射只发生在这两处生成默认值的时刻）。
  const mapFont = f => (f && fontFamilies.length ? resolveFontForWrite(f) : f);
  const def = {
    key, name: key,
    category: isTable ? "table" : (isChar ? "char" : "text"),
    matchType: "none", match: "", mode: "whole",
    // v1.31：新建文本样式默认后续段落样式 = 正文段落（table/char/page 忽略 next）
    ...(isTable || isChar ? {} : { next: "正文段落" }),
    defaults: isTable
      ? { font: mapFont("仿宋"), size: 14, align: "justify", line: 28, bold: false, indentChars: 0, spaceBefore: 0, spaceAfter: 0 }
      : isChar
        ? { font: "", size: 16, bold: false }
        : { font: mapFont("仿宋"), size: 16, align: "justify", line: 28, bold: false, indentChars: 2, spaceBefore: 0, spaceAfter: 0 },
  };
  if (isTable) {
    def.group = group || TABLE_GROUP_DEFAULT;
    def.role = "内容";
    state.collapsed["table:" + def.group] = false;        // 新行分组保持展开可见
  }
  // v1.30 解耦：样式定义随当前模板自包含（styles[]），不落集级
  const tpl = currentTemplate();
  if (!tpl) { toast("请先创建模板，样式随模板保存", "err"); return; }
  if (!Array.isArray(tpl.styles)) tpl.styles = [];
  tpl.styles.push(def);
  ensureStyleOrderComplete(tpl);   // 2026-09-01：新建样式同步进 styleOrder，防残缺甩尾
  if (!Array.isArray(tpl.include)) tpl.include = [];
  if (!tpl.include.includes(key)) tpl.include.push(key);
  // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults（新建样式自带），不再写顶层扁平字段
  state.collapsed[cat] = false;                                  // 确保新行可见
  markDirty();
  renderStyleList();
  // 新行立即进入重命名（草稿态：取消/空名则丢弃该样式）
  const row = document.querySelector(`#style-list .style-item[data-elem="${CSS.escape(key)}"]`);
  if (row) beginInlineRenameStyle(row, key, { draft: true });
  else await flushSave();
}

// ── 组头 ＋ 下拉菜单（段落/字符一级组）：新建空样式 / 从默认样式中选（按 category 筛未 include 的种子）──
async function openAddStyleMenu(anchor, cat) {
  if (!state.templateSet) return;
  if (!SEED_STYLE_CACHE) await ensureSeedStyles();   // 种子未加载先取（菜单依赖种子清单）
  const tpl = currentTemplate();
  const items = [];
  items.push({
    icon: ICONS.newTpl,
    label: "新建空样式",
    onClick: () => createStyleInline(cat),
  });
  // 「从默认样式中选」：种子清单里同 category、且当前模板未 include 的项（已存在的自然隐去 = 防重名）
  const seeds = (SEED_STYLE_CACHE || []).filter(s => styleCategory(s) === cat);
  if (seeds.length) {
    const have = new Set(allStyleDefs().map(d => d.key));
    const subs = seeds
      .filter(s => !have.has(s.key))
      .map(s => ({
        label: s.name || s.key,
        onClick: () => cloneSeedToTemplate(s, cat),
      }));
    if (subs.length) {
      items.push({ icon: ICONS.set, label: "从默认样式中选", sub: subs });
    }
  }
  openSubmenu(anchor, items);
}

// 从默认样式（种子）克隆进当前模板：defaults + matchType/match 原样携带，
// 自动 include + 初始化扁平值 → 进入改名态（草稿：空名/Esc 丢弃）。
// 字符样式不携带 next/matchType/match/mode（字符样式无段落级属性）。
function cloneSeedToTemplate(seed, cat) {
  if (!seed) return;
  const tpl = currentTemplate();
  if (!tpl) { toast("请先创建模板，样式随模板保存", "err"); return; }
  if (!Array.isArray(tpl.styles)) tpl.styles = [];
  const copy = seedToTemplateDef(seed);   // 字体按当前开关落值（种子存的是字族名）
  copy.key = seed.key;
  copy.name = seed.name || seed.key;
  if (cat === "char") {
    delete copy.matchType; delete copy.match; delete copy.mode; delete copy.next;
    // 字号空 = 跟随段落（保留空，不强制 16）；font/bold 同上
    copy.defaults = { font: copy.defaults.font || "", size: copy.defaults.size || "", bold: !!copy.defaults.bold };
  }
  tpl.styles.push(copy);
  ensureStyleOrderComplete(tpl);   // 2026-09-01：克隆种子样式同步进 styleOrder，防残缺甩尾
  if (!Array.isArray(tpl.include)) tpl.include = [];
  if (!tpl.include.includes(copy.key)) tpl.include.push(copy.key);
  // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults（克隆自带），不再初始化扁平值
  state.collapsed[cat] = false;
  markDirty();
  renderStyleList();
  const row = document.querySelector(`#style-list .style-item[data-elem="${CSS.escape(copy.key)}"]`);
  if (row) beginInlineRenameStyle(row, copy.key, { draft: true });
  else flushSave();
}

// 丢弃草稿样式（新建后取消重命名）：清定义 + 模板引用（v1.30：样式随当前模板）
async function discardDraftStyle(key) {
  if (!state.templateSet) return;
  const tpl = currentTemplate();
  if (tpl && Array.isArray(tpl.styles)) tpl.styles = tpl.styles.filter(s => s.key !== key);
  if (tpl) {
    const inc = tpl.include || [];
    const i = inc.indexOf(key);
    if (i >= 0) inc.splice(i, 1);
    if (tpl[key] !== undefined) delete tpl[key];
  }
  if (state.selectedElem === key) state.selectedElem = null;
  markDirty();
  await flushSave();
  renderAll();
}

// 样式重命名：key 与显示名强制同步（v5.4 一条链）；返回 false = 重名被拒（调用方处理提示）
// 引用迁移：include / 其他样式的 next / selectedElem（方案A：值在 styles 条目 defaults，随条目走，无需扁平迁移）
function renameStyleKey(def, newKey) {
  const oldKey = def.key;
  if (newKey === oldKey) return true;
  // 防重名：key 或 name 撞名都拒绝（不允许静默仅改显示名造成漂移）
  if (allStyleDefs().some(s => s !== def && (s.key === newKey || s.name === newKey))) return false;
  def.key = newKey;
  def.name = newKey;
  // v1.30：样式随当前模板，只同步当前模板的 include
  const tpl = currentTemplate();
  if (tpl) {
    const inc = tpl.include || [];
    const i = inc.indexOf(oldKey);
    if (i >= 0) inc[i] = newKey;
    // v1.31：同步更新指向旧 key 的 next 引用，防悬挂
    if (Array.isArray(tpl.styles)) for (const s of tpl.styles) if (s.next === oldKey) s.next = newKey;
  }
  if (state.selectedElem === oldKey) state.selectedElem = newKey;
  return true;
}

// 按范围打开不同内容的 ⋯ 子菜单（两类：全局 / 定制）
function openStyleMenu(key, scope, btn) {
  const renameItem = { icon: ICONS.rename, label: "重命名", onClick: () => { const row = btn.closest(".style-item"); if (row) beginInlineRenameStyle(row, key); } };
  const items = [];
  items.push(renameItem);
  items.push({ icon: ICONS.copy, label: "复制样式", onClick: () => copyStyleFrom(key, "set") });
  if (isFixedStyle(key)) {
    // 内置固定样式：可「重置」回种子默认格式（定义保留在种子，随时可从「从默认样式中选」加回）
    items.push({ icon: ICONS.reset, label: "重置样式", onClick: () => resetStyle(key) });
  }
  // 删除样式对内置/自定义一视同仁：从当前模板移除（内置样式移除后转换按正文类样式降级渲染，内容不丢）
  items.push({ icon: ICONS.del, label: "删除样式", danger: true, onClick: () => deleteStyle(key, "set") });
  openSubmenu(btn, items);
}

// ── styleOrder 完整性维护（2026-09-01）：残缺 styleOrder（缺 key）是面板乱序根因，
// 新建/克隆/补建样式与拖拽前都调用本函数，把缺失 key 按种子权威序插回原锚点（前驱之后）。
// 无 styleOrder（从未拖过）时不创建、不干预 —— 面板完全走 seedOrder 兜底。
function ensureStyleOrderComplete(tpl) {
  if (!tpl || !Array.isArray(tpl.styleOrder)) return;
  const order = tpl.styleOrder;
  const have = new Set(order);
  const missing = allStyleDefs().filter(d => !have.has(d.key));
  if (!missing.length) return;
  // 种子基准与 renderStyleList 一致：SEED_STYLE_CACHE（懒加载缓存）优先，seedStyles 兜底
  const seedOrder = new Map((SEED_STYLE_CACHE || seedStyles || []).map((s, i) => [s.key, i]));
  const sorted = missing.slice().sort((a, b) => (seedOrder.get(a.key) ?? 9999) - (seedOrder.get(b.key) ?? 9999));
  for (const d of sorted) {
    const si = seedOrder.get(d.key);
    let pos = order.length;
    if (si !== undefined) {
      let anchor = -1;
      for (let i = 0; i < order.length; i++) {
        const sk = seedOrder.get(order[i]);
        if (sk !== undefined && sk < si && i > anchor) anchor = i;
      }
      pos = anchor + 1;
    }
    order.splice(pos, 0, d.key);
  }
}

// ── 拖拽落组：target 为 "text" / "table" / "table:<二级组名>"（page 不接受） ──
// 跨组 = 改样式分类（category/group），同组 = 移到组末尾；仅自定义样式可移动。
// 拖拽排序（2026-08-29 用户定调）：把样式移到目标行前，持久化到模板 styleOrder（随模板落盘）。
// styleOrder 缺失时以当前全部样式序为基准（= 权威清单序兜底）；仅影响面板展示序，不影响识别（matchers 用 styles[] 序）。
function reorderStyle(key, beforeKey) {
  const tpl = currentTemplate();
  if (!tpl || key === beforeKey) return;
  ensureStyleOrderComplete(tpl);   // 拖拽前先补全残缺 styleOrder，拖一次即自愈
  let order = Array.isArray(tpl.styleOrder) ? tpl.styleOrder.slice() : allStyleDefs().map(s => s.key);
  const i = order.indexOf(key);
  if (i >= 0) order.splice(i, 1);
  const bi = beforeKey ? order.indexOf(beforeKey) : -1;
  if (bi >= 0) order.splice(bi, 0, key); else order.push(key);
  tpl.styleOrder = order;
  markDirty();
  renderStyleList();
  flushSave();   // 落盘（失败有 toast）
}

async function moveStyleToGroup(key, target) {
  const tpl = currentTemplate();
  if (!tpl || !target) return;
  if (isFixedStyle(key)) { toast("固定样式不可移动分类（可拖动排序）", "err"); return; }
  const arr = Array.isArray(tpl.styles) ? tpl.styles : [];
  const i = arr.findIndex(x => x.key === key);
  if (i < 0) { toast("内置固定样式不可移动分类", "err"); return; }
  let cat = target, group = "";
  if (target.indexOf("table:") === 0) { cat = "table"; group = target.slice("table:".length); }
  if (cat === "page") { toast("页面样式组不接受其他样式", "err"); return; }
  const [def] = arr.splice(i, 1);
  def.category = cat;
  if (cat === "table") { def.group = group || TABLE_GROUP_DEFAULT; if (!def.role) def.role = "内容"; }
  else { delete def.group; delete def.role; }
  arr.push(def);
  markDirty();
  try { await apiSaveTemplateset(); } catch (e) { toast("保存失败：" + e.message, "err"); }
  renderAll();
}

// ── 就地重命名：样式（双击整行 → 行内输入框，改的是样式显示名） ──
// opts.draft=true：行内新建的草稿样式 —— 取消或空名则丢弃该样式，提交时 key 与名称同步
function beginInlineRenameStyle(el, key, opts = {}) {
  const draft = !!opts.draft;
  const def = styleDef(key);
  if (!def) return; // footer 等无定义的不重命名
  const nameSpan = el.querySelector(".style-name");
  const oldName = nameSpan ? nameSpan.textContent : (def.name || def.key || "");
  styleRenaming = true; // 必须在任何渲染前置位，防止并发 renderStyleList 拆掉输入框
  el.classList.add("editing");
  const input = document.createElement("input");
  input.type = "text";
  input.className = "style-rename-input";
  input.value = oldName;
  input.placeholder = draft ? "输入样式名称" : "";
  nameSpan.replaceWith(input);
  input.focus();
  input.select();
  let done = false;
  const finish = async (save) => {
    if (done) return;
    done = true;
    const v = input.value.trim();
    styleRenaming = false;
    if (draft && (!save || !v)) {          // 草稿取消/空名 → 丢弃新建的样式
      el.classList.remove("editing");
      await discardDraftStyle(key);
      return;
    }
    if (!save || !v || v === oldName) {
      const span = document.createElement("span");
      span.className = "style-name";
      span.textContent = oldName;
      input.replaceWith(span);
      el.classList.remove("editing");
      if (draft) { await flushSave(); renderAll(); }
      return;
    }
    // 防重名 + 强制 key===name（v5.4 一条链）：key 或 name 撞名一律拒绝
    if (!renameStyleKey(def, v)) {
      toast("已存在同名样式", "err");
      const span = document.createElement("span");
      span.className = "style-name";
      span.textContent = oldName;
      input.replaceWith(span);
      el.classList.remove("editing");
      renderStyleList();       // 重建列表（含输入框还原）
      return;
    }
    markDirty();                              // 显式置脏，保证 flushSave 真正落盘
    el.classList.remove("editing");
    try {
      await apiSaveTemplateset();
    } catch (e) { toast("保存失败：" + e.message, "err"); }
    renderAll();
      };
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); finish(true); }
    else if (e.key === "Escape") { e.preventDefault(); finish(false); }
  });
  input.addEventListener("blur", () => finish(true));
}

// 当前模板集落盘别名（样式编辑/删除等沿用）；改为立即落盘（flushSave 含自动派生文件名）
async function apiSaveTemplateset() {
  return flushSave();
}

// ── 复制样式（到当前模板，v1.30：样式随模板） ──
async function copyStyleFrom(key, scope) {
  const src = styleDef(key);
  if (!src) return;
  const tpl = currentTemplate();
  if (!tpl) { toast("请先创建模板，样式随模板保存", "err"); return; }
  if (!Array.isArray(tpl.styles)) tpl.styles = [];
  const targetArr = tpl.styles;
  let base = key + "_copy", nk = base, i = 1;
  const ks = allStyleDefs().map(s => s.key);
  while (ks.includes(nk)) { i++; nk = base + i; }
  const copy = JSON.parse(JSON.stringify(src));
  copy.key = nk; copy.name = (src.name || src.key) + " 副本";
  targetArr.push(copy);
  if (!tpl.include.includes(nk)) tpl.include.push(nk);
  // 2026-09-02 方案A：值随条目（defaults 已深拷贝），不再初始化扁平值
  markDirty();
  await flushSave();
  renderAll();
  }

// ── 删除样式（清理定义 + 模板引用，仅本模板） ──
// 内置样式同样可删：v1.30 起模板自包含 styles[] + include 是启用开关，移除后转换走 safeKey 降级链
// 落到正文类样式（内容不丢），需要时用「从默认样式中选」加回，故不再区分「内置不可删」。
async function deleteStyle(key, scope) {
  const def = styleDef(key);
  if (!def) return;
  const tpl0 = currentTemplate();
  if (!tpl0) return;
  // 兜底闸：正文/正文段落至少要留一个 —— 二者皆无时 safeKey 降级链无处可落，全篇段落失格式
  if (key === "正文" || key === "正文段落") {
    const rest = new Set((tpl0.styles || []).filter(s => s.key !== key).map(s => s.key));
    if (!rest.has("正文") && !rest.has("正文段落")) {
      toast("至少保留「正文」或「正文段落」其一（正文格式兜底）", "err");
      return;
    }
  }
  // 样式只服务当前模板（v1.30 自包含），删除即从当前模板移除定义 + 引用
  const msg = isFixedStyle(key)
    ? `确定从当前模板移除内置样式「${esc(def.name || def.key)}」？<br>移除后本模板不再使用它（相关段落按正文格式渲染，内容不丢），可从「从默认样式中选」重新添加。`
    : `确定删除样式「${esc(def.name || def.key)}」？<br>删除后不可恢复。`;
  const ok = await modalConfirm("删除样式", msg, "删除");
  if (!ok) return;
  const tpl = currentTemplate();
  if (tpl && Array.isArray(tpl.styles)) tpl.styles = tpl.styles.filter(s => s.key !== key);
  if (tpl) {
    const inc = tpl.include || [];
    const i = inc.indexOf(key);
    if (i >= 0) inc.splice(i, 1);
    if (tpl[key] !== undefined) delete tpl[key];
    if (Array.isArray(tpl.styleOrder)) {          // 2026-09-01：删除样式同步清 styleOrder 残留
      const si = tpl.styleOrder.indexOf(key);
      if (si >= 0) tpl.styleOrder.splice(si, 1);
    }
    markDirty();
  }
  // 清理预览覆盖里指向被删样式的条目（toKey 悬空，避免渲染落到不存在的样式）
  if (state.paraOverrides.some(o => o.toKey === key)) {
    state.paraOverrides = state.paraOverrides.filter(o => o.toKey !== key);
    saveParaOverrides();
  }
  if (state.selectedElem === key) state.selectedElem = null;
  await flushSave();
  renderAll();
  }

// ── 重置内置样式（仅清本模板的自定义值，定义与 include 保留，转换不受影响） ──
// 固定样式定义来自 config/style-seeds.json 种子；「重置」= 回到种子默认格式（不动定义、不移除样式）。
// 想让本模板彻底不用某个内置样式 → 用「删除样式」（移除 styles[] + include，转换走 safeKey 降级到正文类），
// 需要时用组头「+ 从默认样式中选」把该种子样式加回来。
async function resetStyle(key) {
  if (!state.templateSet) return;
  const def = styleDef(key);
  if (!def) return;
  if (!isFixedStyle(key)) { toast("自定义样式请用「删除样式」", "err"); return; }
  const ok = await modalConfirm("重置样式",
    `确定将内置样式「${def.name || def.key}」恢复为种子默认格式？<br>本模板集所有模板对该样式的自定义值将被清除；样式定义保留、转换不受影响。`,
    "重置");
  if (!ok) return;
  // 方案A：重置 = 各模板 styles 条目 defaults 回种子默认（样式值唯一真相在 defaults）
  const seed = (SEED_STYLE_CACHE || []).find(s => s.key === key);
  const seedDefaults = materializeSeedDefaults(seed ? seed.defaults : {});   // 字体按当前开关落值
  for (const t of state.templateSet.templates) {
    const s = Array.isArray(t.styles) ? t.styles.find(x => x.key === key) : null;
    if (s) s.defaults = JSON.parse(JSON.stringify(seedDefaults));
  }
  // 预览覆盖里指向被重置样式的条目一并清掉（该样式已回种子默认，覆盖无意义）
  if (state.paraOverrides.some(o => o.toKey === key)) {
    state.paraOverrides = state.paraOverrides.filter(o => o.toKey !== key);
    saveParaOverrides();
  }
  if (state.selectedElem === key) state.selectedElem = null;
  markDirty();
  await flushSave();
  renderAll();
}

// 全局样式库已移除：无「恢复默认」概念，样式全部随模板集走（见 docs/样式定义体系-总体设计.md）


// ── 样式行点击：局部展开/折叠编辑表单（不整体重建列表，保证双击可触发） ──
function toggleStyleAccordion(key) {
  // 样式行点击 = 展开/折叠编辑表单；不改动预览段落选中状态，也不高亮相关段落
  if (state.selectedElem === key) state.selectedElem = null;
  else state.selectedElem = key;
  document.querySelectorAll("#style-list .style-item").forEach(it => {
    it.classList.toggle("active", it.dataset.elem === state.selectedElem);
  });
  document.querySelectorAll("#style-list .style-edit").forEach(e => e.remove());
  if (state.selectedElem) {
    const row = document.querySelector(`#style-list .style-item[data-elem="${CSS.escape(state.selectedElem)}"]`);
    const tpl = currentTemplate();
    if (row && tpl) {
      const selDef = styleDef(state.selectedElem);
      if (selDef && !selDef.defaults) selDef.defaults = {};
      const spec = (selDef && selDef.defaults) || {};
      const edit = document.createElement("div");
      edit.className = "style-edit";
      edit.innerHTML = buildElemFormHTML(state.selectedElem, spec);
      row.after(edit);
      if (selDef) bindElemForm(state.selectedElem, selDef.defaults);
      checkFontInstalled(spec.font);
    }
  }
}

// ═══════════════════════════════════════════
// 段落级样式覆盖（Word 语义：选中预览文字区域 → 点样式行 → 应用到该段）
// 数据模型：[{ match: 段落原文, toKey: 换成什么样式 }]，按段落原文匹配（精确到段）。
// 定位：纯**预览演示态**（用户拍板 A）——只存浏览器 localStorage，不落模板文件，
//       不影响导出 Word（导出按模板规则走）。刷新页面保留、重启应用丢失。
// ═══════════════════════════════════════════
const PARA_OV_KEY = "klzl_para_overrides";

// ── 应用即落盘辅助（2026-08-30）──
// 前端样本重建：与服务端 sampleBlocksToTxt 互逆一致（含表格/页脚），应用样式落盘用。
function sampleTxtFromBlocks(blocks) {
  const lines = [];
  for (const b of blocks) {
    if (b.type === "table") {
      lines.push(`[表格·${b.title || "示例数据表"}]`);
      if (Array.isArray(b.header) && b.header.length) lines.push("| " + b.header.join(" | ") + " |");
      for (const row of (Array.isArray(b.body) ? b.body : [])) lines.push("| " + row.join(" | ") + " |");
    } else if (b.type === "footer") {
      lines.push("[页脚]");
      if (b.text) lines.push(b.text);
    } else if (b.type === "hline") {
      lines.push("[红线]");
    } else if (b.type === "pagebreak") {
      lines.push("[分页]");
    } else {
      lines.push(b.blank || !b.text ? "[空行]" : `[${b.style}]`);
      if (b.text) lines.push(b.text);
    }
    lines.push("");
  }
  return lines.join("\n").trimEnd() + "\n";
}

// 前端样本解析（与服务端 parseSampleTxt 同源互逆；保存前改样式标记用，不改内容结构）
function parseSampleTxtLocal(text) {
  const lines = String(text).split(/\r?\n/);
  const blocks = [];
  let cur = null;
  const pushCur = () => { if (cur) blocks.push(cur); cur = null; };
  for (const raw of lines) {
    const line = raw.replace(/\s+$/, "");
    const m = line.match(/^\[(.*)\]$/);
    if (m) {
      pushCur();
      const tag = m[1].trim();
      if (tag === "页脚") cur = { type: "footer", text: "" };
      else if (tag === "空行") cur = { type: "para", style: "正文段落", text: "", blank: true };
      else if (tag === "红线") cur = { type: "hline", text: "" };
      else if (tag === "分页") cur = { type: "pagebreak", text: "" };
      else if (tag.startsWith("表格")) {
        const parts = tag.split("·");
        const title = parts.length > 1 ? parts.slice(1).join("·").trim() : "";
        cur = { type: "table", title, rows: [] };
      } else cur = { type: "para", style: tag, text: "" };
      continue;
    }
    if (!cur) continue;
    if (cur.type === "table") {
      if (line.startsWith("|")) cur.rows.push(line.split("|").slice(1, -1).map(c => c.trim()));
      continue;
    }
    if (!line) continue;   // 空行仅作块分隔（与服务端同口径，保证往返互逆）
    cur.text = cur.text ? cur.text + "\n" + line : line;
  }
  pushCur();
  for (const b of blocks) {
    if (b.type === "table") {
      const all = (b.rows || []).filter(r => r.length > 0);
      b.header = all[0] || [];
      b.body = all.slice(1).filter(r => !(r.length && r.every(c => /^[-:\s]*$/.test(c))));
      delete b.rows;
    }
  }
  return { ok: true, blocks };
}

// 导入保存：样本里「未保留样式」的段落 → 降级为正文类样式标记，内容不丢（2026-09-05 用户口径：
// 确认环节取消勾选 = 这个样式不入模板，但段落内容仍是文档内容，样本里按正文兜底呈现）。
// 样本由服务端按 LLM 分类生成（含版头结构/脱敏/表格），前端只改样式标记、不重建内容。
function importBodyFallbackKey(keep, styles) {
  const arr = Array.isArray(styles) ? styles : [];
  const isText = s => s && (s.category === "text" || s.category === undefined);
  const specialKey = m => {
    const d = arr.find(s => isText(s) && s.matchType === "special" && s.match === m);
    return d && keep.has(d.key) ? d.key : null;
  };
  const named = k => (k && keep.has(k) && arr.some(s => isText(s) && s.key === k)) ? k : null;
  const fb = specialKey("default") || named("正文段落") || specialKey("normal") || named("正文") || named("body");
  if (fb) return fb;
  const any = arr.find(s => isText(s) && keep.has(s.key));
  return any ? any.key : null;
}
function normalizeImportSample(sampleTxt, keep, styles) {
  if (typeof sampleTxt !== "string" || !sampleTxt.trim()) return sampleTxt;
  const fb = importBodyFallbackKey(keep, styles);
  if (!fb) return sampleTxt;
  const { blocks } = parseSampleTxtLocal(sampleTxt);
  let changed = false;
  for (const b of blocks) {
    if (b.type !== "para" || b.blank || !b.style) continue;
    if (keep.has(b.style)) continue;
    b.style = fb;
    changed = true;
  }
  return changed ? sampleTxtFromBlocks(blocks) : sampleTxt;
}

// 清除某段残留的 localStorage 覆盖记录（应用即落盘后覆盖不再写入，仅清历史数据过渡）
function clearParaOverride(paraText) {
  const existing = state.paraOverrides.find(o => o.match === paraText);
  if (existing) {
    state.paraOverrides.splice(state.paraOverrides.indexOf(existing), 1);
    if (!state.paraOverrides.length) state.paraOverrides = [];
    saveParaOverrides();
  }
}

// 应用即落盘核心：改 sampleBlocks[idx] 的样式标记 → 重建 tpl.sample → 随模板集保存 → 重拉样本。
async function commitParaStyleToSample(idx, toKey, paraText) {
  const tpl = currentTemplate();
  if (!tpl) return;
  const blk = (state.sampleBlocks || [])[idx];
  if (!blk || blk.type !== "para") return;
  blk.style = toKey;
  clearParaOverride(paraText);
  tpl.sample = sampleTxtFromBlocks(state.sampleBlocks);
  await persistCurrentSet();
  _sampleKey = "";               // 强制重拉（服务端重新解析，含新样式标记）
  await loadSample();
}

function ovStorageKey() {
  return (state.currentFile || "(未保存)") + "|" + (state.currentTemplateKey || "");
}

// 切换模板集/模板时从 localStorage 读回当前模板的覆盖
function loadParaOverrides() {
  try {
    const map = JSON.parse(localStorage.getItem(PARA_OV_KEY) || "{}");
    const list = map[ovStorageKey()];
    state.paraOverrides = Array.isArray(list) ? list : [];
  } catch (e) { state.paraOverrides = []; }
}

// 应用/撤销覆盖后写回 localStorage（替换整个 map 里当前模板的条目）
function saveParaOverrides() {
  try {
    const map = JSON.parse(localStorage.getItem(PARA_OV_KEY) || "{}");
    map[ovStorageKey()] = state.paraOverrides;
    localStorage.setItem(PARA_OV_KEY, JSON.stringify(map));
  } catch (e) { /* 容量/隐私模式忽略 */ }
}

// 查询某段落的实际生效样式：有覆盖 → 目标样式；无覆盖 → null（用段落本身标定样式）
function effKeyOf(paraText) {
  if (!paraText) return null;
  const o = state.paraOverrides.find(x => x.match === paraText);
  return (o && o.toKey) ? o.toKey : null;
}

// 段落的正文原文（不含段首样式标签文本）：
// .style-tag 是 .para 的子节点，直接读 el.textContent 会把标签名（如「正文段落」）
// 一并带上，导致与覆盖记录（按段落原文精确匹配）对不上。
function paraTextOf(paraEl) {
  if (!paraEl) return "";
  const body = paraEl.querySelector(".para-body");
  return body ? body.textContent : paraEl.textContent;
}

function clearParaSelection() {
  state.paraSelected = false;
  state.selectedParaKey = null;
  state.selectedParaIdx = null;       // 同步清除精确段索引
  // 取消段落选中 → 由选中联动展开的编辑表单一并收起，避免留下"无段落但表单开着"的悬空态
  state.selectedElem = null;
  document.querySelectorAll(".para").forEach(el => {
    el.classList.remove("selected");
    el.removeAttribute("data-para-selected");
  });
}

// 把目标样式应用到当前选中段落（2026-08-30 应用即落盘：改样本 [样式名] 标记并随模板集保存；
// 不再写 localStorage 覆盖。编辑样本可见、刷新保留、导出生效。旧覆盖记录在应用时一并清除过渡）
async function applyStyleToSelectedPara(toKey) {
  const paraEl = document.querySelector(".para.selected");
  if (!paraEl) return;
  const srcKey = paraEl.dataset.key;
  const idx = Number(paraEl.dataset.paraIdx);
  const paraText = paraTextOf(paraEl);
  if (!srcKey || !Number.isInteger(idx) || !paraText) return;
  const effKey = effKeyOf(paraText) || srcKey;
  if (toKey === effKey) {
    showUrlHint(`已是「${styleName(toKey)}」`, "");   // no-op：该段已经是这个样式
    return;
  }
  if (toKey === srcKey) {
    // 点回段落本身标定的样式：样本已是 srcKey，无标记变更，仅清残留覆盖
    clearParaOverride(paraText);
    state.selectedParaKey = srcKey;
    const linkKey = syncPanelToPara(srcKey);        // 联动跟随：高亮+展开恢复后的样式
    saveParaOverrides(); renderPreview(); renderStyleList();
    scrollStyleRowIntoView(linkKey);
    showUrlHint(`已恢复为「${styleName(srcKey)}」`, "ok");
    return;
  }
  await commitParaStyleToSample(idx, toKey, paraText);
  state.selectedParaKey = toKey;   // 应用后段落保持选中，样式行高亮切到新样式
  const linkKey = syncPanelToPara(toKey);          // 联动跟随：高亮+展开新生效样式
  renderStyleList();
  scrollStyleRowIntoView(linkKey);
  showUrlHint(`已应用「${styleName(toKey)}」并保存`, "ok");
}

// 预览兜底样式：样本段落标了一个不在 include 里的样式（样式被删 / 导入时未勾选）→
// 降级到正文类样式渲染（内容不丢）。返回 null = 模板里连正文类样式都没有（极端，维持隐藏）。
function previewBodyFallbackKey(include) {
  const arr = allStyleDefs();
  const isText = s => s && (s.category === "text" || s.category === undefined);
  const special = m => {
    const d = arr.find(s => isText(s) && s.matchType === "special" && s.match === m);
    return d && include.includes(d.key) ? d.key : null;
  };
  const named = k => (k && include.includes(k) && arr.some(s => isText(s) && s.key === k)) ? k : null;
  const fb = special("default") || named("正文段落") || special("normal") || named("正文") || named("body");
  if (fb) return fb;
  const any = arr.find(s => isText(s) && include.includes(s.key));
  return any ? any.key : null;
}

// ── 手动把某段设为/取消标题引导段落（需求②）──
// 2026-08-30 应用即落盘：设引导段 = 样本标记改为 [标题引导段落] 并随模板集保存；
// 预览按简单规则切标题句/正文：段首第一个结束标志（默认 。：），最长 50 汉字，后续正文。
async function setParaLead(paraEl, on) {
  if (!paraEl) return;
  const srcKey = paraEl.dataset.key;
  const idx = Number(paraEl.dataset.paraIdx);
  const paraText = paraTextOf(paraEl);
  if (!srcKey || !Number.isInteger(idx) || !paraText) return;
  if (on) await ensureLeadStyleDef();   // 模板缺引导段样式 → 先补建，否则标记无处落地
  await commitParaStyleToSample(idx, on ? LEAD_STYLE_KEY : srcKey, paraText);
  state.selectedParaKey = on ? LEAD_STYLE_KEY : srcKey;
  const linkKey = syncPanelToPara(state.selectedParaKey);   // 引导段 → 联动「标题引导段落」样式行
  renderStyleList();
  scrollStyleRowIntoView(linkKey);
  showUrlHint(on ? "已设为标题引导段并保存（标题句止于第一个句号或冒号）" : "已取消标题引导段并保存", "ok");
}

// 段落右键菜单：设为/取消标题引导段（＋有覆盖时可一步恢复原样式）
function openParaContextMenu(x, y, paraEl) {
  const srcKey = paraEl.dataset.key;
  const paraText = paraTextOf(paraEl);
  if (!srcKey || !paraText) return;
  const effKey = effKeyOf(paraText) || srcKey;
  const isLead = isLeadKey(effKey);
  const items = [{
    icon: ICONS.lead,
    label: isLead ? "取消标题引导段" : "设为标题引导段",
    onClick: () => {
      const el = document.querySelector(".para.selected") || paraEl;
      setParaLead(el, !isLead);
    },
  }];
  if (effKey !== srcKey) {
    items.push({ icon: ICONS.reset, label: "恢复原样式", onClick: () => applyStyleToSelectedPara(srcKey) });
  }
  openContextMenu(x, y, items);
}

// 在光标位置弹出菜单（复用浮动子菜单原语：造一个"虚拟锚点"，以 topleft 对齐在鼠标处展开）
function openContextMenu(x, y, items) {
  const anchor = {
    classList: { add() {}, remove() {} },
    getBoundingClientRect: () => ({ left: x, right: x, top: y, bottom: y, width: 0, height: 0 }),
  };
  openSubmenu(anchor, items, { align: "topleft" });
}


// ═══════════════════════════════════════════
// 预览渲染
// ═══════════════════════════════════════════
function specToStyle(spec) {
  const sizePx = (spec.size || 16) * 96 / 72;
  const linePx = (spec.line || 28) * 96 / 72;
  const style = {
    fontFamily: spec.font || "inherit",
    fontSize: sizePx.toFixed(1) + "px",
    lineHeight: linePx.toFixed(1) + "px",
    ...(spec.align === "distribute" ? { textAlign: "justify", textAlignLast: "justify" } : { textAlign: spec.align || "justify" }),
    fontWeight: spec.bold ? "bold" : "normal",
  };
  if (spec.color) style.color = "#" + spec.color;
  if (spec.indentChars) style.textIndent = (spec.indentChars * sizePx).toFixed(1) + "px";
  if (spec.spaceBefore) style.marginTop = ((spec.spaceBefore) * 96 / 72).toFixed(1) + "px";
  if (spec.spaceAfter) style.marginBottom = ((spec.spaceAfter) * 96 / 72).toFixed(1) + "px";
  return style;
}

function applyStyle(el, style) {
  for (const k of Object.keys(style)) el.style[k] = style[k];
}

/*
 * 预览端落款右缩进（像素级视觉居中）。
 * 与 server/convert.js 共用 SignLayout.computeSignRightIndents 的字符网格算法作为 Word 导出真相源；
 * 但浏览器渲染是比例字体（仿宋中数字比汉字窄），直接用"字符数 × 字号"算 padding 会导致
 * "署名以日期为准居中"在视觉上偏左/偏右。这里用 Canvas 测量真实文本宽度，重新算 paddingRight，
 * 让预览看起来居中；Word 导出仍走字符网格，不受影响。
 */
function computePreviewSignRightIndents(blocks, signKey, baseRightIndentChars, font, sizePx) {
  const map = {};
  if (!signKey || !blocks || !blocks.length) return map;
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  ctx.font = `${sizePx}px ${font || "仿宋"}`;
  const hanziW = ctx.measureText("国").width || sizePx;
  const basePx = baseRightIndentChars * hanziW;
  let i = 0;
  while (i < blocks.length) {
    if (blocks[i].style === signKey) {
      const run = [];
      while (i < blocks.length && blocks[i].style === signKey) { run.push(i); i++; }
      if (run.length >= 2) {
        const dateIdx = run[run.length - 1];
        const dateText = (blocks[dateIdx].text || "").trim();
        const dateW = ctx.measureText(dateText).width;
        const dateCenterPx = basePx + dateW / 2;                 // 日期中心线距容器右缘
        for (const ri of run.slice(0, -1)) {
          const signText = (blocks[ri].text || "").trim();
          const signW = ctx.measureText(signText).width;
          // 署名中心与日期中心重合 → 署名右内边距 = 日期中心线 − 署名宽度/2
          map[ri] = Math.max(0, dateCenterPx - signW / 2);
        }
        map[dateIdx] = basePx;
      } else {
        map[run[0]] = basePx;
      }
    } else i++;
  }
  return map;
}

function renderPreview() {
  // v1.50 多页预览（D1=B/D2=自动分页）：.paper 下动态生成 N 个 .page，替换原单页 #page。
  //   分页机制：①样本分页符块 → 强制开新页；②块底超出内容区底界（页高-下边距）且当前页已有内容
  //   → 整块挪新页（块级分页：段/表整体挪页，不拆行）；③单块自身超整页 → 该页如实长高不截断。
  //   已知平台边界：预览自动分页位置 ≠ Word 实际分页（浏览器断行/字体度量与 Word 排版引擎不同）；
  //   分页符位置、页眉/页脚/页码三态与导出 1:1（口径镜像 convert.js buildDocx）。
  const paperEl = document.getElementById("paper");
  const tpl = currentTemplate();
  if (!paperEl) return;
  const keepScroll = paperEl.scrollTop;   // 全量重建后恢复滚动位置
  paperEl.innerHTML = "";
  if (!tpl) return;   // 无模板：清空预览
  // include 降级口径（include 为空时全量展示，2026-08-30：避免「导入结果页全取消勾选」等场景下
  // 服务端已生成样本、前端却把段落全部 display:none → 预览空白。见 previewIncludeFallbackKeys 注释）。
  const include = (Array.isArray(tpl.include) && tpl.include.length) ? tpl.include : previewIncludeFallbackKeys();

  // ── 页面几何（twip→px @96dpi：1/1440 英寸 = 1/15 px；提前计算：动态建页需要）──
  // 旧模板无对应字段回退出厂缺省（A4 11906×16838 twip = 794×1123px，与 CSS 缺省一致）。
  const page = tpl.page || {};
  const mT = (page.marginTop || 2098) / 15;
  const mR = (page.marginRight || 1474) / 15;
  const mB = (page.marginBottom || 1984) / 15;
  const mL = (page.marginLeft || 1587) / 15;
  const pw = (page.width || 11906) / 15;
  const ph = (page.height || 16838) / 15;
  const contentBottom = ph - mB;   // 内容区底界（px，相对页 padding box 顶）——装箱判定界

  // ── 动态建页（v1.50：页序列由装箱结果决定，几何样式统一在此设置）──
  const pages = [];
  const newPageEl = () => {
    const p = document.createElement("div");
    p.className = "page";
    p.style.width = pw.toFixed(1) + "px";
    p.style.minHeight = ph.toFixed(1) + "px";
    p.style.paddingTop = mT.toFixed(1) + "px";
    p.style.paddingRight = mR.toFixed(1) + "px";
    p.style.paddingBottom = mB.toFixed(1) + "px";
    p.style.paddingLeft = mL.toFixed(1) + "px";
    p.style.setProperty("--mT", mT.toFixed(1) + "px");
    p.style.setProperty("--mR", mR.toFixed(1) + "px");
    p.style.setProperty("--mB", mB.toFixed(1) + "px");
    p.style.setProperty("--mL", mL.toFixed(1) + "px");
    paperEl.appendChild(p);
    pages.push(p);
    return p;
  };
  let pageEl = newPageEl();   // 第一页

  // 按样品块清单生成 DOM（数据驱动：页面不再写死任何段落/表格/页脚内容）
  // v1.3：模板无内嵌样本 → sampleBlocks 为空 → 显示空态提示，不再自动生成 ×××× 占位。
  pageEl.innerHTML = "";
  if (!(state.sampleBlocks || []).length) {
    const empty = document.createElement("div");
    empty.className = "sample-empty";
    empty.textContent = "该模板暂无样本，可在「编辑样本」中添加内容。";
    pageEl.appendChild(empty);
  }
  (state.sampleBlocks || []).forEach((blk, idx) => {
    if (blk.type === "table") {
      const wrap = document.createElement("div"); wrap.className = "sample-table";
      if (blk.title) { const cap = document.createElement("div"); cap.className = "sample-table-caption"; cap.textContent = blk.title; wrap.appendChild(cap); }
      const table = document.createElement("table");
      if (blk.header && blk.header.length) {
        const thead = document.createElement("thead");
        const trh = document.createElement("tr");
        blk.header.forEach(c => { const th = document.createElement("th"); th.textContent = c; trh.appendChild(th); });
        thead.appendChild(trh); table.appendChild(thead);
        // 有 thead = 数据表：加 has-thead 让 CSS 偶数行底色生效；版记（header=[]）不加，保留原 word 视觉
        table.classList.add("has-thead");
      }
      if (blk.body && blk.body.length) {
        const tbody = document.createElement("tbody");
        blk.body.forEach(row => {
          const tr = document.createElement("tr");
          row.forEach(c => { const td = document.createElement("td"); td.textContent = c; tr.appendChild(td); });
          tbody.appendChild(tr);
        });
        table.appendChild(tbody);
      }
      wrap.appendChild(table); pageEl.appendChild(wrap);
      return;
    }
    if (blk.type === "footer") {
      // 旧样 [页脚] 示意块（v1.50：字体样式创建时内联——多页后它随流装载，不再依赖单例查询）；
      // 有真实页脚（page.footer 配置）时由 renderHfPreview 隐藏。
      const foot = document.createElement("div"); foot.className = "page-footer"; foot.dataset.key = "footer";
      foot.textContent = blk.text || "";
      const fSpec = tpl.footer || {};
      foot.style.fontFamily = fSpec.font || "inherit";
      foot.style.fontSize = ((fSpec.size || 12) * 96 / 72).toFixed(1) + "px";
      pageEl.appendChild(foot); return;
    }
    // 红线（版头分隔线）：整行红色横线，不参与段落样式体系
    if (blk.type === "hline") {
      const line = document.createElement("div"); line.className = "sample-hline"; line.dataset.key = "hline";
      pageEl.appendChild(line); return;
    }
    // 分页符（v1.50 多页预览）：分页符 = 真实页边界——渲染为隐藏哨兵，装箱阶段触发开新页（不再画虚线示意）
    if (blk.type === "pagebreak") {
      const pb = document.createElement("div"); pb.className = "pagebreak-mark"; pb.dataset.key = "pagebreak";
      pb.style.display = "none";
      pageEl.appendChild(pb); return;
    }
    // 段落（含空行：[空行] 解析为空的正文段落，文本为空时注入 <br> 让其占一行、高度=样式行距，与 Word 一致）
    const el = document.createElement("div");
    el.className = "para";
    el.dataset.key = blk.style;
    el.dataset.paraIdx = String(idx);
    if (idx === state.selectedParaIdx) el.dataset.paraSelected = "1";
    // 正文包一层 .para-body：首行缩进只加在正文块上。
    // 若不包层、把 text-indent 直接放 .para，绝对定位的样式标签（.style-tag）的
    // shrink-to-fit 框宽会被首行缩进撑大（实测缩进 48px → 框宽 +48px），
    // 导致各标签框宽不一致、长标签甚至探出页面左缘。
    const body = document.createElement("div");
    body.className = "para-body";
    if (blk.text && blk.text.trim()) body.textContent = blk.text;
    else body.innerHTML = "<br>";   // 空段落在 Word 里仍占一行（高度=样式行距）；注入 <br> 获行盒，避免 0 高度
    el.appendChild(body);
    pageEl.appendChild(el);
  });

  // 落款右缩进：预览端用 Canvas 真实文本宽度做视觉居中；Word 导出仍走 SignLayout 字符网格算法。
  // 预览块形状 { type:"para", style:<key>, text }；分组按 blk.style === signKey。
  const signDef = allStyleDefs().find(d => d.matchType === "special" && d.match === "sign");
  const signKey = signDef ? signDef.key : null;
  const signEffSpec = signKey ? (signDef.defaults || {}) : {};
  const baseRightIndent = signEffSpec.rightIndentChars || 4;
  const signSizePx = (signEffSpec.size || 16) * 96 / 72;
  // rightIndentMap 值为像素（px），后面直接设置 paddingRight，不再乘 sizePx。
  const rightIndentMap = signKey && state.sampleBlocks
    ? computePreviewSignRightIndents(state.sampleBlocks, signKey, baseRightIndent, signEffSpec.font, signSizePx)
    : {};

  // 段落上样式（沿用原逻辑；重建 DOM 后按 selectedParaKey 恢复选中高亮）
  // 未保留样式（模板里已删除 / 导入时未勾选）的样本段落 → 降级为正文类样式渲染，内容不丢
  // （2026-09-05 用户口径：样式不要可以，内容不能凭空消失；与保存侧 normalizeImportSample 同口径）
  const fbKey = previewBodyFallbackKey(include);
  document.querySelectorAll(".para[data-key]").forEach(el => {
    let key = el.dataset.key;
    const body = el.querySelector(".para-body");
    if (!include.includes(key)) {
      if (!fbKey) { el.style.display = "none"; return; }   // 无正文类样式可兜底（极端）→ 维持原隐藏行为
      key = fbKey;
      el.dataset.key = key;
    }
    el.style.display = "block";
    // 段落样式 = 覆盖后的生效样式（sample.txt 标定源样式，按段落原文查覆盖得实际生效样式；覆盖仅预览演示态）
    // 用 paraTextOf 取正文原文——此刻标签虽未 prepend，但统一走 paraTextOf 更稳，避免将来顺序变化引入标签名
    const effKey = effKeyOf(paraTextOf(el)) || key;
    const effName = styleName(effKey);
    if (effName) el.dataset.styleName = effName; else el.removeAttribute("data-style-name");
    // 选中恢复：仅恢复用户手动选中的那一段（按 sampleBlocks 序号精确判定），同文本多段不一起被选中
    if (el.dataset.paraSelected === "1") el.classList.add("selected");
    // 生效值 = 样式定义自带默认值（styles[].defaults，方案A 唯一真相；迁移后扁平字段已废弃）
    // 口径与右侧面板摘要（styleSummaryOf）、服务端 defsForTemplate（defaults 即生效值）一致。
    const spec = (styleDef(effKey) && styleDef(effKey).defaults) || {};
    const sizePx = (spec.size || 16) * 96 / 72;
    // 行距渲染（与数据 1:1，不兜底）：
    //   auto 倍数（lineMult，如 1.0=单倍）→ line-height = lineMult × 行基准 1.2（与表单 lineBasePt「1 行 ≈ 1.2×字号」同源，
    //     Word 单倍实际行高 ≈ 字号×1.2，直译 1.0 会让折行两行紧贴）；
    //   固定磅值（line）→ line-height 严格 = 磅值（Word exact 同口径）。
    const linePx = spec.lineMult != null ? null : (spec.line || 28) * 96 / 72;
    const baseStyle = {
      fontFamily: spec.font, fontSize: sizePx + "px",
      lineHeight: spec.lineMult != null ? String(spec.lineMult * 1.2) : linePx + "px",
      textAlign: spec.align || "justify",
      ...(spec.color ? { color: "#" + spec.color } : {}),
    };
    // 首行缩进只作用于正文块（.para-body），不落在 .para 上——
    // 否则 text-indent 会参与 .style-tag 的 shrink-to-fit 宽度计算、撑大标签框。
    if (body && spec.indentChars) body.style.textIndent = (spec.indentChars * sizePx).toFixed(1) + "px";
    // 字符缩放/间距（样式级，红头压缩）：缩放 = scaleX 视觉压扁；间距 = letter-spacing（缇→px：1 缇=1/15px @96dpi）。
    // 负 letter-spacing 收窄布局宽度 → 与 Word 一致的单行排布；scaleX 只做视觉（不改布局）。
    if (body && (spec.charScale != null || spec.charSpacing != null)) {
      if (spec.charScale != null) {
        body.style.transform = "scaleX(" + (spec.charScale / 100).toFixed(3) + ")";
        body.style.transformOrigin = "left center";   // WebView2：缩放时基线左对齐，避免 box 错位导致行间空隙
      }
      if (spec.charSpacing != null) body.style.letterSpacing = ((spec.charSpacing) / 15).toFixed(2) + "px";
    }
    // 落款右缩进：右空 N 字 + 署名以日期为中心居中。rightIndentMap 值为 px（预览端已用 Canvas 真实宽度校正）。
    if (body) {
      const riPx = rightIndentMap[Number(el.dataset.paraIdx)];
      if (riPx !== undefined) body.style.paddingRight = riPx.toFixed(1) + "px";
    }
    if (spec.spaceBefore) baseStyle.marginTop = ((spec.spaceBefore) * 96 / 72).toFixed(1) + "px";
    if (spec.spaceAfter) baseStyle.marginBottom = ((spec.spaceAfter) * 96 / 72).toFixed(1) + "px";

    const isLeadMode = !!(styleDef(effKey) && styleDef(effKey).mode === "lead");
    if (isLeadMode) {
      // 标题引导段落：标题句字体/字号/加粗 = lead 样式自身生效值（defaults + tpl 覆盖），
      // 与导出侧 paraLead 首句直接写样式定义一致——2026-09-01 删字符样式「段首标题句」，无独立字符入口。
      const titleFont = spec.font;
      const titleSize = spec.size || 16;
      const titleBold = !!spec.bold;
      // 正文：恒跟随「正文段落」样式生效值（与输出方案 B 一致；bodyFont/bodyBold 兼容分支已删）
      const bodySpec = (styleDef("正文段落") && styleDef("正文段落").defaults) || {};
      applyStyle(el, { ...baseStyle, fontWeight: titleBold ? "bold" : "normal" });
      // 标题句切分（简单规则，预览仅演示）：从段首找第一个结束标志（spec.leadEnd 可配，默认 "。："），
      // 最长 50 个汉字；50 字内找到 → 标题句止于标志（含）；未找到 → 超 50 字取前 50 字、否则整段为标题句，后续是正文。
      const ends = (spec && spec.leadEnd) || "。：";
      const sample = el.textContent || "";
      let splitAt = -1;
      const window50 = sample.slice(0, 50);
      for (const c of ends) {
        const i = window50.indexOf(c);
        if (i >= 0 && (splitAt < 0 || i < splitAt)) splitAt = i;
      }
      let leadTitle = "", leadBody = "";
      if (splitAt >= 0) { leadTitle = sample.slice(0, splitAt + 1); leadBody = sample.slice(splitAt + 1); }
      else if (sample.length > 50) { leadTitle = sample.slice(0, 50); leadBody = sample.slice(50); }
      else leadTitle = sample;
      const titleSizePx = (titleSize * 96 / 72).toFixed(1);
      const bodySizePx = ((bodySpec.size || spec.size || 16) * 96 / 72).toFixed(1);
      body.innerHTML =
        (leadTitle
          ? `<span class="lead-title" style="font-family:${esc(titleFont || "inherit")};font-size:${titleSizePx}px;font-weight:${titleBold ? "bold" : "normal"}">${esc(leadTitle)}</span>`
          : "") +
        (leadBody
          ? `<span style="font-family:${esc(bodySpec.font || "inherit")};font-size:${bodySizePx}px;font-weight:${bodySpec.bold ? "bold" : "normal"}">${esc(leadBody)}</span>`
          : "");
    } else {
      // 其余元素：整体应用样式（含加粗）
      // 若上一次渲染是 lead 模式（el 内有 .lead-title 子节点），先还原为纯文本，
      // 否则残留的 <span> 内联样式会覆盖父元素的新样式。
      // N-3：用 replaceChildren(textNode) 取代 el.textContent = el.textContent，
      // 显式重建文本节点，避免自赋值写法在将来支持段内联格式时误丢子节点。
      if (body && body.querySelector(".lead-title")) body.replaceChildren(document.createTextNode(body.textContent));
      applyStyle(el, { ...baseStyle, fontWeight: spec.bold ? "bold" : "normal" });
    }
    // 段落样式标签（段内行首 inline-block，hover/selected 时浮现）
    // v1.34：放弃段外绝对定位——纸张 padding-left 不固定（默认 106px ≈ 1587 DXA，用户可改小到 30 DXA），
    // 段外位置受纸张边距约束，严重时整块推出纸张白边界、字压蓝色高亮条。改为 prepend 到段落首，
    // 让标签永远在段落里，绝对不超纸张，视觉上也更稳定。
    if (effName) {
      const tag = document.createElement("span");
      tag.className = "style-tag";
      tag.appendChild(document.createTextNode(effName));
      el.prepend(tag);
    }
  });

  // ── 自动分页装箱（v1.50，D1=B/D2=自动分页）──
  // 逐块真实测量（offsetTop/offsetHeight 相对当前页 padding box）：
  //   分页符哨兵 → 结束当前页（自身移除，不渲染）；块底超内容区底界且当前页已有内容块 → 整块挪新页重试；
  //   单块自身超整页 → 留在本页如实长高（min-height 语义，不截断不丢内容）。
  // 非流块（旧样页脚 .page-footer 绝对定位）不参与装箱；display:none 块（include 未勾选/哨兵）不计入
  // curHasContent，避免「空页 + 超高块」误判成已装载页。
  const flowBlocks = Array.from(pageEl.children);
  let curPage = pageEl, curHasContent = false;
  for (let i = 0; i < flowBlocks.length; i++) {
    const el = flowBlocks[i];
    if (el.dataset.key === "pagebreak") {
      el.remove();   // 哨兵移除：分页符本身不占内容
      if (i < flowBlocks.length - 1) { curPage = newPageEl(); curHasContent = false; }
      continue;
    }
    const pos = getComputedStyle(el).position;
    if (pos === "absolute" || pos === "fixed") continue;   // 非流块：留原页，不参与装箱
    if (el.parentNode !== curPage) curPage.appendChild(el);
    if (el.offsetTop + el.offsetHeight > contentBottom + 0.5 && curHasContent) {
      curPage = newPageEl(); curHasContent = false;
      i--;   // 该块在新页重试（首块永不触发切页 → 无死循环）
      continue;
    }
    if (el.offsetHeight > 0) curHasContent = true;
  }

  // ── 每页页眉/页脚/页码（v1.50 多页：按页角色渲染，第 i+1/共 N 页）──
  const totalPages = pages.length;
  pages.forEach((p, idx) => renderHfPreview(p, tpl, mL, mR, idx + 1, totalPages));

  paperEl.scrollTop = keepScroll;   // 恢复滚动位置（重建后总高可能变化，浏览器自动 clamp）
}

// ── 页眉/页脚预览（v1.49 建立口径，v1.50 多页化）：镜像服务端 buildDocx 的装配口径（convert.js）──
// 数据：page.header / page.footer = { font, size, variants:{default,first,even} }，
//   variants[k] = 段落数组（runs：text/page/tab/image）；page.marginHeader = 页眉距纸边缇。
// 口径与导出一致：
//   页眉 —— variants 存在才渲染（旧行为无页眉）；
//   页脚 —— variants 存在原样渲染；page.footer 缺失（旧模板/旧结构）→ 默认页码页脚
//           （format：plain/pageOf/dash/wave；字体 page.footer → footer → 种子 pageFont 兜底）；
//           page.footer = null（导入取消勾选）→ 不渲染。
// v1.50 多页：renderHfPreview 按页角色（第 pageNo/共 totalPages 页）取变体——
//   首页不同开（titlePage || firstNoPg || 首页变体存在）→ 第 1 页取 first 态（未定义 → 空白，Word 行为）；
//   奇偶不同开（oddEven || 偶数变体存在）→ 偶数页取 even 态（未定义 → 空白）；
//   其余页取 default 态。PAGE=当前页号、NUMPAGES=总页数；外侧对齐按页奇偶解析（奇右偶左）；
//   firstNoPg（首页无页码，导出 start=0）→ 第 2 页起显示 pageNo-1。
// 位置（v1.50.1 修锚点 → v1.51 与 ECMA-376 / Word「页面设置→版式」语义一致）：
//   页眉 top = marginHeader（缺省 851 缇 ≈1.5cm）—— 顶端锚定，内容向下排（w:header 语义）
//   页脚 bottom —— **底端**锚定、内容**向上**排（w:footer 语义）；位置一律由 page.numberGap 按国标反算（见下），
//     Word 缺省 w:footer=992 不再参与。实现：ftrEl.style.bottom = 反算值恒定，div 上沿随内容总高 h 自动向上撑开；
//     多段页脚（如「备参资料」default = 页码+空段）段顺序保留（P0 在上、P1 在下），最末段底边贴那条线
//     —— 与 Word 真实渲染一致。
//     内容超高（总高 > 页面底部可用区）→ 顶端进入正文区（Word 同行为：侵占正文区，不裁剪）。
//   左右 = 正文左右边距；纸张/边距/页码在面板调整后 renderPreview 重跑即自动重排。
const HF_HEADER_DEFAULT = 851;   // Word 缺省页眉距纸边（缇）；页脚无缺省——一律走 numberGap 锚点

// ── 页码垂直锚点（GB/T 9704 7.5：页码在版心下边缘**之下**，一字线距版心下边缘 7mm）──
// 与导出侧 convert.js 的同名常量/函数**严格同源**：预览与 Word 必须落在同一位置（一致性铁律）。
//   · PN_LINE_RATIO：页码行高比（锁定精确行距 = 字号 × 1.5）→ 页脚高度可精确计算，不随 Word 默认行距漂移；
//   · 反算：marginFooter = 下边距 − numberGap − 页码锚点偏移（w:footer 为底端锚定，见上）。
//     减号是硬约束：版心下边缘距纸底 = 下边距，页码在其**下方**，距纸底更小。
// page.numberGap 缺失（旧模板文件）→ 按国标缺省 7mm（与导出侧 PAGE_DEFAULT 同值）。
const DEFAULT_NUMBER_GAP = 7;
const PN_LINE_RATIO = 1.5;
/** mm → 缇 */
function mmToDxa(mm) { return Math.round(parseFloat(mm || 0) * DXA_PER_CM / 10); }
/** 页码行高（缇） */
function pnLineTwip(size) { return Math.round((size || 14) * PN_LINE_RATIO * 20); }
/** 单段行高（缇）：页码段用锁定行高，其余段用显式精确行距，缺省按段内最大字号估 */
function paraHeightTwip(p, defSize) {
  const runs = (p && p.runs) || [];
  if (runs.some(r => r.t === "page")) return pnLineTwip(defSize);
  const sizes = runs.map(r => (typeof r.size === "number" && r.size > 0 ? r.size : 0)).filter(Boolean);
  const size = sizes.length ? Math.max(...sizes) : defSize;
  const lh = (p.line != null && p.lineRule) ? p.line : size * PN_LINE_RATIO;
  return Math.round(lh * 20);
}
/** 页脚内容高度（缇）：段行高累加 */
function footerHeightTwip(paras, defSize) {
  return (paras || []).reduce((sum, p) => sum + paraHeightTwip(p, defSize), 0);
}
/** 页码锚点偏移（缇）：含 PAGE 段的行中心相对页脚内容底边的距离（底端锚定下自底向上累加）。
 *  无 PAGE 段（空页脚）时退化为整块中心，再退化到半行高——锚点此时无视觉意义，只求公式稳定。 */
function pnAnchorOffsetTwip(paras, defSize) {
  const list = paras || [];
  let below = 0;
  for (let i = list.length - 1; i >= 0; i--) {
    const h = paraHeightTwip(list[i], defSize);
    if (((list[i].runs) || []).some(r => r.t === "page")) return below + Math.round(h / 2);
    below += h;
  }
  return Math.round(footerHeightTwip(list, defSize) / 2) || Math.round(pnLineTwip(defSize) / 2);
}

// 武文线（复合线）细段/间隙固定宽度，单位 pt（2026-09-09 D）
// 复合边框的 w:sz 只指定**粗段**宽度（ECMA-376），细段与间隙为不随 sz 放大的小常量。
// ✅ 已完成 Word 实测校准（2026-09-09，docx→PDF→解析矢量矩形）：细段 ≈0.72pt、间隙 ≈0.72pt（96dpi 下均取 1px）。
//    旧值为估猜，实测后确认量级正确；若换渲染环境需重测，方法见 hfRunHtml 内注释。
const RULE_THIN_PT = 0.72;   // 细段（实测值，96dpi 下 1px）
const RULE_GAP_PT = 0.72;    // 粗段与细段之间的间隙（实测值）
// 复合线粗段占源 stroke 总宽的比例（2026-09-09 Word→PDF 矢量实测，与导出侧 convert.js RULE_THICK_RATIO 同口径）
// DML 复合线（cmpd=thickThin/thinThick）分段恒为 粗:隙:细 = 3:1:1 —— 声明宽 1/2/3/4.75/6pt 五组实测，
// 粗段恰为 0.6/1.2/1.8/2.85/3.6pt（= 60%），无一例外。故粗段 = RULE_THICK_RATIO × 总宽。
// ⚠ 勿改回「总宽 − 固定量」：边框细段+间隙是 Word 固定 1.44pt，减固定量会把全部误差堆到粗段且非线性
//   （源 2pt → 粗段 0.56 vs 真值 1.20，细 53%；源 6pt → 4.56 vs 3.60，粗 27%）。
const RULE_THICK_RATIO = 0.6;

// 页眉/页脚预览区提示（2026-09-09）：Word 内置 dimming——正文编辑态下页眉页脚被淡化显示，
// 属 Word 自身的显示行为（非文档内容差异），打印/导出 PDF 为正常颜色。挂 title 防用户误判为渲染不一致。
const HF_DIM_TIP = "Word 正文视图下页眉页脚会淡化显示，这是 Word 自身的显示行为，不是文档内容变了；打印或导出 PDF 为正常颜色。";

// 完整字体族 → CSS font-family 回退链（2026-09-09 B）
// 主字体（提取侧已按 w:hint 选好）排第一，其余槽位（eastAsia/ascii/hAnsi/cs）依次作回退——
// 旧实现只写一个字体名，浏览器回退结果与 Word 的回退（Panose 相似字体匹配）不同源。
// 外层 style="…" 用双引号，故字体名一律用**单引号**包裹（CSS 合法，且含空格如「等线 Light」也安全）。
function fontFamilyCss(set, main) {
  const seen = new Set(), list = [];
  const push = n => { if (n && !seen.has(n)) { seen.add(n); list.push(n); } };
  push(main);
  push(set && set.eastAsia);
  push(set && set.ascii);
  push(set && set.hAnsi);
  push(set && set.cs);
  // 字体名里若有单引号（极罕见）先去掉，避免提前闭合
  return list.map(n => "'" + String(n).replace(/'/g, "") + "'").join(",");
}

// run → HTML：run 级字体/字号/加粗/颜色覆盖块级缺省；
// PAGE/NUMPAGES 字段按所属页取值（v1.50：kind=current → 当前页号 pgNo、kind=total → 总页数 pgTotal）
// 横线 run（t:"rule"，红头红色双线，2026-01 D1）：整行横线——
// thickThin=上细下粗（细 border-top + 粗 border-bottom）、thinThick=上粗下细（粗细互换）、dbl/sng=等宽双线/单线；
// 颜色/粗细取导入提取值（默认红）。段序依据见下方「线型 → 自上而下的段序」Word 实测注释。
function hfRunHtml(r, pgNo, pgTotal, paraAlign) {
  if (r.t === "tab") return null;
  if (r.t === "rule") {
    // 颜色白名单：只认 6 位 hex，防异常值拼进 style 造成注入/样式错乱
    const c = (r.color && /^[0-9A-Fa-f]{6}$/.test(r.color)) ? "#" + r.color : "#e02020";
    const ls = r.lineStyle || "thickThin";
    const compound = ls === "thickThin" || ls === "thinThick";
    const pxOf = pt => Math.max(1, Math.round(pt * 96 / 72));   // pt → px（96dpi）
    // 粗段 = 源 stroke 总宽 × RULE_THICK_RATIO（DML 比例恒 3:1:1），与导出 sz=round(总宽*0.6*8)/8pt 同口径。
    // ⚠ 旧式先按 总宽×0.5（2026-09-06 截图校准，偏细约 20%）、后按 总宽−1.44（2026-09-09 总高对齐，粗段误差非线性）；
    //   两次都错在拿总高做对齐目标，而总高里含 Word 固定不可控的 1.44pt。直接对齐粗段才是同语义换算。
    const thickPx = pxOf(compound ? Math.max(0.5, (r.weight || 2) * RULE_THICK_RATIO) : (r.weight || 2));
    // ⚠ 细段/间隙为**固定值**、不随 sz 放大（2026-09-09 D 修正）：
    //   ECMA-376 规定复合边框 w:sz 只表示**粗段**宽度，Word 的细线恒为发丝级、间隙亦为小常量；
    //   旧实现按 thickPx*0.4 / *0.25 等比放大 → sz 越大细段越粗，与 Word 实际渲染相反。
    //   ⚠ 两常量为**近似值**（Word 未公开细段/间隙规格），需实测校准：
    //     校准法 = 导出两个 sz 相差 3 倍的 docx → Word 转 PDF → 量总高差是否等于粗段差（验证固定值假设），
    //     不符则只调这两个常量即可，不必改结构。
    const thinPx = compound ? pxOf(RULE_THIN_PT) : 0;
    const gapPx = compound ? pxOf(RULE_GAP_PT) : 0;
    // 显式三段结构（粗/隙/细）替代 border-top+border-bottom 近似：比例可控、间隙真实存在，
    // 而 border 方案两条边直接贴边、只能靠撑高内容区伪造间隙。
    const bar = hh => `<span style="display:block;height:${hh}px;background:${c}"></span>`;
    const spc = hh => (hh > 0 ? `<span style="display:block;height:${hh}px"></span>` : "");
    // 线型 → 自上而下的段序（★Word 实测真值 2026-09-09：docx→PDF→解析矢量填充矩形得到）：
    //   thickThin（导出 thickThinSmallGap）= **上细下粗**；thinThick（导出 thinThickSmallGap）= **上粗下细**。
    //   ⚠ 与字面直觉相反：lineStyle 已是 **ECMA-376 段落边框口径**（ruleFromShape 已从 DML 形状名换算过），
    //     而边框命名的 thick/thin 次序对 bottom 边框是「先外后内」（外=下），故 thickThin → 粗段在下。
    //     实测：thickThinSmallGap 粗 2.28pt 在下 / 细 0.72pt 在上 / 间隙 0.72pt；thinThickSmallGap 反之。
    //     源 DML「cmpd=thickThin + flipV=1」真实渲染也是上细下粗（粗 2.85pt 下、细 0.95pt 上，总宽 4.75pt）。
    let inner, total;
    if (ls === "sng") { inner = bar(thickPx); total = thickPx; }
    else if (ls === "dbl") { inner = bar(thinPx) + spc(gapPx) + bar(thinPx); total = thinPx * 2 + gapPx; }
    else if (ls === "thinThick") { inner = bar(thickPx) + spc(gapPx) + bar(thinPx); total = thickPx + gapPx + thinPx; }
    else { inner = bar(thinPx) + spc(gapPx) + bar(thickPx); total = thinPx + gapPx + thickPx; }
    // 横向几何复刻（与导出口径一致：线宽保持源值、左右均分出血整体居中）：
    // 定宽 + left:50% 相对位移居中（1缇=1/15px @96dpi）——线宽大于版心容器时 margin:auto 会失效归零，
    // 相对位移对任意宽度都对称。无几何（旧模板）回落版心撑满。
    const geo = (r.w != null) ? `width:${(r.w / 15).toFixed(1)}px;position:relative;left:50%;transform:translateX(-50%);` : "";
    return `<span class="hf-rule" style="display:block;height:${total}px;${geo}">${inner}</span>`;
  }
  const st = [];
  if (r.fontSet) st.push(`font-family:${fontFamilyCss(r.fontSet, r.font)}`);   // 2026-09-09 B：完整字体族回退链
  else if (r.font) st.push(`font-family:${esc(r.font)}`);
  if (r.size) st.push(`font-size:${(r.size * 96 / 72).toFixed(1)}px`);
  if (r.bold) st.push("font-weight:bold");
  if (r.color) st.push(`color:#${r.color}`);
  // 提取侧本就抓全的其余 4 项 run 级格式 → 预览回放（2026-09-06 D2 闭环补齐，与导出口径一致）
  if (r.italic) st.push("font-style:italic");
  if (r.bgColor) st.push(`background-color:#${r.bgColor}`);
  const deco = [r.underline ? "underline" : "", r.strike ? "line-through" : ""].filter(Boolean).join(" ");
  if (deco) st.push(`text-decoration:${deco}`);   // 下划线+删除线并存须合并成一个 text-decoration（后者覆盖前者）
  // 字符间距缇 → letter-spacing（1缇=1/15px @96dpi，与正文样式级同口径）；缩放 → scaleX。
  // distribute（分散对齐）段：整段 inline-block 会成为单个盒子，text-align-last:justify 无法在盒内分配字距
  // → 逐字压缩：每字符 inline-block scaleX（布局宽度不变，justify 照常分配字距），视觉 ≈ Word「压字形宽后重排撑满」。
  const scale = r.charScale != null && r.charScale !== 100 ? (r.charScale / 100).toFixed(3) : null;
  if (scale && paraAlign !== "distribute") st.push("display:inline-block", `transform:scaleX(${scale})`, "transform-origin:left");
  if (r.charSpacing) st.push(`letter-spacing:${(r.charSpacing / 15).toFixed(2)}px`);
  const attr = st.length ? ` style="${st.join(";")}"` : "";
  if (r.t === "page") return `<span${attr}>${r.kind === "total" ? pgTotal : pgNo}</span>`;
  if (r.t === "image") {
    if (!r.data || !r.width || !r.height) return "";
    return `<img src="data:image/${r.mediaType || "png"};base64,${r.data}" width="${r.width}" height="${r.height}" alt="">`;
  }
  if (scale && paraAlign === "distribute") {
    const chars = [...String(r.text || "")].map(ch => `<span style="display:inline-block;transform:scaleX(${scale})">${esc(ch)}</span>`).join("");
    return `<span${attr}>${chars}</span>`;
  }
  return `<span${attr}>${esc(r.text || "")}</span>`;
}

// 页码格式归一化（与 convert.js 同口径）：自由形态 → 4 键，4 键幂等
function normalizePageFormat(fmt) {
  const s = String(fmt || "").trim();
  if (s === "plain" || s === "pageOf" || s === "dash" || s === "wave") return s;
  const t = s.replace(/\s+/g, "");
  if (!t) return "dash";
  if (/^\{n\}$/.test(t)) return "plain";
  if (/^[~～]\{n\}[~～]$/.test(t)) return "wave";   // 半角/全角波浪都识别（与 convert.js 同口径）
  if (/^—\{n\}—$/.test(t)) return "dash";
  if (/^第\{n\}页共\{N\}页$/.test(t) || /\{N\}/.test(t) || /第.*\{n\}.*页/.test(t)) return "pageOf";
  return "dash";
}
// 页码取态（2026-08-31 全局化，与 convert.js 同口径）：{ format, align, first }；
//   首页无页码（first=false）→ {none:true}；「无」（format=none）= 全篇无页码（2026-09-01 用户澄清）→ 所有态 none+whole
//   （whole：页码段整段清空，两侧横线/波浪线等装饰一并清除，2026-09-01 用户需求）
function pageNumberForState(pn, key) {
  if (!pn) return null;
  if (pn.format === "none") return { none: true, whole: true };
  if (pn.format !== undefined || pn.align !== undefined || pn.first !== undefined) {
    if (key === "first" && pn.first === false) return { none: true };
    return { format: pn.format, align: pn.align };
  }
  if (pn[key]) return pn[key];
  if (pn.default) return pn.default;
  if (pn.format !== undefined) return pn;
  return null;
}
// 段 → HTML：无 tab 走行内流；有 tab + 制表位 → 相对容器 + 制表位段绝对定位
//   （右对齐段 translateX(-100%)、居中段 -50%，无需测宽；无制表位的 tab 兜底 2em 间隔）
// pageNum：页码设置（{format, align}，v1.50：align 的 outside 已由调用方按页奇偶解析成 left/right）；
//   页脚页码段按设置渲染（改页码样式/对齐预览联动，2026-08-30）
//   pgNo/pgTotal（v1.50 多页）：PAGE=当前页号、NUMPAGES=总页数；
//   无页码（{none:true}）：剔除 PAGE/NUMPAGES 字段，保留段其他内容（首页页脚空白）
function hfParaHtml(p, pageNum, pgNo, pgTotal) {
  let runs = p.runs || [];
  if (pageNum && runs.some(r => r.t === "page")) {
    if (pageNum.whole) return "";   // 全篇无页码：页码段整段清空——两侧横线/波浪线等装饰一并清除（2026-09-01 用户需求）
    if (pageNum.none) {
      const rest = runs.filter(r => r.t !== "page");
      if (!rest.length) return "";   // 纯页码段 → 首页页脚空白
      p = Object.assign({}, p, { runs: rest });
      runs = rest;   // 修复（v1.50）：剔除 PAGE 后渲染也用 rest——原版仅改 p.runs，tab 分组仍用旧 runs，PAGE 字段残留
      pageNum = null;
    } else {
      // 页码段一律按设置重出标准形态（2026-09-01 用户拍板：删除自然形态保留/空格归一旁路——
      // 页码由 pageNumber 设置唯一决定，docx 原文的空格/文本差异（紧贴「—1—」等）不得绕过设置）
      const fmt = normalizePageFormat(pageNum.format);
      const align = pageNum.align || "left";
      const txt = fmt === "pageOf" ? `第 ${pgNo} 页 共 ${pgTotal} 页`
        : fmt === "wave" ? `～ ${pgNo} ～`
        : fmt === "plain" ? String(pgNo)
        : `— ${pgNo} —`;
      return `<div class="page-hf-line" style="text-align:${align}">${esc(txt)}</div>`;
    }
  }
  const align = (p.align === "outside" ? "right" : p.align) || "left";
  const alignCss = align === "distribute" ? "justify;text-align-last:justify" : align;   // 分散对齐：CSS 无同名值 → justify+末行 justify（单行段即撑满版心，与 Word distribute 一致）
  // 段级格式回放（2026-09-06 D2 闭环，与导出 hfParaSpacing/hfParaIndent 同口径）：
  //   行高 exact/atLeast → line-height px（磅→px）；段前后间距 → margin；缩进（缇→px ÷15）→ text-indent/padding
  const ps = [];
  if (p.line != null && p.lineRule) ps.push(`line-height:${(p.line * 96 / 72).toFixed(1)}px`);
  if (p.spaceBefore) ps.push(`margin-top:${(p.spaceBefore * 96 / 72).toFixed(1)}px`);
  if (p.spaceAfter) ps.push(`margin-bottom:${(p.spaceAfter * 96 / 72).toFixed(1)}px`);
  if (p.indFirstLine) ps.push(`text-indent:${(p.indFirstLine / 15).toFixed(1)}px`);
  if (p.indLeft) ps.push(`padding-left:${(p.indLeft / 15).toFixed(1)}px`);
  if (p.indRight) ps.push(`padding-right:${(p.indRight / 15).toFixed(1)}px`);
  const psAttr = ps.length ? ";" + ps.join(";") : "";
  if (!runs.some(r => r.t === "tab")) {
    // 空段渲染占位（2026-09-03 C）：纯空段（runs=[]，源文档页脚回车残留）此前渲染成空 div
    // 高度 0 不占位 → 预览页脚「只有页码 1 行」、与 Word 真实（空段占 1 行高、页码被顶高 1 行）不符。
    // 空 body 补 &nbsp; 让空行如实占一行高 → 预览与 Word 一致（一致性铁律）。
    const body = runs.map(r => hfRunHtml(r, pgNo, pgTotal, align)).join("");
    return `<div class="page-hf-line" style="text-align:${alignCss}${psAttr}">${body || "&nbsp;"}</div>`;
  }
  const stops = Array.isArray(p.tabStops) ? p.tabStops : [];
  const segs = [{ runs: [], stop: null }];
  let tabIdx = 0;
  for (const r of runs) {
    if (r.t === "tab") { tabIdx++; segs.push({ runs: [], stop: stops[tabIdx - 1] || null }); }
    else segs[segs.length - 1].runs.push(r);
  }
  const inner = segs.map((s, i) => {
    const body = s.runs.map(r => hfRunHtml(r, pgNo, pgTotal, align)).join("");
    if (i === 0) return body;                                   // 首段行内流
    if (!s.stop) return `<span class="hf-tab"></span>${body}`;  // 无制表位兜底
    const posPx = ((s.stop.pos || 0) / 15).toFixed(1);
    const tf = s.stop.align === "right" ? "translateX(-100%)" : s.stop.align === "center" ? "translateX(-50%)" : "";
    const ldCss = s.stop.leader ? ";border-bottom:1px dotted currentColor" : "";   // 前导符（点线等）→ 虚线下边框近似
    return `<span class="hf-seg" style="left:${posPx}px;transform:${tf}${ldCss}">${body}</span>`;
  }).join("");
  return `<div class="page-hf-line" style="text-align:${alignCss}${psAttr}">${inner || "&nbsp;"}</div>`;
}

// 三态取段（v1.50 多页化）：按页角色取变体——
//   default：无变体 → null（不渲染，旧行为）
//   first：变体未定义/无内容 → []（空白，Word「缺 reference = 空白」，9/1 D2 拍板）
//   even：变体未定义/无内容 → 回落到 default 段（v1.50.2）。
//     用户常配 oddEven 但只配 default 段，期望奇偶页沿用 default 内容只切位置/对齐——
//     配了空 even（hfHasVariant=false）按未配处理；导出侧 oddEven=true + variants.even=null/空 时
//     也回填 default 段给 even ref（避免缺 ref 在 settings.evenAndOddHeaders 下偶数页脚空白，
//     「备参」翻车即此场景——前后页页码内容不一致）。
function hfPickParas(cfg, role) {
  if (!cfg || !cfg.variants) return null;
  if (role === "default") return cfg.variants.default || null;
  if (role === "first") {
    const f = cfg.variants.first;
    return (f && f.length) ? f : [];
  }
  // role === "even"：变体存在且有内容 → 用 even；否则（未定义/空数组/空段）回落到 default
  const e = cfg.variants.even;
  if (Array.isArray(e) && e.length && hfVarHasContent(e)) return e;
  return cfg.variants.default || [];
}

// 变体内容判定（与 convert.js hfVarHasContent 同口径）：段内任一 run 有实质内容（非空文本/PAGE/横线/图片）→ 有内容
function hfVarHasContent(paras) {
  return (paras || []).some(p => (p.runs || []).some(r => {
    if (r.t === "text") return !!(r.text || "").trim();
    return r.t === "page" || r.t === "rule" || r.t === "image";
  }));
}

// 变体存在判定（镜像导出 hasFirst/hasEven：variants[k] 有实质内容 ⇔ 生成对应部件）
// 2026-09-01：空变体（源文档引用空 first/even 部件）不算存在 → 首页/偶数页回落到 default 页眉
function hfHasVariant(cfg, k) {
  return !!(cfg && cfg.variants && hfVarHasContent(cfg.variants[k]));
}

function renderHfPreview(pageEl, tpl, mLpx, mRpx, pageNo, totalPages) {
  const pg = tpl.page || {};
  const pn = pg.pageNumber || null;
  const isFirst = pageNo === 1;
  const isEven = pageNo % 2 === 0;
  // 首页不同开（镜像导出 titlePage 写入条件：hasFirst || page.titlePage || firstNoPg）；
  // firstNoPg 与 convert.js 同口径：format≠none 且 first===false（format=none 全篇无页码，不因它开首页态）
  const firstNoPg = !!(pn && pn.format !== "none" && pn.first === false);
  const firstMode = !!(pg.titlePage || firstNoPg || hfHasVariant(pg.header, "first") || hfHasVariant(pg.footer, "first"));
  // 奇偶不同开（镜像导出 evenAndOddHeaderAndFooters 写入条件 hasEven，面板开关 page.oddEven 同义）
  const evenMode = !!(pg.oddEven || hfHasVariant(pg.header, "even") || hfHasVariant(pg.footer, "even"));
  const role = (isFirst && firstMode) ? "first" : (isEven && evenMode) ? "even" : "default";
  // 页码取态（同口径 pageNumberForState）；外侧对齐按页奇偶解析（奇数页右/偶数页左，Word 帧行为），
  // hfParaHtml 收到的 align 已是 left/center/right
  let pnState = pageNumberForState(pn, role);
  if (pnState && pnState.align === "outside") {
    pnState = Object.assign({}, pnState, { align: isEven ? "left" : "right" });
  }
  // 显示页号：firstNoPg（导出 start=0）→ 第 2 页起显示 pageNo-1、第 1 页无页码（pnState 已 none）
  const displayNo = firstNoPg ? pageNo - 1 : pageNo;

  // 页眉：无配置不渲染（与导出口径一致）
  const hdrParas = hfPickParas(pg.header, role);
  if (pg.header && hdrParas && hdrParas.length) {
    const el = document.createElement("div");
    el.className = "page-hf page-hf-header";
    el.title = HF_DIM_TIP;   // 2026-09-09：Word dimming 提示（非内容差异）
    el.style.top = ((pg.marginHeader || HF_HEADER_DEFAULT) / 15).toFixed(1) + "px";
    el.style.left = mLpx + "px";
    el.style.right = mRpx + "px";
    if (pg.header.font) el.style.fontFamily = pg.header.font;
    if (pg.header.size) el.style.fontSize = (pg.header.size * 96 / 72).toFixed(1) + "px";
    el.innerHTML = hdrParas.map(p => hfParaHtml(p, null, displayNo, totalPages)).join("");
    pageEl.appendChild(el);
  }
  // 页脚：variants 原样；页码联动兜底（2026-09-01，format=页码唯一真源，与导出 hfPart/装配同口径）：
  //   footer 未配置（旧模板）→ 页码兜底；footer=null 或变体不含 PAGE 段 → 页码设置有样式（非 none）时补标准页码
  const ftrCfg = pg.footer !== undefined ? pg.footer : undefined;
  const ftrParas = (ftrCfg && ftrCfg.variants) ? hfPickParas(ftrCfg, role) : null;
  // 标准页码行 HTML（与导出 pageNumParagraph 同文案/对齐口径）
  const pnLineHtml = (st, no) => {
    const fmt = normalizePageFormat(st && st.format);
    const align = (st && st.align) || "center";
    const txt = fmt === "pageOf" ? `第 ${no} 页 共 ${totalPages} 页`
      : fmt === "wave" ? `～ ${no} ～`
      : fmt === "plain" ? String(no)
      : `— ${no} —`;
    return `<div class="page-hf-line" style="text-align:${align}">${esc(txt)}</div>`;
  };
  const pageSeedDef = (typeof seedStyles !== "undefined" ? seedStyles : []).find(s => s.category === "page");
  const seedD = (pageSeedDef && pageSeedDef.defaults) || {};
  const fFont = (ftrCfg && ftrCfg.font) || (tpl.footer && tpl.footer.font) || seedD.pageFont || "宋体";
  const fSize = (ftrCfg && ftrCfg.size) || (tpl.footer && tpl.footer.size) || seedD.pageSize || 14;
  let ftrEl = null;
  if (ftrCfg && ftrCfg.variants) {
    if (ftrParas && ftrParas.length) {
      ftrEl = document.createElement("div");
      ftrEl.className = "page-hf page-hf-footer";
      if (ftrCfg.font) ftrEl.style.fontFamily = ftrCfg.font;
      if (ftrCfg.size) ftrEl.style.fontSize = (ftrCfg.size * 96 / 72).toFixed(1) + "px";
      // 页码段（含 PAGE 字段）按当前态 pageNumber 设置渲染（改样式/对齐预览联动；无设置走原样）
      ftrEl.innerHTML = ftrParas.map(p => hfParaHtml(p, pnState, displayNo, totalPages)).join("");
      // 该态页脚段不含 PAGE 字段且页码设置有样式 → 追加标准页码行（导入无页码文档改页码样式的联动）
      if (pnState && !pnState.none && !ftrParas.some(p => (p.runs || []).some(r => r.t === "page"))) {
        ftrEl.innerHTML += pnLineHtml(pnState, displayNo);
      }
    } else if (pnState && !pnState.none) {
      // 变体为空段（如导入无页码文档）+ 页码设置有样式 → 整个页脚按标准页码渲染（与导出追加口径一致）
      ftrEl = document.createElement("div");
      ftrEl.className = "page-hf page-hf-footer";
      ftrEl.style.fontFamily = fFont;
      ftrEl.style.fontSize = (fSize * 96 / 72).toFixed(1) + "px";
      ftrEl.innerHTML = pnLineHtml(pnState, displayNo);
    }
    // 首页不同且未定义首页态 → 空白（不渲染、也不落页码兜底，与 Word 一致）
  } else if (ftrCfg !== null || (pnState && !pnState.none)) {
    if (pnState && pnState.none) { /* 该页无页码且无页脚内容 → 页脚空白（首页无页码态） */ }
    else {
      // 默认页码页脚（镜像 renderPageNumber）：页号取所属页（v1.50 多页）
      ftrEl = document.createElement("div");
      ftrEl.className = "page-hf page-hf-footer";
      ftrEl.style.fontFamily = fFont;
      ftrEl.style.fontSize = (fSize * 96 / 72).toFixed(1) + "px";
      ftrEl.innerHTML = pnLineHtml(pnState || {}, displayNo);
    }
  }
  if (ftrEl) {
    ftrEl.title = HF_DIM_TIP;   // 2026-09-09：Word dimming 提示（三处页脚分支统一在此挂载）
    ftrEl.style.left = mLpx + "px";
    ftrEl.style.right = mRpx + "px";
    // 锁定页脚行高（与导出 pageNumParagraph 的精确行距同源）：反算页脚位置需要**确定的内容高度**，
    // 否则预览行高按浏览器默认、Word 按 docDefaults → 页码位置不一致。国标锚点是唯一口径 ⇒ 无条件锁定。
    ftrEl.style.lineHeight = (pnLineTwip(fSize) / 20 * 96 / 72).toFixed(1) + "px";
    pageEl.appendChild(ftrEl);
    // 页脚定位（v1.50.1 修锚点 → v1.51 与 ECMA-376 / Word 一致）：
    // w:footer = 页脚底端距纸底 → div.bottom = 反算值恒定（底端锚定），内容向上排。
    // 旧实现写成 bottom = max(0, 该值 - h)（顶端锚定、内容向下）→ 一行页码即越过下边距虚线、与 Word 真实渲染相反。
    // 正确：底端锚定 + 内容向上排 → 段顺序保留、最末段底边贴该线；多段时页码在上方段（与 Word 视觉一致）。
    //   内容超高 → 顶端向上进入正文区（Word 同行为：侵占正文区）。
    // 需先入 DOM 才能测高 → append 后定位；图片等异步内容在下一帧复测修正。
    // 页码锚点（GB/T 9704 7.5）：与导出 FOOTER_TWIP **同公式**反算页脚底端距纸底
    // （底端锚定：下边距 − numberGap − 页码锚点偏移；页码在版心下边缘之下，故为减）。
    // 唯一口径：numberGap 缺失按国标缺省 7mm，不再回落到 marginFooter。
    const mfTw = (() => {
      const gap = pg.numberGap != null ? pg.numberGap : DEFAULT_NUMBER_GAP;
      const paras = (ftrParas && ftrParas.length) ? ftrParas : [{ runs: [{ t: "page" }] }];
      return Math.max(0, (pg.marginBottom || 1984) - mmToDxa(gap) - pnAnchorOffsetTwip(paras, fSize));
    })();
    const mfPx = mfTw / 15;
    const placeFooter = () => {
      // 底端锚定：bottom 恒定 = 反算值（缇 → px），div 上沿随内容总高自动向上撑开；
      // 多段页脚（如「备参资料」default = 页码+空段）的页码段在空段上方，与 Word 一致。
      ftrEl.style.bottom = mfPx.toFixed(1) + "px";
    };
    placeFooter();
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(placeFooter);
    // 真实页脚已呈现 → 隐藏旧样 [页脚] 块（.page-footer），避免与真实页脚重叠
    const legacy = pageEl.querySelector(".page-footer");
    if (legacy) legacy.style.display = "none";
  }
}

// ── 模板样本加载（每模板一份：内嵌 templates[].sample；v1.3 无样本 → 预览空态提示，不再自动补齐占位） ──
let _sampleKey = "";   // 上次已加载样本的 "模板集|模板" 标识（去重，renderAll 频繁调用不重载）

async function loadSample() {
  const tset = state.currentFile;
  const tpl = state.currentTemplateKey;
  if (!tset || !tpl) {
    state.sampleBlocks = [];
    state.sampleRaw = "";
    _sampleKey = "";
    renderPreview();
    return;
  }
  const key = tset + "|" + tpl;
  if (key === _sampleKey) return;   // 未切换模板，保留现有样本
  _sampleKey = key;
  state.sampleBlocks = [];
  try {
    const data = await api(`/api/sample?tset=${encodeURIComponent(tset)}&tpl=${encodeURIComponent(tpl)}`);
    if (key !== _sampleKey) return;   // 已切换到其他模板，丢弃过期结果
    state.sampleBlocks = Array.isArray(data.blocks) ? data.blocks : [];
    state.sampleFile = data.file || "";
    state.sampleRaw = (typeof data.raw === "string") ? data.raw : "";
  } catch (e) {
    console.error("加载样本失败:", e);
    state.sampleBlocks = [];
    state.sampleRaw = "";
    state.sampleFile = "";
    // 失败自愈（2026-08-30）：仅当失败请求仍是当前模板时才复位去重标记，让下一次 renderAll()
    // 自动重拉样本——否则一次性失败（如落盘竞态：模板刚保存、文件尚未含新 key）会让预览永久空白。
    // 若期间已切到其他模板（key !== _sampleKey），保留新请求的去重态，不干扰其正常返回。
    if (key === _sampleKey) _sampleKey = "";
  }
  renderPreview();
}

// 页面内样本编辑器：编辑当前模板样本，保存落 tpl.sample（随模板集落盘），刷新预览
async function editSampleInPage() {
  const tpl = currentTemplate();
  if (!tpl || !state.currentFile) { showUrlHint("请先选择模板", "err"); return; }
  const tset = state.currentFile;
  const tplKey = tpl.key;
  // 预填：优先用 loadSample 缓存的 raw；为空（切换后尚未加载完成）再请求一次
  let raw = state.sampleRaw || "";
  if (!raw) {
    try {
      const d = await api(`/api/sample?tset=${encodeURIComponent(tset)}&tpl=${encodeURIComponent(tplKey)}`);
      raw = (d && typeof d.raw === "string") ? d.raw : "";
    } catch (e) { /* 忽略，raw 留空 */ }
  }
  const bodyHtml =
    `<div class="modal-msg">编辑「${esc(tpl.name)}」样本内容，保存后仅应用于当前模板，预览即时刷新。</div>` +
    `<textarea id="sample-editor-area" class="sample-editor-textarea" spellcheck="false">${esc(raw)}</textarea>` +
    `<div class="sample-editor-foot">` +
      `<span class="sample-editor-hint" id="sample-editor-hint"></span>` +
    `</div>`;
  openModal({ title: "样本编辑 · " + tpl.name, body: bodyHtml, hideButtons: true });
  const modalEl = document.querySelector("#modal-overlay .modal");
  if (modalEl) modalEl.classList.add("sample-editor-modal");
  const area = document.getElementById("sample-editor-area");
  const hint = document.getElementById("sample-editor-hint");
  setTimeout(() => area && area.focus(), 50);
  // 自定义保存 / 取消按钮（默认 OK/Cancel 已隐藏；「恢复默认」按钮已删，2026-09-01 用户拍板）
  const foot = document.querySelector(".sample-editor-foot");
  const cancelBtn = document.createElement("button");
  cancelBtn.className = "small"; cancelBtn.textContent = "取消";
  const saveBtn = document.createElement("button");
  saveBtn.className = "small primary"; saveBtn.textContent = "保存并刷新";
  foot.appendChild(cancelBtn); foot.appendChild(saveBtn);
  cancelBtn.addEventListener("click", () => closeModal(null));
  saveBtn.addEventListener("click", async () => {
    saveBtn.disabled = true;
    try {
      const text = area.value;
      if (text && text.trim()) tpl.sample = text;
      else delete tpl.sample;           // 清空 → 模板无样本（预览显示「暂无样本」提示，v1.3 不再自动生成占位）
      await persistCurrentSet();        // 样本随模板集落盘
      _sampleKey = "";                  // 强制重新拉取样本
      await loadSample();
      closeModal(true);
      toast("已保存并刷新预览", "ok");
    } catch (e) {
      saveBtn.disabled = false;
      hint.textContent = "保存失败：" + e.message;
      hint.className = "sample-editor-hint err";
    }
  });
}

// ═══════════════════════════════════════════
// 元素样式展开编辑（点击样式项后显示）
// ═══════════════════════════════════════════
function buildField(label, controlHtml) {
  return `<div class="prop-row"><label>${label}</label>${controlHtml}</div>`;
}

function fontSelect(id, value) {
  const v = String(value || "").trim();
  let opts = FONTS.map(f => `<option value="${esc(f)}" ${f === v ? "selected" : ""}>${esc(f)}</option>`).join("");
  if (v && !FONTS.includes(v)) opts = `<option value="${esc(v)}" selected>${esc(v)}</option>` + opts;
  // 空值（未设置扁平字体 = 跟随种子默认 / 跟随段落）：显式占位并选中，
  // 避免浏览器无 selected 时暗选列表第一项（如「方正小标宋简体」）造成显示假象
  if (!v) opts = `<option value="" selected>（跟随段落）</option>` + opts;
  return `<select id="${id}">${opts}</select>`;
}

function sizeSelect(id, value, allowEmpty) {
  const pt = parseFloat(value);
  let opts = "";
  // 空值（char 样式字号空 = 跟随段落）：显式占位并选中，避免暗选第一项造成显示假象
  if (allowEmpty && !(pt > 0)) opts = `<option value="" selected>（跟随段落）</option>`;
  opts += CN_SIZES.map(s => `<option value="${s.pt}" ${Math.abs(s.pt - pt) < 0.01 ? "selected" : ""}>${s.label}（${s.pt}pt）</option>`).join("");
  if (pt > 0) {
    const exact = ptToCnSize(pt);
    if (!exact) opts += `<option value="${pt}" selected>自定义（${pt}pt）</option>`;
  }
  return `<select id="${id}">${opts}</select>`;
}

// 磅/行 单位共用取值范围（2026-09-02 修）：磅单位 min/max 即磅口径（step 1）；
// 行单位下 min 用行语义下限（行距 0.5 行起步、段前段后 0）——直接沿用磅的 min（如行距 12）
// 会把「行」当磅数校验，行单位下连 1.5 行都不让输（spinner/输入红框），step 0.5 为倍数粒度；
// max 按行基准 base 换算同磅上限（120 磅 ≈ 6.25 行），避免行单位下还能填出 120 行的荒谬值。
function spacingRange(unit, base, ptMin, ptMax, lineMin = 0) {
  if (unit !== "line") return { min: ptMin, max: ptMax, step: 1 };
  const max = base > 0 ? Math.round((ptMax / base) * 100) / 100 : ptMax;
  return { min: lineMin, max: Math.max(lineMin, max), step: 0.5 };
}

function spacingField(id, value, base, opts = {}) {
  // value: pt；base: 1"行"对应的磅值；单位切换 磅/行；unit 回填单位选择（记忆后重建列表不回弹「磅」）
  const unit = opts.unit === "line" ? "line" : "pt";
  const ptMin = opts.min !== undefined ? opts.min : 0;
  const ptMax = opts.max !== undefined ? opts.max : 200;
  const r = spacingRange(unit, base, ptMin, ptMax, opts.lineMin);
  const fmt = v => (+v).toFixed(2).replace(/\.?0+$/, "");
  const shown = (unit === "line" && base > 0) ? fmt(value / base) : (+value || 0);
  return `<span class="spacing-wrap">
    <input type="number" id="${id}" min="${r.min}" max="${r.max}" step="${r.step}" value="${shown}">
    <select id="${id}-unit" class="spacing-unit">
      <option value="pt"${unit === "pt" ? " selected" : ""}>磅</option>
      <option value="line"${unit === "line" ? " selected" : ""}>行</option>
    </select>
  </span>`;
}

// 行距的"行"基准：1 行 ≈ 1.2 × 字号（单倍行距语义）
function lineBasePt(spec) {
  return Math.round((parseFloat(spec.size) || 16) * 1.2 * 10) / 10;
}

// 磅/行 单位切换时同步输入框范围（min/max/step 随单位换算，见 spacingRange）
function applySpacingStep(input, unit, base, ptMin, ptMax, lineMin = 0) {
  if (!input) return;
  const r = spacingRange(unit, base, ptMin, ptMax, lineMin);
  input.min = r.min;
  input.max = r.max;
  input.step = r.step;
}

// 对齐按钮组（图标按钮；values 可选子集，如页码只需 左/中/右）
function alignButtons(id, value, values = ALIGNS, disabled = []) {
  const icons = {
    left:    `<i style="width:14px"></i><i style="width:10px"></i><i style="width:12px"></i>`,
    center:  `<i style="width:12px;margin:0 auto"></i><i style="width:14px;margin:0 auto"></i><i style="width:10px;margin:0 auto"></i>`,
    right:   `<i style="width:14px;margin-left:auto"></i><i style="width:10px;margin-left:auto"></i><i style="width:12px;margin-left:auto"></i>`,
    justify: `<i style="width:14px"></i><i style="width:14px"></i><i style="width:14px"></i>`,
    // 外侧：书脊竖线（高≈按钮 80%）+ 左右靠下圆点（页码贴各自外缘底部：奇数页右、偶数页左）
    outside: `<span style="position:relative;display:block;width:100%;height:100%">` +
      `<i style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:1.5px;height:22px;background:currentColor"></i>` +
      `<i style="position:absolute;left:-5px;bottom:-2px;width:4px;height:4px;border-radius:50%;background:currentColor"></i>` +
      `<i style="position:absolute;right:-5px;bottom:-2px;width:4px;height:4px;border-radius:50%;background:currentColor"></i>` +
      `</span>`,
  };
  const labels = { left: "左对齐", center: "居中", right: "右对齐", justify: "两端对齐", outside: "外侧（奇数右、偶数左）" };
  const dis = new Set(disabled);
  return `<div class="align-btns" id="${id}" role="group">
    ${values.map(a => {
      const isDis = dis.has(a.v);
      return `<button type="button" class="align-btn${value === a.v ? " active" : ""}" data-align="${a.v}" title="${labels[a.v]}" aria-label="${labels[a.v]}"${isDis ? " disabled" : ""}>
        <span class="align-icon">${icons[a.v] || ""}</span>
      </button>`;
    }).join("")}
  </div>`;
}

// 开关按钮（加粗等布尔属性）
function switchField(id, checked, title) {
  // 统一 .switch 组件：视觉由 CSS `input:checked + .slider` 驱动（规范 §7），不维护 .on 类
  // title（第 3 参）：动态 tooltip 文案（如首页开关「开/关」随状态提示），随 renderStyleList 重渲染刷新
  return `<label class="switch"${title ? ` title="${title}"` : ""}>
    <input type="checkbox" id="${id}" ${checked ? "checked" : ""}>
    <span class="slider"></span>
  </label>`;
}

// 开关切换：更新 spec（视觉由 :checked 自动同步）
function bindSwitch(id, getter, setter) {
  const input = document.getElementById(id);
  if (!input) return;
  input.addEventListener("change", () => {
    setter(input.checked);
    const markDirtyAndRender = () => { markDirty(); renderPreview(); renderStyleList(); };
    markDirtyAndRender();
  });
}

// 后续段落样式下拉：候选 = 当前模板全部 text 样式 key（允许自引用，如 副标题→副标题）；
// 值存样式定义条目级 styles[].next（v1.31），非 defaults 格式值
function nextStyleSelect(id, value) {
  const keys = allStyleDefs().filter(d => styleCategory(d) === "text").map(d => d.key);
  const cur = value || "正文段落";
  let opts = keys.map(k => `<option value="${esc(k)}" ${k === cur ? "selected" : ""}>${esc(styleName(k))}</option>`).join("");
  if (cur && !keys.includes(cur)) opts = `<option value="${esc(cur)}" selected>${esc(cur)}</option>` + opts;
  return `<select id="${id}">${opts}</select>`;
}

// 检测字体是否本机安装（document.fonts.check，WebView2 支持）：false = 未装（预览将回退系统字体）
// 未装 → 字体字段下红字提示 + toast（同一字体只弹一次）；装/留空 → 清提示
function checkFontInstalled(fontName) {
  const warn = document.getElementById("font-warn");
  const name = String(fontName || "").trim();
  if (!warn) return;
  let installed = null;
  try {
    if (document.fonts && document.fonts.check && name) {
      installed = document.fonts.check(`36pt "${name}"`);
    }
  } catch (e) { installed = null; }
  if (installed === false) {
    warn.textContent = `本机未安装：${name}（预览将回退系统字体）`;
    warn.style.color = "#d33";
    if (warn.dataset.toasted !== name) {
      warn.dataset.toasted = name;
      toast(`本机未安装字体「${name}」，预览将回退系统字体`);
    }
  } else {
    warn.textContent = "";
    delete warn.dataset.toasted;
  }
}

// 构建元素样式编辑表单 HTML（手风琴展开区）
function buildElemFormHTML(key, spec) {
  // lead 只设「段首标题句」半边（方案 B 收口 2026-09-07）：标题后正文恒跟随「正文段落」样式，不再单独设置
  const def = styleDef(key);
  // 字符样式：仅 字体/字号/加粗（font 空 = 跟随段落；字号空 = 跟随段落）
  if (def && styleCategory(def) === "char") {
    return buildField("字体", fontSelect("f-font", spec.font || ""))
      + `<div id="font-warn" class="font-warn"></div>`
      + buildField("字号", sizeSelect("f-size", spec.size, true))
      + buildField("加粗", switchField("f-bold", spec.bold));
  }
  const isLeadMode = !!(def && def.mode === "lead");
  let html = isLeadMode ? `<div class="prop-group-title">段首标题句</div>` : "";
  html += buildField("字体", fontSelect("f-font", spec.font || ""));
  html += `<div id="font-warn" class="font-warn"></div>`;
  html += buildField("字号", sizeSelect("f-size", spec.size || 16));
  html += buildField("对齐", alignButtons("f-align", spec.align || "justify"));
  const lineBase = lineBasePt(spec);
  // 行距单位：lineMult（行倍数数据）强制「行」；固定磅样式按记忆单位（spec.lineUnit，默认磅）
  const lineUnit = spec.lineMult != null ? "line" : (spec.lineUnit || "pt");
  // 行距显示值：lineMult 样式按 倍数×行基准 折算成磅（line 可能为 null，直接 28 兜底会显示错倍数的行数）
  const lineVal = spec.lineMult != null ? Math.round(spec.lineMult * lineBase * 10) / 10 : (spec.line || 28);
  html += buildField("行距", spacingField("f-line", lineVal, lineBase, { min: 12, max: 120, lineMin: 0.5, unit: lineUnit }));
  html += buildField("段前", spacingField("f-before", spec.spaceBefore || 0, parseFloat(spec.line) || 28, { min: 0, max: 200, lineMin: 0, unit: spec.spaceBeforeUnit || "pt" }));
  html += buildField("段后", spacingField("f-after", spec.spaceAfter || 0, parseFloat(spec.line) || 28, { min: 0, max: 200, lineMin: 0, unit: spec.spaceAfterUnit || "pt" }));
  html += buildField("加粗", switchField("f-bold", spec.bold));
  html += buildField("首行缩进", `<input type="number" id="f-indent" min="0" max="8" step="0.5" value="${spec.indentChars || 0}"><span class="unit-suffix">字符</span>`);
  if (def && styleCategory(def) === "text") {
    html += buildField("后续段落", nextStyleSelect("f-next", def.next || ""));
  }
  return html;
}

// 预览渲染防抖（2026-09-02 排查）：行距/段前段后/缩进连续输入（input 连发、spinner 连点）合并为一次渲染。
// renderPreview 全量重建多页 .paper 且含落款 Canvas 测量，高频触发会卡顿（观感近似"不联动"）。
// 值读取发生在回调执行时（写值先行、schedule 只排渲染），无过期快照问题。
let _previewTimer = null;
const PREVIEW_DEBOUNCE_MS = 150;
function schedulePreviewRender() {
  if (_previewTimer) clearTimeout(_previewTimer);
  _previewTimer = setTimeout(() => { _previewTimer = null; renderPreview(); }, PREVIEW_DEBOUNCE_MS);
}

// 绑定元素样式编辑表单事件
function bindElemForm(key, spec) {
  const markDirtyAndRender = () => { markDirty(); renderPreview(); renderStyleList(); };
  // 输入过程中（input 事件）只刷新预览，不重建样式列表 —— 重建会销毁正在编辑的输入框，导致值被重置
  const previewOnly = () => { markDirty(); renderPreview(); };
  const on = (id, fn) => {
    const el = document.getElementById(id);
    if (el) { el.addEventListener("input", fn); el.addEventListener("change", fn); }
  };
  // 字符样式：字体/字号/加粗改动 → 全量刷新（摘要显示"跟随段落/加粗"变化）
  if (styleCategory(styleDef(key)) === "char") {
    on("f-font", e => { spec.font = e.target.value; checkFontInstalled(spec.font); markDirtyAndRender(); });
    // 字号空 = 跟随段落（不写字号）；有值才存数字
    on("f-size", e => { const v = e.target.value; spec.size = (v === "" ? "" : parseFloat(v) || 16); markDirtyAndRender(); });
    bindSwitch("f-bold", () => spec.bold, v => { spec.bold = v; markDirtyAndRender(); });
    return;
  }

  on("f-font", e => { spec.font = e.target.value; checkFontInstalled(spec.font); });
  on("f-size", e => { spec.size = parseFloat(e.target.value) || 16; });
  bindSwitch("f-bold", () => spec.bold, v => { spec.bold = v; });
  // 后续段落样式：写回样式定义条目级 next（v1.31）
  const def = styleDef(key);
  if (def && styleCategory(def) === "text") {
    on("f-next", e => { def.next = e.target.value; markDirtyAndRender(); });
  }
  // lead 样式无额外字段绑定：标题后正文已废弃（方案 B 收口），正文恒跟随「正文段落」样式

  // 对齐按钮组
  const alignBox = document.getElementById("f-align");
  if (alignBox) {
    alignBox.querySelectorAll(".align-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        spec.align = btn.dataset.align;
        alignBox.querySelectorAll(".align-btn").forEach(b => b.classList.toggle("active", b === btn));
        markDirtyAndRender();
      });
    });
  }

  // 行距：单位 磅/行（1行 = 1.2×字号）；lineMult 存在时优先按倍数显示（Word 单倍 lineRule=auto → 1.0 行）
  const lineInput = document.getElementById("f-line");
  const lineUnit = document.getElementById("f-line-unit");
  if (lineInput && lineUnit) {
    const lineBase = () => lineBasePt(spec);
    const syncLine = () => {
      // lineMult 存在（行单位数据）→ 初始渲染单位置"行"，显示倍数（1.0 行）；line 固定值 → 可按磅/行显示
      const hasMult = spec.lineMult != null;
      if (hasMult && lineUnit.value !== "line") { lineUnit.value = "line"; applySpacingStep(lineInput, "line", lineBase(), 12, 120, 0.5); }
      if (lineUnit.value === "line") {
        lineInput.value = hasMult ? lineMultTxt(spec.lineMult) : ((spec.line || 28) / lineBase()).toFixed(2).replace(/\.?0+$/, "");
      } else {
        lineInput.value = spec.line || 28;
      }
    };
    const commitLine = (list) => {
      const v = parseFloat(lineInput.value) || 0;
      if (spec.lineMult != null && lineUnit.value === "line") {
        // 行单位下改数值 = 改倍数（保留 lineMult，不解锁；解锁只发生在切单位）
        spec.lineMult = Math.round(v * 100) / 100;
      } else {
        // 磅单位 / 已解锁：写固定磅值（行单位时按行基准换算成磅）
        spec.line = lineUnit.value === "line" ? Math.round(v * lineBase() * 10) / 10 : Math.round(v);
      }
      markDirty(); schedulePreviewRender();
      if (list) renderStyleList();
    };
    lineInput.addEventListener("input", () => commitLine(false));
    lineUnit.addEventListener("change", () => {
      // 用户手动切单位 → 解锁：lineMult（行倍数）转 line（磅 = 当前倍数 × 行基准），单位恢复自由
      let unlocked = false;
      if (spec.lineMult != null) {
        const v = parseFloat(lineInput.value) || 1;
        spec.line = Math.round(v * lineBase() * 10) / 10;
        delete spec.lineMult;
        unlocked = true;
      }
      // 单位选择持久化进样式数据（spec.lineUnit，随模板落盘）：重建样式列表/重新打开文档时单位不回弹「磅」（2026-09-02 补）
      spec.lineUnit = lineUnit.value;
      // 切单位不重建样式列表：重建会让 .style-edit 表单里的单位 select 回到默认「磅」，
      // 刚切的「行」被回滚 → 无法在行单位下改数值，预览看起来"没联动"（2026-09-02 排查）。
      // 纯显示换算（磅↔行，数值不变）不改数据值（避免往返舍入漂移与误重渲预览）；
      // 但单位记忆需落盘 → markDirty 触发一次保存（不动数值不重渲）。
      syncLine(); applySpacingStep(lineInput, lineUnit.value, lineBase(), 12, 120, 0.5);
      markDirty();
      if (unlocked) commitLine(false);
    });
    applySpacingStep(lineInput, lineUnit.value, lineBase(), 12, 120, 0.5);
    syncLine();
  }

  // 段前段后：输入框按当前单位
  const beforeInput = document.getElementById("f-before");
  const afterInput = document.getElementById("f-after");
  const beforeUnit = document.getElementById("f-before-unit");
  const afterUnit = document.getElementById("f-after-unit");
  const baBase = () => parseFloat(spec.line) || 28; // 段前段后"行"按当前行距换算
  if (beforeInput && beforeUnit) {
    const syncDisplay = () => {
      if (beforeUnit.value === "line") beforeInput.value = ((spec.spaceBefore || 0) / baBase()).toFixed(2).replace(/\.?0+$/, "");
      else beforeInput.value = spec.spaceBefore || 0;
      if (afterUnit.value === "line") afterInput.value = ((spec.spaceAfter || 0) / baBase()).toFixed(2).replace(/\.?0+$/, "");
      else afterInput.value = spec.spaceAfter || 0;
    };
    const commit = (list) => {
      const b = parseFloat(beforeInput.value) || 0;
      const a = parseFloat(afterInput.value) || 0;
      // 磅为单位时按整数磅四舍五入，确保最小调节单位 1 磅
      spec.spaceBefore = beforeUnit.value === "line" ? Math.round(b * baBase() * 10) / 10 : Math.round(b);
      spec.spaceAfter = afterUnit.value === "line" ? Math.round(a * baBase() * 10) / 10 : Math.round(a);
      markDirty(); schedulePreviewRender();
      if (list) renderStyleList();
    };
    beforeInput.addEventListener("input", () => commit(false));
    afterInput.addEventListener("input", () => commit(false));
    // 切单位仅重算显示 + 持久化单位选择（spec.spaceBeforeUnit/spaceAfterUnit，随模板落盘，
    // 重建列表/重新打开不回弹「磅」）；不 commit 数值（显示换算无数据变化，commit 反而引入
    // 往返舍入漂移如 20磅→0.71行→19.9磅，且不必为纯换算重渲预览）
    const persistUnit = () => { spec.spaceBeforeUnit = beforeUnit.value; spec.spaceAfterUnit = afterUnit.value; markDirty(); };
    const syncUnitRange = () => { applySpacingStep(beforeInput, beforeUnit.value, baBase(), 0, 200); applySpacingStep(afterInput, afterUnit.value, baBase(), 0, 200); };
    beforeUnit.addEventListener("change", () => { persistUnit(); syncDisplay(); syncUnitRange(); });
    afterUnit.addEventListener("change", () => { persistUnit(); syncDisplay(); syncUnitRange(); });
    syncUnitRange();
    syncDisplay();
  }

  // 首行缩进：写回 spec.indentChars（原缺失导致值永不保存）；输入中只刷预览，
  // 上下箭头/连续输入不重建列表，避免输入框被销毁重置；失焦后再刷新列表摘要
  const indentInput = document.getElementById("f-indent");
  if (indentInput) {
    const commitIndent = (list) => {
      spec.indentChars = parseFloat(indentInput.value) || 0;
      markDirty(); schedulePreviewRender();
      if (list) renderStyleList();
    };
    indentInput.addEventListener("input", () => commitIndent(false));
    indentInput.addEventListener("change", () => commitIndent(true));
  }
  // select（字体/字号）选择完成 → 写回 + 刷新列表摘要
  ["f-font", "f-size"]
    .forEach(id => { const el = document.getElementById(id); if (el) { el.addEventListener("input", markDirtyAndRender); el.addEventListener("change", markDirtyAndRender); } });
}

// 原独立「页面设置」卡片已并入样式面板「页面样式」组（点「页面」行展开编辑，见 buildPageFormHTML）

function selectElem(key, paraEl) {
  // 再次点击同一段落 → 取消选中（以 .selected 类判定，不依赖 DOM 引用）
  if (state.selectedParaIdx !== null && paraEl.classList.contains("selected")) {
    state.selectedParaKey = null;
    state.selectedParaIdx = null;
    state.selectedElem = null;      // 取消选中 → 右侧编辑表单一并收起（组的折叠态不回滚，避免面板跳动）
    paraEl.classList.remove("selected");
    renderStyleList();
    return;
  }
  // 选中当前段落（仅该段，按 sampleBlocks 序号精确记录），并高亮其**生效**样式行
  // 段落可能被覆盖（paraOverrides），此时 data-key 仍是源样式，须按段落原文查实际生效样式
  const effKey = effKeyOf(paraTextOf(paraEl)) || key;
  const idx = parseInt(paraEl.dataset.paraIdx || "-1", 10);
  state.selectedParaKey = effKey;
  state.selectedParaIdx = (idx >= 0) ? idx : null;
  document.querySelectorAll(".para.selected").forEach(el => el.classList.remove("selected"));
  paraEl.classList.add("selected");
  paraEl.dataset.paraSelected = "1";
  // 联动（D1+D2）：展开生效样式所在分组 → 展开该样式的编辑表单 → 滚动可见
  // 引导段（lead）样式行就在面板里，联动指回它自身（2026-09-01 删字符样式「段首标题句」后无第二入口）
  const linkKey = syncPanelToPara(effKey);
  renderStyleList();
  scrollStyleRowIntoView(linkKey);
}

// ═══════════════════════════════════════════
// 登录功能已移除（本版只处理本地文档上传，不需要在线文档 / 扫码登录）。
// 以下保留空实现，避免历史任务回显等路径引用报错。
// ═══════════════════════════════════════════
let _loginOk = false;      // 兼容旧引用：恒 false（无登录）
let _convEchoTask = false; // 转换页当前是否在回显任务来源（openTask / restoreConvProcess）而非新建任务态

async function refreshLoginStatus(prefetched) {
  return { ok: false, reason: "已禁用登录" };
}

function applyOnlineLinkGate(ok) {
  _loginOk = !!ok;
  updateInputSummary();
  applyInheritDisabled();
}

function onlineLinkShown() {
  return false;
}

function renderLoginEntry(st) {
  _loginOk = !!(st && st.ok);
  updateInputSummary();
  applyInheritDisabled();
  if (typeof renderSidebarCollapse === "function") renderSidebarCollapse();
}

async function onNavLogin() {
  // 无登录入口，不做任何事
}

function loginMenuHeader(st) {
  return "";
}

async function openAccountMenu(anchor, st) {
  // 无账号菜单，不做任何事
}

async function doLogout() {
  // 无登录，不做任何事
}

// ── 转换任务（每次成功生成自动记录，左侧「文档排版」下展示，可点击恢复重转）──
// 存储（2026-08-22）：服务端落盘 %APPDATA%/doc-audit/Tasks/tasks.json，数量不设上限；
// 过期清理（>90 天）启动时自动执行（2026-09-01 起不再询问，见 initTasks）。localStorage 仅离线兜底。
const TASKS_KEY = "klunl_tasks";
// ── 格式处理选项（升级自「继承源文档格式」白名单）：全局偏好（独立持久化，与任务历史快照分离）
// - 十一项勾选，分三组：段落结构 lead / 保留原文文字格式 8 项（bold italic underline strike color bgColor charSpacing charScale）/ 内容处理 review + namecheck
// - 默认开 lead + bold + review；其余 7 项文字属性（含颜色/底色，规范性风险）+ namecheck（需先有名单）默认关，由用户自行开启
// - 用户改 checkbox 后实时落盘，作为下次进入的默认值；旧数据兼容：仅保留白名单键（旧 emphasisFont 对象结构取 checks）
const INHERIT_PREFS_KEY = "klunl_inherit_prefs";
const INHERIT_PREFS_V2_KEY = "klunl_inherit_prefs_v2";   // review（内容校对）引入后：老存档首次升级补一次默认勾选，此后尊重用户存档
const DEFAULT_INHERIT = ["lead", "bold", "review"];   // 首次默认：lead/bold/review 开；其余 6 项文字属性（含颜色/底色）默认关
// 允许项 = 后端继承白名单 8 项 + lead（段落结构）+ review（内容校对）+ namecheck（人员姓名校对）
// ⚠ namecheck **不进 DEFAULT_INHERIT、也不进 v2 迁移补勾**：它比 review 多一个前提（必须先有名单），
//   默认开会让没建名单的用户一进来就撞见"未执行"提示。只走正常偏好持久化。
const INHERIT_ALLOW = ["lead", "review", "namecheck", "bold", "italic", "underline", "strike", "color", "bgColor", "charSpacing", "charScale"];
// 文字属性 8 项 = 提交给后端的 inheritAttrs（后端 convert.js 继承白名单同名键）；
// lead / review 不是文字属性，走独立布尔开关 recognizeLeadParagraph / reviewEnabled
const INHERIT_ATTR_KEYS = ["bold", "italic", "underline", "strike", "color", "bgColor", "charSpacing", "charScale"];
// 旧任务兼容：老 UI「保留源字符缩放与间距」一个勾同时下发 charSpacing + charScale，
// 恢复时还原成两项勾选，忠实反映当时后端实际应用的属性
function expandLegacyInheritAttrs(list) {
  const out = [];
  for (const x of list) {
    if (x === "charScale") { out.push("charSpacing", "charScale"); continue; }
    out.push(x);
  }
  return out.filter(x => INHERIT_ALLOW.includes(x));
}
function loadInheritPrefs() {
  let list = null;
  try {
    const v = JSON.parse(localStorage.getItem(INHERIT_PREFS_KEY) || "null");
    if (Array.isArray(v)) list = v.filter(x => INHERIT_ALLOW.includes(x));
    else if (v && typeof v === "object" && Array.isArray(v.checks)) list = v.checks.filter(x => INHERIT_ALLOW.includes(x));
  } catch (e) {}
  if (!list) {
    // 无任何存档（全新用户首启）：按出厂默认返回，并立即打 v2 标——否则用户随后首次改动
    // 会被下面「老存档迁移」误判（无标记 → 补勾），导致取消「内容校对」后刷新又弹回勾选
    try { localStorage.setItem(INHERIT_PREFS_V2_KEY, "1"); } catch (e) {}
    return DEFAULT_INHERIT.slice();
  }
  // v2 迁移：老存档（内容校对功能加入前保存）补上默认勾选的「内容校对」，保证升级后默认行为与出厂一致
  if (!localStorage.getItem(INHERIT_PREFS_V2_KEY)) {
    if (!list.includes("review")) list = list.concat("review");
    try { localStorage.setItem(INHERIT_PREFS_V2_KEY, "1"); } catch (e) {}
  }
  return list;
}
function saveInheritPrefs(list) {
  try { localStorage.setItem(INHERIT_PREFS_KEY, JSON.stringify(list)); } catch (e) { /* 忽略 */ }
}
/** #inherit-grid 内**格式处理选项本身**的勾选框（保序）。
 *  ⚠ 必须用 .opt-check 限定：人员姓名校对的名单选择区也在 #inherit-grid 内，
 *    那里的 checkbox 是「本次用哪几份名单」，不是格式处理选项——不排除就会被当成选项一起存取。 */
function inheritOptInputs() {
  const grid = document.getElementById("inherit-grid");
  if (!grid) return [];
  return [...grid.querySelectorAll('.inherit-opt > label.opt-check > input[type="checkbox"]')];
}

// 把偏好应用到 #inherit-grid 的 checkbox；returns 收集到的 value 数组（保序、按当前 DOM 顺序）
function applyInheritPrefs(list) {
  const inputs = inheritOptInputs();
  if (!inputs.length) return list.slice();
  const set = new Set(list);
  inputs.forEach(inp => { inp.checked = set.has(inp.value); });
  // 返回"当前 DOM 的勾选顺序"，与传入门控顺序无关，保序便于展示 & 与后端字段对齐
  return inputs.filter(i => i.checked).map(i => i.value);
}
function setupInheritGrid() {
  const grid = document.getElementById("inherit-grid");
  if (!grid) return;
  // 1) 应用持久化偏好（含首次默认）
  applyInheritPrefs(loadInheritPrefs());
  // 2) 监听 checkbox 变化 → 立即把当前全量勾选写回 localStorage
  //    ⚠ 名单选择区的 change 会冒泡到这里（它嵌套在 .inherit-opt 内）：取值必须用 inheritOptInputs()，
  //       否则名单 id 会被当成格式选项写进偏好存档。
  grid.addEventListener("change", ev => {
    const t = ev.target;
    // 名单选择区的改动：只存名单选择。它不是格式选项，不能写进格式处理偏好（否则会被白名单校验滤掉）
    if (t && t.closest && t.closest("#roster-pick")) {
      saveRosterSelection(selectedRosterIds());
      updateRosterPickHint();
      return;
    }
    // 勾/取消「人员姓名校对」→ 同步名单选择区的显隐
    if (t && t.value === "namecheck") syncRosterPickVisible();
    saveInheritPrefs(inheritOptInputs().filter(i => i.checked).map(i => i.value));
  });
  // 3) 按当前输入源类型灰置不可达属性（仅 bold 受源格式能力限制；lead 走形态识别，任何源可用）
  applyInheritDisabled();
  // 4) 名单选择区：按当前勾选状态定显隐（清单**懒加载**——没勾「人员姓名校对」时不请求名单接口）
  syncRosterPickVisible();
}

// ── 格式属性 × 输入源 可达矩阵（与后端能力一一对应；置灰 = 该源类型无法提供）
// 在线（link）：launcher/blocks.go DocRun 只有 bold/italic/underline/strike/color/bgColor 六个字段，
//   无字符间距/缩放 → 这两项在在线源一律置灰（2026-09-05 修正：此前误标为可达，勾了也不生效）
// 本地 docx：<w:b>/<w:i>/<w:u>/<w:strike>/<w:color>/<w:shd>/<w:spacing>/<w:w> 八项全可达
// 本地 doc/md：经 markdown 中转，只有 **加粗** 与 *斜体*
const LINK_ATTRS = ["bold", "italic", "underline", "strike", "color", "bgColor"];
const DOCX_ATTRS = ["bold", "italic", "underline", "strike", "color", "bgColor", "charSpacing", "charScale"];
const INHERIT_KIND_SUPPORT = {
  link: LINK_ATTRS,
  docx: DOCX_ATTRS,
  doc:  ["bold", "italic"],
  md:   ["bold", "italic"],
  other: [],                                                             // txt/json：无 run 格式
};
// 判定当前输入源类型：看本地文件扩展名；无输入 → null（全启用）。
function currentInputKind() {
  if (convertLocalFiles.length) {
    const name = (convertLocalFiles[0].name || "").toLowerCase();
    const ext = name.includes(".") ? name.split(".").pop() : "";
    if (ext === "docx") return "docx";
    if (ext === "doc") return "doc";
    if (ext === "md") return "md";
    return "other";
  }
  return null;   // 未选择输入源 → 不灰任何项
}
// 应用灰置：不可达属性 checkbox disabled + 卡片降级；可达/未选源保持可勾；lead（形态识别）、review（内容校对）、namecheck（姓名校对）永不置灰
function applyInheritDisabled() {
  const grid = document.getElementById("inherit-grid");
  if (!grid) return;
  const kind = currentInputKind();
  const support = kind ? INHERIT_KIND_SUPPORT[kind] : null;
  grid.querySelectorAll(".inherit-opt").forEach(opt => {
    const inp = opt.querySelector('input[type="checkbox"]');
    if (!inp) return;
    const disabled = !["lead", "review", "namecheck"].includes(inp.value) && Array.isArray(support) && !support.includes(inp.value);
    inp.disabled = disabled;
    opt.classList.toggle("disabled", disabled);
    // 不覆盖 opt.title：HTML 已写好 tooltip 全文；disabled 视觉降级 + cursor:not-allowed 已足够透明传达
  });
}

// ── 人员姓名校对的名单选择区（第 3 步「内容处理」内就地展开，不弹窗） ──
// 名单本身在设置页维护；这里只决定**本次转换用哪几份**。选择结果持久化（与格式处理偏好同范式，
// 但**分开存**：它不是格式选项，混在一个键里会让 inheritOptInputs 的白名单校验把它滤掉）。
const ROSTER_PICK_KEY = "klunl_roster_pick";
let rosterCache = null;        // 名单清单（/api/rosters 的 rosters 数组）；null = 读取失败/未加载
let rosterPinnedPick = null;   // 一次性渲染覆盖：历史任务回显用（非 null 时按它勾选；不写 localStorage）

function loadRosterSelection() {
  try {
    const a = JSON.parse(localStorage.getItem(ROSTER_PICK_KEY) || "null");
    return Array.isArray(a) ? a.map(String) : null;   // null = 无存档（首次 → 默认全选）
  } catch (e) { return null; }
}
function saveRosterSelection(ids) {
  try { localStorage.setItem(ROSTER_PICK_KEY, JSON.stringify(ids)); } catch (e) { /* 忽略 */ }
}
/** 当前勾选的名单 id（保序、按 DOM 顺序） */
function selectedRosterIds() {
  const box = document.getElementById("roster-pick-list");
  if (!box) return [];
  return [...box.querySelectorAll('input[type="checkbox"]:checked')].map(i => i.value);
}

/** 勾/取消「人员姓名校对」时同步名单选择区显隐（首次展开时补拉名单） */
function syncRosterPickVisible() {
  const wrap = document.getElementById("roster-pick");
  if (!wrap) return;
  const cb = document.querySelector('#inherit-grid input[value="namecheck"]');
  const on = !!(cb && cb.checked);
  wrap.style.display = on ? "" : "none";
  if (on && rosterCache === null) loadRosterPick();
}

/** 历史任务回显：把该任务当时勾的名单应用到选择区。
 *  ★ **不写 localStorage** —— 回显只用于展示，改了全局偏好就变成「点开历史任务，下次的默认值被悄悄改掉」。
 *  清单尚未渲染（启动期异步未回）→ 记为一次性渲染覆盖，等渲染时按它落。 */
function applyRosterSelection(ids) {
  const sel = (Array.isArray(ids) ? ids : []).map(String);
  const box = document.getElementById("roster-pick-list");
  if (box && box.children.length) {
    const set = new Set(sel);
    [...box.querySelectorAll('input[type="checkbox"]')].forEach(i => { i.checked = set.has(i.value); });
    updateRosterPickHint();
    return;
  }
  rosterPinnedPick = sel;
  loadRosterPick();
}

/** 拉名单清单并渲染选择区（设置页保存/删除名单后也调用，两处口径始终一致） */
async function loadRosterPick() {
  try {
    const r = await api("/api/rosters");
    rosterCache = (r && Array.isArray(r.rosters)) ? r.rosters : [];
  } catch (e) { rosterCache = null; }   // 读取失败 → 选择区给明确提示，不静默当成"没有名单"
  renderRosterPick();
}

/** 渲染名单选择区：每份名单一行 checkbox（value = 名单 id） */
function renderRosterPick() {
  const box = document.getElementById("roster-pick-list");
  const hint = document.getElementById("roster-pick-hint");
  if (!box) return;
  const list = rosterCache || [];
  if (!list.length) {
    box.innerHTML = "";
    if (hint) {
      hint.textContent = rosterCache === null
        ? "名单读取失败，请到「设置 → 人员名单」检查后重试"
        : "还没有名单，请先到「设置 → 人员名单」新建";
      hint.classList.add("warn");
    }
    return;
  }
  const saved = loadRosterSelection();
  const has = new Set(list.map(r => String(r.id)));
  // 勾选来源优先级：历史任务回显（一次性）> 用户存档 > 首次默认全选；失效 id 一律自然淘汰
  const pinned = rosterPinnedPick;
  rosterPinnedPick = null;
  const picked = pinned
    ? pinned.filter(id => has.has(id))
    : (saved ? saved.filter(id => has.has(id)) : list.map(r => String(r.id)));
  box.innerHTML = list.map(r => {
    const id = esc(String(r.id));
    const n = (r.members || []).length;
    return `<label class="roster-pick-item">`
      + `<input type="checkbox" value="${id}"${picked.includes(String(r.id)) ? " checked" : ""}>`
      + `<span>${esc(String(r.name || ""))}</span>`
      + `<span class="rp-count">${n} 人</span>`
      + `</label>`;
  }).join("");
  updateRosterPickHint();
}

/** 选择区下方提示：只在「一份都没勾」时给出说明；有勾选则不占位 */
function updateRosterPickHint() {
  const hint = document.getElementById("roster-pick-hint");
  if (!hint) return;
  const list = rosterCache || [];
  if (!list.length) return;   // 空清单/读取失败的提示由 renderRosterPick 写
  const picked = new Set(selectedRosterIds());
  if (list.some(r => picked.has(String(r.id)))) {
    // 有勾选时不留提示：每行名单自带人数、勾选态本身就表达得清楚。
    // 这里必须显式清空，否则会残留上一次「未选择任何名单」的红字。
    hint.classList.remove("warn");
    hint.textContent = "";
    return;
  }
  hint.textContent = "未选择任何名单：本次不会执行姓名校对";
  hint.classList.add("warn");
}

// 历史任务存储（2026-08-22 起服务端落盘 %APPDATA%/doc-audit/Tasks/tasks.json）：
// 内存缓存为同步真源（沿用既有同步调用点），写入异步 POST 服务端；localStorage 仅作离线兜底。
let tasksCache = null;

async function initTasks() {
  let fromServer = false;
  try {
    const r = await api("/api/tasks");
    if (r && Array.isArray(r.tasks)) {
      if (r.tasks.length) { tasksCache = r.tasks; fromServer = true; }
      else tasksCache = [];
      // 一次性迁移：历史总表里内联的过程区 HTML（本版本起归档到 Tasks\<id>\）→ 首次加载即回写瘦身。
      // 阈值与后端 PROC_HTML_INLINE_MAX 同量级粗判，只用于决定"要不要多写一次"，不参与任何判定。
      if (fromServer && tasksCache.some(t => t && t.process && typeof t.process.procHtml === "string" && t.process.procHtml.length > 4096)) {
        saveTasks(tasksCache);
      }
    }
  } catch (e) { /* 服务不可用 → 走本地兜底 */ }
  if (!fromServer) {
    // 兜底：迁移旧 localStorage 数据到服务端（老版本历史不丢）
    try {
      const a = JSON.parse(localStorage.getItem(TASKS_KEY) || "[]");
      if (Array.isArray(a) && a.length) {
        tasksCache = a.filter(t => t && typeof t.url === "string");
        saveTasks(tasksCache);
      }
    } catch (e) { /* ignore */ }
    if (!tasksCache) tasksCache = [];
  }
  await cleanupExpiredTasks();
}

// 启动过期清理（2026-09-01 用户定调）：>90 天未使用的历史任务自动清理，不再弹确认
const TASK_EXPIRE_DAYS = 90;

async function cleanupExpiredTasks() {
  const tasks = loadTasks();
  if (!tasks.length) return;
  const cutoff = Date.now() - TASK_EXPIRE_DAYS * 24 * 3600 * 1000;
  // 2026-09-01 用户定调：无 createdAt 的老数据也一律清理，不留例外
  const expired = tasks.filter(t => !t.createdAt || t.createdAt < cutoff);
  if (!expired.length) return;
  const ids = new Set(expired.map(t => t.id));
  saveTasks(tasks.filter(t => !ids.has(t.id)));
  // 过期清理同步联动删产物目录（同 deleteTask）
  expired.forEach(t => purgeTaskOutputs(t.id));
}

function loadTasks() {
  if (!tasksCache) {
    try {
      const a = JSON.parse(localStorage.getItem(TASKS_KEY) || "[]");
      tasksCache = Array.isArray(a) ? a.filter(t => t && typeof t.url === "string") : [];
    } catch (e) { tasksCache = []; }
  }
  return tasksCache.filter(t => t && typeof t.url === "string");
}

function saveTasks(tasks) {
  tasksCache = (Array.isArray(tasks) ? tasks : []).filter(t => t && typeof t.url === "string");
  try { localStorage.setItem(TASKS_KEY, JSON.stringify(slimTasksForCache(tasksCache))); } catch (e) { /* 忽略 */ }
  // 异步落盘服务端（不阻塞 UI）；失败静默保留本地兜底。
  // 与其他任务写共用串行队列：并发 POST（转换完成 × 删除/重命名任务等）会丢更新
  // （后完成的写覆盖先完成的写），串行后队尾写入的必是 tasksCache 最新值。
  enqueueSetWrite(() => fetch("/api/tasks", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ tasks: tasksCache }),
  }).then(r => (r && r.ok) ? r.json() : null).then(j => {
    // 服务端把过程区 HTML 抽离成 Tasks\<id>\过程记录-*.html 后，回传 id→rel，
    // 前端同步换成引用：① 后续全量保存不再重复写同一份快照；② 内存不再驻留 130KB 级 HTML。
    if (j && Array.isArray(j.extracted) && j.extracted.length) applyExtractedProcHtml(j.extracted);
  }).catch(() => {}));
}

// 历史任务过程区渲染：过程 HTML + 兼容旧快照（statsHtml 曾独立于过程区）+ 重绑折叠交互。
// 归档回显与内联回显走同一入口，保证两种来源渲染结果完全一致。
function fillHistoryProc(procEl, task, html) {
  procEl.innerHTML = html || "";
  // 兼容旧历史任务：转换统计曾独立存于 statsHtml（旧快照过程区不含统计），合并回过程区尾部
  if (task.process.statsHtml && !procEl.querySelector(".conv-stats")) {
    procEl.insertAdjacentHTML("beforeend", '<div class="conv-stats">' + task.process.statsHtml + "</div>");
  }
  rebindReasonToggle(procEl);   // 恢复思考流折叠/展开交互（innerHTML 序列化会丢失事件绑定）
}

// 过程记录归档版回显：从 Tasks\<id>\过程记录-*.html 按需读回（历史任务点开时才加载，不占启动与内存）
async function loadArchivedProcHtml(task, procEl) {
  const rel = task.process.procHtmlFile && task.process.procHtmlFile.rel;
  if (!rel) return;
  procEl.innerHTML = '<div class="conv-line">载入过程记录…</div>';
  try {
    const r = await fetch("/api/task-html?f=" + encodeURIComponent(rel));
    const j = r && r.ok ? await r.json() : null;
    if (!j || typeof j.html !== "string") throw new Error("empty");
    fillHistoryProc(procEl, task, j.html);
  } catch (e) {
    procEl.innerHTML = '<div class="conv-line">过程记录已随任务清理，无法回显</div>';
  }
}

// localStorage 有 5MB 硬限（超出即静默失败）：过程记录已随任务归档，本地只存引用不存全文
function slimTasksForCache(tasks) {
  return tasks.map(t => (t && t.process && typeof t.process.procHtml === "string")
    ? { ...t, process: { ...t.process, procHtml: "" } }
    : t);
}

// 服务端抽离结果回填内存：procHtml → procHtmlFile 引用（不再次触发 POST，纯内存+本地兜底更新）
function applyExtractedProcHtml(list) {
  let changed = false;
  for (const e of list) {
    if (!e || !e.id || !e.rel) continue;
    const t = tasksCache.find(x => x && x.id === e.id);
    if (!t || !t.process) continue;
    if (typeof t.process.procHtml === "string") { delete t.process.procHtml; changed = true; }
    t.process.procHtmlFile = { rel: e.rel, bytes: e.bytes };
  }
  if (changed) { try { localStorage.setItem(TASKS_KEY, JSON.stringify(slimTasksForCache(tasksCache))); } catch (e) { /* 忽略 */ } }
}
// 记录任务：按「链接」去重——同一文档只保留一条任务（更新并移到顶部，
// 模板集/勾选记录最近一次使用的；换模板集转换也更新到该任务）。
function upsertTask({ taskId, title, url, inputType, files, tsFile, tsName, checked, checkedNames, inheritAttrs, recognizeLeadParagraph, reviewEnabled, nameCheck, nameRosters, process, telemetry }) {
  const tasks = loadTasks();
  const idx = tasks.findIndex(t => t.url === url);
  const prev = idx >= 0 ? tasks[idx] : null;
  const task = {
    id: prev ? prev.id : (taskId || "t" + Date.now()),
    title: title || url,
    url,
    inputType: inputType || (prev && prev.inputType) || "link",
    files: files || (prev && prev.files) || [],
    tsFile: tsFile || "",
    tsName: tsName || "",
    checked: checked || [],
    // 勾选模板的**名字**（与 checked 同步记录）：模板 key 会因「删除重建 / 重新导入」而变，
    // 只记 key 的历史任务在模板重建后重转会整体失效；名字是用户可见身份，可用来把旧勾选认回（2026-09-13）
    checkedNames: checkedNames !== undefined ? checkedNames : ((prev && prev.checkedNames) || []),
    // 历史任务里同时记录"当时勾选的格式处理选项"；旧任务无字段时回退 localStorage 偏好（不含 lead，lead 由 recognizeLeadParagraph 单独管）
    inheritAttrs: inheritAttrs !== undefined ? inheritAttrs : (prev && prev.inheritAttrs) || loadInheritPrefs().filter(x => INHERIT_ATTR_KEYS.includes(x)),
    recognizeLeadParagraph: recognizeLeadParagraph !== undefined
      ? recognizeLeadParagraph
      : (prev && prev.recognizeLeadParagraph !== undefined ? prev.recognizeLeadParagraph : true),
    // 内容校对开关：新任务记录布尔；无字段（老任务/直传）→ 默认开（与出厂「默认勾选」及老版本总是校对的行为一致）
    reviewEnabled: reviewEnabled !== undefined
      ? reviewEnabled
      : (prev && prev.reviewEnabled !== undefined ? prev.reviewEnabled : true),
    // 人员姓名校对开关 + 本次勾选的名单 id（重跑还原用）。
    // 开关：旧任务无字段 → 默认**关**（该能力是后加的，老任务当时没有这个选项，不该凭空补勾）。
    // 名单 id：名单被删后 id 失效，还原时由 renderRosterPick 按现存名单自然淘汰。
    nameCheck: nameCheck !== undefined
      ? nameCheck
      : (prev && prev.nameCheck !== undefined ? prev.nameCheck : false),
    nameRosters: nameRosters !== undefined
      ? nameRosters
      : ((prev && prev.nameRosters) || []),
    createdAt: Date.now(),
    // 生成成功时写入最新过程快照；重命名等非生成更新则沿用旧快照（process 省略即保留）
    process: process ? process : (prev ? prev.process : null),
    // 过程记录（源文件/提取结果/LLM 输入输出/统计/输出），服务端落盘可审计；非生成更新沿用旧值
    telemetry: telemetry ? telemetry : (prev ? prev.telemetry : null),
  };
  if (idx >= 0) tasks.splice(idx, 1);
  tasks.unshift(task);
  saveTasks(tasks);   // 数量不设上限（2026-08-22 放开）；过期清理启动时自动执行（2026-09-01 起）
  renderTasks();
  return task;
}
// 任务列表相对时间（2026-09-02 定稿）：刚刚 / N分钟前 / N小时前 / 昨天 / N天前，
// 超 1 天一概天数化，永不回落绝对日期。列表随 renderTasks 刷新，无需定时器。
function fmtTaskTime(ts) {
  const d = new Date(ts);
  if (isNaN(d.getTime())) return "";
  const diff = Date.now() - ts;
  if (diff < 60e3) return "刚刚";
  if (diff < 3600e3) return `${Math.floor(diff / 60e3)}分钟前`;
  if (diff < 86400e3) return `${Math.floor(diff / 3600e3)}小时前`;
  const days = Math.floor(diff / 86400e3);
  return days < 2 ? "昨天" : `${days}天前`;
}

// ── 就地重命名：转换任务（双击任务行 → 行内改标题） ──
function beginInlineRenameTask(el, taskId) {
  // 收口旧编辑态：同一时刻只允许一个就地编辑行
  finishInlineEditing(true);
  const titleSpan = el.querySelector(".task-title");
  if (!titleSpan) return;
  const oldName = titleSpan.textContent;
  const input = document.createElement("input");
  input.className = "tree-rename-input";
  input.type = "text";
  input.value = oldName;
  titleSpan.replaceWith(input);
  el.classList.add("editing");
  _inlineEditing = true;
  input.focus();
  input.select();
  let done = false;
  const finish = (save) => {
    if (done) return;
    done = true;
    _inlineEditFinish = null;
    _inlineEditing = false;
    el.classList.remove("editing");
    const v = input.value.trim();
    if (!save || !v || v === oldName) {
      input.replaceWith(titleSpan);
      renderTasks();
      return;
    }
    const tasks = loadTasks();
    const t = tasks.find(x => x.id === taskId);
    if (t) {
      t.title = v;
      t.createdAt = Date.now();
      saveTasks(tasks);
    }
    renderTasks();
  };
  _inlineEditFinish = finish;
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); finish(true); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); finish(false); }
  });
  input.addEventListener("blur", () => finish(true));
}

let activeTaskId = null;

// 删除历史任务
// 删除任务 = 删记录 + 删产物目录（Tasks\<taskId>\，N7+方案A 同名联动，2026-08-27 拍板）。
// purge 失败不阻塞记录删除：目录残留交给 GC 后议。
function purgeTaskOutputs(taskId) {
  try { fetch("/api/tasks/purge", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: taskId }) }).catch(() => {}); } catch (e) { /* ignore */ }
}

function deleteTask(taskId) {
  const wasActive = activeTaskId === taskId;
  let tasks = loadTasks();
  tasks = tasks.filter(t => t.id !== taskId);
  saveTasks(tasks);
  purgeTaskOutputs(taskId);
  if (wasActive) {
    // 删除的是当前打开的任务：主显示区仍挂着该任务的过程/结果快照，
    // 复用「新建任务」清空逻辑回到新建态（清显示区 + 脱离任务流 + 隐藏右上角重跑，
    // 内部已置 activeTaskId=null 并重渲染两侧树）。进行中的转换不中断（并行模型不变）。
    newConvertTask();
    return;
  }
  renderTasks();
}

// 任务列表折叠：总数超过上限时默认只列最近 N 条（loadTasks 为新→旧序），尾部给展开/收起行。
// 展开态是会话级状态（重渲染不丢），但刷新/重启回到折叠态 —— 侧边栏是任务入口而非浏览面板。
const TASKS_FOLD_LIMIT = 8;
let tasksListExpanded = false;

// 渲染任务列表（左侧「文档排版」下）
function renderTasks() {
  const tree = document.getElementById("tasks-tree");
  if (!tree) return;
  const tasks = loadTasks();
  // 仅当「文档排版」导航激活时才高亮任务，避免与模板设计侧选中并存
  const convertActive = document.getElementById("nav-convert") &&
    document.getElementById("nav-convert").classList.contains("active");
  if (!tasks.length) {
    tree.innerHTML = `<div class="hint">暂无任务：转换成功后自动记录在这里，可点击重新转换</div>`;
    return;
  }
  const foldable = tasks.length > TASKS_FOLD_LIMIT;
  // 条数回落到上限以内 → 复位展开态，下次再超限时仍从「折叠」起步
  if (!foldable) tasksListExpanded = false;
  const shown = foldable && !tasksListExpanded ? tasks.slice(0, TASKS_FOLD_LIMIT) : tasks;
  const foldRow = foldable
    ? `<div class="tree-node"><div class="tree-more" role="button" tabindex="0"
         title="${tasksListExpanded ? "收起列表" : `展开查看全部 ${tasks.length} 个任务`}"
       >${tasksListExpanded ? "收起" : `展开全部（共 ${tasks.length} 个）`}</div></div>`
    : "";
  tree.innerHTML = `<div class="tree">${shown.map(t => {
    // 进行中的任务行只挂转圈（klUI `.kl-spinner`），不显示「转换中」文案与时间：
    // 状态源 = 内存 convTasks（running），任务完成/失败后由 renderTasks 刷新，转圈与时间互换。
    const running = convTasks.some(c => c.running && c.taskId === t.id);
    return `
    <div class="tree-node">
      <div class="tree-set task-item ${convertActive && t.id === activeTaskId ? "current" : ""}" data-task="${esc(t.id)}">
        <span class="task-title">${esc(t.title)}</span>
        ${running ? `<span class="kl-spinner task-spin" role="img" aria-label="转换进行中" title="转换进行中"></span>` : ""}
        ${running ? "" : `<span class="task-meta">${fmtTaskTime(t.createdAt)}</span>`}
        <span class="task-del" data-task="${esc(t.id)}" title="删除">删除</span>
      </div>
    </div>`;}).join("")}${foldRow}</div>`;
}

// 打开任务：填入链接 + 模板集/勾选 → 用户自行点击生成
async function openTask(task) {
  clearTemplateSelection(); // 切到文档排版区，清除「模板设计」选中（同时重渲染树以去除高亮）
  // 先切视图：从其他视图进入时 switchView 会触发 restoreConvProcess（恢复聚焦任务显示），
  // 随后本函数立即用历史快照覆盖显示区——因此 focusedRunId 必须在 switchView 之后置 null，
  // 否则会被恢复逻辑改回聚焦任务 runId → 进行中任务的流事件污染历史快照、右上角重跑对象错位。
  switchView("convert");
  snapshotConvState();   // 覆盖显示区前保住当前聚焦任务的最新进度（并行模型：进行中任务的快照在 convTasks）
  // 该任务最近一次运行（重转进行中/刚结束未回存）：存在则切回时聚焦该次运行，而非回显旧历史快照
  const liveRun = convTasks.filter(x => x.taskId === task.id).pop() || null;
  focusedRunId = null;   // 共享显示区脱离任务流：历史快照不被进行中任务的流事件写入
  _historyViewTask = null;
  // 载入历史任务：若上次生成成功，恢复已保存的生成过程快照（直接回显，不再提示"已载入"）；
  // 无快照的任务保持干净，给中性占位（不含任务标题）。
  const rb = document.getElementById("result-block");
  const procEl = document.getElementById("conv-process");
  const rl = document.getElementById("result-list");
  const blockInput = document.getElementById("block-input");
  const blockTemplate = document.getElementById("block-template");
  const blockInherit = document.getElementById("block-inherit");
  // 打开历史快照：三个卡片默认折叠（摘要见各 update*Summary），聚焦回显的过程/结果区
  if (blockInput) blockInput.classList.add("collapsed");
  if (blockTemplate) blockTemplate.classList.add("collapsed");
  if (blockInherit) blockInherit.classList.add("collapsed");
  if (rb) rb.style.display = "block";
  if (liveRun) {
    // 点开的是重转进行中/刚结束的任务：恢复该次运行的实时快照并重新聚焦其流。
    // 旧行为是回显 task.process（上一次成功转换的旧快照）→「转换中」徽标 + 旧结果同屏，
    // 且 focusedRunId=null 后该任务的流不再写显示区，切换期间的进展永远看不见。
    convState = liveRun;
    focusedRunId = liveRun.runId;   // 重新挂接：该任务的流事件继续写共享显示区
    if (procEl) {
      procEl.innerHTML = liveRun.procHtml || "";
      rebindReasonToggle(procEl);
    }
    if (rl) {
      rl.innerHTML = liveRun.resultHtml || "";
      upgradeResultItems(rl);
    }
    const resultPhaseLive = document.getElementById("result-phase");
    if (resultPhaseLive) resultPhaseLive.style.display = liveRun.resultHtml ? "" : "none";
  } else if (task.process && procEl) {
    // 过程记录已随任务归档（Tasks\<id>\过程记录-*.html），总表里只有引用 → 按需读回再渲染
    fillHistoryProc(procEl, task, task.process.procHtml || "");
    if (!task.process.procHtml && task.process.procHtmlFile) loadArchivedProcHtml(task, procEl);
    if (rl) rl.innerHTML = task.process.resultHtml || "";
    upgradeResultItems(rl);   // 历史快照若为早期版本（含下载链接、无打开按钮），就地升级
    const resultPhase = document.getElementById("result-phase");
    if (resultPhase) resultPhase.style.display = task.process.resultHtml ? "" : "none";   // 历史快照有结果 → ⑤下载结果
  } else {
    if (procEl) procEl.innerHTML = "";
    if (rl) rl.innerHTML = "";
    const resultPhase = document.getElementById("result-phase");
    if (resultPhase) resultPhase.style.display = "none";
  }
  // 历史任务：第④步「开始转换」发起区隐藏（右上角「重新转换」负责重跑该历史）
  const actionsRowHist = document.getElementById("actions-row");
  if (actionsRowHist) actionsRowHist.style.display = "none";
  // 历史任务输入源只显示本地文件；在线文档任务已不支持，回显时清空并提示
  _convEchoTask = true;
  if (blockInput) blockInput.classList.remove("link-only", "local-only");
  if (task.inputType === "local") {
    if (blockInput) blockInput.classList.add("local-only");
    convertLocalFiles = (task.files || []).map(f => ({ name: f.name || "", size: f.size || 0, path: f.path || "" }));
  } else {
    if (blockInput) blockInput.classList.add("link-only");
    clearLocalFiles();
    toast("该任务来源为在线文档，本版本已不再支持在线文档", "warn");
  }
  renderUploadList();
  // 等模板集列表填充后恢复选择
  await refreshConvertView(task.tsFile);
  if (task.tsFile) {
    window.__tplPickChecked = window.__tplPickChecked || {};
    window.__tplPickCheckedNames = window.__tplPickCheckedNames || {};
    window.__tplPickChecked[task.tsFile] = (task.checked || []).filter(k => typeof k === "string");
    window.__tplPickCheckedNames[task.tsFile] = taskCheckedNames(task);
    renderTplPick();
  }
  // 恢复"当时转换用的格式处理选项"到 ③ 卡片（来源：task.inheritAttrs + task.recognizeLeadParagraph；旧任务无字段则保持当前全局偏好）
  // 注意：仅把"本次任务勾选集合"应用回 DOM，不再回写 localStorage，避免破坏"全局默认偏好"语义。
  if (Array.isArray(task.inheritAttrs)) {
    // 兼容旧继承属性数组：只认白名单键；旧语义的 charScale 还原成「字符间距 + 字符缩放」两项
    const restored = expandLegacyInheritAttrs(task.inheritAttrs);
    // recognizeLeadParagraph：新任务记录布尔；旧任务无此字段 → 默认开（lead 勾选，与默认一致）
    if (task.recognizeLeadParagraph === false) {
      const li = restored.indexOf("lead");
      if (li >= 0) restored.splice(li, 1);
    } else if (!restored.includes("lead")) {
      restored.push("lead");
    }
    // reviewEnabled：内容校对开关（布尔）；旧任务无此字段 → 默认开（老版本转换本就执行校对，重转不倒退）
    if (task.reviewEnabled === false) {
      const ri = restored.indexOf("review");
      if (ri >= 0) restored.splice(ri, 1);
    } else if (!restored.includes("review")) {
      restored.push("review");
    }
    // nameCheck：人员姓名校对开关（布尔）；旧任务无此字段 → 默认关（该能力后加，老任务当时没这个选项）
    if (task.nameCheck === true) {
      if (!restored.includes("namecheck")) restored.push("namecheck");
    } else {
      const ni = restored.indexOf("namecheck");
      if (ni >= 0) restored.splice(ni, 1);
    }
    applyInheritPrefs(restored);
    // 名单选择：按该任务当时勾的名单回显（**不写 localStorage**，与上面同理）；再同步选择区显隐
    applyRosterSelection(task.nameRosters);
    syncRosterPickVisible();
  }
  applyInheritDisabled();   // 历史任务源类型 → 同步③卡片置灰
  // 折叠摘要：输入源 / 模板集 / 继承勾选 已在上面全部恢复，此时刷新三卡摘要行
  updateInputSummary();
  updateTemplateSummary();
  updateInheritSummary();
  activeTaskId = task.id;
  renderTasks();
  // 记录当前查看的历史任务：右上角「重新转换」用它的输入重发（不中断进行中的转换）。
  // 聚焦了进行中/刚结束的运行（liveRun）→ 重跑对象归聚焦任务，不设历史查看态。
  _historyViewTask = liveRun ? null : task;
  updateReconvertBtn();
  // 复位到页面顶部（不再自动往下滚）：过程/结果区就在表头三卡之下，首屏已可见，
  // 「滚到结果区」只会把用户从原位拽走（短过程时还因夹取停在底部）。
  resetConvScroll();
}

// 思考流折叠/展开：完成后默认折叠成单行摘要，点击展开全文。
// 绑定记在 DOM 上（innerHTML 序列化会丢 onclick，但 data-full-text 保留），快照恢复后重绑即复原交互。
function toggleReasonCollapsed(el) {
  if (!el) return;
  if (el.classList.contains("is-collapsed")) {
    el.classList.remove("is-collapsed");
    el.textContent = el.dataset.fullText || "";
  } else {
    el.classList.add("is-collapsed");
    el.textContent = "▸ 思考过程 · " + (el.dataset.fullText || "").length + " 字（点击展开）";
  }
  // 不做滚动：这是用户点开/收起某个思考块的动作，块就在点击处，把页面甩到底部会让人丢掉刚展开的内容
  snapshotConvState();
}
function bindReasonToggle(el) {
  if (!el) return;
  el.style.cursor = "pointer";
  el.onclick = () => {
    // 用户正在/已完成文本选择（拖选/双击选词）→ 不切折叠，避免选区被 textContent 替换吞掉、复制不下来
    const sel = window.getSelection && window.getSelection();
    if (sel && sel.toString && sel.toString().trim().length > 0) return;
    toggleReasonCollapsed(el);
  };
}
// 过程区被快照整体重写后重建交互：思考块（data-line="reason|<key>"）+ 阶段标签行（data-line="reasontag|<key>"）
function rebindReasonToggle(procEl) {
  if (!procEl) return;
  procEl.querySelectorAll(".conv-reason").forEach(el => {
    if (!el.dataset.fullText) return;
    bindReasonToggle(el);
  });
  procEl.querySelectorAll(".conv-reason-tag[data-line]").forEach(tag => {
    const key = (tag.getAttribute("data-line") || "").replace(/^reasontag\|/, "");
    if (!key) return;
    tag.style.cursor = "pointer";
    tag.onclick = () => {
      const sel = window.getSelection && window.getSelection();
      if (sel && sel.toString && sel.toString().trim().length > 0) return;
      let box = null;
      procEl.querySelectorAll(".conv-reason[data-line]").forEach(el => {
        if ((el.getAttribute("data-line") || "") === "reason|" + key) box = el;
      });
      if (box && box.dataset.fullText) toggleReasonCollapsed(box);
    };
  });
}

// ── 文档链接历史已移除（不再支持在线文档）──
let _urlHitHint = null;
function onUrlInput() {
  // 在线文档链接已移除，保留空实现避免引用报错
}

// ═══════════════════════════════════════════
// 转换视图
// ═══════════════════════════════════════════
async function refreshConvertView(forceFile, clearSelection) {
  const st = await refreshStatus().catch(() => null);
  if (!st) return;
  const all = [...st.templatesets];
  if (state.templateSet && state.currentFile && !all.some(t => t.file === state.currentFile)) {
    all.unshift({ file: state.currentFile, name: state.templateSet.name + "（编辑中）", templates: state.templateSet.templates.map(t => t.name) });
  }
  if (!all.length && state.templateSet) {
    all.push({ file: "(current)", name: state.templateSet.name + "（未保存）", templates: state.templateSet.templates.map(t => t.name) });
  }
  const prev = clearSelection ? "" : (forceFile || currentTsFile);
  renderTsPick(all, prev, clearSelection);
  renderTplPick();
  renderTasks(); // 同步左侧任务列表
}

// 渲染「模板集」单选列表（替代原下拉框）
function renderTsPick(all, prevFile, clearSelection) {
  const box = document.getElementById("ts-pick");
  box.innerHTML = "";
  if (!all.length) {
    box.innerHTML = `<div class="style-group-empty">暂无模板集，可在「模板设计」新建</div>`;
    currentTsFile = "";
    currentTsName = "";
    return;
  }
  const target = (prevFile && all.some(t => t.file === prevFile)) ? prevFile : (clearSelection ? "" : all[0].file);
  currentTsFile = target;
  currentTsName = target ? (all.find(t => t.file === target) || all[0]).name : "";
  for (const t of all) {
    const item = document.createElement("div");
    item.className = "ts-item" + (t.file === target ? " active" : "");
    item.dataset.file = t.file;
    item.textContent = t.name;
    item.title = t.name;
    item.addEventListener("click", () => {
      if (currentTsFile === t.file) return;
      currentTsFile = t.file;
      currentTsName = t.name;
      box.querySelectorAll(".ts-item").forEach(x => x.classList.toggle("active", x.dataset.file === t.file));
      renderTplPick();
    });
    box.appendChild(item);
  }
}

// 渲染「该集下模板」多选列表（基于 GET /api/templateset 取到的 templates 数组）
// 请求序号守卫：并发触发（如任务恢复 + 视图刷新）时只让最后一次渲染生效，
// 过期响应直接丢弃，避免模板项叠加/错乱。
let _tplPickSeq = 0;
let currentTsFile = "";   // 当前选中的模板集 file
let currentTsName = "";   // 当前选中的模板集显示名
function renderTplPick() {
  const mySeq = ++_tplPickSeq;
  const selFile = currentTsFile;
  const box = document.getElementById("tpl-pick");
  const lbl = document.getElementById("lbl-tpl-pick");
  const hint = document.getElementById("ts-hint");
  const cur = window.__tplPickChecked = window.__tplPickChecked || {};
  box.innerHTML = "";
  lbl.style.display = "";
  if (!selFile) { hint.style.display = "none"; return; }
  if (selFile === "(current)") {
    // 内存中编辑但未保存的模板集：直接从 state 取（同步渲染，无需守卫）
    const ts = state.templateSet;
    const tpls = (ts && ts.templates) || [];
    // 2026-09-01：选择模板集后模板默认不勾选，由用户自行勾选（历史勾选记录优先）
    const curKeys = cur[selFile] || [];
    fillTplPick(box, lbl, tpls, curKeys);
    const msg = !tpls.length ? "当前模板集为空。" : (!curKeys.length ? "（请勾选至少一个模板）" : "");
    hint.textContent = msg; hint.style.display = msg ? "" : "none";
    return;
  }
  // 已保存的：拉详情拿到完整 templates（含 key）
  api("/api/templateset?file=" + encodeURIComponent(selFile)).then(r => {
    if (mySeq !== _tplPickSeq) return; // 过期响应，丢弃
    const tpls = (r.data && r.data.templates) || [];
    const key = selFile;
    // 2026-09-01：首次选择该模板集默认不勾选，由用户自行勾选（已有勾选记录则保留）
    if (!cur[key]) cur[key] = [];
    fillTplPick(box, lbl, tpls, cur[key]);
    const msg = !tpls.length ? "该模板集为空。" : (!cur[key].length ? "（请勾选至少一个模板）" : "");
    hint.textContent = msg; hint.style.display = msg ? "" : "none";
  }).catch(e => {
    if (mySeq !== _tplPickSeq) return;
    hint.textContent = "加载模板集失败：" + e.message; hint.style.display = "";
  });
}

function fillTplPick(box, lbl, tpls, checkedKeys) {
  box.innerHTML = ""; // 防叠加：任何一次填充都先重置（配合 renderTplPick 序号守卫）
  lbl.style.display = "";
  if (!tpls.length) return;
  const checkedSet = new Set(checkedKeys);
  for (const t of tpls) {
    const id = "tpl-pick-" + t.key;
    // 模板说明 tooltip：悬停模板名显示（导入时填写的 templates[].description）；无说明不挂 title
    const desc = (t.description || "").trim();
    const label = document.createElement("label");
    label.innerHTML = `<input type="checkbox" id="${id}" data-key="${esc(t.key)}"${checkedSet.has(t.key) ? " checked" : ""}><span class="tpl-name"${desc ? ` title="${esc(desc)}"` : ""}>${esc(t.name || t.key)}</span>`;
    box.appendChild(label);
  }
  // 每次重建后重绑 change：维护勾选缓存 + 错误提示
  box.querySelectorAll('input[type=checkbox]').forEach(cb => {
    cb.addEventListener("change", () => {
      const selFile = currentTsFile;
      const checkedNow = Array.from(box.querySelectorAll('input[type=checkbox]:checked')).map(x => x.dataset.key);
      window.__tplPickChecked[selFile] = checkedNow;
      // 同步记下勾选模板的**名字**：用户手点的勾选同样要能在 key 失效时按名认回
      window.__tplPickCheckedNames = window.__tplPickCheckedNames || {};
      window.__tplPickCheckedNames[selFile] = selectedTplNames(tpls, checkedNow);
      const hint = document.getElementById("ts-hint");
      const msg = checkedNow.length ? "" : "（请勾选至少一个模板）";
      hint.textContent = msg; hint.style.display = msg ? "" : "none";
    });
  });
}

function selectedTplNames(tpls, keys) {
  const set = new Set(keys);
  return tpls.filter(t => set.has(t.key)).map(t => t.name || t.key);
}

// ── 勾选的「模板身份」解析（2026-09-13）──
// 背景：模板 key 是创建时生成的一次性标识（"imp"+时间戳+随机，见 ensureImportPlaceholder），
// 模板被删掉重建 / 同名重新导入 → key 必换；而历史任务记的是 key，重转时就「勾选项整体失效」。
// 用户看到的身份是模板**名**，所以：key 精确命中优先，命不中的按同名模板重挂——等价于「还是用日程安排跑一遍」。
function resolveCheckedKeys(tpls, keys, names) {
  const list = Array.isArray(tpls) ? tpls : [];
  const byKey = new Set(list.map(t => t.key));
  const hit = (Array.isArray(keys) ? keys : []).filter(k => byKey.has(k));
  const usedKeys = new Set(hit);
  const usedNames = new Set(list.filter(t => usedKeys.has(t.key)).map(t => t.name));
  for (const n of (Array.isArray(names) ? names : [])) {
    if (!n || usedNames.has(n)) continue;
    const t = list.find(x => x.name === n && !usedKeys.has(x.key));
    if (!t) continue;
    hit.push(t.key);
    usedKeys.add(t.key);
    usedNames.add(n);
  }
  return hit;
}

function sameKeyList(a, b) {
  const x = a || [], y = b || [];
  return x.length === y.length && x.every((k, i) => k === y[i]);
}

// 任务记录的「勾选模板名」：新任务直接存 checkedNames；存量任务（无该字段）从过程记录反查——
// telemetry[].outputs[].name 就是当次产出的模板名（= 当时勾选并实际使用的模板），可用于把旧 key 认回名字。
function taskCheckedNames(t) {
  if (!t) return [];
  if (Array.isArray(t.checkedNames) && t.checkedNames.length) return t.checkedNames.filter(n => typeof n === "string");
  const out = [];
  const tels = Array.isArray(t.telemetry) ? t.telemetry : (t.telemetry ? [t.telemetry] : []);
  for (const e of tels) {
    for (const o of ((e && e.outputs) || [])) {
      const n = o && o.name;
      if (n && !out.includes(n)) out.push(n);
    }
  }
  return out;
}

// ── 文档排版：本地文件（方案A，链接/本地互斥，可多选） ──
let convertLocalFiles = [];   // [{ name, size, path }]，path 为 /api/upload 落盘后返回

// ── 多转换并行运行态（2026-08-26 用户定调：允许多个转换任务并行，不中断旧流） ──
// convTasks：全部转换任务（进行中/已完成/失败），每项独立快照与输入。
// convState：指向「聚焦任务」（最近发起），兼容既有读法（running/procHtml…）；
//           无任务时为空态对象（不置 null，避免读 .running 崩）。
// focusedRunId：当前共享过程区（#conv-process/#result-list）显示的是哪个任务的流；
//           发起新任务/新建/历史查看时更新——非聚焦任务静默跑完（只收 done/error，
//           完成时用完成数据构造快照保存任务），不污染显示区，也不被中断。
let currentView = "convert";   // 当前激活视图（用于判断「转换在切走后完成」）
let convTasks = [];            // 全部转换任务（含已完成）
let convState = {
  runId: null,
  running: false,
  completedAway: false,   // 转换在用户切走其他视图后才完成：回来需恢复结果快照
  procHtml: "",
  resultHtml: "",
  input: null,
  tsFile: "",
  tsName: "",
  checked: []
};
let focusedRunId = null;   // 共享过程区当前归属的任务 runId（null=显示历史快照/空态）

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result).split(",")[1] || "");
    fr.onerror = () => reject(new Error("文件读取失败"));
    fr.readAsDataURL(file);
  });
}

function renderUploadList() {
  const list = document.getElementById("upload-list");
  if (!list) return;
  list.innerHTML = convertLocalFiles.map((f, i) => `
    <div class="upload-file">
      <span class="uf-ic">📄</span>
      <span class="uf-name" title="${esc(f.name)}">${esc(f.name)}</span>
      <span class="uf-size">${formatBytes(f.size)}</span>
      <span class="uf-remove" data-i="${i}" title="移除">✕</span>
    </div>`).join("");
  list.querySelectorAll(".uf-remove").forEach(el => {
    el.addEventListener("click", e => {
      e.stopPropagation();
      convertLocalFiles.splice(+el.dataset.i, 1);
      renderUploadList();
      hideUrlHint();   // 「已添加 N 个本地文件」随文件移除失效，不清理会一直挂着
    });
  });
  const hint = document.getElementById("upload-hint");
  if (hint) hint.style.display = convertLocalFiles.length ? "none" : "";
  applyInheritDisabled();   // 本地文件增减 → 同步③卡片置灰
}

// 链接 / 本地互斥已移除（仅本地）；保留空实现避免引用报错
function clearLocalFiles() { convertLocalFiles = []; renderUploadList(); }
function clearLink() { /* 在线链接已移除 */ }
function hideUrlHint() { /* 在线链接已移除 */ }

// 本地文件展示名：上传落盘的临时路径 → 原始文件名（addLocalFiles 里的映射）；无匹配原样返回。
// 服务端 read 事件只知道落盘路径，原始名只在前端有，读取行/统计卡源行经此换名，不向用户暴露 temp 路径
function localDisplayName(p) {
  const hit = convertLocalFiles.find(f => f.path === p);
  return (hit && hit.name) || p || "";
}

// 「复制链接」按钮已移除；保留空实现避免引用报错
function syncCopyUrlBtn() {
  // 在线链接已移除，无按钮需要同步
}

function buildConvertInputs() {
  // 仅支持本地上传；在线文档链接入口已移除。
  if (convertLocalFiles.length) return { inputType: "local", inputs: convertLocalFiles.map(f => f.path), label: convertLocalFiles.map(f => f.name).join("、"), files: convertLocalFiles.slice() };
  return null;
}

function truncateMiddle(s, max) {
  if (!s || s.length <= max) return s;
  const half = Math.floor(max / 2);
  return s.slice(0, half) + "…" + s.slice(-half);
}

function updateInputSummary() {
  const sum = document.getElementById("input-summary");
  if (!sum) return;
  if (convertLocalFiles.length) {
    const names = convertLocalFiles.map(f => f.name).join("、");
    sum.textContent = `本地文件（${convertLocalFiles.length}个）：${names}`;
    return;
  }
  sum.textContent = "";
}

function updateTemplateSummary() {
  const sum = document.getElementById("template-summary");
  if (!sum) return;
  const tsName = currentTsName || (state.templateSet && state.templateSet.name) || "";
  const checked = (window.__tplPickChecked || {})[currentTsFile] || [];
  if (tsName && checked.length) sum.textContent = `${tsName} · 已选 ${checked.length} 个模板`;
  else if (tsName) sum.textContent = tsName;
  else sum.textContent = "";
}

// 格式处理选项中文名（折叠摘要用；与 #inherit-grid checkbox 顺序一致）
const INHERIT_CN = {
  lead: "识别段首标题句", review: "内容校对", namecheck: "人员姓名校对",
  bold: "加粗", italic: "斜体", underline: "下划线", strike: "删除线",
  color: "字体颜色", bgColor: "底色高亮", charSpacing: "字符间距", charScale: "字符缩放",
};
// 摘要分组（与 #inherit-grid 三组一致）：单选项组直接显示选项名，多选项组 ≤2 项列名、>2 项收敛为「N 项」
const INHERIT_GROUPS = [
  { label: "段落结构", keys: ["lead"] },
  { label: "文字格式", keys: ["bold", "italic", "underline", "strike", "color", "bgColor", "charSpacing", "charScale"] },
  { label: "内容处理", keys: ["review", "namecheck"] },
];
// 折叠态摘要：按组聚合当前勾选；未勾选 → 按模板规范
function updateInheritSummary() {
  const sum = document.getElementById("inherit-summary");
  if (!sum) return;
  const checked = new Set(inheritOptInputs().filter(i => i.checked).map(i => i.value));
  const parts = INHERIT_GROUPS.map(g => {
    const hit = g.keys.filter(k => checked.has(k));
    if (!hit.length) return null;
    if (g.keys.length === 1) return INHERIT_CN[hit[0]];
    if (hit.length <= 2) return hit.map(k => INHERIT_CN[k]).join(" · ");
    return `${g.label} ${hit.length} 项`;
  }).filter(Boolean);
  sum.textContent = parts.length ? "格式处理：" + parts.join(" · ") : "格式处理：无（按模板规范）";
}

// 同步当前共享过程区 DOM 到「聚焦任务」的快照（并行模型：每个任务独立快照，
// 显示区只属于 focusedRunId 任务；切走/被新任务覆盖后，快照仍完整保留）
function snapshotConvState() {
  const procEl = document.getElementById("conv-process");
  const rl = document.getElementById("result-list");
  const t = focusedRunId ? convTasks.find(x => x.runId === focusedRunId) : null;
  if (t) {
    if (procEl) t.procHtml = procEl.innerHTML;
    if (rl) t.resultHtml = rl.innerHTML;
  }
}

// 从任务列表恢复「最新任务」的过程/结果显示（切回 convert 视图时调用）。
// 并行模型：恢复最近发起/查看的任务，不中断任何任务流；聚焦任务=该任务。
function restoreConvProcess() {
  const t = convTasks[convTasks.length - 1];
  if (!t) return;
  convState = t;                 // 聚焦 = 最新任务
  focusedRunId = t.runId;        // 该任务（若仍在跑）的流事件继续写显示区
  const rb = document.getElementById("result-block");
  const procEl = document.getElementById("conv-process");
  const rl = document.getElementById("result-list");
  if (rb) rb.style.display = "block";
  if (procEl) {
    procEl.innerHTML = t.procHtml || "";
    rebindReasonToggle(procEl);
  }
  if (rl) rl.innerHTML = t.resultHtml || "";
  upgradeResultItems(rl);   // 历史快照若为早期版本（含下载链接、无打开按钮），就地升级
  const resultPhase = document.getElementById("result-phase");
  if (resultPhase) resultPhase.style.display = t.resultHtml ? "" : "none";   // 恢复时按是否有结果显隐⑤
  // 恢复时同样折叠输入/模板卡片，聚焦过程区
  const blockInput = document.getElementById("block-input");
  const blockTemplate = document.getElementById("block-template");
  const blockInherit = document.getElementById("block-inherit");
  if (blockInput) blockInput.classList.add("collapsed");
  if (blockTemplate) blockTemplate.classList.add("collapsed");
  if (blockInherit) blockInherit.classList.add("collapsed");
  // 恢复折叠摘要
  updateInputSummary();
  updateTemplateSummary();
  updateInheritSummary();
  // 恢复输入源显示（该任务实际使用的方式）。在线文档任务已不支持，回显时清空并提示。
  _convEchoTask = true;
  if (blockInput) blockInput.classList.remove("link-only", "local-only");
  if (t.input) {
    if (t.input.inputType === "local") {
      if (blockInput) blockInput.classList.add("local-only");
      convertLocalFiles = (t.input.files || []).slice();
    } else {
      if (blockInput) blockInput.classList.add("link-only");
      clearLocalFiles();
      toast("该任务来源为在线文档，本版本已不再支持在线文档", "warn");
    }
  }
  renderUploadList();
  resetConvScroll();   // 复位到顶部：恢复快照不是「跟随流」，滚到底会把用户扔到过程区末尾
  // 恢复任务视图：第④步「开始转换」发起区保持隐藏（仅在新建表单显示）；重跑由右上角「重新转换」承担
  const actionsRowRes = document.getElementById("actions-row");
  if (actionsRowRes) actionsRowRes.style.display = "none";
  // 已脱离历史查看：右上角「重新转换」指向聚焦任务
  _historyViewTask = null;
  updateReconvertBtn();
  // 已完成任务的恢复只回显一次：切回时若聚焦任务已完成，清除「切走后完成」标记
  if (!t.running) t.completedAway = false;
}

// 右上角「重新转换」按钮状态机（进行中 / 已完成 / 历史查看 三态常驻过程卡右上角）。
// 重跑对象 = 聚焦任务（进行中/完成）或当前打开的历史任务（_historyViewTask）。
let _historyViewTask = null;   // openTask 当前查看的历史任务（供右上角重新转换）
function updateReconvertBtn() {
  const el = document.getElementById("btn-reconvert");
  const stopEl = document.getElementById("btn-stop-convert");
  if (!el) return;
  const src = focusedRunId ? convTasks.find(x => x.runId === focusedRunId) : null;
  const t = src || _historyViewTask;
  // 聚焦任务进行中 → 显示「停止」（中断本任务流）；结束后 / 历史查看 → 显示「重新转换」。
  // 两按钮互斥，历史查看（src=null）无停止。
  if (src && src.running) {
    el.style.display = "none";
    if (stopEl) { stopEl.style.display = ""; stopEl.disabled = false; }
    return;
  }
  if (stopEl) stopEl.style.display = "none";
  if (!t) { el.style.display = "none"; return; }
  el.style.display = "";
  el.disabled = false;
  el.title = "重新转换当前任务";
}

// 右上角「停止」：中断聚焦任务的转换流。
// 与「重新转换」的 abort 共用 AbortController，靠 task.stoppedByUser 区分：
// 用户停止 → 面板追加「已停止」行；重新转换接管 → 静默退出（新任务已发起）。
function stopFocusedTask() {
  const t = convTasks.find(x => x.runId === focusedRunId);
  if (!t || !t.running || !t.ac) return;
  t.stoppedByUser = true;
  t.ac.abort();
}

// 右上角「重新转换」：重跑 已结束的聚焦任务 或 当前历史查看的任务。
// 重跑 = 用**界面当前的表单状态**（输入源 / 模板集 / 模板勾选 / ③格式处理选项）发起一次新转换（聚焦新任务）。
// ★ 不回填任何历史快照（2026-09-13 用户拍板）：历史记录只用于「打开任务时回显」（openTask），
//   绝不用于改写表单——否则用户在界面上刚改的勾选（如③「内容校对」）会在点重转的瞬间被旧快照覆盖，
//   表现为「明明勾了却没生效」，且卡片随即折叠、用户完全看不到勾被吃掉。
//   语义统一为：打开时还原给你看，跑的时候一律以界面当前选择为准。
// 进行中不显示此按钮 → 重跑不中断任何流。
async function reconvertTask() {
  const src = focusedRunId ? convTasks.find(x => x.runId === focusedRunId) : null;
  const t = (src && !src.running) ? src : _historyViewTask;   // 仅已结束任务或历史任务可重跑
  if (!t) return;
  // 重置上一轮的「段落分类 / 内容校对」完成标志（2026-09-08 拍板）：
  // 完成标志就写在过程区的思考流标签上（「段落分类 · 完成，点击展开」），旧运行快照与历史任务
  // 快照都带着它。不清理会两处显旧：①点完重跑到新任务真正起流之间（要先取模板集、校验模型配置）
  // 界面仍停在上一轮的完成态；②新运行被停止/失败后回看该任务，旧完成态又从 process 快照复活。
  // 只清过程快照（完成标志的载体），结果列表保留——重跑失败也仍能下载上一次的产物。
  if (src) { src.procHtml = ""; src.resultHtml = ""; }
  if (t !== src && t.process) t.process.procHtml = "";
  const procEl0 = document.getElementById("conv-process");
  const rl0 = document.getElementById("result-list");
  const rp0 = document.getElementById("result-phase");
  if (procEl0) procEl0.innerHTML = "";
  if (rl0) rl0.innerHTML = "";
  if (rp0) rp0.style.display = "none";
  doConvert();
}

// 真正的滚动容器是 .convert-wrap（overflow-y:auto），#conv-process 自身不滚动，
// 故不能对 procEl 设 scrollTop（无效）。这里从 conv-process 向上找最近的滚动祖先并滚到底。
function scrollConvBottom() {
  const proc = document.getElementById("conv-process");
  if (!proc) return;
  // 1) 内部 reasoning 容器优先滚到底（保持最新思考可见）
  const reasons = proc.querySelectorAll(".conv-reason");
  const reason = reasons[reasons.length - 1];
  if (reason && reason.scrollHeight > reason.clientHeight + 1) {
    reason.scrollTop = reason.scrollHeight;
  }
  // 2) 外层：从 proc 向上找第一个 overflow 滚动祖先
  let el = proc;
  while (el && el !== document.body) {
    const cs = getComputedStyle(el);
    if ((cs.overflowY === "auto" || cs.overflowY === "scroll") && el.scrollHeight > el.clientHeight + 1) {
      el.scrollTop = el.scrollHeight;
      return;
    }
    el = el.parentElement;
  }
  // 3) 兜底：整页
  if (window.scrollTo) window.scrollTo({ top: document.documentElement.scrollHeight });
}

// 进入/切换任务时把转换页复位到顶部。与 scrollConvBottom 是一对相反的原语：
// 自动滚动只服务于「正在进行的流」（跟随最新），其余一切进入动作都回顶。
// ★ 不得改成「滚到过程区顶部」：#result-block 位于卡片下部（表头 3 卡之下），
//   对它用 scrollIntoView({block:"start"}) 时滚动量一旦超过容器可滚动的最大量，
//   浏览器会把它夹到 max ⇒ 页面自己停在「底部」。内容越短、窗口越高，越必然发生：
//   实测 20 行过程记录落在 299/308（≈底部），过程记录更短时直接 clamp 到底。
// ★ 同理，清空内容后（新建任务）若不复位，浏览器会把旧的 scrollTop 夹到新的底部 ——
//   表现就是「点了『文档排版』新页面停在底部、看不到地址/模板卡」。
function resetConvScroll() {
  const proc = document.getElementById("conv-process");
  if (!proc) return;
  let el = proc;
  while (el && el !== document.body) {
    const cs = getComputedStyle(el);
    if (cs.overflowY === "auto" || cs.overflowY === "scroll") { el.scrollTop = 0; return; }
    el = el.parentElement;
  }
  if (window.scrollTo) window.scrollTo({ top: 0 });
}

// 转换流程内的提示只走页面（转换过程有专门页面，不弹 toast）：
// 在「转换过程」面板追加一行，并确保结果区可见、清理旧内容。
function showConvPageMsg(msg, type) {
  const rb = document.getElementById("result-block");
  const procEl = document.getElementById("conv-process");
  const rl = document.getElementById("result-list");
  if (rb) rb.style.display = "block";
  if (procEl) procEl.innerHTML = "";
  if (rl) rl.innerHTML = "";
  const resultPhase = document.getElementById("result-phase");
  if (resultPhase) resultPhase.style.display = "none";   // 错误提示无产出 → 隐藏⑤下载结果
  if (!procEl) return;
  const d = document.createElement("div");
  const ic = type === "err" ? "✗" : type === "ok" ? "✓" : "•";
  d.className = "conv-line" + (type === "info" ? " info" : procLineCls(ic));
  d.innerHTML = `<span class="conv-ic">${procIcon(ic)}</span><span class="conv-msg">${esc(msg)}</span>`;
  procEl.appendChild(d);
  scrollConvBottom();
}

// 输入区（链接 / 上传）的提示也走页面内联；在线链接已移除，改走 toast
function showUrlHint(msg, type) {
  if (msg) toast(msg, type);
}

async function uploadConvertFile(f) {
  // .json = 在线文档源快照（转换时自动存进任务目录），可当本地文件重转：免登录、免联网、结果可复现
  const ALLOWED_EXT = [".docx", ".doc", ".txt", ".md", ".json"];
  const ext = (f.name.split(".").pop() || "").toLowerCase();
  if (!ALLOWED_EXT.includes("." + ext)) { showUrlHint("不支持的格式：" + f.name + "（仅支持 docx / doc / txt / md / json 源快照）", "err"); return null; }
  if (f.size > 50 * 1024 * 1024) { showUrlHint("文件过大：" + f.name + "，请拆分后重试", "err"); return null; }
  const b64 = await readFileAsBase64(f);
  // 服务端失败不再静默：api 非 2xx 抛错 / 响应缺路径均明确提示（上传失败要给原因，不能无声消失）
  let r;
  try {
    r = await api("/api/upload", { method: "POST", body: JSON.stringify({ fileName: f.name, base64: b64 }) });
  } catch (e) {
    showUrlHint("添加失败：" + f.name + "（" + e.message + "）", "err");
    return null;
  }
  if (!r || !r.ok || !r.path) {
    showUrlHint("添加失败：" + f.name + "（" + ((r && r.error) || "服务未返回文件路径") + "）", "err");
    return null;
  }
  return { name: f.name, size: f.size, path: r.path };
}

async function doConvert() {
  const built = buildConvertInputs();
  if (!built) {
    showConvPageMsg("请先选择本地文件", "err");
    return;
  }

  let ts = state.templateSet;
  const selFile = currentTsFile;
  if (!ts || (state.currentFile && selFile !== state.currentFile) || !state.currentFile) {
    if (selFile && selFile !== "(current)") {
      const r = await api("/api/templateset?file=" + encodeURIComponent(selFile));
      ts = r.data;
    }
  }
  if (!ts || !ts.templates || !ts.templates.length) { showConvPageMsg("请先配置模板集（至少一个模板）", "err"); return; }

  // 按勾选过滤：未勾选任何模板直接报错。
  // 勾选按「key 优先、模板名兜底」解析（resolveCheckedKeys）：模板被删掉重建 / 同名重新导入时 key 必换，
  // 历史任务记的旧 key 会整体失效（2026-09-13 排查：日程安排重导入后重转报「所选模板集不含勾选的模板」）。
  // 用户眼里的模板身份是名字，按名重挂 = 仍然用「日程安排」跑一遍，属于程序该自己消化的兼容。
  const checkedRaw = (window.__tplPickChecked || {})[selFile] || [];
  const namesRaw = (window.__tplPickCheckedNames || {})[selFile] || [];
  if (!checkedRaw.length && !namesRaw.length) { showConvPageMsg("请至少勾选一个模板", "err"); return; }
  let checked = resolveCheckedKeys(ts.templates, checkedRaw, namesRaw);
  let picked = ts.templates.filter(t => checked.includes(t.key));
  // 仍匹配不到：内存快照可能已过期（模板集在别处被静默改写——写入非当前集时只改盘、不动内存）。
  // 以磁盘为准重拉一次再解析，避免「界面看到的是新模板集、转换用的是旧快照」。
  if (!picked.length && selFile && selFile !== "(current)") {
    try {
      const r2 = await api("/api/templateset?file=" + encodeURIComponent(selFile));
      const tpls2 = (r2 && r2.data && r2.data.templates) || [];
      const hit2 = resolveCheckedKeys(tpls2, checkedRaw, namesRaw);
      if (hit2.length) { ts = r2.data; checked = hit2; picked = tpls2.filter(t => hit2.includes(t.key)); }
    } catch (e) { /* 拉取失败：按原判定走下面的提示 */ }
  }
  if (!picked.length) {
    const miss = namesRaw.length ? `勾选的模板「${namesRaw.join("、")}」` : "原勾选的模板";
    showConvPageMsg(miss + "已不在该模板集中，请重新勾选", "err");
    return;
  }
  const tsFiltered = { ...ts, templates: picked };
  // 勾选被「按名重挂」改写 → 回写缓存并重渲染模板列表，保证界面勾选与本次转换实际使用的一致
  // （否则会出现「界面看着勾着、转换说没勾」的错位观感）
  if (!sameKeyList(checked, checkedRaw)) {
    window.__tplPickChecked = window.__tplPickChecked || {};
    window.__tplPickCheckedNames = window.__tplPickCheckedNames || {};
    window.__tplPickChecked[selFile] = checked.slice();
    window.__tplPickCheckedNames[selFile] = selectedTplNames(ts.templates, checked);
    renderTplPick();
  }

  // LLM 前置检查：转换主链路依赖 LLM，未配置直接提示——不发起任务、不产生历史记录
  if (!(await checkLlmConfigured())) {
    showConvPageMsg("文档排版需配置模型服务：请到「设置」启用一个模型服务", "err");
    return;
  }

  const btn = document.getElementById("btn-convert");
  const rb = document.getElementById("result-block");
  const procEl = document.getElementById("conv-process");
  const rl = document.getElementById("result-list");
  const blockInput = document.getElementById("block-input");
  const blockTemplate = document.getElementById("block-template");
  const blockInherit = document.getElementById("block-inherit");

  // 第 3 步：格式处理选项（顶部快照一次，本任务内不再变；旧任务无字段时回退 loadInheritPrefs）
  // 勾选拆三路：8 项文字属性 → 继承白名单 inheritAttrs；lead → recognizeLeadParagraph；review → reviewEnabled
  const fmtOptsRun = inheritOptInputs().filter(x => x.checked).map(x => x.value);
  const inheritAttrsRun = fmtOptsRun.filter(x => INHERIT_ATTR_KEYS.includes(x));
  const recognizeLeadRun = fmtOptsRun.includes("lead");
  const reviewRun = fmtOptsRun.includes("review");
  // 人员姓名校对：勾了才跑；名单取选择区当前勾选（**只传 id**，名单内容由服务端读盘解析——
  // 名单是服务端的用户数据，不必让前端把整份名单再传回去）
  const nameCheckRun = fmtOptsRun.includes("namecheck");
  const nameRostersRun = nameCheckRun ? selectedRosterIds() : [];
  // 勾了姓名校对但一份名单都没选（或还没建名单）：本次什么也校对不了，当场拦下。
  // 不拦会白跑一趟 AI，结果卡再告诉用户"未执行"，用户的时间已经花了。
  if (nameCheckRun && !nameRostersRun.length) {
    showConvPageMsg(rosterCache === null
      ? "名单读取失败：请到「设置 → 人员名单」检查后重试"
      : "已勾选「人员姓名校对」，但没有可用名单：请到「设置 → 人员名单」新建一份并勾选", "err");
    // 提示本身走过程区（在页面下部），用户看不到上面的名单选择区 —— 这里把第 3 步收尾的那份
    // 提示同步写清楚并保持可见，避免"报了错却不知道去哪儿改"
    if (nameCheckRun) updateRosterPickHint();
    return;
  }
  // LLM 流式行文案（内容校对/姓名校对都是独立任务；未真正要跑的不提，避免提示与后端行为不符）
  const llmTaskTxt = ["段落分类"]
    .concat(reviewRun ? ["内容校对"] : [])
    .concat((nameCheckRun && nameRostersRun.length) ? ["姓名校对"] : [])
    .join("+");

  // 折叠前先把摘要写入折叠行
  updateInputSummary();
  updateTemplateSummary();
  updateInheritSummary();

  // 折叠输入/模板/继承卡片，让过程区占据可视区域（保留一行摘要）
  if (blockInput) blockInput.classList.add("collapsed");
  if (blockTemplate) blockTemplate.classList.add("collapsed");
  if (blockInherit) blockInherit.classList.add("collapsed");

  rb.style.display = "block";
  // 一次运行 = 一份全新的过程区：思考流状态（reasonPhase/reasonBlocks/reasonBufs）是本次运行的
  // 闭包变量，随函数进入天然归零——「段落分类 / 内容校对」的完成标志不会跨运行残留。
  // 旧运行/历史任务的快照清理在 reconvertTask 里做（重跑入口）。
  procEl.innerHTML = "";
  rl.innerHTML = "";
  const resultPhase = document.getElementById("result-phase");
  if (resultPhase) resultPhase.style.display = "none";   // 新一次生成：⑤下载结果待完成后再显示
  // 用户定调（15:1x）：点击「开始转换」发起后，按钮消失（隐藏第④步发起区），
  // 转换过程中不显示「重新转换」、不中断任何流；结束后/历史查看由右上角「重新转换」承担重跑
  const actionsRow = document.getElementById("actions-row");
  if (actionsRow) actionsRow.style.display = "none";

  // 并行模型：发起新任务 = 聚焦它（显示区展示本任务过程）；进行中的旧任务不中断，
  // 静默跑完（只收完成/错误事件），完成后保存任务，可从左侧历史任务查看结果。
  // N7 + 方案A（2026-08-27）：发起即确定 taskId——按链接去重查已有任务记录，
  // 命中则复用其 id 作 taskId（产物继续落同一 Tasks\<taskId>\ 目录，_N 后缀防覆盖，
  // 重转不产生孤儿目录）；无记录才新生成。记录 id 与产物目录名严格一致。
  const prevTaskRecord = loadTasks().find(t => t.url === built.label);
  const taskId = prevTaskRecord ? prevTaskRecord.id : "t" + Date.now();
  const task = {
    runId: "r" + Date.now(),
    taskId,
    running: true,
    completedAway: false,
    done: false,
    procHtml: "",
    resultHtml: "",
    input: built,
    tsFile: selFile,
    tsName: currentTsName,
    checked: checked.slice(),
    checkedNames: selectedTplNames(tsFiltered.templates, checked),   // 重跑/重转按名认回用（见 resolveCheckedKeys）
    inheritAttrs: inheritAttrsRun.slice(),        // 重跑还原用（右上角「重新转换」）
    recognizeLeadParagraph: recognizeLeadRun,
    reviewEnabled: reviewRun,                    // 重跑还原用（同 recognizeLeadParagraph）
    nameCheck: nameCheckRun,                     // 人员姓名校对开关（重跑还原用）
    nameRosters: nameRostersRun.slice(),         // 本次勾选的名单 id（重跑还原用；名单内容在服务端）
    ac: null                                       // 本任务流 AbortController（仅「重新转换当前任务」时中断本任务）
  };
  task.ac = new AbortController();
  convTasks.push(task);
  // 发起即创建任务占位（方案 A，2026-09-01）：任务从「开始转换」起就出现在左侧列表，
  // 进行中带「转换中」徽标；完成/失败后收尾 upsertTask 按同 url 去重更新同一记录。
  // 不传 process/telemetry：占位不清旧快照（重转同一文档时点开仍见上次结果）。
  upsertTask({
    taskId, title: built.label, url: built.label, inputType: built.inputType, files: built.files,
    tsFile: task.tsFile, tsName: task.tsName, checked: task.checked, checkedNames: task.checkedNames,
    inheritAttrs: task.inheritAttrs, recognizeLeadParagraph: task.recognizeLeadParagraph, reviewEnabled: task.reviewEnabled,
    nameCheck: task.nameCheck, nameRosters: task.nameRosters,
  });
  convState = task;          // 聚焦 = 最新任务
  focusedRunId = task.runId; // 本任务流事件写共享显示区
  updateReconvertBtn();

  // 滚动到过程区顶部，开始实时跟随
  rb.scrollIntoView({ behavior: "smooth", block: "start" });

  // lineKey：行锚点。切走再切回时过程区被快照整体替换，闭包里的元素引用会失效，
  // 靠这个锚点在活动 DOM 里重新定位（否则后续更新写进脱离文档的节点 = 面板冻结）
  const appendLine = (icon, text, lineKey) => {
    const d = document.createElement("div");
    d.className = "conv-line" + procLineCls(icon);
    if (lineKey) d.setAttribute("data-line", lineKey);
    d.innerHTML = `<span class="conv-ic">${procIcon(icon)}</span><span class="conv-msg">${text}</span>`;
    procEl.appendChild(d);
    scrollConvBottom();
    snapshotConvState();
    return d;
  };
  // 失败块：把后端翻好的三段话（出了什么事 / 为什么 / 怎么办）分行摆开，技术原文折进 details。
  // 为什么不用一行：一条 200 字的报错挤在单行里，用户只会看到 ellipsis 截断的前半截。
  // f 为空（后端认不出这类错误）时退回原样单行文本 —— 不在这里编文案。
  const appendFailBlock = (f, fallbackMsg) => {
    if (!f || !f.title) { appendLine("✗", "转换失败：" + esc(fallbackMsg || "未知错误")); return; }
    const d = document.createElement("div");
    d.className = "conv-line err conv-fail";
    const stage = f.stageLabel ? esc(f.stageLabel) : "";
    const tech = [f.tech, f.logFile ? "日志：" + f.logFile : ""].filter(Boolean).map(esc).join("<br>");
    d.innerHTML = `<span class="conv-ic">${procIcon("✗")}</span><span class="conv-msg">`
      + `<span class="cf-head">转换失败${stage ? " · " + stage : ""}</span>`
      + `<span class="cf-title">${esc(f.title)}</span>`
      + (f.detail ? `<span class="cf-detail">${esc(f.detail)}</span>` : "")
      + (f.action ? `<span class="cf-action">→ ${esc(f.action)}</span>` : "")
      + (tech ? `<details class="cf-tech"><summary>技术详情</summary><div>${tech}</div></details>` : "")
      + `</span>`;
    procEl.appendChild(d);
    scrollConvBottom();
    snapshotConvState();
  };
  const setLine = (el, icon, text) => {
    if (!el) return;
    el.innerHTML = `<span class="conv-ic">${procIcon(icon)}</span><span class="conv-msg">${text}</span>`;
    scrollConvBottom();
    snapshotConvState();
  };

  let doneData = null;
  let streamError = null;
  // 后端翻好的失败文案（SSE error 事件的 friendly 字段）。为空时前端退回原始 message。
  let streamFriendly = null;
  let genTimers = [];   // 每个「生成中」行的计时器（原地替换成完成行）
  let readTimer = null; // 「读取中…」行的计时器（在线 fetch 通常 30s+，必须有实时反馈）
  let llmTimer = null;  // LLM 主行计时器（实时累计秒数：从「调用 LLM」一直累计到「完成」）
  let llmStartT = 0;    // LLM 阶段起点；未启动时为 0（避免误算时长）
  let llmLine = null;   // 「调用LLM中…」行元素
  let readCharCount = null; // 读取阶段原始字数（供 parse 行标注校对 Δ）

  // 本任务流标识：渲染段用它判「行/块是否归属本流」——失焦后旧流的残留计时器与缓冲一律停手
  const runId = task.runId;

  // ══ 思考流渲染 START（回归脚本按本标记抽取本段真实实现执行：标记名不能改，段内代码须自包含）══
  let reasonBlocks = {}; // jobKey → { el, tag }：段落分类/内容校对/姓名校对各一块，不再混写同一容器
  let reasonBufs = {};   // jobKey → 待落盘文本（逐 chunk 只做字符串累加，零 DOM 操作）
  let reasonRaf = 0;     // flush 的动画帧句柄（0 = 无待刷帧）
  let lastReasonSnap = 0; // 思考流快照节流时间戳（快照=序列化整个面板，不能逐块做）
  let reasonPhase = {};  // jobKey → think（思考中）/ gen（生成结果中）/ done（完成）/ fail（失败跳过）

  const REASON_LABEL = { convert: "段落分类", review: "内容校对", names: "人员姓名校对" };
  const reasonLabel = (key) => REASON_LABEL[key] || String(key || "模型");
  // 表格识别（table / table1…）是转换内部环节（判表头行与跨页重复），界面上没有对应开关、用户勾不到
  // ⇒ 不进过程面板：不建块、不显示思考流、不驱动主线文案。过程面板列出的阶段必须与主线文案列出的
  //   那几路（段落分类/内容校对/姓名校对）完全一致，多出用户没得选的一路只会让人以为有这功能。
  // 表头/跨页判断仍照跑（决定表头行加粗与跨页重复），只是不出现在界面上；思考与输入输出照旧进 LLM 日志留痕。
  // 返回 "" = 该路不显示；reasonKeyOf 的每个消费者都按「空 key = 不处理」走，新增事件类型天然继承本口径。
  const INTERNAL_STAGE_KEY = /^table\d*$/;
  const reasonKeyOf = (d) => {
    const k = (d && (d.jobKey || d.stage)) || "";
    return INTERNAL_STAGE_KEY.test(k) ? "" : k;
  };
  const lineKey = (kind, k) => kind + "|" + k;

  // 过程区被快照整体重写后（切视图/切任务/并行任务接管显示区），闭包持有的元素引用即失效。
  // 统一按 data-line 锚点在活动 DOM 里重新定位，避免往脱离文档的节点上写（表现：面板冻结/思考流停住不动）
  //
  // 判活用 host.contains 而非 el.isConnected：isConnected 只答"挂在任意节点树下"，
  // 答不了"挂的是不是当前这块面板"。面板被整体替换（旧面板未从文档摘除、只是不再显示）时
  // isConnected 仍为真 → 缓存的死引用被判活 → 写入的节点用户看不见 → 标签永远停在「思考中…」。
  // contains 同时约束"在文档里"与"在当前面板下"，是唯一正确的判据（2026-09-14 实测复现）。
  const reasonHost = () => document.getElementById("conv-process");
  const inHost = (n) => !!n && !!host_ && host_.contains(n);
  let host_ = null;   // 活动面板引用（每次定位前刷新，避免每元素都查一次 DOM）
  // 在当前面板里定位一行：缓存引用仍在面板下 → 直接用；否则按 data-line 锚点重定位；都没有 → null。
  // ⚠ host_ 必须在这里自己刷新：行计时器（读取/LLM/生成中）的 tick 不经过 getReasonBlock，
  //   靠调用顺序拿到的是上一帧的面板引用甚至 null，会把「面板已换」误判成「行丢了」。
  const liveLine = (cached, key) => {
    host_ = reasonHost();
    if (cached && inHost(cached)) return cached;
    if (!host_) return null;
    const all = host_.querySelectorAll("[data-line]");
    for (const el of all) { if (el.getAttribute("data-line") === key) return el; }
    return null;
  };
  // 行计时器的判活口径：行必须仍在**当前**面板里，且归属**本**任务流。返回 null ⇒ 调用方停表。
  // ⚠ 不得用「liveLine 找不到就退回旧引用」兜底：旧引用是脱离面板的死节点，属性判活（data-runid）
  //   永远通过 ⇒ 失焦任务的计时器停不下来，每 500ms 调一次 scrollConvBottom()，
  //   把新建任务页的滚动容器顶到底部（用户点「文档排版」开新任务，页面自己一直往下滚）。
  const liveOwnLine = (cached, key) => {
    const el = liveLine(cached, key);
    return (el && el.getAttribute("data-runid") === runId) ? el : null;
  };
  const getReasonBlock = (key) => {
    host_ = reasonHost();
    if (!host_) return null;
    const cur = reasonBlocks[key];
    // 思考块与阶段标签**都必须**在当前面板下才算有效：早先只判 el，标签丢失时短路命中
    // 旧 blk（tag 为 null）→ setReasonPhase 静默 return → 标签永远停在「思考中…」（2026-09-14 修复）
    if (cur && inHost(cur.el) && inHost(cur.tag)) return cur;
    const el = liveLine(cur && cur.el, lineKey("reason", key));
    const tag = liveLine(cur && cur.tag, lineKey("reasontag", key));
    if (!el && !tag) return null;
    const blk = { el, tag };
    reasonBlocks[key] = blk;
    return blk;
  };
  const ensureReasonBlock = (key) => {
    const found = getReasonBlock(key);
    if (!found) return null;
    // 缺失的那一半就地补建（挂在思考块之后 / 面板末尾），而不是把整个块丢掉重建 ——
    // 重建会把已累积的思考流文本清空（用户看到思考内容凭空消失）。
    if (!found.el || !found.tag) {
      if (!found.tag) {
        const tag = document.createElement("div");
        tag.className = "conv-reason-tag";
        tag.setAttribute("data-line", lineKey("reasontag", key));
        tag.textContent = reasonLabel(key) + " · 思考中…";
        if (found.el && found.el.parentNode) found.el.parentNode.insertBefore(tag, found.el);
        else host_.insertBefore(tag, host_.firstChild);
        found.tag = tag;
      }
      if (!found.el) {
        const el = document.createElement("div");
        el.className = "conv-reason";
        el.setAttribute("data-line", lineKey("reason", key));
        if (found.tag) found.tag.parentNode.appendChild(el);
        else host_.appendChild(el);
        found.el = el;
      }
      reasonBlocks[key] = found;
    }
    if (!reasonPhase[key]) reasonPhase[key] = "think";
    return found;
  };
  // 整块新建（首次收到该路思考流时用）：与补建分开，避免上面的"保内容"补建逻辑被误用于全新块
  const createReasonBlock = (key) => {
    host_ = reasonHost();
    if (!host_) return null;
    const tag = document.createElement("div");
    tag.className = "conv-reason-tag";
    tag.setAttribute("data-line", lineKey("reasontag", key));
    tag.textContent = reasonLabel(key) + " · 思考中…";
    const el = document.createElement("div");
    el.className = "conv-reason";
    el.setAttribute("data-line", lineKey("reason", key));
    host_.appendChild(tag);
    host_.appendChild(el);
    const blk = { el, tag };
    reasonBlocks[key] = blk;
    if (!reasonPhase[key]) reasonPhase[key] = "think";
    return blk;
  };
  // 阶段状态可见化：思考中 → 正在生成结果 → 完成。
  // 「正在生成结果」必须有：LLM 切到正文输出后不再有思考内容，少了这一条界面就长时间无变化 → 体感"卡住"
  const setReasonPhase = (key, phase) => {
    if (!key) return;
    reasonPhase[key] = phase;
    // 走到这里说明该路 job 有动静（思考/结束），块不存在就整块新建；
    // 存在但缺一半则由 ensureReasonBlock 就地补建 —— 两条路都保证 tag 一定拿得到，
    // 绝不因 tag 缺失而静默 return（那正是「一直停在思考中…」的成因）
    const blk = ensureReasonBlock(key) || createReasonBlock(key);
    if (!blk || !blk.tag) return;
    const tail = phase === "gen" ? " · 正在生成结果…"
      : phase === "done" ? " · 完成"
      : phase === "fail" ? " · 失败（已跳过）"
      : phase === "stop" ? " · 已停止"
      : " · 思考中…";
    blk.tag.textContent = reasonLabel(key) + tail;
  };
  // 收尾复位：流被中断（用户停止 / 重新转换接管）或失败时，思考块标签与 LLM 计时行都不会再收到后续事件
  // （单路终态靠 llm-job-done、整体成功靠 parse 折叠），掐流后这两类事件永远不到 ⇒ 面板上「已停止转换」
  // 旁边一直挂着「思考中…」「调用 LLM 中…6.0s」，用户看到的就是"停了还在转"。收尾必须由这里统一置终态。
  // 只改已有块：getReasonBlock 是纯查询，绝不在此新建块（收尾时凭空多出块比残留更糟）。
  const settleStream = (phase) => {   // phase: "stop"（用户停止）| "fail"（转换失败）
    // 只有显示区归属本任务流才动面板：锚点（data-line）按 jobKey 命名，失焦任务的同名块在别的面板里，
    // 不判归属会串改别人正在跑的那几块。
    if (focusedRunId !== runId) return;
    const tail = phase === "stop" ? " · 已停止" : " · 失败（已跳过）";
    for (const key of Object.keys(reasonPhase)) {
      if (reasonPhase[key] === "done" || reasonPhase[key] === "fail") continue;
      reasonPhase[key] = phase;
      const blk = getReasonBlock(key);
      if (blk && blk.tag) blk.tag.textContent = reasonLabel(key) + tail;
    }
    if (llmLine) {
      const l = liveOwnLine(llmLine, "llm") || liveLine(llmLine, "llm");
      if (l) {
        l.classList.remove("cur");
        l.classList.add(phase === "stop" ? "wait" : "err");
        setLine(l, phase === "stop" ? "■" : "✗",
          (phase === "stop" ? "调用 LLM 已停止（" : "调用 LLM 失败（") + llmTaskTxt + "）");
      }
      llmLine = null;
    }
  };
  // 思考流落盘：把缓冲一次性写入 DOM 并跟随滚动（每帧最多一次，替代逐 chunk 写入）
  const flushReason = () => {
    if (reasonRaf) { cancelAnimationFrame(reasonRaf); reasonRaf = 0; }
    // 本流已不是显示区归属（点了「文档排版」新建任务 / 打开别的任务）→ 丢弃缓冲：
    // 缓冲里的文本属于上一轮的思考流，落进当前面板就是串写（同名 jobKey 的块会命中新任务的块）。
    if (focusedRunId !== runId) { reasonBufs = {}; return; }
    const keys = Object.keys(reasonBufs);
    if (!keys.length) return;
    let wrote = false;
    for (const k of keys) {
      const txt = reasonBufs[k];
      delete reasonBufs[k];
      if (!txt) continue;
      const blk = getReasonBlock(k);
      if (!blk) continue;
      blk.el.textContent += txt;
      blk.el.scrollTop = blk.el.scrollHeight; // 内部思考流跟随到最新
      wrote = true;
    }
    if (!wrote) return;
    scrollConvBottom();                               // 外层窗口跟随
    // 快照节流：思考流每 500ms 才落一次快照（切走再切回靠它恢复）。
    // 逐块快照会序列化整个面板（O(n²)），正是主线程卡死的元凶，必须节流。
    const now = Date.now();
    if (now - lastReasonSnap >= 500) { lastReasonSnap = now; snapshotConvState(); }
  };
  // ══ 思考流渲染 END（test/test-reason-throttle.js 抽取本段真实实现执行，勿改标记）══

  const onEvent = (ev, data) => {
    // 并行模型：完成/错误事件无条件收（收尾据此构造快照/判断失败）；
    // 过程渲染仅「聚焦任务」执行——非聚焦任务静默跑完，完成后由收尾构造快照保存任务。
    if (ev === "done") { flushReason(); doneData = data; return; }   // 收尾补写思考流缓冲
    if (ev === "error" && focusedRunId !== runId) {
      streamError = (data && (data.message || data.reason)) || "转换失败";
      streamFriendly = (data && data.friendly) || null;
      return;
    }
    if (focusedRunId !== runId) return;   // 非聚焦：跳过过程渲染（不污染共享显示区）
    if (ev === "phase") {
      const st = data.stage;
      const fi = (typeof data.fileIndex === "number") ? data.fileIndex : 0;
      const gkey = fi + "|" + (data.tplName || "");
      if (st === "read-start") {
        const src = data.source === "local" ? "本地文件" : "在线文档";
        // 读取中不铺源标识：在线文档是动辄几百字符的 URL，截断后仍是噪音；标题等读取完成再出。
        // 多源时带序号（#2 #3）区分同型行，单源（fi=0）不带
        const idx = fi > 0 ? " #" + (fi + 1) : "";
        const label = "读取中（" + src + idx + "）…";
        const line = appendLine("◌", label, "read");
        line.setAttribute("data-runid", runId);   // 行归属标记：上一任务残留的刷新计时器不得改写本任务的行
        const t0 = Date.now();
        const id = setInterval(() => {
          // 行不在当前面板里（本任务失焦 / 面板被新任务清空 / 行已归别的任务）→ 本轮不写：
          // 不串写别人的行、也不触发跟随滚动。表留到阶段结束或收尾统一停，切回该任务即可继续走秒。
          const ln = liveOwnLine(readTimer.line, "read");
          if (!ln) return;
          readTimer.line = ln;
          const s = ((Date.now() - t0) / 1000).toFixed(1);
          setLine(ln, "◌", label + " " + s + "s");
        }, 500);
        readTimer = { id, line, src, idx };
      } else if (st === "read") {
        const src = data.source === "local" ? "本地文件" : "在线文档";
        const idx = (readTimer && readTimer.idx) || (fi > 0 ? " #" + (fi + 1) : "");
        // 读取完成即显示源标题/文件名（2026-09-05 用户反馈：不要等到收尾统计才出现）
        // 本地文件无 title，input 是上传落盘/归档后的路径 → 换回原始文件名：
        // displayName（服务端归档时下发）> localDisplayName（前端暂存路径映射）> 原样路径。
        // ⚠ 归档路径必须靠 displayName 兜：前端映射表只认 temp\drag\ 暂存路径，归档后失配会漏出任务目录（2026-09-09）。
        const inp = (data.title || data.displayName || (data.source === "local" ? localDisplayName(data.input) : data.input) || "").toString();
        const short = inp.length > 48 ? (inp.slice(0, 46) + "…") : inp;
        const charCount = data.charCount || 0;
        readCharCount = charCount;
        const sec = ((data.ms || 0) / 1000).toFixed(1);
        const text = "读取完成（" + src + idx + "）：" + esc(short) + "（" + charCount + " 字 · " + sec + "s）";
        if (readTimer) { clearInterval(readTimer.id); readTimer.line = liveLine(readTimer.line, "read") || readTimer.line; readTimer.line.classList.remove("cur"); readTimer.line.classList.add("ok"); setLine(readTimer.line, "✓", text); readTimer = null; }
        else appendLine("✓", text);
      } else if (st === "parse") {
        flushReason();   // 折叠摘要要读 textContent，先把缓冲补写（否则末尾思考被截）
        // LLM 推理刚结束（parse 紧随其后），收尾 LLM 计时行
        if (llmTimer) { clearInterval(llmTimer); llmTimer = null; }
        if (llmLine) { llmLine = liveLine(llmLine, "llm") || llmLine; llmLine.classList.remove("cur"); llmLine.classList.add("ok"); setLine(llmLine, "✓", "调用 LLM 完成（" + llmTaskTxt + "）" + (llmStartT ? (" · " + ((Date.now() - llmStartT) / 1000).toFixed(1) + "s") : "")); llmLine = null; }
        // 解析行不再重复字数（读取行已有原始字数、统计卡源行有最终字数，2026-09-05 用户反馈去重）：
        // 只留段落/表格计数 + 差异标注（校对开 → 「校对 Δ」；校对关 → Δ 仅来自字符规范化，标签随之改「规范 Δ」）
        const delta = (readCharCount != null) ? ((data.textChars || 0) - readCharCount) : null;
        const deltaStr = delta === null ? "" : (delta === 0
          ? (reviewRun ? "（校对无改动）" : "")
          : "（" + (reviewRun ? "校对" : "规范") + " Δ" + (delta > 0 ? "+" : "") + delta + " 字）");
        appendLine("◆", "解析完成：" + data.paraCount + " 段落 / " + data.tableCount + " 表格" + deltaStr);
        // 思考流默认折叠：完成后改为单行摘要，点击展开/收起（三路并行各一块，不再混写同一容器）
        const procHost = document.getElementById("conv-process");
        if (procHost) {
          for (const key of Object.keys(reasonPhase)) {
            const blk = getReasonBlock(key);
            if (!blk) continue;
            // 半块兜底：el / tag 任一缺失都由上游可能发生（面板被重写只保住一半）——
            // 这里直接跳过，不假设两者都在（否则 .textContent 取到 null 会整段崩掉收尾）
            if (!blk.el) {
              if (blk.tag && blk.tag.parentNode) blk.tag.parentNode.removeChild(blk.tag);
              delete reasonBlocks[key];
              continue;
            }
            const full = blk.el.textContent || "";
            if (!full) {   // 该阶段没有思考内容（非推理模型）→ 不留空块
              if (blk.tag && blk.tag.parentNode) blk.tag.parentNode.removeChild(blk.tag);
              if (blk.el.parentNode) blk.el.parentNode.removeChild(blk.el);
              delete reasonBlocks[key];
              continue;
            }
            blk.el.dataset.fullText = full;
            // 思考块落块即收起：定调是不展示思考链原文（推理模型输出为英文，展开即刷屏），
            // 块头只给「思考过程 · N 字」实时字数，用户想查看再点开。
            blk.el.classList.add("is-collapsed");
            blk.el.textContent = "▸ 思考过程 · " + full.length + " 字（点击展开）";
            if (blk.tag) blk.tag.textContent = reasonLabel(key) + " · " + (reasonPhase[key] === "fail" ? "失败（已跳过）" : "完成") + "，点击展开";
          }
          rebindReasonToggle(procHost);   // 阶段标签 + 思考块都绑上展开/收起
        }
        snapshotConvState();
      } else if (st === "llm") {
        llmLine = appendLine("◌", "调用 LLM 中（" + llmTaskTxt + "）…", "llm");
        llmLine.setAttribute("data-runid", runId);   // 行归属标记：旧任务残留的计时器盖不到本行（曾致提示文案来回闪）
        // 实时计时：覆盖思考 + 出结果整个阶段；setInterval 走闭包引用 llmLine，
        // 切走再切回后 liveLine 按锚点重定位，下一帧就接着写
        llmStartT = Date.now();
        if (llmTimer) clearInterval(llmTimer);
        llmTimer = setInterval(() => {
          // 行不在当前面板里（本任务失焦 / 面板被新任务清空 / 行已归别的任务）→ 本轮不写
          const l = liveOwnLine(llmLine, "llm");
          if (!l) return;
          llmLine = l;
          if (!llmStartT) return;
          const s = ((Date.now() - llmStartT) / 1000).toFixed(1);
          const ph = l.dataset.llmPhase === "gen" ? "生成结果中" : "调用 LLM 中";
          setLine(l, "◌", ph + "（" + llmTaskTxt + "）… " + s + "s");
        }, 500);
      } else if (st === "llm-content") {
        // LLM 思考结束、开始输出结果：状态可见化（这段原先界面全无变化，容易被当成又卡住）
        const rk = reasonKeyOf(data);
        if (!rk) return;   // 内部环节（表格识别）进的是别的 job：不代表主线那几路出了结果，不切文案
        setReasonPhase(rk, "gen");
        const line = liveLine(llmLine, "llm");
        if (line) {
          llmLine = line;
          if (line.dataset.llmPhase !== "gen") {   // 三路并行只切一次
            line.dataset.llmPhase = "gen";
            setLine(line, "◌", "生成结果中（" + llmTaskTxt + "）… " + ((Date.now() - llmStartT) / 1000).toFixed(1) + "s");
          }
        }
      } else if (st === "llm-job-done") {
        // 单路任务结束（成功/失败）：对应块的标签改为完成/失败，用户能看出还有几路在跑
        setReasonPhase(reasonKeyOf(data), (data && data.ok) ? "done" : "fail");
        snapshotConvState();
      } else if (st === "gen-start") {
        const line = appendLine("◌", "生成中「" + esc(data.tplName) + "」…", "gen|" + gkey);
        line.setAttribute("data-runid", runId);   // 行归属标记（同 read/llm 行）
        const t0 = Date.now();
        const id = setInterval(() => {
          const gt = genTimers.find(x => x.gkey === gkey);
          if (!gt) return;
          // 行不在当前面板里（本任务失焦 / 面板被新任务清空 / 行已归别的任务）→ 本轮不写
          const ln = liveOwnLine(gt.line, "gen|" + gkey);
          if (!ln) return;
          gt.line = ln;
          const s = ((Date.now() - t0) / 1000).toFixed(1);
          setLine(ln, "◌", "生成中「" + esc(data.tplName) + "」… " + s + "s");
        }, 500);
        genTimers.push({ id, gkey, line });
      } else if (st === "gen") {
        const gt = genTimers.find(x => x.gkey === gkey);
        const text = "套用「" + esc(data.tplName) + "」→ " + esc(data.fileName) + "（" + formatBytes(data.bytes) + " · " + ((data.ms || 0) / 1000).toFixed(1) + "s）";
        if (gt) {
          clearInterval(gt.id);
          gt.line = liveLine(gt.line, "gen|" + gkey) || gt.line;
          gt.line.classList.remove("cur"); gt.line.classList.add("ok"); setLine(gt.line, "✓", text);
        }
        else appendLine("✓", text);
      }
    } else if (ev === "reasoning") {
      // LLM 思考流：缓冲 + 每帧批量落盘（推理模型 reasoning_content）
      // 关键：逐 chunk 直写 DOM 且每块都 snapshotConvState() 会序列化整个面板（O(n²)），
      // 思考流变长后主线程被占满 → 界面卡死、点菜单无响应。改为纯字符串缓冲 + rAF 单次写入。
      const key = reasonKeyOf(data);
      const txt = (data && data.text) || "";
      if (key && txt) {
        // 首块思考流 → 整块新建（此前用 ensureReasonBlock 会两条腿都不走：无块可补 → 返回 null，
        // 于是 ensureReasonBlock 里的 reasonPhase 初值也没设上，块凭空消失、思考流无处可落）
        if (!ensureReasonBlock(key)) createReasonBlock(key);
        reasonBufs[key] = (reasonBufs[key] || "") + txt;
        if (!reasonRaf) {
          reasonRaf = requestAnimationFrame(() => { reasonRaf = 0; flushReason(); });
        }
      }
    } else if (ev === "reasoning-reset") {
      // 服务端护栏掐断（模型输出退化 / 超过总时长上限）：本轮已回推的思考作废 —— 连缓冲带 DOM 一起丢弃，
      // 由重发的那次思考流重新填充。不丢的后果：退化时刷屏的大段垃圾留在面板上，重发内容接在它后面，
      // 用户看到的是「一直刷的鬼东西 + 正常内容」，无从分辨哪次有效。
      const key = reasonKeyOf(data);
      if (key) {
        reasonBufs[key] = "";                       // 先清未落盘的缓冲（rAF 排队中的那份）
        const blk = getReasonBlock(key);
        if (blk && blk.el) blk.el.textContent = "";
        setReasonPhase(key, "think");               // 标签退回「思考中…」（重发正在跑）
      }
    } else if (ev === "done") {
      doneData = data;
    } else if (ev === "error") {
      flushReason();   // 先补写思考流，保证错误行排在思考内容之后（顺序不乱）
      // 不再抛错（避免被 SSE 底层的"坏帧忽略"静默吞掉），改记到变量，收尾处统一处理。
      // 2026-09-13：错误行**只由收尾处（下方 catch 的 showErr）写一次** —— 这里再写一遍会和收尾那次
      // 叠成两行一模一样的「转换失败：…」（用户看到会以为失败了两次）。这里只做状态标记：
      if (readTimer) { clearInterval(readTimer.id); readTimer.line = liveLine(readTimer.line, "read") || readTimer.line; readTimer.line.classList.remove("cur"); readTimer.line.classList.add("err"); readTimer = null; }
      if (llmTimer) { clearInterval(llmTimer); llmTimer = null; }
      if (llmLine) { llmLine = liveLine(llmLine, "llm") || llmLine; llmLine.classList.remove("cur"); llmLine.classList.add("err"); }
      streamError = data.message || "转换失败";
      streamFriendly = data.friendly || null;
    }
  };

  const runStream = async (inputs) => {
    // 转换第 3 步：继承白名单（源文档有价值属性；未勾选 = 空数组 → 后端回退模板集/默认）
    // 用 doConvert 顶部快照的 inheritAttrsRun，避免 runStream 内重复读 DOM 出现"转换中用户改勾选导致前后不一致"
    const inheritAttrs = inheritAttrsRun;
    const resp = await fetch("/api/convert/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ inputs, templateSet: tsFiltered, revisions: "keep", inheritAttrs, recognizeLeadParagraph: recognizeLeadRun, reviewEnabled: reviewRun, nameCheck: nameCheckRun, nameRosters: nameRostersRun, taskId }),
      // 并行模型：发起新任务/查看历史不中断本任务流；仅「重新转换当前任务」时 abort 本任务的 ac
      signal: task.ac ? task.ac.signal : undefined,
    });
    if (!resp.ok) {
      let msg = "请求失败(" + resp.status + ")";
      try { const j = await resp.json(); if (j && j.error) msg = j.error; } catch (e) {}
      throw new Error(msg);
    }
    if (!resp.body) throw new Error("当前环境不支持流式读取");
    await sseReadStream(resp, (ev, data) => { onEvent(ev, data); });
  };

  const stopTimers = () => {
    flushReason();   // 兜底：流结束时把思考流缓冲补写，避免最后一块丢失
    if (readTimer) { clearInterval(readTimer.id); readTimer = null; }
    if (llmTimer) { clearInterval(llmTimer); llmTimer = null; }
    genTimers.forEach(x => clearInterval(x.id)); genTimers = [];
  };

  try {
    await runStream(built.inputs);
    stopTimers();
    if (streamError) throw new Error(streamError);
  } catch (e) {
    stopTimers();
    // 被「重新转换当前任务」主动中断（AbortError）：静默退出（新任务已发起，不提示、不污染显示区）
    if (e && e.name === "AbortError") {
      task.running = false;
      task.done = true;
      // 用户点「停止」中断：面板交代结果，不静默
      if (task.stoppedByUser) {
        settleStream("stop");   // 先把思考块/LLM 行置终态，再写结果行（否则终态行之上还挂着"思考中…"）
        if (focusedRunId === runId) appendLine("■", "已停止转换（未生成结果）");
        else toast("任务「" + (task.input && task.input.label || "转换") + "」已停止", "warn");
      }
      if (convState === task) { convState.running = false; convState.completedAway = false; }
      updateReconvertBtn();
      renderTasks();   // 任务结束 → 列表「转换中」徽标消退
      return;
    }
    // 并行模型：本任务失败不影响其他任务；聚焦才写显示区，非聚焦 toast 透明提示
    // 有 friendly 就渲染三段式块（技术原文折叠）；没有则原样单行（后端认不出这类错误）
    const showErr = (msg) => {
      settleStream("fail");   // 同上：失败时思考块/LLM 行同样不再有事件驱动，先置终态再写错误块
      const t = (streamFriendly && streamFriendly.title) || msg;
      if (focusedRunId === runId) appendFailBlock(streamFriendly, msg);
      else toast("任务「" + (task.input && task.input.label || "转换") + "」失败：" + t, "err");
    };
    // 优先用 SSE error 事件带回的后端文案（唯一一次显示点）
    showErr(streamError || e.message || "未知错误");
    task.running = false;
    task.done = true;
    if (convState === task) { convState.running = false; convState.completedAway = (currentView !== "convert"); }
    // 失败 = 转换结束：第④步「开始转换」保持消失，右上角「重新转换」显示（重试/重跑）
    updateReconvertBtn();
    renderTasks();   // 任务结束 → 列表「转换中」徽标消退（占位记录保留，可点开重新转换）
    return;
  }
  // ── 收尾：构造快照（过程+统计+结果）→ 无条件保存任务 → 聚焦才渲染显示区 ──
  const d = doneData;
  task.running = false;
  task.done = true;
  if (convState === task) {
    convState.running = false;
    convState.completedAway = (currentView !== "convert");
  }
  // 转换结束（完成/异常）：第④步「开始转换」保持消失，右上角「重新转换」显示（重跑）
  updateReconvertBtn();
  // done 事件缺失等边缘情况：此处兜底刷新一次，徽标不残留（正常路径 upsertTask 内部已刷新，幂等）
  renderTasks();
  if (d) {
    const sec = (d.elapsedMs / 1000).toFixed(1);
    // 源行只留标题+最终字数（2026-09-06 拍板：单源不在统计卡展示；仅多源时列出，供区分各源）：
    // 段落/表格已在「解析完成」行、原始字数已在「读取完成」行，不再重复（2026-09-05 用户反馈）
    const srcLines = (d.sources || []).length > 1 ? (d.sources || []).map(s =>
      `<div class="cs-sub">源（${s.source === "local" ? "本地" : "在线"}）：<b>${esc(s.title || (s.source === "local" ? localDisplayName(s.input) : s.input))}</b> · ${s.charCount} 字</div>`
    ).join("") : "";
    const statKeys = Object.keys(d.stats || {}).filter(k => (d.stats[k] || 0) > 0);
    // title（外部标题计数）与模板样式「标题」同名，显示为「文档标题」消歧（2026-09-05 用户反馈：样式计数不重名）
    const statName = (k) => (k === "title" && d.stats["标题"]) ? "文档标题" : (STYLE_NAME_FALLBACK[k] || k);
    const statLine = statKeys.length ? statKeys.map(k => statName(k) + "×" + d.stats[k]).join(" · ") : "—";
    // 格子按需出现：耗时已在「全部完成」行（不重复），修订/高亮为 0 时不占位（2026-09-05 用户反馈）
    const csCell = (k, v) => `<div class="cs-cell"><span class="cs-k">${k}</span><span class="cs-v">${v}</span></div>`;
    const cells = [csCell("输入源", d.sources.length + " 个"), csCell("生成文件", d.files.length + " 个")];
    if (d.totalEdits) cells.push(csCell("修订（revisions）", d.totalEdits));
    if (d.totalHl) cells.push(csCell("高亮（highlight）", d.totalHl));
    // 校对建议（2026-09-14 用户拍板：一律落成修订，用户逐条接受/拒绝）——正常情况这里没有额外行；
    // 只有真落不了的（原文定位不到 / 无可改建议文本）才提示条数，明细在任务过程记录里（避免建议静默消失）。
    const noteCnt = (d.reviewDropped || 0) + (d.reviewNoop || 0);
    const noteSub = noteCnt
      ? `<div class="cs-sub">校对建议中有 ${noteCnt} 条未能落成修订（原文定位不到 ${d.reviewDropped || 0} · 无可改建议文本 ${d.reviewNoop || 0}），详见任务过程记录</div>`
      : "";
    // 人员姓名校对（2026-09-17 用户定调）：**只报统计、不列具体姓名** —— 具体哪个名字错了，
    // 到 Word 的修订列表里逐条看最直观，结果卡不重复列一遍。why 非空 = 勾了但没跑，显示原因而非一个空的 0。
    const nc = d.nameCheck;
    const ncSub = nc
      ? (nc.why
        ? `<div class="cs-sub" style="color:var(--warn)">人员姓名校对未执行：${esc(nc.why)}</div>`
        : `<div class="cs-sub">人员姓名校对：已核对名单 ${nc.rosterCount || 0} 人，其中 ${nc.mentioned || 0} 人在正文中出现，${nc.applied || 0} 处写法有误（已在文中标出，见修订）</div>`)
      : "";
    const statsHtml = `
      <div class="conv-stats">
        <div class="cs-title">转换统计</div>
        <div class="cs-grid">${cells.join("")}</div>
        ${srcLines}
        <div class="cs-sub">各样式计数：${statLine}</div>
        ${noteSub}
        ${ncSub}
        ${d.degraded ? '<div class="cs-sub" style="color:var(--warn)">⚠ 部分内容经降级处理</div>' : ""}
        ${d.incomplete ? '<div class="cs-sub" style="color:var(--warn)">⚠ 模型输出末尾被截断，已按可用部分生成 —— 请核对文末内容是否完整（上次已自动重试一次，仍缺则建议缩小输入后重转）</div>' : ""}
      </div>`;
    const resultHtml = d.files.map(f => `
      <div class="result-item">
        <span class="result-name">${esc(f.fileName)}</span>
        <span class="uf-size">${formatBytes(f.bytes)}</span>
        <button type="button" class="open-btn" data-download="${esc(f.rel || "")}"${f.rel ? "" : " disabled"}>下载</button>
      </div>`).join("");
    // 快照：过程行（聚焦期间实时累积于 task.procHtml）+ 统计块；结果单独存
    const savedProcess = {
      procHtml: (task.procHtml || "") + statsHtml,
      resultHtml,
      savedAt: Date.now(),
    };
    // 任务保存无条件：并行下每个完成任务都记录，从左侧历史任务可回显/重新转换
    // N7：新建记录沿用本次转换的 taskId（与产物目录同名）；同文档已存在记录则保持旧 id（URL 去重语义不变）
    // 拖拽源随任务归档（2026-09-06 拍板）：服务端把 temp\drag\ 源复制进 Tasks\<taskId>\，
    // sources[i].input 即归档路径——回写输入快照，任务记录与「重新转换」不再依赖重启即清空的临时目录
    if (built.inputType === "local" && Array.isArray(d.sources)) {
      built.files = built.files.map((f, i) => {
        const s = d.sources[i];
        return (s && s.source === "local" && s.input && s.input !== f.path) ? { ...f, path: s.input } : f;
      });
    }
    upsertTask({
      taskId,
      title: d.title || built.label, url: built.label, inputType: built.inputType, files: built.files,
      tsFile: task.tsFile, tsName: task.tsName, checked: task.checked || [], checkedNames: task.checkedNames || [],
      inheritAttrs: inheritAttrsRun.slice(), recognizeLeadParagraph: recognizeLeadRun, reviewEnabled: reviewRun,
      nameCheck: nameCheckRun, nameRosters: nameRostersRun.slice(),
      process: savedProcess, telemetry: d.telemetry
    });
    // 聚焦才渲染共享显示区（非聚焦任务：任务已保存，用户可从左侧历史任务查看）
    if (focusedRunId === runId) {
      appendLine("✓", "全部完成：" + d.files.length + " 个文件 · " + sec + "s");
      procEl.insertAdjacentHTML("beforeend", statsHtml);
      scrollConvBottom();
      rl.innerHTML = resultHtml;
      upgradeResultItems(rl);   // 实时新结果已是新结构，函数幂等跳过
      const resultPhase = document.getElementById("result-phase");
      if (resultPhase) resultPhase.style.display = d.files.length ? "" : "none";
      // 统计块与结果是在 appendLine 内部快照之后才插入 DOM 的 → 必须补一次快照，
      // 否则切视图/切任务用旧快照恢复时「转换统计 + ⑤输出结果」整体丢失（2026-09-10 修复）
      snapshotConvState();
      rb.scrollIntoView({ behavior: "smooth", block: "nearest" });
    } else {
      // 非聚焦完成：显示区不属于本任务，DOM 快照路径走不到 → 直接回写任务快照，
      // 否则点开该历史任务时 liveRun 分支优先于已保存的任务记录，结果卡同样丢失
      task.procHtml = savedProcess.procHtml;
      task.resultHtml = resultHtml;
    }
  }
}

// ═══════════════════════════════════════════
// 登录流程已移除（本版只处理本地文档上传）。以下保留空实现避免历史引用报错。
// ═══════════════════════════════════════════
// 需要认证的错误（登录已移除，恒 false；错误走普通分支）
function needsAuth(errMsg) {
  return false;
}

// 确保已登录：已移除，直接返回 false（不再打开授权窗）。
// pending 落盘/恢复机制保留，但不再用于在线文档链接。
const PENDING_CONVERT_KEY = "klunl_pending_convert";
function savePendingConvert(p) {
  try { localStorage.setItem(PENDING_CONVERT_KEY, JSON.stringify(Object.assign({ savedAt: Date.now() }, p || {}))); } catch (e) { /* 忽略 */ }
}
function loadPendingConvert() {
  try {
    const v = JSON.parse(localStorage.getItem(PENDING_CONVERT_KEY) || "null");
    if (!v || !Array.isArray(v.inputs) || !v.inputs.length) return null;
    return v;
  } catch (e) { return null; }
}
function clearPendingConvert() {
  try { localStorage.removeItem(PENDING_CONVERT_KEY); } catch (e) { /* 忽略 */ }
}
async function doLogin(pending) {
  // 登录功能已移除：直接视为未授权，调用方应给出本地文件提示。
  return false;
}

// 扫码授权中断恢复：登录功能已移除，直接清除遗留的 pending（不再恢复在线文档链接）。
async function maybeResumePendingConvert() {
  const pending = loadPendingConvert();
  if (!pending) return;
  clearPendingConvert();   // 本地文件路径已在磁盘，不恢复在线文档链接
}

// ═══════════════════════════════════════════
// 导入 docx → 提取模板
// ═══════════════════════════════════════════
// 导入 docx → 提取模板：渲染（思考流 + 两类结果 + 中文名/单位换算 + 勾选过滤）
// ═══════════════════════════════════════════

// 导入预览的样式命名：来自服务端下发的 styleNames（与引擎种子一致，Word 命名），未知 key 直显。
// 兜底映射：覆盖种子/模板集被改成英文 key 的场景，UI 仍能展示中文名。
const STYLE_NAME_FALLBACK = {
  title: "标题", subtitle: "副标题",
  h1: "标题 1", h2: "标题 2", h3: "标题 3",
  body: "正文段落", lead: "标题引导段落",
  salutation: "称谓", sign: "落款",
  key: "自定义键", tableTitle: "表格标题", tableHeader: "表头行", tableBody: "表格正文",
};
let IMPORT_STYLE_NAMES = {};
let IMPORT_STYLE_CATS = {};
let IMPORT_EXTRACTED_KEYS = new Set();   // 文档实际提取到的样式 key（未提取到的用默认值，前端备注）
let IMPORT_STYLE_GROUPS = {};            // 表格样式 key → 组名（数据表格/版记…），导入结果卡片按组展示
let IMPORT_EXTRA_STYLE_DEFS = [];        // 导入产生的非种子样式定义（如 版记机关/版记日期），保存时并入模板集 styles[]
// 「默认样式」区（2b 未识别默认样式）的默认勾选集（2026-09-04 用户拍板 D1=A+标题引导段落）：
//   「正文」兜底 + 标题骨架（标题/副标题/标题 1~3）+ 段落骨架（正文段落/标题引导段落）。
//   场景件（称谓/落款/表格标题）默认不勾：模板 styles[] 就是 LLM 分类目标清单，目标越多误判率越高；
//   未勾样式在转换时降级到正文段落，内容不丢，成本远低于误判。
//   英文别名为种子/模板集被改成英文 key 的场景兜底（与 STYLE_NAME_FALLBACK 同源）。
const IMPORT_OPT_DEFAULT_KEYS = [
  "正文", "正文段落", "标题", "副标题", "标题 1", "标题 2", "标题 3", "标题引导段落",
  "body", "title", "subtitle", "h1", "h2", "h3", "lead",
];
let IMPORT_SAMPLES = {};                 // 提取分组的第一样本原文（示例列展示，与 LLM 原始分类解耦）
// v1.33 导入结果页重设计（以识别出的样式为主线）：
let IMPORT_STYLE_GROUPS_DETAIL = {};     // 识别样式 key → { paras: [ {text, font, size, align, line, bold, indentChars, spaceBefore, spaceAfter, color} ] }（2a 展开段落修正用）
let IMPORT_UNCLAIMED = [];               // 未归属段落（被 openingEnd 跳过的开头段：红头/文号等，2c 区可提取为样式）
let IMPORT_REAL_TPL = null;              // 当前导入的真实模板（编辑态派生 / 保存基准）
let IMPORT_EDIT = null;                  // 编辑态（渲染与保存的唯一依据）：
                                         //   rows: [{ key, seed, deleted, expanded, paras }]（2a 识别样式）
                                         //   opt:  { key: { checked } }（2b 未识别默认样式，默认不勾）
                                         //   unclaimed: [ {text, …fmt} ]（2c 剩余未处理段落）
let IMPORT_DESC = "";                    // 模板说明（结果页输入，保存时写入 templates[].description）
function styleNameCn(key) {
  return IMPORT_STYLE_NAMES[key] || STYLE_NAME_FALLBACK[key] || key;
}

function isExtraStyleKey(key) {
  return (IMPORT_EXTRA_STYLE_DEFS || []).some(d => d.key === key);
}
// 是否为种子 text 样式（非本次导入自造/新命名样式）
function isSeedTextKey(key) {
  return IMPORT_STYLE_CATS[key] === "text" && !isExtraStyleKey(key);
}
// 样式默认值（种子 defaults；自造样式无 defaults 时返回空）
function seedDefaultsOf(key) {
  const tpl = IMPORT_REAL_TPL;
  const st = (tpl && Array.isArray(tpl.styles)) ? tpl.styles.find(s => s.key === key) : null;
  return (st && st.defaults) ? st.defaults : {};
}
// 段落原文清洗（去 Word XML 标签字面残留 + 合并空白）
function cleanImportText(s) {
  return (s || "").replace(/<w:\w+(?:\s[^>]*)?\/>/g, "").replace(/<w:\w+[^>]*>[\s\S]*?<\/w:\w+>/g, "").replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}
function truncateText(s, n) {
  return s.length > n ? s.slice(0, n) + "…" : s;
}
// 段落 → 新样式名建议（取原文开头 2~4 个汉字；冲突时加序号）
function suggestStyleName(text, takenKeys) {
  const clean = cleanImportText(text).replace(/^[^\u4e00-\u9fff]+/, "").slice(0, 4);
  let name = clean || "新样式";
  let n = 2;
  while ((takenKeys || []).includes(name)) name = (clean || "新样式") + " " + (n++);
  return name;
}
// 单段格式摘要（展开列表 / 2c 区展示）
function paraSummary(p) {
  const parts = [];
  if (p.font) parts.push(p.font);
  if (p.size != null) parts.push(fontSizeToCn(p.size));
  if (p.align) parts.push(alignToCn(p.align));
  if (p.bold) parts.push("加粗");
  if (p.indentChars != null) parts.push("缩进" + p.indentChars + "字");
  return parts.join(" · ") || "默认格式";
}
// 由段落数组重算样式值（与后端提取同一算法：public/styleAgg.js，单一真相源）
function recomputeRowValue(row) {
  return StyleAgg.elemFromParas(row.paras, seedDefaultsOf(row.key));
}

// 由服务端明细重建编辑态（2a 识别样式 / 2b 未识别默认样式 / 2c 未归属段落）
// reasoning.steps 按文档顺序（每块 {idx, style}），据此把识别样式按「文档首次出现顺序」排序
// （v1.37 用户拍板 D2-B：公文结构序靠文档序自然呈现，如 发文单位→发文字号→标题→…）。
function initImportEdit(realTpl, reasoning) {
  // 种子由 renderImportResultView / restoreImportView 在调用本函数前已预加载
  // （chip 摘要依赖 SEED_STYLE_CACHE，异步不 await 会导致首屏摘要为空 → 需重渲染才出现）
  const rows = [];
  const used = new Set();
  for (const k of Object.keys(IMPORT_STYLE_GROUPS_DETAIL)) {
    const g = IMPORT_STYLE_GROUPS_DETAIL[k];
    if (!g || !g.paras || !g.paras.length) continue;
    rows.push({ key: k, seed: isSeedTextKey(k), deleted: false, checked: true, expanded: false, paras: g.paras.map(p => Object.assign({}, p)) });
    used.add(k);
  }
  // 文档首次出现索引：steps 未命中的 key（如 LLM 自造但 steps 里分类名不同的）保序排后（V8 稳定排序）
  const firstIdx = {};
  if (reasoning && Array.isArray(reasoning.steps)) {
    for (const s of reasoning.steps) {
      if (s && s.style && firstIdx[s.style] === undefined) firstIdx[s.style] = s.idx;
    }
  }
  rows.sort((a, b) => {
    const ia = firstIdx[a.key], ib = firstIdx[b.key];
    if (ia !== undefined && ib !== undefined) return ia - ib;
    if (ia !== undefined) return -1;
    if (ib !== undefined) return 1;
    return 0;
  });
  const opt = {};
  // Normal 兜底样式（matchType none + next 自引用的 text，如「正文」）：默认勾选随模板保存（2026-09-01 用户拍板，
  // 导入模板自包含正文基准样式，做样式降级兜底终点）；其不匹配任何段落，故仅作基准、不参与段落归属。
  // 用 SEED_STYLE_CACHE（renderImportResultView 已 await ensureSeedStyles 预加载），不用 seedStyles（仅设置页打开后才有值）
  const normalFallbackDef = (SEED_STYLE_CACHE || []).find(s => s.category === "text" && s.matchType === "none" && s.next === s.key);
  const normalFallbackKey = normalFallbackDef ? normalFallbackDef.key : null;
  // 勾汇集 = 常量集 ∪ 正文兜底（种子 key 被改成英文时，正文仍能兜住）
  const defaultChecked = new Set(IMPORT_OPT_DEFAULT_KEYS);
  if (normalFallbackKey) defaultChecked.add(normalFallbackKey);
  for (const k of Object.keys(IMPORT_STYLE_CATS)) {
    if (IMPORT_STYLE_CATS[k] !== "text" || isExtraStyleKey(k) || used.has(k)) continue;
    opt[k] = { checked: defaultChecked.has(k) };
  }
  return { rows, opt, unclaimed: (IMPORT_UNCLAIMED || []).map(p => Object.assign({}, p)) };
}

// pt → 中文字号（最接近映射）
function fontSizeToCn(pt) {
  if (pt == null) return "-";
  const map = [[42,"初号"],[36,"小初"],[26,"一号"],[24,"小一"],[22,"二号"],[18,"小二"],
    [16,"三号"],[15,"小三"],[14,"四号"],[12,"小四"],[10.5,"五号"],[9,"小五"],
    [7.5,"六号"],[6.5,"小六"],[5.5,"七号"],[5,"八号"]];
  let best = null, bestDiff = Infinity;
  for (const [p, name] of map) {
    const d = Math.abs(p - pt);
    if (d < bestDiff) { bestDiff = d; best = name; }
  }
  return bestDiff <= 1.5 ? best : (Math.round(pt * 10) / 10) + "pt";
}

function alignToCn(a) {
  return ({ left: "左对齐", right: "右对齐", center: "居中", both: "两端对齐", justify: "两端对齐", outside: "外侧" })[a] || (a || "-");
}

function twipToCm(t) { return t == null ? "-" : String(Math.round((t / 1440 * 2.54) * 10) / 10); }
function twipToMm(t) { return t == null ? "-" : Math.round(t / 1440 * 25.4); }

function paperName(w, h) {
  if (!w || !h) return "未知";
  const set = [[11906,16838,"A4"],[16838,11906,"A4"],[15840,23760,"A3"],[23760,15840,"A3"],
    [12240,15840,"B5"],[15840,12240,"B5"]];
  for (const [a,b,n] of set) {
    if ((Math.abs(w-a)<40 && Math.abs(h-b)<40) || (Math.abs(w-b)<40 && Math.abs(h-a)<40)) return n;
  }
  return (twipToMm(w)/10) + "×" + (twipToMm(h)/10) + " cm";
}

// 模板提取（完成态）：「模板提取」卡片，与「页面设置」「段落样式」并列，默认展开。
// 内容：统计一行。（2026-09-02 用户拍板：原始思考链不进结果页、不进缓存；
// 推理过程仅在活跃态实时滚动展示（特征中文化后为中文白话），完成态只留统计行。）
function buildImportProcessHtml(tpl, reasoning, thinkText, elapsedMs) {
  const r = reasoning || {};
  const par = r.paragraphCount || 0;
  const tbl = r.tableCount || 0;
  // 识别样式数：编辑态行（未删）优先；legacy 缓存无编辑态时按 reasoning.steps 去重样式数兜底
  const styleCount = IMPORT_EDIT
    ? IMPORT_EDIT.rows.filter(row => !row.deleted).length
    : new Set((r.steps || []).map(s => s.style).filter(Boolean)).size;
  const uncCount = (IMPORT_UNCLAIMED || []).length;
  // 与转换页「转换统计」卡同构（.conv-stats 零新增 CSS）；0 值不占位（表格/未归属按需出现，
  // 2026-09-05 转换页同规则）；单源不设源行（文件名已在面板顶栏，2026-09-06 拍板）
  const cell = (k, v) => `<div class="cs-cell"><span class="cs-k">${k}</span><span class="cs-v">${v}</span></div>`;
  const cells = [cell("解析段落", par), cell("识别样式", styleCount)];
  if (tbl > 0) cells.push(cell("表格", tbl));
  if (uncCount > 0) cells.push(cell("未归属段落", uncCount));
  if (elapsedMs) cells.push(cell("LLM 耗时", (elapsedMs / 1000).toFixed(1) + "s"));
  // id="import-process"：与活跃态卡同名，供 renderImportResultView 定位（默认收起卡 + 填充思考折叠摘要）
  return `<details class="imp-card" id="import-process" open>
    <summary>模板提取</summary>
    <div class="imp-card-body">
      <div class="conv-stats">
        <div class="cs-title">模板提取统计</div>
        <div class="cs-grid">${cells.join("")}</div>
      </div>
      ${thinkText ? `<div class="imp-think-fold conv-reason is-collapsed" data-think="1"></div>` : ""}
    </div>
  </details>`;
}

// 示例摘录：优先取「提取分组」的样本原文（跳过开头的红头/文号即使被 LLM 误判也不出现）；
// 对样本原文做防御性清洗，去除被错误编码进 <w:t> 的 Word XML 标签字面量（如 <w:autoSpaceDE/>）。
// 旧缓存无 samples 时回退到 reasoning.steps 按 style key 取首个命中段，剥掉残留 XML 标签字面文本。
// 若清理后仍以 `<` 开头（旧缓存被 80 字符截断在标签中间的残段），视为无意义，返回空。
function exampleForStyle(styleKey, reasoning) {
  const clean = s => (s || "").replace(/<w:\w+(?:\s[^>]*)?\/>/g, "").replace(/<w:\w+[^>]*>[\s\S]*?<\/w:\w+>/g, "").replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
  const raw = IMPORT_SAMPLES && IMPORT_SAMPLES[styleKey];
  if (raw) {
    const cleaned = clean(raw);
    return cleaned;
  }
  const steps = (reasoning && reasoning.steps) || [];
  const hit = steps.find(s => s.style === styleKey && (s.text || "").trim());
  if (!hit) return "";
  const cleaned = clean(hit.text);
  if (cleaned.startsWith("<")) return "";
  return cleaned;
}

// ═══════════════════════════════════════════
// v1.33 结果区重设计：以「识别出的样式」为主线（可修正清单）
//  2a 识别到的样式（默认保留：可改名 / 整行删除 / 展开段落修正）
//  2b 未识别到的默认样式（默认不勾选，勾选以默认值补入模板）
//  2c 未归属段落（被规则跳过的开头段：可单独提取为样式 / 忽略）
//  表格样式 / 页面设置 保持勾选语义
// ═══════════════════════════════════════════

// 页面设置卡（勾选保留：纸张/页边距必留，页眉/页脚/页号可取消）
// v1.38：页眉/页脚三态详情默认展开——run 级展示（文字/字段/制表位/图片缩略图），未设置的态隐藏，多段 ┊ 串一行
function buildPageCardHtml(tpl) {
  const pg = tpl.page || {};
  const paper = paperName(pg.width, pg.height);
  const paperMm = (pg.width && pg.height) ? `（${twipToMm(pg.width)} × ${twipToMm(pg.height)} mm）` : "";
  // 页眉/页脚摘要：2026-09-02 用户拍板
  //   顺序：识别内容（文字/字段/制表位） → 页码 → 文字属性（字体/字号）→ 多态标签
  //   cfg 完全没传过来 → "无"；cfg 存在但 variants 没文字/字段/图形 → "有"
  //   「普通」永不出现（默认态隐含语义），仅 first/even 多态作区分
  // 单态时无多态标签；多态时仅 first/even 作区分
  const hfSummary = (cfg) => {
    if (!cfg) return "无";
    const v = cfg.variants || {};
    const content = collectContentText(v);
    const labels = [];
    if (v.default && v.first) labels.push("首页");
    if (v.default && v.even) labels.push("偶数页");
    return [content].filter(Boolean).join(" · ") || "有";
  };
  // 收集识别内容：文字/字段/制表位三类一并视为"识别内容"（与字体/字号分离）
  // 多态按 default → first → even 串接；空段跳过；多段以 ┊ 串接
  // 2026-09-02 用户拍板：横线/图形只走第二行，第一行只保留文字类（文字/字段/制表位）
  const collectContentText = (v) => {
    if (!v) return "";
    const order = ["default", "first", "even"];
    const parts = [];
    for (const k of order) {
      const paras = v[k];
      if (!paras || !paras.length) continue;
      const segs = [];
      for (const p of paras) {
        const runs = (p.runs || []).filter(r => r.t !== "rule" && r.t !== "image");
        if (!runs.length) continue;
        const txt = runs.map(hfRunView).join("").trim();
        if (txt) segs.push(txt);
      }
      if (segs.length) parts.push(segs.join(" ┊ "));
    }
    return parts.join(" ┊ ");
  };
  // 文字属性（字体/字号）：cfg 任一态有非图形 run 时显示；优先 default，多态时按 default/first/even 任一可用
  // 返回「字体 字号」字符串；多态时仅取代表性那一份（避免长尾）
  const collectProp = (cfg) => {
    if (!cfg || !cfg.variants) return "";
    const order = ["default", "first", "even"];
    for (const k of order) {
      const paras = cfg.variants[k];
      if (!paras || !paras.length) continue;
      for (const p of paras) {
        for (const r of (p.runs || [])) {
          // 跳过纯图形 run
          if (r.t === "rule" || r.t === "image") continue;
          const parts = [];
          if (r.font) parts.push(esc(r.font));
          if (r.size) parts.push(`${fontSizeToCn(r.size)}(${r.size}pt)`);
          if (parts.length) return ` · ${parts.join(" ")}`;
        }
      }
    }
    return "";
  };
    // 单条 run 视图：纯文字描述（2026-09-02 用户拍板：所有 run 一律正常文字，不做特殊修饰）
  // 字段/制表位/横线/图片统一走「特征描述 + 颜色/字号/文件名/尺寸」行文
  const hfRunView = (r) => {
    if (r.t === "text") return esc(r.text || "");
    if (r.t === "page") {
      const total = r.kind === "total";
      return total ? "总页数字段" : "当前页字段";
    }
    if (r.t === "tab") return "制表位";
    if (r.t === "rule") {
      const c = (r.color || "FF0000").toUpperCase();
      const w = r.weight != null ? r.weight : 2;
      const ls = r.lineStyle || "thickThin";
      // 线型文案：只描述**真实视觉**（方向已由 Word 实测校准，见 hfRunHtml 注释），不用「武文线/文武线」等公文术语
      const styleMap = { thickThin: "双线（上细下粗）", thinThick: "双线（上粗下细）", single: "单线", double: "双线", triple: "三线" };
      const cn = styleMap[ls] || ls;
      return `横线（${cn} · 粗 ${w}pt · 颜色 ${c}）`;
    }
    if (r.t === "image") {
      const fname = r.fileName || "image";
      return `图片（${esc(fname)} · ${r.width}×${r.height}px）`;
    }
    return "";
  };
  // 段 body：r.t==="rule" 时该段算"有内容"（不再被「纯下框线场景」过滤吞掉）
  const runHasContent = (r) => {
    if (r.t === "text") return !!(r.text || "").trim();
    return r.t === "page" || r.t === "rule" || r.t === "image";
  };
  // 文字类 run（与字体/字号同级——「识别文字·字体/字号」属同源特征，2026-09-02 用户拍板合并为同一行）
  const isTextual = (r) => r.t === "text" || r.t === "page" || r.t === "tab";
  // 图形类 run（横线 / 图片 → 单列第二行，文字不做特殊修饰）
  const isGraphic = (r) => r.t === "rule" || r.t === "image";
  // 三态详情：v1.38 拆分文字组与图形组
  //   文字组 → 已并入 hfSummary（识别文字 + 字体/字号 + 多态标签）同一行
  //   图形组（横线/图片）→ 第二行详情 pv-graphic 块（56px 缩进）
  // 单态时无标签；多态时 default 默认不加「普通」前缀，仅 first/even 作区分
  const hfGraphicRows = (cfg) => {
    if (!cfg || !cfg.variants) return "";
    const labels = { first: "首页", even: "偶数" };
    const rows = [];
    for (const k of ["default", "first", "even"]) {
      const paras = cfg.variants[k];
      if (!paras || !paras.length) continue;
      const showable = paras.filter(p => (p.runs || []).some(r => r.t === "rule" || r.t === "image"));
      if (!showable.length) continue;
      const bodies = showable.map(p => (p.runs || []).filter(r => r.t === "rule" || r.t === "image").map(hfRunView).join("")).filter(s => s.length);
      if (!bodies.length) continue;
      rows.push({ k, body: bodies.join(" ┊ ") });
    }
    if (!rows.length) return "";
    const single = rows.length === 1;
    const html = rows.map(r => {
      const lab = single || r.k === "default" ? "" : `<span class="pv-var-label">${labels[r.k]}</span>`;
      return `<div class="pv-var">${lab}<span class="pv-var-content">${r.body}</span></div>`;
    }).join("");
    return `<div class="pv-detail pv-graphic">${html}</div>`;
  };
  const hdr = hfSummary(pg.header);
  const ftr = hfSummary(pg.footer);
  const hdrGraphic = hfGraphicRows(pg.header);
  const ftrGraphic = hfGraphicRows(pg.footer);
  // 文字属性（字体/字号）：多态优先 default 的字体字号；任何 cfg 存在都可能显示（字段也走"识别内容"路径，属性是另一个维度）
  // 2026-09-02 用户拍板：属性放页码之后
  const hdrProp = collectProp(pg.header);
  const ftrProp = collectProp(pg.footer);
  // 页码属性 inline：v1.38 用户拍板「彻底合并」
  // 2026-09-02 用户拍板：format="none" → 显示"无页码"；有页码时显示对齐+格式+首页变体
  // 文字不做特殊修饰，与"识别内容/特征"走一致普通字体
  const PN_FMT_DEMO = { plain: "1", pageOf: "第 1 页 共 1 页", dash: "— 1 —", wave: "～ 1 ～" };
  const pnFmt = pg.pageNumber && (pg.pageNumber.format !== undefined ? pg.pageNumber.format : (pg.pageNumber.default || {}).format) || "";
  const pnAlign = pg.pageNumber && (pg.pageNumber.align !== undefined ? pg.pageNumber.align : (pg.pageNumber.default || {}).align);
  const fmtDemo = PN_FMT_DEMO[pnFmt] || pnFmt || "";
  const pgnTxt = pg.pageNumber
    ? (pnFmt === "none"
        ? `无页码`
        : `有页码 · ${alignToCn(pnAlign)}${fmtDemo ? ` · ${fmtDemo}` : ""}${pg.pageNumber.first === false ? " · 首页无页码" : ""}`)
    : "";
  // 页码仅贴页脚（页眉不放——常识）；无页脚时整段消失
  const pgnSuffix = (pg.pageNumber && pg.footer) ? ` · ${pgnTxt}` : "";
  const pgnInFooter = pgnSuffix;
  const pgnInHeader = "";
  // 单双页 / 首页不同：仅在 cfg 存在时附加（cfg 不存在时只显示"无"，不污染）
  const oddEvenTxt = (pg.oddEven && pg.header) ? " · 单双页" : "";
  const oddEvenFtrTxt = (pg.oddEven && pg.footer) ? " · 单双页" : "";
  const firstTxt = (pg.titlePage || (pg.pageNumber && pg.pageNumber.first === false)) ? " · 首页不同" : "";
  // 多态区分标签（first/even 出现时显示）—— 属性后追加
  const hfMultiLabels = (cfg) => {
    if (!cfg || !cfg.variants) return "";
    const v = cfg.variants;
    const labels = [];
    if (v.default && v.first) labels.push("首页");
    if (v.default && v.even) labels.push("偶数页");
    return labels.length ? ` · ${labels.join("/")}` : "";
  };
  // 页眉行：内容 · 页码 · 属性 · 多态标签
  const hdrLine = `${hdr}${pgnInHeader}${hdrProp}${hfMultiLabels(pg.header)}${oddEvenTxt}${firstTxt}`;
  const ftrLine = `${ftr}${pgnInFooter}${ftrProp}${hfMultiLabels(pg.footer)}${oddEvenFtrTxt}${firstTxt}`;
  return `<details class="imp-card" open>
    <summary>页面设置</summary>
    <div class="imp-card-body">
      <div class="pv-row"><label class="pv-chk"><input type="checkbox" data-pv="paper" checked> 纸张</label> ${esc(paper)} ${paperMm}</div>
      <div class="pv-row"><label class="pv-chk"><input type="checkbox" data-pv="margin" checked> 页边距</label> 上 ${twipToCm(pg.marginTop)} / 下 ${twipToCm(pg.marginBottom)} / 左 ${twipToCm(pg.marginLeft)} / 右 ${twipToCm(pg.marginRight)} cm</div>
      <div class="pv-row"><label class="pv-chk"><input type="checkbox" data-pv="header" ${pg.header ? "checked" : ""}> 页眉</label> ${hdrLine}</div>
      ${hdrGraphic}
      <div class="pv-row"><label class="pv-chk"><input type="checkbox" data-pv="footer" ${pg.footer ? "checked" : ""}> 页脚</label> ${ftrLine}</div>
      ${ftrGraphic}
    </div>
  </details>`;
}

// 表格样式卡（按组分段 + 勾选；id 供保存读取勾选，与旧版一致）
function buildTableCardHtml(tpl, reasoning) {
  const catOf = k => IMPORT_STYLE_CATS[k] || (seedStyles.find(s => s.key === k) || {}).category || "";
  // 方案A（2026-09-02）：样式枚举来自 styles[]（importDocx 不再写顶层扁平字段）
  const styleKeys = (Array.isArray(tpl.styles) ? tpl.styles.map(s => s.key) : [])
    .filter(k => catOf(k) === "text" || catOf(k) === "table");
  const tableKeys = styleKeys.filter(k => catOf(k) === "table");
  if (!tableKeys.length) return "";
  const rowHtml = k => {
    // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults（服务端导入 doc 已无扁平字段）
    const e = (Array.isArray(tpl.styles) && (tpl.styles.find(x => x.key === k) || {}).defaults) || {};
    const szTxt = e.size != null ? `${fontSizeToCn(e.size)}(${e.size}pt)` : "-";
    const ex = exampleForStyle(k, reasoning);
    const exShow = ex ? truncateText(ex, 32) : "-";
    return `<div class="style-row" data-key="${esc(k)}">
      <input type="checkbox" class="style-chk" data-key="${esc(k)}" checked>
      <span class="s-name" title="${esc(styleNameCn(k))}">${esc(styleNameCn(k))}</span>
      <span class="s-font" title="${esc(e.font || "-")}">${esc(e.font || "-")}</span>
      <span class="s-size">${szTxt}</span>
      <span class="s-align">${alignToCn(e.align) || "-"}</span>
      <span class="s-line">${lineCellText(e)}</span>
      <span class="s-spb">${e.spaceBefore != null ? e.spaceBefore : "-"}</span>
      <span class="s-spa">${e.spaceAfter != null ? e.spaceAfter : "-"}</span>
      <span class="s-indent">${indentCellText(e)}</span>
      <span class="s-bold">${e.bold ? "✓" : "-"}</span>
      <span class="s-color">${e.color ? `<span class="s-color-block" style="background:#${e.color}" title="#${e.color}"></span>` : "-"}</span>
      <span class="s-example" title="${esc(exShow)}">${esc(exShow)}</span>
    </div>`;
  };
  const headHtml = `<div class="style-row style-row-head">
    <span class="s-chk-h"></span>
    <span class="s-name">样式名称</span>
    <span class="s-font">字体</span>
    <span class="s-size">字号</span>
    <span class="s-align">对齐</span>
    <span class="s-line">行距</span>
    <span class="s-spb">段前</span>
    <span class="s-spa">段后</span>
    <span class="s-indent">缩进</span>
    <span class="s-bold">加粗</span>
    <span class="s-color">颜色</span>
    <span class="s-example">示例/备注</span>
  </div>`;
  return `<details class="imp-card" open id="imp-table-card">
    <summary>表格样式</summary>
    <div class="imp-card-body style-table-wrap">
      ${headHtml}
      ${(() => {
        const out = [];
        let lastGroup = null;
        for (const k of tableKeys) {
          const g = IMPORT_STYLE_GROUPS[k] || "数据表格";
          if (g !== lastGroup) { out.push(`<div class="imp-tbl-group">${esc(g)}</div>`); lastGroup = g; }
          out.push(rowHtml(k));
        }
        return out.join("");
      })()}
    </div>
  </details>`;
}

// 旧版段落样式卡（v1 缓存无段落明细时的只读兼容：全样式勾选表）
function buildLegacyParaCardHtml(tpl, reasoning) {
  const catOf = k => IMPORT_STYLE_CATS[k] || (seedStyles.find(s => s.key === k) || {}).category || "";
  // 方案A（2026-09-02）：样式枚举来自 styles[]（importDocx 不再写顶层扁平字段）
  const styleKeys = (Array.isArray(tpl.styles) ? tpl.styles.map(s => s.key) : [])
    .filter(k => catOf(k) === "text" || catOf(k) === "table");
  const paraKeys = styleKeys.filter(k => catOf(k) === "text");
  const rowHtml = k => {
    // 2026-09-02 方案A：样式值唯一真相 = styles[].defaults
    const e = (Array.isArray(tpl.styles) && (tpl.styles.find(x => x.key === k) || {}).defaults) || {};
    const szTxt = e.size != null ? `${fontSizeToCn(e.size)}(${e.size}pt)` : "-";
    const extracted = IMPORT_EXTRACTED_KEYS.has(k);
    const ex = extracted ? exampleForStyle(k, reasoning) : "";
    const exShow = extracted ? truncateText(ex, 32) : "未提取到样式，使用默认样式";
    return `<div class="style-row" data-key="${esc(k)}">
      <input type="checkbox" class="style-chk" data-key="${esc(k)}" checked>
      <span class="s-name" title="${esc(styleNameCn(k))}">${esc(styleNameCn(k))}</span>
      <span class="s-font" title="${esc(e.font || "-")}">${esc(e.font || "-")}</span>
      <span class="s-size">${szTxt}</span>
      <span class="s-align">${alignToCn(e.align) || "-"}</span>
      <span class="s-line">${lineCellText(e)}</span>
      <span class="s-spb">${e.spaceBefore != null ? e.spaceBefore : "-"}</span>
      <span class="s-spa">${e.spaceAfter != null ? e.spaceAfter : "-"}</span>
      <span class="s-indent">${indentCellText(e)}</span>
      <span class="s-bold">${e.bold ? "✓" : "-"}</span>
      <span class="s-color">${e.color ? `<span class="s-color-block" style="background:#${e.color}" title="#${e.color}"></span>` : "-"}</span>
      <span class="s-example" title="${esc(exShow)}">${esc(exShow)}</span>
    </div>`;
  };
  const headHtml = `<div class="style-row style-row-head">
    <span class="s-chk-h"></span>
    <span class="s-name">样式名称</span>
    <span class="s-font">字体</span>
    <span class="s-size">字号</span>
    <span class="s-align">对齐</span>
    <span class="s-line">行距</span>
    <span class="s-spb">段前</span>
    <span class="s-spa">段后</span>
    <span class="s-indent">缩进</span>
    <span class="s-bold">加粗</span>
    <span class="s-color">颜色</span>
    <span class="s-example">示例/备注</span>
  </div>`;
  return `<details class="imp-card" open>
    <summary>段落样式</summary>
    <div class="imp-card-body style-table-wrap">
      ${headHtml}
      ${paraKeys.map(rowHtml).join("")}
    </div>
  </details>`;
}

// ── 2a 识别样式行 ──
// 缩进列文案（v1.37 用户拍板 D3-B：首行缩进/右缩进两列合并为「缩进」单列）：
// 右对齐 + 右缩进 → 「右空 N 字」（落款语义：成文日期右空）；其余 + 首行缩进 → 「首行 N 字」；无缩进 → 「-」。
function indentCellText(e) {
  if (e.align === "right" && e.rightIndentChars) return "右空 " + e.rightIndentChars + "字";
  if (e.indentChars) return "首行 " + e.indentChars + "字";
  return "-";
}
function rowValueCells(eff) {
  const szTxt = eff.size != null ? `${fontSizeToCn(eff.size)}(${eff.size}pt)` : "-";
  return `
    <span class="s-font" title="${esc(eff.font || "-")}">${esc(eff.font || "-")}</span>
    <span class="s-size">${szTxt}</span>
    <span class="s-align">${alignToCn(eff.align) || "-"}</span>
    <span class="s-line">${lineCellText(eff)}</span>
    <span class="s-spb">${eff.spaceBefore != null ? eff.spaceBefore : "-"}</span>
    <span class="s-spa">${eff.spaceAfter != null ? eff.spaceAfter : "-"}</span>
    <span class="s-indent">${indentCellText(eff)}</span>
    <span class="s-bold">${eff.bold ? "✓" : "-"}</span>
    <span class="s-color">${eff.color ? `<span class="s-color-block" style="background:#${eff.color}" title="#${eff.color}"></span>` : "-"}</span>`;
}

// 行距显示（用户口径 2026-08-31）：lineMult（倍数/行）→ "1.0行"；line（固定磅）→ "28磅"；都无 → "-"
function lineMultTxt(m) {
  const v = +m;
  return Number.isInteger(v) ? v.toFixed(1) : String(v);
}
function lineCellText(e) {
  if (e.lineMult != null) return lineMultTxt(e.lineMult) + "行";
  if (e.line != null) return e.line + "磅";
  return "-";
}

// 引导段行判定：种子中 mode:"lead" 的样式（标题引导段落）。改名到其他 key 后自然失效。
function isLeadRowKey(key) {
  const def = seedStyles.find(s => s.key === key);
  return !!(def && def.mode === "lead");
}

function buildParaRowHtml(row, idx) {
  const eff = recomputeRowValue(row);
  const ex = cleanImportText((row.paras[0] && row.paras[0].text) || "");
  const exShow = ex ? truncateText(ex, 32) : "（无段落，将使用默认值）";
  // 引导段（mode=lead）：表格值列体现标题句属性（parseParaXml 对引导段 p.font 取首个 run），
  // 正文部分输出时不改格式、按「正文段落」样式排版 → 加提示说明语义
  const leadHint = isLeadRowKey(row.key)
    ? `<div class="imp-tx-muted" style="padding:2px 8px">引导段：标题句按样式定义的「字体 / 加粗」渲染（以句号或冒号结束），其后的正文不改格式，自动套用「正文段落」样式。</div>`
    : "";
  // v1.37 用户拍板 D1-B：识别行只留勾选（取消勾选=该样式不入模板），删除按钮移除
  const off = row.checked === false ? " opt-off" : "";
  return `
  <div class="style-row style-row-edit${off}" data-row="${idx}">
    <input type="checkbox" class="row-chk" data-row="${idx}" title="勾选=保存到模板，取消=不保存" ${row.checked === false ? "" : "checked"}>
    <span class="s-name s-name-edit" title="点击选择 / 修改样式名">${esc(row.key)}</span>
    ${rowValueCells(eff)}
    <span class="s-example">${esc(exShow)} <span class="imp-n-badge">${row.paras.length} 段</span></span>
    <span class="s-flex-gap"></span>
    <span class="s-ops">
      <button type="button" class="imp-link imp-row-expand" data-row="${idx}">${row.expanded ? "收起" : "展开"}</button>
    </span>
  </div>
  ${leadHint}
  ${row.expanded ? buildParasHtml(row, idx) : ""}`;
}

// 展开：该样式全部识别段落（每段可 归入其它样式 / 移除）
function buildParasHtml(row, idx) {
  if (!row.paras.length) return `<div class="imp-para-list" data-row="${idx}"><div class="imp-tx-muted" style="padding:6px 8px">（本样式已无段落，将使用默认值）</div></div>`;
  // 段落改判（2026-09-05 用户口径：LLM 把这一段判错样式时，必须有地方改）：
  //   ① 归入已识别的其它样式行——最常见场景（这段被判到错的样式，改到正确的那一行）
  //   ② 归入未使用的默认样式（原「单独成样式」，种子里识别结果没用到的样式）
  //   ③ 自定义新样式名（就地输入）
  // 三类目标共用同一个下拉；移入已存在的行 = 目标行吸收该段并重算格式，源行同步重算。
  const usedKeys = new Set(IMPORT_EDIT.rows.filter(r => !r.deleted).map(r => r.key));
  const freeKeys = Object.keys(IMPORT_EDIT.opt).filter(k => !usedKeys.has(k));
  const seedSummary = (k) => {
    const sd = (SEED_STYLE_CACHE || []).find(s => s.key === k) || {};
    const d = sd.defaults || {};
    return [d.font, d.size ? fontSizeToCn(d.size) : "", alignToCn(d.align),
      d.indentChars ? "缩进" + d.indentChars + "字" : "", d.bold ? "加粗" : ""].filter(Boolean).join(" · ") || "（种子默认值）";
  };
  // 已识别行的重算值摘要（与行表格同一算法 recomputeRowValue，用户据此判断该归到哪一行）
  const rowSummary = (k) => {
    const r = IMPORT_EDIT.rows.find(x => !x.deleted && x.key === k);
    if (!r) return "";
    const e = recomputeRowValue(r);
    return [e.font, e.size ? fontSizeToCn(e.size) : "", alignToCn(e.align),
      indentCellText(e) === "-" ? "" : indentCellText(e), e.bold ? "加粗" : ""].filter(Boolean).join(" · ") || `（${r.paras.length} 段）`;
  };
  // 归入目标不含：本行自身 / 已删除行 / 已取消勾选的行（取消勾选 = 该样式不入模板，归进去等于丢内容）
  const rowKeys = IMPORT_EDIT.rows.map((r, ri) => (r.deleted || r.checked === false || ri === idx ? null : r.key)).filter(Boolean);
  const moveOpts = rowKeys.length
    ? rowKeys.map(k => `<option value="${esc(k)}">${esc(k)} · ${esc(rowSummary(k))}</option>`).join("")
    : `<option value="" disabled>（暂无其它已识别样式）</option>`;
  const splitOpts = freeKeys.length
    ? freeKeys.map(k => `<option value="${esc(k)}">${esc(k)} · ${esc(seedSummary(k))}</option>`).join("")
    : `<option value="" disabled>（暂无未使用的默认样式）</option>`;
  const rows = row.paras.map((p, pi) => {
    const txt = cleanImportText(p.text);
    return `
    <div class="imp-para" data-row="${idx}" data-pi="${pi}">
      <span class="imp-para-txt" title="${esc(txt)}">${esc(truncateText(txt, 40)) || "-"}</span>
      <span class="imp-para-fmt" title="${esc(paraSummary(p))}">${esc(paraSummary(p))}</span>
      <span class="imp-para-ops">
        <select class="imp-para-split" data-row="${idx}" data-pi="${pi}" title="本段判错了就在这里改：归入其它样式（已识别样式 / 未使用的默认样式 / 自定义新名）">
          <option value="">归入…</option>
          <optgroup label="已识别样式">${moveOpts}</optgroup>
          <optgroup label="未使用的默认样式">${splitOpts}</optgroup>
          <optgroup label="自定义">
            <option value="__custom__">✏️ 输入新样式名…</option>
          </optgroup>
        </select>
        <button type="button" class="imp-link imp-link-danger imp-para-act" data-act="remove" data-row="${idx}" data-pi="${pi}">移除</button>
      </span>
    </div>`;
  }).join("");
  return `<div class="imp-para-list" data-row="${idx}">
    ${rows}
  </div>`;
}

// ── 2b 未识别默认样式：卡片已删（2026-08-27 用户定调）。
// opt 仅存内部簿记（勾选恒 false → 永不补入模板）；保留结构以兼容改名/删除行的回退逻辑。

// ── 2c 未归属段落（可提取为样式 / 忽略）──
function buildUnclaimedHtml() {
  const edit = IMPORT_EDIT;
  if (!edit.unclaimed.length) return `<div class="imp-tx-muted" style="padding:6px 8px">（无未归属段落）</div>`;
  // 未归属 = 版头区被跳过的段（红头/文号）+ LLM 判不准的低置信度段落：
  // 转换时按正文处理（内容不丢），用户可在这里「提取为样式」把它学成独立样式。
  const tip = `<div class="imp-tx-muted" style="padding:2px 8px 4px">以下段落未归入任何样式（按正文处理）。要单独学它的格式 → 「提取为样式」；不需要 → 「忽略」。</div>`;
  return tip + edit.unclaimed.map((p, i) => {
    const txt = cleanImportText(p.text);
    return `
    <div class="imp-unc-row" data-ui="${i}">
      <span class="imp-unc-txt" title="${esc(txt)}">${esc(truncateText(txt, 40)) || "-"}</span>
      <span class="imp-unc-fmt" title="${esc(paraSummary(p))}">${esc(paraSummary(p))}</span>
      <span class="imp-unc-ops">
        <button type="button" class="imp-link imp-unc-act" data-act="extract" data-ui="${i}">提取为样式</button>
        <button type="button" class="imp-link imp-link-danger imp-unc-act" data-act="ignore" data-ui="${i}">忽略</button>
      </span>
    </div>`;
  }).join("");
}

// 段落样式区（2a + 2b + 2c 三卡；挂 #imp-para-section 容器便于局部重渲染）
function buildParaSectionHtml() {
  const edit = IMPORT_EDIT;
  if (!edit) return "";
  const headHtml = `<div class="style-row style-row-head style-row-edit">
    <span class="s-chk-h"></span>
    <span class="s-name">样式名称</span>
    <span class="s-font">字体</span>
    <span class="s-size">字号</span>
    <span class="s-align">对齐</span>
    <span class="s-line">行距</span>
    <span class="s-spb">段前</span>
    <span class="s-spa">段后</span>
    <span class="s-indent">缩进</span>
    <span class="s-bold">加粗</span>
    <span class="s-color">颜色</span>
    <span class="s-example">示例</span>
    <span class="s-flex-gap"></span>
    <span class="s-ops">操作</span>
  </div>`;
  const rowsHtml = edit.rows.map((r, i) => buildParaRowHtml(r, i)).join("");
  const uncHtml = buildUnclaimedHtml();
  const optN = Object.keys(edit.opt || {}).length;
  const sum = optN
    ? `段落样式（已识别 ${edit.rows.length} + 默认样式 ${optN}）`
    : `段落样式（已识别 ${edit.rows.length}）`;
  const uncCard = edit.unclaimed.length
    ? `<details class="imp-card" id="imp-unc-card">
    <summary>未归属段落（${edit.unclaimed.length}）· 可单独提取为样式</summary>
    <div class="imp-card-body">
      ${uncHtml}
    </div>
  </details>`
    : "";
  return `
  <details class="imp-card" open id="imp-para-card">
    <summary>${sum}</summary>
    <div class="imp-card-body style-table-wrap">
      ${headHtml}
      ${rowsHtml || `<div class="imp-tx-muted" style="padding:8px">（未识别到段落样式）</div>`}
      ${buildOptSectionHtml()}
    </div>
  </details>
  ${uncCard}`;
}

// 未识别默认样式勾选子区（2026-08-29 用户定调：保存 = 识别到的 + 勾选补入；未勾选不落模板）。
// 勾选的样式以种子默认值补入模板（styles[]/include/扁平字段）；转换兜底降级保证内容不丢。
// 2026-08-30：并入「段落样式」卡做子区，横排勾选，默认直接可见。
function buildOptSectionHtml() {
  const edit = IMPORT_EDIT;
  if (!edit) return "";
  const keys = Object.keys(edit.opt);
  if (!keys.length) return "";
  const rows = keys.map(k => {
    const sd = (SEED_STYLE_CACHE || []).find(s => s.key === k) || {};
    const d = (sd.defaults) || {};
    const sum = [d.font, d.size ? fontSizeToCn(d.size) : "", alignToCn(d.align),
      d.indentChars ? "缩进" + d.indentChars + "字" : "", d.bold ? "加粗" : ""].filter(Boolean).join(" · ");
    return `<label class="opt-row" data-key="${esc(k)}">
      <input type="checkbox" class="opt-chk" data-key="${esc(k)}" ${edit.opt[k].checked ? "checked" : ""}>
      <span class="s-name" title="${esc(k)}">${esc(k)}</span>
      <span class="s-summary" title="${esc(sum || "（种子默认值）")}">${esc(sum || "（种子默认值）")}</span>
    </label>`;
  }).join("");
  return `
  <div class="imp-opt-sec">
    <div class="imp-opt-head">默认样式（${keys.length}）<span class="imp-opt-hint">勾选后随模板保存</span></div>
    <div class="opt-list">
      ${rows}
    </div>
  </div>`;
}

// 模板说明卡（结果页最前，默认展开）：选填一句话说明，保存写入 templates[].description，
// 在「文档排版 → 选择模板」悬停模板名时以 tooltip 展示（说明为空则不挂 title）。
function buildImportDescHtml() {
  return `
  <details class="imp-card" open id="imp-desc-card">
    <summary>模板说明<span class="imp-desc-note">在「文档排版」选择模板时，鼠标悬停模板名称即可看到，选填</span></summary>
    <div class="imp-card-body">
      <textarea id="imp-desc-input" class="imp-desc-input" rows="2" maxlength="200"
        placeholder="一句话说明这个模板的用途或适用场景，例如：用于向上级报送的正式请示文件">${esc(IMPORT_DESC)}</textarea>
    </div>
  </details>`;
}

// 结果区：模板说明 + 页面设置 + 段落样式（编辑态 or 旧版兼容）+ 表格样式
function buildImportResultHtml(tpl, reasoning) {
  const pageCard = buildPageCardHtml(tpl);
  const paraSection = IMPORT_EDIT
    ? `<div id="imp-para-section">${buildParaSectionHtml()}</div>`
    : buildLegacyParaCardHtml(tpl, reasoning);
  return buildImportDescHtml() + pageCard + paraSection + buildTableCardHtml(tpl, reasoning);
}

// ── 编辑态操作（2a/2b/2c 交互；每次变更 → 局部重渲染段落区 + 同步缓存）──
function rerenderParaSection() {
  const sec = document.getElementById("imp-para-section");
  if (sec) sec.innerHTML = buildParaSectionHtml();
}
function syncImportCache() {
  const c = loadImportViewCache();
  if (!c || (c.v !== 2 && c.v !== 3)) return;
  c.checks = readImportChecks(document.getElementById("import-body"));
  c.edit = IMPORT_EDIT ? JSON.parse(JSON.stringify(IMPORT_EDIT)) : undefined;
  c.desc = IMPORT_DESC;   // 模板说明（输入框内容，F5 恢复后保留）
  saveImportViewCache(c);
}
function toggleImportRow(idx) {
  const row = IMPORT_EDIT.rows[idx];
  if (!row) return;
  row.expanded = !row.expanded;
  rerenderParaSection(); syncImportCache();
}
function paraAct(act, idx, pi) {
  const row = IMPORT_EDIT.rows[idx];
  if (!row || !row.paras[pi]) return;
  if (act === "remove") {
    row.paras.splice(pi, 1);
  }
  rerenderParaSection(); syncImportCache();
}

// ── 「单独成样式」下拉（2026-08-30 用户定调：行内下拉替代弹层，仅列未使用默认样式 + 自定义新名）──
// 选中默认样式 → 段落移入（升为识别行并出 opt）；选中 __custom__ → 就地输入新名 → 建新行。
function applySplitToStyle(idx, pi, targetKey, isSeed) {
  const row = IMPORT_EDIT.rows[idx];
  if (!row || !row.paras[pi] || !targetKey) return;
  const p = row.paras.splice(pi, 1)[0];
  const existing = IMPORT_EDIT.rows.find((r, ri) => ri !== idx && !r.deleted && r.key === targetKey);
  if (existing) {
    existing.paras.push(p);
  } else {
    IMPORT_EDIT.rows.push({ key: targetKey, seed: isSeed, deleted: false, checked: true, expanded: false, paras: [p] });
  }
  if (Object.prototype.hasOwnProperty.call(IMPORT_EDIT.opt, targetKey)) delete IMPORT_EDIT.opt[targetKey];
  rerenderParaSection(); syncImportCache();
}
// 选中「✏️ 输入新样式名…」：段落行的 select 就地替换为 input+确定/取消
function startParaCustomInput(sel, idx, pi) {
  const wrap = document.createElement("span");
  wrap.className = "imp-para-custom";
  wrap.innerHTML = `<input class="imp-para-input" data-row="${idx}" data-pi="${pi}" placeholder="新样式名" maxlength="30">
    <button type="button" class="imp-link imp-para-custom-ok" data-row="${idx}" data-pi="${pi}">确定</button>
    <button type="button" class="imp-link imp-para-custom-cancel" data-row="${idx}" data-pi="${pi}">取消</button>`;
  sel.replaceWith(wrap);
  const inp = wrap.querySelector(".imp-para-input");
  inp.focus();
  const finish = () => {
    const name = (inp.value || "").trim();
    if (!name) { rerenderParaSection(); return; }   // 空名 = 取消
    // 新样式不允许重名（撞已识别行或默认样式 → 拒绝并提示）
    const used = new Set(IMPORT_EDIT.rows.filter(r => !r.deleted).map(r => r.key));
    Object.keys(IMPORT_EDIT.opt).forEach(k => used.add(k));
    if (used.has(name)) { toast(`「${name}」已存在，请换一个名称`, "err"); return; }
    applySplitToStyle(idx, pi, name, false);
  };
  wrap.querySelector(".imp-para-custom-ok").onclick = (e) => { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } finish(); };
  wrap.querySelector(".imp-para-custom-cancel").onclick = (e) => { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } rerenderParaSection(); };
  inp.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } finish(); }
    else if (e.key === "Escape") { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } rerenderParaSection(); }
  });
  inp.addEventListener("blur", () => { if (inp.isConnected) scheduleRenameCancel(inp.closest(".imp-para-custom")); });
}
function uncAct(act, ui) {
  const p = IMPORT_EDIT.unclaimed[ui];
  if (!p) return;
  if (act === "extract") {
    IMPORT_EDIT.rows.push({ key: suggestStyleName(p.text, IMPORT_EDIT.rows.map(r => r.key)), seed: false, deleted: false, checked: true, expanded: false, paras: [p] });
    IMPORT_EDIT.unclaimed.splice(ui, 1);
  } else if (act === "ignore") {
    IMPORT_EDIT.unclaimed.splice(ui, 1);
  }
  rerenderParaSection(); syncImportCache();
}
// 改名候选：种子 text 样式 + 其它识别行 + 2b 未识别默认样式 + 自定义新名
function renameCandidates(row, idx) {
  const set = new Set();
  const opts = [];
  const push = k => { if (k && k !== row.key && !set.has(k)) { set.add(k); opts.push(`<option value="${esc(k)}">${esc(k)}</option>`); } };
  for (const k of Object.keys(IMPORT_STYLE_CATS)) if (IMPORT_STYLE_CATS[k] === "text") push(k);
  IMPORT_EDIT.rows.forEach((r, ri) => { if (ri !== idx && !r.deleted) push(r.key); });
  for (const k of Object.keys(IMPORT_EDIT.opt)) push(k);
  return `<option value="${esc(row.key)}" selected>${esc(row.key)}（当前）</option>
    <option disabled>──────────</option>
    ${opts.join("")}
    <option value="__custom__">✏️ 输入新样式名…</option>`;
}
// 改名编辑态失焦取消：点击外部 / 焦点离开 → 恢复为 span 原样式名（不残留下拉/输入控件）。
// timer 延迟 150ms，允许同一交互中先执行的「确定 / 取消 / change」处理掉，避免误取消。
let _renameCancelTimer = null;
function scheduleRenameCancel(cell) {
  if (_renameCancelTimer) clearTimeout(_renameCancelTimer);
  _renameCancelTimer = setTimeout(() => {
    _renameCancelTimer = null;
    // 该 cell 仍处编辑态才恢复（元素已因 change/Enter/Esc 被替换时跳过）
    if (cell && cell.querySelector(".imp-rename-sel, .imp-rename-input, .imp-para-input")) rerenderParaSection();
  }, 150);
}
// 取消当前活动改名（恢复 span 显示原名称）；开启新编辑前调用，防止残留控件叠行
function cancelActiveRename() {
  if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; }
  const active = document.querySelector(".imp-rename-sel, .imp-rename-input");
  if (!active) return;
  const cell = active.closest(".s-name");
  const row = IMPORT_EDIT.rows[+active.dataset.row];
  if (cell && row) cell.innerHTML = esc(row.key);
}
function startRenameRow(idx) {
  const row = IMPORT_EDIT.rows[idx];
  if (!row) return;
  cancelActiveRename();
  const cell = document.querySelector(`.style-row-edit[data-row="${idx}"] .s-name`);
  if (!cell) return;
  cell.innerHTML = `<select class="imp-rename-sel" data-row="${idx}">${renameCandidates(row, idx)}</select>`;
  const sel = cell.querySelector(".imp-rename-sel");
  sel.focus();
  // isConnected 检查：change 分支用 innerHTML 替换移除 select 时其 blur 也会派发，此时不应触发取消
  sel.addEventListener("blur", () => { if (sel.isConnected) scheduleRenameCancel(cell); });
  sel.onchange = () => {
    if (sel.value === "__custom__") {
      cell.innerHTML = `<input class="imp-rename-input" data-row="${idx}" placeholder="新样式名" maxlength="30">
        <button type="button" class="imp-link imp-rename-ok" data-row="${idx}">确定</button>
        <button type="button" class="imp-link imp-rename-cancel" data-row="${idx}">取消</button>`;
      const inp = cell.querySelector(".imp-rename-input");
      inp.focus();
      inp.addEventListener("blur", () => { if (inp.isConnected) scheduleRenameCancel(cell); });
      return;
    }
    applyRename(idx, sel.value);
  };
}
function applyRename(idx, newName) {
  const row = IMPORT_EDIT.rows[idx];
  if (!row) return;
  newName = (newName || "").trim();
  if (!newName) { rerenderParaSection(); return; }   // 空名 = 取消
  const oldKey = row.key;
  const target = IMPORT_EDIT.rows.find((r, ri) => ri !== idx && !r.deleted && r.key === newName);
  if (target) {
    // 重名 → 合并（目标行吸收本行段落，格式重算）
    target.paras = target.paras.concat(row.paras);
    IMPORT_EDIT.rows.splice(idx, 1);
  } else {
    if (Object.prototype.hasOwnProperty.call(IMPORT_EDIT.opt, newName)) delete IMPORT_EDIT.opt[newName];   // 改名到未识别默认样式 → 升为识别行
    row.key = newName;
    row.seed = isSeedTextKey(newName);
  }
  // 旧 key 若为种子样式且已无行引用 → 回 2b（默认不勾）
  if (isSeedTextKey(oldKey) && !IMPORT_EDIT.rows.some(r => r.key === oldKey) && !IMPORT_EDIT.opt[oldKey]) {
    IMPORT_EDIT.opt[oldKey] = { checked: false };
  }
  rerenderParaSection(); syncImportCache();
}
// 结果区交互（事件委派）：2a 改名/展开/删除、段落三操作、2b 勾选、2c 提取/忽略
function bindImportEditEvents(bodyEl) {
  bodyEl.onclick = (e) => {
    const t = e.target;
    // 改名控件存在期间，点击其内部不重开（否则 select 的原生下拉被 innerHTML 重置而"抬起鼠标即消失"）
    if (t.closest(".imp-rename-sel") || t.closest(".imp-rename-input")) return;
    const renameOk = t.closest(".imp-rename-ok");
    if (renameOk) { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } const idx = +renameOk.dataset.row; const inp = bodyEl.querySelector(`.imp-rename-input[data-row="${idx}"]`); applyRename(idx, inp ? inp.value : ""); return; }
    const renameCancel = t.closest(".imp-rename-cancel");
    if (renameCancel) { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } rerenderParaSection(); return; }
    const expBtn = t.closest(".imp-row-expand");
    if (expBtn) { e.preventDefault(); toggleImportRow(+expBtn.dataset.row); return; }
    const paraBtn = t.closest(".imp-para-act");
    if (paraBtn) { e.preventDefault(); paraAct(paraBtn.dataset.act, +paraBtn.dataset.row, +paraBtn.dataset.pi); return; }
    const uncBtn = t.closest(".imp-unc-act");
    if (uncBtn) { e.preventDefault(); uncAct(uncBtn.dataset.act, +uncBtn.dataset.ui); return; }
    const nameEl = t.closest(".s-name-edit");
    if (nameEl) { e.preventDefault(); startRenameRow(+nameEl.parentElement.dataset.row); return; }
  };
  bodyEl.onchange = (e) => {
    const t = e.target;
    // 段落「单独成样式」下拉：未使用默认样式选中 → 直接移入；__custom__ → 就地输入新名
    const splitSel = t.closest(".imp-para-split");
    if (splitSel && splitSel.value !== "") {
      const val = splitSel.value;
      if (val === "__custom__") { startParaCustomInput(splitSel, +splitSel.dataset.row, +splitSel.dataset.pi); return; }
      applySplitToStyle(+splitSel.dataset.row, +splitSel.dataset.pi, val, isSeedTextKey(val)); return;
    }
    // 识别样式行勾选（v1.37 D1-B：取消勾选=该样式不入模板；不重渲染，保持勾选态）
    const rowChk = t.closest(".row-chk");
    if (rowChk) {
      const idx = +rowChk.dataset.row;
      const row = IMPORT_EDIT && IMPORT_EDIT.rows[idx];
      if (row) {
        row.checked = rowChk.checked;
        const rowEl = rowChk.closest(".style-row-edit");
        if (rowEl) rowEl.classList.toggle("opt-off", !rowChk.checked);
        syncImportCache();
      }
      return;
    }
    // 未识别默认样式勾选（补入模板；保存时 keep 收集）——仅同步缓存，不重渲染（保持勾选态）
    const optChk = t.closest(".opt-chk");
    if (optChk) {
      const k = optChk.dataset.key;
      if (IMPORT_EDIT && IMPORT_EDIT.opt[k]) IMPORT_EDIT.opt[k].checked = optChk.checked;
      syncImportCache();
      return;
    }
    syncImportCache();   // 页面设置 / 表格勾选 → 仅同步缓存
  };
  bodyEl.onkeydown = (e) => {
    const inp = e.target && e.target.closest && e.target.closest(".imp-rename-input");
    if (!inp) return;
    if (e.key === "Enter") { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } applyRename(+inp.dataset.row, inp.value); }
    else if (e.key === "Escape") { e.preventDefault(); if (_renameCancelTimer) { clearTimeout(_renameCancelTimer); _renameCancelTimer = null; } rerenderParaSection(); }
  };
}

// 保存：由编辑态生成最终模板（识别行保留/重算/改名映射 + 2b 勾选补入 + 表格/页面勾选）
function buildImportDocFromEdit(realTpl, bodyEl) {
  const edit = IMPORT_EDIT;
  const d = JSON.parse(JSON.stringify(realTpl));
  // 1) 保留集合 + 重算值（2a 勾选识别行 + 2b 勾选 + 表格卡勾选；v1.37 D1-B：识别行取消勾选=不入模板）
  const keep = new Set();
  const recompute = {};
  for (const r of edit.rows) {
    if (r.deleted || r.checked === false) continue;
    keep.add(r.key);
    recompute[r.key] = recomputeRowValue(r);
  }
  for (const k of Object.keys(edit.opt)) if (edit.opt[k].checked) keep.add(k);
  const tableCard = bodyEl.querySelector("#imp-table-card");
  if (tableCard) tableCard.querySelectorAll(".style-chk:checked").forEach(c => keep.add(c.dataset.key));
  // 引擎功能样式（char「重点词语」承载加粗继承、page 是渲染配置载体）恒保留——
  // 与服务端 importDocx 注释「char / page 类：引擎功能样式恒保留」口径一致，
  // 避免前端按勾选过滤把 char/page 删光导致模板失活（2026-09-01 用户反馈）。
  for (const s of (d.styles || [])) {
    if (s && (s.category === "char" || s.category === "page")) keep.add(s.key);
  }
  // 2) 删除非保留 key（结构字段除外）；保存 = 识别到的 + 勾选补入（2026-08-29 策略：未勾选默认样式不落模板）
  //    sample 为模板级字段（导入 docx 时由服务端生成的真实样本），必须保留
  for (const k of Object.keys(d)) {
    if (["name", "include", "page", "key", "footer", "styles", "styleOverrides", "sample"].includes(k)) continue;
    if (!keep.has(k)) delete d[k];
  }
  if (Array.isArray(d.include)) {
    d.include = d.include.filter(k => keep.has(k));
    for (const k of keep) if (!d.include.includes(k)) d.include.push(k);
  }
  if (Array.isArray(d.styles)) {
    d.styles = d.styles.filter(s => keep.has(s.key));
    // 勾选补入的种子样式（识别结果里没有）：补样式定义（defaults = 种子默认值），保证模板自包含
    // 2026-09-02 方案A：值随条目（defaults 自带），不再写顶层扁平字段
    const have = new Set(d.styles.map(s => s.key));
    for (const k of keep) {
      if (have.has(k)) continue;
      const sd = (SEED_STYLE_CACHE || []).find(s => s.key === k);
      if (sd) {
        d.styles.push(seedToTemplateDef(sd));   // 字体按当前开关落值（种子存的是字族名）
        have.add(k);
      }
    }
  }
  // 3) 覆盖识别行重算值（方案A：写回 styles 条目 defaults）
  for (const k of Object.keys(recompute)) {
    const s = (Array.isArray(d.styles) && d.styles.find(x => x.key === k)) || null;
    if (s) s.defaults = { ...(s.defaults || {}), ...recompute[k] };
  }
  // 4) 页面设置勾选（纸张/页边距必留）；页号并入页眉/页脚：任一保留则页码保留，都取消则页码清空（v1.38）
  const pvUnchecked = Array.from(bodyEl.querySelectorAll("[data-pv]:not(:checked)")).map(c => c.dataset.pv);
  if (d.page) {
    if (pvUnchecked.includes("header")) d.page.header = null;
    if (pvUnchecked.includes("footer")) d.page.footer = null;
    if (pvUnchecked.includes("header") && pvUnchecked.includes("footer")) d.page.pageNumber = { format: "none", align: "center", first: true };   // 落「无」不落 null（null 会被面板 dash 兜底误显示）
  }
  // 5) extraStyleDefs：保留的非种子 text 样式按行重建（defaults = 重算值；写回时并入模板 styles[]）
  const extras = [];
  for (const r of edit.rows) {
    if (r.deleted || r.checked === false) continue;
    if (isSeedTextKey(r.key)) continue;
    extras.push({ key: r.key, category: "text", matchType: "none", match: "", mode: "whole", defaults: recompute[r.key] });
  }
  // 6) 样本兜底：未保留样式的段落 → 标记降级为正文类样式（内容不丢；d.styles 此时已按 keep 过滤）
  d.sample = normalizeImportSample(d.sample, keep, d.styles);
  return { doc: d, extraStyleDefs: extras };
}


// 按勾选过滤：未勾选段落样式删除 key；页眉/页脚未勾选修空，两者都取消时页码一并清空（纸张/页边距必留）
// 同时过滤本次导入产生的非种子样式定义（extraStyleDefs）与模板 styles[]，保证落盘一致。
function filterImportByChecks(doc, bodyEl) {
  const d = JSON.parse(JSON.stringify(doc));
  const unchecked = Array.from(bodyEl.querySelectorAll(".style-chk:not(:checked)")).map(c => c.dataset.key);
  for (const k of unchecked) {
    if (k === "sample") continue;   // 模板级样本字段：不随样式勾选删除
    delete d[k];
    if (Array.isArray(d.include)) d.include = d.include.filter(x => x !== k);
    if (Array.isArray(d.styles)) d.styles = d.styles.filter(s => s.key !== k);
  }
  const pvUnchecked = Array.from(bodyEl.querySelectorAll("[data-pv]:not(:checked)")).map(c => c.dataset.pv);
  if (d.page) {
    if (pvUnchecked.includes("header")) d.page.header = null;
    if (pvUnchecked.includes("footer")) d.page.footer = null;
    if (pvUnchecked.includes("header") && pvUnchecked.includes("footer")) d.page.pageNumber = { format: "none", align: "center", first: true };   // 落「无」不落 null（null 会被面板 dash 兜底误显示）
  }
  const filteredExtras = (IMPORT_EXTRA_STYLE_DEFS || []).filter(s => !unchecked.includes(s.key));
  // 样本兜底（与 buildImportDocFromEdit 同口径）：取消勾选的样式，其样本段落降级为正文类样式，内容不丢
  d.sample = normalizeImportSample(d.sample, new Set(d.include || []), d.styles);
  return { doc: d, extraStyleDefs: filteredExtras };
}

// ═══════════════════════════════════════════
// 导入结果页：F5 / WebView2 刷新后从缓存恢复
// 模板尚未落盘：此处仅缓存「预览 + 编辑态」视图（保存才写回）；F5 后从缓存重放，
// 使重载后仍能看到结果页并继续修改 / 保存，避免刷新即丢失。
// v2：含段落明细与编辑态（可完整恢复编辑能力）；v1 旧缓存回退只读展示（无段落明细）。
// ═══════════════════════════════════════════
const IMPORT_VIEW_CACHE_KEY = "klws_import_view_v1";     // 旧版（无段落明细，只读回退）
const IMPORT_VIEW_CACHE_KEY_V2 = "klws_import_view_v2";  // 新版（含编辑态）

function saveImportViewCache(payload) {
  try { localStorage.setItem(IMPORT_VIEW_CACHE_KEY_V2, JSON.stringify(payload)); } catch (e) { /* 容量/隐私模式忽略 */ }
}
function loadImportViewCache() {
  try {
    const s2 = localStorage.getItem(IMPORT_VIEW_CACHE_KEY_V2);
    if (s2) { const p = JSON.parse(s2); if (p && (p.v === 2 || p.v === 3)) return p; }
    const s1 = localStorage.getItem(IMPORT_VIEW_CACHE_KEY);
    if (s1) { const p = JSON.parse(s1); if (p && p.realTpl) return Object.assign({ legacy: true }, p); }
  } catch (e) { /* 忽略 */ }
  return null;
}
function clearImportViewCache() {
  try { localStorage.removeItem(IMPORT_VIEW_CACHE_KEY_V2); localStorage.removeItem(IMPORT_VIEW_CACHE_KEY); } catch (e) { /* 忽略 */ }
}

// 读取当前勾选状态（页面项 + 表格样式；2a/2b/2c 编辑态存 IMPORT_EDIT）
function readImportChecks(bodyEl) {
  const style = {}, pv = {};
  bodyEl.querySelectorAll(".style-chk").forEach(c => { style[c.dataset.key] = c.checked; });
  bodyEl.querySelectorAll("[data-pv]").forEach(c => { pv[c.dataset.pv] = c.checked; });
  return { style, pv };
}
// 还原勾选状态（F5 后保留用户已改的勾选）
function applyImportChecks(bodyEl, checks) {
  if (!checks) return;
  if (checks.style) bodyEl.querySelectorAll(".style-chk").forEach(c => {
    if (checks.style[c.dataset.key] !== undefined) c.checked = checks.style[c.dataset.key];
  });
  if (checks.pv) bodyEl.querySelectorAll("[data-pv]").forEach(c => {
    if (checks.pv[c.dataset.pv] !== undefined) c.checked = checks.pv[c.dataset.pv];
  });
}

// 渲染「导入结果」视图（完成态 与 F5 恢复态 共用），并绑定 编辑交互 + 保存/关闭 按钮
async function renderImportResultView({ realTpl, reasoning, thinkText, name, tsFile, cached, elapsedMs }) {
  const bodyEl = document.getElementById("import-body");
  const okBtn = document.getElementById("import-ok");
  const cancelBtn = document.getElementById("import-cancel");
  const stopBtn = document.getElementById("import-stop");
  const metaEl = document.getElementById("import-doc-meta");

  switchView("import");
  stopBtn.style.display = "none";
  okBtn.style.display = "";
  okBtn.disabled = false;
  okBtn.textContent = "保存";
  cancelBtn.textContent = "放弃";
  cancelBtn.disabled = false;
  metaEl.textContent = name || (realTpl && realTpl.name) || "导入模板";

  // 预加载种子：保证「默认样式」chip 摘要首屏就有（不再依赖点击展开触发重渲染）
  await ensureSeedStyles();

  IMPORT_REAL_TPL = realTpl;
  // 模板说明：F5 恢复优先取缓存（用户在结果页已输入的内容），否则取模板自带值
  IMPORT_DESC = (cached && typeof cached.desc === "string")
    ? cached.desc
    : ((realTpl && realTpl.description) || "");
  if (!IMPORT_EDIT && !(cached && cached.legacy)) IMPORT_EDIT = initImportEdit(realTpl, reasoning);

  bodyEl.innerHTML = buildImportProcessHtml(realTpl, reasoning, thinkText, elapsedMs) + buildImportResultHtml(realTpl, reasoning);

  const procEl = bodyEl.querySelector("#import-process");
  if (procEl) procEl.open = false;

  // 模板说明输入：实时写入 IMPORT_DESC 并同步缓存（不走段落区重渲染，输入焦点不丢）
  const descEl = bodyEl.querySelector("#imp-desc-input");
  if (descEl) {
    descEl.addEventListener("input", () => {
      IMPORT_DESC = (descEl.value || "").trim() ? descEl.value : "";
      syncImportCache();
    });
  }

  // A：思考过程折叠摘要（完成态）：默认折叠为单行「思考过程 · N 字」，点击展开回看全文
  //    （思考全文从活跃态累积/缓存恢复；用 dataset 存原文避免 HTML 转义风险）
  const foldEl = bodyEl.querySelector("#import-process [data-think]");
  if (foldEl && thinkText) {
    foldEl.dataset.fullText = thinkText;
    foldEl.textContent = "▸ 思考过程 · " + thinkText.length + " 字（点击展开）";
    bindReasonToggle(foldEl);   // 折叠/展开交互（与转换面板思考块同款）
  }

  // F5 恢复：还原用户已调整的勾选（表格 / 页面；2a/2b/2c 由 IMPORT_EDIT 还原）
  if (cached && cached.checks) applyImportChecks(bodyEl, cached.checks);

  // 编辑交互（改名 / 展开 / 删除 / 段落修正 / 2b / 2c）
  bindImportEditEvents(bodyEl);

  okBtn.onclick = async () => {
    okBtn.disabled = true; cancelBtn.disabled = true; okBtn.textContent = "保存中…";
    try {
      const { doc: filtered, extraStyleDefs: filteredExtras } = IMPORT_EDIT
        ? buildImportDocFromEdit(realTpl, bodyEl)
        : filterImportByChecks(realTpl, bodyEl);
      // 模板说明随模板落盘（空则不写字段，避免一堆 "description": ""）
      if (IMPORT_DESC) { filtered.description = IMPORT_DESC; realTpl.description = IMPORT_DESC; }
      else { delete filtered.description; delete realTpl.description; }
      const savedFile = await saveImportedToSet(filtered, name || realTpl.name, tsFile, filteredExtras);
      if (!savedFile) throw new Error("已取消保存");
      bodyEl.insertAdjacentHTML("afterbegin", `<div class="imp-ok">✓ 已按勾选保存到模板集，可在左侧菜单查看</div>`);
      okBtn.style.display = "none";
      cancelBtn.textContent = "完成";
      clearImportViewCache();
      removeImportPlaceholder(realTpl.key);   // 已落盘：移除「待保存」占位，菜单回到干净态
      try {
        await openImportedTemplateInDesign(savedFile, realTpl.key);
        toast(`✓ 已保存模板「${realTpl.name}」，进入模板编辑`, "ok");
      } catch (e2) {
        toast("已保存，但切换编辑视图失败：" + e2.message, "err");
      }
    } catch (e) {
      bodyEl.insertAdjacentHTML("afterbegin", `<div class="imp-err"><b>保存失败</b><br>${esc(e.message)}</div>`);
      okBtn.disabled = false; cancelBtn.disabled = false; okBtn.textContent = "保存修改";
    }
  };
  cancelBtn.onclick = () => {
    if (realTpl) removeImportPlaceholder(realTpl.key);  // 放弃：移除「待保存」占位，不残留未确认模板
    clearImportViewCache();
    switchView("convert");
  };
}

// 启动后调用：若上次导入结果被刷新中断，自动恢复预览页
async function restoreImportView() {
  const c = loadImportViewCache();
  if (!c || !c.realTpl) return false;
  // 预加载种子：保证「默认样式」chip 摘要首屏就有（与 renderImportResultView 同源）
  await ensureSeedStyles();
  IMPORT_STYLE_NAMES = c.styleNames || {};
  IMPORT_STYLE_CATS = c.styleCats || {};
  IMPORT_EXTRACTED_KEYS = new Set(c.extractedKeys || []);
  IMPORT_STYLE_GROUPS = c.styleGroups || {};
  IMPORT_EXTRA_STYLE_DEFS = c.extraStyleDefs || [];
  IMPORT_SAMPLES = c.samples || {};
  IMPORT_STYLE_GROUPS_DETAIL = c.styleGroupsDetail || {};
  IMPORT_UNCLAIMED = c.unclaimed || [];
  IMPORT_REAL_TPL = c.realTpl;
  IMPORT_EDIT = c.edit ? JSON.parse(JSON.stringify(c.edit)) : (c.legacy ? null : initImportEdit(c.realTpl, c.reasoning));
  restoreImportPlaceholder(c);   // 重建「待保存」占位：切走/刷新后仍可从菜单找回结果页
  await renderImportResultView({
    realTpl: c.realTpl,
    reasoning: c.reasoning,
    thinkText: c.reasoningText || "",
    name: c.name,
    tsFile: c.tsFile,
    cached: c,
    elapsedMs: c.elapsedMs || null,
  });
  return true;
}

// F5 恢复时重建「待保存」占位：刷新后内存占位清空，若用户从结果页切走，
// 仍可凭左侧菜单的「待保存·…」项点回结果页，避免「刷新后找不到导入模板」。
function restoreImportPlaceholder(c) {
  const realTpl = c && c.realTpl;
  if (!realTpl) return;
  const key = c.tplKey || realTpl.key;
  if (_importPlaceholders.has(key)) return;   // 已有同 key 占位，不重复建
  const ph = blankTemplateStruct(key, `待保存·${realTpl.name}`);
  ph._placeholder = true;
  ph._status = "ready";
  ph._realTpl = realTpl;
  ph._tsFile = c.tsFile || "";
  ph._importFile = c.name || "";
  _importPlaceholders.set(key, ph);
  if (c.tsFile) uiState.expandedSets.add(c.tsFile); // 展开目标集，让占位可见
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
}

// ═══════════════════════════════════════════
// 导入后台化：占位模板管理
// 导入开始时在左侧菜单对应模板集下创建占位模板（全局内存 Map，不落盘），
// LLM 推理期间用户可随意切换视图 / 模板 / 模板集 / 转换任务；推理完成后
// 占位升级为「待保存」态（ready，不落盘）继续留在菜单，由结果页「保存」显式写回目标集
// （保存成功 / 明确放弃后占位才移除），失败则占位标记失败态（可删除）。
// 这正是「占位仅内存、保存才落盘、放弃不落盘」模型：用户在任何阶段「放弃」都不会在
// 磁盘 / 菜单残留未确认模板。占位在菜单保留到保存/放弃，保证导入结果始终有找回入口
// （若完成即删占位 + 模板未落盘 + 导航无导入入口，切走视图的用户将找不到刚导入的模板）。
// 占位挂在「导入时的目标集」下（_tsFile），与当前编辑集解耦：
// 中途切换模板集再切回来，占位依然在菜单中显示（导入中 / 待保存 / 失败态均可见）。
// ═══════════════════════════════════════════
const _importPlaceholders = new Map(); // phKey → 占位对象（含 _tsFile 目标集）
const _canceledImports = new Set();  // 已被用户取消（删除占位）的导入 key：返回结果将被丢弃
const _importAbortMap = new Map();   // phKey → AbortController：导入中取消时中止推理流（弹窗放弃 / 菜单取消导入共用）

// 默认空模板结构（与 addTemplate 一致，保证设计视图渲染不崩）
function blankTemplateStruct(key, name) {
  const sp = seedPageDefaults();
  return {
    key, name,
    include: [...styleKeys()],
    page: {
      width: 11906, height: 16838,
      marginTop: 2098, marginRight: 1474, marginBottom: 1984, marginLeft: 1587,
      linePitch: sp.linePitch != null ? sp.linePitch : 579,
      numberGap: sp.numberGap != null ? sp.numberGap : DEFAULT_NUMBER_GAP,
      pageNumber: { format: "dash", align: "center", first: true },
    },
    title: {}, subtitle: {}, h1: {}, h2: {}, body: {}, footer: {},
  };
}

// 创建导入占位并显示在左侧菜单；返回占位对象
// tsFile：入口集（点击 ⋯ 时钉死），直接作为落点，不再依赖 state.currentFile
function ensureImportPlaceholder(fileName, tsFile) {
  const usedKeys = new Set(_importPlaceholders.keys());
  let key;
  // key 必须跨会话唯一：不能用内存递增序号（F5/重启后归 0，新导入会复用磁盘上
  // 上次导入保存过的 key；保存时 writeTemplateToSet 按 key 匹配 → 覆盖已有模板）。
  // 改用时间戳 + 随机数，刷新后生成的新 key 与任何历史导入 key 都不冲突。
  do {
    key = "imp" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  } while (usedKeys.has(key));
  const ph = blankTemplateStruct(key, `导入中… ${fileName}`);
  ph._placeholder = true;
  ph._status = "importing";
  ph._importFile = fileName;
  ph._tsFile = tsFile || "";   // 入口集钉死：从哪个模板集 ⋯ 进来的就存哪个，对话框期间切集不影响
  _importPlaceholders.set(key, ph);
  if (tsFile) uiState.expandedSets.add(tsFile); // 自动展开入口集，让占位可见
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
  return ph;
}

// 完成：占位升级为「内存态真实模板」（不落盘）并保留为「待保存」态；返回 real；占位已从 Map 删除则返回 null。
// 落盘交由结果页「保存」显式触发（saveImportedToSet），保证「放弃」可无损回退。
// 关键：占位【不删除】而是标记 ready 留在左侧菜单——若用户导入中切走了视图，菜单里仍能点回结果页保存，
// 否则「占位消失 + 模板未落盘 + 导航无导入入口」三者叠加，用户将彻底找不到刚导入的模板。
function completeImportPlaceholder(phKey, doc, fileName) {
  const ph = _importPlaceholders.get(phKey);
  if (!ph) return null;
  const name = (fileName || "导入模板").replace(/\.docx$/i, "").replace(/[\\/:*?"<>|]/g, "_") || "导入模板";
  const real = Object.assign({}, doc, { key: phKey, name });
  ph._status = "ready";        // 待保存：菜单可见可点（点击 → 恢复结果页）
  ph._realTpl = real;          // 挂真实模板备用（结果页渲染主用全局 IMPORT_REAL_TPL / 缓存）
  ph.name = `待保存·${name}`;  // 菜单显示名带状态前缀，不与集内模板撞名（保存后占位移除、恢复本名）
  _importAbortMap.delete(phKey);   // 推理已结束，清理 abort 登记
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
  return real;
}

// 失败：占位标记失败态（红色显示在菜单，可删除）
function failImportPlaceholder(phKey, errMsg) {
  const t = _importPlaceholders.get(phKey);
  if (!t) return;
  t._status = "error";
  t._error = errMsg || "未知错误";
  _importAbortMap.delete(phKey);   // 流已结束，清理 abort 登记
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
}

// 取消 / 删除占位：导入中删除 → 中止推理流 + 标记取消（返回结果丢弃）；失败态删除 → 直接移除
// 收敛点：菜单「取消导入」与弹窗「放弃」均经此函数，故在此统一 abort 流，两路径行为一致。
function removeImportPlaceholder(phKey) {
  const t = _importPlaceholders.get(phKey);
  if (!t) return;
  if (t._status === "importing") {
    _canceledImports.add(phKey);                 // 标记取消：迟到的流式结果在 doImport 中丢弃
    const ac = _importAbortMap.get(phKey);       // 中止上游 LLM 推理流（避免白跑）
    if (ac) { ac.abort(); _importAbortMap.delete(phKey); }
  }
  _importPlaceholders.delete(phKey);
  if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
}

// 同名模板去重（方案 C 变体 2026-08-30）：同名但不同 key（如重复导入同一份 docx）不再弹确认，
// 直接追加序号后缀区分：「xxx」已存在 → 新名「xxx2」→「xxx3」…（跳过已占用的名字）
function uniqueTemplateName(existingNames, baseName) {
  const names = new Set(existingNames);
  if (!names.has(baseName)) return baseName;
  for (let i = 2; ; i++) {
    const cand = baseName + i;
    if (!names.has(cand)) return cand;
  }
}

// 把模板写入目标模板集：目标集=当前内存集 → 覆盖/追加内存并落盘；
// 目标集=其他集 → 静默读改写目标集文件（不打扰当前编辑）；tsFile 空 → 新建集（弹框命名）
// tpl 带 key：同 key 覆盖更新，否则追加（同名但不同 key 自动加序号后缀，不弹确认）。
// activate：是否切为当前编辑模板（仅用户主动保存时 true）。
// extraDefs：本次导入/保存需要合并的非种子样式定义（已按用户勾选过滤）；默认读全局 IMPORT_EXTRA_STYLE_DEFS。
// 返回：成功 = 落盘的文件名（字符串）；用户取消 = false。
async function writeTemplateToSet(tpl, fileName, tsFile, activate, extraDefs) {
  const defsToMerge = extraDefs || IMPORT_EXTRA_STYLE_DEFS || [];
  // v1.30 解耦：导入产生的非种子样式定义直接并入当前模板的自包含 styles[]（随模板走）。
  // 不再写集级 customStyles；defsToMerge 仅兜底补 styles[] 缺失项（去重、只补 include 引用的）。
  const mergeExtraStyles = (t) => {
    if (!Array.isArray(t.styles)) t.styles = [];
    for (const d of defsToMerge) {
      if (!(t.include || []).includes(d.key)) continue;
      if (!t.styles.some(s => s.key === d.key)) {
        t.styles.push(JSON.parse(JSON.stringify(d)));
      }
    }
  };
  let file = tsFile;
  if (!file && state.currentFile) file = state.currentFile;
  if (!file) {
    const base = (fileName || "导入模板").replace(/\.docx$/i, "").replace(/[\\/:*?"<>|]/g, "_");
    file = await modalPrompt("保存模板集", { value: base + ".json", placeholder: "模板集文件名（如：十五五规划.json）" });
    if (!file) return false;
    if (!/\.json$/i.test(file)) file += ".json";
  }
  if (file === state.currentFile && state.templateSet) {
    // 先把本次模板写入目标集，再统计 usedKeys，否则新模板 include 还没生效
    const idx = state.templateSet.templates.findIndex(x => x.key === tpl.key);
    if (idx >= 0) {
      state.templateSet.templates[idx] = tpl;
    } else {
      // 同名但不同 key（重复导入同一份 docx）：直接加序号后缀区分，不弹确认
      const existingNames = state.templateSet.templates.map(x => x.name);
      if (existingNames.includes(tpl.name)) tpl.name = uniqueTemplateName(existingNames, tpl.name);
      state.templateSet.templates.push(tpl);
    }
    mergeExtraStyles(tpl);
    // 先落盘、后切当前模板并渲染（2026-08-30 修复「导入模板后预览空白」）：
    // 若先 renderAll()，loadSample() 的 GET /api/sample 会先于 persistCurrentSet() 的写盘
    // 到达服务端，读到不含新模板 key 的旧文件 → 500「模板不存在」→ 预览清空，
    // 且 _sampleKey 去重使后续 renderAll 不再重试，空白永久保持。落盘完成后再渲染即可命中新文件。
    await persistCurrentSet();
    syncCurrentSetInStatus();
    if (activate) { // 仅主动保存时切为当前编辑模板；后台自动保存不抢占用户编辑
      state.currentTemplateKey = tpl.key;
      state.selectedElem = null;
      renderAll();
    }
    return file;
  }
  // 目标集非当前内存集：读改写目标集文件，静默完成，不切走当前集
  let data = null;
  try {
    const r = await api("/api/templateset?file=" + encodeURIComponent(file));
    data = r.data;
  } catch (e) { data = null; }
  if (!data || !Array.isArray(data.templates)) data = { name: "导入模板集", templates: [] };
  const idx = data.templates.findIndex(x => x.key === tpl.key);
  if (idx >= 0) {
    data.templates[idx] = tpl;
  } else {
    // 同名但不同 key（重复导入同一份 docx）：直接加序号后缀区分，不弹确认
    const existingNames = data.templates.map(x => x.name);
    if (existingNames.includes(tpl.name)) tpl.name = uniqueTemplateName(existingNames, tpl.name);
    data.templates.push(tpl);
  }
  mergeExtraStyles(tpl);
  // 与其他模板集写共用串行队列，避免并发 POST 丢更新（见 persistCurrentSet 队列注释）
  await enqueueSetWrite(() => api("/api/templateset", { method: "POST", body: JSON.stringify({ file, data }) }));
  // 同步左侧树数据源，让新模板立即出现在菜单
  if (_lastTreeStatus) {
    const ts = _lastTreeStatus.templatesets.find(t => t.file === file);
    if (ts && !ts.templates.includes(tpl.name)) ts.templates.push(tpl.name);
  }
  return file;
}

// 导入 docx 生成模板：占满整个工作区展示（活跃态 = 4 阶段进度「解析文档→分析样式特征→智能识别样式→
// 生成模板」；2026-09-02 起不再展示 LLM 思考链原文。完成态折叠为摘要 + 结果两类 + 勾选保存）
async function doImport(file, tsFile) {
  // LLM 前置检查：导入依赖 LLM 分类，未配置直接提示——不发起导入、不产生左侧菜单占位
  if (!(await checkLlmConfigured())) {
    toast("导入模板需配置模型服务：请到「设置」启用一个模型服务", "err");
    return;
  }
  // 模板样本前置大小检查（与后端 /api/import-docx 同阈值）：模板是学格式的样本，正常都很小，超限直接拒绝
  if (file.size > 10 * 1024 * 1024) {
    toast("文件过大，不适合作为模板：请选择较小的文档（模板仅用于学习格式）", "err");
    return;
  }
  // 禁止并发导入：已有「导入中」或「待保存」占位时拒绝新导入——并发导入会互相顶掉结果页/缓存，
  // 先完成的导入无处保存直接丢失（覆盖正在进行的导入任务）；待保存占位同理：新导入完成会覆盖
  // 全局唯一的结果缓存，导致待保存模板无法找回。失败态占位不阻止。
  const hasImporting = [..._importPlaceholders.values()].some(p => p._status === "importing");
  const hasPending = [..._importPlaceholders.values()].some(p => p._status === "ready");
  if (hasImporting || hasPending) {
    toast(hasImporting ? "已有导入正在进行，请先完成或放弃当前导入" : "已有生成的模板待保存，请先在左侧菜单点「待保存·…」保存或放弃", "err");
    return;
  }
  const bodyEl = document.getElementById("import-body");
  const okBtn = document.getElementById("import-ok");
  const cancelBtn = document.getElementById("import-cancel");
  const stopBtn = document.getElementById("import-stop");
  const metaEl = document.getElementById("import-doc-meta");

  // 重置上次导入的编辑态（新导入从零开始）
  IMPORT_EDIT = null;
  IMPORT_REAL_TPL = null;
  IMPORT_DESC = "";

  // ── 后台化：先在左侧菜单插入占位模板（全局内存态），LLM 推理期间可随意切走 ──
  const ph = ensureImportPlaceholder(file.name, tsFile);
  const phKey = ph.key;
  const phTsFile = ph._tsFile;   // 记录目标集：finalize 后占位从 Map 移除，需提前保存

  switchView("import");
  bodyEl.dataset.impKey = phKey;   // 标记本视图属于哪次导入：并发导入时旧任务不覆盖新视图
  okBtn.style.display = "none";
  cancelBtn.textContent = "放弃";
  stopBtn.style.display = "none";
  metaEl.textContent = file.name;

  // ── 模板提取面板（活跃态）：阶段日志行（B，仿转换过程面板：里程碑 ✓ 行 + LLM 计时行）
  //    + 思考过程折叠区（A：默认收起不铺开英文原文，点击展开实时滚动）
  //    + LLM 思考→输出切换状态（C：收到 llm-content 提示「正在生成识别结果」，防静默卡住）
  //    （2026-09-03 起弃用旧的「单行状态覆盖 + 思考原文强制滚动展示」）
  bodyEl.innerHTML = `
    <details class="imp-think" id="import-process" open>
      <summary id="import-process-summary" class="imp-think-summary">模板提取</summary>
      <div class="imp-think-body">
        <div class="imp-log" id="import-log"></div>
        <div class="imp-reason-toggle" id="import-reason-toggle" style="display:none"></div>
        <div class="imp-reason-live" id="import-reason-live"></div>
      </div>
    </details>`;
  const logEl = document.getElementById("import-log");
  const toggleRow = document.getElementById("import-reason-toggle");
  const reasonEl = document.getElementById("import-reason-live");
  // 阶段日志行（轻量版转换面板行）：里程碑 ✓ / 进行中 ◌（cur）
  let firstLine = null;   // 「读取」行：后续 read 里程碑原地升 ✓
  const mkLine = (icon, text) => {
    const d = document.createElement("div");
    d.className = "conv-line" + procLineCls(icon);
    d.innerHTML = `<span class="conv-ic">${procIcon(icon)}</span><span class="conv-msg"></span>`;
    d.querySelector(".conv-msg").textContent = text;
    if (logEl) logEl.appendChild(d);
    return d;
  };
  const setLine = (el, icon, text) => {
    if (!el) return;
    el.className = "conv-line" + procLineCls(icon);
    el.querySelector(".conv-ic").innerHTML = procIcon(icon);
    el.querySelector(".conv-msg").textContent = text;
  };
  // LLM 行实时计时（思考 + 输出识别结果整个阶段；500ms 刷新一次）
  let llmRow = null, llmT0 = 0, llmBase = "", llmTimer = 0;
  const llmTick = () => {
    if (!llmRow || !llmT0) return;
    setLine(llmRow, "◌", llmBase + " · " + ((Date.now() - llmT0) / 1000).toFixed(1) + "s");
  };
  const stopLlmTimer = () => { if (llmTimer) { clearInterval(llmTimer); llmTimer = 0; } };
  // 思考过程：默认展开实时滚动（D1 拍板 2026-09-05，沿用转换页行为）；标签行文案与转换页思考块同款
  // （「样式识别 · 思考中…/正在生成结果…/完成/失败」），与转换页思考块同款
  let reasonOpen = true;
  let reasonSeen = false;
  let reasonState = "think";   // think | gen | done | fail（与转换页 conv-reason-tag 状态语义一致）
  const updateToggle = () => {
    if (!toggleRow) return;
    const tail = reasonState === "gen" ? " · 正在生成结果…"
      : reasonState === "done" ? " · 完成，点击展开"
      : reasonState === "fail" ? " · 失败，点击展开"
      : (reasonOpen ? " · 思考中…" : " · 思考中…（已收起，点击展开）");
    toggleRow.textContent = "样式识别" + tail;
  };
  toggleRow.onclick = () => {
    reasonOpen = !reasonOpen;
    reasonEl.style.display = reasonOpen ? "" : "none";
    reasonEl.classList.toggle("streaming", reasonOpen);
    if (reasonOpen) reasonEl.scrollTop = reasonEl.scrollHeight;
    updateToggle();
  };
  // 思考流缓冲：推理模型每秒数十个 chunk，逐块直写 DOM（textContent += 会 O(n) 重建节点，
  // scrollTop 读 scrollHeight 强制同步重排）随内容变长持续放大 → 主线程被占满、界面卡死。
  // 改为纯字符串缓冲 + 每帧一次落盘（与转换过程面板同源修复）。
  let reasonBuf = "";
  let reasonRaf = 0;
  const flushReason = () => {
    if (reasonRaf) { cancelAnimationFrame(reasonRaf); reasonRaf = 0; }
    if (!reasonEl || !reasonBuf) return;
    reasonEl.textContent += reasonBuf;
    reasonBuf = "";
    if (reasonOpen) reasonEl.scrollTop = reasonEl.scrollHeight;   // 仅展开态跟随滚动
  };
  let ac = null;   // 本次导入的 AbortController（流启动后赋值）
  // 读取行 + 实时计时（D3 拍板 2026-09-05）：覆盖前端读文件 + 上传 + 服务端解析整个阶段
  firstLine = mkLine("◌", "正在读取文件…");   // 初始反馈：读取完成前不让面板空着
  let readRow = firstLine, readT0 = Date.now(), readBase = "正在读取文件…", readTimer = 0;
  const stopReadTimer = () => { if (readTimer) { clearInterval(readTimer); readTimer = 0; } };
  readTimer = setInterval(() => {
    if (!readRow || !readRow.isConnected) { stopReadTimer(); return; }   // 面板已重建（结果页/失败）→ 停表
    setLine(readRow, "◌", readBase + " · " + ((Date.now() - readT0) / 1000).toFixed(1) + "s");
  }, 500);

  try {
    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).split(",")[1]);
      reader.onerror = () => reject(new Error("读取文件失败"));
      reader.readAsDataURL(file);
    });

    ac = new AbortController();
    _importAbortMap.set(phKey, ac);   // 登记 abort 控制器，供菜单「取消导入」中止流
    stopBtn.style.display = "";
    stopBtn.onclick = () => { ac.abort(); };   // 停止：中止流并丢弃本次导入
    cancelBtn.onclick = () => {                // 放弃：中止流 + 删除占位模板（迟到的结果丢弃）
      if (ac) ac.abort();
      removeImportPlaceholder(phKey);
      switchView("convert");
    };

    // N8（2026-08-27）：导入也生成 taskId 随请求透传（服务端临时文件名携带，可追溯）
    const importTaskId = "i" + Date.now();
    const resp = await fetch("/api/import-docx/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileName: file.name, base64, taskId: importTaskId }),
      signal: ac.signal,
    });
    if (!resp.ok) {
      let msg = "请求失败(" + resp.status + ")";
      try { const j = await resp.json(); if (j && j.error) msg = j.error; } catch (e) { /* 非 JSON */ }
      throw new Error(msg);
    }
    if (!resp.body) throw new Error("当前浏览器不支持流式读取（无 ReadableStream）");

    // ── 消费 SSE：phase → 阶段日志行（read/parse 里程碑 ✓；llm 计时行；llm-content 切「生成识别结果」）；
    //    reasoning → 思考折叠区累积（默认收起，点击展开）；done 收结果；error 抛错 ──
    let streamResult = null;
    let genLine = null;   // 「正在生成模板…」行：done 时升 ✓
    await sseReadStream(resp, (ev, data) => {
      if (ev === "phase") {
        const stage = (data && data.stage) || "";
        const msg = (data && data.message) || "";
        if (stage === "read-start") {
          // 服务端已收到上传、开始解析：切文案，计时继续（与转换链路同名阶段，D3 拍板）
          readBase = msg || "已上传，正在读取文件…";
          if (readRow && readRow.isConnected) setLine(readRow, "◌", readBase + " · " + ((Date.now() - readT0) / 1000).toFixed(1) + "s");
        } else if (stage === "read") {
          stopReadTimer();
          const sec = ((Date.now() - readT0) / 1000).toFixed(1);
          // 读取完成行带文件名+大小（信息量对齐转换页 read 行）
          const readTxt = "已读取文件：" + file.name + "（" + formatBytes(file.size) + " · " + sec + "s）";
          if (firstLine) { setLine(firstLine, "✓", readTxt); firstLine = null; }
          else mkLine("✓", readTxt);
        } else if (stage === "parse") {
          mkLine("✓", msg || "解析完成");
        } else if (stage === "llm") {
          llmT0 = Date.now();
          llmBase = msg || "调用 LLM 中（样式识别）…";   // 句式与转换页 LLM 行统一
          llmRow = mkLine("◌", llmBase);
          llmTimer = setInterval(llmTick, 500);
        } else if (stage === "llm-content") {
          // LLM 思考结束、开始输出识别结果（JSON）：同一行切换状态，防 LLM 静默期体感「卡住」
          reasonState = "gen"; updateToggle();   // 思考标签同步切「正在生成结果」
          if (llmRow) { llmBase = msg || "生成结果中（样式识别）…"; llmTick(); }
        } else if (stage === "build") {
          // LLM 分类完成 → 本地生成模板（极快）：LLM 行升 ✓（带耗时）
          if (llmRow && llmT0) setLine(llmRow, "✓", "调用 LLM 完成（样式识别） · " + ((Date.now() - llmT0) / 1000).toFixed(1) + "s");
          stopLlmTimer();
          genLine = mkLine("◌", msg || "正在生成模板…");
        }
      } else if (ev === "reasoning") {
        const t = (data && data.text) || "";
        if (t) {
          if (!reasonSeen) { reasonSeen = true; toggleRow.style.display = ""; reasonEl.classList.add("streaming"); updateToggle(); }   // 思考开始：露出开关行 + 光标
          reasonBuf += t;
          if (!reasonRaf) {
            reasonRaf = requestAnimationFrame(() => { reasonRaf = 0; flushReason(); });
          }
        }
      } else if (ev === "done") {
        flushReason();   // 收尾补写，避免最后一块思考丢失
        reasonEl.classList.remove("streaming");   // 思考结束：去光标
        stopReadTimer();
        if (genLine) { setLine(genLine, "✓", "模板生成完成"); genLine = null; }
        if (llmRow && llmT0 && llmTimer) { stopLlmTimer(); setLine(llmRow, "✓", "调用 LLM 完成（样式识别） · " + ((Date.now() - llmT0) / 1000).toFixed(1) + "s"); }
        streamResult = data;
        if (llmT0) streamResult.llmMs = Date.now() - llmT0;   // LLM 阶段耗时：结果页统计卡「LLM 耗时」格
        reasonState = "done"; updateToggle();   // 思考标签切「完成」
        IMPORT_STYLE_NAMES = (data && data.styleNames) || {};
        IMPORT_STYLE_CATS = (data && data.styleCats) || {};
        IMPORT_EXTRACTED_KEYS = new Set((data && data.extractedKeys) || []);
        IMPORT_STYLE_GROUPS = (data && data.styleGroups) || {};
        IMPORT_EXTRA_STYLE_DEFS = (data && data.extraStyleDefs) || [];
        IMPORT_SAMPLES = (data && data.samples) || {};
        IMPORT_STYLE_GROUPS_DETAIL = (data && data.styleGroupsDetail) || {};
        IMPORT_UNCLAIMED = (data && data.unclaimed) || [];
      } else if (ev === "error") {
        flushReason();   // 出错也补写思考流，便于排查卡在哪一步
        stopReadTimer();
        stopLlmTimer();
        reasonState = "fail"; updateToggle();   // 思考标签切「失败」
        if (llmRow) setLine(llmRow, "✗", "调用 LLM 失败（样式识别）");
        if (genLine) { setLine(genLine, "✗", "模板生成失败"); genLine = null; }
        throw new Error((data && (data.message || data.reason)) || "导入流式请求出错");
      }
    });
    if (!streamResult || !streamResult.doc) throw new Error("导入未返回模板结果");
    if (_canceledImports.has(phKey)) { _canceledImports.delete(phKey); return; }  // 用户已取消，丢弃结果
    // ── 完成：占位升级为内存态真实模板（不落盘）；结果页「保存」才写入模板集 ──
    const realTpl = completeImportPlaceholder(phKey, streamResult.doc, file.name);
    if (!realTpl) { _canceledImports.delete(phKey); return; }

    // 思考全文（活跃态累积）：完成态折叠为「思考过程 · N 字」摘要，可点击展开回看（F5 恢复同样保留）
    const thinkText = reasonEl && reasonEl.textContent ? reasonEl.textContent : "";

    // 结果缓存：F5 / WebView2 刷新后可在 init 还原（无论是否正在查看都缓存）。
    // ⚠ 缓存带版本号（v），字段只增不删：reasoningText 存思考全文供刷新还原，删它会让旧缓存升级断层。
    const cachePayload = {
      v: 3, name: file.name, tplKey: realTpl.key, tsFile: phTsFile,
      realTpl, reasoning: streamResult.reasoning, reasoningText: thinkText, checks: null,
      elapsedMs: streamResult.llmMs || null,
      styleNames: IMPORT_STYLE_NAMES,
      styleCats: IMPORT_STYLE_CATS,
      extractedKeys: [...IMPORT_EXTRACTED_KEYS],
      styleGroups: IMPORT_STYLE_GROUPS,
      extraStyleDefs: IMPORT_EXTRA_STYLE_DEFS,
      samples: IMPORT_SAMPLES,
      styleGroupsDetail: IMPORT_STYLE_GROUPS_DETAIL,
      unclaimed: IMPORT_UNCLAIMED,
      edit: null,   // renderImportResultView 渲染后回填编辑态
    };
    // 仅当用户当前仍停留在导入视图时才渲染结果页（不强制切回、不打断正在进行的其他操作）；
    // 切走视图的用户：占位已升级为「待保存」留在左侧菜单，点「待保存·…」即可恢复结果页继续保存。
    // （2026-08-27：成功 toast 已删除——结果页/待保存占位自身即状态提示，toast 冗余）
    if (bodyEl.dataset.impKey === phKey && currentView === "import") {
      await renderImportResultView({ realTpl, reasoning: streamResult.reasoning, thinkText, name: file.name, tsFile: phTsFile, elapsedMs: streamResult.llmMs || null });
      cachePayload.checks = readImportChecks(bodyEl);
      cachePayload.edit = IMPORT_EDIT ? JSON.parse(JSON.stringify(IMPORT_EDIT)) : null;
      cachePayload.desc = IMPORT_DESC;
    } else {
      /* 非聚焦：占位已升级「待保存」，无需额外提示 */
    }
    saveImportViewCache(cachePayload);
  } catch (e) {
    stopReadTimer();  // 失败/放弃：停读取计时器
    stopLlmTimer();   // 失败/放弃：停 LLM 计时器（行已随视图重建脱离文档，不再 tick）
    // 错误不整页替换：整页替换会抹掉过程行，排障信息全丢。
    // 过程面板还在 → 保留全部过程行，末尾追加 ✗ 行 + 错误卡；面板不在（兜底）→ 退回整页错误卡。
    const showImportFail = (title, msg, canRetry) => {
      if (logEl && logEl.isConnected) {
        mkLine("✗", title);
        bodyEl.insertAdjacentHTML("beforeend",
          `<div class="imp-err"><b>${esc(title)}</b><br>${esc(msg)}<br><span class="imp-err-tip">（${canRetry ? "可点击「重试」重新提取本文件" : "左侧菜单中的占位可删除，或重新导入"}）</span></div>`);
      } else {
        bodyEl.innerHTML = `<div class="imp-err"><b>${esc(title)}</b><br>${esc(msg)}</div>`;
      }
      okBtn.style.display = "";
      okBtn.disabled = false;
      okBtn.textContent = canRetry ? "重试" : "重新导入";
      cancelBtn.textContent = "放弃";
      cancelBtn.disabled = false;
      cancelBtn.onclick = () => { switchView("convert"); };
    };
    if (e && e.name === "AbortError") {
      // 用户「停止」：中止上游生成，占位标记取消
      if (_canceledImports.has(phKey)) { _canceledImports.delete(phKey); return; }   // 用户「关闭」取消
      failImportPlaceholder(phKey, "已取消");
      if (bodyEl.dataset.impKey === phKey) {
        showImportFail("已放弃导入", "模板提取已被停止，未生成模板。", true);
        okBtn.onclick = () => {   // 重试：移除占位后同一文件重走导入流程
          removeImportPlaceholder(phKey);
          doImport(file, tsFile);
        };
      }
      return;
    }
    if (_canceledImports.has(phKey)) { _canceledImports.delete(phKey); return; }  // 用户已取消，不标记失败
    failImportPlaceholder(phKey, e.message);
    if (bodyEl.dataset.impKey === phKey) {
      showImportFail("导入失败", e.message, true);
      okBtn.onclick = () => {   // 失败重试：移除失败占位后，同一文件重新走导入流程
        removeImportPlaceholder(phKey);
        doImport(file, tsFile);
      };
    }
  } finally {
    stopBtn.style.display = "none";
  }
}

// 将导入模板落盘到指定模板集（tsFile 空 = 当前集 / 弹框新建集），并刷新左侧菜单树
// 统一走 writeTemplateToSet：目标集=当前内存集直接落盘，目标集=其他集静默读改写
// 用户主动保存 → 切为当前编辑模板；返回落盘的文件名（字符串），取消返回 false。
async function saveImportedToSet(tpl, fileName, tsFile, extraDefs) {
  return writeTemplateToSet(tpl, fileName, tsFile, true, extraDefs);
}

// 保存后切到「模板编辑」状态：加载目标模板集并把新模板置为当前编辑模板，
// 展开左侧所在模板集并高亮新模板行（配合 switchView("design") 进入设计器）。
async function openImportedTemplateInDesign(file, tplKey) {
  if (!file) return;
  if (file !== state.currentFile || !state.templateSet) {
    await loadTemplateSet(file);   // 内部先 flushSave 当前集，再读取目标集
  }
  const tpl = state.templateSet && state.templateSet.templates.find(t => t.key === tplKey);
  if (tpl) {
    state.currentTemplateKey = tpl.key;
    state.selectedElem = null;
    renderAll();
  }
  uiState.expandedSets.add(file);  // 展开所在集，让新模板在左侧菜单可见
  switchView("design");            // 进入模板编辑状态（设计视图）
  // 强制从服务端拉取最新状态再渲染菜单：刚保存的模板集/模板必然出现在列表，
  // 避免依赖可能过期的 _lastTreeStatus 缓存导致「菜单没刷新、新模板不显示」
  // （典型场景：导入落到新建集时，writeTemplateToSet 不会把新集补进缓存）。
  try {
    const st = await refreshStatus();
    if (st) renderSetTree(st);
  } catch (e) {
    if (_lastTreeStatus) renderSetTree(_lastTreeStatus);
  }
}

// ═══════════════════════════════════════════
// 设置页：LLM 服务配置（多服务 + 多模型，OpenAI 兼容）
// ═══════════════════════════════════════════
let llmServices = [];       // 脱敏后的服务列表（自配）
let llmEnabledId = "";      // 当前启用的自配服务 id（单一启用位；空 = 未启用任何自配服务）
let llmBuiltin = null;      // 内置服务状态 { available, using, name, model, state }，不含 Key
let llmUseBuiltin = true;   // 内置服务启用开关（llms.json 顶层 useBuiltin）
let _llmFormState = null;   // 表单弹窗状态 { isNew, id, keyShown, dirty, discardArmed, preset, keyReady }

async function loadLlmList() {
  try {
    const r = await api("/api/llms");
    llmServices = r.services || [];
    llmEnabledId = r.defaultServiceId || "";
    llmBuiltin = null;
    llmUseBuiltin = false;
    state.nodeUseEnvProxy = !!r.nodeUseEnvProxy;
    state.systemProxyEnabled = !!r.systemProxyEnabled;
    renderLlmList();
    renderLlmProxyStatus();
  } catch (e) {
    toast("加载 LLM 配置失败：" + e.message, "err");
  }
}

// ═══════════════════════════════════════════
// 内置服务卡已移除（登录功能已禁用）。保留空实现避免历史引用报错。
// ═══════════════════════════════════════════
function renderBuiltinCard() {
  const box = document.getElementById("builtin-card");
  if (box) box.innerHTML = "";
}

/** 内置服务当前是否生效：已移除，恒 false */
function llmBuiltinOn() {
  return false;
}

/** 是否存在已启用且可用的模型：自配（当前启用的那一个）。
 * 全站"是否已配置模型"的闸门统一走这里——开关是唯一判据（关掉即视为未配置）。 */
function llmConfigured() {
  return llmServices.some(s => s.id === llmEnabledId);
}

/** 内置服务启用开关：已移除，保留空实现 */
async function onBuiltinToggle(on) {
  // 无内置服务，不做任何事
}

// 内置卡动作：已移除，保留空实现
async function onBuiltinAction(act, btn) {
  // 无内置服务，不做任何事
}

/** 渲染代理通道提示（常显，描述连通性，与 LLM 能力开关无关）。
 * 三态：系统代理未启用（灰）/ 系统代理启用 + Node 开关启用（绿）/ 启用但开关未启用（红） */
function renderLlmProxyStatus() {
  const el = document.getElementById("llm-proxy-diag");
  if (!el) return;
  el.classList.remove("on", "off", "neutral");
  if (!state.systemProxyEnabled) {
    el.classList.add("neutral");
    el.textContent = "代理状态：未启用系统代理（直连）";
  } else if (state.nodeUseEnvProxy) {
    el.classList.add("on");
    el.textContent = "代理状态：✓ 代理通道就绪";
  } else {
    el.classList.add("off");
    el.innerHTML = "代理状态：⚠ 系统代理已启用但 Node 开关未启用，公网可能连接失败 → 完全退出后重开应用";
  }
}

function llmKeyText(svc) {
  return svc.apiKey && svc.apiKey.hasKey
    ? "Key ****" + svc.apiKey.keyTail
    : "未填 Key（免鉴权网关适用）";
}

function renderLlmList() {
  const box = document.getElementById("llm-list");
  if (!box) return;
  if (!llmServices.length) {
    box.innerHTML = "";
    return;
  }
  box.innerHTML = llmServices.map(svc => {
    const keyText = svc.apiKey && svc.apiKey.hasKey ? ("****" + svc.apiKey.keyTail) : "未填 Key";
    const model = svc.models && svc.models.length ? (svc.models.find(m => m.isDefault) || svc.models[0]) : null;
    const modelText = model ? model.name : "（未设置）";
    const on = svc.id === llmEnabledId;
    return `
    <div class="llm-card${on ? " is-on" : ""}">
      <div class="llm-card-head">
        <span class="llm-card-name">${esc(svc.name)}</span>
        <span class="builtin-toggle">
          <span class="builtin-toggle-label">启用</span>
          <label class="switch" data-sw="${esc(svc.id)}" title="启用后同一时刻不再使用其他模型服务；关闭即停用该服务">
            <input type="checkbox" ${on ? "checked" : ""}>
            <span class="slider"></span>
          </label>
        </span>
      </div>
      <div class="llm-meta">
        <div class="llm-meta-item" title="${esc(modelText)}"><code>${esc(modelText)}</code></div>
        <div class="llm-meta-item" title="${esc(svc.baseUrl)}"><code>${esc(svc.baseUrl)}</code></div>
        <div class="llm-meta-item"><code>${esc(keyText)}</code></div>
        <div class="llm-ops">
          <button class="small" data-act="test" data-id="${esc(svc.id)}">测试</button>
          <button class="small" data-act="edit" data-id="${esc(svc.id)}">编辑</button>
          <button class="small danger" data-act="del" data-id="${esc(svc.id)}">删除</button>
        </div>
      </div>
      <div class="test-result" data-res="${esc(svc.id)}" style="display:none"></div>
    </div>`;
  }).join("");
  box.querySelectorAll("button[data-act]").forEach(btn => {
    btn.addEventListener("click", () => onLlmAction(btn.dataset.act, btn.dataset.id, btn));
  });
  box.querySelectorAll("label.switch[data-sw]").forEach(lb => {
    const input = lb.querySelector("input");
    input.addEventListener("change", () => onLlmToggle(lb.dataset.sw, input.checked));
  });
}

/** 启用/停用某个自配服务（单一启用位：启用 = 自动停用其他服务与内置）。
 * 停用当前启用的那个 = 此后不启用任何模型服务，需再启用一个。 */
async function onLlmToggle(id, on) {
  try {
    await api("/api/llms/" + encodeURIComponent(id) + "/enable", { method: "PUT", body: JSON.stringify({ on }) });
    await loadLlmList();    // 重画全部卡：互斥关系变了，内置卡与其余自配卡都要跟着变
    toast(on ? "已启用该模型服务" : "已停用该模型服务");
  } catch (e) {
    toast("切换失败：" + e.message, "err");
    loadLlmList();          // 回滚 UI 到服务端实际状态
  }
}

// ═══════════════════════════════════════════
// 默认样式（引擎种子）：设置页只读展示
// 语义：种子 = 新建模板时深拷贝进模板的出厂基准（config/style-seeds.json，出厂唯一源），
//       不回溯存量模板；编辑能力统一收归「模板设计」页，本页无任何写回入口。
// ═══════════════════════════════════════════
let seedStyles = [];        // 内存态种子 styles（只读展示用）
let fontFamilies = [];      // 内存态字体族映射表（只读展示用；config/fonts.json）
let appSettings = {};       // 应用级开关（用户态 settings.json；forceMappedFont 等）
// 只读清单表头：三组样式清单共用同一口径（名称 + 默认格式），保证视觉一致；字体映射组是三列表，自带表头
const SEED_GRID_HEAD = `<thead><tr><th>名称</th><th>格式</th></tr></thead>`;

async function loadSeedStyles() {
  try {
    const r = await api("/api/style-seeds");
    seedStyles = Array.isArray(r.styles) ? JSON.parse(JSON.stringify(r.styles)) : [];
    renderSeedList();
  } catch (e) {
    seedStyles = [];
    toast("加载默认样式失败：" + e.message, "err");
  }
}

/** 读字体族映射表（config/fonts.json）+ 应用级开关（用户态 settings.json），渲染「默认样式」卡
 *  （样式清单与字体映射组一次成型，故只需一次 renderSeedList）。
 *  任一失败不打断设置页其余内容，只让该组不渲染 / 开关走默认值；
 *  数据侧走 ensureFontMapping（与「种子 → 模板」写入路径共用一份，幂等）。 */
async function loadFontFamilies() {
  await ensureFontMapping();
  renderSeedList();
  bindSeedFontSwitch();
}

/** 写模板时的字体映射（与服务端 resolveFontForWrite 同口径，前端只用于「新建样式」预填默认值）。
 *  开（默认）→ 取「模板映射」列（方正字体）；关 → 取「系统默认」列（Windows 中文名）。
 *  非映射表内的字体名原样返回；映射表未加载（fontFamilies 空）时不改写，避免误伤。 */
function resolveFontForWrite(name) {
  const n = String(name == null ? "" : name).trim();
  if (!n || !fontFamilies.length) return name;
  const fam = fontFamilies.find(f => (f.key || "").trim() === n || (f.label || "").trim() === n);
  if (!fam) return name;
  const target = (appSettings.forceMappedFont !== false) ? fam.mapped : fam.system;
  return target || name;
}

// ═══════════════════════════════════════════
// 设置页「人员名单」（「人员姓名校对」的核对依据）
// ═══════════════════════════════════════════
// 名单是**纯用户数据**：可自由增删改，无数量上限、无预设分类（按部门/按会议/就一份全单位都由用户定）。
// 与出厂基准无关 —— 没有哪份名单是所有用户共用的默认值，故无出厂文件、无覆盖层、无「恢复出厂」。
// 保存是**整体写回**（与术语表同范式）：界面持有全量清单，改完一次提交，服务端校验后落盘。
// ⚠ 与转换页的名单选择区共用 rosterCache，任何保存后两处都要刷新，避免"设置页加了名单、
//   转换页选择区看不见"。

let rosterNewSeq = 0;        // 新建名单的临时 id 序号（防同一毫秒内连点撞 id）

function newRosterId() {
  rosterNewSeq += 1;
  return "r" + Date.now().toString(36) + "_" + rosterNewSeq;
}

/** 拉名单并渲染设置页 + 转换页选择区（两处同源） */
async function loadRosterSettings() {
  try {
    const r = await api("/api/rosters");
    rosterCache = (r && Array.isArray(r.rosters)) ? r.rosters : [];
  } catch (e) { rosterCache = null; }
  renderRosterSettings();
  renderRosterPick();
}

/** 整体提交名单；成功 → 刷新内存缓存并返回 true */
async function postRosters(list) {
  // ★ 整体写回的前提是界面持有**全量**清单。rosterCache === null 表示清单读取失败，
  //   此时界面手里只有残片，提交上去等于用残片覆盖服务端 —— 用户的名单会被整批抹掉。
  //   宁可不保存，让用户刷新后重试。
  if (rosterCache === null) { toast("名单读取失败，请刷新页面后重试", "err"); return false; }
  try {
    const r = await api("/api/rosters", { method: "POST", body: JSON.stringify({ rosters: list }) });
    if (r && r.ok) { rosterCache = Array.isArray(r.rosters) ? r.rosters : []; return true; }
    toast((r && r.error) || "保存失败", "err");
    return false;
  } catch (e) {
    toast("保存失败：" + e.message, "err");
    return false;
  }
}

function renderRosterSettings() {
  const box = document.getElementById("roster-list");
  if (!box) return;
  if (rosterCache === null) {
    box.innerHTML = '<div class="roster-empty" style="color:var(--warn)">名单读取失败，请刷新页面重试</div>';
    return;
  }
  if (!rosterCache.length) {
    // 空态不给文案：卡页头已说明这份名单是干什么用的，「＋ 新建名单」按钮就在列表下方，
    // 再补一段操作引导是重复信息。只有**读取失败**才需要落到这里（那是用户该知道、也必须处理的）。
    box.innerHTML = "";
    return;
  }
  // 一行一份：名称 + 人数 + 操作（操作靠右），姓名全文放编辑弹窗里看/改，不在本页铺开。
  // 不再标「含空格」——名单侧的空格在保存时已被删干净（见 server/names.js cleanName），
  // 勾选这里也没有「脏姓名」可言了。
  box.innerHTML = rosterCache.map(r => {
    const id = String(r.id || "");
    const members = Array.isArray(r.members) ? r.members : [];
    const warn = members.length ? "" : '<span class="roster-row-warn">未填姓名</span>';
    return '<div class="roster-row" data-roster-id="' + esc(id) + '">'
      + `<span class="roster-row-name">${esc(String(r.name || ""))}</span>`
      + `<span class="roster-row-count">${members.length} 人</span>`
      + warn
      + '<span class="roster-row-ops">'
      + `<button type="button" class="small" data-roster-edit="${esc(id)}">编辑</button>`
      + `<button type="button" class="small danger" data-roster-del="${esc(id)}">删除</button>`
      + '</span></div>';
  }).join("");
}

/** 打开名单编辑弹窗。
 *  draft 传入时是**新建草稿**：只填进弹窗，不进列表也不落盘，点「保存」才写入；
 *  取消 / 关窗（弹窗带输入控件，Esc 与点窗外都不生效，只能走按钮）即丢弃 —— 绝不留下半成品空名单。
 *  判据：新建是「打开一个编辑框」这个动作本身，不是「往列表里塞一条」。 */
async function openRosterEditor(id, draft) {
  const key = String(id || "");
  const isNew = !!draft;
  const r = isNew
    ? { id: key, name: String(draft.name || ""), members: Array.isArray(draft.members) ? draft.members : [] }
    : (rosterCache || []).find(x => String(x.id) === key);
  if (!r) return;
  const members = Array.isArray(r.members) ? r.members : [];
  // 人数挂在字段名上（不是另起一行「解析为 N 人」）：用户读字段名时就顺带看到条数，
  // 不必再多看一行；姓名的逐条罗列也一并去掉 —— 真正需要用户判断的只有「哪些写法被改了」，
  // 那是下面的空格告警，罗列一遍姓名不提供新信息。
  const bodyHtml =
    '<div class="roster-edit-row"><label>名单名称</label>'
    + `<input type="text" id="roster-edit-name" value="${esc(String(r.name || ""))}" placeholder="如：综合办公室人员"></div>`
    + '<div class="roster-edit-row"><label id="roster-edit-names-label">姓名列表（' + members.length + ' 人）</label>'
    + `<textarea id="roster-edit-names" placeholder="张三&#10;李四&#10;王五">${esc(members.join("\n"))}</textarea></div>`
    + '<div class="roster-edit-hint">多个姓名之间用<b>顿号、逗号、分号或换行</b>分开。'
    + '姓名里不要留空格 —— 保存时会全部删掉。</div>'
    + '<div class="roster-parse-warn" id="roster-parse-warn" hidden></div>';
  // hideButtons：原生 ok/cancel 已被 openModal 隐藏，本模块自己追加「保存」主按钮（保存中可禁用）
  const p = openModal({ title: isNew ? "新建名单" : "编辑名单", body: bodyHtml, hideButtons: true, modalClass: "roster-editor-modal" });
  const actionsEl = document.querySelector("#modal-overlay .modal-actions");
  if (actionsEl) {
    actionsEl.style.display = "";
    // ⚠ 只 **追加** 自建按钮（带 data-modal-custom 标记，由 openModal / closeModal 收尾时摘掉），
    //   不能拿 innerHTML 覆盖 —— 那会把原生 ok/cancel 节点一起销毁，连带丢掉 bindModalEvents()
    //   绑在它们身上的 click 监听，之后全站「确定 / 删除」都点不动。
    actionsEl.insertAdjacentHTML("beforeend",
      '<button type="button" class="small" data-modal-custom id="roster-dlg-cancel">取消</button>'
      + '<button type="button" class="small primary" data-modal-custom id="roster-dlg-save">保存</button>');
  }
  const nameEl = document.getElementById("roster-edit-name");
  const textEl = document.getElementById("roster-edit-names");
  const namesLabelEl = document.getElementById("roster-edit-names-label");
  const warnEl = document.getElementById("roster-parse-warn");
  // 输入即刷新字段名上的人数，并把**改过写法的姓名**列出来（改了哪几处、原样是什么）。
  // 删空格是确定的，但「一个空格是姓名之间的分隔、还是姓与名之间的排版噪声」只有用户知道，
  // 代码不猜；改为把改动摊开：`张三　李四` 被并成 `张三李四` 时，用户对着「原写法 → 清理后」
  // 一眼就能看出原本是两个名字。去抖 400ms，避免逐字请求。
  let auditTimer = null;
  const runAudit = async () => {
    if (!textEl) return;
    const text = String(textEl.value || "");
    let names = [];
    let audit = null;
    try {
      const res = await api("/api/rosters/parse", { method: "POST", body: JSON.stringify({ text }) });
      names = (res && Array.isArray(res.names)) ? res.names : [];
      audit = res && res.audit;
    } catch (e) { return; }   // 解析失败静默（不阻塞输入；保存时会再解析一次并明确报错）
    // 人数以解析结果为准（可能与已保存的条数不同 —— 用户正在改）
    if (namesLabelEl) namesLabelEl.textContent = "姓名列表（" + names.length + " 人）";
    if (!warnEl) return;
    const lines = [];
    const cleaned = (audit && audit.cleaned) || [];
    const tooLong = (audit && audit.tooLong) || [];
    if (cleaned.length) {
      const shown = cleaned.slice(0, 6).map(c => "「" + c.from + "」→「" + c.to + "」").join("、");
      lines.push("这些写法带空格，已删掉：" + shown + (cleaned.length > 6 ? " 等" : "")
        + "。若某处本是两个姓名，请改用顿号分开。");
    }
    if (tooLong.length) {
      lines.push("有 " + tooLong.length + " 条超过 20 字，已被忽略：多半是漏了分隔符，请用顿号把姓名分开。");
    }
    if (!lines.length) { warnEl.hidden = true; warnEl.textContent = ""; return; }
    warnEl.textContent = lines.join("\n");   // 多行：CSS white-space: pre-line
    warnEl.hidden = false;
  };
  if (textEl) {
    textEl.addEventListener("input", () => {
      if (auditTimer) clearTimeout(auditTimer);
      auditTimer = setTimeout(runAudit, 400);
    });
    setTimeout(runAudit, 60);   // 打开即解析一次（已保存的名单也过一遍，保证显示与保存结果一致）
  }
  setTimeout(() => { if (nameEl) { nameEl.focus(); nameEl.select(); } }, 50);
  const close = (val) => {
    if (auditTimer) clearTimeout(auditTimer);
    if (actionsEl) actionsEl.querySelectorAll("[data-modal-custom]").forEach(n => n.remove());
    closeModal(val);
  };
  document.getElementById("roster-dlg-cancel").onclick = () => close(null);
  document.getElementById("roster-dlg-save").onclick = async () => {
    const name = String((nameEl && nameEl.value) || "").trim();
    const text = String((textEl && textEl.value) || "");
    if (!name) { toast("请填写名单名称", "err"); if (nameEl) nameEl.focus(); return; }
    let parsed;
    try {
      const res = await api("/api/rosters/parse", { method: "POST", body: JSON.stringify({ text }) });
      parsed = (res && Array.isArray(res.names)) ? res.names : [];
    } catch (e) { toast("姓名解析失败：" + e.message, "err"); return; }
    // 落盘：已在列表里 → 改这一条；新建草稿 → 追加一条。二者都走整体写回。
    const exists = (rosterCache || []).some(x => String(x.id) === key);
    const next = exists
      ? (rosterCache || []).map(x => String(x.id) === key ? { ...x, name, members: parsed } : x)
      : (rosterCache || []).concat([{ id: key, name, members: parsed }]);
    if (!(await postRosters(next))) return;   // 失败保留弹窗，用户不必重打
    close(true);
    await loadRosterSettings();
    updateRosterPickHint();
    toast(`${isNew ? "已新建" : "已保存"}「${name}」，共 ${parsed.length} 人`);
  };
  await p;
}

/** 删除一份名单（需确认：名单是用户手工维护的数据，误删要重打） */
async function deleteRoster(id) {
  const key = String(id || "");
  const r = (rosterCache || []).find(x => String(x.id) === key);
  if (!r) return;
  // 用项目自建确认框（modalConfirm），不用原生 window.confirm：内置 WebView 下原生对话框样式与站点脱节
  const ok = await modalConfirm("删除名单", `确定删除名单「${esc(r.name)}」？该名单将从转换页的可选列表中移除。`, "删除");
  if (!ok) return;
  const next = (rosterCache || []).filter(x => String(x.id) !== key);
  if (!(await postRosters(next))) return;
  // 选择区存档里可能还留着被删的 id，交给 renderRosterPick 按现存名单自然淘汰
  await loadRosterSettings();
  updateRosterPickHint();
  toast(`已删除名单「${r.name}」`);
}

/** 绑定设置页名单行的事件（一次性；内容靠 renderRosterSettings 整体重渲染）。
 *  ⚠ modal-actions 里那些按钮是本模块自建的（原生 ok/cancel 在 hideButtons 下被隐藏、但**节点保留**），
 *    自建按钮带 data-modal-custom 标记，由 openModal / closeModal 自动摘除。 */
function setupRosterSettings() {
  const box = document.getElementById("roster-list");
  if (box) {
    box.addEventListener("click", ev => {
      const t = ev.target;
      if (!t || !t.dataset) return;
      if (t.dataset.rosterEdit !== undefined) { openRosterEditor(t.dataset.rosterEdit); return; }
      if (t.dataset.rosterDel !== undefined) { deleteRoster(t.dataset.rosterDel); }
    });
  }
  const add = document.getElementById("btn-add-roster");
  if (add) {
    add.addEventListener("click", () => {
      // 新建 = 打开一份**草稿**（不落盘）：点「保存」才进列表，取消 / 关窗就当没建过。
      // 早先是「点了就落盘、关掉弹窗也留一条」，用户会攒出一串 0 人的「新名单」。
      openRosterEditor(newRosterId(), { name: "", members: [] });
    });
  }
}

/** 渲染只读种子清单：段落（含段内字符格式）/ 表格 / 页面 / 字体映射 四组，只展示名称与默认格式，不可编辑。
 *  字体映射组与清单同卡（renderFontMap）：它是「清单里的字族落到哪个具体字体」的取值规则，
 *  与清单是同一件事的两半；数据来自另一份出厂文件（config/fonts.json），作用域也不止新建模板，
 *  故口径说明写在组内，卡页头不重复。 */
function renderSeedList() {
  const box = document.getElementById("seed-list");
  if (!box) return;
  box.innerHTML = renderSeedTable("段落样式", ["text", "char"])
    + renderSeedTable("表格样式", ["table"])
    + renderSeedPage()
    + renderFontMap();
}

/** 只读样式表：名称列 + 默认格式摘要列；cats = 本表纳入的 category 列表。
 * 注：种子的 group 字段（table 类二级分组）当前只有「数据表格」一个取值、与组名重复，
 *     故不渲染二级标题行；将来出现多分组时再按 group 分段。 */
function renderSeedTable(label, cats) {
  const items = seedStyles.filter(s => cats.includes(s.category));
  if (!items.length) return "";
  const rows = items
    .map(s => `<tr><td class="seed-key">${esc(s.key)}</td><td class="seed-val">${esc(seedSummaryOf(s))}</td></tr>`)
    .join("");
  return `<div class="seed-group">
    <div class="seed-group-label">${label}（${items.length}）</div>
    <table class="seed-grid">${SEED_GRID_HEAD}<tbody>${rows}</tbody></table>
  </div>`;
}

/** 字族语义名 → 国标叫法（种子里的「小标宋」→ 展示「小标宋体」）。
 *  只用于只读展示；映射表未加载、或值不是字族时原样返回。 */
function seedFontFamilyLabel(font) {
  const n = String(font == null ? "" : font).trim();
  if (!n) return "";
  const fam = fontFamilies.find(f => (f.key || "").trim() === n || (f.label || "").trim() === n);
  return (fam && fam.label) || n;
}

/** 种子条目 → 只读格式摘要（中文口径；不暴露 next / match 等引擎内部字段）。
 *  字体列展示种子存的**字族语义名**（小标宋体 / 黑体 …）：本层只定「用哪个字族」，
 *  落到哪个具体字体由「字体映射」按开关决定（resolveFontForWrite）——此处解析等于把映射
 *  结果抄进本层，界面显示与开关脱钩，也就看不出这层存的是字族。 */
function seedSummaryOf(s) {
  const d = s.defaults || {};
  const parts = [];
  if (d.font) parts.push(seedFontFamilyLabel(d.font));
  if (d.size) parts.push(ptToCnSize(d.size));
  if (d.bold) parts.push("加粗");
  if (s.category !== "char") {
    const al = (ALIGNS.find(a => a.v === d.align) || {}).label;
    if (al) parts.push(al);
    if (d.line) parts.push(`行距 ${d.line} 磅`);
    if (d.indentChars) parts.push(`首行缩进 ${d.indentChars} 字`);
    if (d.rightIndentChars) parts.push(`右缩进 ${d.rightIndentChars} 字`);
  }
  return parts.join(" · ") || "默认";
}

/** 页面样式（只读）：纸张 / 版面 / 页边距 / 页码四行，与其余两组同构（表格 + 表头） */
function renderSeedPage() {
  const items = seedStyles.filter(s => s.category === "page");
  if (!items.length) return "";
  const rows = items.map(s => {
    const d = s.defaults || {};
    const mm = v => (v == null || v === "") ? "—" : Math.round(v / 1440 * 25.4);
    const al = (ALIGNS.find(a => a.v === d.pageAlign) || {}).label || "";
    const pageNum = [d.pageFont, ptToCnSize(d.pageSize), al].filter(Boolean).join(" · ");
    // 版面：版心高 ÷ 行网格间距（国标「一般每面 22 行」；行网格见 page.linePitch）
    const lines = d.linePitch ? Math.round((d.height - d.marginTop - d.marginBottom) / d.linePitch) : null;
    return [
      ["纸张", `${mm(d.width)} × ${mm(d.height)} mm`],
      ["版面", lines ? `每面 ${lines} 行` : "—"],
      ["页边距", `上 ${mm(d.marginTop)} · 下 ${mm(d.marginBottom)} · 左 ${mm(d.marginLeft)} · 右 ${mm(d.marginRight)} mm`],
      ["页码", [pageNum, `距版心下边缘 ${d.numberGap != null ? d.numberGap : DEFAULT_NUMBER_GAP} mm`].filter(Boolean).join(" · ")],
    ].map(([k, v]) => `<tr><td class="seed-key">${esc(k)}</td><td class="seed-val">${esc(v)}</td></tr>`).join("");
  }).join("");
  return `<div class="seed-group">
    <div class="seed-group-label">页面样式</div>
    <table class="seed-grid">${SEED_GRID_HEAD}<tbody>${rows}</tbody></table>
  </div>`;
}

/** 字体映射（只读表 + 一个开关）：国标规定字族 → 方正字库（开关开） / 系统自带字库（开关关）。
 *  口径：国标 GB/T 9704 只规定字族（小标宋/黑体/楷体/仿宋/宋体），不指定字体文件；
 *        出厂种子与模板里存的具体字体名都按本表取值。
 *  开关作用于所有「系统替用户填字体值」的时刻——新建模板、从默认样式中选、重置样式、
 *  模板导入、手工新建样式；已有模板里的字体名与文档排版流程一律不受影响。
 *  渲染为「默认样式」卡内第四组（#font-map-group）：开关决定该清单里的字族落到哪个字体，
 *  两件事必须同屏才讲得清；但它属另一份出厂文件（config/fonts.json）、作用域也大于新建模板，
 *  故口径说明写在组内，不在卡页头。
 *  表头：单行主标 + 逐列 tooltip（列长说明放 title，不站版面）。 */
function renderFontMap() {
  if (!fontFamilies.length) return "";
  const rows = fontFamilies.map(f =>
    `<tr><td class="seed-key">${esc(f.label || f.key)}</td><td class="seed-val">${esc(f.mapped || "—")}</td><td class="seed-val">${esc(f.system || "—")}</td></tr>`
  ).join("");
  const on = appSettings.forceMappedFont !== false;
  // 提示语常驻不随开关变化：一句话讲清开关两侧各用什么、管到哪些时刻
  const note = "开启使用方正字库，关闭使用 Windows 系统自带字库；只在新建模板、导入模板、新建样式时决定字体取值，不影响已有模板";
  // 开关放标签行右侧（右对齐）：沿用全站唯一 .switch 组件（36×20，:checked 驱动）
  return `<div class="seed-group" id="font-map-group">
    <div class="seed-group-label seed-group-label-row">字体映射（${fontFamilies.length}）
      <span class="seed-group-note">${esc(note)}</span>
      <label class="switch seed-label-switch" title="开启：新建模板 / 导入模板 / 新建样式时使用「方正字库」列字体；关闭：使用「系统自带字库」列字体。已有模板里的字体与文档排版流程不受影响">
        <input type="checkbox" id="chk-force-font" ${on ? "checked" : ""}>
        <span class="slider"></span>
      </label>
    </div>
    <table class="seed-grid">
      <thead><tr>
        <th title="《党政机关公文格式》（GB/T 9704-2012）只规定到字族，不指定字体文件">国标规定字族</th>
        <th title="开关开启时使用">方正字库</th>
        <th title="开关关闭时使用">系统自带字库</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

/** 开关联动：保存到用户态 settings.json。
 *  事件委托挂在容器 #seed-list 上（不挂在开关元素本身）：清单与字体映射组由同一处 innerHTML 整体重渲染，
 *  重渲染会换掉开关 DOM ⇒ 直接绑元素会随重渲染失效（清单渲染与映射表加载是两条并发请求，先后不定）。
 *  容器本身不被替换，故委托只绑一次（dataset 标记防重复）。
 *  「默认样式」清单不随开关重渲染——它展示字族语义名，与开关取值无关。 */
function bindSeedFontSwitch() {
  const box = document.getElementById("seed-list");
  if (!box || box.dataset.fontSwitchBound) return;
  box.dataset.fontSwitchBound = "1";
  box.addEventListener("change", async e => {
    const cb = e.target;
    if (!cb || cb.id !== "chk-force-font") return;
    const on = cb.checked;
    appSettings.forceMappedFont = on;
    // 提示语常驻（同时讲开/关两侧），不随开关变化 ⇒ 这里无需就地改文案
    try {
      await api("/api/settings", { method: "POST", body: JSON.stringify({ forceMappedFont: on }) });
      toast(on ? "已开启：使用方正字库" : "已关闭：使用 Windows 系统自带字库", "ok");
    } catch (e) {
      // 落盘失败 → 回滚勾选态（提示语常驻，无需回滚）
      appSettings.forceMappedFont = !on;
      cb.checked = !on;
      toast("保存失败：" + e.message, "err");
    }
  });
}

/** 成功时展示的模型清单（可滚动 chip）。
 * 表单内（forForm=true）额外带搜索框；设置页服务卡只展示清单、无搜索框。 */
function testModelsHtml(r, forForm) {
  if (r.ok && r.via === "models" && Array.isArray(r.models) && r.models.length) {
    const chips = r.models
      .map(m => `<span class="chip" data-model="${esc(m)}"${forForm ? ` title="双击填入模型名称"` : ""}>${esc(m)}</span>`)
      .join("");
    const search = forForm
      ? `<input type="text" class="test-models-search" placeholder="筛选模型（共 ${r.models.length} 个）">`
      : "";
    return `<div class="test-models-wrap${forForm ? "" : " no-form"}">${search}<div class="test-models">${chips}</div></div>`;
  }
  return "";
}

function showTestResult(el, r) {
  if (!el) return;
  el.style.display = "";
  el.className = "test-result " + (r.ok ? "ok" : "err");
  const forForm = el.id === "llm-form-test-result";   // 仅「添加/编辑模型」表单才有模型名输入框
  if (r.ok) {
    // 成功：结论一行，下方模型清单（表单内带搜索框 + 双击自动填入「模型名称」）
    el.innerHTML = `<div class="test-main">✓ ${esc(r.message || "连接成功")}</div>${testModelsHtml(r, forForm)}`;
    const modelInput = document.getElementById("llm-model");   // 仅添加/编辑表单存在
    const search = el.querySelector(".test-models-search");
    if (search) {
      search.addEventListener("input", () => {
        const q = search.value.trim().toLowerCase();
        el.querySelectorAll(".test-models .chip").forEach(c => {
          const name = (c.dataset.model || "").toLowerCase();
          c.style.display = (!q || name.includes(q)) ? "" : "none";
        });
      });
    }
    if (modelInput) {
      el.querySelectorAll(".test-models .chip").forEach(c => {
        c.addEventListener("dblclick", () => {
          modelInput.value = c.dataset.model || c.textContent;
          _llmFormState.dirty = true;
          c.classList.add("selected");
          setTimeout(() => c.classList.remove("selected"), 300);
        });
      });
    }
  } else {
    // 失败：结论 + 代理 reason/告警（不再折叠诊断）
    let html = `<div class="test-main">✗ ${esc(r.message || "连接失败")}</div>`;
    if (r.proxy && r.proxy.reason) html += `<div class="test-proxy">${esc(r.proxy.reason)}</div>`;
    if (r.proxy && typeof r.proxy.nodeUseEnvProxy === "boolean" && !r.proxy.nodeUseEnvProxy) {
      html += `<div class="test-proxy-warn">⚠ Node 代理开关未启用 → 若要连接公网请完全重启应用（重启应用）</div>`;
    }
    el.innerHTML = html;
  }
}

async function onLlmAction(act, id, btn) {
  if (act === "edit") { editLlmModal(id); return; }
  if (act === "del") {
    const svc = llmServices.find(s => s.id === id);
    const ok = await modalConfirm("删除 LLM 服务", `确定删除「${esc(svc ? svc.name : "")}」？该操作不可恢复。`);
    if (!ok) return;
    try {
      await api("/api/llms/" + encodeURIComponent(id), { method: "DELETE" });
            loadLlmList();
    } catch (e) { toast("删除失败：" + e.message, "err"); }
    return;
  }
  if (act === "test") {
    const resEl = document.querySelector(`[data-res="${CSS.escape(id)}"]`);
    btn.disabled = true;
    const orig = btn.textContent;
    btn.textContent = "测试中…";
    try {
      const r = await api("/api/llms/" + encodeURIComponent(id) + "/test", { method: "POST", body: "{}" });
      showTestResult(resEl, r);
    } catch (e) {
      showTestResult(resEl, { ok: false, message: e.message });
    } finally {
      btn.disabled = false;
      btn.textContent = orig;
    }
  }
}

/** LLM 服务商预设（v1.9：2026-08 最新 baseUrl / 模型） */
const LLM_PRESETS = {
  deepseek: {
    label: "DeepSeek",
    // 官方 base_url（2026-08 确认）：不带 /v1（OpenAI 兼容端点即 https://api.deepseek.com）
    baseUrl: "https://api.deepseek.com",
    modelPlaceholder: "deepseek-v4-flash、deepseek-v4-pro",
  },
  bailian: {
    label: "阿里云百炼",
    // 官方 OpenAI 兼容端点（带 /v1，正确）
    baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1",
    modelPlaceholder: "qwen3.8-max、qwen3.7-plus、qwen3.6-flash",
  },
  custom: { label: "自定义", baseUrl: "", modelPlaceholder: "如 deepseek-v4-flash" },
};

/** 收集表单值（模型 = 单输入框值；v1.10：API Key 必填；已删除备注与 chips 列表） */
function llmCollectForm() {
  const model = document.getElementById("llm-model").value.trim();
  // 最大输出 tokens：模型级上限，必须随保存带回（2026-09-11 修复：此前表单没有该字段，
  // 每次保存都把模型上配好的 maxTokens 抹掉 → 后端落 1M 兜底 → 服务端 400 拒收）
  const mtEl = document.getElementById("llm-maxtokens");
  const mt = parseInt(mtEl ? mtEl.value.trim() : "", 10);
  return {
    name: document.getElementById("llm-name").value.trim(),
    baseUrl: document.getElementById("llm-url").value.trim(),
    apiKey: document.getElementById("llm-key").value.trim(),
    models: model
      ? [Object.assign({ name: model, isDefault: true, note: "" },
          Number.isInteger(mt) && mt > 0 ? { maxTokens: mt } : {})]
      : [],
  };
}

/** 校验 URL：http(s):// 开头且 new URL 可解析；非法返回错误文案，合法返回 null */
function llmValidateUrl(url) {
  if (!/^https?:\/\//i.test(url)) return "API 地址需以 http:// 或 https:// 开头";
  try { new URL(url); } catch (e) { return "API 地址格式不合法"; }
  return null;
}

/** 应用服务商预设：填 baseUrl + 模型 placeholder；名称默认跟随服务商名。
 * 跟随条件：名称为空，或等于任一预设名（切预设时同步更新）；
 * 用户自定义名称（非预设名）不覆盖。编辑态服务名同理（自定义名保留）。 */
function applyLlmPreset(key) {
  const p = LLM_PRESETS[key];
  if (!p) return;
  const urlEl = document.getElementById("llm-url");
  const modelEl = document.getElementById("llm-model");
  if (urlEl) urlEl.value = p.baseUrl;
  if (modelEl) modelEl.placeholder = p.modelPlaceholder || "如 deepseek-v4-flash";
  if (key === "custom") {
    if (urlEl) urlEl.value = "";
    // 自定义：清除由预设自动带出的名称（预设名），用户手填的自定义名保留
    const nameEl = document.getElementById("llm-name");
    if (nameEl) {
      const v = nameEl.value.trim();
      const presetNames = Object.values(LLM_PRESETS).map(x => x.label);
      if (presetNames.includes(v)) nameEl.value = "";
    }
  }
  if (key !== "custom" && _llmFormState) {
    const nameEl = document.getElementById("llm-name");
    if (nameEl) {
      const v = nameEl.value.trim();
      const presetNames = Object.values(LLM_PRESETS).map(x => x.label);
      if (!v || presetNames.includes(v)) nameEl.value = p.label;
    }
  }
  _llmFormState.dirty = true;
}

function editLlmModal(id) {
  const isNew = !id;
  const svc = isNew ? null : llmServices.find(s => s.id === id);
  // 模型名 = 单值（默认模型；无默认取第一个）；v1.10 起模型区为单输入框 + 测试后可选
  const initialModel = svc && Array.isArray(svc.models) && svc.models.length
    ? ((svc.models.find(m => m.isDefault) || svc.models[0]).name || "")
    : "";
  // 已配的最大输出 tokens 回填（缺省留空 = 后端 1M 兜底）
  const initialMaxTokens = svc && Array.isArray(svc.models) && svc.models.length
    ? ((svc.models.find(m => m.isDefault) || svc.models[0]).maxTokens || "")
    : "";
  _llmFormState = { isNew, id, keyShown: false, dirty: false, preset: "custom" };

  const keyPh = svc && svc.apiKey && svc.apiKey.hasKey
    ? `已配置 ****${svc.apiKey.keyTail}，留空保持原值`
    : "留空表示免鉴权网关";
  // P2-5：历史模型名 datalist（汇总已保存服务的全部模型名去重）
  const historyModels = Array.from(new Set(
    (llmServices || []).flatMap(s => (s.models || []).map(m => m.name).filter(Boolean))
  ));
  const datalistHtml = historyModels.length
    ? `<datalist id="llm-model-history">${historyModels.map(m => `<option value="${esc(m)}"></option>`).join("")}</datalist>`
    : "";
  const presetsHtml = Object.entries(LLM_PRESETS).map(([k, p]) =>
    `<button type="button" class="preset-chip" data-preset="${k}">${esc(p.label)}</button>`).join("");

  const modalBody = `
    <form id="llm-form" class="llm-form">
      <div class="llm-form-row">
        <label>服务商</label>
        <div class="service-presets">${presetsHtml}</div>
      </div>
      <div class="llm-form-row"><label>名称 <span class="required-star">*</span></label>
        <input type="text" id="llm-name" value="${esc(svc ? svc.name : "")}" placeholder="如 公司内网网关">
      </div>
      <div class="llm-form-row"><label>API 地址 <span class="required-star">*</span> <span class="label-sub">（OpenAI 兼容协议）</span></label>
        <input type="url" id="llm-url" value="${esc(svc ? svc.baseUrl : "")}" placeholder="${esc(LLM_PRESETS.deepseek.baseUrl)}">
      </div>
      <div class="llm-form-row"><label>API Key <span class="required-star">*</span></label>
        <div class="key-row">
          <input type="password" id="llm-key" placeholder="${esc(keyPh)}" autocomplete="new-password">
          <button type="button" class="small" id="llm-key-toggle">显示</button>
          <button type="button" class="small" id="llm-key-copy">复制</button>
        </div>
      </div>
      <div class="llm-form-row"><label>模型名称 <span class="required-star">*</span> <span class="label-sub">点击测试连接可以查询模型名称，双击查询结果即可自动填入</span></label>
        <input type="text" id="llm-model" value="${esc(initialModel)}" placeholder="${esc(LLM_PRESETS.custom.modelPlaceholder)}" list="llm-model-history">
        ${datalistHtml}
      </div>
      <div class="llm-form-row"><label>最大输出 tokens <span class="label-sub">（该模型单次回复上限，按服务商官方数据填；留空按 1M 兜底，超上限的服务商会直接拒收）</span></label>
        <input type="number" id="llm-maxtokens" min="1" step="1" value="${esc(String(initialMaxTokens))}" placeholder="如 393216（384K，DeepSeek 4.1 Flash 上限）">
      </div>
      <div class="llm-form-actions">
        <button type="button" class="small" id="llm-form-test">测试连接</button>
        <span class="test-result" id="llm-form-test-result" style="display:none"></span>
      </div>
    </form>`;
  openModal({ title: isNew ? "添加 LLM 服务" : ("编辑 LLM 服务 · " + (svc ? svc.name : "")), body: modalBody, okText: "保存" });
  // 注意：不能 await openModal —— 其 Promise 直到弹窗关闭才 resolve，
  // await 会让下面的接管与保存绑定失效（点「保存」只关窗不提交）。
  // 必须同步执行：立即接管 modal-ok，让「保存」走自定义提交。
  _styleModalState = { llm: true };

  const keyInput = document.getElementById("llm-key");
  const keyToggle = document.getElementById("llm-key-toggle");
  keyToggle.addEventListener("click", () => {
    _llmFormState.keyShown = !_llmFormState.keyShown;
    keyInput.type = _llmFormState.keyShown ? "text" : "password";
    keyToggle.textContent = _llmFormState.keyShown ? "隐藏" : "显示";
  });
  // 编辑态回填已保存 Key 明文（v1.1 拍板「Key 只在表单编辑态可见」）：
  // 让「显示 / 复制 / 表单内测试」都能用真实 Key（修复 401 鉴权失败与空 Key 问题）。
  // 列表接口一律脱敏，仅编辑单条时经 GET /api/llms/:id 取回明文（本地单机信任边界）。
  // keyReady 存入 state：测试/复制前 await，避免用户在回填完成前误操作（仍按空 Key 测）。
  if (svc && svc.id) {
    _llmFormState.keyReady = api("/api/llms/" + encodeURIComponent(svc.id))
      .then(d => {
        if (d && d.apiKey) {
          keyInput.value = d.apiKey;
          keyInput.placeholder = "sk-xxx（已回填，可直接修改）";
        }
      })
      .catch(() => { /* 静默：取不到明文则保持空 + placeholder 提示已配置 */ });
  }
  // P2-4：复制 Key（先等编辑态 Key 回填完成，避免复制到空；clipboard API 失败时降级 execCommand）
  document.getElementById("llm-key-copy").addEventListener("click", async () => {
    if (_llmFormState && _llmFormState.keyReady) { try { await _llmFormState.keyReady; } catch (e) {} }
    const v = keyInput.value.trim();
    if (!v) { toast("Key 为空，无可复制", "err"); return; }
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(v);
      } else {
        const ta = document.createElement("textarea");
        ta.value = v;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        ta.remove();
      }
          } catch (e) { toast("复制失败：" + e.message, "err"); }
  });
  // P0-1：服务商预设 chips
  document.querySelectorAll(".preset-chip").forEach(chip => {
    chip.addEventListener("click", () => {
      document.querySelectorAll(".preset-chip").forEach(c => c.classList.remove("active"));
      chip.classList.add("active");
      applyLlmPreset(chip.dataset.preset);
      _llmFormState.preset = chip.dataset.preset;
    });
  });
  // 编辑态自动匹配预设（baseUrl 完全等于某预设 → 高亮该 chip 并同步 placeholder）
  if (svc && svc.baseUrl) {
    const matched = Object.entries(LLM_PRESETS).find(([k, p]) => k !== "custom" && p.baseUrl === svc.baseUrl);
    if (matched) {
      const [k] = matched;
      _llmFormState.preset = k;
      const chip = document.querySelector(`.preset-chip[data-preset="${k}"]`);
      if (chip) chip.classList.add("active");
      const modelEl = document.getElementById("llm-model");
      if (modelEl) modelEl.placeholder = LLM_PRESETS[k].modelPlaceholder;
    }
  }
  // P1-3：URL / 名称失焦即时校验
  document.getElementById("llm-url").addEventListener("blur", () => {
    const el = document.getElementById("llm-url");
    const v = el.value.trim();
    el.classList.remove("input-error");
    if (v) { const err = llmValidateUrl(v); if (err) { el.classList.add("input-error"); toast(err, "err"); } }
  });
  document.getElementById("llm-name").addEventListener("blur", () => {
    document.getElementById("llm-name").value = document.getElementById("llm-name").value.trim();
  });
  // P1-2：dirty 检测 —— 任一输入变化置脏
  document.querySelectorAll("#llm-form input").forEach(inp => {
    inp.addEventListener("input", () => { _llmFormState.dirty = true; });
  });
  // 表单内测试：用「当前未保存的表单值」测（v1.10：API Key 必填，不再有免鉴权测试）
  document.getElementById("llm-form-test").addEventListener("click", async function() {
    const btn = this;
    const resEl = document.getElementById("llm-form-test-result");
    // 先等编辑态 Key 回填完成（避免仍按空 Key 测 → 401）
    if (_llmFormState && _llmFormState.keyReady) { try { await _llmFormState.keyReady; } catch (e) {} }
    const cfg = llmCollectForm();
    if (!cfg.name) { toast("请先填写服务名称", "err"); return; }
    if (!cfg.apiKey) { toast("请填写 API Key（必填）", "err"); return; }
    const urlErr = llmValidateUrl(cfg.baseUrl);
    if (urlErr) { toast(urlErr, "err"); return; }
    btn.disabled = true;
    btn.textContent = "测试中…";
    try {
      const r = await api("/api/llms/test", { method: "POST", body: JSON.stringify(cfg) });
      showTestResult(resEl, r);
    } catch (e) {
      showTestResult(resEl, { ok: false, message: "测试失败：" + e.message });
    } finally {
      btn.disabled = false;
      btn.textContent = "测试连接";
    }
  });

  // P2-1：form 回车 → 保存
  const llmForm = document.getElementById("llm-form");
  const onLlmSubmit = e => { e.preventDefault(); doLlmSave(); };
  llmForm.addEventListener("submit", onLlmSubmit);

  const okBtn = document.getElementById("modal-ok");
  // ⚠ 接管 modal-ok 一律用 onclick（覆盖式）而非 addEventListener：该节点是常驻的（见 openModal 复位段），
  //   用 addEventListener 会随每次打开弹窗累积，且用户「取消」关窗时不会解绑（只在本函数里解绑，
  //   而取消路径不走这里）⇒ 下次点「保存」会并发跑多个 handler：重复提交、重复 toast。
  okBtn.onclick = doLlmSave;
  // 保存逻辑（ok 按钮与 form 回车共用）
  async function doLlmSave() {
    if (!_llmFormState) return;
    const st = _llmFormState;
    const cfg = llmCollectForm();
    if (!cfg.name) { toast("请填写服务名称", "err"); return; }
    if (!cfg.apiKey) { toast("请填写 API Key（必填）", "err"); return; }
    if (!cfg.models.length) { toast("请填写模型名称", "err"); return; }
    const urlErr = llmValidateUrl(cfg.baseUrl);
    if (urlErr) { toast(urlErr, "err"); return; }
    try {
      if (st.isNew) {
        await api("/api/llms", { method: "POST", body: JSON.stringify(cfg) });
      } else {
        await api("/api/llms/" + encodeURIComponent(st.id), { method: "PUT", body: JSON.stringify(cfg) });
      }
      _llmFormState = null;
      _styleModalState = null;
      okBtn.onclick = null;
      llmForm.removeEventListener("submit", onLlmSubmit);
      closeModal(true);
            loadLlmList();
    } catch (e) {
      toast("保存失败：" + e.message, "err");
    }
  }
}

// ═══════════════════════════════════════════
// LLM 演示：智能体式流式输出（消费 POST /api/llm/stream 的 SSE）
// 事件 → 分区渲染：reasoning → 思考过程 / tool_call → 工具调用 / content → 输出
// done → 聚合统计（tokens / finish_reason）；error → 状态条报错；Abort → 上游取消
// ═══════════════════════════════════════════
let llmDemoAbort = null;        // 当前流的 AbortController
let llmDemoTimer = null;        // 状态条计时器
let llmDemoStart = 0;           // 本次开始时间（performance.now）
let llmDemoStats = null;        // { thinkChars, toolCalls, answerChars, phase }
let llmDemoToolEls = null;      // tool_call index → 步骤 DOM（流式分片就地更新）
let llmDemoLastAnswer = "";     // 最近一次回答正文（复制用）

function llmDemoEl(id) { return document.getElementById(id); }

/** 当前启用的服务 + 模型 徽标文案 */
function llmDemoModelText() {
  const svc = llmServices.find(s => s.id === llmEnabledId);
  if (!svc) return "";
  const m = svc.models && svc.models.length ? (svc.models.find(x => x.isDefault) || svc.models[0]) : null;
  return (svc.name || "") + (m ? " · " + m.name : "");
}

/** 依据 LLM 启用态刷新面板可用性：未启用 → 提示 + 禁用发送 */
function llmDemoRefreshGate() {
  const ok = llmConfigured();
  const dis = llmDemoEl("llmd-disabled");
  if (dis) dis.style.display = ok ? "none" : "";
  const send = llmDemoEl("btn-llmd-send");
  if (send) send.disabled = !ok;
  const badge = llmDemoEl("llmd-model-badge");
  if (badge) {
    const t = llmDemoModelText();
    badge.textContent = t ? "当前：" + t : "";
  }
  if (!ok) llmDemoSetStatus("尚未启用模型服务：请到「设置」页启用一个模型", "warn");
}

function llmDemoSetStatus(html, cls) {
  const el = llmDemoEl("llmd-status");
  if (!el) return;
  el.innerHTML = html;
  el.className = "llmd-status" + (cls ? " " + cls : "");
}

/** 追加文本到滚动区：用户停在底部时跟随滚动，上翻查看时不打扰 */
function llmDemoAppend(el, text) {
  if (!el) return;
  const pinned = el.scrollTop + el.clientHeight >= el.scrollHeight - 4;
  el.textContent += text;
  if (pinned) el.scrollTop = el.scrollHeight;
}

/** 逐块解析 SSE（块间空行分帧，event:/data: 两行）；onEvent(event, data)。
 *  供 LLM 演示与流式导入共用。 */
async function sseReadStream(resp, onEvent) {
  const reader = resp.body.getReader();
  const decoder = new TextDecoder();
  let buf = "";
  const handleBlock = block => {
    const ev = /^event: (.+)$/m.exec(block);
    const da = /^data: (.+)$/m.exec(block);
    if (!ev || !da) return;
    let data;
    try { data = JSON.parse(da[1]); } catch (e) { return; }   // 只忽略坏 JSON 帧
    onEvent(ev[1], data);   // 回调异常必须向上抛——error 事件的 throw（如导入失败真实原因）不能被吞成"坏帧"
  };
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    let idx;
    while ((idx = buf.indexOf("\n\n")) >= 0) {
      handleBlock(buf.slice(0, idx));
      buf = buf.slice(idx + 2);
    }
  }
  if (buf.trim()) handleBlock(buf);   // 结尾无空行的最后一个分片
}

function llmDemoTick() {
  const s = ((performance.now() - llmDemoStart) / 1000).toFixed(1);
  const st = llmDemoEl("llmd-status");
  if (!st || st.classList.contains("ok") || st.classList.contains("err") || st.classList.contains("warn")) return;
  const what = llmDemoStats && llmDemoStats.phase === "answer" ? "生成中" : "思考中";
  llmDemoSetStatus('<span class="kl-spinner"></span> ' + what + "… " + s + "s");
}

function llmDemoOnDone(data, seconds) {
  const u = data && data.usage;
  const parts = [];
  if (u) parts.push("prompt " + (u.prompt_tokens || 0) + " · completion " + (u.completion_tokens || 0) + " · total " + (u.total_tokens || 0) + " tokens");
  if (data && data.finishReason) parts.push("finish=" + data.finishReason);
  parts.push(seconds + "s");
  let html = "✓ 完成（" + parts.join(" · ") + "）";
  if (data && data.finishReason === "length") {
    html += ' <span class="err">· ⚠ 输出被截断（max_tokens 可能过小）</span>';
  }
  if (data && (!data.content || !data.content.trim())) {
    html += ' <span class="warn">· ⚠ 未产出正文（可能只有思考内容）</span>';
  }
  llmDemoSetStatus(html, "ok");
  const copyBtn = llmDemoEl("btn-llmd-copy");
  if (copyBtn) {
    if (data && data.content) { copyBtn.style.display = ""; llmDemoLastAnswer = data.content; }
    else copyBtn.style.display = "none";
  }
}

/** 工具调用增量：按 index 就地更新步骤卡片（arguments 为分片拼接后的原始 JSON） */
function llmDemoToolStep(data) {
  const tools = llmDemoEl("llmd-tools");
  if (!tools) return;
  tools.style.display = "";
  let step = llmDemoToolEls[data.index];
  if (!step) {
    step = document.createElement("div");
    step.className = "llmd-tool-step";
    step.innerHTML = '<div class="llmd-tool-name"></div><div class="llmd-tool-args"></div>';
    llmDemoEl("llmd-tools-body").appendChild(step);
    llmDemoToolEls[data.index] = step;
    llmDemoStats.toolCalls++;
    llmDemoEl("llmd-tools-badge").textContent = llmDemoStats.toolCalls + " 次";
  }
  step.querySelector(".llmd-tool-name").textContent = data.name || ("工具 " + (data.index + 1));
  const argsEl = step.querySelector(".llmd-tool-args");
  const raw = data.arguments || "";
  try { argsEl.textContent = raw ? JSON.stringify(JSON.parse(raw), null, 2) : ""; }
  catch (e) { argsEl.textContent = raw + (raw ? "" : "（等待参数…）"); }
}

async function llmDemoSend() {
  const promptEl = llmDemoEl("llmd-prompt");
  const prompt = (promptEl.value || "").trim();
  if (!prompt) { toast("请输入提示词", "warn"); return; }
  if (!llmConfigured()) { toast("尚未启用模型服务：请到「设置」页启用一个模型", "err"); return; }

  // ── 重置面板 ──
  llmDemoAbort = new AbortController();
  llmDemoToolEls = {};
  llmDemoStats = { thinkChars: 0, toolCalls: 0, answerChars: 0, phase: "think" };
  llmDemoLastAnswer = "";
  ["llmd-think-body", "llmd-tools-body", "llmd-answer-body"].forEach(id => { llmDemoEl(id).textContent = ""; });
  llmDemoEl("llmd-tools").style.display = "none";
  llmDemoEl("llmd-think-badge").textContent = "0 字";
  llmDemoEl("llmd-tools-badge").textContent = "0 次";
  llmDemoEl("llmd-answer-badge").textContent = "0 字";
  llmDemoEl("btn-llmd-copy").style.display = "none";
  const answerBody = llmDemoEl("llmd-answer-body");
  answerBody.classList.add("streaming");
  llmDemoEl("btn-llmd-send").disabled = true;
  llmDemoEl("btn-llmd-stop").style.display = "";

  llmDemoStart = performance.now();
  llmDemoSetStatus('<span class="kl-spinner"></span> 思考中…');
  llmDemoTimer = setInterval(llmDemoTick, 250);
  const started = performance.now();
  const elapsed = () => ((performance.now() - started) / 1000).toFixed(1);
  let finished = false;

  try {
    const resp = await fetch("/api/llm/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ messages: [{ role: "user", content: prompt }] }),
      signal: llmDemoAbort.signal,
    });
    if (!resp.ok) {
      let msg = "请求失败(" + resp.status + ")";
      try { const j = await resp.json(); if (j && j.error) msg = j.error; } catch (e) { /* 非 JSON */ }
      throw new Error(msg);
    }
    if (!resp.body) throw new Error("当前浏览器不支持流式读取（无 ReadableStream）");
    await sseReadStream(resp, (ev, data) => {
      if (ev === "reasoning") {
        llmDemoStats.thinkChars += (data.text || "").length;
        llmDemoEl("llmd-think-badge").textContent = llmDemoStats.thinkChars + " 字";
        llmDemoAppend(llmDemoEl("llmd-think-body"), data.text || "");
      } else if (ev === "tool_call") {
        llmDemoToolStep(data);
      } else if (ev === "content") {
        if (llmDemoStats.phase !== "answer") {
          llmDemoStats.phase = "answer";
          llmDemoSetStatus('<span class="kl-spinner"></span> 生成中…');
        }
        llmDemoStats.answerChars += (data.text || "").length;
        llmDemoEl("llmd-answer-badge").textContent = llmDemoStats.answerChars + " 字";
        llmDemoAppend(answerBody, data.text || "");
      } else if (ev === "done") {
        finished = true;
        llmDemoOnDone(data, elapsed());
      } else if (ev === "error") {
        throw new Error((data && (data.message || data.reason)) || "流式请求出错");
      }
    });
    if (!finished) llmDemoOnDone(null, elapsed());   // 流结束但未收到 done（异常兜底）
  } catch (e) {
    if (e && e.name === "AbortError") {
      llmDemoSetStatus("⏹ 已停止（客户端中断，上游生成已取消）", "warn");
    } else {
      llmDemoSetStatus("✗ " + esc(e.message || String(e)), "err");
    }
  } finally {
    clearInterval(llmDemoTimer); llmDemoTimer = null;
    answerBody.classList.remove("streaming");
    llmDemoEl("btn-llmd-send").disabled = false;
    llmDemoEl("btn-llmd-stop").style.display = "none";
    llmDemoAbort = null;
  }
}

function llmDemoStop() { if (llmDemoAbort) llmDemoAbort.abort(); }

function llmDemoCopy() {
  if (!llmDemoLastAnswer) return;
  const done = () => toast("已复制回答");
  const fallback = () => {
    const ta = document.createElement("textarea");
    ta.value = llmDemoLastAnswer;
    ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); done(); } catch (e) { toast("复制失败", "err"); }
    ta.remove();
  };
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(llmDemoLastAnswer).then(done).catch(fallback);
  } else fallback();
}

// ═══════════════════════════════════════════
// 样式管理（样式定义的新增/编辑，统一经 editStyleModal）
// 全局样式库已移除：样式全部随模板集走，仅本模板集定制样式可编辑。
// ═════════════════════════════════════════════

async function editStyleModal(style, scope) {
  const isNew = !style;
  const s = style || { key: "", name: "", matchType: "none", match: "", mode: "whole", defaults: {} };
  // 结构槽位（special）候选 = 引擎固定样式的结构角色。原「识别方式：正则/前缀/特殊」下拉已删除
  //（2026-09-10 用户拍板）：段落归类一律交模型，正则/前缀匹配不再参与识别；结构槽位只对既有 special 样式可改，
  // 其余样式（含旧模板残留的 regex/prefix/none）只读展示、保存时不覆写（避免把别的样式顶到 firstLine/default 槽位）。
  const SPECIAL_LABELS = { firstLine: "首行标题", subtitle: "副标题", salutation: "称谓", default: "默认正文", sign: "落款署名" };
  const specialOpts = Object.keys(SPECIAL_LABELS)
    .map(v => `<option value="${v}" ${s.match === v ? "selected" : ""}>${SPECIAL_LABELS[v]}</option>`).join("");
  const slotRow = s.matchType === "special"
    ? `<div class="prop-row"><label>结构判定</label><select id="sm-special">${specialOpts}</select></div>`
    : `<div class="prop-row"><label>结构判定</label><span class="sm-static">无（该样式由模型判出，不占结构槽位）</span></div>`;
  const modalBody = `
    <div class="sm-form">
      <div class="prop-row"><label>名称</label><input type="text" id="sm-name" value="${esc(s.name || (isNew ? "" : s.key))}" placeholder="如 三级标题"></div>
      <div class="prop-row"><label>标识 key</label><input type="text" id="sm-key" value="${esc(s.key)}" placeholder="如 h3" ${isNew ? "" : "disabled"}></div>
      ${slotRow}
      <div class="prop-row"><label>渲染方式</label>
        <select id="sm-mode">
          <option value="whole" ${s.mode !== "lead" ? "selected" : ""}>整段统一</option>
          <option value="lead" ${s.mode === "lead" ? "selected" : ""}>首句标题+句号后正文</option>
        </select>
      </div>
      <div class="prop-row"><label>默认字体</label>${fontSelect("sm-font", (s.defaults && s.defaults.font) || "")}</div>
    </div>`;
  const title = isNew
    ? "新建样式（仅本模板集生效）"
    : "编辑样式";
  openModal({ title, body: modalBody, okText: "保存" });
  // 注意：不能 await openModal —— 其 Promise 直到弹窗关闭才 resolve，
  // await 会让下面的接管与保存绑定失效（点「保存」只关窗不提交）。
  _styleModalState = { s, isNew, scope, editExisting: !!style };
  const okBtn = document.getElementById("modal-ok");
  // ⚠ 接管 modal-ok 用 onclick（覆盖式）：该节点常驻，用 addEventListener 会随每次打开累积，
  //   且「取消」关窗不解绑 ⇒ 下次保存并发跑多个 handler（重复保存、重复 toast）。
  okBtn.onclick = async function _saveStyleHandler() {
    const st = _styleModalState;
    if (!st) return;
    const name = document.getElementById("sm-name").value.trim();
    let key = document.getElementById("sm-key").value.trim();
    const specialEl = document.getElementById("sm-special");       // 仅 special 类样式有该下拉
    const specialVal = specialEl ? specialEl.value : "";
    const mode = document.getElementById("sm-mode").value;
    const font = document.getElementById("sm-font").value;
    const scopeOut = st.scope;            // 新建时由按钮决定，编辑时沿用原桶
    if (!name) { toast("请填写样式名称", "err"); return; }
    // v1.30：样式定义随当前模板自包含
    const tpl = currentTemplate();
    const targetArr = (tpl && Array.isArray(tpl.styles)) ? tpl.styles : [];
    if (st.isNew) {
      if (!/^[a-z][a-z0-9_]*$/.test(key)) { toast("标识 key 需为小写字母开头（如 h3）", "err"); return; }
      if (allStyleDefs().some(x => x.key === key)) { toast("标识 key 已存在", "err"); return; }
      // 新建样式不占结构槽位（槽位归引擎固定样式）；匹配方式字段不再由面板产生
      targetArr.push({ key, name, matchType: "none", match: "", mode, defaults: { font: font || undefined } });
    } else {
      const obj = st.s;
      obj.name = name; obj.mode = mode;
      if (specialEl && obj.matchType === "special") obj.match = specialVal;   // 只有 special 类可改槽位
      obj.defaults = Object.assign({}, obj.defaults, { font: font || undefined });
    }
    await flushSave();
    if (tpl && !tpl.include.includes(key)) { tpl.include.push(key); markDirty(); }
    _styleModalState = null;
    okBtn.onclick = null;
    closeModal(true);
    renderAll();
  };
}

let _styleModalState = null;

// ═══════════════════════════════════════════
// 渲染入口
// ═══════════════════════════════════════════
function renderAll() {
  loadParaOverrides();   // 切换模板集/模板后同步当前模板的预览覆盖（localStorage，不落模板）
  renderTemplateTitle();
  renderStyleList();
  renderPreview();
  // 段落选中态由 selectElem / clearParaSelection 独占管理（不再与 selectedElem 联动）；
  // 此处仅作安全网：渲染完成后若 selectedParaKey 已清但 DOM 残留 .selected 则清理
  if (!state.selectedParaKey) {
    document.querySelectorAll(".para.selected").forEach(el => el.classList.remove("selected"));
  }
  // 模板级样本：切换模板/模板集后异步重载（loadSample 按 currentFile|currentTemplateKey 去重，未切换直接跳过）
  loadSample();
}

// 自动保存：任何改动实时落盘（文本输入防抖），以极简状态字替代保存/取消按钮
let _saveTimer = null;
function markDirty() {
  state.dirty = true;
  scheduleSave();
}
function scheduleSave() {
  setSaveStatus("保存中…");
  if (_saveTimer) clearTimeout(_saveTimer);
  _saveTimer = setTimeout(() => { persistCurrentSet(); }, 500);
}
function setSaveStatus(text) {
  const el = document.getElementById("tpl-save-status");
  if (el) el.textContent = text || "";
}
// 切换 / 新建前若有待保存改动，立即落盘，避免防抖窗口内丢失
async function flushSave() {
  if (_saveTimer) { clearTimeout(_saveTimer); _saveTimer = null; }
  if (state.dirty) await persistCurrentSet();
}

// ── 左侧菜单折叠（Font Awesome 双箭头，随折叠方向切换）──
const SIDEBAR_ICONS = {
  open:  '<i class="fas fa-angles-left"></i>',
  closed:'<i class="fas fa-angles-right"></i>'
};
function renderSidebarCollapse() {
  const sb = document.getElementById("sidebar");
  const btn = document.getElementById("sidebar-collapse");
  if (!sb || !btn) return;
  const collapsed = sb.classList.contains("collapsed");
  btn.innerHTML = collapsed ? SIDEBAR_ICONS.closed : SIDEBAR_ICONS.open;
  const baseTitle = collapsed ? "展开菜单" : "折叠菜单";
  // 诊断折叠态图标底色压扁问题：把按钮/nav-icon 实际渲染尺寸追加到 title，hover 看 tooltip 即可确认
  const br = btn.getBoundingClientRect();
  btn.title = `${baseTitle}（${Math.round(br.width)}×${Math.round(br.height)}）`;
  btn.setAttribute("aria-label", baseTitle);
  sb.querySelectorAll(".nav-item").forEach(item => {
    const icon = item.querySelector(".nav-icon");
    if (icon) {
      const ir = icon.getBoundingClientRect();
      icon.title = collapsed
        ? `${item.dataset.label || ""}（${Math.round(ir.width)}×${Math.round(ir.height)}）`
        : "";
    }
  });
}
function bindSidebarCollapse() {
  const sb = document.getElementById("sidebar");
  const btn = document.getElementById("sidebar-collapse");
  if (!sb || !btn) return;
  // 把菜单文字同步到 data-label / nav-icon title，供折叠态悬停提示使用
  sb.querySelectorAll(".nav-item").forEach(item => {
    const label = item.querySelector(".nav-label");
    if (label && !item.dataset.label) item.dataset.label = label.textContent.trim();
  });
  // 手动折叠时通知自动折叠逻辑（记录用户意图，避免画布变宽后被自动展开顶掉）
  function notifyManual() {
    sb.dispatchEvent(new CustomEvent("sidebar:manual", { detail: sb.classList.contains("collapsed") }));
  }
  btn.addEventListener("click", e => {
    e.stopPropagation();
    sb.classList.toggle("collapsed");
    renderSidebarCollapse();
    notifyManual();
  });
  // 折叠按钮已从 <button> 改为 <div>，补齐键盘可达性（Enter / Space 触发）
  btn.addEventListener("keydown", e => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      sb.classList.toggle("collapsed");
      renderSidebarCollapse();
      notifyManual();
    }
  });
  renderSidebarCollapse();
}

// 画布过窄自动折叠左侧菜单：画布 .preview-pane 已设最小宽度 842px（见样式），窗口再窄也不再被压小，
// 而是由整个模板设计视图横向溢出（.layout overflow-x:auto）。这里监听 .layout 是否横向溢出——
// 溢出即画布被压到最小、放不下，先自动折叠侧栏腾出约 188px；仍放不下才靠横向滚动条。
// 自动展开用滞回（展开会让可用空间 -188px），避免"折叠→变宽→展开→又折叠"振荡；用户手动折叠后不自动展开。
let sidebarAutoApply = () => {};
function bindSidebarAutoCollapse() {
  const sb = document.getElementById("sidebar");
  const vd = document.getElementById("view-design");
  const layout = vd ? vd.querySelector(".layout") : null;
  if (!sb || !layout) return;
  const FREED = 244 - 56;   // 折叠释放给画布的宽度
  const MARGIN = 20;        // 展开滞回缓冲（防止临界抖动）
  let userCollapsed = null; // null=自动；true=用户手动折叠；false=用户手动展开
  let forceOpen = false;    // 用户手动展开过 → 自动折叠让位（2026-08-26 方案 A：修「窗口过窄点展开没反应」）
  sb.addEventListener("sidebar:manual", e => {
    userCollapsed = e.detail;
    // 用户手动展开（false）时挂起自动折叠，尊重手动意图；手动折叠（true）时解除挂起
    forceOpen = (e.detail === false);
  });
  function apply() {
    if (vd.classList.contains("hidden")) return;   // 未进入模板设计视图：不干预
    const overflowing = layout.scrollWidth > layout.clientWidth + 1;   // 设计视图横向溢出
    // 窗口已放得下 → 解除强制展开锁，自动折叠逻辑完全恢复
    if (!overflowing) forceOpen = false;
    const collapsed = sb.classList.contains("collapsed");
    if (!collapsed) {
      // 已展开：画布被压到最小、放不下 → 自动折叠腾空间（用户手动展开过则让位，不再抢）
      if (overflowing && !forceOpen) { sb.classList.add("collapsed"); renderSidebarCollapse(); }
    } else {
      // 已折叠：仅当用户未手动折叠、且展开后也放得下（滞回）才自动展开
      if (userCollapsed !== true && !forceOpen) {
        const overAfterExpand = layout.scrollWidth > (layout.clientWidth - FREED - MARGIN);
        if (!overAfterExpand) { sb.classList.remove("collapsed"); renderSidebarCollapse(); }
      }
    }
  }
  let ro;
  if (typeof ResizeObserver !== "undefined") { ro = new ResizeObserver(apply); ro.observe(layout); }
  else { window.addEventListener("resize", apply); }
  sidebarAutoApply = apply;
  apply();
}

// ═══════════════════════════════════════════
// 初始化
// ═══════════════════════════════════════════
async function init() {
  clearParaSelection();   // 启动时强制清除段落选中态（防止任何残留）
  bindModalEvents();
  // Toast 关闭只走右上角 × 按钮（见上方 toast() 注册的 close 监听）：
  // 不再「点哪都关」，正文区留给用户选中复制，避免选中文字一抬鼠标就消失。
  await initTasks(); // 从服务端加载历史任务（落盘 %APPDATA%/doc-audit/Tasks/tasks.json）
  renderTasks(); // 填充左侧转换任务列表
  // 转换第 3 步：格式处理选项勾选卡 · 装载本地持久化偏好 + 监听 change 实时落盘
  setupInheritGrid();
  // 设置页「人员名单」卡：事件委托一次性绑定（内容由 loadRosterSettings 整体重渲染）
  setupRosterSettings();
  // 任务列表事件委托：单击打开任务，双击就地重命名，删除按钮 hover 显示
  document.getElementById("tasks-tree").addEventListener("click", e => {
    const more = e.target.closest(".tree-more");
    if (more) {
      e.stopPropagation();
      tasksListExpanded = !tasksListExpanded;
      renderTasks();
      return;
    }
    const del = e.target.closest(".task-del");
    if (del) {
      e.stopPropagation();
      deleteTask(del.dataset.task);
      return;
    }
    const item = e.target.closest(".task-item");
    // 双击会重绘 DOM，旧元素可能已不在文档中，忽略避免重复 openTask
    if (!item || item.classList.contains("editing") || !document.contains(item)) return;
    const t = loadTasks().find(x => x.id === item.dataset.task);
    if (t) openTask(t);
  });
  document.getElementById("tasks-tree").addEventListener("dblclick", e => {
    const item = e.target.closest(".task-item");
    if (!item || item.classList.contains("editing")) return;
    const taskId = item.dataset.task;
    // 双击会先触发两次 click（已重绘 DOM），用 taskId 重新定位当前行
    const liveItem = document.querySelector(`.task-item[data-task="${taskId}"]`);
    if (liveItem) beginInlineRenameTask(liveItem, taskId);
  });
  // 展开/收起行是 role=button（可 Tab 到达），需自带键盘触发
  document.getElementById("tasks-tree").addEventListener("keydown", e => {
    if (e.key !== "Enter" && e.key !== " ") return;
    const more = e.target.closest(".tree-more");
    if (!more) return;
    e.preventDefault();
    tasksListExpanded = !tasksListExpanded;
    renderTasks();
  });
  // 启动只取一次 /api/status：模板集树依赖这份响应
  let st;
  try { st = await refreshStatus(); }
  catch (e) { setStatus("✗ 无法连接服务", "err"); return; }
  // 全局样式库已移除：无「内置默认」概念（样式随模板集走）
  renderSetTree(st);
  // 根图标
  document.getElementById("ic-design").innerHTML = '<i class="fas fa-table-cells"></i>';
  document.getElementById("ic-convert").innerHTML = '<i class="fas fa-arrows-rotate"></i>';
  document.getElementById("ic-settings").innerHTML = '<i class="fas fa-gear"></i>';

  // 左侧菜单折叠按钮（方案 C 图标）+ 画布过窄自动折叠
  bindSidebarCollapse();
  bindSidebarAutoCollapse();

  // 左侧导航：切换视图
  document.querySelectorAll(".nav-item").forEach(n => {
    n.addEventListener("click", e => {
      if (e.target.closest(".nav-dots")) return;
      if (!n.dataset.view) return;
      if (n.dataset.view === "design") {
        // 模板集菜单项：仅展开/折叠子菜单（模板集列表），不切换右侧视图、不清空当前模板
        document.getElementById("sets-tree").classList.toggle("collapsed");
      } else if (n.dataset.view === "convert") {
        // 新建任务的入口 = 左侧菜单「文档排版」：已在转换视图时点它 = 新建任务（清空输入/显示区）；
        // 从其他视图点它 = 切回并恢复最新任务显示（并行：不中断任何流）
        if (currentView === "convert") {
          newConvertTask();
        } else {
          switchView("convert");
        }
      } else {
        switchView(n.dataset.view);
      }
    });
  });
  // 模板设计根 ＋：直接新建模板集（无子菜单，tooltip 即「新建模板集」）
  document.getElementById("dots-design-root").addEventListener("click", e => {
    e.stopPropagation();
    newTemplateSet();
  });
  // 文档排版根 ＋：新建排版任务（与点标题等效；在其他视图点时先切回排版视图）
  document.getElementById("dots-convert-root").addEventListener("click", e => {
    e.stopPropagation();
    if (currentView !== "convert") switchView("convert");
    newConvertTask();
  });
  // 转换视图：折叠摘要行点击展开
  document.getElementById("input-summary")?.addEventListener("click", () => {
    document.getElementById("block-input")?.classList.remove("collapsed");
  });
  document.getElementById("template-summary")?.addEventListener("click", () => {
    document.getElementById("block-template")?.classList.remove("collapsed");
  });
  document.getElementById("inherit-summary")?.addEventListener("click", () => {
    document.getElementById("block-inherit")?.classList.remove("collapsed");
  });
  // 转换视图内登录/转换按钮
  // 样式面板搜索框（实时过滤，不重建列表）
  const styleSearch = document.getElementById("style-search");
  if (styleSearch) {
    styleSearch.addEventListener("input", () => { styleFilterText = styleSearch.value; applyStyleFilter(); });
    styleSearch.addEventListener("keydown", e => {
      if (e.key === "Escape") { styleSearch.value = ""; styleFilterText = ""; applyStyleFilter(); }
    });
  }
  document.getElementById("btn-convert").addEventListener("click", doConvert);
  // 转换过程卡右上角「重新转换」：重跑聚焦任务/历史查看的任务（进行中先确认再中断本任务流）
  const reconvertBtn = document.getElementById("btn-reconvert");
  if (reconvertBtn) reconvertBtn.addEventListener("click", reconvertTask);
  // 转换过程卡右上角「停止」：中断聚焦任务（D4 拍板 2026-09-05）
  const stopConvertBtn = document.getElementById("btn-stop-convert");
  if (stopConvertBtn) stopConvertBtn.addEventListener("click", stopFocusedTask);
  // 设置页：添加 LLM 服务
  document.getElementById("btn-add-llm").addEventListener("click", () => editLlmModal(null));
  // 设置页：默认样式为出厂基准只读展示，编辑统一在「模板设计」页（本页无任何写回入口）
  // 设置页页脚：使用引导（重看，不改 done 状态）· 关于本软件 · 支持格式 · 更新日志
  const lnkOnboarding = document.getElementById("lnk-onboarding");
  if (lnkOnboarding) lnkOnboarding.addEventListener("click", () => {
    if (typeof showOnboarding === "function") showOnboarding({ force: true });
  });
  const lnkAbout = document.getElementById("lnk-about");
  if (lnkAbout) lnkAbout.addEventListener("click", async () => {
    await ensureVersion();
    const ver = versionText(appVersion);
    infoModal("关于文档审核", aboutBody(),
      ver ? `关于文档审核<span class="modal-ver">${esc(ver)}</span>` : undefined);
  });
  const lnkFormats = document.getElementById("lnk-formats");
  if (lnkFormats) lnkFormats.addEventListener("click", () => infoModal("支持格式", formatsBody()));
  const lnkChangelog = document.getElementById("lnk-changelog");
  if (lnkChangelog) lnkChangelog.addEventListener("click", openChangelogModal);
  // 设置页：软件更新的入口已移到侧栏账号菜单（见 openAccountMenu），此处不再绑定按钮
  // LLM 演示页：导航图标 + 交互（菜单入口已移除，代码保留待重新启用；发布时清理整块）
  const icLlmDemo = document.getElementById("ic-llmdemo");
  if (icLlmDemo) icLlmDemo.innerHTML = ICONS.llmdemo;
  document.getElementById("btn-llmd-send").addEventListener("click", llmDemoSend);
  document.getElementById("btn-llmd-stop").addEventListener("click", llmDemoStop);
  document.getElementById("btn-llmd-copy").addEventListener("click", llmDemoCopy);
  document.querySelectorAll(".llmd-example").forEach(ch => {
    ch.addEventListener("click", () => {
      const ta = document.getElementById("llmd-prompt");
      ta.value = ch.dataset.p || "";
      ta.focus();
    });
  });
  // Enter 发送（Shift+Enter 换行；组合输入中不触发）
  document.getElementById("llmd-prompt").addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
      e.preventDefault();
      if (!document.getElementById("btn-llmd-send").disabled) llmDemoSend();
    }
  });
  document.getElementById("file-input").addEventListener("change", e => {
    if (e.target.files && e.target.files[0]) {
      const f = e.target.files[0];
      const entrySet = importEntrySet;   // 钉死的入口集（点 ⋯ 时确定），不随对话框期间切集变化
      importEntrySet = null;
      doImport(f, entrySet);             // 直接传入口集，彻底不依赖 state.currentFile
    }
    e.target.value = "";
  });

  // ESC 键退出段落选中
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && state.selectedParaKey) {
      e.preventDefault();
      clearParaSelection();
      renderStyleList();   // 同步样式面板高亮
    }
  });
  // 预览区点击：点段落=选中该段（事件委托，适配动态生成 DOM）；点空白=退出选中
  const paperEl = document.getElementById("paper");
  if (paperEl) {
    paperEl.addEventListener("click", e => {
      const para = e.target.closest(".para[data-key]");
      if (para) { selectElem(para.dataset.key, para); return; }
      if (state.selectedParaKey && !e.target.closest(".para")) {
        clearParaSelection();
        renderStyleList();
      }
    });
    // 段落右键：先选中该段（已在选中态则不取消），再弹「设为/取消标题引导段」菜单
    paperEl.addEventListener("contextmenu", e => {
      const para = e.target.closest(".para[data-key]");
      if (!para) return;
      e.preventDefault();
      if (!para.classList.contains("selected")) selectElem(para.dataset.key, para);
      openParaContextMenu(e.clientX, e.clientY, para);
    });
  }
  // 样品工具条：打开样本（恢复默认已移入样本编辑弹窗内）
  const openSampleBtn = document.getElementById("btn-open-sample");
  if (openSampleBtn) openSampleBtn.addEventListener("click", editSampleInPage);
  // 模板说明：编辑当前模板的 description（导入时填写的说明在此可改；清空即移除）
  const descBtn = document.getElementById("btn-tpl-desc");
  if (descBtn) descBtn.addEventListener("click", editTemplateDesc);

  // 在线文档链接输入已移除，仅保留本地上传（上传区事件见下方）

  // 方案A：上传区点击 / 拖拽 → 选本地文件（多选），上传落盘后列入清单
  const uploadZone = document.getElementById("upload-zone");
  const convertFileInput = document.getElementById("convert-file-input");
  const pickLocal = () => convertFileInput && convertFileInput.click();
  if (uploadZone) {
    uploadZone.addEventListener("click", pickLocal);
    uploadZone.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); pickLocal(); } });
  }
  const addLocalFiles = async (files) => {
    const arr = Array.from(files || []);
    if (!arr.length) return;
    const blockInput = document.getElementById("block-input");
    if (blockInput) blockInput.classList.remove("link-only", "local-only");
    let added = 0;
    for (const f of arr) {
      const rec = await uploadConvertFile(f);
      if (rec) { convertLocalFiles.push(rec); added++; }
    }
    renderUploadList();
    if (added) showUrlHint("已添加 " + added + " 个本地文件", "ok");
  };
  if (convertFileInput) {
    convertFileInput.addEventListener("change", async e => {
      // 先拷贝成数组再清 value：置空 value 会同时清空 FileList（活动引用），拖拽不受影响、点击选择会拿到空列表
      const files = Array.from(e.target.files || []);
      e.target.value = "";
      await addLocalFiles(files);
    });
  }
  // 拖拽到上传区：等价于点击选择（多选）。接线与来源框共用 bindFileDropZone。
  bindFileDropZone(uploadZone, files => addLocalFiles(files));

  window.addEventListener("beforeunload", e => {
    if (_saveTimer) { e.preventDefault(); e.returnValue = ""; }   // 防抖窗口内仍有待保存改动
  });

  if (st.templatesets.length) {
    uiState.expandedSets.add(st.templatesets[0].file);
    await loadTemplateSet(st.templatesets[0].file);
  } else {
    // 无模板集：不自动加载内置集、不自动新建，主窗口留空
    state.templateSet = null;
    state.currentFile = "";
    state.currentTemplateKey = null;
  }

  // 样本加载已并入 renderAll（模板集加载/切换后自动重载）；此处兜底：无模板集时清空预览。
  await loadSample();

  // 启动后：若上次导入结果被 F5 / WebView2 刷新中断，自动恢复预览页；否则默认进入「文档排版」
  if (!await restoreImportView()) switchView("convert");
  // 扫码授权续跑：授权页导航卸载本页后，登录成功切回（launcher 带 ?resume=1）自动继续上次转换
  await maybeResumePendingConvert();

  // 首次启动使用引导：init 末尾判定（done===false 才弹）；设置页可重看
  if (typeof maybeShowOnboarding === "function") await maybeShowOnboarding();

  // 自动检查更新：冷启动跑完后静默查一次。
  // 只有「真有新版本」才弹窗打扰；网络不通 / 清单异常一律静默（下次启动或手动点按钮再试）。
  setTimeout(() => { refreshUpdateState(); checkForUpdate(false); }, 4000);
}

// ── 结果列表：下载产物（服务端部署，产物在服务端任务目录，浏览器需下载到本地）──
// 事件委托挂在容器上，实时结果与历史任务快照（innerHTML 回显）同样生效。
// 历史快照里保存的 resultHtml 是早期版本（含 a.download-link、无打开按钮）→ 渲染前用 upgradeResultItems 升级为新结构。
function upgradeResultItems(rl) {
  if (!rl) return;
  rl.querySelectorAll(".result-item").forEach(item => {
    const a = item.querySelector("a.download-link");
    if (!a) return;                       // 已是新结构（含 .open-btn），跳过
    let rel = "";
    try {
      const u = new URL(a.getAttribute("href") || "", location.origin);
      rel = u.searchParams.get("f") || u.searchParams.get("file") || "";
    } catch {}
    a.remove();
    const open = document.createElement("button");
    open.type = "button";
    open.className = "open-btn";
    open.setAttribute("data-download", rel);
    if (!rel) open.disabled = true;
    open.textContent = "下载";
    item.appendChild(open);
  });
}
function bindResultActions() {
  const rl = document.getElementById("result-list");
  if (!rl) return;
  function downloadResult(rel) {
    if (!rel) return;
    // 服务端部署：走 /api/download 流式下载（Content-Disposition: attachment 触发浏览器保存）
    const a = document.createElement("a");
    a.href = "/api/download?f=" + encodeURIComponent(rel);
    a.download = "";
    document.body.appendChild(a);
    a.click();
    a.remove();
  }
  rl.addEventListener("click", e => {
    const btn = e.target.closest("[data-download]");
    if (btn) downloadResult(btn.getAttribute("data-download"));
  });
  // 首屏若已回显历史快照，立即升级一次
  upgradeResultItems(rl);
}

// 全局兜底：拖到**框外**时也不许触发"导航到 file://"。
// 两台页面的接收区都是框内（bindFileDropZone），框外这一层只负责拦住默认行为、不接管内容。
// 只拦带文件的拖动 —— 画布里拖元素是 HTML5 内部拖动（types 里没有 Files），碰不得。
function bindFileDropGuard() {
  const guard = e => {
    const t = e.dataTransfer;
    if (t && Array.prototype.indexOf.call(t.types || [], "Files") >= 0) e.preventDefault();
  };
  document.addEventListener("dragover", guard);
  document.addEventListener("drop", guard);
}

bindFileDropGuard();
bindResultActions();

init();
