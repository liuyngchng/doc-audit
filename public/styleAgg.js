/*
 * 样式格式聚合（众数）—— 导入 docx 提取（server/convert.js）与导入结果页段落修正（public/app.js）共用单一真相源。
 * UMD：浏览器挂 window.StyleAgg，Node 侧 module.exports。
 *
 * 用途：
 *  - mode(arr, defaultVal)：统计数组里出现最多的值（众数），空数组/全空值返回 defaultVal。
 *  - elemFromParas(paras, defDefaults)：由一组段落样本聚合样式格式值（各字段取众数）。
 *    paras 为空时用 defDefaults 兜底（与"未提取到样式 → 使用默认样式"语义一致）。
 *    数值字段转数字、布尔字段转布尔；color 仅当样本中出现过颜色时才输出。
 *
 * 注意：本模块与 convert.js 既有 elemFrom 行为完全一致（v1.33 抽取，勿在此引入新语义；
 * 若提取规则调整，改这里即可两端生效）。
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.StyleAgg = factory();
})(typeof self !== "undefined" ? self : this, function () {
  /** 统计数组里出现最多的值（众数）；clean 为空（全 null/undefined/""）返回 defaultVal */
  function mode(arr, defaultVal = null) {
    const clean = arr.filter(v => v !== null && v !== undefined && v !== "");
    if (!clean.length) return defaultVal;
    const count = {};
    for (const v of clean) count[v] = (count[v] || 0) + 1;
    let best = null, bestN = -1;
    for (const k of Object.keys(count)) if (count[k] > bestN) { bestN = count[k]; best = k; }
    return best;
  }

  /** 由段落样本聚合样式格式值（众数）；ps 为空时用 def 兜底（语义同"无样本用默认"） */
  function elemFromParas(ps, def) {
    const arr = ps && ps.length ? ps : [def || {}];
    // 数值字段转数字，布尔字段转布尔
    const num = v => (v === null || v === undefined || v === "" ? null : parseFloat(v));
    const bool = v => (v === true || v === "true" || v === "1" || v === "on" ? true : (v === false || v === "false" || v === "0" ? false : null));
    const elem = {
      font: mode(arr.map(p => p.font), def && def.font),
      size: num(mode(arr.map(p => p.size), def && def.size)),
      align: mode(arr.map(p => p.align), def && def.align),
      // 行距：auto 倍数（lineMult）与固定磅值（line）互斥——有倍数样本时 line 置 null（防回退固定值）
      line: arr.some(p => p.lineMult != null)
        ? null
        : num(mode(arr.map(p => p.line), def && def.line)),
      ...(arr.some(p => p.lineMult != null) ? { lineMult: num(mode(arr.map(p => p.lineMult))) } : {}),
      bold: bool(mode(arr.map(p => p.bold), def ? !!def.bold : false)),
      indentChars: num(mode(arr.map(p => p.indentChars), def && def.indentChars)),
      rightIndentChars: num(mode(arr.map(p => p.rightIndentChars), def && def.rightIndentChars)),
      spaceBefore: num(mode(arr.map(p => p.spaceBefore), 0)),
      spaceAfter: num(mode(arr.map(p => p.spaceAfter), 0)),
      // bodyFont/bodyBold 已废弃（方案 B 收口 2026-09-07）：lead 正文恒跟随「正文段落」，不再聚合输出
      ...(arr.some(p => p.color) ? { color: mode(arr.map(p => p.color)) } : {}),
      ...(arr.some(p => p.charScale != null) ? { charScale: num(mode(arr.map(p => p.charScale))) } : {}),
      ...(arr.some(p => p.charSpacing != null) ? { charSpacing: num(mode(arr.map(p => p.charSpacing))) } : {}),
    };
    return elem;
  }

  return { mode, elemFromParas };
});
