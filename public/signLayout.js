/*
 * 落款（署名/成文日期）右缩进算法 —— 预览端(app.js)与转换端(convert.js)共用单一真相源。
 * UMD：浏览器挂 window.SignLayout，Node 侧 module.exports。
 *
 * 算法规则（GB/T 9704-2012 §7.3.5.2 不盖章公文）：
 *  - 成文日期右空 baseRightIndent 字（通常来自落款样式 rightIndentChars，兜底 4）。
 *  - 署名以日期为中心上下居中：署名右缩进 = base + (日期显示宽 − 署名显示宽)/2，允许小于 base；
 *    仅当结果 < 0 才截为 0，防止署名越过右边界。
 *  - 落款成组（连续 ≥2 个落款块）时：末块=日期固定右空 base；前面块=署名按居中公式上调。
 *  - 落款仅 1 个块时：右空 base。
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) module.exports = factory();
  else root.SignLayout = factory();
})(typeof self !== "undefined" ? self : this, function () {
  // 中文/全角符记 1 字宽，其它（含半角、数字、标点）记 0.5 字宽。
  function charWidth(s) {
    let w = 0;
    for (const ch of (s || "")) {
      w += /[一-鿿　-〿＀-￯]/.test(ch) ? 1 : 0.5;
    }
    return w;
  }

  // blocks: 段落块数组（预览端 {type:"para", style, text} / 转换端 {type:<样式key>, text}）
  // getStyleKey(b): 返回该块的样式 key（预览端取 b.style，转换端取 b.type）
  // signKey: 落款样式 key（如 "落款"）
  // baseRightIndent: 基准右空字符数（来自落款 rightIndentChars，兜底 4）
  // 返回：{ blockIndex -> 该段右缩进字符数 }
  function computeSignRightIndents(blocks, getStyleKey, signKey, baseRightIndent) {
    const map = {};
    let i = 0;
    while (i < blocks.length) {
      if (getStyleKey(blocks[i]) === signKey) {
        const run = [];
        while (i < blocks.length && getStyleKey(blocks[i]) === signKey) { run.push(i); i++; }
        if (run.length >= 2) {
          const dateChars = charWidth((blocks[run[run.length - 1]].text || "").trim());
          for (const ri of run.slice(0, -1)) {
            const signChars = charWidth((blocks[ri].text || "").trim());
            map[ri] = Math.max(0, Math.round(baseRightIndent + (dateChars - signChars) / 2));
          }
          map[run[run.length - 1]] = baseRightIndent;
        } else {
          map[run[0]] = baseRightIndent;
        }
      } else i++;
    }
    return map;
  }

  return { charWidth, computeSignRightIndents };
});
