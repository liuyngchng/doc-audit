#!/usr/bin/env node
'use strict'

// 从 cfg.yml（OpenAI 兼容 LLM 配置）生成原版服务端读取的 llms.json。
// 用法：node apply-cfg.js /path/to/cfg.yml /data/config/llms.json
//
// cfg.yml 结构（标准 YAML，支持 # 注释、单/双引号）：
//   llm:
//     name: 我的模型服务
//     baseUrl: https://api.example.com/v1
//     apiKey: sk-xxx
//     models:
//       - name: gpt-4o
//         default: true
//         maxTokens: 8192
//     timeoutMs: 60000
//     note: 自建网关

const fs = require('fs')
const path = require('path')
const crypto = require('crypto')

function fail(msg) {
  console.error(`[apply-cfg] ${msg}`)
  process.exit(1)
}

let YAML
try {
  YAML = require('yaml')
} catch (e) {
  fail('缺少 yaml 依赖，请先 npm install yaml（容器镜像内已预装）')
}

function str(v) {
  return String(v ?? '').trim()
}

function main() {
  const cfgPath = process.argv[2]
  const outPath = process.argv[3]
  if (!cfgPath || !outPath) fail('用法：node apply-cfg.js <cfg.yml> <llms.json 输出路径>')

  if (!fs.existsSync(cfgPath)) fail(`配置文件不存在：${cfgPath}`)

  let parsed
  try {
    parsed = YAML.parse(fs.readFileSync(cfgPath, 'utf8')) || {}
  } catch (e) {
    fail(`cfg.yml 解析失败：${e.message}`)
  }

  const llm = (parsed.llm && typeof parsed.llm === 'object') ? parsed.llm : parsed

  const name = str(llm.name) || 'OpenAI 兼容服务'
  const baseUrl = str(llm.baseUrl).replace(/\/+$/, '')
  const apiKey = str(llm.apiKey)

  if (!/^https?:\/\//i.test(baseUrl)) fail('llm.baseUrl 需以 http:// 或 https:// 开头')
  if (!apiKey) fail('llm.apiKey 不能为空')

  const rawModels = Array.isArray(llm.models) ? llm.models : []
  const seen = new Set()
  const models = []
  for (const m of rawModels) {
    const modelName = str(m && m.name)
    if (!modelName || seen.has(modelName)) continue
    seen.add(modelName)
    const maxTokens = parseInt(m && m.maxTokens, 10)
    const model = { name: modelName, isDefault: !!(m && m.default) }
    if (Number.isInteger(maxTokens) && maxTokens > 0) model.maxTokens = maxTokens
    if (m && m.note) model.note = str(m.note)
    models.push(model)
  }
  if (!models.length) fail('llm.models 至少需要一项，且每项要有 name')

  // 有且仅有一个默认模型；未标注默认时取第一个
  const defaults = models.filter((m) => m.isDefault)
  if (defaults.length > 1) {
    models.forEach((m) => (m.isDefault = m === defaults[0]))
  } else if (defaults.length === 0) {
    models[0].isDefault = true
  }

  const timeoutMs = Math.min(120, Math.max(3, parseInt(llm.timeoutMs, 10) || 15)) * 1000
  const now = new Date().toISOString()
  const id = 'svc-' + crypto.randomUUID()

  const llmsJson = {
    defaultServiceId: id,
    services: [
      {
        id,
        name,
        baseUrl,
        apiKey,
        timeoutMs,
        note: str(llm.note),
        models,
        createdAt: now,
        updatedAt: now,
      },
    ],
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true })
  fs.writeFileSync(outPath, JSON.stringify(llmsJson, null, 2), 'utf8')
  console.log(`[apply-cfg] 已生成 ${outPath}`)
  console.log(`[apply-cfg] 服务：${name} (${baseUrl})，模型：${models.map((m) => m.name).join(', ')}`)
}

main()
