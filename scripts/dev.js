#!/usr/bin/env node
'use strict'

// 本地开发启动器：对齐 entrypoint.sh 的行为
//   1) 读 cfg.yml → 生成 data/config/llms.json
//   2) 启动 server/server.bundle.js
// 环境变量默认值见下，均可用环境变量覆盖。

const { spawn } = require('child_process')
const fs = require('fs')
const path = require('path')

const root = path.join(__dirname, '..')
const dataDir = process.env.DOC_AUDIT_DATA_DIR || path.join(root, 'data')
const bindHost = process.env.DOC_AUDIT_BIND_HOST || '0.0.0.0'
const port = process.env.DOC_AUDIT_PORT || '30001'
const cfgFile = process.env.DOC_AUDIT_CFG || path.join(root, 'cfg.yml')

process.env.DOC_AUDIT_DATA_DIR = dataDir
process.env.DOC_AUDIT_BIND_HOST = bindHost

function run(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, args, { stdio: 'inherit' })
    child.on('exit', (code) => (code === 0 ? resolve() : reject(new Error(`退出码 ${code}`))))
  })
}

async function main() {
  fs.mkdirSync(path.join(dataDir, 'config'), { recursive: true })
  const llmsJson = path.join(dataDir, 'config', 'llms.json')

  if (fs.existsSync(cfgFile)) {
    console.log(`[dev] 检测到 ${cfgFile}，正在生成 llms.json ...`)
    await run([path.join(root, 'config', 'apply-cfg.js'), cfgFile, llmsJson])
  } else {
    console.log(`[dev] 未检测到 ${cfgFile}，使用已有的 ${llmsJson}（如有）或 Web 设置页配置。`)
  }

  console.log(`[dev] 启动文档审核服务，监听 ${bindHost}:${port}，数据目录 ${dataDir}`)
  const child = spawn(process.execPath, [path.join(root, 'server', 'server.bundle.js'), port], { stdio: 'inherit' })
  child.on('exit', (code) => process.exit(code ?? 0))
}

main().catch((e) => {
  console.error('[dev]', e)
  process.exit(1)
})
