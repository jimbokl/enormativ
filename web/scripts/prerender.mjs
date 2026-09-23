import { readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import React from 'react'
import { renderToString } from 'react-dom/server'
import { createServer } from 'vite'

const root = fileURLToPath(new URL('..', import.meta.url))
const server = await createServer({
  configFile: join(root, 'vite.config.js'),
  root,
  server: { middlewareMode: true },
  appType: 'custom',
})

try {
  const { App } = await server.ssrLoadModule('/src/main.jsx')
  const pages = [
    { page: 'portal', file: 'index.html' },
    { page: 'check', file: 'tools/check/index.html' },
    { page: 'compare', file: 'tools/compare/index.html' },
  ]
  const titles = new Set()
  for (const { page, file } of pages) {
    const path = join(root, 'dist', file)
    const html = await readFile(path, 'utf8')
    const markup = renderToString(React.createElement(App, { page }))
    const title = html.match(/<title>(.*?)<\/title>/)?.[1]
    if (!title || titles.has(title)) throw new Error(`Missing or duplicate title: ${file}`)
    titles.add(title)
    if (!markup.includes('<h1>') || !markup.includes('</h1>')) throw new Error(`Missing rendered H1: ${file}`)
    if (!html.includes('<div id="root"></div>')) throw new Error(`Missing empty root: ${file}`)
    await writeFile(path, html.replace('<div id="root"></div>', `<div id="root">${markup}</div>`))
    process.stdout.write(`Prerendered ${file}\n`)
  }
} finally {
  await server.close()
}
