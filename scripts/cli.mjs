#!/usr/bin/env node
import fs from "node:fs/promises"
import path from "node:path"
import process from "node:process"
import { build, preview } from "vite"
import { chromium } from "playwright-core"

function usage() {
  console.log(`Usage: marp-editable-pptx <slides.md> [options]

Options:
  -t, --theme <theme.css>  Marp theme CSS
  -o, --output <file.pptx> Output path (default: next to Markdown)
  --headed                  Show the browser while converting
  -h, --help                Show this help`)
}

function parseArgs(argv) {
  const options = { input: "", theme: "", output: "", headed: false }
  for (let index = 0; index < argv.length; index++) {
    const argument = argv[index]
    if (argument === "-h" || argument === "--help") return { ...options, help: true }
    if (argument === "--headed") options.headed = true
    else if (argument === "-t" || argument === "--theme") options.theme = argv[++index] ?? ""
    else if (argument === "-o" || argument === "--output") options.output = argv[++index] ?? ""
    else if (!argument.startsWith("-") && !options.input) options.input = argument
    else throw new Error(`Unknown argument: ${argument}`)
  }
  return options
}

const mimeTypes = {
  ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
  ".gif": "image/gif", ".webp": "image/webp", ".svg": "image/svg+xml",
  ".woff": "font/woff", ".woff2": "font/woff2", ".ttf": "font/ttf", ".otf": "font/otf",
}

async function fileDataUrl(reference, baseDir) {
  if (/^(?:data:|https?:|blob:|#)/i.test(reference)) return reference
  const clean = decodeURIComponent(reference.split(/[?#]/, 1)[0])
  const candidate = path.isAbsolute(clean) ? clean : path.resolve(baseDir, clean)
  try {
    const data = await fs.readFile(candidate)
    const mime = mimeTypes[path.extname(candidate).toLowerCase()] ?? "application/octet-stream"
    return `data:${mime};base64,${data.toString("base64")}`
  } catch {
    return reference
  }
}

async function replaceAsync(source, pattern, replacer) {
  const matches = [...source.matchAll(pattern)]
  if (!matches.length) return source
  let output = ""
  let cursor = 0
  for (const match of matches) {
    output += source.slice(cursor, match.index) + await replacer(match)
    cursor = match.index + match[0].length
  }
  return output + source.slice(cursor)
}

async function inlineMarkdownAssets(markdown, baseDir) {
  let result = await replaceAsync(markdown, /!\[([^\]]*)\]\(<?([^)>\s]+)>?(?:\s+["'][^"']*["'])?\)/g, async (match) => {
    const source = await fileDataUrl(match[2], baseDir)
    if (source === match[2]) return match[0]
    const dimensions = [...match[1].matchAll(/\b(width|height):\s*([^\s]+)/gi)]
      .map((dimension) => `${dimension[1].toLowerCase()}:${dimension[2]}`).join(";")
    const alt = match[1].replace(/["&<>]/g, (character) => ({ '"': "&quot;", "&": "&amp;", "<": "&lt;", ">": "&gt;" })[character])
    return `<img src="${source}" alt="${alt}"${dimensions ? ` style="${dimensions}"` : ""} />`
  })
  result = await replaceAsync(result, /(<img\b[^>]*?\bsrc=["'])([^"']+)(["'])/gi, async (match) =>
    `${match[1]}${await fileDataUrl(match[2], baseDir)}${match[3]}`)
  return result
}

async function inlineCssAssets(css, baseDir) {
  return replaceAsync(css, /url\(\s*(["']?)([^"')]+)\1\s*\)/gi, async (match) =>
    `url("${await fileDataUrl(match[2].trim(), baseDir)}")`)
}

function themeName(markdown) {
  const frontMatter = markdown.match(/^---\s*\r?\n([\s\S]*?)\r?\n---/)
  return frontMatter?.[1].match(/^theme:\s*["']?([^\s"']+)/m)?.[1] ?? ""
}

async function detectTheme(inputPath, markdown, explicitTheme) {
  if (explicitTheme) return path.resolve(explicitTheme)
  const directory = path.dirname(inputPath)
  const name = themeName(markdown)
  const candidates = [
    name && path.join(directory, "themes", `${name}.css`),
    name && path.join(process.cwd(), "themes", `${name}.css`),
    path.join(directory, "theme.css"),
  ].filter(Boolean)
  for (const candidate of candidates) {
    try { await fs.access(candidate); return candidate } catch { /* try next */ }
  }
  return ""
}

async function launchBrowser(headless) {
  const attempts = process.platform === "win32" ? ["msedge", "chrome"] : ["chrome", "msedge"]
  for (const channel of attempts) {
    try { return await chromium.launch({ channel, headless }) } catch { /* try next */ }
  }
  throw new Error("EdgeまたはChromeが見つかりません。どちらかをインストールしてください。")
}

async function newestModifiedTime(directory) {
  let newest = 0
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const target = path.join(directory, entry.name)
    if (entry.isDirectory()) newest = Math.max(newest, await newestModifiedTime(target))
    else newest = Math.max(newest, (await fs.stat(target)).mtimeMs)
  }
  return newest
}

async function ensureBuild(projectRoot) {
  const output = path.join(projectRoot, "dist", "index.html")
  let outputTime = 0
  try { outputTime = (await fs.stat(output)).mtimeMs } catch { /* build below */ }
  const sourceTime = Math.max(
    await newestModifiedTime(path.join(projectRoot, "src")),
    (await fs.stat(path.join(projectRoot, "index.html"))).mtimeMs,
  )
  if (outputTime < sourceTime) await build({ root: projectRoot, logLevel: "error" })
}

async function main() {
  const options = parseArgs(process.argv.slice(2))
  if (options.help) { usage(); return }
  if (!options.input) { usage(); throw new Error("Markdown file is required.") }

  const inputPath = path.resolve(options.input)
  const rawMarkdown = await fs.readFile(inputPath, "utf8")
  const markdown = await inlineMarkdownAssets(rawMarkdown, path.dirname(inputPath))
  const themePath = await detectTheme(inputPath, rawMarkdown, options.theme)
  const themeCss = themePath
    ? await inlineCssAssets(await fs.readFile(themePath, "utf8"), path.dirname(themePath))
    : ""
  const outputPath = path.resolve(options.output || path.join(path.dirname(inputPath), `${path.basename(inputPath, path.extname(inputPath))}.pptx`))

  const projectRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(.:)/, "$1")), "..")
  await ensureBuild(projectRoot)
  const server = await preview({ root: projectRoot, logLevel: "error", preview: { host: "127.0.0.1", port: 0 } })
  const browser = await launchBrowser(!options.headed)
  try {
    const url = server.resolvedUrls?.local[0]
    if (!url) throw new Error("Conversion server failed to start.")
    const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } })
    await page.goto(url, { waitUntil: "networkidle" })
    await page.evaluate(({ markdown, themeCss }) => {
      const markdownInput = document.querySelector("#markdown")
      const themeInput = document.querySelector("#theme-css")
      markdownInput.value = markdown
      themeInput.value = themeCss
      markdownInput.dispatchEvent(new Event("input", { bubbles: true }))
      themeInput.dispatchEvent(new Event("input", { bubbles: true }))
    }, { markdown, themeCss })
    await page.waitForFunction(() => document.querySelector("#status")?.textContent === "プレビュー更新済み", null, { timeout: 60_000 })
    const downloadPromise = page.waitForEvent("download", { timeout: 120_000 })
    await page.locator("#convert").click()
    const download = await downloadPromise
    await fs.mkdir(path.dirname(outputPath), { recursive: true })
    await download.saveAs(outputPath)
    console.log(`Created: ${outputPath}`)
    if (themePath) console.log(`Theme: ${themePath}`)
  } finally {
    await browser.close()
    server.httpServer.close()
  }
}

main().catch((error) => {
  console.error(`Error: ${error instanceof Error ? error.message : String(error)}`)
  process.exitCode = 1
})
