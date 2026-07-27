#!/usr/bin/env node
import fs from "node:fs/promises"
import path from "node:path"
import process from "node:process"
import { build, preview } from "vite"
import { chromium } from "playwright-core"

function usage() {
  console.log(`Usage: marp-pptx <slides.md> [options]

Options:
  -t, --theme <theme.css>  Marp theme CSS
  -o, --output <file.pptx> Output path (default: next to Markdown)
  --html-output <file.html> Save the resolved Marp HTML
  --html-only                Skip PowerPoint generation
  --callout-style <style>    blockquote rendering: shape (default) or textbox
  --headed                  Show the browser while converting
  -h, --help                Show this help`)
}

function parseArgs(argv) {
  const options = { input: "", theme: "", output: "", htmlOutput: "", htmlOnly: false, headed: false, calloutStyle: "shape" }
  for (let index = 0; index < argv.length; index++) {
    const argument = argv[index]
    if (argument === "-h" || argument === "--help") return { ...options, help: true }
    if (argument === "--headed") options.headed = true
    else if (argument === "--html-only") options.htmlOnly = true
    else if (argument === "--html-output") options.htmlOutput = argv[++index] ?? ""
    else if (argument === "--callout-style") {
      options.calloutStyle = argv[++index] ?? ""
      if (!["shape", "textbox"].includes(options.calloutStyle)) throw new Error("--callout-style must be shape or textbox")
    }
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
  ".mp4": "video/mp4", ".webm": "video/webm", ".ogv": "video/ogg",
  ".woff": "font/woff", ".woff2": "font/woff2", ".ttf": "font/ttf", ".otf": "font/otf",
}

async function fileDataUrl(reference, baseDir) {
  if (/^(?:data:|blob:|#)/i.test(reference)) return reference
  let localReference = reference
  if (/^https?:/i.test(reference)) {
    try {
      const response = await fetch(reference)
      if (response.ok) {
        const data = Buffer.from(await response.arrayBuffer())
        const contentType = response.headers.get("content-type")?.split(";", 1)[0] || "application/octet-stream"
        return `data:${contentType};base64,${data.toString("base64")}`
      }
    } catch { /* try a checked-out copy below */ }
    // Offline-friendly fallback for themes that point at their own GitHub raw
    // assets even though the repository contains the same file.
    localReference = path.basename(new URL(reference).pathname)
  }
  const clean = decodeURIComponent(localReference.split(/[?#]/, 1)[0])
  const candidates = []
  if (path.isAbsolute(clean)) {
    candidates.push(clean)
  } else {
    // Community themes are frequently published with CSS in a `themes/`
    // directory while their url("./images/...") paths assume the repository
    // root. Prefer the CSS/Markdown directory, then search its nearest
    // ancestors so the resolved browser HTML remains portable.
    let directory = path.resolve(baseDir)
    for (let depth = 0; depth < 6; depth += 1) {
      candidates.push(path.resolve(directory, clean))
      const parent = path.dirname(directory)
      if (parent === directory) break
      directory = parent
    }
  }
  for (const candidate of candidates) {
    try {
      const data = await fs.readFile(candidate)
      const mime = mimeTypes[path.extname(candidate).toLowerCase()] ?? "application/octet-stream"
      return `data:${mime};base64,${data.toString("base64")}`
    } catch { /* try the next nearest base directory */ }
  }
  return reference
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
    const directive = match[1].trim()
    // Advanced backgrounds must remain Marp image syntax so Marp can create
    // its split/background layers. Ordinary SVG data URLs are rejected by
    // Markdown-it's link sanitizer, so emit those as HTML images instead.
    if (/(^|\s)bg(\s|$)/.test(directive)) return `![${match[1]}](${source})`
    if (!/^data:image\/svg\+xml/i.test(source)) return `![${match[1]}](${source})`
    const filters = [...directive.matchAll(/\b(blur|brightness|contrast|drop-shadow|grayscale|hue-rotate|invert|opacity|saturate|sepia):([^\s]+)/g)]
      .map((item) => `${item[1]}(${item[2]})`)
    const width = directive.match(/\bwidth:([^\s]+)/)?.[1]
    const height = directive.match(/\bheight:([^\s]+)/)?.[1]
    const styles = [
      filters.length ? `filter:${filters.join(" ")}` : "",
      width ? `width:${width}` : "",
      height ? `height:${height}` : "",
    ].filter(Boolean).join(";")
    return `<img src="${source}" alt="${match[1].replaceAll('"', "&quot;")}"${styles ? ` style="${styles}"` : ""}>`
  })
  result = await replaceAsync(result, /(<img\b[^>]*?\bsrc=["'])([^"']+)(["'])/gi, async (match) =>
    `${match[1]}${await fileDataUrl(match[2], baseDir)}${match[3]}`)
  result = await replaceAsync(result, /(<(?:video|source)\b[^>]*?\bsrc=["'])([^"']+)(["'])/gi, async (match) =>
    `${match[1]}${await fileDataUrl(match[2], baseDir)}${match[3]}`)
  result = await replaceAsync(result, /url\(\s*(["']?)([^"')]+)\1\s*\)/gi, async (match) =>
    `url("${await fileDataUrl(match[2].trim(), baseDir)}")`)
  return result
}

async function inlineCssAssets(css, baseDir) {
  return replaceAsync(css, /url\(\s*(["']?)([^"')]+)\1\s*\)/gi, async (match) =>
    `url("${await fileDataUrl(match[2].trim(), baseDir)}")`)
}

const THEME_BOUNDARY = "/* @marp-pptx-theme-boundary */"

async function loadThemeCss(cssPath, visited = new Set()) {
  const resolvedPath = path.resolve(cssPath)
  if (visited.has(resolvedPath)) return ""
  visited.add(resolvedPath)
  const baseDir = path.dirname(resolvedPath)
  const css = await inlineCssAssets(await fs.readFile(resolvedPath, "utf8"), baseDir)
  const dependencies = []
  const imports = [...css.matchAll(/@import\s+(?:url\(\s*)?(["'])([^"']+)\1\s*\)?\s*;/gi)]
  for (const match of imports) {
    const reference = match[2].trim()
    if (/^(?:https?:|data:)/i.test(reference) || ["default", "gaia", "uncover"].includes(reference)) continue
    const candidates = [
      path.resolve(baseDir, reference),
      path.resolve(baseDir, `${reference}.css`),
    ]
    for (const candidate of candidates) {
      try {
        await fs.access(candidate)
        const dependency = await loadThemeCss(candidate, visited)
        if (dependency) dependencies.push(dependency)
        break
      } catch { /* try next candidate */ }
    }
  }
  return [...dependencies, css].join(`\n${THEME_BOUNDARY}\n`)
}

function themeName(markdown) {
  const frontMatter = markdown.match(/^---\s*\r?\n([\s\S]*?)\r?\n---/)
  return frontMatter?.[1].match(/^theme:\s*["']?([^\s"']+)/m)?.[1] ?? ""
}

function parseJsonc(source) {
  return JSON.parse(source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/^\s*\/\/.*$/gm, "")
    .replace(/,\s*([}\]])/g, "$1"))
}

async function vscodeThemeCandidates(inputPath) {
  const candidates = []
  let directory = path.dirname(inputPath)
  const root = path.parse(directory).root
  while (true) {
    const settingsPath = path.join(directory, ".vscode", "settings.json")
    try {
      const settings = parseJsonc(await fs.readFile(settingsPath, "utf8"))
      const themes = settings["markdown.marp.themes"]
      if (Array.isArray(themes)) {
        for (const theme of themes) {
          if (typeof theme === "string" && !/^https?:/i.test(theme)) candidates.push(path.resolve(directory, theme))
        }
      }
    } catch { /* no settings at this level */ }
    if (directory === root) break
    directory = path.dirname(directory)
  }
  return candidates
}

async function cssThemeName(cssPath) {
  try {
    const css = await fs.readFile(cssPath, "utf8")
    return css.match(/\/\*[\s\S]*?@theme\s+([^*\s]+)[\s\S]*?\*\//)?.[1] ?? ""
  } catch {
    return ""
  }
}

async function detectTheme(inputPath, markdown, explicitTheme) {
  if (explicitTheme) return path.resolve(explicitTheme)
  const directory = path.dirname(inputPath)
  const name = themeName(markdown)
  const configuredThemes = await vscodeThemeCandidates(inputPath)
  for (const configuredTheme of configuredThemes) {
    if (!name || await cssThemeName(configuredTheme) === name) return configuredTheme
  }
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
  console.log(`Reading Markdown: ${inputPath}`)
  const rawMarkdown = await fs.readFile(inputPath, "utf8")
  const markdown = await inlineMarkdownAssets(rawMarkdown, path.dirname(inputPath))
  const themePath = await detectTheme(inputPath, rawMarkdown, options.theme)
  console.log(themePath ? `Using theme: ${themePath}` : "Using Marp built-in theme")
  const themeCss = themePath ? await loadThemeCss(themePath) : ""
  const outputPath = path.resolve(options.output || path.join(path.dirname(inputPath), `${path.basename(inputPath, path.extname(inputPath))}.pptx`))

  const projectRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(.:)/, "$1")), "..")
  await ensureBuild(projectRoot)
  const server = await preview({ root: projectRoot, logLevel: "error", preview: { host: "127.0.0.1", port: 0 } })
  const browser = await launchBrowser(!options.headed)
  try {
    console.log("Rendering Marp layout in the browser...")
    const url = server.resolvedUrls?.local[0]
    if (!url) throw new Error("Conversion server failed to start.")
    const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } })
    page.on("pageerror", (error) => console.error(`[browser] ${error.message}`))
    await page.goto(url, { waitUntil: "networkidle" })
    await page.waitForFunction(() => document.querySelector("#status")?.textContent === "プレビュー更新済み", null, { timeout: 60_000 })
    await page.evaluate((style) => { globalThis.__marpPptxCalloutStyle = style }, options.calloutStyle)
    await page.evaluate(({ markdown, themeCss }) => {
      const markdownInput = document.querySelector("#markdown")
      const themeInput = document.querySelector("#theme-css")
      markdownInput.value = markdown
      themeInput.value = themeCss
      markdownInput.dispatchEvent(new Event("input", { bubbles: true }))
      themeInput.dispatchEvent(new Event("input", { bubbles: true }))
    }, { markdown, themeCss })
    await page.waitForFunction(() => document.querySelector("#status")?.textContent === "プレビュー更新中…", null, { timeout: 60_000 })
    await page.waitForFunction(() => document.querySelector("#status")?.textContent === "プレビュー更新済み", null, { timeout: 60_000 })
    const conversionHost = page.locator(".marp-conversion-host")
    await conversionHost.evaluate(async (host) => {
      await Promise.all(Array.from(host.querySelectorAll("img")).map(async (image) => {
        try { await image.decode() } catch { /* conversion reports real failures later */ }
      }))
      await Promise.all(Array.from(host.querySelectorAll("video")).map((video) => {
        if (video.readyState >= HTMLMediaElement.HAVE_CURRENT_DATA) return
        return new Promise((resolve) => {
          const done = () => resolve()
          video.addEventListener("loadeddata", done, { once: true })
          video.addEventListener("error", done, { once: true })
          setTimeout(done, 10_000)
          video.load()
        })
      }))
    })
    if (options.htmlOutput) {
      const resolvedHtml = await conversionHost.evaluate((host) => {
        const clone = host.cloneNode(true)
        clone.classList.remove("marp-conversion-host")
        return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Marp HTML export</title>
<style>
html, body { margin: 0; padding: 0; background: #111; }
body { display: flex; flex-direction: column; align-items: center; gap: 24px; padding: 24px; }
svg[data-marpit-svg] { display: block; width: 1280px; height: 720px; max-width: 100%; }
@media print {
  body { display: block; padding: 0; background: white; }
  svg[data-marpit-svg] { break-after: page; page-break-after: always; max-width: none; }
}
</style>
</head>
<body>${clone.innerHTML}</body>
</html>`
      })
      const htmlPath = path.resolve(options.htmlOutput)
      await fs.mkdir(path.dirname(htmlPath), { recursive: true })
      await fs.writeFile(htmlPath, resolvedHtml, "utf8")
      console.log(`HTML: ${htmlPath}`)
    }
    if (options.htmlOnly) return
    if (process.env.MARP_PPTX_DEBUG_LAYOUT === "1") {
      const layout = await conversionHost.evaluate((host) =>
        Array.from(host.querySelectorAll("section")).map((section) => ({
          className: section.className,
          section: section.getBoundingClientRect().toJSON(),
          text: Array.from(section.querySelectorAll("h1,h2,h3,h4,p,pre,marp-pre,ul,ol")).map((element) => {
            const range = document.createRange()
            range.selectNodeContents(element)
            const style = getComputedStyle(element)
            return {
              tag: element.tagName,
              text: element.textContent,
              element: element.getBoundingClientRect().toJSON(),
              range: range.getBoundingClientRect().toJSON(),
              fontSize: style.fontSize,
              lineHeight: style.lineHeight,
              padding: [style.paddingTop, style.paddingRight, style.paddingBottom, style.paddingLeft],
              listItems: ["UL", "OL", "MARP-UL", "MARP-OL"].includes(element.tagName)
                ? Array.from(element.querySelectorAll(":scope > li")).map((li) => {
                    const liStyle = getComputedStyle(li)
                    const markerStyle = getComputedStyle(li, "::marker")
                    return {
                      fontSize: liStyle.fontSize,
                      fontFamily: liStyle.fontFamily,
                      markerFontSize: markerStyle.fontSize,
                      markerFontFamily: markerStyle.fontFamily,
                      markerColor: markerStyle.color,
                    }
                  })
                : [],
              shadow: element.shadowRoot
                ? Array.from(element.shadowRoot.querySelectorAll("*")).map((child) => ({
                    tag: child.tagName,
                    part: child.getAttribute("part"),
                    rect: child.getBoundingClientRect().toJSON(),
                    transform: getComputedStyle(child).transform,
                  }))
                : [],
            }
          }),
        })))
      console.log(JSON.stringify(layout, null, 2))
    }
    await conversionHost.evaluate((host) => {
      host.style.setProperty("left", "0", "important")
      host.style.setProperty("position", "absolute", "important")
      host.style.setProperty("z-index", "99999", "important")
    })
    const previewSlides = conversionHost.locator("svg[data-marpit-svg]")
    console.log(`Capturing backgrounds for ${await previewSlides.count()} slides...`)
    const backgroundDataUrls = []
    for (let index = 0; index < await previewSlides.count(); index++) {
      const slide = previewSlides.nth(index)
      await slide.evaluate((svg) => {
        const maskedTextStyles = []
        const isFixedFooter = (element) => {
          const blockquote = element.closest("blockquote")
          if (!blockquote) return false
          const style = getComputedStyle(blockquote)
          return (style.position === "absolute" || style.position === "fixed") && style.bottom !== "auto"
        }
        svg.querySelectorAll("section blockquote").forEach((blockquote) => {
          if (!isFixedFooter(blockquote)) return
          ;[blockquote, ...blockquote.querySelectorAll("*")].forEach((target) => {
            const style = getComputedStyle(target)
            target.setAttribute("data-marp-pptx-fixed-footer", "")
            target.style.setProperty("--marp-pptx-fixed-footer-color", style.color)
            target.style.setProperty("--marp-pptx-fixed-footer-shadow", style.textShadow)
          })
        })
        svg.querySelectorAll(
          "section h1, section h2, section h3, section h4, section h5, section h6," +
          "section p, section li, section blockquote, section pre, section code," +
          "section header, section footer, section figcaption",
        ).forEach((element) => {
          if (isFixedFooter(element)) return
          ;[element, ...element.querySelectorAll("*")].forEach((target) => {
            const value = target.style.getPropertyValue("color")
            const priority = target.style.getPropertyPriority("color")
            const textShadow = target.style.getPropertyValue("text-shadow")
            const textShadowPriority = target.style.getPropertyPriority("text-shadow")
            maskedTextStyles.push({ target, value, priority, textShadow, textShadowPriority })
            target.style.setProperty("color", "transparent", "important")
            target.style.setProperty("text-shadow", "none", "important")
          })
        })
        svg.querySelectorAll("section code").forEach((code) => {
          if (isFixedFooter(code)) return
          if (code.closest("pre,marp-pre")) return
          const value = code.style.getPropertyValue("background-color")
          const priority = code.style.getPropertyPriority("background-color")
          maskedTextStyles.push({ target: code, backgroundColor: value, backgroundColorPriority: priority })
          code.style.setProperty("background-color", "transparent", "important")
        })
        svg.querySelectorAll("section pre, section marp-pre, section blockquote").forEach((pre) => {
          if (isFixedFooter(pre)) return
          const backgroundColor = pre.style.getPropertyValue("background-color")
          const backgroundColorPriority = pre.style.getPropertyPriority("background-color")
          const borderColor = pre.style.getPropertyValue("border-color")
          const borderColorPriority = pre.style.getPropertyPriority("border-color")
          const boxShadow = pre.style.getPropertyValue("box-shadow")
          const boxShadowPriority = pre.style.getPropertyPriority("box-shadow")
          maskedTextStyles.push({
            target: pre,
            backgroundColor,
            backgroundColorPriority,
            borderColor,
            borderColorPriority,
            boxShadow,
            boxShadowPriority,
          })
          pre.style.setProperty("background-color", "transparent", "important")
          pre.style.setProperty("border-color", "transparent", "important")
          pre.style.setProperty("box-shadow", "none", "important")
        })
        svg.__marpPptxMaskedTextStyles = maskedTextStyles
        svg.querySelectorAll("section img").forEach((image) => {
          const style = getComputedStyle(image)
          if (style.filter !== "none" || Number(style.opacity) < 1) {
            image.setAttribute("data-marp-pptx-raster", "")
            // Chromium can lose element opacity when an HTML image is painted
            // through Marpit's SVG foreignObject screenshot path. Bake it into
            // the filter chain so the raster background matches the final DOM.
            if (Number(style.opacity) < 1) {
              const filter = style.filter === "none" ? "" : `${style.filter} `
              image.style.setProperty("filter", `${filter}opacity(${style.opacity})`, "important")
              image.style.setProperty("opacity", "1", "important")
            }
          }
        })
        const mask = document.createElementNS("http://www.w3.org/2000/svg", "style")
        mask.setAttribute("data-marp-pptx-cli-mask", "")
        mask.textContent = `
          section h1, section h2, section h3, section h4, section h5, section h6,
          section p, section li, section blockquote, section pre, section code,
          section header, section footer, section figcaption {
            color: transparent !important;
            text-shadow: none !important;
          }
          section h1 *, section h2 *, section h3 *, section h4 *,
          section h5 *, section h6 *, section p *, section li *,
          section blockquote *, section pre *, section code *,
          section header *, section footer *, section figcaption * {
            color: transparent !important;
            text-shadow: none !important;
          }
          section li::marker { color: transparent !important; content: "" !important; }
          section li::before, section li::after {
            content: none !important;
            display: none !important;
          }
          section img:not([data-marpit-advanced-background]) {
            visibility: hidden !important;
          }
          section video {
            visibility: hidden !important;
          }
          section img[alt*="blur:"], section img[alt*="brightness:"],
          section img[alt*="contrast:"], section img[alt*="drop-shadow:"],
          section img[alt*="grayscale:"], section img[alt*="hue-rotate:"],
          section img[alt*="invert:"], section img[alt*="opacity:"],
          section img[alt*="saturate:"], section img[alt*="sepia:"] {
            visibility: visible !important;
          }
          section img[data-marp-pptx-raster] {
            visibility: visible !important;
          }
          section[data-marpit-advanced-background="content"] img,
          section table { visibility: hidden !important; }
          section [data-marp-pptx-fixed-footer] {
            color: var(--marp-pptx-fixed-footer-color) !important;
            text-shadow: var(--marp-pptx-fixed-footer-shadow) !important;
          }
        `
        svg.appendChild(mask)
      })
      const initialBox = await slide.boundingBox()
      if (!initialBox) throw new Error(`Slide ${index + 1} has no rendered HTML bounds.`)
      await conversionHost.evaluate((host, y) => {
        host.style.setProperty("transform", `translateY(${-y}px)`, "important")
      }, initialBox.y)
      const box = await slide.boundingBox()
      if (!box) throw new Error(`Slide ${index + 1} has no rendered HTML bounds.`)
      const png = await page.screenshot({
        type: "png",
        animations: "disabled",
        clip: { x: box.x, y: box.y, width: box.width, height: box.height },
      })
      backgroundDataUrls.push(`data:image/png;base64,${png.toString("base64")}`)
      await conversionHost.evaluate((host) => host.style.removeProperty("transform"))
      await slide.evaluate((svg) => {
        svg.querySelector("[data-marp-pptx-cli-mask]")?.remove()
        svg.querySelectorAll("[data-marp-pptx-fixed-footer]").forEach((target) => {
          target.removeAttribute("data-marp-pptx-fixed-footer")
          target.style.removeProperty("--marp-pptx-fixed-footer-color")
          target.style.removeProperty("--marp-pptx-fixed-footer-shadow")
        })
        for (const saved of svg.__marpPptxMaskedTextStyles ?? []) {
          if (Object.hasOwn(saved, "backgroundColor")) {
            if (saved.backgroundColor) saved.target.style.setProperty("background-color", saved.backgroundColor, saved.backgroundColorPriority)
            else saved.target.style.removeProperty("background-color")
            if (Object.hasOwn(saved, "borderColor")) {
              if (saved.borderColor) saved.target.style.setProperty("border-color", saved.borderColor, saved.borderColorPriority)
              else saved.target.style.removeProperty("border-color")
              if (saved.boxShadow) saved.target.style.setProperty("box-shadow", saved.boxShadow, saved.boxShadowPriority)
              else saved.target.style.removeProperty("box-shadow")
            }
            continue
          }
          if (saved.value) saved.target.style.setProperty("color", saved.value, saved.priority)
          else saved.target.style.removeProperty("color")
          if (saved.textShadow) saved.target.style.setProperty("text-shadow", saved.textShadow, saved.textShadowPriority)
          else saved.target.style.removeProperty("text-shadow")
        }
        delete svg.__marpPptxMaskedTextStyles
      })
    }
    await conversionHost.evaluate((host) => {
      host.style.removeProperty("left")
      host.style.removeProperty("position")
      host.style.removeProperty("z-index")
    })
    await page.evaluate((backgrounds) => { globalThis.__marpPptxBackgrounds = backgrounds }, backgroundDataUrls)
    console.log("Building editable PowerPoint objects...")
    const downloadPromise = page.waitForEvent("download", { timeout: 120_000 })
      .then((download) => ({ download }), (error) => ({ error }))
    const errorPromise = page.waitForFunction(() => document.querySelector("#status")?.textContent === "変換エラー", null, { timeout: 120_000 })
      .then(async () => ({ error: new Error(await page.locator("#report").textContent() || "変換エラー") }), (error) => ({ error }))
    // Custom theme CSS may contain broad selectors such as `header` or
    // `section` that visually overlap the app chrome. CLI export does not need
    // pointer hit-testing, so dispatch the button action through the DOM.
    await page.locator("#convert").evaluate((button) => button.click())
    const outcome = await Promise.race([downloadPromise, errorPromise])
    if (outcome.error || !outcome.download) {
      const error = outcome.error ?? new Error("PPTX download did not start.")
      const status = await page.locator("#status").textContent().catch(() => "")
      throw new Error(`${error instanceof Error ? error.message : String(error)}${status ? `\nStatus: ${status}` : ""}`)
    }
    const download = outcome.download
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
