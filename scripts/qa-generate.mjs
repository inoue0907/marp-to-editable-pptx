import { createRequire } from "node:module"
import fs from "node:fs"
import path from "node:path"

const requireFromWorkspace = createRequire("C:/Users/kirim/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/package.json")
const { chromium } = requireFromWorkspace("playwright")

const outputDir = path.resolve("qa")
fs.mkdirSync(outputDir, { recursive: true })

const browser = await chromium.launch({ channel: "msedge", headless: true })
try {
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } })
  await page.goto("http://127.0.0.1:4173/", { waitUntil: "networkidle" })
  await page.locator("#status").filter({ hasText: /更新済み|準備完了/ }).waitFor({ timeout: 60_000 })

  await page.locator("#preview").screenshot({ path: path.join(outputDir, "marp-preview.png") })
  const previewSlides = page.locator("#preview section:not([data-marpit-advanced-background]), #preview section[data-marpit-advanced-background=\"content\"]")
  for (let index = 0; index < await previewSlides.count(); index++) {
    await previewSlides.nth(index).screenshot({ path: path.join(outputDir, `marp-slide-${index + 1}.png`) })
  }
  const downloadPromise = page.waitForEvent("download")
  await page.locator("#convert").click()
  const download = await downloadPromise
  await download.saveAs(path.join(outputDir, "generated-next.pptx"))

  const diagnostics = await page.evaluate(() => ({
    status: document.querySelector("#status")?.textContent,
    report: document.querySelector("#report")?.textContent,
    math: [...document.querySelectorAll("#preview mjx-container, #preview .katex")].map((node) => node.outerHTML),
    slides: [...document.querySelectorAll("#preview section")]
      .filter((section) => !section.dataset.marpitAdvancedBackground || section.dataset.marpitAdvancedBackground === "content")
      .map((section, slideIndex) => ({
        slide: slideIndex + 1,
        advanced: section.dataset.marpitAdvancedBackground ?? null,
        split: section.dataset.marpitAdvancedBackgroundSplit ?? null,
        section: {
          width: section.clientWidth,
          height: section.clientHeight,
          padding: getComputedStyle(section).padding,
        },
        blocks: [...section.querySelectorAll("h1,h2,h3,h4,h5,h6,p,ul,ol")].map((element) => {
          const style = getComputedStyle(element)
          const visualLines = []
          let lineTop = null
          let line = ""
          const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
          while (walker.nextNode()) {
            const node = walker.currentNode
            const text = node.textContent ?? ""
            for (let index = 0; index < text.length; index++) {
              const range = document.createRange()
              range.setStart(node, index)
              range.setEnd(node, index + 1)
              const rect = range.getClientRects()[0]
              if (rect && lineTop !== null && Math.abs(rect.top - lineTop) > 1) {
                visualLines.push(line.trimEnd())
                line = ""
              }
              line += text[index]
              if (rect) lineTop = rect.top
            }
          }
          if (line.trim()) visualLines.push(line.trimEnd())
          return {
            tag: element.tagName,
            text: element.textContent,
            x: element.offsetLeft,
            y: element.offsetTop,
            width: element.offsetWidth,
            height: element.offsetHeight,
            fontSize: style.fontSize,
            fontFamily: style.fontFamily,
            fontWeight: style.fontWeight,
            lineHeight: style.lineHeight,
            paddingLeft: style.paddingLeft,
            visualLines,
          }
        }),
      })),
  }))
  fs.writeFileSync(path.join(outputDir, "browser-layout.json"), JSON.stringify(diagnostics, null, 2))
  console.log(JSON.stringify(diagnostics, null, 2))
} finally {
  await browser.close()
}
