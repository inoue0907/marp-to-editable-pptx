import "./style.css"
import { convertMarp } from "./marpConverter"
import { exportEditablePptx } from "./exportPptx"
import { ConversionResult } from "./model"
import { SAMPLE_CSS, SAMPLE_MARKDOWN } from "./sample"

const app = document.querySelector<HTMLDivElement>("#app")!

app.innerHTML = `
  <header class="topbar">
    <div class="brand"><span class="brand-mark">M</span><span>Marp to PowerPoint</span><small>editable exporter</small></div>
    <div class="actions">
      <label class="file-button">Marpファイル<input id="md-file" type="file" accept=".md,.markdown,text/markdown" /></label>
      <label class="file-button secondary">テーマCSS<input id="css-file" type="file" accept=".css,text/css" /></label>
      <button id="convert" class="primary">変換してPPTX</button>
    </div>
  </header>
  <main>
    <section class="editor-pane">
      <div class="pane-title"><span>Marp Markdown</span><span id="status" class="status">準備完了</span></div>
      <textarea id="markdown" spellcheck="false"></textarea>
      <details>
        <summary>Theme CSS <span>任意</span></summary>
        <textarea id="theme-css" class="css-editor" spellcheck="false"></textarea>
      </details>
    </section>
    <section class="preview-pane">
      <div class="pane-title"><span>Browser preview</span><span id="slide-count" class="muted"></span></div>
      <div id="preview" class="preview"></div>
      <div id="report" class="report empty">変換すると、編集可能要素とフォールバック情報を表示します。</div>
    </section>
  </main>
  <div id="conversion-stage" aria-hidden="true"></div>
`

const markdown = document.querySelector<HTMLTextAreaElement>("#markdown")!
const themeCss = document.querySelector<HTMLTextAreaElement>("#theme-css")!
const preview = document.querySelector<HTMLDivElement>("#preview")!
const report = document.querySelector<HTMLDivElement>("#report")!
const status = document.querySelector<HTMLSpanElement>("#status")!
const slideCount = document.querySelector<HTMLSpanElement>("#slide-count")!
const convertButton = document.querySelector<HTMLButtonElement>("#convert")!
markdown.value = SAMPLE_MARKDOWN
themeCss.value = SAMPLE_CSS

let latest: ConversionResult | null = null
let renderTimer = 0

function setPreview(html: string, css: string) {
  preview.innerHTML = `<style>${css}</style>${html}`
  const slides = Array.from(preview.querySelectorAll<HTMLElement>("section")).filter(
    (section) => {
      const layer = section.dataset.marpitAdvancedBackground
      return !layer || layer === "content"
    },
  )
  slideCount.textContent = `${Math.max(1, slides.length)} slides`
}

async function render(exportAfter = false) {
  status.textContent = exportAfter ? "PPTXを生成中…" : "プレビュー更新中…"
  status.className = "status working"
  convertButton.disabled = true
  try {
    const converted = await convertMarp(markdown.value, themeCss.value)
    latest = converted.result
    setPreview(converted.previewHtml, converted.previewCss)
    const total = latest.nativeElements + latest.rasterElements
    const nativeRate = total ? Math.round(latest.nativeElements / total * 100) : 0
    report.className = "report"
    report.innerHTML = `
      <div class="metrics"><strong>${nativeRate}%</strong><span>editable</span><strong>${latest.slides.length}</strong><span>slides</span><strong>${latest.warnings.length}</strong><span>warnings</span></div>
      ${latest.warnings.length ? `<ul>${latest.warnings.map((warning) => `<li>Slide ${warning.slide}: ${warning.message}</li>`).join("")}</ul>` : `<p>主要コンテンツをすべて編集可能要素として変換しました。</p>`}
    `
    if (exportAfter) await exportEditablePptx(latest)
    status.textContent = exportAfter ? "PPTXを保存しました" : "プレビュー更新済み"
    status.className = "status success"
  } catch (error) {
    console.error(error)
    status.textContent = "変換エラー"
    status.className = "status error"
    report.className = "report error-box"
    report.textContent = error instanceof Error ? error.message : String(error)
  } finally {
    convertButton.disabled = false
  }
}

function scheduleRender() {
  window.clearTimeout(renderTimer)
  renderTimer = window.setTimeout(() => void render(false), 600)
}

markdown.addEventListener("input", scheduleRender)
themeCss.addEventListener("input", scheduleRender)
convertButton.addEventListener("click", () => void render(true))

document.querySelector<HTMLInputElement>("#md-file")!.addEventListener("change", async (event) => {
  const file = (event.currentTarget as HTMLInputElement).files?.[0]
  if (!file) return
  markdown.value = await file.text()
  void render(false)
})

document.querySelector<HTMLInputElement>("#css-file")!.addEventListener("change", async (event) => {
  const file = (event.currentTarget as HTMLInputElement).files?.[0]
  if (!file) return
  themeCss.value = await file.text()
  void render(false)
})

void render(false)
