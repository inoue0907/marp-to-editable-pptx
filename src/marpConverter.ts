import { Marp } from "@marp-team/marp-core"
import marpBrowser from "@marp-team/marp-core/browser"
import { toPng } from "html-to-image"
import {
  ConversionResult,
  ImageElement,
  ListElement,
  SLIDE_H,
  SLIDE_W,
  SlideElement,
  SlideModel,
  ShapeElement,
  TextElement,
  TextRun,
  TableElement,
  VideoElement,
} from "./model"

const EDITABLE_SELECTOR = "h1,h2,h3,h4,h5,h6,p,blockquote,pre,marp-pre,ul,ol,marp-ul,marp-ol,table,img,video,header,footer"
const THEME_BOUNDARY = "/* @marp-pptx-theme-boundary */"

declare global {
  var __marpPptxBackgrounds: string[] | undefined
  var __marpPptxCleanup: (() => void) | undefined
  var __marpPptxCalloutStyle: "shape" | "textbox" | undefined
}

function calloutStyle(): "shape" | "textbox" {
  return globalThis.__marpPptxCalloutStyle === "textbox" ? "textbox" : "shape"
}

function hasThemeMetadata(source: string): boolean {
  return /\/\*[\s\S]*?@theme\s+[^\s*]+[\s\S]*?\*\//.test(source)
}

function isPreElement(element: Element): boolean {
  return element.tagName === "PRE" || element.tagName === "MARP-PRE"
}

function isListElement(element: Element): boolean {
  return ["UL", "OL", "MARP-UL", "MARP-OL"].includes(element.tagName)
}

function isFixedFooterBlockquote(element: Element): boolean {
  const blockquote = element.closest("blockquote")
  if (!(blockquote instanceof HTMLElement)) return false
  const style = getComputedStyle(blockquote)
  return (style.position === "absolute" || style.position === "fixed") && style.bottom !== "auto"
}

function isOrderedList(element: Element): boolean {
  return element.tagName === "OL" || element.tagName === "MARP-OL"
}
const PX_PER_IN = 96
const CSS_VIRTUAL_FONTS = new Set([
  "-apple-system",
  "blinkmacsystemfont",
  "system-ui",
  "sans-serif",
  "serif",
  "monospace",
])
const EAST_ASIAN_FONT_CANDIDATES = [
  "Yu Gothic UI", "Yu Gothic", "Meiryo", "Noto Sans JP", "Noto Sans CJK JP",
  "Hiragino Sans", "Hiragino Kaku Gothic ProN", "MS PGothic",
]

function cssColor(value: string, fallback = "#000000"): string {
  const match = value.match(/rgba?\((\d+)[, ]+(\d+)[, ]+(\d+)(?:\s*[,/]\s*([\d.]+))?/)
  if (match) {
    if (match[4] !== undefined && Number(match[4]) === 0) return fallback
    return `#${[match[1], match[2], match[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("")}`
  }
  const srgb = value.match(/color\(\s*srgb\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)(?:\s*\/\s*([\d.]+))?\s*\)/i)
  if (srgb) {
    if (srgb[4] !== undefined && Number(srgb[4]) === 0) return fallback
    return `#${[srgb[1], srgb[2], srgb[3]]
      .map((channel) => Math.round(Math.max(0, Math.min(1, Number(channel))) * 255).toString(16).padStart(2, "0"))
      .join("")}`
  }
  return value.startsWith("#") ? value : fallback
}

function px(value: string, fallback = 0): number {
  const parsed = Number.parseFloat(value)
  return Number.isFinite(parsed) ? parsed : fallback
}

function resolvedLineHeightPx(
  element: Element,
  style: CSSStyleDeclaration,
  fontPx = px(style.fontSize, 32),
): number {
  if (style.lineHeight !== "normal") return px(style.lineHeight, fontPx * 1.2)

  // Browsers intentionally expose the keyword "normal" instead of its used
  // pixel value. Measure the actual two-line box with the final computed font
  // rather than assuming the conventional but non-portable 1.2 multiplier.
  const probe = document.createElement("span")
  probe.textContent = "M\nM"
  Object.assign(probe.style, {
    position: "fixed",
    left: "-10000px",
    top: "0",
    visibility: "hidden",
    whiteSpace: "pre",
    padding: "0",
    margin: "0",
    border: "0",
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontStyle: style.fontStyle,
    fontWeight: style.fontWeight,
    fontStretch: style.fontStretch,
    fontVariant: style.fontVariant,
    lineHeight: "normal",
    letterSpacing: style.letterSpacing,
    writingMode: style.writingMode,
  })
  ;(element.ownerDocument.body ?? document.body).appendChild(probe)
  const measured = probe.getBoundingClientRect().height / 2
  probe.remove()
  return measured > 0 ? measured : fontPx * 1.2
}

function fontStack(value: string): string[] {
  return value.split(",")
    .map((font) => font.trim().replace(/^['"]|['"]$/g, ""))
    .filter((font) => font && !CSS_VIRTUAL_FONTS.has(font.toLowerCase()))
}

function resolvedFont(fontFamily: string): string {
  const stack = fontStack(fontFamily)
  const style = getComputedStyle(document.body)
  for (const candidate of stack) {
    if (isFontAvailable(candidate, style)) return candidate
  }
  return "Arial"
}

function weightedFontFace(fontFace: string, fontWeight: string): string {
  const weight = fontWeight === "bold" ? 700 : Number.parseInt(fontWeight, 10)
  if (weight === 600) return `${fontFace} SemiBold`
  if (weight === 500) return `${fontFace} Medium`
  if (weight === 300) return `${fontFace} Light`
  return fontFace
}

function measuredFontWidth(text: string, fontFamily: string, style: CSSStyleDeclaration): number {
  const canvas = document.createElement("canvas")
  const context = canvas.getContext("2d")
  if (!context) return 0
  context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${fontFamily}`
  return context.measureText(text).width
}

function isFontAvailable(candidate: string, style: CSSStyleDeclaration): boolean {
  const sample = "MarpWQ 123 日本語かなカナ"
  const candidateWidth = measuredFontWidth(sample, `"${candidate}", sans-serif`, style)
  const fallbackWidth = measuredFontWidth(sample, `"__marp_pptx_missing_font__", sans-serif`, style)
  return Math.abs(candidateWidth - fallbackWidth) > 0.05
}

function eastAsianFont(fontFamily: string, style: CSSStyleDeclaration): string {
  const sample = "漢字かなカナ、。配置編集"
  const targetWidth = measuredFontWidth(sample, fontFamily, style)
  const namedEastAsianFonts = fontStack(fontFamily).filter((font) =>
    /(gothic|meiryo|mincho|noto sans.*(?:jp|cjk)|hiragino|yahei|simsun|malgun|japanese)/i.test(font),
  )
  const candidates = [...namedEastAsianFonts, ...EAST_ASIAN_FONT_CANDIDATES]
  let best = resolvedFont(fontFamily)
  let difference = Number.POSITIVE_INFINITY
  for (const candidate of [...new Set(candidates)]) {
    if (!isFontAvailable(candidate, style)) continue
    const nextDifference = Math.abs(measuredFontWidth(sample, `"${candidate}"`, style) - targetWidth)
    if (nextDifference < difference) {
      best = candidate
      difference = nextDifference
    }
  }
  return best
}

function isEastAsianCharacter(character: string): boolean {
  return /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/u.test(character)
}

function applyResolvedFonts(runs: TextRun[], style: CSSStyleDeclaration): TextRun[] {
  const baseWeight = style.fontWeight === "bold" ? 700 : Number.parseInt(style.fontWeight, 10)
  const latin = weightedFontFace(resolvedFont(style.fontFamily), style.fontWeight)
  // PowerPoint does not reliably resolve named weight aliases of CJK variable
  // fonts (for example "Noto Sans JP SemiBold"). Keep the real family name;
  // semantic strong runs still use the bold flag independently.
  const eastAsian = eastAsianFont(style.fontFamily, style)
  const resolved: TextRun[] = []
  for (const run of runs) {
    for (const character of run.text) {
      const isEastAsian = isEastAsianCharacter(character)
      const fontFace = isEastAsian ? eastAsian : latin
      const bold = Boolean(run.bold) || baseWeight >= 700 || (isEastAsian && baseWeight >= 600)
      const previous = resolved.at(-1)
    if (previous && previous.fontFace === fontFace && previous.bold === bold && previous.italic === run.italic && previous.code === run.code && previous.fontSize === run.fontSize && previous.color === run.color && previous.highlight === run.highlight && previous.superscript === run.superscript && previous.subscript === run.subscript && previous.mathLatex === run.mathLatex) {
        previous.text += character
      } else {
        resolved.push({ ...run, text: character, fontFace, bold })
      }
    }
  }
  return resolved
}

function textRuns(
  root: Element,
  preserveVisualLines = false,
  skipNestedLists = false,
  skipDirectBlocks = false,
): TextRun[] {
  const runs: TextRun[] = []
  let lastLineTop: number | null = null
  let lastGlyphRight: number | null = null
  const rootStyle = getComputedStyle(root)
  const visualLineThreshold = Math.max(4, px(rootStyle.fontSize, 16) * 0.45)
  const preserveSourceLines = /^(pre|pre-wrap|break-spaces)$/.test(rootStyle.whiteSpace)
  const append = (text: string, state: Omit<TextRun, "text">) => {
    const previous = runs[runs.length - 1]
    if (previous && previous.bold === state.bold && previous.italic === state.italic && previous.code === state.code && previous.fontSize === state.fontSize && previous.color === state.color && previous.highlight === state.highlight && previous.superscript === state.superscript && previous.subscript === state.subscript) {
      previous.text += text
    } else {
      runs.push({ text, ...state })
    }
  }
  const visit = (node: Node, state: Omit<TextRun, "text">) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? ""
      if (!preserveVisualLines || preserveSourceLines) {
        if (text) append(text, state)
        return
      }
      for (let index = 0; index < text.length; index++) {
        const range = document.createRange()
        range.setStart(node, index)
        range.setEnd(node, index + 1)
        const rect = range.getClientRects()[0]
        const previousText = runs.at(-1)?.text ?? ""
        const changedLine = Boolean(
          rect && lastLineTop !== null && Math.abs(rect.top - lastLineTop) > visualLineThreshold,
        )
        if (changedLine && !previousText.endsWith("\n")) {
          append("\n", state)
          lastGlyphRight = null
        }
        // In normal HTML flow, source newlines/tabs are collapsed whitespace.
        // They must not become extra PowerPoint paragraphs in addition to the
        // visual line transition measured above.
        if (/\s/.test(text[index])) {
          const tail = runs.at(-1)?.text ?? ""
          if (tail && !/[\s\n]$/.test(tail)) append(" ", state)
        } else {
          // Inline boxes may add padding or margins between adjacent text
          // nodes. PowerPoint runs have no CSS box model, so preserve the
          // browser-resolved horizontal gap with non-breaking spacer glyphs.
          if (rect && !changedLine && lastGlyphRight !== null) {
            const gap = rect.left - lastGlyphRight
            if (gap > px(rootStyle.fontSize, 16) * 0.15 && !/[\s\n]$/.test(previousText)) {
              const spaces = Math.min(12, Math.max(1, Math.round(gap / (px(rootStyle.fontSize, 16) * 0.32))))
              append("\u00a0".repeat(spaces), state)
            }
          }
          append(text[index], state)
        }
        if (rect) {
          lastLineTop = rect.top
          if (!/\s/.test(text[index])) lastGlyphRight = rect.right
        }
      }
      return
    }
    if (!(node instanceof Element)) return
    const tag = node.tagName.toLowerCase()
    if (
      skipDirectBlocks
      && node.parentElement === root
      && ["block", "flex", "grid", "table", "list-item"].includes(getComputedStyle(node).display)
    ) return
    if (node.matches(".katex, mjx-container")) {
      const latex = node.querySelector('annotation[encoding="application/x-tex"]')?.textContent?.trim() ?? ""
      const compact = latex.match(/^\^(\d+)$/)?.[1] ?? latex
      if (compact) {
        const footnote = /^\^\d+$/.test(latex)
        append(compact, footnote
          ? { ...state, superscript: true }
          : {
              ...state,
              fontSize: px(rootStyle.fontSize, 16) * 0.7 * 72 / PX_PER_IN,
              mathLatex: latex,
              mathDisplay: false,
            })
      }
      return
    }
    // List items must be measured in their live, styled DOM. Do not clone them:
    // detached clones lose selector context, CSS variables, and inherited theme
    // styles. Nested lists are exported separately, so exclude only those
    // subtrees while walking the original item.
    if (skipNestedLists && node !== root && ["ul", "ol", "marp-ul", "marp-ol"].includes(tag)) return
    const nodeStyle = getComputedStyle(node)
    const nodeWeight = nodeStyle.fontWeight === "bold" ? 700 : Number.parseInt(nodeStyle.fontWeight, 10)
    const next = {
      bold: state.bold || tag === "strong" || tag === "b" || nodeWeight >= 700,
      italic: state.italic || tag === "em" || tag === "i",
      code: state.code || tag === "code",
      superscript: state.superscript || tag === "sup",
      subscript: state.subscript || tag === "sub",
      fontSize: px(nodeStyle.fontSize, px(rootStyle.fontSize, 16)) * 72 / PX_PER_IN,
      color: cssColor(nodeStyle.color),
      highlight: tag === "code" && isVisibleColor(nodeStyle.backgroundColor)
        ? cssColor(nodeStyle.backgroundColor)
        : state.highlight,
    }
    if (tag === "br") {
      const tail = runs.at(-1)?.text ?? ""
      if (!tail.endsWith("\n")) append("\n", next)
      lastLineTop = null
      lastGlyphRight = null
      return
    }
    node.childNodes.forEach((child) => visit(child, next))
    if (node !== root && (nodeStyle.display === "inline" || nodeStyle.display === "inline-block")) {
      const trailingSpace = px(nodeStyle.paddingRight) + px(nodeStyle.marginRight)
      if (trailingSpace > px(rootStyle.fontSize, 16) * 0.1) {
        const spaces = Math.min(
          12,
          Math.max(1, Math.round(trailingSpace / (px(nodeStyle.fontSize, px(rootStyle.fontSize, 16)) * 0.18))),
        )
        append("\u00a0".repeat(spaces), next)
        lastGlyphRight = null
      }
    }
  }
  root.childNodes.forEach((child) => visit(child, {}))
  return runs
    .map((run) => ({
      ...run,
      // Whitespace around visual line transitions is noise for normal HTML,
      // but in preformatted code it is the authored indentation.
      text: preserveSourceLines
        ? run.text
        : run.text.replace(/[ \t]+\n/g, "\n").replace(/\n[ \t]+/g, "\n"),
    }))
    .filter((run) => run.text.length > 0)
}

function firstTextGlyphRect(root: Element, skippedDirectBlocks: ReadonlySet<Element> = new Set()): DOMRect | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  while (walker.nextNode()) {
    const node = walker.currentNode
    if (Array.from(skippedDirectBlocks).some((block) => block.contains(node))) continue
    const text = node.textContent ?? ""
    const index = text.search(/\S/)
    if (index < 0) continue
    const range = document.createRange()
    range.setStart(node, index)
    range.setEnd(node, index + 1)
    const rect = range.getClientRects()[0]
    if (rect && rect.width > 0 && rect.height > 0) return rect
  }
  return null
}

function textContentRect(root: Element, skippedDirectBlocks: ReadonlySet<Element> = new Set()): DOMRect {
  const rects: DOMRect[] = []
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  while (walker.nextNode()) {
    const node = walker.currentNode
    if (Array.from(skippedDirectBlocks).some((block) => block.contains(node))) continue
    if (!(node.textContent ?? "").trim()) continue
    const range = document.createRange()
    range.selectNodeContents(node)
    rects.push(...Array.from(range.getClientRects()).filter((rect) => rect.width > 0 && rect.height > 0))
  }
  if (!rects.length) return new DOMRect()
  const left = Math.min(...rects.map((rect) => rect.left))
  const top = Math.min(...rects.map((rect) => rect.top))
  const right = Math.max(...rects.map((rect) => rect.right))
  const bottom = Math.max(...rects.map((rect) => rect.bottom))
  return new DOMRect(left, top, right - left, bottom - top)
}

function isBackgroundImage(img: HTMLImageElement): boolean {
  return /(^|\s)bg(\s|$)/.test(img.alt) ||
    img.hasAttribute("data-marpit-advanced-background") ||
    Boolean(img.closest("[data-marpit-advanced-background]"))
}

function isFilteredImage(img: HTMLImageElement): boolean {
  const style = getComputedStyle(img)
  return img.hasAttribute("data-marp-pptx-raster") ||
    style.filter !== "none" ||
    Number(style.opacity) < 1 ||
    /\b(?:blur|brightness|contrast|drop-shadow|grayscale|hue-rotate|invert|opacity|saturate|sepia):/.test(img.alt)
}

async function imageToDataUrl(src: string): Promise<string> {
  const blob = src.startsWith("data:")
    ? await (await fetch(src)).blob()
    : await (async () => {
        const response = await fetch(src)
        if (!response.ok) throw new Error(`画像を取得できません: ${src}`)
        return await response.blob()
      })()
  if (blob.type === "image/svg+xml") {
    const objectUrl = URL.createObjectURL(blob)
    try {
      const image = await loadImage(objectUrl)
      const canvas = document.createElement("canvas")
      canvas.width = Math.max(1, image.naturalWidth || 1280)
      canvas.height = Math.max(1, image.naturalHeight || 720)
      const context = canvas.getContext("2d")
      if (!context) throw new Error("SVG rasterization canvas is unavailable")
      context.drawImage(image, 0, 0, canvas.width, canvas.height)
      return canvas.toDataURL("image/png")
    } finally {
      URL.revokeObjectURL(objectUrl)
    }
  }
  return await new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(blob)
  })
}

async function videoElement(element: HTMLVideoElement, slide: HTMLElement): Promise<VideoElement> {
  const src = element.currentSrc || element.src
  const mime = src.match(/^data:([^;,]+)/i)?.[1] ?? ""
  const extn = mime === "video/webm" ? "webm" : mime === "video/ogg" ? "ogv" : "mp4"
  let cover: string | undefined
  if (element.poster) {
    try { cover = await imageToDataUrl(element.poster) } catch { /* use a frame below */ }
  }
  if (!cover && Number.isFinite(element.duration) && element.duration > 0) {
    const originalTime = element.currentTime
    let bestScore = -1
    const candidates = [0.98, 0.95, 0.9, 0.75].map((ratio) =>
      Math.min(Math.max(0.1, element.duration * ratio), Math.max(0.1, element.duration - 0.05)),
    )
    for (const thumbnailTime of candidates) {
      try {
        await new Promise<void>((resolve) => {
          let settled = false
          const done = () => {
            if (settled) return
            settled = true
            const videoWithFrameCallback = element as HTMLVideoElement & {
              requestVideoFrameCallback?: (callback: () => void) => number
            }
            if (videoWithFrameCallback.requestVideoFrameCallback) {
              videoWithFrameCallback.requestVideoFrameCallback(() => resolve())
            } else {
              requestAnimationFrame(() => resolve())
            }
          }
          element.addEventListener("seeked", done, { once: true })
          element.addEventListener("error", done, { once: true })
          setTimeout(() => resolve(), 3_000)
          element.currentTime = thumbnailTime
        })
        if (element.videoWidth <= 0 || element.videoHeight <= 0) continue
        const canvas = document.createElement("canvas")
        canvas.width = element.videoWidth
        canvas.height = element.videoHeight
        const context = canvas.getContext("2d", { willReadFrequently: true })
        if (!context) continue
        context.drawImage(element, 0, 0, canvas.width, canvas.height)
        const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
        let sum = 0
        let sumSquares = 0
        const stride = Math.max(4, Math.floor(pixels.length / 16_000 / 4) * 4)
        let count = 0
        for (let index = 0; index < pixels.length; index += stride) {
          const luminance = pixels[index] * 0.2126 + pixels[index + 1] * 0.7152 + pixels[index + 2] * 0.0722
          sum += luminance
          sumSquares += luminance * luminance
          count += 1
        }
        const mean = sum / count
        const variance = Math.max(0, sumSquares / count - mean * mean)
        const score = variance + mean * 0.25
        const isUsefulEndFrame = variance > 20 || (mean > 8 && mean < 247)
        if (score > bestScore) {
          bestScore = score
          cover = canvas.toDataURL("image/png")
        }
        if (isUsefulEndFrame) break
      } catch { /* try the next representative frame */ }
    }
    try { element.currentTime = originalTime } catch { /* no-op */ }
  }
  if (!cover && element.readyState >= HTMLMediaElement.HAVE_CURRENT_DATA && element.videoWidth > 0 && element.videoHeight > 0) {
    try {
      const canvas = document.createElement("canvas")
      canvas.width = element.videoWidth
      canvas.height = element.videoHeight
      const context = canvas.getContext("2d")
      if (context) {
        context.drawImage(element, 0, 0, canvas.width, canvas.height)
        cover = canvas.toDataURL("image/png")
      }
    } catch { /* PptxGenJS will provide its default media cover */ }
  }
  return { kind: "video", ...elementBox(element, slide), src, extn, cover }
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error("背景画像を読み込めませんでした"))
    image.src = src
  })
}

async function composeBackground(
  background: string,
  layers: string[],
  positionedImages: ImageElement[],
): Promise<string> {
  const scale = 2
  const canvas = document.createElement("canvas")
  canvas.width = 1280 * scale
  canvas.height = 720 * scale
  const context = canvas.getContext("2d")
  if (!context) throw new Error("背景合成用Canvasを作成できません")
  context.fillStyle = background
  context.fillRect(0, 0, canvas.width, canvas.height)
  for (const positioned of positionedImages) {
    const image = await loadImage(positioned.src)
    const x = positioned.x / SLIDE_W * canvas.width
    const y = positioned.y / SLIDE_H * canvas.height
    const w = positioned.w / SLIDE_W * canvas.width
    const h = positioned.h / SLIDE_H * canvas.height
    const imageRatio = image.naturalWidth / image.naturalHeight
    const boxRatio = w / h
    const sourceWidth = imageRatio > boxRatio ? image.naturalHeight * boxRatio : image.naturalWidth
    const sourceHeight = imageRatio > boxRatio ? image.naturalHeight : image.naturalWidth / boxRatio
    const sourceX = (image.naturalWidth - sourceWidth) / 2
    const sourceY = (image.naturalHeight - sourceHeight) / 2
    context.drawImage(image, sourceX, sourceY, sourceWidth, sourceHeight, x, y, w, h)
  }
  for (const layer of layers) {
    const image = await loadImage(layer)
    context.drawImage(image, 0, 0, canvas.width, canvas.height)
  }
  return canvas.toDataURL("image/png")
}

function offsetWithin(element: HTMLElement, ancestor: HTMLElement): { left: number; top: number } {
  let left = 0
  let top = 0
  let current: HTMLElement | null = element
  while (current && current !== ancestor) {
    left += current.offsetLeft
    top += current.offsetTop
    current = current.offsetParent as HTMLElement | null
  }
  if (current === ancestor) return { left, top }

  // Fallback for unusual custom-element/shadow-root boundaries.
  const rect = element.getBoundingClientRect()
  const ancestorRect = ancestor.getBoundingClientRect()
  return { left: rect.left - ancestorRect.left, top: rect.top - ancestorRect.top }
}

/**
 * Freeze the box from Marp's final rendered layout before writing any text.
 *
 * offsetTop/offsetLeft are not the final visual coordinates for flex/grid
 * alignment, transforms, foreignObject layers, or positioned theme content.
 * getBoundingClientRect() is the browser's resolved result for all of those.
 * Normalizing it by the section rect also removes any preview-scale transform.
 */
function elementBox(element: HTMLElement, slide: HTMLElement) {
  const rect = element.getBoundingClientRect()
  const slideRect = slide.getBoundingClientRect()
  const sx = SLIDE_W / slideRect.width
  const sy = SLIDE_H / slideRect.height
  return {
    x: (rect.left - slideRect.left) * sx,
    y: (rect.top - slideRect.top) * sy,
    w: rect.width * sx,
    h: rect.height * sy,
  }
}

function isVisibleColor(value: string): boolean {
  const alpha = value.match(/rgba?\([^)]*[,/]\s*([\d.]+)\s*\)$/)?.[1]
  return value !== "transparent" && (alpha === undefined || Number(alpha) > 0)
}

function clientRectBox(rect: DOMRect, slide: HTMLElement) {
  const slideRect = slide.getBoundingClientRect()
  const logicalX = 1280 / slideRect.width
  const logicalY = 720 / slideRect.height
  return {
    x: (rect.left - slideRect.left) * logicalX / PX_PER_IN,
    y: (rect.top - slideRect.top) * logicalY / PX_PER_IN,
    w: rect.width * logicalX / PX_PER_IN,
    h: rect.height * logicalY / PX_PER_IN,
  }
}

function elementDecorations(element: HTMLElement, slide: HTMLElement): ShapeElement[] {
  const shapes: ShapeElement[] = []
  const style = getComputedStyle(element)
  const box = elementBox(element, slide)
  if (element.tagName === "BLOCKQUOTE" && calloutStyle() === "shape") {
    if (isVisibleColor(style.backgroundColor)) {
      shapes.push({ kind: "shape", shape: "rect", ...box, fill: cssColor(style.backgroundColor) })
    }
    const borders = [
      { width: px(style.borderLeftWidth), color: style.borderLeftColor, x: box.x, y: box.y, w: 0, h: box.h },
      { width: px(style.borderTopWidth), color: style.borderTopColor, x: box.x, y: box.y, w: box.w, h: 0 },
      { width: px(style.borderRightWidth), color: style.borderRightColor, x: box.x + box.w, y: box.y, w: 0, h: box.h },
      { width: px(style.borderBottomWidth), color: style.borderBottomColor, x: box.x, y: box.y + box.h, w: box.w, h: 0 },
    ]
    for (const border of borders) {
      if (border.width <= 0 || !isVisibleColor(border.color)) continue
      shapes.push({
        kind: "shape",
        shape: "line",
        x: border.x,
        y: border.y,
        w: border.w,
        h: border.h,
        lineColor: cssColor(border.color),
        lineWidth: border.width * 72 / PX_PER_IN,
      })
    }
  }
  if (isPreElement(element) && isVisibleColor(style.backgroundColor)) {
    shapes.push({ kind: "shape", shape: "rect", ...box, fill: cssColor(style.backgroundColor) })
  }
  if (!isPreElement(element)) {
    element.querySelectorAll<HTMLElement>("code").forEach((code) => {
      const codeStyle = getComputedStyle(code)
      if (!isVisibleColor(codeStyle.backgroundColor)) return
      Array.from(code.getClientRects()).forEach((rect) => {
        const codeBox = clientRectBox(rect, slide)
        if (codeBox.w < 0.01 || codeBox.h < 0.01) return
        shapes.push({ kind: "shape", shape: "rect", ...codeBox, fill: cssColor(codeStyle.backgroundColor) })
      })
    })
  }
  return shapes
}

function mathTextElement(element: HTMLElement, slide: HTMLElement, latexQueue: { latex: string; display: boolean }[]): TextElement {
  const liveMath = Array.from(element.querySelectorAll<HTMLElement>("mjx-container, .katex")).map((node) => {
    const annotation = node.querySelector('annotation[encoding="application/x-tex"]')?.textContent?.trim()
    return {
      latex: annotation || latexQueue.shift()?.latex || "",
      display: node.hasAttribute("display") || node.classList.contains("katex-display") || Boolean(node.closest(".katex-display")),
    }
  })
  const clone = element.cloneNode(true) as HTMLElement
  const mathNodes = Array.from(clone.querySelectorAll("mjx-container, .katex"))
  const mathRuns: TextRun[] = []
  let markerIndex = 0
  for (const mathNode of mathNodes) {
    const marker = `__MARPMATH_${markerIndex++}__`
    mathNode.replaceWith(document.createTextNode(marker))
  }
  const plain = clone.textContent ?? ""
  const markerPattern = /__MARPMATH_(\d+)__/g
  let cursor = 0
  for (const match of plain.matchAll(markerPattern)) {
    if (match.index! > cursor) mathRuns.push({ text: plain.slice(cursor, match.index) })
    const math = liveMath[Number(match[1])]
    if (math) mathRuns.push({ text: match[0], mathLatex: math.latex, mathDisplay: math.display })
    cursor = match.index! + match[0].length
  }
  if (cursor < plain.length) mathRuns.push({ text: plain.slice(cursor) })
  const converted = textElement(element, slide)
  converted.runs = applyResolvedFonts(mathRuns, getComputedStyle(element))
  const nonMathText = plain.replace(markerPattern, "").trim()
  if (liveMath.some((math) => math.display) && !nonMathText) {
    // KaTeX/MathJax centers display math inside a full-width container. Their
    // hidden accessibility trees expose misleading "first glyph" geometry,
    // so using textElement's glyph anchor shifts the native equation far to
    // the right. Preserve the real parent box and let PowerPoint center the
    // editable OMML equation in that same box.
    Object.assign(converted, elementBox(element, slide))
    converted.align = "center"
    converted.valign = "middle"
    converted.margin = [0, 0, 0, 0]
  }
  return converted
}

function extractMath(markdown: string): { latex: string; display: boolean }[] {
  const matches: { index: number; latex: string; display: boolean }[] = []
  const blockRanges: [number, number][] = []
  for (const match of markdown.matchAll(/\$\$([\s\S]+?)\$\$/g)) {
    matches.push({ index: match.index!, latex: match[1].trim(), display: true })
    blockRanges.push([match.index!, match.index! + match[0].length])
  }
  for (const match of markdown.matchAll(/(?<!\$)\$(?!\$)([^$\n]+?)\$(?!\$)/g)) {
    if (blockRanges.some(([start, end]) => match.index! >= start && match.index! < end)) continue
    matches.push({ index: match.index!, latex: match[1].trim(), display: false })
  }
  return matches.sort((a, b) => a.index - b.index).map(({ latex, display }) => ({ latex, display }))
}

function textElement(
  element: HTMLElement,
  slide: HTMLElement,
  skippedDirectBlocks: ReadonlySet<Element> = new Set(),
): TextElement {
  const style = getComputedStyle(element)
  const box = elementBox(element, slide)
  const artworkBox = { ...box }
  const slideRect = slide.getBoundingClientRect()
  const scaleX = SLIDE_W / slideRect.width
  const scaleY = SLIDE_H / slideRect.height
  const paddingLeftPx = px(style.paddingLeft)
  const paddingRightPx = px(style.paddingRight)
  const paddingTopPx = px(style.paddingTop)
  const paddingBottomPx = px(style.paddingBottom)
  const left = paddingLeftPx * scaleX
  const right = paddingRightPx * scaleX
  box.x += left
  box.w = Math.max(0.1, box.w - left - right)

  // A theme may use the element's padding/background as artwork (for example,
  // a title-logo panel). The text itself does not begin at the element border.
  // Range geometry gives us the final browser-resolved line area without
  // splitting the content into one PowerPoint box per line.
  const fontPx = px(style.fontSize, 32)
  const contentRect = skippedDirectBlocks.size
    ? textContentRect(element, skippedDirectBlocks)
    : (() => {
        const contentRange = document.createRange()
        contentRange.selectNodeContents(element)
        return contentRange.getBoundingClientRect()
      })()
  if (!isPreElement(element) && contentRect.width > 0 && contentRect.height > 0) {
    const align = style.textAlign
    if (align !== "center" && align !== "right" && align !== "end") {
      const firstGlyph = firstTextGlyphRect(element, skippedDirectBlocks)
      box.x = ((firstGlyph?.left ?? contentRect.left) - slideRect.left) * scaleX
      // Browser and PowerPoint font metrics differ slightly. Keep the text box
      // anchored at the measured glyph origin but provide horizontal headroom
      // so short headings (especially CJK) do not wrap unexpectedly.
      box.w = Math.min(SLIDE_W - box.x, Math.max(box.w, contentRect.width * scaleX * 1.12))
    }
    box.y = (contentRect.top - slideRect.top) * scaleY
    box.h = Math.max(0.1, contentRect.height * scaleY * 1.08)
  }
  // For auto-scaled code, preserve authored newlines. Character ranges inside
  // Marp's scaling wrapper can report pre-transform wrap rows that are not
  // present in the final visual result.
  const preserveVisualLines = !isPreElement(element) && element.tagName !== "BLOCKQUOTE"
  let sourceRuns: TextRun[]
  if (element.tagName === "BLOCKQUOTE") {
    const paragraphs = Array.from(element.children).filter((child) => child.tagName === "P")
    sourceRuns = paragraphs.length > 0
      ? paragraphs.flatMap((paragraph, index) => {
          const paragraphRuns = textRuns(paragraph, true)
          if (paragraphRuns.length > 0) {
            paragraphRuns[0].text = paragraphRuns[0].text.trimStart()
            paragraphRuns[paragraphRuns.length - 1].text = paragraphRuns.at(-1)!.text.trimEnd()
          }
          return index < paragraphs.length - 1 ? [...paragraphRuns, { text: "\n" }] : paragraphRuns
        })
      : textRuns(element, false)
  } else {
    sourceRuns = textRuns(element, preserveVisualLines, false, skippedDirectBlocks.size > 0)
  }
  let unscaledRuns = applyResolvedFonts(sourceRuns, style)
  if (isPreElement(element)) {
    // Keep the browser-computed color of every syntax-highlighting span.
    // Collapsing the block to one run here used to discard token colors.
    unscaledRuns = applyResolvedFonts(textRuns(element, false), style)
      .map((run) => ({ ...run, code: true }))
  }
  const artworkPadding = paddingTopPx + paddingBottomPx > fontPx * 1.5
  let visualFontScale = 1
  if (artworkPadding) {
    const textRect = contentRect
    if (textRect.width > 0 && textRect.height > 0) {
      box.y = (textRect.top - slideRect.top) * scaleY
      box.h = Math.max(0.1, textRect.height * scaleY)
    }
  } else if (isPreElement(element)) {
    const range = document.createRange()
    range.selectNodeContents(element)
    const textRect = range.getBoundingClientRect()
    const lineHeightPx = style.lineHeight === "normal" ? fontPx * 1.2 : px(style.lineHeight, fontPx * 1.2)
    const lineCount = Math.max(1, unscaledRuns.map((run) => run.text).join("").trimEnd().split("\n").length)
    visualFontScale = Math.max(0.1, Math.min(1, textRect.height / (lineCount * lineHeightPx)))
    // Leave a small PowerPoint metric allowance after Marp has downscaled a
    // code block. Otherwise the final visual line can spill below the box even
    // when Chromium fits it exactly.
    if (visualFontScale < 0.9) visualFontScale *= 0.88
    if (textRect.height > 0) {
      box.y = (textRect.top - slideRect.top) * scaleY
      box.h = Math.max(0.1, textRect.height * scaleY)
    }
  }
  const lineHeightPx = resolvedLineHeightPx(element, style, fontPx)
  if (element.tagName === "BLOCKQUOTE") {
    const isBottomAnchored =
      style.position === "absolute"
      && style.bottom !== "auto"
    // Range geometry is measured at the browser glyph bounds. PowerPoint,
    // however, positions the text frame at the line box and adds the font's
    // ascent before drawing the first glyph. That difference is especially
    // visible for bottom-anchored footnotes: the box is correct but every line
    // is rendered lower and the last one falls off the slide. Compensate the
    // line-box/ascent offset without changing the resolved font size or width.
    if (isBottomAnchored) box.y = Math.max(0, box.y - lineHeightPx * scaleY * 1.1)
  }
  // A blockquote's editable text frame shares the full CSS background box.
  // Center its line box vertically so PowerPoint's extra font-top leading does
  // not make the text appear lower than the browser rendering.
  const valign = element.tagName === "BLOCKQUOTE" || style.justifyContent === "center"
    ? "middle"
    : style.justifyContent === "flex-end"
      ? "bottom"
      : "top"
  const runs = unscaledRuns.map((run) =>
    run.fontSize ? { ...run, fontSize: run.fontSize * visualFontScale } : run)
  const pre = isPreElement(element)
  const textboxCallout = element.tagName === "BLOCKQUOTE" && calloutStyle() === "textbox"
  if (pre || element.tagName === "BLOCKQUOTE") {
    box.x = artworkBox.x
    box.y = artworkBox.y
    box.w = artworkBox.w
    box.h = artworkBox.h
  }
  const codeFill = (pre || (element.tagName === "BLOCKQUOTE" && calloutStyle() === "textbox"))
    && isVisibleColor(style.backgroundColor)
    ? cssColor(style.backgroundColor)
    : undefined
  return {
    kind: "text",
    ...box,
    runs,
    fontSize: fontPx * visualFontScale * 72 / PX_PER_IN,
    fontFace: weightedFontFace(resolvedFont(style.fontFamily), style.fontWeight),
    color: cssColor(style.color),
    bold: Number(style.fontWeight) >= 700 || style.fontWeight === "bold",
    align: style.textAlign === "center" ? "center" : style.textAlign === "right" || style.textAlign === "end" ? "right" : "left",
    valign,
    lineSpacingMultiple: lineHeightPx / fontPx,
    lineSpacing: lineHeightPx * visualFontScale * 72 / PX_PER_IN,
    pre,
    fill: codeFill,
    lineColor: undefined,
    lineWidth: undefined,
    margin: pre || element.tagName === "BLOCKQUOTE"
      ? [
          paddingTopPx * 72 / PX_PER_IN,
          paddingRightPx * 72 / PX_PER_IN,
          paddingBottomPx * 72 / PX_PER_IN
            // PowerPoint vertically centers the font line box, whose ascent /
            // descent balance sits visibly lower than Chromium's glyph box.
            // Extra bottom inset shifts only the visible text upward while
            // preserving the authored blockquote geometry.
            + (element.tagName === "BLOCKQUOTE" ? fontPx * 0.5 * 72 / PX_PER_IN : 0),
          paddingLeftPx * 72 / PX_PER_IN,
        ]
      : undefined,
    fit: undefined,
  }
}

function textElements(element: HTMLElement, slide: HTMLElement): TextElement[] {
  if (element.tagName === "BLOCKQUOTE" || isPreElement(element)) return [textElement(element, slide)]
  const independentBlocks = new Set(
    Array.from(element.children).filter((child) => {
      if (!(child instanceof HTMLElement) || !(child.textContent ?? "").trim()) return false
      return ["block", "flex", "grid", "table", "list-item"].includes(getComputedStyle(child).display)
    }),
  )
  if (!independentBlocks.size) return [textElement(element, slide)]

  // A CSS block inside a heading/paragraph has its own horizontal origin and
  // box model. It cannot be represented faithfully as a rich-text run inside
  // the parent's PowerPoint box. Split only these authored CSS blocks (never
  // browser-wrapped visual lines), keeping ordinary inline content together.
  const output: TextElement[] = []
  const parent = textElement(element, slide, independentBlocks)
  if (parent.runs.some((run) => run.text.trim())) output.push(parent)
  independentBlocks.forEach((block) => output.push(...textElements(block as HTMLElement, slide)))
  return output
}

function listElement(element: HTMLElement, slide: HTMLElement): ListElement {
  const style = getComputedStyle(element)
  const firstLi = element.querySelector(":scope > li")
  const itemStyle = firstLi ? getComputedStyle(firstLi) : style
  const fontPx = px(itemStyle.fontSize, 32)
  const lineHeight = resolvedLineHeightPx(firstLi ?? element, itemStyle, fontPx)
  const box = elementBox(element, slide)
  const rootListRect = element.getBoundingClientRect()
  const items: ListElement["items"] = []
  const numberStyle = (value: string, suffix: string) => {
    const paren = suffix.includes(")")
    const both = suffix.includes("(") && paren
    const ending = both ? "ParenBoth" : paren ? "ParenR" : "Period"
    const styles: Record<string, string> = {
      "decimal": `arabic${ending}`,
      "lower-alpha": `alphaLc${ending}`,
      "lower-latin": `alphaLc${ending}`,
      "upper-alpha": `alphaUc${ending}`,
      "upper-latin": `alphaUc${ending}`,
      "lower-roman": `romanLc${ending}`,
      "upper-roman": `romanUc${ending}`,
    }
    return styles[value] ?? `arabic${ending}`
  }
  const markerProperties = (li: HTMLLIElement, list: HTMLElement) => {
    const pseudo = getComputedStyle(li, "::before")
    const marker = getComputedStyle(li, "::marker")
    const pseudoContent = pseudo.content && pseudo.content !== "none" ? pseudo.content : ""
    const markerStyle = pseudoContent ? pseudo : marker
    const quoted = pseudoContent.match(/^["'](.+)["']$/)?.[1]
    const counter = pseudoContent.match(/counter\([^,)]*(?:,\s*([^)]+))?\)(?:\s*["']([^"']*)["'])?/)
    const listStyle = counter?.[1]?.trim() || getComputedStyle(li).listStyleType || getComputedStyle(list).listStyleType
    const suffix = counter?.[2] || (pseudoContent.includes(")") ? ")" : ".")
    const fallbackCharacter = listStyle === "square" ? "▪" : listStyle === "circle" ? "◦" : "•"
    const markerCharacter = !isOrderedList(list) ? (quoted || fallbackCharacter) : undefined
    // PowerPoint's U+2022/U+25E6 glyphs are optically smaller than Chromium's
    // CSS list markers at the same point size. Keep the computed CSS size as
    // the source of truth and compensate only for the glyph-metric difference.
    const markerFontPx = px(markerStyle.fontSize, px(getComputedStyle(li).fontSize, fontPx))
    // PowerPoint's solid bullet diverges more from Chromium as the marker gets
    // larger. Interpolate the optical correction instead of using a
    // theme-specific constant: 16px and below needs none, 32px and above 1.25x.
    const solidBulletScale = 1 + Math.min(1, Math.max(0, (markerFontPx - 16) / 16)) * 0.25
    const markerScale = markerCharacter === "•" ? solidBulletScale
      : markerCharacter === "●" ? 0.5
      : markerCharacter === "◦" ? 1
      : markerCharacter === "▪" ? 1.15
      : 1
    return {
      markerCharacter,
      markerColor: cssColor(markerStyle.color || getComputedStyle(li).color),
      markerFontSize: markerFontPx * markerScale * 72 / PX_PER_IN,
      markerBold: Number(markerStyle.fontWeight) >= 600 || markerStyle.fontWeight === "bold",
      numberStyle: isOrderedList(list) ? numberStyle(listStyle, suffix) : undefined,
      markerFontPx,
      markerFontFamily: markerStyle.fontFamily || getComputedStyle(li).fontFamily,
      markerSuffix: suffix,
      pseudoMarker: Boolean(pseudoContent),
    }
  }
  const alphaIndex = (value: number) => String.fromCharCode(96 + Math.max(1, Math.min(26, value)))
  const romanIndex = (value: number) => {
    const pairs: [number, string][] = [[10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]]
    let remaining = Math.max(1, value)
    let result = ""
    for (const [amount, symbol] of pairs) {
      while (remaining >= amount) {
        result += symbol
        remaining -= amount
      }
    }
    return result
  }
  const orderedMarkerText = (index: number, style: string, suffix: string) => {
    const value = style.includes("alpha") || style.includes("latin")
      ? alphaIndex(index)
      : style.includes("roman")
        ? romanIndex(index)
        : String(index)
    const upper = style.startsWith("upper-") ? value.toUpperCase() : value
    return suffix.includes(")") ? `${upper})` : `${upper}.`
  }
  const measuredMarkerLeft = (
    li: HTMLLIElement,
    list: HTMLElement,
    textLeft: number,
    index: number,
    marker: ReturnType<typeof markerProperties>,
  ) => {
    // A generated ::before marker participates in the list item's own inline
    // flow, so its box starts at the list-item content edge.
    if (marker.pseudoMarker) return li.getBoundingClientRect().left

    // Native ::marker geometry is not exposed by the DOM API. Its right edge
    // is, however, defined immediately before the measured first text glyph.
    // Measure the actual computed marker font and include Chromium's marker
    // separator advance instead of inserting spaces into the list text.
    const markerStyle = getComputedStyle(li, "::marker")
    const listStyle = markerStyle.listStyleType || getComputedStyle(li).listStyleType || getComputedStyle(list).listStyleType
    const label = isOrderedList(list)
      ? orderedMarkerText(index, listStyle, marker.markerSuffix)
      : marker.markerCharacter ?? "•"
    const canvas = document.createElement("canvas")
    const context = canvas.getContext("2d")
    const weight = marker.markerBold ? "700" : (markerStyle.fontWeight || "400")
    if (context) context.font = `${weight} ${marker.markerFontPx}px ${marker.markerFontFamily}`
    const glyphWidth = context?.measureText(label).width ?? marker.markerFontPx * 0.5
    const separator = marker.markerFontPx * 0.5
    return textLeft - glyphWidth - separator
  }
  const normalizedRuns = (li: HTMLLIElement, itemStyle: CSSStyleDeclaration) => {
    const runs = applyResolvedFonts(textRuns(li, false, true), itemStyle)
      .map((run) => ({ ...run, text: run.text.replace(/\s+/g, " ") }))
      .filter((run) => run.text.length > 0)
    if (!runs.length) return runs
    runs[0] = { ...runs[0], text: runs[0].text.trimStart() }
    runs[runs.length - 1] = { ...runs[runs.length - 1], text: runs[runs.length - 1].text.trimEnd() }
    return runs.filter((run) => run.text.length > 0)
  }
  const directTextLeft = (li: HTMLLIElement) => {
    let left: number | null = null
    const visit = (node: Node) => {
      if (left !== null) return
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? ""
        const index = text.search(/\S/)
        if (index >= 0) {
          const range = document.createRange()
          range.setStart(node, index)
          range.setEnd(node, index + 1)
          left = range.getBoundingClientRect().left
        }
        return
      }
      if (!(node instanceof Element) || (node !== li && isListElement(node))) return
      node.childNodes.forEach(visit)
    }
    li.childNodes.forEach(visit)
    return left ?? li.getBoundingClientRect().left
  }
  const walk = (list: HTMLElement, level: number) => {
    const listLeft = list.getBoundingClientRect().left
    Array.from(list.children).filter((child): child is HTMLLIElement => child instanceof HTMLLIElement).forEach((li, index) => {
      const itemStyle = getComputedStyle(li)
      const textLeft = directTextLeft(li)
      const marker = markerProperties(li, list)
      const markerLeft = measuredMarkerLeft(li, list, textLeft, index + 1, marker)
      const {
        markerFontPx: _markerFontPx,
        markerFontFamily: _markerFontFamily,
        markerSuffix: _markerSuffix,
        pseudoMarker: _pseudoMarker,
        ...exportedMarker
      } = marker
      items.push({
        runs: normalizedRuns(li, itemStyle),
        index: index + 1,
        level,
        ordered: isOrderedList(list),
        fontSize: px(itemStyle.fontSize, fontPx) * 72 / PX_PER_IN,
        color: cssColor(itemStyle.color),
        markerIndent: (markerLeft - rootListRect.left) / PX_PER_IN,
        textIndent: (textLeft - rootListRect.left) / PX_PER_IN,
        lineSpacing: resolvedLineHeightPx(li, itemStyle, px(itemStyle.fontSize, fontPx)) * 72 / PX_PER_IN,
        paraSpaceBefore: Math.max(0, px(itemStyle.marginTop)) * 72 / PX_PER_IN,
        paraSpaceAfter: Math.max(0, px(itemStyle.marginBottom)) * 72 / PX_PER_IN,
        ...exportedMarker,
      })
      Array.from(li.children)
        .filter((child): child is HTMLElement => child instanceof HTMLElement && isListElement(child))
        .forEach((nested) => walk(nested, level + 1))
    })
  }
  walk(element, 0)
  return {
    kind: "list",
    ...box,
    items,
    fontSize: fontPx * 72 / PX_PER_IN,
    fontFace: weightedFontFace(resolvedFont(itemStyle.fontFamily), itemStyle.fontWeight),
    color: cssColor(itemStyle.color),
    lineSpacingMultiple: lineHeight / fontPx,
    lineSpacing: lineHeight * 72 / PX_PER_IN,
    // PptxGenJS defines bullet.indent as the marker-to-text distance, not
    // the list's whole CSS padding-left. One em matches Marp's marker gap.
    indent: px(itemStyle.paddingLeft, fontPx * 1.5) / PX_PER_IN,
  }
}

function tableElement(element: HTMLTableElement, slide: HTMLElement): TableElement {
  const box = elementBox(element, slide)
  const domRows = Array.from(element.rows)
  const firstRow = domRows[0]
  const colWidths = firstRow
    ? Array.from(firstRow.cells).map((cell) => cell.offsetWidth / PX_PER_IN)
    : [box.w]
  const rowHeights = domRows.map((row) => row.offsetHeight / PX_PER_IN)
  const rows = domRows.map((row) => Array.from(row.cells).map((cell) => {
    const style = getComputedStyle(cell)
    const fontPx = px(style.fontSize, 24)
    const background = cssColor(style.backgroundColor, "#FFFFFF")
    return {
      text: cell.textContent?.trim() ?? "",
      runs: applyResolvedFonts(textRuns(cell, false), style),
      fontSize: fontPx * 72 / PX_PER_IN,
      fontFace: isEastAsianCharacter(cell.textContent ?? "")
        ? eastAsianFont(style.fontFamily, style)
        : weightedFontFace(resolvedFont(style.fontFamily), style.fontWeight),
      color: cssColor(style.color),
      fill: background,
      bold: cell.tagName === "TH" || Number(style.fontWeight) >= 700,
      align: style.textAlign === "center" ? "center" as const : style.textAlign === "right" || style.textAlign === "end" ? "right" as const : "left" as const,
      valign: style.verticalAlign === "middle" ? "middle" as const : style.verticalAlign === "bottom" ? "bottom" as const : "top" as const,
      margin: [
        px(style.paddingTop) * 72 / PX_PER_IN,
        Math.max(0, px(style.paddingRight) * 72 / PX_PER_IN - 1),
        px(style.paddingBottom) * 72 / PX_PER_IN,
        Math.max(0, px(style.paddingLeft) * 72 / PX_PER_IN - 1),
      ] as [number, number, number, number],
      borderColor: cssColor(style.borderTopColor, "#CCCCCC"),
      borderWidth: Math.max(0.5, px(style.borderTopWidth, 1) * 72 / PX_PER_IN),
    }
  }))
  return { kind: "table", ...box, rows, colWidths, rowHeights }
}

async function captureBackground(slide: HTMLElement, editable: Element[], omitSlideBackground = false): Promise<string> {
  const backgroundColor = getComputedStyle(slide).backgroundColor
  const slideBackground = {
    color: slide.style.getPropertyValue("background-color"),
    colorPriority: slide.style.getPropertyPriority("background-color"),
    image: slide.style.getPropertyValue("background-image"),
    imagePriority: slide.style.getPropertyPriority("background-image"),
  }
  if (omitSlideBackground) {
    slide.style.setProperty("background-color", "transparent", "important")
    slide.style.setProperty("background-image", "none", "important")
  }
  const captureStyle = document.createElement("style")
  captureStyle.textContent = `
    [data-marp-pptx-background-capture],
    [data-marp-pptx-background-capture] * {
      color: transparent !important;
      text-shadow: none !important;
    }
    [data-marp-pptx-background-capture]::marker,
    [data-marp-pptx-background-capture] *::marker {
      color: transparent !important;
      content: "" !important;
    }
    [data-marp-pptx-background-capture] li,
    [data-marp-pptx-background-capture] li::marker {
      list-style: none !important;
      list-style-type: none !important;
    }
    [data-marp-pptx-background-list-item]::before,
    [data-marp-pptx-background-list-item]::after {
      content: none !important;
      display: none !important;
    }
    [data-marp-pptx-background-hidden] { visibility: hidden !important; }
    blockquote[data-marp-pptx-background-capture] {
      background: transparent !important;
      border-color: transparent !important;
      box-shadow: none !important;
    }
  `
  slide.appendChild(captureStyle)
  const marked = editable.map((element) => {
    const htmlElement = element as HTMLElement
    const attribute = element instanceof HTMLImageElement || element instanceof HTMLVideoElement || element.tagName === "TABLE"
      ? "data-marp-pptx-background-hidden"
      : "data-marp-pptx-background-capture"
    htmlElement.setAttribute(attribute, "")
    return { element: htmlElement, attribute }
  })
  const listItems = editable.flatMap((element) =>
    isListElement(element) ? Array.from(element.querySelectorAll<HTMLElement>("li")) : [],
  ).map((item) => {
    const value = item.style.getPropertyValue("list-style")
    const priority = item.style.getPropertyPriority("list-style")
    item.style.setProperty("list-style", "none", "important")
    item.setAttribute("data-marp-pptx-background-list-item", "")
    return { item, value, priority }
  })
  try {
    return await toPng(slide, {
      pixelRatio: 2,
      cacheBust: true,
      skipFonts: false,
      backgroundColor: omitSlideBackground ? "transparent" : backgroundColor,
    })
  } finally {
    marked.forEach(({ element, attribute }) => element.removeAttribute(attribute))
    listItems.forEach(({ item, value, priority }) => {
      item.removeAttribute("data-marp-pptx-background-list-item")
      if (value) item.style.setProperty("list-style", value, priority)
      else item.style.removeProperty("list-style")
    })
    if (omitSlideBackground) {
      if (slideBackground.color) slide.style.setProperty("background-color", slideBackground.color, slideBackground.colorPriority)
      else slide.style.removeProperty("background-color")
      if (slideBackground.image) slide.style.setProperty("background-image", slideBackground.image, slideBackground.imagePriority)
      else slide.style.removeProperty("background-image")
    }
    captureStyle.remove()
  }
}

function topLevelEditable(slide: HTMLElement): Element[] {
  return Array.from(slide.querySelectorAll(EDITABLE_SELECTOR)).filter((element) => {
    // Bottom-anchored footnotes are fixed slide furniture rather than flowing
    // content. Preserve their dashed rules, inline highlights, and exact line
    // wrapping together in the slide background.
    if (isFixedFooterBlockquote(element)) return false
    if (element instanceof HTMLImageElement && isBackgroundImage(element)) return false
    if (element instanceof HTMLImageElement && isFilteredImage(element)) return false
    if (element instanceof HTMLImageElement || element instanceof HTMLVideoElement) return true
    const nativeImages = Array.from(element.querySelectorAll("img")).filter((img) => !isBackgroundImage(img))
    if (nativeImages.length > 0 && !(element.textContent ?? "").trim()) return false
    if (element.matches("li *")) return false
    if (element.matches("ul *, ol *, marp-ul *, marp-ol *") && !isListElement(element)) return false
    return !element.parentElement?.closest(EDITABLE_SELECTOR) || element.parentElement === slide
  })
}

function extractNotes(markdown: string): string[] {
  const frontmatterStripped = markdown.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n?/, "")
  return frontmatterStripped.split(/^---\s*$/m).map((block) =>
    Array.from(block.matchAll(/<!--(?!\s*[_a-zA-Z-]+\s*:)([\s\S]*?)-->/g))
      .map((match) => match[1].trim()).filter(Boolean).join("\n\n")
  )
}

export async function convertMarp(markdown: string, customThemeCss = ""): Promise<{ result: ConversionResult; previewHtml: string; previewCss: string }> {
  globalThis.__marpPptxCleanup?.()
  globalThis.__marpPptxCleanup = undefined
  document.querySelectorAll(".marp-conversion-host").forEach((existing) => existing.remove())
  // Keep Unicode emoji as text. Marp's default Twemoji output depends on a
  // remote CDN and becomes a broken <img> in offline/corporate environments;
  // native emoji remains editable and portable in the PPTX.
  const marp = new Marp({ html: true, emoji: { unicode: true, shortcode: true } })
  const themeSources = customThemeCss.split(THEME_BOUNDARY).map((source) => source.trim()).filter(Boolean)
  // Marpit's theme registry accepts only CSS containing `@theme` metadata.
  // Plain CSS is still useful as an override, so append it after Marp's
  // generated CSS without trying to register it as a named theme.
  for (const themeSource of themeSources) {
    if (hasThemeMetadata(themeSource)) marp.themeSet.add(themeSource)
  }
  const rendered = marp.render(markdown)
  const host = document.createElement("div")
  host.className = "marp-conversion-host"
  host.innerHTML = rendered.html
  const style = document.createElement("style")
  style.textContent = `${rendered.css}\n${customThemeCss}`
  host.appendChild(style)
  document.body.appendChild(host)
  const browserController = marpBrowser(host)

  const warnings: ConversionResult["warnings"] = []
  const notes = extractNotes(markdown)
  const mathQueue = extractMath(markdown)
  const slides: SlideModel[] = []
  let nativeElements = 0
  let rasterElements = 0

  try {
    await document.fonts.ready
    browserController.update()
    // Auto-scaling uses ResizeObserver + requestAnimationFrame internally.
    // Wait until Marp has applied its transforms before reading any boxes.
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
    await new Promise((resolve) => setTimeout(resolve, 50))
    // Advanced backgrounds (`![bg ...]`) are emitted by Marp as additional
    // background-only <section> layers. They belong to the following content
    // slide and must never be counted as independent PowerPoint slides.
    const allSections = Array.from(host.querySelectorAll<HTMLElement>("section"))
    const sections = allSections.filter((section) => {
      const layer = section.dataset.marpitAdvancedBackground
      return !layer || layer === "content"
    })
    if (!sections.length) throw new Error("Marpスライドを検出できませんでした")

    for (let slideIndex = 0; slideIndex < sections.length; slideIndex++) {
      const section = sections[slideIndex]
      const slideRoot = (section.closest("svg[data-marpit-svg]") as unknown as HTMLElement | null) ?? section
      const editable = topLevelEditable(section)
      const elements: SlideElement[] = []
      const injectedBackground = globalThis.__marpPptxBackgrounds?.[slideIndex]

      for (const element of editable) {
        try {
          if (element instanceof HTMLImageElement) {
            const box = elementBox(element, slideRoot)
            elements.push({ kind: "image", ...box, src: await imageToDataUrl(element.currentSrc || element.src) } satisfies ImageElement)
          } else if (element instanceof HTMLVideoElement) {
            elements.push(await videoElement(element, slideRoot))
          } else if (element instanceof HTMLElement && isListElement(element)) {
            elements.push(listElement(element, slideRoot))
          } else if (element.tagName === "TABLE") {
            elements.push(tableElement(element as HTMLTableElement, slideRoot))
          } else if (element.querySelector("mjx-container, .katex")) {
            elements.push(mathTextElement(element as HTMLElement, slideRoot, mathQueue))
          } else {
            // CLI background capture already preserves CSS backgrounds and
            // inline-code highlights at exact browser geometry. Adding native
            // rectangles again would create a second, offset highlight.
            if (!injectedBackground || (element.tagName === "BLOCKQUOTE" && calloutStyle() === "shape")) {
              elements.push(...elementDecorations(element as HTMLElement, slideRoot))
            }
            elements.push(...textElements(element as HTMLElement, slideRoot))
          }
          nativeElements++
        } catch (error) {
          rasterElements++
          warnings.push({ slide: slideIndex + 1, message: error instanceof Error ? error.message : String(error) })
        }
      }

      const sectionStyle = getComputedStyle(section)
      const backgroundDataUrls: string[] = []
      try {
        if (injectedBackground) {
          backgroundDataUrls.push(injectedBackground)
        } else {
          const allIndex = allSections.indexOf(section)
          const advancedBackground = allSections[allIndex - 1]?.dataset.marpitAdvancedBackground === "background"
            ? allSections[allIndex - 1]
            : null
          if (advancedBackground) backgroundDataUrls.push(await captureBackground(advancedBackground, [], true))
          backgroundDataUrls.push(await captureBackground(section, editable))
        }
        rasterElements++
      } catch (error) {
        warnings.push({ slide: slideIndex + 1, message: `背景装飾の画像化に失敗: ${error instanceof Error ? error.message : String(error)}` })
      }
      const composedBackground = await composeBackground(
        cssColor(sectionStyle.backgroundColor, "#ffffff"),
        backgroundDataUrls,
        [],
      )
      slides.push({
        background: cssColor(sectionStyle.backgroundColor, "#ffffff"),
        backgroundDataUrls: [composedBackground],
        backgroundImages: [],
        elements,
        notes: notes[slideIndex],
      })
    }
  } finally {
    globalThis.__marpPptxCleanup = () => browserController.cleanup()
  }

  const titleMatch = markdown.match(/^title:\s*(.+)$/m)
  const authorMatch = markdown.match(/^author:\s*(.+)$/m)
  return {
    result: {
      slides,
      title: titleMatch?.[1].trim().replace(/^['"]|['"]$/g, "") || "marp-presentation",
      author: authorMatch?.[1].trim().replace(/^['"]|['"]$/g, "") || "",
      warnings,
      nativeElements,
      rasterElements,
    },
    previewHtml: rendered.html,
    previewCss: `${rendered.css}\n${customThemeCss}`,
  }
}
