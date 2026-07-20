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
} from "./model"

const EDITABLE_SELECTOR = "h1,h2,h3,h4,h5,h6,p,blockquote,pre,ul,ol,table,img"
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
  if (!match) return value.startsWith("#") ? value : fallback
  if (match[4] !== undefined && Number(match[4]) === 0) return fallback
  return `#${[match[1], match[2], match[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("")}`
}

function px(value: string, fallback = 0): number {
  const parsed = Number.parseFloat(value)
  return Number.isFinite(parsed) ? parsed : fallback
}

function fontStack(value: string): string[] {
  return value.split(",")
    .map((font) => font.trim().replace(/^['"]|['"]$/g, ""))
    .filter((font) => font && !CSS_VIRTUAL_FONTS.has(font.toLowerCase()))
}

function resolvedFont(fontFamily: string): string {
  const stack = fontStack(fontFamily)
  return stack[0] ?? "Arial"
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
    if (!document.fonts.check(`${style.fontSize} "${candidate}"`, sample)) continue
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
      const fontFace = run.code ? "Courier New" : isEastAsian ? eastAsian : latin
      const bold = Boolean(run.bold) || baseWeight >= 700 || (isEastAsian && baseWeight >= 600)
      const previous = resolved.at(-1)
      if (previous && previous.fontFace === fontFace && previous.bold === bold && previous.italic === run.italic && previous.code === run.code && previous.fontSize === run.fontSize && previous.color === run.color && previous.mathLatex === run.mathLatex) {
        previous.text += character
      } else {
        resolved.push({ ...run, text: character, fontFace, bold })
      }
    }
  }
  return resolved
}

function textRuns(root: Element, preserveVisualLines = false): TextRun[] {
  const runs: TextRun[] = []
  let lastLineTop: number | null = null
  const rootStyle = getComputedStyle(root)
  const visualLineThreshold = Math.max(4, px(rootStyle.fontSize, 16) * 0.45)
  const preserveSourceLines = /^(pre|pre-wrap|break-spaces)$/.test(rootStyle.whiteSpace)
  const append = (text: string, state: Omit<TextRun, "text">) => {
    const previous = runs[runs.length - 1]
    if (previous && previous.bold === state.bold && previous.italic === state.italic && previous.code === state.code && previous.fontSize === state.fontSize && previous.color === state.color) {
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
        if (rect && lastLineTop !== null && Math.abs(rect.top - lastLineTop) > visualLineThreshold && !previousText.endsWith("\n")) {
          append("\n", state)
        }
        // In normal HTML flow, source newlines/tabs are collapsed whitespace.
        // They must not become extra PowerPoint paragraphs in addition to the
        // visual line transition measured above.
        if (/\s/.test(text[index])) {
          const tail = runs.at(-1)?.text ?? ""
          if (tail && !/[\s\n]$/.test(tail)) append(" ", state)
        } else {
          append(text[index], state)
        }
        if (rect) lastLineTop = rect.top
      }
      return
    }
    if (!(node instanceof Element)) return
    const tag = node.tagName.toLowerCase()
    const nodeStyle = getComputedStyle(node)
    const next = {
      bold: state.bold || tag === "strong" || tag === "b",
      italic: state.italic || tag === "em" || tag === "i",
      code: state.code || tag === "code",
      fontSize: px(nodeStyle.fontSize, px(rootStyle.fontSize, 16)) * 72 / PX_PER_IN,
      color: cssColor(nodeStyle.color),
    }
    if (tag === "br") {
      const tail = runs.at(-1)?.text ?? ""
      if (!tail.endsWith("\n")) append("\n", next)
      lastLineTop = null
      return
    }
    node.childNodes.forEach((child) => visit(child, next))
  }
  root.childNodes.forEach((child) => visit(child, {}))
  return runs
    .map((run) => ({ ...run, text: run.text.replace(/[ \t]+\n/g, "\n").replace(/\n[ \t]+/g, "\n") }))
    .filter((run) => run.text.length > 0)
}

function isBackgroundImage(img: HTMLImageElement): boolean {
  return /(^|\s)bg(\s|$)/.test(img.alt) ||
    img.hasAttribute("data-marpit-advanced-background") ||
    Boolean(img.closest("[data-marpit-advanced-background]"))
}

async function imageToDataUrl(src: string): Promise<string> {
  if (src.startsWith("data:")) return src
  const response = await fetch(src)
  if (!response.ok) throw new Error(`画像を取得できません: ${src}`)
  const blob = await response.blob()
  return await new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(blob)
  })
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
 * Convert Marp's already-resolved 1280x720 local layout box into inches.
 * offset* values deliberately avoid viewport transforms applied by Marp's SVG
 * polyfill. The resulting box is frozen before any text is written to PPTX.
 */
function elementBox(element: HTMLElement, slide: HTMLElement) {
  const { left, top } = offsetWithin(element, slide)
  // Every Marp layer uses the same logical 1280x720 canvas. Advanced content
  // sections may themselves be only 58% wide; scaling that smaller width to a
  // full PPTX slide would enlarge and shift all text.
  const sx = SLIDE_W / 1280
  const sy = SLIDE_H / 720
  return {
    x: left * sx,
    y: top * sy,
    w: element.offsetWidth * sx,
    h: element.offsetHeight * sy,
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
  if (element.tagName === "PRE" && isVisibleColor(style.backgroundColor)) {
    shapes.push({ kind: "shape", shape: "rect", ...box, fill: cssColor(style.backgroundColor) })
  }
  if (element.tagName !== "PRE") {
    element.querySelectorAll<HTMLElement>("code").forEach((code) => {
      const codeStyle = getComputedStyle(code)
      if (!isVisibleColor(codeStyle.backgroundColor)) return
      Array.from(code.getClientRects()).forEach((rect) => {
        shapes.push({ kind: "shape", shape: "rect", ...clientRectBox(rect, slide), fill: cssColor(codeStyle.backgroundColor) })
      })
    })
  }
  return shapes
}

function backgroundImageUrl(value: string): string | null {
  const match = value.match(/^url\(["']?(.*?)["']?\)$/)
  return match?.[1] ?? null
}

async function advancedBackgroundImages(section: HTMLElement): Promise<ImageElement[]> {
  const images: ImageElement[] = []
  for (const figure of section.querySelectorAll<HTMLElement>("figure")) {
    const src = backgroundImageUrl(getComputedStyle(figure).backgroundImage)
    if (!src) continue
    images.push({
      kind: "image",
      ...elementBox(figure, section),
      src: await imageToDataUrl(src),
    })
  }
  return images
}

function mathTextElement(element: HTMLElement, slide: HTMLElement, latexQueue: { latex: string; display: boolean }[]): TextElement {
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
    const math = latexQueue.shift()
    if (math) mathRuns.push({ text: match[0], mathLatex: math.latex, mathDisplay: math.display })
    cursor = match.index! + match[0].length
  }
  if (cursor < plain.length) mathRuns.push({ text: plain.slice(cursor) })
  const converted = textElement(element, slide)
  converted.runs = applyResolvedFonts(mathRuns, getComputedStyle(element))
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

function textElement(element: HTMLElement, slide: HTMLElement): TextElement {
  const style = getComputedStyle(element)
  const box = elementBox(element, slide)
  if (element.tagName === "PRE") {
    const left = px(style.paddingLeft) / PX_PER_IN
    const right = px(style.paddingRight) / PX_PER_IN
    const top = px(style.paddingTop) / PX_PER_IN
    const bottom = px(style.paddingBottom) / PX_PER_IN
    box.x += left
    box.y += top
    box.w = Math.max(0.1, box.w - left - right)
    box.h = Math.max(0.1, box.h - top - bottom)
  }
  const fontPx = px(style.fontSize, 32)
  const lineHeightPx = style.lineHeight === "normal" ? fontPx * 1.2 : px(style.lineHeight, fontPx * 1.2)
  const valign = style.justifyContent === "center" ? "middle" : style.justifyContent === "flex-end" ? "bottom" : "top"
  return {
    kind: "text",
    ...box,
    runs: applyResolvedFonts(textRuns(element, true), style),
    fontSize: fontPx * 72 / PX_PER_IN,
    fontFace: weightedFontFace(resolvedFont(style.fontFamily), style.fontWeight),
    color: cssColor(style.color),
    bold: Number(style.fontWeight) >= 700 || style.fontWeight === "bold",
    align: style.textAlign === "center" ? "center" : style.textAlign === "right" || style.textAlign === "end" ? "right" : "left",
    valign,
    lineSpacingMultiple: lineHeightPx / fontPx,
    pre: element.tagName === "PRE",
  }
}

function listElement(element: HTMLOListElement | HTMLUListElement, slide: HTMLElement): ListElement {
  const style = getComputedStyle(element)
  const firstLi = element.querySelector(":scope > li")
  const itemStyle = firstLi ? getComputedStyle(firstLi) : style
  const fontPx = px(itemStyle.fontSize, 32)
  const lineHeight = itemStyle.lineHeight === "normal" ? fontPx * 1.2 : px(itemStyle.lineHeight, fontPx * 1.2)
  const box = elementBox(element, slide)
  const markerOffsetPx = Math.max(0, px(style.paddingLeft, 40) - fontPx)
  const markerOffsetIn = markerOffsetPx / PX_PER_IN
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
  const markerProperties = (li: HTMLLIElement, list: HTMLOListElement | HTMLUListElement) => {
    const pseudo = getComputedStyle(li, "::before")
    const marker = getComputedStyle(li, "::marker")
    const pseudoContent = pseudo.content && pseudo.content !== "none" ? pseudo.content : ""
    const markerStyle = pseudoContent ? pseudo : marker
    const quoted = pseudoContent.match(/^["'](.+)["']$/)?.[1]
    const counter = pseudoContent.match(/counter\([^,)]*(?:,\s*([^)]+))?\)(?:\s*["']([^"']*)["'])?/)
    const listStyle = counter?.[1]?.trim() || getComputedStyle(li).listStyleType || getComputedStyle(list).listStyleType
    const suffix = counter?.[2] || (pseudoContent.includes(")") ? ")" : ".")
    const fallbackCharacter = listStyle === "square" ? "▪" : listStyle === "circle" ? "◦" : "•"
    return {
      markerCharacter: list instanceof HTMLUListElement ? (quoted || fallbackCharacter) : undefined,
      markerColor: cssColor(markerStyle.color || getComputedStyle(li).color),
      markerFontSize: px(markerStyle.fontSize, px(getComputedStyle(li).fontSize, fontPx)) * 72 / PX_PER_IN,
      markerBold: Number(markerStyle.fontWeight) >= 600 || markerStyle.fontWeight === "bold",
      numberStyle: list instanceof HTMLOListElement ? numberStyle(listStyle, suffix) : undefined,
    }
  }
  const normalizedRuns = (li: HTMLLIElement, itemStyle: CSSStyleDeclaration) => {
    const runs = applyResolvedFonts(textRuns(li, false), itemStyle)
      .map((run) => ({ ...run, text: run.text.replace(/\s+/g, " ") }))
      .filter((run) => run.text.length > 0)
    if (!runs.length) return runs
    runs[0] = { ...runs[0], text: runs[0].text.trimStart() }
    runs[runs.length - 1] = { ...runs[runs.length - 1], text: runs[runs.length - 1].text.trimEnd() }
    return runs.filter((run) => run.text.length > 0)
  }
  const walk = (list: HTMLOListElement | HTMLUListElement, level: number) => {
    Array.from(list.children).filter((child): child is HTMLLIElement => child instanceof HTMLLIElement).forEach((li, index) => {
      const itemStyle = getComputedStyle(li)
      const clone = li.cloneNode(true) as HTMLLIElement
      clone.querySelectorAll("ul,ol").forEach((nested) => nested.remove())
      items.push({
        runs: normalizedRuns(clone, itemStyle),
        index: index + 1,
        level,
        ordered: list instanceof HTMLOListElement,
        fontSize: px(itemStyle.fontSize, fontPx) * 72 / PX_PER_IN,
        color: cssColor(itemStyle.color),
        ...markerProperties(li, list),
      })
      Array.from(li.children)
        .filter((child): child is HTMLOListElement | HTMLUListElement => child instanceof HTMLOListElement || child instanceof HTMLUListElement)
        .forEach((nested) => walk(nested, level + 1))
    })
  }
  walk(element, 0)
  return {
    kind: "list",
    ...box,
    x: box.x + markerOffsetIn,
    w: Math.max(0.1, box.w - markerOffsetIn),
    items,
    fontSize: fontPx * 72 / PX_PER_IN,
    fontFace: weightedFontFace(resolvedFont(itemStyle.fontFamily), itemStyle.fontWeight),
    color: cssColor(itemStyle.color),
    lineSpacingMultiple: lineHeight / fontPx,
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

async function captureBackground(slide: HTMLElement, editable: Element[]): Promise<string> {
  const previous = editable.map((element) => {
    const htmlElement = element as HTMLElement
    const keepHeadingDecoration = /^H[1-6]$/.test(htmlElement.tagName)
    const property = keepHeadingDecoration ? "color" : "visibility"
    return {
      element: htmlElement,
      property,
      value: htmlElement.style.getPropertyValue(property),
      priority: htmlElement.style.getPropertyPriority(property),
    }
  })
  editable.forEach((element) => {
    const htmlElement = element as HTMLElement
    if (/^H[1-6]$/.test(htmlElement.tagName)) htmlElement.style.setProperty("color", "transparent", "important")
    else htmlElement.style.setProperty("visibility", "hidden", "important")
  })
  try {
    return await toPng(slide, { pixelRatio: 2, cacheBust: true, skipFonts: false })
  } finally {
    previous.forEach(({ element, property, value, priority }) => {
      if (value) element.style.setProperty(property, value, priority)
      else element.style.removeProperty(property)
    })
  }
}

function topLevelEditable(slide: HTMLElement): Element[] {
  return Array.from(slide.querySelectorAll(EDITABLE_SELECTOR)).filter((element) => {
    if (element instanceof HTMLImageElement && isBackgroundImage(element)) return false
    if (element instanceof HTMLImageElement) return true
    const nativeImages = Array.from(element.querySelectorAll("img")).filter((img) => !isBackgroundImage(img))
    if (nativeImages.length > 0 && !(element.textContent ?? "").trim()) return false
    if (element.matches("li *")) return false
    if (element.matches("ul *, ol *") && !element.matches("ul,ol")) return false
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
  const marp = new Marp({ html: true })
  // Marpit's theme registry accepts only CSS containing `@theme` metadata.
  // Plain CSS is still useful as an override, so append it after Marp's
  // generated CSS without trying to register it as a named theme.
  if (/^\s*\/\*\s*@theme\s+[^*]+\*\//m.test(customThemeCss)) {
    marp.themeSet.add(customThemeCss)
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

      for (const element of editable) {
        try {
          if (element.querySelector("mjx-container, .katex")) {
            elements.push(mathTextElement(element as HTMLElement, section, mathQueue))
          } else if (element instanceof HTMLImageElement) {
            const box = elementBox(element, section)
            elements.push({ kind: "image", ...box, src: await imageToDataUrl(element.currentSrc || element.src) } satisfies ImageElement)
          } else if (element instanceof HTMLUListElement || element instanceof HTMLOListElement) {
            elements.push(listElement(element, section))
          } else if (element.tagName === "TABLE") {
            elements.push(tableElement(element as HTMLTableElement, section))
          } else {
            elements.push(...elementDecorations(element as HTMLElement, section))
            elements.push(textElement(element as HTMLElement, section))
          }
          nativeElements++
        } catch (error) {
          rasterElements++
          warnings.push({ slide: slideIndex + 1, message: error instanceof Error ? error.message : String(error) })
        }
      }

      const sectionStyle = getComputedStyle(section)
      const backgroundDataUrls: string[] = []
      let backgroundImages: ImageElement[] = []
      const allIndex = allSections.indexOf(section)
      const advancedBackground = allSections[allIndex - 1]?.dataset.marpitAdvancedBackground === "background"
        ? allSections[allIndex - 1]
        : null

      if (advancedBackground) {
        try {
          backgroundImages = await advancedBackgroundImages(advancedBackground)
          nativeElements += backgroundImages.length
        } catch (error) {
          warnings.push({ slide: slideIndex + 1, message: `高度背景画像の取得に失敗: ${error instanceof Error ? error.message : String(error)}` })
        }
      }

      try {
        // Capture only the content section's decorations. Marp's advanced
        // background and pseudo sections are support layers; capturing their
        // SVG wrappers produces duplicate or opaque full-slide images.
        backgroundDataUrls.push(await captureBackground(section, editable))
        rasterElements++
      } catch (error) {
        warnings.push({ slide: slideIndex + 1, message: `背景装飾の画像化に失敗: ${error instanceof Error ? error.message : String(error)}` })
      }
      const composedBackground = await composeBackground(
        cssColor(sectionStyle.backgroundColor, "#ffffff"),
        backgroundDataUrls,
        backgroundImages,
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
    browserController.cleanup()
    host.remove()
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
