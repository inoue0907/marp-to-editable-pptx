import pptxgen from "pptxgenjs"
import JSZip from "jszip"
import { ConversionResult, ListElement, ShapeElement, SlideModel, TableElement, TextElement, TextRun } from "./model"
import { injectNativeMath, MathPatch } from "./nativeMath"
import { injectListMarkerStyles, ListMarkerPatch } from "./listMarkers"

const color = (value: string) => value.replace("#", "").slice(0, 6).toUpperCase()

interface CodeIndentPatch {
  slideIndex: number
  marker: string
  indentPoints: number
}

async function injectCodeIndents(buffer: ArrayBuffer, patches: CodeIndentPatch[]): Promise<ArrayBuffer> {
  if (!patches.length) return buffer
  const zip = await JSZip.loadAsync(buffer)
  for (const patch of patches) {
    const path = `ppt/slides/slide${patch.slideIndex + 1}.xml`
    const file = zip.file(path)
    if (!file) continue
    let xml = await file.async("string")
    const paragraphPattern = new RegExp(`<a:p>(?:(?!<a:p>).)*?${patch.marker}(?:(?!<a:p>).)*?<\\/a:p>`, "s")
    xml = xml.replace(paragraphPattern, (paragraph) => {
      let cleaned = paragraph.replace(
        new RegExp(`<a:r>(?:(?!<a:r>).)*?<a:t>${patch.marker}<\\/a:t>(?:(?!<a:r>).)*?<\\/a:r>`, "s"),
        "",
      )
      const marL = Math.round(patch.indentPoints * 12700)
      if (/<a:pPr\b[^>]*\/>/.test(cleaned)) {
        cleaned = cleaned.replace(/<a:pPr\b([^>]*)\/>/, (_match, attrs) =>
          `<a:pPr${attrs.replace(/\s+marL="[^"]*"/g, "")} marL="${marL}"/>`)
      } else if (/<a:pPr\b/.test(cleaned)) {
        cleaned = cleaned.replace(/<a:pPr\b([^>]*)>/, (_match, attrs) =>
          `<a:pPr${attrs.replace(/\s+marL="[^"]*"/g, "")} marL="${marL}">`)
      } else {
        cleaned = cleaned.replace("<a:p>", `<a:p><a:pPr marL="${marL}"/>`)
      }
      // PptxGenJS repeats paragraph properties before rich-text runs. A
      // DrawingML paragraph may have only one pPr; retain the leading one.
      const firstPropertiesEnd = cleaned.indexOf("</a:pPr>") >= 0
        ? cleaned.indexOf("</a:pPr>") + "</a:pPr>".length
        : cleaned.indexOf("/>") + 2
      cleaned = cleaned.slice(0, firstPropertiesEnd) +
        cleaned.slice(firstPropertiesEnd).replace(/<a:pPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:pPr>)/g, "")
      return cleaned
    })
    zip.file(path, xml)
  }
  return await zip.generateAsync({ type: "arraybuffer" })
}

async function normalizeOpenXmlGeometry(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer)
  const slideNames = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  await Promise.all(slideNames.map(async (name) => {
    const entry = zip.file(name)
    if (!entry) return
    const xml = await entry.async("string")
    // DrawingML coordinates and extents are integer EMUs. PptxGenJS can emit
    // a decimal for sub-pixel browser geometry, which PowerPoint tolerates in
    // some cases but strict Open XML readers reject.
    zip.file(name, xml
      .replace(
        /\b(x|y|cx|cy)="(-?\d+(?:\.\d+)?)"/g,
        (_match, attribute, value) => `${attribute}="${Math.round(Number(value))}"`,
      )
      // PptxGenJS serializes `line: { type: "none" }` as an empty line
      // element. Make the intent explicit so PowerPoint never falls back to a
      // theme/default outline.
      .replace(/<a:ln><\/a:ln>/g, "<a:ln><a:noFill/></a:ln>"))
  }))
  return await zip.generateAsync({ type: "arraybuffer" })
}

function addText(
  slide: pptxgen.Slide,
  element: TextElement,
  slideIndex: number,
  mathPatches: MathPatch[],
  codeIndentPatches: CodeIndentPatch[],
) {
  let atLineStart = true
  const preparedRuns = element.runs.flatMap((run) => {
    const lines = run.text.split("\n")
    return lines.flatMap((rawText, index) => {
      const breakLine = index < lines.length - 1
      // split() emits a phantom final segment when the run ends at a newline.
      // Keep the line-start state for the following styled run.
      if (!rawText && index === lines.length - 1 && lines.length > 1) return []
      let text = rawText
      const output: (TextRun & { breakLine?: boolean })[] = []
      if (element.pre && atLineStart) {
        const leading = text.match(/^[ \t]+/)?.[0] ?? ""
        if (leading) {
          const spaces = Array.from(leading).reduce((sum, character) => sum + (character === "\t" ? 4 : 1), 0)
          text = text.slice(leading.length)
          const marker = `__MARPCODEINDENT_${slideIndex}_${codeIndentPatches.length}__`
          const fontSize = run.fontSize ?? element.fontSize
          codeIndentPatches.push({ slideIndex, marker, indentPoints: spaces * fontSize * 0.6 })
          output.push({ ...run, text: marker, fontSize: 1 })
        }
      }
      if (text || breakLine) output.push({
        ...run,
        text: text || " ",
        breakLine,
      })
      atLineStart = breakLine
      return output
    })
  })
  slide.addText(preparedRuns.map((run, runIndex) => {
    const marker = run.mathLatex ? `__OMML_${slideIndex}_${mathPatches.length}_${runIndex}__` : run.text
    if (run.mathLatex) mathPatches.push({
      slideIndex,
      marker,
      latex: run.mathLatex,
      display: Boolean(run.mathDisplay),
      color: color(element.color),
      fontSize: element.fontSize,
    })
    return {
    // PowerPoint trims leading ASCII spaces and may wrap code at spaces even
    // with wrapping disabled. Non-breaking spaces preserve the browser's
    // resolved preformatted lines and indentation in one editable text box.
    text: element.pre
      ? marker
          .replace(/\t/g, "\u00a0\u00a0\u00a0\u00a0")
          .replace(/ /g, "\u00a0")
      : marker.replace(/^ /, "\u00a0").replace(/ $/, "\u00a0"),
    options: {
      bold: run.bold ?? element.bold,
      italic: run.italic,
      fontFace: run.fontFace ?? (run.code ? "Courier New" : element.fontFace),
      color: color(run.color ?? element.color),
      fontSize: run.fontSize ?? (run.code ? element.fontSize * 0.9 : element.fontSize),
      highlight: run.highlight ? color(run.highlight) : undefined,
      superscript: run.superscript,
      subscript: run.subscript,
      breakLine: run.breakLine,
    },
  }}), {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fontFace: element.fontFace, fontSize: element.fontSize,
    color: color(element.color), bold: element.bold,
    align: element.align, valign: element.valign,
    lineSpacing: element.lineSpacing,
    margin: element.margin ?? 0, breakLine: false, wrap: !element.pre,
    fill: element.fill ? { color: color(element.fill) } : undefined,
    line: element.pre
      ? { type: "none" }
      : element.lineColor
        ? { color: color(element.lineColor), width: element.lineWidth ?? 1.5 }
        : undefined,
    fit: element.fit,
  })
}

function addList(
  slide: pptxgen.Slide,
  element: ListElement,
  slideIndex: number,
  mathPatches: MathPatch[],
  markerPatches: ListMarkerPatch[],
) {
  const runs: pptxgen.TextProps[] = []
  element.items.forEach((item) => {
    const marker = item.markerCharacter ?? "•"
    const patchMarker = `__MARPLIST_${slideIndex}_${markerPatches.length}__`
    markerPatches.push({
      slideIndex,
      marker: patchMarker,
      color: color(item.markerColor),
      fontSize: item.markerFontSize,
      bold: item.markerBold,
      fontFace: element.fontFace,
      ordered: item.ordered,
      characterCode: marker.codePointAt(0)?.toString(16).toUpperCase() ?? "2022",
      numberStyle: item.numberStyle,
      numberStartAt: item.index,
      markerIndentPoints: item.markerIndent * 72,
      textIndentPoints: item.textIndent * 72,
    })
    runs.push({
      text: patchMarker,
      options: {
        fontSize: 1,
        lineSpacing: item.lineSpacing,
        paraSpaceBefore: item.paraSpaceBefore,
        paraSpaceAfter: item.paraSpaceAfter,
      },
    })
    const itemRuns = item.runs.length ? item.runs : [{ text: " " }]
    itemRuns.forEach((run, runIndex) => {
      const text = run.mathLatex ? `__OMML_${slideIndex}_${mathPatches.length}_${runIndex}__` : run.text
      if (run.mathLatex) mathPatches.push({
        slideIndex,
        marker: text,
        latex: run.mathLatex,
        display: false,
        color: color(run.color ?? item.color),
        fontSize: run.fontSize ?? item.fontSize,
      })
      runs.push({
        text,
        options: {
          // Keep every body run explicit. PowerPoint's native auto-number
          // otherwise inherits bold from the first run of the paragraph.
          bold: run.bold ?? false, italic: run.italic,
          fontFace: run.fontFace ?? (run.code ? "Courier New" : element.fontFace),
          fontSize: run.fontSize ?? (run.code ? item.fontSize * 0.9 : item.fontSize),
          color: color(run.color ?? item.color),
          highlight: run.highlight ? color(run.highlight) : undefined,
          superscript: run.superscript,
          subscript: run.subscript,
          breakLine: runIndex === itemRuns.length - 1,
        },
      })
    })
  })
  slide.addText(runs, {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fontFace: element.fontFace, fontSize: element.fontSize,
    color: color(element.color), margin: 0, wrap: true, valign: "top",
    lineSpacing: element.lineSpacing,
  })
}

function addTable(
  slide: pptxgen.Slide,
  element: TableElement,
  slideIndex: number,
  mathPatches: MathPatch[],
) {
  slide.addTable(element.rows.map((row) => row.map((cell) => ({
    text: cell.runs.map((run, runIndex) => {
      const text = run.mathLatex ? `__OMML_${slideIndex}_${mathPatches.length}_${runIndex}__` : run.text
      if (run.mathLatex) mathPatches.push({
        slideIndex,
        marker: text,
        latex: run.mathLatex,
        display: false,
        color: color(run.color ?? cell.color),
        fontSize: run.fontSize ?? cell.fontSize,
      })
      return {
        text,
        options: {
          bold: run.bold,
          italic: run.italic,
          fontFace: run.fontFace ?? cell.fontFace,
          fontSize: run.fontSize ?? cell.fontSize,
          color: color(run.color ?? cell.color),
          superscript: run.superscript,
          subscript: run.subscript,
        },
      }
    }),
    options: {
      fontFace: cell.fontFace,
      fontSize: cell.fontSize,
      color: color(cell.color),
      fill: { color: color(cell.fill) },
      bold: cell.bold,
      align: cell.align,
      valign: cell.valign,
      margin: cell.margin,
      border: { type: "solid", color: color(cell.borderColor), pt: cell.borderWidth },
    },
  }))), {
    x: element.x, y: element.y, w: element.w, h: element.h,
    colW: element.colWidths,
    rowH: element.rowHeights,
  })
}

function addShape(slide: pptxgen.Slide, element: ShapeElement) {
  if (element.shape === "line") {
    slide.addShape("line" as pptxgen.ShapeType, {
      x: element.x, y: element.y, w: element.w, h: element.h,
      line: { color: color(element.lineColor ?? "#000000"), width: element.lineWidth ?? 1 },
    })
    return
  }
  slide.addShape("rect" as pptxgen.ShapeType, {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fill: { color: color(element.fill ?? "#FFFFFF") },
    line: element.lineColor
      ? { color: color(element.lineColor), width: element.lineWidth ?? 1 }
      : { color: color(element.fill ?? "#FFFFFF"), transparency: 100 },
  })
}

function addSlide(
  pptx: pptxgen,
  model: SlideModel,
  slideIndex: number,
  mathPatches: MathPatch[],
  markerPatches: ListMarkerPatch[],
  codeIndentPatches: CodeIndentPatch[],
) {
  const slide = pptx.addSlide()
  const backgroundImage = model.backgroundDataUrls?.[0]
  slide.background = backgroundImage
    ? { data: backgroundImage, path: "background.png" }
    : { color: color(model.background) }
  for (const element of model.elements) {
    const geometry = [element.x, element.y, element.w, element.h]
    if (!geometry.every(Number.isFinite)) continue
    // Marp Browser can retain a hidden auto-scaling source node far outside
    // the SVG, or expose it as a sub-pixel rectangle. It is not visible in the
    // resolved slide and may serialize as an invalid decimal EMU.
    if (element.kind !== "shape" && (
      element.w < 0.01 || element.h < 0.01 ||
      element.x + element.w <= 0 || element.y + element.h <= 0 ||
      element.x >= 13.333 || element.y >= 7.5
    )) continue
    if (element.kind === "text") addText(slide, element, slideIndex, mathPatches, codeIndentPatches)
    else if (element.kind === "list") addList(slide, element, slideIndex, mathPatches, markerPatches)
    else if (element.kind === "table") addTable(slide, element, slideIndex, mathPatches)
    else if (element.kind === "shape") addShape(slide, element)
    else if (element.kind === "video") slide.addMedia({
      type: "video",
      data: element.src,
      extn: element.extn,
      cover: element.cover,
      x: element.x, y: element.y, w: element.w, h: element.h,
    })
    else slide.addImage({ data: element.src, x: element.x, y: element.y, w: element.w, h: element.h })
  }
  if (model.notes) slide.addNotes(model.notes)
}

export async function exportEditablePptx(result: ConversionResult): Promise<void> {
  const pptx = new pptxgen()
  pptx.layout = "LAYOUT_WIDE"
  pptx.title = result.title
  pptx.author = result.author
  const mathPatches: MathPatch[] = []
  const markerPatches: ListMarkerPatch[] = []
  const codeIndentPatches: CodeIndentPatch[] = []
  result.slides.forEach((slide, index) =>
    addSlide(pptx, slide, index, mathPatches, markerPatches, codeIndentPatches))
  const buffer = await pptx.write({ outputType: "arraybuffer" }) as ArrayBuffer
  const normalizedBuffer = await normalizeOpenXmlGeometry(buffer)
  const styledBuffer = await injectListMarkerStyles(normalizedBuffer, markerPatches)
  const indentedBuffer = await injectCodeIndents(styledBuffer, codeIndentPatches)
  const blob = await injectNativeMath(indentedBuffer, mathPatches)
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = `${result.title.replace(/[\\/:*?"<>|]/g, "_")}.pptx`
  anchor.click()
  URL.revokeObjectURL(url)
}
