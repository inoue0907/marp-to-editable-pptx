import pptxgen from "pptxgenjs"
import { ConversionResult, ListElement, ShapeElement, SlideModel, TableElement, TextElement } from "./model"
import { injectNativeMath, MathPatch } from "./nativeMath"
import { injectListMarkerStyles, ListMarkerPatch } from "./listMarkers"

const color = (value: string) => value.replace("#", "").slice(0, 6).toUpperCase()

function addText(slide: pptxgen.Slide, element: TextElement, slideIndex: number, mathPatches: MathPatch[]) {
  slide.addText(element.runs.map((run, runIndex) => {
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
    text: marker,
    options: {
      bold: run.bold ?? element.bold,
      italic: run.italic,
      fontFace: run.fontFace ?? (run.code ? "Courier New" : element.fontFace),
      color: color(run.color ?? element.color),
      fontSize: run.fontSize ?? (run.code ? element.fontSize * 0.9 : element.fontSize),
    },
  }}), {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fontFace: element.fontFace, fontSize: element.fontSize,
    color: color(element.color), bold: element.bold,
    align: element.align, valign: element.valign,
    lineSpacingMultiple: element.lineSpacingMultiple / 1.12,
    margin: 0, breakLine: false, wrap: !element.pre,
  })
}

function addList(slide: pptxgen.Slide, element: ListElement, slideIndex: number, markerPatches: ListMarkerPatch[]) {
  const runs: pptxgen.TextProps[] = []
  element.items.forEach((item) => {
    markerPatches.push({
      slideIndex,
      color: color(item.markerColor),
      fontSize: item.markerFontSize,
      bold: item.markerBold,
      fontFace: element.fontFace,
    })
    item.runs.forEach((run, runIndex) => runs.push({
      text: run.text,
      options: {
        bold: run.bold, italic: run.italic,
        fontFace: run.fontFace ?? (run.code ? "Courier New" : element.fontFace),
        fontSize: run.fontSize ?? (run.code ? item.fontSize * 0.9 : item.fontSize),
        color: color(run.color ?? item.color),
        ...(runIndex === 0 ? {
          bullet: item.ordered
            ? { type: "number", style: item.numberStyle ?? "arabicPeriod", indent: element.indent * 72, numberStartAt: item.index }
            : { characterCode: (item.markerCharacter ?? "•").codePointAt(0)!.toString(16).padStart(4, "0"), indent: element.indent * 72 },
          indentLevel: item.level,
        } : {}),
      },
    }))
  })
  slide.addText(runs, {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fontFace: element.fontFace, fontSize: element.fontSize,
    color: color(element.color), margin: 0, wrap: true, valign: "top",
    lineSpacingMultiple: element.lineSpacingMultiple / 1.12,
  })
}

function addTable(slide: pptxgen.Slide, element: TableElement) {
  slide.addTable(element.rows.map((row) => row.map((cell) => ({
    text: cell.text,
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

function addSlide(pptx: pptxgen, model: SlideModel, slideIndex: number, mathPatches: MathPatch[], markerPatches: ListMarkerPatch[]) {
  const slide = pptx.addSlide()
  const backgroundImage = model.backgroundDataUrls?.[0]
  slide.background = backgroundImage
    ? { data: backgroundImage, path: "background.png" }
    : { color: color(model.background) }
  for (const element of model.elements) {
    if (element.kind === "text") addText(slide, element, slideIndex, mathPatches)
    else if (element.kind === "list") addList(slide, element, slideIndex, markerPatches)
    else if (element.kind === "table") addTable(slide, element)
    else if (element.kind === "shape") addShape(slide, element)
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
  result.slides.forEach((slide, index) => addSlide(pptx, slide, index, mathPatches, markerPatches))
  const buffer = await pptx.write({ outputType: "arraybuffer" }) as ArrayBuffer
  const styledBuffer = await injectListMarkerStyles(buffer, markerPatches)
  const blob = await injectNativeMath(styledBuffer, mathPatches)
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = `${result.title.replace(/[\\/:*?"<>|]/g, "_")}.pptx`
  anchor.click()
  URL.revokeObjectURL(url)
}
