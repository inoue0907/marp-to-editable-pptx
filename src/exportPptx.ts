import pptxgen from "pptxgenjs"
import { ConversionResult, ListElement, SlideModel, TableElement, TextElement } from "./model"
import { injectNativeMath, MathPatch } from "./nativeMath"

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
      color: color(element.color),
      fontSize: run.code ? element.fontSize * 0.9 : element.fontSize,
    },
  }}), {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fontFace: element.fontFace, fontSize: element.fontSize,
    color: color(element.color), bold: element.bold,
    align: element.align, valign: element.valign,
    lineSpacingMultiple: element.lineSpacingMultiple / 1.12,
    margin: 0, breakLine: false, wrap: true,
  })
}

function addList(slide: pptxgen.Slide, element: ListElement) {
  const runs: pptxgen.TextProps[] = []
  element.items.forEach((item) => {
    item.runs.forEach((run, runIndex) => runs.push({
      text: run.text,
      options: {
        bold: run.bold, italic: run.italic,
        fontFace: run.fontFace ?? (run.code ? "Courier New" : element.fontFace),
        fontSize: run.code ? element.fontSize * 0.9 : element.fontSize,
        color: color(element.color),
        ...(runIndex === 0 ? { bullet: element.ordered
          ? { type: "number", style: "arabicPeriod", indent: element.indent * 72, numberStartAt: item.index }
          : { characterCode: "2022", indent: element.indent * 72 } } : {}),
      },
    }))
  })
  slide.addText(runs, {
    x: element.x, y: element.y, w: element.w, h: element.h,
    fontFace: element.fontFace, fontSize: element.fontSize,
    color: color(element.color), margin: 0, wrap: true,
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

function addSlide(pptx: pptxgen, model: SlideModel, slideIndex: number, mathPatches: MathPatch[]) {
  const slide = pptx.addSlide()
  slide.background = { color: color(model.background) }
  for (const background of model.backgroundDataUrls ?? []) {
    slide.addImage({ data: background, x: 0, y: 0, w: 13.333, h: 7.5 })
  }
  for (const image of model.backgroundImages ?? []) {
    slide.addImage({
      data: image.src,
      x: image.x, y: image.y, w: image.w, h: image.h,
      sizing: { type: "cover", w: image.w, h: image.h },
    })
  }
  for (const element of model.elements) {
    if (element.kind === "text") addText(slide, element, slideIndex, mathPatches)
    else if (element.kind === "list") addList(slide, element)
    else if (element.kind === "table") addTable(slide, element)
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
  result.slides.forEach((slide, index) => addSlide(pptx, slide, index, mathPatches))
  const buffer = await pptx.write({ outputType: "arraybuffer" }) as ArrayBuffer
  const blob = await injectNativeMath(buffer, mathPatches)
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = `${result.title.replace(/[\\/:*?"<>|]/g, "_")}.pptx`
  anchor.click()
  URL.revokeObjectURL(url)
}
