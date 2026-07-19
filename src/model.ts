export const SLIDE_W = 13.333
export const SLIDE_H = 7.5

export interface TextRun {
  text: string
  bold?: boolean
  italic?: boolean
  code?: boolean
  fontFace?: string
  mathLatex?: string
  mathDisplay?: boolean
}

export interface TextElement {
  kind: "text"
  x: number
  y: number
  w: number
  h: number
  runs: TextRun[]
  fontSize: number
  fontFace: string
  color: string
  bold?: boolean
  align: "left" | "center" | "right"
  valign: "top" | "middle" | "bottom"
  lineSpacingMultiple: number
  pre?: boolean
}

export interface ListElement {
  kind: "list"
  x: number
  y: number
  w: number
  h: number
  items: { runs: TextRun[]; index: number }[]
  ordered: boolean
  fontSize: number
  fontFace: string
  color: string
  lineSpacingMultiple: number
  indent: number
}

export interface ImageElement {
  kind: "image"
  x: number
  y: number
  w: number
  h: number
  src: string
}

export interface TableCell {
  text: string
  fontSize: number
  fontFace: string
  color: string
  fill: string
  bold: boolean
  align: "left" | "center" | "right"
  valign: "top" | "middle" | "bottom"
  margin: [number, number, number, number]
  borderColor: string
  borderWidth: number
}

export interface TableElement {
  kind: "table"
  x: number
  y: number
  w: number
  h: number
  rows: TableCell[][]
  colWidths: number[]
  rowHeights: number[]
}

export type SlideElement = TextElement | ListElement | ImageElement | TableElement

export interface SlideModel {
  background: string
  backgroundDataUrls?: string[]
  backgroundImages?: ImageElement[]
  elements: SlideElement[]
  notes?: string
}

export interface ConversionWarning {
  slide: number
  message: string
}

export interface ConversionResult {
  slides: SlideModel[]
  title: string
  author: string
  warnings: ConversionWarning[]
  nativeElements: number
  rasterElements: number
}
