// Exact 16:9 PowerPoint wide canvas (equivalent to 1920x1080).
export const SLIDE_W = 40 / 3
export const SLIDE_H = 7.5

export interface TextRun {
  text: string
  bold?: boolean
  italic?: boolean
  code?: boolean
  fontFace?: string
  fontSize?: number
  color?: string
  highlight?: string
  superscript?: boolean
  subscript?: boolean
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
  lineSpacing: number
  pre?: boolean
  fill?: string
  lineColor?: string
  lineWidth?: number
  margin?: [number, number, number, number]
  fit?: "shrink"
}

export interface ListElement {
  kind: "list"
  x: number
  y: number
  w: number
  h: number
  items: {
    runs: TextRun[]
    index: number
    level: number
    ordered: boolean
    fontSize: number
    color: string
    markerCharacter?: string
    markerColor: string
    markerFontSize: number
    markerBold: boolean
    numberStyle?: string
    markerIndent: number
    textIndent: number
    lineSpacing: number
    paraSpaceBefore: number
    paraSpaceAfter: number
  }[]
  fontSize: number
  fontFace: string
  color: string
  lineSpacingMultiple: number
  lineSpacing: number
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

export interface VideoElement {
  kind: "video"
  x: number
  y: number
  w: number
  h: number
  src: string
  extn: string
  cover?: string
}

export interface ShapeElement {
  kind: "shape"
  shape: "rect" | "line"
  x: number
  y: number
  w: number
  h: number
  fill?: string
  lineColor?: string
  lineWidth?: number
}

export interface TableCell {
  text: string
  runs: TextRun[]
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

export type SlideElement = TextElement | ListElement | ImageElement | VideoElement | TableElement | ShapeElement

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
