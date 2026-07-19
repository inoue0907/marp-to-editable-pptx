import JSZip from "jszip"
import katex from "katex"
// mathml2omml ships ESM and is bundled by Vite.
import { mml2omml } from "mathml2omml"

export interface MathPatch {
  slideIndex: number
  marker: string
  latex: string
  display: boolean
  color: string
  fontSize: number
}

const NS_MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

const escapeXml = (value: string) => value
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;")

function plainOmml(latex: string): string {
  return `<m:oMath><m:r><m:t>${escapeXml(latex)}</m:t></m:r></m:oMath>`
}

function groupNaryBodies(mathml: string): string {
  try {
    const doc = new DOMParser().parseFromString(mathml, "application/xml")
    if (doc.querySelector("parsererror")) return mathml
    const tags = new Set(["msubsup", "munderover", "msub", "munder"])
    const chars = new Set(["∫", "∬", "∭", "∮", "∑", "∏", "⋂", "⋃"])
    const process = (element: Element): void => {
      const children = Array.from(element.children)
      for (let index = 0; index < children.length; index++) {
        const child = children[index]
        const base = child.children[0]
        if (tags.has(child.localName) && base?.localName === "mo" && chars.has(base.textContent?.trim() ?? "")) {
          const body = children.slice(index + 1)
          if (body.length > 1 || (body.length === 1 && body[0].localName !== "mrow")) {
            const wrapper = doc.createElementNS("http://www.w3.org/1998/Math/MathML", "mrow")
            body.forEach((node) => wrapper.appendChild(node))
            element.appendChild(wrapper)
          }
          return
        }
        process(child)
      }
    }
    process(doc.documentElement)
    return new XMLSerializer().serializeToString(doc.documentElement)
  } catch {
    return mathml
  }
}

function sanitizeOmml(omml: string): string {
  try {
    const doc = new DOMParser().parseFromString(`<root xmlns:m="${NS_MATH}" xmlns:w="${NS_W}">${omml}</root>`, "text/xml")
    if (doc.querySelector("parsererror")) return omml
    const root = doc.documentElement
    for (const run of Array.from(root.getElementsByTagNameNS(NS_MATH, "r"))) {
      const text = run.getElementsByTagNameNS(NS_MATH, "t")[0]
      if (text && !text.textContent) run.remove()
    }
    const argumentTags = ["e", "num", "den", "sub", "sup", "deg", "lim", "oMath", "fName"]
    for (const tag of argumentTags) {
      for (const element of Array.from(root.getElementsByTagNameNS(NS_MATH, tag))) {
        for (const child of Array.from(element.childNodes)) {
          if (child.nodeType === Node.TEXT_NODE && child.textContent?.trim()) {
            const run = doc.createElementNS(NS_MATH, "m:r")
            const text = doc.createElementNS(NS_MATH, "m:t")
            text.textContent = child.textContent
            run.appendChild(text)
            element.replaceChild(run, child)
          }
        }
      }
    }
    return new XMLSerializer().serializeToString(root)
      .replace(/^<root[^>]*>/, "").replace(/<\/root>$/, "")
      .replace(/ xmlns:[a-zA-Z0-9]+="[^"]*"/g, "")
  } catch {
    return omml
  }
}

function latexToOmml(latex: string): string {
  const html = katex.renderToString(latex, { output: "mathml", throwOnError: false })
  const math = html.match(/<math[\s\S]*?<\/math>/)?.[0]
  if (!math) return plainOmml(latex)
  try {
    return sanitizeOmml((mml2omml(groupNaryBodies(math)) as string)
      .replace(/ xmlns:[a-zA-Z0-9]+="[^"]*"/g, ""))
  } catch {
    return plainOmml(latex)
  }
}

function paragraphBounds(xml: string, markerIndex: number): [number, number] | null {
  const plain = xml.lastIndexOf("<a:p>", markerIndex)
  const attributed = xml.lastIndexOf("<a:p ", markerIndex)
  const start = Math.max(plain, attributed)
  const end = xml.indexOf("</a:p>", markerIndex)
  return start < 0 || end < 0 ? null : [start, end + 6]
}

function replaceDisplay(xml: string, patch: MathPatch, omml: string): string {
  const markerIndex = xml.indexOf(patch.marker)
  if (markerIndex < 0) return xml
  const bounds = paragraphBounds(xml, markerIndex)
  if (!bounds) return xml
  const size = Math.round(patch.fontSize * 100)
  const paragraph = `<a:p><a:pPr><a:defRPr sz="${size}"><a:solidFill><a:srgbClr val="${patch.color}"/></a:solidFill></a:defRPr></a:pPr><a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>${omml}</m:oMathPara></a14:m></a:p>`
  return xml.slice(0, bounds[0]) + paragraph + xml.slice(bounds[1])
}

function replaceInline(xml: string, patch: MathPatch, omml: string): string {
  const markerIndex = xml.indexOf(patch.marker)
  if (markerIndex < 0) return xml
  const textOpen = xml.lastIndexOf("<a:t", markerIndex)
  const textStart = xml.indexOf(">", textOpen) + 1
  const textEnd = xml.indexOf("</a:t>", markerIndex)
  const runOpen = xml.lastIndexOf("<a:r>", textOpen)
  const runEnd = xml.indexOf("</a:r>", textEnd) + 6
  if (textOpen < 0 || textEnd < 0 || runOpen < 0 || runEnd < 6) return xml
  const text = xml.slice(textStart, textEnd)
  const markerAt = text.indexOf(patch.marker)
  const before = text.slice(0, markerAt)
  const after = text.slice(markerAt + patch.marker.length)
  const runXml = xml.slice(runOpen, runEnd)
  const properties = runXml.match(/(<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>))/)?.[1] ?? "<a:rPr/>"
  const replacement = `${before ? `<a:r>${properties}<a:t xml:space="preserve">${before}</a:t></a:r>` : ""}<a14:m>${omml}</a14:m>${after ? `<a:r>${properties}<a:t xml:space="preserve">${after}</a:t></a:r>` : ""}`
  let result = xml.slice(0, runOpen) + replacement + xml.slice(runEnd)

  // PowerPoint renders inline OMML smaller than adjacent DrawingML text unless
  // the paragraph default run size is stated explicitly. The 1.21 correction
  // matches the visual em-size used by Marp/MathJax.
  const paragraph = paragraphBounds(result, runOpen)
  if (paragraph) {
    const paragraphXml = result.slice(paragraph[0], paragraph[1])
    const size = Math.round(patch.fontSize * 1.21 * 100)
    const defaultRun = `<a:defRPr sz="${size}"><a:solidFill><a:srgbClr val="${patch.color}"/></a:solidFill></a:defRPr>`
    let patchedParagraph: string
    const propertiesEnd = paragraphXml.indexOf("</a:pPr>")
    if (propertiesEnd >= 0) {
      patchedParagraph = paragraphXml.includes("<a:defRPr")
        ? paragraphXml
        : paragraphXml.slice(0, propertiesEnd) + defaultRun + paragraphXml.slice(propertiesEnd)
    } else {
      const openingEnd = paragraphXml.indexOf(">") + 1
      patchedParagraph = paragraphXml.slice(0, openingEnd) + `<a:pPr>${defaultRun}</a:pPr>` + paragraphXml.slice(openingEnd)
    }
    result = result.slice(0, paragraph[0]) + patchedParagraph + result.slice(paragraph[1])
  }
  return result
}

function ensureNamespaces(xml: string): string {
  const namespaces: string[] = []
  if (!xml.includes("xmlns:m=")) namespaces.push(`xmlns:m="${NS_MATH}"`)
  if (!xml.includes("xmlns:w=")) namespaces.push(`xmlns:w="${NS_W}"`)
  if (!xml.includes("xmlns:a14=")) namespaces.push('xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"')
  return namespaces.length ? xml.replace("<p:sld ", `<p:sld ${namespaces.join(" ")} `) : xml
}

export async function injectNativeMath(pptxBuffer: ArrayBuffer, patches: MathPatch[]): Promise<Blob> {
  const zip = await JSZip.loadAsync(pptxBuffer)
  const converted = new Map<string, string>()
  for (const patch of patches) {
    if (!converted.has(patch.latex)) converted.set(patch.latex, latexToOmml(patch.latex))
  }
  for (const slidePath of Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))) {
    const slideNumber = Number(slidePath.match(/slide(\d+)\.xml$/)?.[1]) - 1
    const relevant = patches.filter((patch) => patch.slideIndex === slideNumber)
    if (!relevant.length) continue
    let xml = await zip.files[slidePath].async("string")
    for (const patch of relevant) {
      const omml = converted.get(patch.latex) ?? plainOmml(patch.latex)
      xml = patch.display ? replaceDisplay(xml, patch, omml) : replaceInline(xml, patch, omml)
    }
    zip.file(slidePath, ensureNamespaces(xml))
  }
  return zip.generateAsync({ type: "blob" })
}
