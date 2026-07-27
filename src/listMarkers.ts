import JSZip from "jszip"

export interface ListMarkerPatch {
  slideIndex: number
  marker: string
  color: string
  fontSize: number
  bold: boolean
  fontFace: string
  ordered: boolean
  characterCode?: string
  numberStyle?: string
  numberStartAt?: number
  markerIndentPoints: number
  textIndentPoints: number
}

export async function injectListMarkerStyles(buffer: ArrayBuffer, patches: ListMarkerPatch[]): Promise<ArrayBuffer> {
  if (!patches.length) return buffer
  const zip = await JSZip.loadAsync(buffer)
  const bySlide = new Map<number, ListMarkerPatch[]>()
  patches.forEach((patch) => bySlide.set(patch.slideIndex, [...(bySlide.get(patch.slideIndex) ?? []), patch]))
  for (const [slideIndex, slidePatches] of bySlide) {
    const path = `ppt/slides/slide${slideIndex + 1}.xml`
    const file = zip.file(path)
    if (!file) continue
    let xml = await file.async("string")
    for (const patch of slidePatches) {
      const paragraphPattern = new RegExp(`<a:p>(?:(?!<a:p>).)*?${patch.marker}(?:(?!<a:p>).)*?<\\/a:p>`, "s")
      xml = xml.replace(paragraphPattern, (paragraph) => {
        const withoutMarkerRun = paragraph.replace(
          new RegExp(`<a:r>(?:(?!<a:r>).)*?<a:t>${patch.marker}<\\/a:t>(?:(?!<a:r>).)*?<\\/a:r>`, "s"),
          // PowerPoint's automatic marker inherits weight from the first
          // actual run, even when defRPr says otherwise. Retain a zero-width
          // leading run with the marker's own weight so a bold first word does
          // not make the list number/bullet bold.
          `<a:r><a:rPr lang="en-US" sz="100" b="${patch.bold ? 1 : 0}"/><a:t>&#x200B;</a:t></a:r>`,
        )
        const bulletXml = patch.ordered
          ? `<a:buAutoNum type="${patch.numberStyle ?? "arabicPeriod"}" startAt="${patch.numberStartAt ?? 1}"/>`
          : `<a:buChar char="&#x${patch.characterCode ?? "2022"};"/>`
        const markerXml = `<a:buClr><a:srgbClr val="${patch.color}"/></a:buClr><a:buSzPts val="${Math.round(patch.fontSize * 100)}"/><a:buFont typeface="${patch.fontFace.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"/>${bulletXml}`
        const patchedParagraph = withoutMarkerRun.replace(/<a:pPr\b([^>]*)>([\s\S]*?)<\/a:pPr>/, (_pPr, attrs, body) => {
          const cleaned = body
            .replace(/<a:buNone\/>/g, "")
            .replace(/<a:buClr>[\s\S]*?<\/a:buClr>/g, "")
            .replace(/<a:buSz(?:Pct|Pts)\b[^>]*\/>/g, "")
            .replace(/<a:buFont\b[^>]*\/>/g, "")
            .replace(/<a:bu(?:Char|AutoNum)\b[^>]*\/>/g, "")
            .replace(/<a:defRPr\b[^>]*\/>/g, "")
            .replace(/<a:defRPr\b[^>]*>[\s\S]*?<\/a:defRPr>/g, "")
          const marL = Math.round(patch.textIndentPoints * 12700)
          const hanging = Math.round((patch.markerIndentPoints - patch.textIndentPoints) * 12700)
          const cleanAttrs = attrs
            .replace(/\s+marL="[^"]*"/g, "")
            .replace(/\s+indent="[^"]*"/g, "")
          const markerDefaults = `<a:defRPr b="${patch.bold ? 1 : 0}"/>`
          return `<a:pPr${cleanAttrs} marL="${marL}" indent="${hanging}">${cleaned}${markerXml}${markerDefaults}</a:pPr>`
        })
        // Rich-text arrays emitted by PptxGenJS can carry a repeated pPr
        // before each run. A native list paragraph may have only one pPr, so
        // retain the patched leading one and remove the redundant copies.
        const firstPropertiesEnd = patchedParagraph.indexOf("</a:pPr>") + "</a:pPr>".length
        return patchedParagraph.slice(0, firstPropertiesEnd) +
          patchedParagraph.slice(firstPropertiesEnd).replace(/<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g, "")
      })
    }
    zip.file(path, xml)
  }
  return await zip.generateAsync({ type: "arraybuffer" })
}
