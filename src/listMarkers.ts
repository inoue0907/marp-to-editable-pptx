import JSZip from "jszip"

export interface ListMarkerPatch {
  slideIndex: number
  color: string
  fontSize: number
  bold: boolean
  fontFace: string
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
    let markerIndex = 0
    xml = xml.replace(/<a:pPr\b([^>]*)>([\s\S]*?(?:<a:buChar\b[^>]*\/>|<a:buAutoNum\b[^>]*\/>)[\s\S]*?)<\/a:pPr>/g, (paragraph, attrs, body) => {
      const patch = slidePatches[markerIndex++]
      if (!patch) return paragraph
      const cleaned = body
        .replace(/<a:buClr>[\s\S]*?<\/a:buClr>/g, "")
        .replace(/<a:buSz(?:Pct|Pts)\b[^>]*\/>/g, "")
      const markerXml = `<a:buClr><a:srgbClr val="${patch.color}"/></a:buClr><a:buSzPts val="${Math.round(patch.fontSize * 100)}"/>`
      const bulletStart = cleaned.search(/<a:bu(?:Font|Char|AutoNum)\b/)
      const styled = bulletStart >= 0
        ? `${cleaned.slice(0, bulletStart)}${markerXml}${cleaned.slice(bulletStart)}`
        : `${cleaned}${markerXml}`
      const markerRun = patch.bold
        ? `<a:r><a:rPr lang="en-US" sz="${Math.round(patch.fontSize * 100)}" b="1" dirty="0"><a:solidFill><a:srgbClr val="${patch.color}"/></a:solidFill><a:latin typeface="${patch.fontFace}"/><a:ea typeface="${patch.fontFace}"/><a:cs typeface="${patch.fontFace}"/></a:rPr><a:t>​</a:t></a:r>`
        : ""
      return `<a:pPr${attrs}>${styled}</a:pPr>${markerRun}`
    })
    zip.file(path, xml)
  }
  return await zip.generateAsync({ type: "arraybuffer" })
}
