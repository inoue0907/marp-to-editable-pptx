import { Marp } from "@marp-team/marp-core"

const markdown = `---
marp: true
theme: default
---
# One
---
# Two
---
![bg right:42%](/code-art.svg)
## Three
Body
`

const { html } = new Marp({ html: true }).render(markdown)
const sections = [...html.matchAll(/<section\b([^>]*)>([\s\S]*?)<\/section>/g)]
console.log(`sections=${sections.length}`)
sections.forEach((match, index) => {
  const attrs = match[1].replace(/\s+/g, " ").trim()
  const text = match[2].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 100)
  console.log(`${index + 1}: attrs=[${attrs}] text=[${text}]`)
  if (attrs.includes('advanced-background="background"')) {
    console.log(match[2])
  }
})
