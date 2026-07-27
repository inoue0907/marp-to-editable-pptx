# Marp to Editable PowerPoint

[日本語](./README.ja.md)

Convert Marp Markdown into a visually faithful, editable PowerPoint presentation.

The converter lets Marp render the Markdown and theme CSS first. It then reads the final browser layout and computed styles, preserves complex decoration as the non-selectable slide background, and rebuilds the main content as native PowerPoint objects.

## Highlights

- Works with built-in and custom Marp themes
- Detects the theme used by the Markdown automatically
- Exports headings, paragraphs, inline styles, superscript, and subscript as editable text
- Exports bullets and multilevel numbering as native PowerPoint lists
- Preserves the resolved font size, color, line height, paragraph spacing, marker color, and marker size
- Exports ordinary images as independent PowerPoint images
- Embeds local HTML5 videos as native PowerPoint media
- Exports Markdown tables as native PowerPoint tables
- Exports inline and display math as editable PowerPoint equations (OMML)
- Exports code blocks as editable text boxes with their background fill and no outline
- Preserves `![bg]`, pseudo-elements, logos, decorative shapes, and theme artwork in the slide background
- Preserves Marp speaker notes

## Requirements

- Node.js LTS
- npm or pnpm
- Microsoft Edge or Google Chrome
- Fonts required by the Marp theme installed on the conversion machine

PowerPoint is not required for conversion, but is required to open and edit the generated `.pptx`.

## Installation

```bash
git clone https://github.com/inoue0907/marp-to-editable-pptx.git
cd marp-to-editable-pptx
npm install
npm link
```

`npm link` registers the short `marp-pptx` command on your PATH. This is required only once.

Verify the installation:

```bash
marp-pptx --help
```

## Usage

You can now run the converter from any directory. Open a terminal in your Marp project and run:

```bash
marp-pptx slides.md
```

The generated PowerPoint is written next to the Markdown file:

```text
slides.md
slides.pptx
```

Choose another output path:

```bash
marp-pptx slides.md --output dist/slides.pptx
```

You normally do not need to specify the theme. If necessary, it can be overridden explicitly:

```bash
marp-pptx slides.md --theme ./themes/my-theme.css
```

### CLI options

```text
marp-pptx <slides.md> [options]

-t, --theme <theme.css>    Override the Marp theme CSS
-o, --output <file.pptx>   Output path
--html-output <file.html>  Save the resolved, self-contained Marp HTML
--html-only                Save HTML without generating PowerPoint
--callout-style <style>    Render blockquotes as shape (default) or textbox
--headed                   Show the browser during conversion
-h, --help                 Show help
```

`--callout-style shape` creates a separate editable rectangle and border line
behind each blockquote. `--callout-style textbox` applies the background fill
directly to the editable text box. Neither mode bakes the callout into the
slide background image.

## Automatic theme detection

When `--theme` is omitted, the converter reads `theme:` from the Markdown front matter and resolves the CSS in this order:

1. Theme paths configured in the nearest `.vscode/settings.json` through `markdown.marp.themes`
2. `themes/<theme-name>.css` next to the Markdown file
3. `themes/<theme-name>.css` under the current working directory
4. `theme.css` next to the Markdown file

For VS Code settings, the converter walks upward from the Markdown file and uses the configured CSS whose `@theme` metadata matches the front matter.

Example:

```json
{
  "markdown.marp.themes": [
    "./themes/company.css",
    "./themes/technical.css"
  ]
}
```

```markdown
---
marp: true
theme: company
---
```

With this setup, the following command automatically uses `company.css`:

```bash
marp-pptx slides.md
```

Local images referenced by Markdown are resolved relative to the Markdown file. Assets referenced by theme CSS are resolved relative to the CSS file, with nearby parent directories also checked for community themes whose published paths assume the repository root.

## Using it from VS Code

Open your own Marp project in VS Code, open the integrated terminal, and run:

```bash
marp-pptx slides.md
```

The converter does not require the slide project to be inside this repository.

If you want `Ctrl+Shift+B` to export the currently active Markdown file, create `.vscode/tasks.json` in your slide project:

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Export active Marp file to PPTX",
      "type": "process",
      "command": "marp-pptx",
      "args": ["${file}"],
      "group": {
        "kind": "build",
        "isDefault": true
      },
      "problemMatcher": []
    }
  ]
}
```

After that, open the Markdown file and press `Ctrl+Shift+B`. VS Code passes the active file path to the globally installed command.

## Web interface

Start the development UI:

```bash
npm run dev
```

On Windows, you can also run:

```powershell
.\start-dev.ps1
```

If PowerShell blocks the script:

```powershell
powershell -ExecutionPolicy Bypass -File .\start-dev.ps1
```

Alternatively, double-click `start-dev.cmd`.

## How conversion works

1. Marp Core renders the Markdown and theme CSS.
2. Marp Browser resolves the final 1280×720 HTML layout.
3. The converter reads final DOM geometry and computed styles rather than interpreting arbitrary theme selectors.
4. Text, lists, images, tables, code, and math are removed temporarily to capture the remaining decoration.
5. The captured decoration becomes the non-selectable PowerPoint slide background.
6. Main content is recreated as native, editable PowerPoint objects at the measured positions.

This design is intended to work across differently authored themes without reimplementing each theme's CSS rules in PowerPoint.

## HTML export

Save the exact resolved Marp HTML used by the converter:

```bash
marp-pptx slides.md \
  --html-output ./dist/slides.html \
  --html-only
```

To generate both HTML and PowerPoint, omit `--html-only`.

## Fonts

The converter uses the font selected by the final computed Marp style. It does not embed font files into PowerPoint.

Install every required font on:

- the machine performing the conversion; and
- every machine that will open the PowerPoint without font substitution.

If a font is unavailable, the converter selects an installed fallback. Browser and PowerPoint font metrics may still produce small differences in wrapping.

## Native and background content

Editable PowerPoint content:

- Headings and paragraphs
- Bold, italic, colored, highlighted, superscript, and subscript runs
- Native bullet and numbered lists
- Ordinary images
- Local videos embedded with `<video src="...">`
- Tables
- Inline and display equations
- Code-block text and fill

Background content:

- `![bg]` and advanced Marp backgrounds
- Logos and theme artwork
- Pseudo-element decoration
- Decorative borders, underlines, shadows, and complex CSS effects
- Page furniture that is not intended for editing

## Current limitations

- Exact text wrapping can vary when browser and PowerPoint font metrics differ.
- Fonts must be installed; font embedding is not implemented.
- Remote fonts and images may fail when blocked by the network or the source server.
- PowerPoint media support depends on the video codec available on the viewing machine; MP4/H.264 is recommended.
- Complex CSS effects remain rasterized in the slide background.
- Code syntax-highlighting runs may not reproduce every highlighter-specific effect.
- Native PowerPoint does not support every CSS typography and layout feature.

## Privacy

Conversion runs locally through a temporary local web server and an installed browser. Files are not uploaded by this project.

Remote URLs referenced by Markdown or CSS may still be fetched from their original servers. For confidential presentations, use local assets and review the theme CSS for external URLs.

## Updating the CLI

```bash
cd marp-to-editable-pptx
git pull
npm install
npm link
```

Run `npm link` again if the command is no longer found after changing Node.js installations.
