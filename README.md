# PDF Customiser

An Obsidian plugin that lets you fully style your PDF exports, and also export notes to Word (`.docx`) and PowerPoint (`.pptx`).

## Install

Copy `main.js`, `manifest.json` and `styles.css` into `<your vault>/.obsidian/plugins/pdf-customiser/`, then enable **PDF Customiser** under *Settings → Community plugins*.

## Install with BRAT (auto-updating)

Install the BRAT plugin, run **BRAT: Add a beta plugin for testing**, and enter `mkdogs25/obsidian-exporter`. Every push to the default branch that changes the plugin files publishes a new release automatically (see `.github/workflows/release.yml`), and BRAT picks it up.

## Using it

Open the **⋯ menu** at the top of a note. It has three buttons:

- **Export to PDF** opens Obsidian's own *Export to PDF* dialog with your custom style applied. The dialog also gains a **PDF Customiser** section where you can switch designs and toggle the header/footer, page numbers and watermark. Keep *Margin* on **Default** so the plugin's margins apply.
- **Export to Word (.docx)** creates an editable Word document with real Heading 1–6 styles, lists, tables, images, links, a header and footer, and an optional table of contents.
- **Export to PowerPoint (.pptx)** turns the note into slides. Put a horizontal rule (`---`) on its own line to start a new slide. The first heading on each slide becomes its title.

The same actions are available as commands, along with **Choose export design…** and **Toggle custom PDF style**.

## Settings

The settings are split into tabs. Most tabs include a live A4 preview, which you can expand to a larger view.

| Tab | What's inside |
| --- | --- |
| General | Turn the style on or off, choose which menu buttons appear, set where files are saved and their names, default author, date format, and embedded notes |
| Designs | A gallery of 12 designs with category filters and undo, plus fine-tuning for heading decoration, quotes, tables, images, page border, accent bar, background pattern, drop cap, link underline and corner roundness |
| Colours | Separate colours for H1–H6, body text, background, links, highlights, quotes, code, tables, rules and the header/footer, plus "generate palette from accent" |
| Typography | Separate font, size, weight, colour, italic, uppercase, letter spacing and alignment for each heading level; body and code fonts; quick tools to use one font for all headings and to set heading sizes on a scale |
| Page & layout | Margins, starting each H1 on a new page, keeping blocks together, image width, printing link addresses, and a watermark |
| Header & footer | Left, centre and right slots, a separator line, the first-page option and page numbers (position, format, numbering style), with clickable tokens |
| Presets | Save, apply, overwrite, delete, import and export (JSON) your own styles |
| Word / PowerPoint | Options specific to each format |
| Advanced | Custom CSS (scoped to the note) and a button to copy the generated CSS |

Tokens: `{title}` `{filename}` `{folder}` `{date}` `{time}` `{author}` `{vault}` `{page}` `{pages}` `{fm.anyFrontmatterKey}`.

Fonts must be installed on your computer to appear in exports.
