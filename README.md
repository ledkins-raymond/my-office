# My Office

Personal replacements for Word, Excel and PowerPoint, built for Windows. Every
file is saved in the standard Office format (.docx, .xlsx, .pptx), so it opens
in Microsoft Office, Google Docs, LibreOffice, Apple Pages and phone apps.

**My Word** and **My Excel** exist so far. PowerPoint comes next. Both are in
one program and one installer.

## Installing on Windows

1. Open the repository's **Actions** tab on GitHub, click the latest green
   "Build Windows installer" run, and download **My-Office-Setup** from the
   Artifacts section. (For tagged versions, the installer is also on the
   **Releases** page.)
2. Unzip it and run `My Office Setup x.y.z.exe`.
3. Windows SmartScreen warns that the app is from an unknown publisher,
   because it isn't code-signed. Click **More info**, then **Run anyway**.
4. Start it from the **My Word** or **My Excel** shortcut (Start menu and
   desktop).

To open a .docx or .xlsx with it, right-click the file, choose **Open with**,
and pick My Office. From either program, **File > Open** opens both kinds and
starts the right one.

## What My Word does

- New, Open, Save, Save As (.docx), Export to PDF, Print
- Styles (Normal, Heading 1 to 3), fonts, sizes, colour, highlight
- Bold, italic, underline, strikethrough, clear formatting
- Alignment: left, centre, right, justify
- Bulleted and numbered lists, with nesting (Tab and Shift+Tab)
- Tables: insert, add or delete rows and columns, merge cells, header row
- Pictures: insert, paste, drag in, resize by the corners
- Links, page breaks (Ctrl+Enter), horizontal lines
- Find and replace (Ctrl+F)
- Word count, spell check, unsaved-changes prompt
- Page is US Letter with 1-inch margins

## Known limits

- The page is one long sheet on screen. Real page breaks appear in print, in
  PDF, and in other programs.
- Opening a complicated .docx from someone else keeps the text, headings,
  lists, tables, pictures, links and character formatting. It drops headers,
  footers, text boxes, footnotes, comments, columns, line spacing and
  paragraph indents.
- No track changes, comments, mail merge or templates.

## What My Excel does

- New, Open, Save, Save As (.xlsx), Export to PDF, Print
- Formulas with about 400 Excel functions (SUM, IF, VLOOKUP, XLOOKUP, SUMIFS,
  INDEX/MATCH, TEXT, dates, financial and statistical functions…), calculated
  by the HyperFormula engine. References to other sheets and named ranges work.
- Point at cells while typing a formula (click or arrow keys), function name
  suggestions (Tab to accept), coloured outlines on referenced cells
- Typing a date, money ($12.50) or a percentage picks the matching number
  format, as in Excel
- Number formats: General, Number, Currency, Accounting, dates, time,
  percentage, fraction, scientific, text, or any custom Excel format code;
  increase and decrease decimals
- Fonts, sizes, bold, italic, underline, strikethrough, font and fill colour,
  alignment, wrap text, borders
- Several sheets: add, rename (double-click the tab), delete, drag to reorder
- Insert and delete rows and columns (formulas adjust), resize by dragging
  the header edges, double-click an edge to fit, hide and unhide
- Copy, cut and paste (formulas adjust; pasting from other programs works),
  fill handle and Ctrl+D / Ctrl+R (number series continue), AutoSum (Alt+=),
  sort A to Z / Z to A (keeps a header row on top), find and replace
- Pivot tables: click in a block of data, choose **PivotTable**, pick the
  row field(s), an optional column field and up to two values (Sum, Count or
  Average). The table goes on a new sheet with grand totals. **Refresh**
  rebuilds it to take in new rows and new categories.
- Freeze panes (top row, first column, or at the selected cell)
- Sum, average and count of the selection in the status bar
- Undo and redo for everything

Saved files store each formula's last result too, so phone apps and file
previews that don't calculate still show the numbers. Every save is checked
by the tests against LibreOffice.

## Known limits (My Excel)

- Pivot tables are built from ordinary formulas (SUMIFS, COUNTIFS), not
  Excel PivotTable objects. Their numbers update everywhere, including in
  Excel and Google Sheets, but in real Excel they are a plain table: no field
  list, and new categories appear only after Refresh in My Excel. The
  definitions are kept in a hidden sheet so Refresh still works after you
  reopen the file. Pivot tables made in real Excel open as plain values.
- No charts, pictures, merged cells, conditional formatting, data validation,
  comments, filters, or macros. Opening a file that has them
  keeps the data and formulas but drops those parts when you save.
- A few Excel functions are missing from the engine (for example AVERAGEIFS, the CUBE
  functions and some newer dynamic-array ones). A cell using one keeps the
  result it had when the file was last saved in Excel, but won't recalculate.
- Printing prints every sheet's used cells as a table, landscape, with no page
  setup options yet.

## Developing

Requires Node.js 22.

```
npm install
npm start        # build and run My Word
npm run start:excel  # build and run My Excel
npm test         # .docx and .xlsx save/open tests (the LibreOffice check runs if it's installed)
npm run dist     # build the Windows installer (run on Windows)
```

Layout:

- `electron/main.cjs` – windows, menus, file dialogs, reading and writing files
- `electron/docx-import.cjs` – reads .docx into the editor
- `src/word/docx-export.js` – writes the editor's content as .docx
- `src/word/main.js` – the editor (built on TipTap/ProseMirror) and toolbar
- `electron/xlsx-io.cjs` – reads and writes .xlsx (with ExcelJS)
- `src/excel/workbook.js` – cell data, formulas (HyperFormula), formatting, undo
- `src/excel/grid.js` – the spreadsheet grid, drawn on a canvas
- `src/excel/main.js` – ribbon, formula bar, sheet tabs, clipboard, find
- `src/excel/pivot.js` – pivot tables
- `build/installer.nsh` – adds the My Excel shortcut to the installer

HyperFormula is used under its GPLv3 licence, which is fine for a program you
build for yourself. If you ever sell or hand out My Office as closed-source,
that licence would need another look.
