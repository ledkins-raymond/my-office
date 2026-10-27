# My Office

Personal replacements for Word, Excel and PowerPoint, built for Windows. Every
file is saved in the standard Office format (.docx, .xlsx, .pptx), so it opens
in Microsoft Office, Google Docs, LibreOffice, Apple Pages and phone apps.

Only **My Word** exists so far. Excel and PowerPoint come next.

## Installing on Windows

1. Open the repository's **Actions** tab on GitHub, click the latest green
   "Build Windows installer" run, and download **My-Office-Setup** from the
   Artifacts section. (For tagged versions, the installer is also on the
   **Releases** page.)
2. Unzip it and run `My Office Setup x.y.z.exe`.
3. Windows SmartScreen warns that the app is from an unknown publisher,
   because it isn't code-signed. Click **More info**, then **Run anyway**.
4. Start it from the **My Word** shortcut.

To open a .docx with it, right-click the file, choose **Open with**, and pick
My Office.

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

## Developing

Requires Node.js 22.

```
npm install
npm start        # build and run the app
npm test         # .docx save/open tests
npm run dist     # build the Windows installer (run on Windows)
```

Layout:

- `electron/main.cjs` – windows, menus, file dialogs, reading and writing files
- `electron/docx-import.cjs` – reads .docx into the editor
- `src/word/docx-export.js` – writes the editor's content as .docx
- `src/word/main.js` – the editor (built on TipTap/ProseMirror) and toolbar
