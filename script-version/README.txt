Arabic PDF Tools 1.3.2 - PDF to PowerPoint / PDF to Word (editable, correct Arabic)
by Beedo El-Omda
============================================================================
1. Unzip this folder anywhere (e.g. Downloads).
2. Double-click Install.cmd and wait for "Done" (first time 1-3 minutes;
   it installs Python if needed, using Windows' built-in winget).
   Already installed the PowerPoint tool before? Just run Install.cmd again.
3. Use it: select one or more PDFs > right-click > Send to >
     "PDF to PPTX (Arabic, editable)"  -> "<name> (editable).pptx"
     "PDF to Word (Arabic, editable)"  -> "<name> (Word).docx"
   (Windows 11: right-click > Show more options > Send to).
   The new file appears next to the PDF; a message confirms when done.

PDF to Word:
- Books / reports / theses: rebuilt as flowing right-to-left Word text with
  the original fonts, sizes, footnote numbers, page numbers, pictures in
  place and one PDF page per Word page. Letters lost by Word's PDF export
  (kashida, ain, taa, harakat, lam-alef...) are recovered.
- Slide PDFs (landscape): kept in their layout; use the PowerPoint tool for
  the best result on slides.
- Requires the original fonts (e.g. Simplified Arabic) on the PC for the
  same look; they come with Windows / Office.

Why converters showed boxes: Mac-made PDFs map joined Arabic letters to rare
"ligature" codes (e.g. U+FC70 for تر, U+FCCC for لم) that Windows fonts lack.
These tools convert them back to normal letters.

Extra:
- fix_arabic_pdf.py repairs a PDF only ("<name> (fixed).pdf"); the fixed PDF
  can then be converted by PDFGear/Word/Acrobat without boxes.
- To remove: run Uninstall.cmd in %LOCALAPPDATA%\PDF2PPTX-Arabic

Enjoy , Beedo El-Omda
