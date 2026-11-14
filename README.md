# Arabic PDF Tools

Convert PDFs into **editable PowerPoint and Word files with correct Arabic text** — no more boxes (□) instead of letters, broken words or reversed letters.

By **Beedo El-Omda** · Windows · free and open source ([MIT](LICENSE))

[Download](https://github.com/BeedoElOmda/Arabic-PDF-Tools/releases/latest) · [Code signing policy](#code-signing-policy) · [Privacy](#privacy) · [Uninstall](#uninstall)

---

## What it does

Many PDFs — especially ones exported from Microsoft Word or made on a Mac — store Arabic text in a way that converters (PDFGear, Word, Acrobat) can't read back: joined letters become boxes, kashida turns into extra letters, harakat separate from their letters, and spaces appear inside words. Arabic PDF Tools repairs the text first, then converts:

- **PDF → PowerPoint (editable):** each slide keeps the original look; every shape becomes one editable right-to-left text box.
- **PDF → Word (editable):** books and reports become flowing right-to-left Word text with the original fonts, sizes, footnotes, page numbers, pictures and tables (one PDF page per Word page). Slide PDFs keep their layout.
- **Repair Arabic in PDF:** writes a fixed copy of the PDF that any other converter can then read correctly.

It recovers letters the PDF lost (from the embedded fonts' own data and the document's vocabulary), reorders numbers, brackets and English words inside Arabic, and removes kashida and stray spaces.

## Download and use

From [Releases](https://github.com/BeedoElOmda/Arabic-PDF-Tools/releases/latest), pick one:

| Package | What it is |
|---|---|
| `ArabicPDFTools-Setup-<version>.exe` | Installer for your user account (no admin rights). Adds a Start-menu entry and right-click menu items on PDFs. |
| `ArabicPDFTools-Portable-<version>.zip` | No installation: unzip anywhere and run `ArabicPDFTools.exe`, or drag PDFs onto it. |
| `Arabic-PDF-Tools-Script-<version>.zip` | Script version: run `Install.cmd` (installs Python if needed; needs internet). |

All three are offline once installed and do the same conversions. The new file is saved next to each PDF.

## Code signing policy

Free code signing provided by [SignPath.io](https://about.signpath.io), certificate by [SignPath Foundation](https://signpath.org).

- **Committers and reviewers:** [Beedo El-Omda](https://github.com/BeedoElOmda)
- **Approvers:** [Beedo El-Omda](https://github.com/BeedoElOmda)

Only files built from this repository by its GitHub Actions workflow ([`build.yml`](.github/workflows/build.yml)) are signed, and every signing request is approved by hand. The signed files are the installer `ArabicPDFTools-Setup-<version>.exe` and the portable launcher `ArabicPDFTools.exe`. Third-party files bundled inside the packages (the Python runtime and the libraries listed below) are included as published by their authors.

## Privacy

This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it.

Arabic PDF Tools works entirely on your computer: it reads the PDFs you choose and writes the converted files next to them. It has no telemetry, no update check and makes no network connections. Only the *script version* installer (`Install.cmd`) uses the internet, once, to download Python (via Windows' `winget`) and the libraries below from the Python Package Index (pypi.org); their privacy policies apply to those downloads.

## Uninstall

- **Installer version:** Settings → Apps → Installed apps → *Arabic PDF Tools* → Uninstall (or *Uninstall Arabic PDF Tools* in the Start menu). This also removes the right-click menu items.
- **Portable version:** delete the folder.
- **Script version:** run `Uninstall.cmd` in `%LOCALAPPDATA%\PDF2PPTX-Arabic`.

## Build from source

The Windows packages are built on Linux (as GitHub Actions does) with:

```bash
sudo apt-get install nsis zip
bash packaging/build.sh        # outputs go to dist/
```

`VERSION` holds the version number used everywhere. Every bundled component is pinned in `packaging/build.sh`.

| Folder | Contents |
|---|---|
| `src/` | The app (`ArabicPDFTools.pyw`) and converters: `fix_arabic_pdf.py` (text repair), `pdf2pptx.py`, `pdf2docx_arabic.py`, `docx_flow.py` |
| `packaging/` | `build.sh`, NSIS scripts for the installer and the portable launcher |
| `script-version/` | `Install.cmd` / `Uninstall.cmd` for the script version |

## Third-party components

Arabic PDF Tools' own source code is MIT-licensed. The Windows packages bundle these components under their own licenses:

| Component | License |
|---|---|
| CPython 3.12 ([python-build-standalone](https://github.com/astral-sh/python-build-standalone)) | PSF-2.0 |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) 2.0 | GNU AGPL-3.0 (used under the AGPL; Artifex also offers a commercial license) |
| [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) 0.5.13 | MIT |
| [python-pptx](https://github.com/scanny/python-pptx), [python-docx](https://github.com/python-openxml/python-docx) | MIT |
| [pikepdf](https://github.com/pikepdf/pikepdf) | MPL-2.0 |
| [Pillow](https://github.com/python-pillow/Pillow) | MIT-CMU |
| [fontTools](https://github.com/fonttools/fonttools) | MIT |
| [NumPy](https://numpy.org) | BSD-3-Clause (and bundled permissive licenses) |
| [OpenCV](https://github.com/opencv/opencv-python) (headless) | Apache-2.0 |
| [lxml](https://lxml.de) | BSD-3-Clause |
| XlsxWriter, wrapt | BSD-2-Clause |
| typing_extensions | PSF-2.0 |
| Deprecated, termcolor | MIT |
| fire, packaging | Apache-2.0 (packaging: or BSD-2-Clause) |

Because the packages include PyMuPDF, the packages as a whole are distributed under the terms of the GNU AGPL-3.0; the complete source code is this repository.

---

Enjoy , Beedo El-Omda
