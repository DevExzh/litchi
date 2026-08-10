# Office interoperability evidence inventory

This directory records cross-format application-resave evidence separately from
producer-created inputs. A fixture belongs here only when its exact lineage is:

1. Litchi opens and changes an input;
2. the named desktop application opens and saves that changed artifact; and
3. Litchi reopens the application-saved artifact and verifies the intended
   semantic change.

Opening a file without saving it, an exact Litchi no-op, a producer-created
input, or a filter declaration in application source code is not resave
evidence.

## 2026-08-10 Linux host inventory

- Host: Ubuntu 24.04.4 LTS, x86-64.
- Microsoft Office: unavailable. This Linux host has no Word, Excel, or
  PowerPoint executable and no Wine or PowerShell automation path.
- LibreOffice/OpenOffice: unavailable. There is no `libreoffice`, `soffice`,
  `lowriter`, `localc`, or `loimpress` executable or installed package.
- Other compatible applications: unavailable. There is no installed
  OnlyOffice, Calligra, AbiWord, Gnumeric, Pandoc, or unoconv executable.
- Containers: the Docker client is installed, but no local Office-compatible
  application image is present. No external image was pulled.
- Ubuntu advertises LibreOffice 24.2.7 packages, but they are not installed.
  Package availability is not executable interoperability evidence.

The following commands reproduced the read-only inventory:

```sh
command -v libreoffice soffice lowriter localc loimpress unoconv pandoc \
  abiword gnumeric onlyoffice-desktopeditors calligrawords calligrasheets \
  calligrastage wine powershell pwsh
dpkg-query -W -f='${binary:Package}\t${Version}\n' | \
  grep -Ei 'libreoffice|openoffice|onlyoffice|calligra|abiword|gnumeric|unoconv|pandoc|wine'
flatpak list --app
snap list
apt-cache policy libreoffice-core
```

All executable and installed-package queries produced no matching application.
The `apt-cache` query reported `Installed: (none)` and candidate
`4:24.2.7-0ubuntu0.24.04.6`.

## Exact filter status

There is no runnable compatible-application filter on this host for any format
in this wave. The names below are LibreOffice filter declarations inspected in
the local LibreOffice source checkout; they describe what a corresponding
LibreOffice build can support, not what is installed or tested here.

| Format | Declared LibreOffice filter | Declared direction | Host status |
|---|---|---|---|
| DOCX | `MS Word 2007 XML` | import and export | unavailable |
| XLSX | `Calc MS Excel 2007 XML` | import and export | unavailable |
| PPTX | `Impress MS PowerPoint 2007 XML` | import and export | unavailable |
| DOC | `MS Word 97` | import and export | unavailable |
| XLS | `MS Excel 97` | import and export | unavailable |
| XLSB | `Calc MS Excel 2007 Binary` | import only | unavailable; no declared same-format export |
| PPT | `MS PowerPoint 97` | import and export | unavailable |
| RTF | `Rich Text Format` | import and export | unavailable |

The declaration paths are
`filter/source/config/fragments/filters/MS_Word_2007_XML.xcu`,
`calc_MS_Excel_2007_XML.xcu`,
`impress_MS_PowerPoint_2007_XML.xcu`, `MS_Word_97.xcu`,
`MS_Excel_97.xcu`, `calc_MS_Excel_2007_Binary.xcu`,
`MS_PowerPoint_97.xcu`, and `Rich_Text_Format.xcu`. The XLSB declaration has
`IMPORT` but no `EXPORT` flag.

## Evidence produced

None. No Office-compatible application was available to perform a genuine
resave, so this inventory adds no DOCX, XLSX, PPTX, DOC, XLS, XLSB, PPT, or RTF
fixture and makes no Microsoft Office or compatible-application
interoperability claim.
