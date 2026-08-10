# Native LibreOffice ODF resave provenance

This directory intentionally contains no native-resaved artifact yet. On
2026-08-10, the test host had no `libreoffice`/`soffice` executable, no
`unoconv`, and no Python `uno` module. Consequently, no native-open or
native-resave claim can be made from this environment.

The reproducible second stage is
[`tools/native_odf_resave.py`](../../../tools/native_odf_resave.py). A format
owner must first produce a Litchi-changed file, invoke that tool with a distinct
output directory, and then reopen the resulting file with the same Litchi
format API. Only that complete chain is native interoperability evidence.

The filter map comes from LibreOffice's checked-in primary registry under
`3rdparty/libreoffice-core/filter/source/config/fragments`. The accompanying
Python test verifies the registry entries directly.

| Family | Extension | CLI filter | Registry capability |
|---|---:|---|---|
| ODT | `.odt` | `writer8` | import/export |
| ODS | `.ods` | `calc8` | import/export |
| ODP | `.odp` | `impress8` | import/export |
| Formula | `.odf` | `math8` | import/export |
| ODC | `.odc` | `chart8` | import/export, hidden from file dialog/chooser |
| ODG | `.odg` | `draw8` | import/export |
| ODI | `.odi` | none | no registered ODI media type or filter |
| ODM | `.odm` | `writerglobal8` | import/export |
| OTH | `.oth` | `writerweb8_writer_template` | import/export template |
| ODB | `.odb` | `StarOffice XML (Base)` | import-only; requires UNO document `store()` for a same-package save |

Run the read-only probe with:

```text
python3 tools/native_odf_resave.py --probe
```

Run a supported resave with an isolated LibreOffice profile and a distinct,
pre-existing output directory:

```text
python3 tools/native_odf_resave.py path/to/litchi-changed.odt path/to/output
```

The harness refuses ODI and ODB rather than silently converting either into a
different family. Synthetic documents and chart/image subdocuments embedded in
another ODF package must retain those provenance labels and must not be entered
here as standalone native evidence.
