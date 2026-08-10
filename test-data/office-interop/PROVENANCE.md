# Current LibreOffice interoperability evidence

These artifacts certify a Litchi change followed by a save in a current
Office-compatible application and a final Litchi semantic readback. They are
LibreOffice evidence, not Microsoft Office evidence.

## Runtime

- Application: LibreOffice 26.2.5.2, build
  `cd7284b4cbbfeb507e630c1aac019f4157393acb`.
- Official archive:
  `https://download.documentfoundation.org/libreoffice/stable/26.2.5/deb/x86_64/LibreOffice_26.2.5_Linux_x86-64_deb.tar.gz`.
- Archive SHA-256:
  `2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`.
- Runtime license: Mozilla Public License 2.0; the official distribution also
  carries its bundled component notices. An exact MPL-2.0 copy is stored at
  `../odf/native-resave/source/LICENSE-MPL-2.0.txt`, SHA-256
  `1f256ecad192880510e84ad60474eab7589218784b9a50bc7ceee34c2b91f1d5`.
- Execution: extracted DEB packages in a temporary directory, headless mode,
  and a new temporary user profile per save. The archive, extracted runtime,
  and profiles were removed after validation.

The generator is `tools/native-resave`. The save stage is
`tools/native_odf_resave.py`. A representative command is:

```sh
LIBREOFFICE_BIN=/temporary/runtime/opt/libreoffice26.2/program/soffice \
  python3 tools/native_odf_resave.py \
  test-data/office-interop/litchi-changed/document-properties-litchi.docx \
  test-data/office-interop/libreoffice-resaved
```

## Successful chains

| Format | Genuine source SHA-256 | Litchi change and SHA-256 | LibreOffice filter and output SHA-256 | Final Litchi readback |
|---|---|---|---|---|
| DOCX | `test-data/ooxml/docx/documentProperties.docx`; `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` | Paragraph 0, `Hello World!` -> `Litchi native resave 2026-08-10`; `e4c76a4cd17a2cc5e66d52fe2109ec62f0a09aa510e97fd291b87073b84697f2` | `MS Word 2007 XML`; `0c10b6489e3b5c02d0470f213611e32fbaa0143746d69d0333bdca821d2c5e47` | Paragraph 0 exactly matches the sentinel |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx`; `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` | First worksheet `A1`, `ID` -> sentinel; `0f1c13c528c75b5293b18f4342d5fa95024fcc604af81c1b66c058a06575bac9` | `Calc MS Excel 2007 XML`; `ade103f651e3f7cac423cd6b01d4d5a004d03470667723234105ba39672bc74d` | `Munka1!A1` is the sentinel text cell |
| PPTX | `test-data/ooxml/pptx/shapes.pptx`; `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` | First text-bearing shape on slide 0 -> sentinel; `41cc73dc78e6506900628ec4518a8dee85544f164f5ecb4134c5240c414928df` | `Impress MS PowerPoint 2007 XML`; `fcc8acffad88f5091316f67403c099a4c9eaa372e17927edfee29c35fb132034` | Slide 0 contains the sentinel in a shape text owner |
| RTF | `test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf`; `1a079582281767c1bf7afa5ef2e63553400cdbc4704aa25d9dbcc34e2c22569d` | Shape 0 text, `Textbox text.` -> sentinel; `d0bf70e50972bbf15dc9b0da96b9702d64a92a676cb81529bf28729a2cd91d71` | `Rich Text Format`; `224707aea42c7b38712bc66a76424d3666b52794e2aed010fe64d5adda54f3d9` | Shape 0 text is the sentinel plus LibreOffice's terminal paragraph newline |

Raw save-stage logs are in `logs/`. Generation and final readback commands:

```sh
cargo run --manifest-path tools/native-resave/Cargo.toml \
  --features docx,xlsx,pptx,rtf -- generate FORMAT OUTPUT
cargo run --manifest-path tools/native-resave/Cargo.toml \
  --features docx,xlsx,pptx,rtf -- readback FORMAT RESAVED
```

## Honest failures and unavailable routes

- An initial XLSX attempt used
  `ConditionalFormattingSamples.xlsx`. LibreOffice saved it, but Litchi
  correctly rejected a root-escaping `/../customXml/item1.xml` relationship.
  The successful simpler Calc fixture above replaced that attempt; the failed
  artifact is not retained as evidence.
- PPT changes fully reopened before the native stage, but LibreOffice did not
  preserve the tested semantics. It restored a hidden flag, restored the
  original two-slide order, and canonicalized a `+100` anchor edit from
  `(left=370, top=585, right=1360, bottom=1018)` to
  `(left=270, top=585, right=1260, bottom=1014)`. No PPT compatibility claim
  or artifact is retained.
- DOC and XLS were not attempted in this pass. Their absence is not evidence.
- XLSB filter `Calc MS Excel 2007 Binary` is `IMPORT`-only, so LibreOffice has
  no same-format XLSB save route.

Existing source fixtures and their derived changed/resaved artifacts retain
their checked-in upstream attribution and applicable licensing; this evidence
inventory does not relicense them. LibreOffice itself and the copied
LibreOffice corpus files are covered by MPL-2.0.
