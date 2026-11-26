# Real-world metafile fixtures

Windows Metafiles (`wmf/`), Enhanced Metafiles including EMF+ (`emf/`),
Macintosh PICT pictures (`pict/`) and raw OfficeArt BLIP records (`blip/`)
carved out of legacy documents. Each file is byte-for-byte as stored in its
document, except that compressed BLIP payloads under `emf/`, `wmf/` and
`pict/` are stored decompressed.

`manifest.tsv` lists every fixture with its SHA-256, the source corpus and
document, the container inside that document, the byte offset in that
container, the extraction route, and how many corpus documents contain the
identical bytes.

`fuzz/` holds inputs that libFuzzer minimized from seeds drawn from these
metafiles; each one once panicked, exhausted memory, ran far too long,
crashed the rasterizer or produced SVG that is not well-formed XML, and must
now convert or fail with an error.

## Sources

- **GovDocs1** (`govdocs1`): the Digital Corpora "GovDocs1" corpus of
  documents collected from U.S. government web servers and redistributed for
  research (<https://digitalcorpora.org/corpora/file-corpora/files/>;
  Garfinkel et al., "Bringing science to digital forensics with standardized
  forensic corpora", DFRWS 2009). The `document` column is the corpus file
  path (`NNN/NNNNNN.ext`). Works of the U.S. federal government are not
  subject to copyright in the United States (17 U.S.C. 105); individual
  documents may carry their own notices.
- **3GPP** (`3gpp`): pictures from releases of 3GPP TS 25.212 "Multiplexing
  and channel coding (FDD)" published at
  <https://www.3gpp.org/ftp/Specs/archive/25_series/25.212/>. The `document`
  column is the archive path and the document inside it. 3GPP specifications
  are copyright of the 3GPP Organizational Partners (ARIB, ATIS, CCSA, ETSI,
  TSDSI, TTA, TTC); these small extracts are kept for interoperability
  testing.

## Containers and extraction routes

The `container` column names the OLE2 stream (for example `Pictures`,
`Data`, `WordDocument`, `ObjectPool/_1234/\x02OlePres000`,
`\x05SummaryInformation`), the package part, or the RTF group the bytes came
from. `Workbook#MsoDrawingGroupN` means the BIFF `MSODRAWINGGROUP` record
data with its `CONTINUE` records joined; `#ExOleObjStg@N` marks an embedded
object inside a PowerPoint document; `!name` separates a ZIP member.

| `extraction` | Meaning |
|---|---|
| `blip` | A complete OfficeArt BLIP record (`msofbtBlipEMF`, `msofbtBlipWMF`, `msofbtBlipPICT`), header included |
| `blip-emf`, `blip-wmf`, `blip-pict` | The metafile payload of such a BLIP, decompressed when stored with DEFLATE |
| `raw-emf` | An EMF found by its header signature inside a stream (document thumbnails, print caches, embedded objects) |
| `raw-wmf`, `raw-wmf-placeable` | A WMF found by its header signature, without or with a placeable header (for example Word picture data) |
| `raw-pict2` | A PICT version 2 picture found by its header signature |
| `olepres-wmf`, `olepres-emf` | The picture of an OLE presentation stream ([MS-OLEDS] 2.3.4) |
| `package-emf`, `package-wmf`, `package-pict` | A media part of an OOXML or ODF package |
| `rtf-wmetafile`, `rtf-emfblip`, `rtf-macpict` | The hex payload of an RTF `\pict` group |
| `imdata-pict` | The Macintosh picture of a BIFF `IMDATA` record |
| `eps-wmf` | The WMF preview of a DOS EPS binary header |

## Selection

The fixtures were chosen from all metafiles in both corpora to include, with
the smallest available sample, every EMF and WMF record type, EMF+ record and
object type, PICT opcode, QuickTime codec and font character set observed,
each extraction route, inputs that exercise each distinct conversion
outcome, plus a random sample of small files for every combination of format,
extraction route and source document type.
