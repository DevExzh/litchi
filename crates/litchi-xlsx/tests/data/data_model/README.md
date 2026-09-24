# Native Data Model fixture

`tdf167689_x15_namespace.xlsx` is an unchanged copy of LibreOffice core's
`sc/qa/unit/data/xlsx/tdf167689_x15_namespace.xlsx`, from the local checkout at
commit `02211da428a009ea0b25ef725de8cb2f761ce4da`.

- Original repository: <https://git.libreoffice.org/core>
- File size: 174,497 bytes.
- SHA-256: `b9064db5694c5e60928a7b6732aabf440c7d2f997d907834e38f927342ec8688`.
- Upstream license texts are retained in `COPYING`, `COPYING.LGPL`, and `COPYING.MPL`.

The fixture contains a native XLSX Data Model and a type-102 linked-table
connection named `LinkedTable_Tabelle1`. Its separate `ThisWorkbookDataModel`
connection carries the model flag. Tests read these declarations as inert data;
they do not execute queries, refresh connections, or evaluate the model.

This input uses outer header version 150. The typed XLDM inspector selects the
fixture-backed `StorageProfile::Tabular150`; canonical version-140 storage keeps
its own allocation and metadata rules. The fixture reports Microsoft Excel
AppVersion `16.0300` and workbook `rupBuild="20417"`; these are source metadata,
not certification of every Office build with those values.

The standalone storage regression reads all 48 allocations, retains the source
slice without copying the complete payload, and verifies exact output through
`xldm::write`. It rejects a changed CRC in each allocation. The storage inspector
does not author inner table/column structures or execute the model. Package-level
descriptor edits and native application open/save evidence are separate work.

The separate compression API also exercises all 46 compressed members through
explicit `XpressFraming::Tabular16`. Their 78 frames use two little-endian 16-bit
lengths: 75 contain Xpress data and three store raw bytes with equal lengths.
These are observations of this fixture. The existing `decompress_xpress` API
continues to select the two signed 32-bit lengths specified by
[MS-WUSP §2.1.1](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-WUSP/3e24630e-8000-4894-a967-315df7ed996e).
Framing is never auto-detected; equal lengths have different meanings in the two
profiles. There is no version-150 inner-payload writer or tabular framing encoder.

Run `python3 crates/litchi-xlsx/tests/data/data_model/native_tabular_members.py`
from the repository root to reproduce the member offsets, stored lengths,
independent decoded lengths from the backup log, and frame counts as JSON.
The extractor pins the original fixture hash and checks all 48 native CRC
markers. It is a fixture-specific evidence tool, not a general XLDM reader.
The Rust regression checks each decoded length and a one-byte-smaller output
limit; the first decoded member also has the expected Load XML boundaries.

The native storage profile includes marker BOMs in their allocations, checks
complemented CRCs, reads backup-log version `11.53`, and maps logical source
paths separately from hexadecimal storage keys. It accepts the native nonzero
folder versions, spaces, and Unicode identifiers without relaxing canonical
path classification. It verifies framed ranges and decoded-size declarations
without decompressing members during outer inspection. Complete decompression
remains an explicitly bounded operation. The strict version-140 checks remain
unchanged.
