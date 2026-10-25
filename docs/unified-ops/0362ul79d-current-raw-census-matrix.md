# 0362ul79d — current raw-census profile matrix and coverage prerequisites

This is a source-only reconciliation of the raw-public-API census at
a574488eaa6ec62913fb8c380235e132e07f967d. It does not run Cargo, rustdoc,
native code, a benchmark, or a final inventory capture. It does not relabel
the older manifest source or any older capture as evidence for this source.
The current glossary and package counts are current observations, not a final
zero-debt result.

## Inputs and source-only checks

The checkout was clean at the candidate before this record was added. The
candidate tree is 2d0a237d9d7648172986a778f43b281af8f24ec4. The relevant
tracked identities are:

| input | observation |
|---|---|
| tools/api_glossary.json | 1,049,135 bytes; SHA-256 d8051202b778dcb002178f82bad5e79e4229a43fd91db5215135813285b028e5; 68 profile rows and 52 distinct package names |
| tools/raw_api_coverage.json | 20,249 bytes; SHA-256 c162284a7b8d7c7cc80505855b4e4c33640a31a2f45207613832f5852df63fd3; 15 declared profiles, 22 entries, 12 lanes; coverage.state=incomplete |
| root Cargo.toml | SHA-256 d46816f76fedce8373a1735f960654a3ba9f5ac66193a06830018c1c5653f2ff |
| rust-toolchain.toml | SHA-256 e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310 |
| direct crates/* library packages | 68; tools.final_api_inventory.package_set_sha256 = 60cb4d20c5d41b89c6789028fd27f5a1c40b10923e8c9120da423e1967069f65 |

The ignored root Cargo.lock was not needed for this source-only audit. A
future capture must seed the coordinator-approved ignored lock
efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811, as
required by 0362ul79b; this record makes no lockfile or capture claim.

The dynamic checks were run without Cargo and with bytecode disabled:

    PYTHONDONTWRITEBYTECODE=1 python3 tools/raw_api_coverage.py \
      --root . --manifest tools/raw_api_coverage.json \
      --glossary tools/api_glossary.json --json

The 12,142-byte JSON report was kept outside Git at
/tmp/raw79d-current.json (SHA-256
01bb258f25bf20286e188ca7c3c492003fd1660a5388f16b5e129d8b27c88f99). It
reports complete=false, profile_mismatches=0, source_failures=0 at the
manifest's pinned source ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb, and:

* 59 of the 68 glossary profiles are not declared in the provisional
  manifest;
* 56 direct package names are not represented by the provisional 12-package
  declaration set;
* all 15 declared profiles lack a supplied rustdoc input and receipt; and
* current_source_failures=26 because the provisional source/anchor blobs
  predate this candidate. Those failures are not repaired by changing the
  old source commit or by reusing its captures.

The exact missing glossary profile IDs are:

    biff cfb core-minimal doc-full doc-minimal docx-full docx-minimal drawingml
    facade-iwork facade-sheet-text fonts-full fonts-minimal formula-raw imgconv iwa
    iwa-archive iwa-common iwa-core iwa-detect iwa-protos-wire-codecs iwa-structured
    iwa-text-wire keynote-full keynote-minimal markdown numbers-full numbers-minimal
    odb odc odf-common odg odi odm odp odraw odt ograph ole-common
    ooxml-common-full ooxml-common-minimal opc-full opc-minimal ops-iwork oth pages
    ppt-full ppt-minimal sheet-text sign soapberry-zip spreadsheet-drawing-full
    spreadsheet-drawing-minimal vba xldm xls-minimal xlsb-full xlsb-minimal
    xlsx-full xlsx-minimal

The exact direct packages that have no glossary package row are:

    litchi-cli litchi-codepage litchi-eval litchi-iwa-cache litchi-iwa-graph
    litchi-iwa-index litchi-iwa-package litchi-iwa-text litchi-numbers-wire
    litchi-odf litchi-py litchi-sheet litchi-slide litchi-word xml-minifier
    xml-minifier-macros

The source-only tooling checks passed independently:

    PYTHONDONTWRITEBYTECODE=1 python3 -B -m unittest -v tools.test_raw_api_coverage
      17 tests, exit 0
    PYTHONDONTWRITEBYTECODE=1 python3 -B -m unittest -v \
      tools.test_api_debt_inventory tools.test_final_api_inventory
      23 tests, exit 0

These are validator and fixture tests only. They do not close a rustdoc
inventory or an API-debt row.

## Reconciled final profile matrix

The final matrix is derived from the current glossary tuples
(id, package, default_features, features), not from a package-name or
feature-name guess. The exact 68 current rows are:

| package | profile(s), with exact feature closure |
|---|---|
| litchi | facade-iwork (no-default; iwork); facade-sheet-text (no-default; sheet-text) |
| litchi-biff | biff (no-default; none) |
| litchi-cfb | cfb (no-default; write) |
| litchi-core | core-minimal (no-default; none) |
| litchi-crypto | crypto (no-default; ooxml) |
| litchi-doc | doc-minimal (no-default; none); doc-full (no-default; formula,performance-diagnostics) |
| litchi-docx | docx-minimal (no-default; none); docx-full (no-default; automatic-fonts,encryption,sign,vba-inspection) |
| litchi-drawingml | drawingml (no-default; none) |
| litchi-fonts | fonts-minimal (no-default; subset); fonts-full (no-default; automatic) |
| litchi-formula | formula-raw (default; none) |
| litchi-imgconv | imgconv (no-default; none) |
| litchi-iwa | iwa (default; none) |
| litchi-iwa-archive | iwa-archive (default; internal-iwork-source) |
| litchi-iwa-common | iwa-common (default; none) |
| litchi-iwa-core | iwa-core (default; none) |
| litchi-iwa-detect | iwa-detect (default; none) |
| litchi-iwa-protos | iwa-protos-wire-codecs (default; none) |
| litchi-iwa-structured | iwa-structured (default; none) |
| litchi-iwa-text-wire | iwa-text-wire (default; none) |
| litchi-keynote | keynote-minimal (default; none); keynote-full (default; internal-iwork-source) |
| litchi-markdown | markdown (default; none) |
| litchi-numbers | numbers-minimal (default; none); numbers-full (default; internal-iwork-source) |
| litchi-odb | odb (default; none) |
| litchi-odc | odc (default; none) |
| litchi-odf-common | odf-common (default; none) |
| litchi-odf-formula | odf-formula (default; none) |
| litchi-odg | odg (default; none) |
| litchi-odi | odi (default; none) |
| litchi-odm | odm (default; none) |
| litchi-odp | odp (default; none) |
| litchi-odraw | odraw (no-default; none) |
| litchi-ods | ods (default; none) |
| litchi-odt | odt (default; none) |
| litchi-ograph | ograph (no-default; none) |
| litchi-ole-common | ole-common (no-default; performance-diagnostics) |
| litchi-ooxml-common | ooxml-common-minimal (no-default; none); ooxml-common-full (no-default; encryption,vba-inspection) |
| litchi-opc | opc-minimal (no-default; none); opc-full (no-default; sign,performance-diagnostics) |
| litchi-ops | ops-rtf (no-default; rtf); ops-odf (default; odt,ods,odp,oth,odm,odg,odc,odi,odb,odf-formula); ops-iwork (default; pages,keynote,numbers); ops-all (no-default; doc,docx,encryption,eval,formula,iwork-adapter,keynote,markdown,numbers,odb,odc,odf-formula,odf-objects,odg,odi,odm,odp,ods,odt,opc,oth,pages,ppt,pptx,rtf,test-support,xls,xlsb,xlsx,yaml) |
| litchi-oth | oth (default; none) |
| litchi-pages | pages (default; none) |
| litchi-ppt | ppt-minimal (no-default; none); ppt-full (no-default; automatic-fonts,encryption,formula,performance-diagnostics,sign,vba-inspection) |
| litchi-pptx | pptx-minimal (no-default; none); pptx-full (no-default; automatic-fonts,encryption,vba-inspection) |
| litchi-rtf | rtf-minimal (no-default; none) |
| litchi-sheet-text | sheet-text (no-default; none) |
| litchi-sign | sign (no-default; none) |
| litchi-spreadsheet-drawing | spreadsheet-drawing-minimal (no-default; none); spreadsheet-drawing-full (no-default; none) |
| litchi-vba | vba (no-default; none) |
| litchi-xldm | xldm (no-default; none) |
| litchi-xls | xls-minimal (no-default; none) |
| litchi-xlsb | xlsb-minimal (no-default; none); xlsb-full (no-default; sign,vba-inspection) |
| litchi-xlsx | xlsx-minimal (no-default; none); xlsx-full (no-default; encryption,vba-inspection) |
| soapberry-zip | soapberry-zip (no-default; none) |

The 16 package-only rows needed to make the direct-library set complete are
the following. They are not glossary rows and therefore still require the
same callable-level disposition audit:

| profile ID | package | exact capture configuration |
|---|---|---|
| workspace-litchi-cli | litchi-cli | no-default; no features |
| workspace-litchi-codepage | litchi-codepage | no-default; no features |
| workspace-litchi-eval-full | litchi-eval | no-default; web_functions |
| workspace-litchi-iwa-cache | litchi-iwa-cache | no-default; no features |
| workspace-litchi-iwa-graph | litchi-iwa-graph | no-default; no features |
| workspace-litchi-iwa-index | litchi-iwa-index | no-default; no features |
| workspace-litchi-iwa-package | litchi-iwa-package | no-default; no features |
| workspace-litchi-iwa-text | litchi-iwa-text | no-default; no features |
| numbers-wire-0362ul66 | litchi-numbers-wire | no-default; no features |
| workspace-litchi-odf-full | litchi-odf | no-default; all |
| workspace-litchi-py | litchi-py | no-default; no features |
| workspace-litchi-sheet | litchi-sheet | no-default; no features |
| workspace-litchi-slide | litchi-slide | no-default; no features |
| workspace-litchi-word | litchi-word | no-default; no features |
| workspace-xml-minifier | xml-minifier | no-default; no features |
| workspace-xml-minifier-macros | xml-minifier-macros | no-default; no features |

Thus the normalized final input count remains 84, but its composition is now
68 glossary profiles plus 16 package-only profiles. The earlier 65-plus-19
description in 0362ul79 is historical for its pre-Formula/VBA/XLDM glossary;
it is superseded for this candidate, not edited retroactively.

The newly merged rows are all represented canonically: formula-raw, vba,
xldm, fonts-minimal, fonts-full, and markdown are present in the 68-row
glossary. The provisional formula-0362ul69 and markdown-0362ul70 tuples
duplicate the canonical rows; vba-0362ul67, xldm-0362ul68, and
fonts-0362za are provisional profile spellings/configs. None of those
provisional artifacts may be reused as a canonical capture. At final freeze,
migrate their provisional entries to the canonical profile IDs after fresh
canonical rustdoc captures; keep only the package-only numbers-wire-0362ul66
row from that group. A separate exact capture would be needed if a
coordinator deliberately retains any provisional feature tuple.

## Provisional lanes and source-anchor state

The current manifest declares these lanes:

    0362ul44  0362ul61  0362ul62  0362ul63  0362ul65  0362ul66
    0362ul67  0362ul68  0362ul69  0362ul70  0362za  crypto-raw

Current engineering returns are PPTX 0362ul61/62/63, PPTX metadata/history
0362ul81/82 (not yet represented in this provisional lane table), and Numbers
wire 0362ul66. The final PPTX pptx-minimal/pptx-full rustdoc census must
enumerate every callable added or changed by 61/62/63/81/82, including the
physical aggregate getters. Their disposition follows the actual owner/raw
boundary: a physical OpcPackage graph consumed through Package::edit_opc is
not an owner violation merely because a method has a mutation-sounding name,
while a public format-facing raw getter or unbounded codec remains a real row.
No name-only exemption is valid.

The other provisional lanes are source-integrated or awaiting final-source
capture/review, not final active exceptions: 0362ul44 (ODF), 0362ul65 (RTF),
0362ul67 (VBA), 0362ul68 (XLDM), 0362ul69 (Formula), 0362ul70 (Markdown),
0362za (fonts), and crypto-raw. Every provisional source anchor must be
rebound to the final frozen commit, and every active entry must be classified
or repaired. Final --require-complete requires no active entries and
lanes=[]; the present lane table cannot satisfy it.

The 26 current-source failures are grouped as follows:

* two trusted coverage blobs (tools/api_glossary.json and Cargo.toml);
* seventeen callable anchors: four litchi_ops::detect functions, four RTF
  validation/compression functions, one VBA Project::read, two XLDM
  functions, three Formula functions, one Markdown Snapshot::read, one ODS
  tracked-changes parser, and one ODF-formula serializer; and
* seven lane anchors: 0362ul65, 0362ul67, 0362ul68, 0362ul69, 0362ul70,
  0362ul44 and 0362za.

Those are expected source drift from the provisional ef1 pin, not evidence
that the old captures match a574. They must be re-anchored only after the
coordinator freezes the final source.

## Current boundary verification outside the active PPTX/Numbers returns

The earlier version of this section repeated pre-migration findings from
0362rd. The current a574 source was checked directly. The following are
closures of the raw-limit findings, with exact current signatures and source
identities; they are not new engineering returns:

| surface | current canonical signatures and source identity | closure record |
|---|---|---|
| ODF package codec | package/codec.rs blob SHA-1 6afb1bc006bffc0d9f1e123694c2fd58b0b7e720: Archive::new(data: &'data [u8], limits: ArchiveLimits) -> Result<Self> at line 126; Archive::read_manifest_xml(&self, limits: XmlStreamLimits) -> Result<String> at 223; read_manifest(archive: &Archive<'_>, limits: XmlStreamLimits) at 284; parse_manifest(xml: &str, limits: XmlStreamLimits) at 592 | 0362ro/source commit 49b447f673986863ee5dacd23d13d1d8aae2089e: old default constructor, fixed manifest ceiling, and read_manifest_xml_with_limit are gone |
| ODF writer | core/writer.rs blob SHA-1 68941b901e2779a12dd7f83c08ae5db31319fd30: PackageWriter::new(limits: StreamingArchiveLimits) at 727; new_bounded(limit: usize, limits: StreamingArchiveLimits) at 750; with_writer(writer: W, limits: StreamingArchiveLimits) at 763 | 0362ro/source commit 49b447f673986863ee5dacd23d13d1d8aae2089e: only explicit-policy constructors remain; old PackageWriterLimits and default twins are absent |
| CFB ingress and writers | file.rs blob SHA-1 f5798a420ba961c7ffca3c456969e154fdd699d2, OleFile::open(reader, limits: OleFileLimits) at 1048; shared.rs blob SHA-1 88f09a1f69ec2242434f3802a01437a5c0f54667, SharedOleFile::open(source, limits), open_owned(source, version, limits), open_owned_vec(source, version, limits) at 634–681; validation.rs blob SHA-1 784c49f28cb8577d8ed41b08ee255babcd90fbf4, validate_source(source, source_limits, report_limits) at 76; writer/core.rs blob SHA-1 413d9e9e25bcb19e40cd36a5be48792a3750d06e, move_storage(..., limits: StorageMoveLimits) at 1282 and adopt_source_layout(..., limits: OleFileLimits) at 466; writer/sequential.rs blob SHA-1 379867bc442067bd5048e98edbf0fb440acce7f9, consuming limits(limits: SequentialWriterLimits) and limits_value() at 198–205 | 0362zad/source commit bf7f48741bc02d1f5e736a7a732c25808d28af36: default twins, public _with_limits names, and OleWriter::update_stream are removed; no public CFB _with_limits, with_limits, or update_stream declaration remains in the current source scan |
| generic ZIP raw readers/writers | office.rs blob SHA-1 0166220513e005779bd6e979e5c8ee1f729b0f96: ArchiveReader::new(data, limits: ArchiveLimits) at 2893; IndexedArchive::preservation_index(..., limits), from_reader(..., limits), from_reader_with_policy(..., limits, policy), from_zip_archive(..., limits), and from_zip_archive_with_policy(..., limits, policy) at 3972–4019; LocatedArchive::index(limits) and index_with_policy(limits, policy) at 5744–5755; StreamingArchiveWriter::new(limits: StreamingArchiveLimits), with_writer(writer, limits), with_writer_and_spool(..., limits, ...), and with_writer_and_spool_and_name_plan(..., limits, ...) at 7235, 7263, 7302 and 7342; LazyArchiveReader::new(data, limits: ArchiveLimits, cache_limits: LazyArchiveCacheLimits) at 8757; preserve.rs blob SHA-1 42b0d399b151ab20e396d97293688055282d8fec, PreservationIndex::new(..., limits: ArchiveLimits) at 776 | 0362zae/source commit 0e3cfc0e1cb0efe8c371623e42c347aa90640461 and 0362zaea: raw ZIP defaults and public _with_limits/with_limits twins are removed; lazy cache constructors/setters use their canonical explicit policies |

The limit-closure statements above are narrower than an assertion that every
historical alias spelling in ZIP has disappeared. The current writer.rs blob
SHA-1 is 1787c62b5cf891af784034ada81ebf7435b7b7d0: `start_entry_owned` at
line 1960 delegates to the canonical `start_file_owned` at line 1932, while
`start_entry_owned_zip64` at line 2002 delegates to
`start_file_owned_zip64` at line 1974. They are public entry-oriented
forwarders, not limit-bearing constructors. The current errors.rs blob
SHA-1 is 6ff49eada6a333e28a310c372e883df708b6712b: `ErrorKind::IO` at line
159 is accompanied by the explicitly documented compatibility variant
`ErrorKind::Io` at line 162. These exact callable and compatibility roles are
separate surfaces for the final rustdoc and glossary disposition; their names
alone do not make either a required limit migration or an accepted
exemption. This record neither waives nor closes them.

The OPC boundary is likewise current-source checked, not copied from the
older scope note. package.rs blob SHA-1
db4365c323f4338e9df916916500456894f39459 has public
OpcPackage::{open,from_reader,from_vec,from_shared_vec,from_bytes} with
SourceLimits at lines 1382–1541; phys_pkg.rs blob SHA-1
02df68f316e0902dea4193f71b18781d575fed7d has
OwnedPhysPkgReader::{open,from_bytes,from_reader} and PhysPkgReader::new
with SourceLimits at lines 93–256. source_backed.rs blob SHA-1
cc8f9cf38c7c7dd4e0c49af52bae8ec55c5275ba has the deferred source-owner
constructors, while the remaining _with_limits helpers are crate-private.
The format-facing raw wrappers actually audited here are
`litchi_xlsx::Package::edit_opc`, `litchi_xlsb::Package::edit_opc`,
`litchi_pptx::Package::edit_opc` and `litchi_docx::Package::edit_opc`; the
shared OOXML staged publication paths
`litchi_ooxml_common::custom_xml::{Edit::new,Patch::apply}` accept
`&OpcPackage`, while `litchi_ooxml_common::custom_xml::Edit::commit`
publishes the new package.
This is not a claim
that every `&mut OpcPackage` helper is private: for example, the public
`litchi_ooxml_common::custom_xml::add(&mut OpcPackage, ...)` re-export (defined
at package.rs line 81; package.rs blob SHA-1
2b7ddc1819796c98cdcd1bd9e1dc4977dc9aa15b) is a public free format helper
and must be classified from its actual rustdoc signature and 0351 boundary,
not excused because it lives in `common`.
Conversely, public `&mut self` methods defined on `litchi_opc::OpcPackage`
itself—such as `get_part_mut`, `add_part`, `remove_part` and `rels_mut`—are
raw physical-graph operations, not the format-owner wrappers listed above.
0362rl is the closure record for the source-policy and ingress migration; it
does not resolve those distinct raw-graph roles. A fresh rustdoc profile is
still required, and this record schedules no new lane from a name-only match.
The PPTX 61/62/63/81/82 aggregate/getter rows still require fresh
profile-level census and disposition.

The historical 0362ul1 diagnostic contains 440 inventory rows from source
49c25ef67f62a8bb943e7c6a0bf464e49ae87d07, and the 84-row reviewed
disposition files are anchored to 082fbeafe8. Both are stale relative to
this candidate. With no current rustdoc inputs in this source-only worktree,
there is no evidence-backed current row list to close or to reclassify as a
semantic exemption. The final run must separately regenerate the 68
all-feature inventory inputs and the 84 raw-census profile inputs, then review
every non-draft row with a current source anchor; the only permitted deferral
remains the redaction audit.

## Exact prerequisites before final capture

1. Freeze a clean candidate and bind the current 68-row glossary, 68-package
   direct set, 16 package-only profiles, exact root lock, toolchain, runner,
   environment and external target/TMP identities.
2. Capture 84 distinct feature-profile rustdoc JSON inputs and receipts for
   the raw census. Separately run the 68-package all-feature capture required
   by api_debt_inventory. The 84 profile captures do not prove the 68
   all-feature inventory, and an all-feature artifact cannot silently
   substitute for a feature-profile input. The current 15 provisional profiles
   have no usable input/receipt; no old artifact can be substituted merely
   because its package name matches.
3. Rebind the expected source blobs, every callable anchor and every
   provisional lane anchor to the frozen commit. Migrate Formula, VBA, XLDM,
   fonts and Markdown entries to their canonical profile IDs, and add the
   final PPTX 61/62/63/81/82 and Numbers-wire callable rows from the fresh
   surfaces.
4. Generate the 68-package all-feature api_debt_inventory from current
   rustdoc and attach a concrete disposition/evidence row to every non-draft
   finding. Existing historical TSVs cannot satisfy this prerequisite.
5. Close every active entry into a required-limits, configured/retained-owner,
   semantic/fixed-scalar, or source-proven private-removed disposition. The
   final manifest must have no active entries and no lanes.
6. Run the final raw coverage validator with exactly the 84 profile bindings;
   only an empty missing/stale/unclassified/signature/source/receipt report
   and complete=true is admissible.

No final capture, API-debt closure, performance claim, or acceptance claim is
made by this record.
