# Change 0344: catalog-first OOXML detection

Date: 2026-08-31

Status: implementation described; focused validation recorded

performance_claim: none

## Problem and eager seam

The established smart detector accepts an owned Vec<u8> and returns an eager
DetectedFormat. For an OOXML result, the enum still carries an eager
OpcPackage. That is the compatibility boundary: callers that use the smart
API receive the existing parsed owner and its existing ownership semantics.

The resource problem is the separate detection path used by typed facades.
Family identification needs only the ZIP/OPC catalog, content-type metadata,
and relationship manifests, but an eager package path can retain or load
package-wide state before the caller knows whether the input is DOCX, PPTX,
XLSX, or XLSB. Ordinary application parts such as word/document.xml,
ppt/presentation.xml, xl/workbook.xml, worksheets, media, and embeddings are
not format-detection inputs and must not be read merely to classify a package.

This change records the catalog-first seam used by the OOXML probes: validate
bounded package structure, retain structural metadata, classify from declared
content types, and leave ordinary payloads in the source. It does not turn the
eager smart result into a borrowed result and does not add semantic payload
detection.

## Actual structural owner: PackageCatalog

litchi-opc owns the public PackageCatalog. It is an owned structural catalog,
not a public borrowed catalog. Its private source contains admitted part
names, content types, relationship manifests, and non-part classification.
PackageCatalog exposes only:

- part_count(), the number of admitted ordinary OPC parts, excluding
  relationship parts;
- part_content_types(), an exact-size iterator borrowing the catalog's
  content-type strings without allocating or reading ordinary payloads.

Catalog construction reads the ZIP directory and validates OPC structure. It
reads [Content_Types].xml and the relationship members needed to build the
structural relationship graph. The graph walk may decompress small .rels
members, including dangling relationship sources, because relationships are
structural metadata. It does not read or decompress ordinary part payloads.
The catalog retains no ordinary payload bytes; those remain in the ZIP/source
for a later, explicit owner operation.

The seekable-reader adapter is private. It temporarily borrows the caller's
mutable reader through a fallible RefCell adapter and exposes a
PackageCatalog, never an archive handle or reader handle. This is the only
borrowed reader boundary in the implementation. There is no public borrowed
catalog, catalog cursor, path-catalog constructor, or payload-bearing catalog
iterator.

## Implemented typed and legacy APIs

The public structural OPC APIs are:

    litchi_opc::probe_package_catalog_from_reader(&mut R)
    litchi_opc::probe_package_catalog_from_reader_with_limits(&mut R, ReadLimits)

Both return litchi_opc::Result<PackageCatalog>. They validate the same catalog
admission used by the source-backed reader while leaving ordinary payloads
unread.

The public OOXML facade APIs are:

    litchi::detection_smart::ooxml::try_detect_zip_format_with_limits(
        &[u8], ReadLimits,
    )
    litchi::detection_smart::ooxml::try_detect_zip_format_from_reader_with_limits(
        &mut (impl Read + Seek), ReadLimits,
    )
    litchi::detection_smart::ooxml::try_detect_ooxml_format_from_source_backed_package(
        &SourceBackedPackage,
    )

The first two return crate::opc::Result<Option<FileFormat>>. The byte probe
checks ZIP magic, wraps the slice in a seekable cursor, and uses the OPC
reader probe. The reader probe consumes the resulting PackageCatalog and
classifies its admitted content types. The source-backed classifier returns
the same typed Result<Option<FileFormat>> while preserving execution-policy
and source-freshness failures.

The established Option wrappers remain available and intentionally flatten
errors:

- detect_zip_format and detect_zip_format_with_limits;
- detect_zip_format_from_reader and detect_zip_format_from_reader_with_limits;
- detect_ooxml_format, detect_ooxml_format_with_limits,
  detect_ooxml_format_from_bytes, and detect_ooxml_format_from_bytes_with_limits;
- detect_ooxml_format_from_package for an already eager OpcPackage;
- detect_file_format, detect_file_format_from_bytes, and
  detect_format_from_reader at the facade boundary;
- detect_format_smart and detect_format_smart_with_limits, whose OOXML
  DetectedFormat variants remain eager OpcPackage owners.

The source-backed document, presentation, and workbook probes are private
facade handoffs, including the DOCX, PPTX, workbook, and ODT path helpers.
They are not additional public catalog APIs. No public API promises a generic
catalog-to-owner conversion; each existing facade keeps its own feature
gates, semantic admission, error mapping, and source owner.

## Catalog reads and ordinary-payload exclusion

The catalog path is deliberately structural:

- ZIP signatures, EOCD/ZIP64 location, central-directory records, names,
  offsets, sizes, and bounded archive metadata are inspected;
- [Content_Types].xml is read as a designated structural member;
- package and part relationship members needed by OPC topology validation are
  read as structural members;
- PackageCatalog::part_content_types() supplies the classifier's input
  without opening an admitted ordinary part.

The classifier does not open WordprocessingML, PresentationML, SpreadsheetML,
XLSB payload records, media, embeddings, custom XML, formulas, previews,
macros, controls, links, or arbitrary opaque members. Structural
relationship/content-type reads are expected and are not ordinary-payload
reads. A successful catalog is not semantic validation of the application
parts; a later format owner may still reject the package.

## Exact content-type matching and family precedence

OoxmlContentTypeMarkers compares each admitted content-type string with the
owning OPC constants using eq_ignore_ascii_case. It does not use substring
matching or contains. The recognized exact constants are:

- Word main document/template and their macro-enabled variants;
- PowerPoint presentation, slideshow, template, and their macro-enabled
  variants;
- the XLSB binary sheet constant;
- the XLSX sheet/template and macro-enabled sheet/template constants.

If a catalog carries more than one recognized marker, the existing precedence
is Word, PowerPoint, XLSB, then XLSX. Casing differences in the declared MIME
token are accepted by the classifier; an arbitrary token that merely contains
an Office MIME substring is not. No content-type marker authorizes a payload
read or executable behavior.

The ODF MIME classifier separately trims leading/trailing ASCII whitespace,
then matches the complete UTF-8 value against its exact supported ODF MIME
constants. Unknown, invalid, or noncanonical values do not establish an ODF
family.

## Reader cursor, limits, freshness, and execution checks

probe_package_catalog_from_reader_with_limits records the reader's original
position, seeks to the end to apply the input limit, builds an indexed ZIP
view through the private borrowed-reader adapter, and restores the original
position before returning. Its result precedence is explicit:

- successful probe plus successful restore returns the catalog;
- successful probe plus failed restore returns the restore I/O error;
- failed probe plus either restore result returns the primary probe error.

Therefore a primary ZIP/OPC probe error is never replaced by a secondary
cursor-restore error when both operations fail. The try_detect reader
functions preserve that typed error; the detect reader wrappers flatten it to
None for compatibility.

ReadLimits is applied to input bytes, archive members and metadata, member
names, compressed and uncompressed entry sizes, total bytes, content-type
bytes, relationship parts, and checked allocation paths during catalog
admission. A limit, ZIP, relationship, XML, allocation, or I/O failure is
typed by the try APIs. The legacy wrappers do not turn a failed probe into a
successful non-OOXML classification.

try_detect_ooxml_format_from_source_backed_package calls
package.check_execution() and package.source_version() before scanning,
periodically while scanning, and after scanning. It checks every 64 admitted
content-type entries, using the implementation's fixed
SOURCE_CLASSIFICATION_CHECK_INTERVAL. This preserves the existing execution
policy, including cancellation configured on the source-backed package, and
rejects a stale source. The classifier has no separate cancellation argument.
Positional source probes also retain their source-version checks; reader
probes have cursor restoration rather than filesystem freshness.

## Catalog-first ODF arbitration

ODF precedence is resolved with the existing bounded catalog probe before an
unrelated full OPC owner path is selected:

- flat ODF is handled by the existing flat-XML detector;
- for ZIP input, packaged_mime first recognizes the local stored ODF
  mimetype entry;
- packaged_has_ooxml_catalog_with_limits inspects only the central directory
  and does not read or decompress [Content_Types].xml or any other member;
- Some(false) means a canonical, in-budget ZIP layout with no reserved OPC
  content-types entry, so the package may take the ordinary ODF path;
- Some(true) means the canonical normalized catalog contains the reserved OPC
  content-types entry, so OOXML precedence remains eligible;
- None means the layout is uncertain or out of the cheap probe's contract,
  including noncanonical names, duplicates, ZIP64, encryption,
  data-descriptor entries, prefixed/trailing layouts, or a bound/allocation
  failure. It is not proof of ordinary ODF.

An ordinary ODF package therefore avoids an unrelated full OPC probe. A valid
OPC catalog wins over an ODF marker in a polyglot. An uncertain ODF catalog
continues through the existing bounded OOXML probe; if no OOXML family is
proven, the existing ODF preparation and typed owner validation may classify
it. Precedence is decided from structure, never by opening an ordinary
payload.

Native ODF MIME and central-directory catalog probes use their own finite ODF
limits, and ordinary ODF bypasses the full OPC path. An ODF-marked package
with a present or uncertain OOXML catalog enters the caller-bounded OPC path;
hard OPC input-size and content-types errors fail closed before ODF fallback
and are not converted into an ordinary ODF classification.

The positional and seekable ODF catalog helpers retain their own contracts:
the positional helper checks source version before and after central-directory
metadata reads and returns source I/O/change/allocation errors, while layout
uncertainty is an Ok(None) classification. The seekable helper restores its
reader position and exposes the established Option result.

## ODT, ODP, and wrong-family fixes

The source-backed ODT path recognizes the exact package MIME rather than a
filename suffix, so extensionless and incorrectly suffixed ODT files remain
eligible. It preserves the native ODF catalog tri-state: only Some(false)
owns the package directly as ordinary ODT without an unrelated OOXML owner
probe. Some(true), and None for an alias or otherwise uncertain catalog, enter
the caller-bounded OPC path. The source-backed OOXML classifier then
arbitrates the package while checking the same source version; an OOXML family
wins rather than silently remaining ODT.

When an uncertain ODT source reports missing OPC content-types, native ODT
fallback is returned only after a source-freshness recheck. An
OOXML-suffixed ODT path is matched case-insensitively and checks the same
source length against the caller input ceiling before reading MIME; a native
.odt path continues to use the ODF policy.

The ODP paths retain the corresponding behavior. Presentation bytes
pre-arbitrate ordinary ODP before the PPTX source-backed probe. A valid
ordinary ODP is allowed after that arbitration when the OOXML probe does not
identify an OOXML family. Invalid ODF body content proceeds to typed ODP
validation and is not accepted as PPTX. An OOXML/ODP polyglot with a disabled
or wrong OOXML owner does not fall through to ODP merely because the ODP
marker is present. A PPTX OPC ReadLimit is mapped through the facade to the
structured core ResourceLimit error rather than a generic parse or format
miss.

The public ODT regression covers both Document::from_bytes_with_limits and
Document::open_with_limits. An ordinary ODT succeeds at DOCX max_input=1,
while an ODT/DOCX polyglot returns the exact InputBytes error for both the
bytes and filesystem-path entry points.

When a source-backed path identifies an enabled but wrong OOXML family, its
private result is OtherOoxml(FileFormat). When the identified family owner is
disabled, it is DisabledOtherOoxml(FileFormat). The document, presentation,
and workbook facades map those results to their established wrong-family or
NotOfficeFile errors. This prevents a lower-precedence ODF owner, a different
semantic facade, or a filename suffix from taking the package after the
structural classifier has established its family.

These are private facade routing changes around the existing owners. They do
not add an eager DetectedFormat variant or a public generic owner handoff.
Byte-backed smart detection retains its old result. The non-Unix/Windows
fallback uses a fixed 512-byte MIME prefix and the native bounded
from-reader PackageCatalog probe; it does not require an unbounded payload
read. That fallback is statically reviewed only in this Linux validation.

## Acceptance evidence: zero ordinary-payload reads

Validation must use a counting ReadAt or reader wrapper and classify every
observed range. Central-directory, EOCD/ZIP64, [Content_Types].xml, and
relationship-member reads are approved structural reads. Any read or
decompression of an ordinary application part during catalog construction or
source-backed family classification is a failure.

The acceptance record must name the actual fixtures and paths exercised. It
must show an expected FileFormat, zero ordinary-payload reads, zero
ordinary-payload decompressions, and the configured limit/error outcome for
each exercised case. It should also show that an explicit later semantic
operation is the first operation allowed to read an ordinary part. Fixture
breadth must be reported honestly; no four-family or current-tree coverage
claim is made until those cases have actually been run.

## OOM-safe serial validation plan

Validation is planned as a serialized resource gate. Each validation root uses
one private disk-backed target and:

    CARGO_BUILD_JOBS=1
    CARGO_INCREMENTAL=0
    CARGO_PROFILE_DEV_DEBUG=0
    CARGO_PROFILE_TEST_DEBUG=0
    test harness: --test-threads=1

The planned roots are the focused litchi-opc catalog/probe checks, focused
facade OOXML probe and classifier checks, ODF arbitration checks, and the
ODT/ODP/wrong-family routing checks that are actually present in the current
tree. Roots run one at a time; no concurrent Cargo invocation, test runner,
or feature matrix is permitted. Captured diagnostics, fixture enumeration,
and any target scan have finite caps. A cancelled or OOM-killed child is
reaped before the next root.

The focused run used the disk-backed target
/home/zhuhe/CodeProjects/.cargo-targets/change-0344. Every root used
CARGO_BUILD_JOBS=1 and a single test thread, and roots were executed strictly
serially with no concurrent Cargo invocation or test runner. The target was
shared only in that serialized order.

If resource evidence is collected, it must be tied to the exact current tree
and named fixtures. Record fixture/input/catalog sizes, wall time, sampled
descendant high-water RSS with an available/partial/unavailable status, and
allocation count/bytes only when an approved allocator instrument is active.
None of those observations may be generalized into a benchmark claim from a
synthetic case.

## Validation status and evidence

Focused validation passed under the serialized protocol above. The recorded
breadth is limited to the named tests and checks; it is not a broad workspace
gate or a four-family benchmark claim:

- litchi-opc package_catalog_probe: 5/5 passed.
- litchi docx catalog_first_detection: 4/4 passed.
- litchi with pptx, odp, xlsx, and ods catalog_detection_arbitration: 8/8
  passed.
- litchi with docx and odt catalog_detection_arbitration: 1/1 passed,
  covering .odt success, .DOCX typed refusal, and exact/lowercase catalog
  polyglots.
- litchi with docx, pptx, xlsx, ods, and odp lib detection_smart: 16/16
  passed.
- The ods-only, odp-only, odt-only, odt+pptx, and docx+odt checks all passed
  warning-free after the configuration fixes.

Configuration isolation, including the feature=opc split, and ODF/OOXML
arbitration issues discovered during implementation were remediated; only the
final passing reruns above are validation results.

The counting-source zero ordinary-payload-read acceptance evidence is covered
only to the extent represented by the named focused checks; no broader fixture
or current-tree coverage is implied. The non-Unix/Windows fallback was
statically reviewed but was not runtime-compiled on this Linux run.

No benchmark, profiler, RSS measurement, or allocation measurement was run.
The validation above is correctness and configuration evidence only, and no
quantitative performance result is asserted.

## Known residuals

- PackageCatalog owns structural names, strings, relationship records, and
  classification vectors. It is bounded by ReadLimits, but it is not a
  zero-allocation or zero-RSS design.
- Structural relationship and content-type members are read as part of
  catalog admission. Zero payload reads does not mean zero package reads.
- Central-directory metadata and relationship graphs can be adversarially
  large; checked arithmetic, allocation errors, and limits remain required.
- A recognized content-type family can still fail semantic owner admission,
  encryption/protection/signature policy, or feature gating. The classifier
  does not decode, repair, decrypt, verify, or edit ordinary payloads.
- The cheap ODF catalog probe is deliberately conservative. None causes the
  existing bounded fallback rather than authorizing ordinary ODF precedence.
- Source-version checks use the source provider's supported freshness signals;
  they do not make a changing filesystem source kernel-transactional.
- The typed try APIs preserve errors, while compatibility Option wrappers
  flatten them. Callers needing diagnostics must use the typed path.
- Execution checks and cancellation are available through the existing
  source-backed package policy, not through a new public cancellation API.
- Private source-backed facade routing remains feature- and platform-gated;
  this change does not make every format owner source-backed or alter
  non-OOXML CRUD and publication behavior.

## Performance claim boundary

performance_claim: none

No latency, throughput, allocation, RSS, decompression, I/O, or OOM-freedom
improvement is claimed. Catalog locality, structural-only reads, and source
freshness/error behavior are implementation and acceptance properties. A
future record may add a performance claim only after current-tree benchmarks,
bounded RSS high-water evidence, and allocation evidence exist for named
fixtures under the serialized validation contract above.
