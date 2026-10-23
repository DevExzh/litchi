# 0362rd — substrate and ODF scope evidence

Status: read-only scope audit. This record contains no production API or test
change and no build or benchmark result. It records the boundary that a
coordinator can use to schedule the remaining work without treating a shared
crate as an unexamined exception.

The original survey was made from the 0362 integration tree. A source-only
refresh at assembled commit `94123c0efb` is reflected below where later
commits closed a named row; no fresh rustdoc or build result is implied.

The source survey was made from the 0362 integration tree. Consumer counts
below are `rg -l` planning counts over tracked `crates/` sources; they include
definitions and tests where the expression matches, so they are not a claim
about unique call sites. The exact call sites named for the ODF raw API were
checked separately.

## Rules that determine scope

| Evidence | Explicit rule | Consequence for this audit |
|---|---|---|
| `docs/unified-ops/0351-api-contract.md`, §2 opening and lines 119–122 | “Every rule applies to every format crate and to the unified `File`.” | A public API in an owned format family cannot be excused by moving it into a `common` module. |
| 0351 §2.1, items 1, 4 and 6 (lines 124–151) | Owners are immutable; mutators use scoped edits; a `raw` module may retain a mutable codec model under a non-owner noun. | A raw Reader or writer may remain mutable, but a format owner or a free format helper must not publish a second owner mutation path. |
| 0351 §2.5, items 1–3 (lines 291–314) | Owners carry `OpenOptions`; there is no public `*_with_limits`; raw, record, detect and codec APIs require a limits argument and have no limit-less twin. | A public raw constructor with a default limit is still debt even when its crate is shared. The limit type must belong to the format-facing API. |
| 0351 §2.6, items 1–3 (lines 346–363) | Every wave migrates every consumer in the same commits, removes aliases, preserves fidelity, and measures performance with ABBA. | A substrate lane must enumerate format crates, facade/ops, tests, examples, fuzz, perf-baseline and native-resave consumers before changing names. |
| ADR 0034 §1 (`docs/adr/0034-unified-api-glossary-and-default-policies.md`, lines 83–89) | One public name per concept; a renamed API is removed rather than retained as an alias. | The aliases in `litchi-odf-common`, `litchi-crypto` and `litchi-ole-common` are contract debt even when they are harmless forwarders. |
| 0351 §0 wave table (lines 28–38) | W5 explicitly owns `litchi-ole-common` and `litchi-crypto`; W6 explicitly owns `litchi-odf-common` and the ODF leaves. `litchi-opc`, `litchi-cfb` and `soapberry-zip` are not listed. | The ODF and legacy aliases are in existing format waves. The three physical substrate crates need explicit follow-up lanes, not an implicit exemption and not an unplanned expansion of W4/W5/W6. |

The absence of a substrate crate from the wave table is a scheduling boundary,
not a claim that its APIs are permanently outside the engineering goal. The
crate descriptions also support the distinction: `litchi-cfb` describes a CFB
container substrate consumed by DOC/PPT/XLS and encrypted OOXML
(`crates/litchi-cfb/src/lib.rs:1-5`), while `soapberry-zip` describes a
high-performance ZIP reader for OOXML, ODF and iWork
(`crates/soapberry-zip/Cargo.toml:2-6`).

### Scope qualification against 0350, 0351 and 0343 D5

0350 §1 (`docs/unified-ops/0350-owner-api-debt-decision.md:8-17`) makes the
required experience explicit: every **format crate** and the unified `File`
facade, with every affected consumer migrated in the same change. It does not
assign every workspace implementation crate to a current wave. 0351 §1
describes the inventory method as rustdoc JSON for 25 format/shared crates and
`litchi-ops` (`docs/unified-ops/0351-api-contract.md:45-60`), and its wave table
lists the owned paths (`:28-38`). That inventory scope is evidence about what
the current format contract measures; it is not permission to hide a public
format re-export, and it is not evidence that every standalone transport method
was accepted as current-format debt.

The original 0343 D5 is equally specific. Its closure record describes PPT
`raw::Reader` and 173 record/family `*_with_limits` entry points
(`docs/unified-ops/0343-glossary-closure.md:35-40,115-118`), and 0350 overrides
that D5 default with the canonical limits outcome. D5 therefore supplies a
textual basis for the PPT raw migration; it does not, by itself, authorize a
rewrite of every CFB or ZIP primitive used by PPT.

The resulting classification is:

| Surface | Textual basis | Status |
|---|---|---|
| Public format-facing re-export, alias or leak | 0350’s every-format requirement; 0351 W5/W6 owned paths; ADR 0034 §1’s one-name/no-alias rule. Examples are the W5 `Revision`/`Snapshot` aliases and the W6 writer’s public ZIP limit/error types. | Required in the owning format wave. The three W5 removals in this branch are this category. A public ODF path cannot be excused as “generic ZIP”; the concrete ODF repair must still distinguish a canonical re-export from a migration alias. |
| Office-specific OPC surface | ADR 0034 §5 explicitly names OPC `ReadLimits` → `SourceLimits` and says OOXML re-exports follow (`docs/adr/0034-unified-api-glossary-and-default-policies.md:169-181`); 0343’s closure also records OPC stream-verb renames. The W4 brief fixes the raw `Package::edit_opc` boundary. | The required source/consumer canonicalization is integrated by 0362rl across DOCX/PPTX/XLSX/XLSB and raw consumers. Its absence from the 0351 wave table remains a qualification boundary: rerun the profile and native checks before calling the substrate lane fully gated. |
| Standalone CFB or soapberry-ZIP methods with no public format operation or free foreign-result signature | 0350 names format crates and `File`; 0351’s inventory and ownership table do not include these implementation crates. ADR 0034’s vocabulary policy does not assign their independent transport surface to W5/W6. A typed foreign causal source nested in a crate-owned error is a separate, permitted boundary shape. | Separate scope requiring an explicit owner/coordinator lane. Do not claim the entire generic API is mandatory in the current format wave solely because formats depend on it. Any method re-exported through an owned format surface moves back into the first row and is required there. |

This distinction is narrower than a blanket exemption. For example,
`litchi-ole-common::source_backed_overlay` publicly accepts and returns
`SharedOleFile`, and ODF publicly aliases a ZIP limits type; those are
format-facing paths to review under W5/W6. Conversely, a `soapberry-zip`
reader or a CFB writer method that is only consumed through private codec
code has no direct 0350/0351 current-wave assignment. The latter remains an
explicit follow-up candidate with its consumer and performance inventory,
not an “accepted” API exception.

## Confirmed debt in the owned ODF family (W6)

### `SourceMemberReaderError` aliases removed in 0362rda

Before 0362rda, `crates/litchi-odf-common/src/core/package.rs:123-191`
exposed the canonical accessors `core`, `callback`, `into_core` and
`into_callback`, plus these six methods whose documentation called them
aliases:

| Alias | Canonical method |
|---|---|
| `core_error` | `core` |
| `error` | `core` |
| `callback_error` | `callback` |
| `into_core_error` | `into_core` |
| `into_error` | `into_core` |
| `into_callback_error` | `into_callback` |

The type is re-exported from `litchi-odf-common` (`src/lib.rs:35-37`) and was
therefore part of every ODF leaf's shared public API. A workspace search found
no consumer call to the six methods; the existing source-member test matches
the canonical enum variants. 0362rda removes the six methods, keeps the four
canonical accessors, and registers the exact-method bans in the ODF common API
profile. This is an ODF change, not a generic ZIP exception.

### The public ODF archive codec has an unbounded constructor path

`crates/litchi-odf-common/src/package/mod.rs:8-30` keeps `codec` private but
re-exports `Archive`, `ArchiveLimits`, `parse_manifest` and `read_manifest`.
Consequently these are public raw package APIs at
`litchi_odf_common::package::*`:

* `Archive::new(data)` (`src/package/codec.rs:120-139`) silently supplies
  `ArchiveLimits::default()` while the explicit implementation is only
  `pub(crate) new_with_limits`.
* `Archive::read_manifest_xml()` (`src/package/codec.rs:221-247`) silently
  supplies a fixed 16 MiB manifest ceiling, while the caller-selected method
  is `read_manifest_xml_with_limit`.
* `read_manifest(&Archive)` (`src/package/codec.rs:280-287`) and
  `parse_manifest(&str)` (`src/package/codec.rs:584-605`) have no caller
  limit parameter. They are public codec/record entry points, not package-owner
  methods.

The direct `Archive::new` consumers are six call sites in five files:

* `crates/litchi-odf-common/tests/parser_hardening.rs:37`;
* `crates/litchi-odf-common/src/package/tests.rs:13,23`;
* `crates/litchi-odf-common/src/core/package.rs:2235`; and
* `crates/litchi-ops/src/raw/package.rs:67,105`.

`parse_manifest`/`read_manifest` are used by the ODF common manifest code,
package tests and the public raw package facade. The raw lane should make the
caller-selected archive/manifest limits explicit, with one of these designs
chosen before implementation:

1. extend the format-owned `ArchiveLimits` to include the manifest ceiling and
   pass it through `Archive::new`, `read_manifest` and `parse_manifest`; or
2. keep archive transport limits and manifest parse limits as separate
   required arguments, both owned by ODF.

Either design must remove the default raw constructor and avoid a public
limit-less twin. The bounded package owners may continue to obtain their
limits from `OpenOptions`; this finding concerns the public raw/codec routes.
`read_manifest_xml_with_limit` is a singular-name API, so it is not caught by
the literal suffix check, but it should be folded into the chosen canonical
raw limits shape rather than left as a second spelling.

### ODF writer limits: canonical re-export versus an owned wrapper

Before 0362rda, `crates/litchi-odf-common/src/core/writer.rs:222-227`
publicly defined

```rust
pub use soapberry_zip::office::StreamingArchiveLimits as PackageWriterLimits;
```

That is a public re-export with a second name, but the contract does not by
itself require an ODF-owned wrapper. 0351 §2.5.1 explicitly permits a shared
canonical `OpenOptions` re-export (`:293-300`), and §2.5.3 permits a raw or
codec API to take its module's limits as the required argument (`:308-314`).
ADR 0034 §1 requires one name per concept and removes renamed migration aliases
(`docs/adr/0034-unified-api-glossary-and-default-policies.md:83-89`); that is
different from choosing one canonical public re-export. `StreamingArchiveLimits`
is the actual public transport type (`crates/soapberry-zip/src/office.rs:6238-6272`),
and its documented dimensions are exactly the finite ZIP ceilings described by
the ODF re-export (`writer.rs:222-227`). It is `Copy`, so an owned newtype would
not allocate, but it would add field conversions, builders/accessors and another
public vocabulary without adding an ODF dimension.

The zero-cost canonical choice landed in 0362rda: the ODF core module now
re-exports `StreamingArchiveLimits` without renaming it, and every ODF
consumer uses that one name. This is the same kind of canonical sharing that
0351 permits for ODF `OpenOptions`; it is not a retained compatibility alias.
An ODF-owned wrapper remains an option only if the owner deliberately adds
semantics that the transport type cannot express. Before the migration,
production streaming consumers were in the ODF common writer and in
`litchi-odt/src/streaming.rs`, `litchi-ods/src/streaming.rs` and
`litchi-odp/src/streaming.rs`; the search found ten source/test files naming
`PackageWriterLimits` and fifteen files naming `with_writer_and_limits`.

There is a separate, factual foreign-error leak. The public ODF
`PackageWriterError` has `Archive(soapberry_zip::Error)` and
`ArchiveFailure(StreamingArchiveFailure)` variants (`writer.rs:314-330`), and
the latter exposes the foreign error, progress and transport-limit reports
through its public methods (`soapberry-zip/src/office.rs:6536-6568`). The ODF
`PackageWriterLimitResource`/`PackageWriterLimitExceeded` types and the mapping
from `StreamingLimitExceeded` are ODF-owned (`writer.rs:238-300,2028-2047`),
so the limit carrier and the error-source leak are distinct decisions. ADR
0034 §6 permits typed boundary variants (`:198-206`), but does not prove that
foreign payload types must be copied; if the coordinator requires a strict
format-owned error payload, that mapping is its own change and may preserve the
foreign error only as a private causal source. A limits wrapper would not solve
that leak.

Finally, `PackageWriter::with_writer` and `new` default the transport limits
(`writer.rs:723-780`). Whether those defaults are acceptable depends on whether
this output/builder surface is classified as a raw untrusted-data ingress under
0351 §2.5.3; that policy question is independent of the limit type and should
not be used to justify an owned wrapper. The lane must measure streamed output
and memory, including the ODF streaming profiles, after selecting the canonical
name and error boundary.

## W5 format-domain aliases (closed by the assembled tree)

The original survey identified two direct format-domain alias families. The
subsequent W5 commits `953f9720bc` and `c905a7d63d` removed them and migrated
their consumers: current `litchi-crypto::spaces::Revision` and
`litchi-ole-common::ole1::Revision` expose only the canonical `value`, and the
former `ole1::Snapshot` alias is gone. No generic CFB or ZIP exemption was
used. This section therefore supplies historical scope evidence, not an open
W5 alias exception; a fresh rustdoc inventory is still required before
claiming the whole substrate profile passed.

## Shared substrate classification

### `litchi-opc`: dedicated Office-package substrate lane

`litchi-opc` is Office-specific physical packaging, so it is closer to a
format substrate than to generic ZIP. It is nevertheless absent from the
0351 wave ownership table. The W4 common brief gives the immediate boundary:
`litchi-xlsx`/`litchi-xlsb` may not publish free public functions that mutate a
raw `OpcPackage`; their raw escape is `Package::edit_opc`, while
`litchi-ooxml-common` may stage against `&OpcPackage` and commit a new package
(`.coord/w4-xlsx/brief-common.md:82-95`). ADR 0034 also explicitly names OPC's
`ReadLimits` to `SourceLimits` rename (`docs/adr/0034...:169-181`).

The original 0362 survey recorded 31 public declarations with a
`*_with_limits` suffix. A source-only audit of assembled commit `94123c0efb`
finds no public OPC declaration with that suffix; the remaining occurrences
are private or test helpers. The old list was:

* `phys_pkg::{open,from_bytes,from_reader,new}_with_limits`;
* `package::{open,from_reader,from_vec,from_shared_vec,from_bytes}_with_limits`
  and content-type replacement/source methods;
* `source_backed::{from_path,from_reader,from_vec,from_read_at}_with_limits`
  plus four longer combinations with cache/policy/execution context;
* package-catalog probing, relationship parsing/replacement and validation
  `_with_limits` methods.

The corresponding limit-less twins were public in the surveyed tree. A source
survey finds direct `litchi_opc::` users in 182 DOCX, 331 PPTX, 281 XLSX, 126
XLSB and 51 ops source files (the counts include tests and definitions). The
perf-baseline source has 98 `litchi_opc::` lines in 42 files and 204
`OpcPackage` lines in 25 files. These are enough consumers that changing the
OPC surface belongs in an atomic substrate wave with all four OOXML families,
ops, facade/raw callers, tests, examples, fuzz and perf-baseline.

The OPC alias debt was also rechecked separately from the old limit migration:

* `VerifiedDecodedReaderError` now exposes the canonical `opc`, `callback`,
  `into_opc` and `into_callback` accessors; the old alias spellings are gone.
* `OpcOperationAccounting` now exposes the canonical payload/accounting names;
  the old duplicate spellings are gone.

The survey found no public OPC operation or free `Result` signature whose
return or argument type is a `soapberry_zip` transport type. The public,
crate-owned `OpcError` does retain typed causal sources: `ParallelRead` carries
`#[source] soapberry_zip::Error` (`crates/litchi-opc/src/error.rs:166-168`),
while `XmlPublication` carries `xml_minifier::audit::Error` and
`QuickXmlError` carries `quick_xml::Error` (`error.rs:145-152,304-305`). The
private `ArchiveAccess` and worker helpers also use foreign `Result` types, but
they do not enter the OPC public surface. Under the settled error reading, a
foreign causal payload inside the crate-owned `OpcError` is a typed source, not
a free transport API leak; it requires no wrapper or string-flattening rewrite.
Thus the old dedicated OPC limit/alias findings are closed in the assembled
source; a fresh rustdoc inventory is still required before claiming the whole
OPC profile passed. A blanket ZIP API rewrite is not justified by these
retained causes.

### `litchi-cfb`: dedicated CFB substrate lane

The CFB crate publicly re-exports both default and explicit-limit ingress:

* `OleFile::open` delegates to `open_with_limits` at
  `crates/litchi-cfb/src/file.rs:1034-1055`;
* `SharedOleFile::{open,open_owned,open_owned_vec}` delegate to their
  `*_with_limits` variants at `src/shared.rs:620-684`;
* `SharedOleFile::open_with_limits` and
  `validate_source_with_limits` are public (`src/shared.rs:770-785`,
  `src/validation.rs:77-95`);
* writer-side `SequentialWriterOptions::with_limits` and
  `OleWriter::move_storage_with_limits` are also public.

The public alias `OleWriter::update_stream` is explicitly documented as an
alias for `create_stream` (`src/writer/core.rs:880-895`). Its only external
source use found in the current tree is the CFB writer test, but deleting it
still requires the test migration and API gate update.

This is not an owner exception: DOC, PPT and XLS raw Readers may keep a
mutable codec model under the 0351 raw rule, and immutable owner projections
that return `SharedOleFile` are not automatically `&mut self` violations. It
does mean that all public CFB ingress and writer aliases need a named CFB
lane. Planning counts are 213 files for `OleFile::open(`, 27 for
`OleFile::open_with_limits(`, 23 for `SharedOleFile::open(` and 15 for
`SharedOleFile::open_with_limits(` across `crates/`; the lane must include
`litchi-doc`, `litchi-ppt`, `litchi-xls`, `litchi-ole-common`, `litchi-crypto`,
`litchi-ograph`, encrypted OOXML integration, ops, tests and perf-baseline.

The CFB lane should canonicalize required caller limits for raw ingress,
remove the public suffix/default twins, remove `update_stream`, and preserve
the source-freshness and writer behavior with ABBA measurements. It should
not be folded into W3 merely because PPT uses the crate: 0351 explicitly
assigned W3 `litchi-ppt` and excluded `litchi-ole-common`, `litchi-ograph` and
`litchi-crypto`.

### `soapberry-zip`: generic physical transport lane, with one immediate ODF leak

The generic ZIP crate is absent from the wave table and its Cargo description
calls it a ZIP reader shared by OOXML, ODF and iWork. The assembled source
still has twelve public `*_with_limits`/`with_limits` declarations, including
archive reader/indexer constructors, preservation indexing, lazy reader
constructors and `StreamingArchiveWriter::with_limits`
(`crates/soapberry-zip/src/office.rs`), plus
`PreservationIndex::new_with_limits` (`src/preserve.rs:786`). It also has
explicit alias debt:

* `VerifiedEntryReaderError`'s `archive_error`, `transport_error`,
  `callback_error` and `into_callback_error` (`src/office.rs:2118-2147`);
* four reader-oriented `StreamingArchiveWriter` aliases
  (`src/office.rs:8283-8308`);
* three `LazyArchiveCacheLimits` aliases (`src/office.rs:8472-8545`);
* three `ZipOperationAccounting` aliases (`src/accounting.rs:110-122`);
* `start_entry_owned` and `start_entry_owned_zip64` aliases
  (`src/writer.rs:1958-2000`); and
* the public `ErrorKind::Io` compatibility alias for `ErrorKind::IO`
  (`src/errors.rs:91-161`).

The generic transport surface appears in 264 source files through
`soapberry_zip::`, with 198 `StreamingArchiveWriter`, 117 `ArchiveReader` and
29 `PreservationIndex` files. The corresponding perf-baseline source has 97
`soapberry_zip` lines in 56 files. Those counts justify a dedicated ZIP lane
with all consumers and streaming/deflate ABBA measurements; they do not
justify changing generic APIs opportunistically during an ODF or OPC owner
wave.

The public-leak audit found no public free `soapberry_zip` operation or result
signature in OPC; it did find the typed causal sources nested in the public
`OpcError` described above. ODF source-publication uses of `PreservationIndex`
are `pub(crate)`. The public ODF writer re-exports the ZIP limit and error types
above, so W6 must make the canonical-name and error-boundary decisions even
though the remaining ZIP surface waits for its own lane. Internal calls such as
`litchi-odf-common/src/core/writer.rs:1309` using
`write_deflated_reader` belong to that eventual generic alias migration but
do not by themselves expand the current ODF owner scope.

## Atomic lane proposal and required evidence

1. **W6 ODF raw/alias lane:** 0362rda removes the six source-member aliases,
   canonicalizes the writer limit re-export and migrates every named ODF
   consumer. The remaining archive/manifest/raw-writer work is assigned to
   `odf_raw_contract`: make limits caller-selected where the public raw/codec
   contract requires it, decide the writer error boundary and default policy,
   then run ODF round-trip/parser-hardening, streaming and facade tests and
   measure memory and output performance.
2. **W5 domain alias lane:** closed by `953f9720bc` and `c905a7d63d`; the
   current source has no `Revision::fingerprint` or `ole1::Snapshot` alias.
3. **W4 format boundary:** keep the W4 change limited to public XLSX/XLSB
   routes and the shared OOXML staging rule. Any public `&mut OpcPackage`
   helper outside the raw `edit_opc`/codec boundary is W4 debt and needs a
   Package route in the same commit.
4. **OPC substrate lane:** source/consumer canonicalization is closed by
   `0362rl` in the assembled tree: limits and aliases were migrated across
   DOCX/PPTX/XLSX/XLSB, ops and raw/perf/test consumers. A fresh rustdoc
   profile and native qualification are still required before calling the
   substrate lane fully gated; preserve source-backed cache, splice and
   publication performance while doing so.
5. **CFB substrate follow-up (separate scope):** the current source still has
   public default/`*_with_limits` ingress, validation and writer methods plus
   `OleWriter::update_stream` (`file.rs:1055`, `shared.rs:650,680,780`,
   `validation.rs:91`, `writer/core.rs:1297`, `writer/sequential.rs:197`).
   This is not current format-wave debt merely because DOC/PPT/XLS depend on
   CFB; a future owner must decide its scope and migrate consumers together.
6. **soapberry ZIP follow-up (separate scope):** the current source still has
   public office transport `*_with_limits` methods in `office.rs` and
   `preserve.rs:786`, plus its documented aliases. This is not a blanket
   current format requirement; a public format re-export or foreign surface
   remains required in that owning format lane.

Before each lane is merged, rerun `tools/api_debt_inventory.py` and the API
glossary gate over the full workspace, then show that all listed consumers
have moved. The present record deliberately makes no claim that the generic
substrate lanes are complete; it supplies the textual basis and concrete
surface lists needed to schedule them without a blanket exemption.
