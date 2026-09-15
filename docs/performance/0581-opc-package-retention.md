# 0581: the eager OPC package's cost is in the open, not the save — and the ordinary DOCX/XLSX/PPTX door is the eager one

Status: design only. No production change and `performance_claim: none` — this
record carries deterministic peak-retained-byte and allocation counts only. **No
timing is measured and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

Change [0578](0578-zip-passthrough-is-already-bounded.md) set out to prove that
compressed-entry passthrough on save copies whole members into memory, refuted
that, and in refuting it measured something larger which it did not explore:

> **Path A holds its whole source in memory.** `pkgwriter.rs:197` opens the
> preservation source with `ZipArchive::from_slice(source)`, so a
> `PackageWriter::write` save has the entire source archive resident as `&[u8]`
> before the copy loop starts. Axis 5 prices it: **135,268,336 peak bytes against
> Path B's 532,626 for byte-identical output** […] It is recorded as the largest
> measured opportunity this investigation found, and deliberately not attempted
> here.

This record takes that up. It verifies the 254× figure independently, **corrects
its attribution**, establishes from source what retains the bytes and why, and
answers the question 0578 did not ask: which public entry points reach the
expensive path.

Three results, in order of how much they change the picture:

1. **The save is not where the payload bytes are.** Path A's publish region is
   **flat at 514,544 peak bytes across a 2,048-fold growth in the media member**
   — independent of payload size — and `pkgwriter.rs:197` takes a *borrow*, not a
   copy. The whole of 0578's difference is `OpcPackage::open`'s **retention**, and
   the two figures decompose to the byte. The publish is **not** constant in
   general: it is linear in *member count*, reaching 7,003,258 bytes at 8,000
   members where the source-backed publish stays flat at 88,841. That is a
   separate, smaller finding, recorded in axis 2 rather than folded into this one.
2. **"254×" is one point on an unbounded curve.** It is 2.2× at 64 KiB, 254× at
   64 MiB and **506× at 128 MiB**, because Path A retains the compressed archive
   *plus the sum of every decompressed part payload*. On a real XML-heavy XLSX
   that is **11.19× the archive** and a 97.6× ratio.
3. **The ordinary documented door is the expensive one.** DOCX, XLSX, PPTX and
   XLSB all reach it from `Package::open` … `save`, and **every update-then-save
   example in `docs/OFFICE_API_GUIDE.md` is that path**. It is not legacy, niche,
   or test-only. This also confirms two standing `docs/GOAL.md:327-331` hypotheses.

The decision is **not to implement**. The retention that matters is held behind a
public, **infallible, borrowing** accessor — `Part::blob(&self) -> &[u8]`, with
**950 production call sites across seven crates** — and removing it would move a
typed limit refusal out of `open()`, where it is reportable, into an accessor that
has no error channel. That is an observable-contract change, so this record
freezes the design instead, in the shape of
[0566](0566-xls-worksheet-window-design.md) and
[0577](0577-ooxml-open-relationship-parts.md).

## What was changed

**Nothing.** No production code, test, or fixture was modified. The only files
added are this record and its evidence directory.

## Stage 1: the figure reproduces exactly

Change 0578's axis-5 probe was re-run unmodified from
[`results/change-0578/opc-save-peak-probe.rs`](results/change-0578), over a
detached git worktree of `32d25e08806d93f792ffd4954d83acc9db9c5301` with an
isolated `CARGO_TARGET_DIR`. Every figure reproduced **byte for byte**, including
the output digests — see
[`results/change-0581/stage1-repro-0578-axis5.csv`](results/change-0581/stage1-repro-0578-axis5.csv)
against 0578's retained `stage1-peak-e2e-opc.csv`. The 254× is real and is not a
measurement artefact.

## The attribution is wrong: the save is already bounded

0578 measured one region spanning open *and* save. Splitting it into an open
region and a publish region, and additionally reporting what each region's
product still **retains** afterwards, relocates the entire cost.

Path A is `OpcPackage::open` → `get_part_mut().set_blob()` →
`PackageWriter::write_to_stream`; Path B is `SourceBackedPackage::from_path` →
`write_part_overlay_to_stream`. Both publish to a sequential, non-seek sink.
Fixtures are change 0578's: one real DOCX plus one large and one small binary
media part, with the **small** part edited.

| media member | archive | **A open retained** | **A save peak** | A open + A save | 0578's single-region figure | match |
| ---: | ---: | ---: | ---: | ---: | ---: | :---: |
| 64 KiB | 164,712 | 646,673 | **514,544** | 1,161,217 | 1,161,217 | **yes** |
| 1 MiB | 1,148,053 | 2,613,054 | **514,544** | 3,127,598 | 3,127,598 | **yes** |
| 4 MiB | 4,294,743 | 8,905,472 | **514,544** | 9,420,016 | 9,420,016 | **yes** |
| 16 MiB | 16,881,495 | 34,075,136 | **514,544** | 34,589,680 | 34,589,680 | **yes** |
| 64 MiB | 67,228,503 | 134,753,792 | **514,544** | 135,268,336 | 135,268,336 | **yes** |
| 128 MiB | 134,357,847 | 268,992,000 | **514,544** | 269,506,544 | not measured by 0578 | — |

**The publish region is flat to the byte at 514,544 peak and 448 allocations at
every size, across a 2,048-fold growth in the media member**, and `A open retained
+ A save peak` reproduces 0578's number exactly at all five of its points. Along
this axis the save contributes a constant, and everything else is residency
established before the writer is called. The constant is only constant *in payload
size*: axis 2 shows the same region growing linearly in member count, so the
correct statement is that the publish is independent of how large the members are,
not that it is independent of the package.

The mechanism is visible in source. `pkgwriter.rs:197`'s argument is a borrow of
bytes the package already owns, not a new copy:

```rust
// package.rs:1147
pub(crate) fn preservation_source(&self) -> Option<(&[u8], &PreservationProvenance)> {
    self.source_archive
        .as_deref()          // Option<Arc<Vec<u8>>> -> Option<&Vec<u8>>
        .map(Vec::as_slice)  // -> Option<&[u8]>
        .zip(self.preservation.as_deref())
}
```

A second, separate fast path streams that same borrow in bounded chunks when the
package is still exact-source authorized (`EXACT_SOURCE_CHUNK_BYTES` = 64 KiB,
`pkgwriter.rs:18`). **This path does not produce the 514,544 figure above** — the
axis-1 scenario edits a part, which revokes the authorization, so
`package.exact_source()` returns `None` and the flat peak comes from the
preservation/`PublicationPlan` path instead. It is quoted here because it is what
produces axis 3's *zero*-allocation unedited saves:

```rust
// pkgwriter.rs:922
fn write_counted<W: Write>(writer: W, package: &OpcPackage) -> Result<()> {
    if let Some(source) = package.exact_source() {
        let mut writer = writer;
        for chunk in source.chunks(EXACT_SOURCE_CHUNK_BYTES) {
            writer.write_all(chunk)?;
        }
```

`PackageWriter` itself is `pub struct PackageWriter;` (`pkgwriter.rs:872`) — a
unit struct with no fields — and `PublicationPlan`'s per-part entry holds
`blob: &'package [u8]` (`pkgwriter.rs:67`), a borrow of the part's `Arc`. **0578's
sentence naming `ZipArchive::from_slice` as the thing that "holds its whole source
in memory" is the wrong half of its own finding**; the clause it wrote immediately
afterwards — "the residency is `OpcPackage`'s eager ownership model" — is the
correct one, and this record is the evidence for it.

## What actually retains the bytes

`OpcPackage` (`package.rs:106`) has four byte-bearing fields:

```rust
pub struct OpcPackage {
    /// All parts in the package, indexed by partname
    parts: HashMap<PackURI, Box<dyn Part + Send + Sync>>,          // package.rs:113
    /// Exact XML payloads materialized from the opened source package.
    source_xml_parts: HashMap<PackURI, Arc<Vec<u8>>>,              // package.rs:116
    /// Owned source archive retained for exact and targeted publication.
    source_archive: Option<Arc<Vec<u8>>>,                          // package.rs:119
    /// Source identity used to prove safe targeted publication.
    preservation: Option<Arc<PreservationProvenance>>,             // package.rs:153
    ...
}
```

Two of the four are **`Arc` aliases, not copies**, and cost no extra bytes at
open: `source_xml_parts` stores `part.blob_arc()` (`package.rs:501`), and
`PreservationProvenance::SourcePart::blob` is likewise an `Arc` snapshot
(`package.rs:65`). The measured residency below is consistent with this — it lands
between 2.00× and 11.19× the archive, tracking compressibility, rather than at the
3× or 4× a genuine second copy of every XML part would produce.

The two that do cost are:

**1. The whole compressed source archive.** Every owned-ingress constructor —
`open`, `open_with_limits`, `from_reader*`, `from_vec*` — funnels through one
function that slurps the file and then *keeps* the buffer:

```rust
// package.rs:1185
fn from_owned_bytes_with_limits(data: Vec<u8>, limits: ReadLimits) -> Result<Self> {
    let mut package = { ... Self::unmarshal(pkg_reader)? };
    package.authorize_owned_source(data);   // the Vec<u8> is moved in, not dropped
    Ok(package)
}

// package.rs:1224
fn authorize_owned_source(&mut self, source: Vec<u8>) {
    let source = Arc::new(source);
    self.preservation = PreservationProvenance::from_package(source.as_slice(), self).map(Arc::new);
    self.source_archive = Some(source);
    self.exact_source_authorized = true;
}
```

**2. Every part's decompressed payload, eagerly, at open.**
`PackageReader::load_parts_eager` (`pkgreader.rs:852`) decompresses *all* admitted
parts in one bulk call before any part object exists. `from_phys_reader`'s own doc
comment states it (`pkgreader.rs:530`): *"Uses the eager payload path: … 3.
Decompress all admitted Part payloads for the owning package"*. There is no lazy,
`OnceCell`, or deferred variant anywhere in this path, and `OpcPackage` has **no
by-value method at all** — no `into_`, `take_`, blob-eviction or source-dropping
affordance a caller could use to shed either one.

So at save time **the compressed archive and the complete decompressed part set
are simultaneously live**, which is exactly what the measurement shows.

## What the retention scales with

This is the part 0578's single 254× figure obscures. Three axes below. `A total`
is `A open retained + A save peak` and `B total` is `B open retained + B save
peak`, the same composition on both sides; `A retained / archive` uses the open's
retention alone. Full data, including every region's separate peak, allocation
count and retention, is in [`results/change-0581/`](results/change-0581).

### Axis 1 — largest member varies, member count fixed

| media member | archive | A total | B total | **A/B** | **A retained / archive** |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 64 KiB | 164,712 | 1,161,217 | 532,599 | 2.2× | 3.93× |
| 256 KiB | 361,380 | 1,554,493 | 532,599 | 2.9× | 2.88× |
| 1 MiB | 1,148,053 | 3,127,598 | 532,599 | 5.9× | 2.28× |
| 4 MiB | 4,294,743 | 9,420,016 | 532,599 | 17.7× | 2.07× |
| 16 MiB | 16,881,495 | 34,589,680 | 532,599 | 64.9× | 2.02× |
| 64 MiB | 67,228,503 | 135,268,336 | 532,599 | **254.0×** | 2.00× |
| 128 MiB | 134,357,847 | 269,506,544 | 532,599 | **506.0×** | 2.00× |

**B is flat to the byte at every size. The A/B ratio is not a constant — it
doubles with every doubling of the member**, because B's denominator does not
move. Quoting "254×" as a property of the two paths is therefore misleading; the
property is that A is O(package content) and B is O(index).

### Axis 2 — member count varies, member size fixed at 4,096 bytes

| members | archive | A total | B total | A/B | A retained / archive | **A save peak** | **B save peak** |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 99,008 | 610,008 | 114,266 | 5.3× | 5.20× | 95,034 | 88,841 |
| 20 | 184,119 | 805,043 | 129,526 | 6.2× | 3.78× | 109,498 | 88,841 |
| 100 | 524,512 | 1,607,148 | 197,926 | 8.1× | 2.72× | 181,050 | 88,841 |
| 500 | 2,226,536 | 5,576,260 | 517,846 | 10.8× | 2.27× | 522,298 | 88,841 |
| 2,000 | 8,608,935 | 20,501,711 | 1,739,626 | 11.8× | 2.17× | 1,818,490 | 88,841 |
| 8,000 | 34,138,114 | 80,203,098 | 6,690,906 | **12.0×** | 2.14× | **7,003,258** | **88,841** |

Here the ratio **plateaus near 12×** rather than growing, because B's index cost
is itself linear in member count (change 0578's axis 2 measured the same shape at
the ZIP layer). Member count is a cost both paths pay; payload bytes are a cost
only A pays.

The two save columns are carried here explicitly because they qualify this
record's first finding. **Path A's publish region is flat in payload size but
linear in member count** — 95,034 bytes at zero added members rising 73.7-fold to
7,003,258 at 8,000, with allocations rising 433 to 64,452 — while **Path B's
publish stays flat at 88,841 bytes throughout**. So "the save is bounded" is true
of the axis 0578 measured and false of this one, and on this axis the eager
publish is genuinely worse than the source-backed one rather than merely
differently sized. It remains the smaller term: at 8,000 members the publish peak
is 7,003,258 against the open's 73,199,840 retained — 8.7% of the 80,203,098
total.
Attributing that growth is outside this record's scope; it is a `PublicationPlan`
question, not a retention one.

### Axis 3 — real corpus, unedited open-and-save

| fixture | archive | A total | B total | A/B | **A retained / archive** |
| --- | ---: | ---: | ---: | ---: | ---: |
| `ArtisticEffectSample.pptx` | 972,788 | 2,355,882 | 133,015 | 17.7× | 2.42× |
| `saut_page.docx` | 2,959,626 | 8,352,590 | 98,684 | 84.6× | 2.82× |
| `no_drawing_patriarch.xlsx` | 672,414 | 7,525,438 | 77,081 | **97.6×** | **11.19×** |
| `ConditionalFormattingSamples.xlsx` | 654,688 | 1,836,295 | 216,384 | 8.5× | 2.80× |
| `EmbeddedVideo.pptx` | 201,418 | 504,504 | 110,423 | 4.6× | 2.50× |
| `drawing.docx` | 95,247 | 506,648 | 90,382 | 5.6× | 5.32× |

On these unedited saves **Path A's publish region allocates nothing at all** —
zero peak, zero allocations, at every one of the six fixtures — because
`write_counted`'s exact-source fast path streams the retained archive straight to
the sink in 64 KiB chunks. `A total` is therefore exactly `A open retained` here.
That is the cleanest statement of this record's first finding: on the scenario the
DEFINITION OF DONE names, the eager path's save is already free, and every byte it
costs was spent before the writer was called.

`no_drawing_patriarch.xlsx` is the instructive one: not the largest archive
(672,414, fourth of six) nor the largest retention (7,525,438, second to
`saut_page.docx`), but by far the largest **inflation** — 11.19× its archive
against 2.42–5.32× for the rest — because highly compressible XML inflates on
decode. **The retention is consistent with `archive + Σ decompressed payloads`**
rather than with the largest member or the member count: on axis 1's synthetic
incompressible media the ratio converges on 2.00×, where a stored member barely
inflates, and it rises with compressibility from there.

That formula is an *inference from the ratios*, not a direct measurement. The
probe records retained bytes, not a decompressed-payload sum, so no figure here
isolates Σ. What is measured is that A's retention is a multiple of the archive
between 2.00× and 11.19×, that the multiple tracks how compressible the package
is, and that the two structures holding bytes are the source archive and the part
blobs.

**All 19 fixtures across the three axes produced byte-identical output from the
two paths**, verified by an FNV-1a digest of the complete output stream folded in
the sink. The paths disagree only in residency.

## Which public entry points reach the expensive path

Plainly: **the ordinary, documented one, in every OOXML format.** This is read
from source, not executed.

| crate | documented open | documented save | reaches |
| --- | --- | --- | --- |
| `litchi-docx` | `Package::open` `codec.rs:273` → `OpcPackage::open_with_limits` `codec.rs:286` | `Package::save` `codec.rs:517` → `write_plain` `codec.rs:697` → `opc.to_stream` `codec.rs:1193` → `PackageWriter::write_to_stream` `package.rs:1115` | **A** |
| `litchi-xlsx` | `Workbook::open` `workbook/model.rs:432` → `OpcPackage::open_with_limits` `:439` | `Workbook::save` `:952` → `writer::save` `:959` → `PackageWriter::write` `writer.rs:34` | **A** |
| `litchi-xlsx` | `Package::open` `package.rs:44` → `:51` | `Package::save` `:506` → `writer::save` `:513` | **A** |
| `litchi-pptx` | `Package::open` `package/codec.rs:177` → `:190` | `Package::save` `:287` → `PackageWriter::write` `:296` | **A** |
| `litchi-xlsb` | `Package::open` `package/mod.rs:225` → `:260` | `Package::save` `:456` → `PackageWriter::write` `:457` | **A** |

None of these is feature-gated, opt-in, or marked legacy. The XLSX package type
**is** an eager package by definition:

```rust
pub struct Package(OpcPackage, #[cfg(feature = "encryption")] PackageEncryption);  // package.rs:35
```

The strongest evidence is not the doc comments but the project's own CRUD guide.
`docs/OFFICE_API_GUIDE.md` has eleven `.save(` call sites. Six of them open an
existing file first, and **all six are Path A**:

| guide line | opened at | shape |
| ---: | ---: | --- |
| `:146` | `:135` `Package::open("document.docx")` | update-then-save |
| `:169` | `:160` `Package::open("document.docx")` | update-then-save |
| `:195` | `:181` `Package::open("document.docx")` | update-then-save |
| `:426` | `:414` `Workbook::open("workbook.xlsx")` | update-then-save |
| `:568` | `:559` `Package::open("presentation.pptx")` | update-then-save |
| `:746` | `:670` `Workbook::open(path)` | update-then-save |

The other five — `:82`, `:320`, `:361`, `:520`, `:634` — are **create-then-save**
(`Package::new()` / `Workbook::create()`) and are deliberately excluded: a created
package has no `source_archive` and no eagerly decompressed source parts, so it
pays none of the retention measured here. By this record's own finding those
examples are cheap, and counting them would inflate the evidence. The same
correction applies to the facade's three "Quick Start … (Write)" blocks
(`crates/litchi/src/lib.rs:64`, `:108`, `:129`): all three are create-then-save,
not update-then-save. Their *reopen-and-verify* lines (`lib.rs:67`, `:111`) are
genuine eager opens and pay the open-side cost, but they are not saves.

The guide mentions source-backed packages exactly once, and only about reads
(`:718`).

This is the serious case rather than the benign one. **Every open-then-save
example the project documents is the expensive path**, and there is no documented
alternative.

Two qualifications matter for what can be done about it:

- **The `litchi` facade is read-only.** `Document::open`, `Presentation::open` and
  `Workbook::open` exist with no `save` or `to_stream` of any kind, and their read
  paths already use Path B types (`document/types.rs:37`,
  `sheet/workbook.rs:273`, `presentation/prs.rs:872`). Its only production eager opens are
  two, both in detection (`detection_smart/detected.rs:1866` and `:1986`); the
  many others in `crates/litchi/src/document/doc.rs` are inside its `mod tests`,
  which begins at `doc.rs:1838`. Writing through the facade
  means the re-exported owner crates (`lib.rs:365-378`), i.e. Path A.
- **Path B is not a drop-in replacement.** No source-backed type has a
  `save(path)` at all. Every publication entry point is one guarded editor's
  `publish_*_to_stream`: counting `pub fn` definitions outside `tests/` at
  `32d25e088`, **45 of them — `litchi-docx` 24, `litchi-xlsx` 16, `litchi-pptx`
  5** (85 call sites in total: 48, 19 and 18). None is a general save. Path A, by
  contrast, has **6** `PackageWriter::write*` call sites outside `litchi-opc` and
  outside `tests/` — pptx `codec.rs:296` and `encryption.rs:170`, xlsb
  `package/mod.rs:452` and `:457`, xlsx `writer.rs:29` and `:34` — and each sits
  directly under a public `save`/`to_stream`. Path B is an additive, bounded fast path
  layered beside Path A, which `HOTSPOTS.md`'s source-backed publication inventory
  describes with explicitly enumerated refusals and a 64-part ceiling.

Neither path carries `#[deprecated]`, and no ADR, changelog, or perf record calls
either legacy. `ADR_COMPLIANCE.md` carves `OpcPackage` atomic saves *out* of
Path B's accounting — a scoping statement that only makes sense because Path A
remains the live general-purpose route — and `HOTSPOTS.md` records Path A
being actively optimized through its `PublicationPlan` work.

### This confirms two standing GOAL.md hypotheses

`docs/GOAL.md:327-331` lists, among the unverified hypotheses about this path:

> 1. Path and generic-reader OPC input may be fully slurped into `Vec<u8>` before
>    parsing instead of retaining a positional source.
>
> 2. Ordinary `OpcPackage` opening may decompress and retain every admitted Part,
>    even for selective structural queries.

**Both are confirmed, and they compose.** Hypothesis 1 is `open_with_limits` →
`read_owned_path_with_limits` → `from_owned_bytes_with_limits`, which slurps the
file and then moves the buffer into `source_archive` rather than dropping it.
Hypothesis 2 is `load_parts_eager`. Together they account for the measured
residency — roughly *twice* the archive on media-heavy packages rather than merely
equal to it, and more where the XML inflates.

`docs/GOAL.md:367-368` attaches the standing instruction:

> These are hypotheses. Do not change an architecture solely because it appears
> suboptimal in source code. Require profiles and scenario measurements.

This record supplies the scenario measurements those hypotheses were waiting for,
and still does not change the architecture — for the contract reason below, not
for want of evidence. `HOTSPOTS.md`'s shared-OOXML-data-path inventory already stated the source-level
half ("`OpcPackage` retains every inflated Part"); what was missing was the price, the
scaling rule, and the fact that the documented door is the one that pays it.

## Why the retention exists

The two retained things have **different** contract status, and conflating them
is what makes this look simpler than it is.

### The source archive is contract-required

ADR 0005, `## 2026-08-21 amendment: OPC exact-source authorization`, lines 341-348:

> For OPC, direct byte-identical no-op publication has one authority: **the owning
> package or source-backed object must retain its exact source artifact and an
> unrevoked exact-source authorization.** Preservation provenance, ZIP indexes,
> and reconstructed graph equality are planning evidence only; none of them can
> authorize exact passthrough or a normalizing full-writer fallback. Any mutable
> OPC seam revokes that authorization. A changed owned source is publishable only
> through a proven preservation plan; if physical framing or opaque members cannot
> be preserved, publication returns a typed capability refusal before output.

The code implements exactly this, including the typed refusal:

```rust
// package.rs:1171
pub(crate) fn requires_owned_source_preservation(&self) -> bool {
    self.source_archive.is_some() && !self.exact_source_authorized
}

// pkgwriter.rs:936
if package.requires_owned_source_preservation() {
    return Err(owned_source_preservation_error());
}
```

Dropping `source_archive` would forfeit byte-identical no-op publication. **This
half is load-bearing and is not a candidate.** Note the clause's disjunction —
"the owning package **or source-backed object**" — which is why Path B satisfies
the same contract while retaining nothing.

### The decoded part blobs are not contract-required

No accepted ADR requires them. ADR 0005 `## Input and lazy state`, lines 14-17,
requires the opposite:

> Opening performs container, relationship/catalog, security, and mandatory
> structural validation. **Semantic payloads load lazily into thread-safe weighted
> caches.** Clean parsed values are evictable; active handles pin them; dirty edit
> state is never silently evicted. Cache behavior is semantically invisible.

and the repository has already adjudicated this exact question for the OLE2 side,
in `ADR_COMPLIANCE.md`'s "0574: attribution only, with no boundary touched":

> the top candidate is noted as **aligned** with ADR 0005 rather than in tension
> with it, because a shared-string table is a semantic payload and **no accepted
> ADR requires eager decoding.**

The eager OPC residency is recorded there as a **disclosed debit**, never as a
contract. `ADR_COMPLIANCE.md`'s design-gate matrix row "Exact owned-source OPC
no-op publication" carries "current eager Part memory remains" in its
I/O-and-memory column and "+22.6% profiled peak heap disclosed" in its status, and
its row "OPC source-backed reader ingress" records the equivalent reduction
already **accepted** for `SourceBackedPackage`, with no ADR exception requested:

> **Reduces compressed-plus-all-decompressed eager retention to one compressed
> buffer plus indexed metadata and deferred selected payloads**

So on ADR grounds the eager blobs are incidental. **They are still not removable**,
for a reason that is not an ADR but is an observable contract.

### …but they are not removable behind the current `Part` trait

```rust
pub trait Part: PartClone + Send + Sync {
    /// Get the binary content of this part.
    fn blob(&self) -> &[u8];                    // part.rs:52
```

`blob` is **public, infallible, and returns a borrow tied to `&self`**, and it has
**950 production call sites across seven crates** outside `litchi-opc` —
`litchi-pptx` 393, `litchi-xlsx` 265, `litchi-xlsb` 127, `litchi-docx` 106,
`litchi-ooxml-common` 47, `litchi-ppt` 7, `litchi` 5, spread over 234 files. A
lazy part behind this signature has no error
channel, so a decompression failure, an I/O error, or a limit violation on first
access could only panic or silently yield empty bytes. The facade is required to
be panic-free (`docs/adr/README.md:54-62`, decision hierarchy item 2), and silent
empty bytes would breach lossless preservation, item 1.

The decisive detail is **where the limits are charged**. Today the eager load
charges them at open, fallibly:

```rust
// pkgreader.rs:916
let blob = result?;
limits.check(ReadResource::PartBytes, blob.len() as u64, limits.max_part_bytes())?;
retained_part_bytes = checked_add(
    retained_part_bytes, blob.len() as u64,
    ReadResource::TotalPartBytes, limits.max_total_part_bytes(),
)?;
```

Deferring the decode moves `ReadResource::PartBytes` and
`ReadResource::TotalPartBytes` refusals out of `open()` — where ADR 0005:19-23
requires them to "identify the resource, observed value, limit, and object path" —
into an accessor that cannot report them. **A caller who today receives a typed
limit error from `Package::open` would instead receive a successfully opened
package.** That is an observable contract change, and it is the reason this record
does not implement.

## The decision

**Branch (c): freeze the design, implement nothing.**

The two cheaper branches were considered against the source and both fail:

- **Removing the eager retention** cannot be done without either changing
  `Part::blob`'s signature across 950 call sites in seven crates, or moving two
  typed limit refusals out of `open()` into an accessor with no error channel.
  The retention is therefore *incidental on ADR grounds but load-bearing on API
  grounds* — a distinction that matters, and that a purely ADR-based reading
  would miss.
- **Fixing the path selection** presupposes a cheap path callers could be routed
  to. There isn't one: no source-backed type exposes a general save, only
  per-editor `publish_*_to_stream` entry points with a 64-part ceiling. The cheap
  path is not merely unreachable, it is not general-purpose.

What remains is to state the price, the scaling rule, the contract that pins it,
and the gates a future change must clear. That is this record.

## Candidates, with predicted effect

Predictions below are arithmetic on the measured columns, not separate
measurements, and are labelled as such. Both tables compare **open retention
only** — `A retained (open)` against `B retained (open)` — because that is the
term a candidate would change; the publish peaks are unaffected by C2 and are
carried in axes 1 and 2 instead. The two synthetic rows therefore read
134,753,792 and 268,992,000, not the 135,268,336 and 269,506,544 that axis 1
reports for the same fixtures as `A total`.

**C1 — lazy blobs behind the existing `Part::blob(&self) -> &[u8]`.**
**Refused.** No error channel; forces panic-or-empty on first-access failure, and
silently relocates two typed limit refusals out of `open()`. Breaches ADR
0005:19-23 and decision-hierarchy items 1 and 2.

**C2 — fallible `Part::blob`, then lazy decode with the archive still owned.**
Admissible in principle; ADR 0005:14-17 actively favours it. Requires a proposed
ADR for the public trait change and migration of 950 call sites across seven
crates. Retention would fall from `archive + Σ decompressed` to `archive + index`:

| fixture | A retained (open) | **C2 predicted** | predicted factor |
| --- | ---: | ---: | ---: |
| `ArtisticEffectSample.pptx` | 2,355,882 | 1,040,267 | 2.26× |
| `saut_page.docx` | 8,352,590 | 2,992,774 | 2.79× |
| `no_drawing_patriarch.xlsx` | 7,525,438 | 683,959 | **11.00×** |
| `ConditionalFormattingSamples.xlsx` | 1,836,295 | 805,536 | 2.28× |
| `EmbeddedVideo.pptx` | 504,504 | 246,305 | 2.05× |
| `drawing.docx` | 506,648 | 120,093 | 4.22× |
| **corpus total** | **21,081,357** | **5,888,934** | **3.58×** |
| 64 MiB media | 134,753,792 | 67,254,495 | 2.00× |

C2 is bounded above by 2× on media-heavy packages, because the archive it keeps is
most of the cost there. It is worth most exactly where the archive is small and
the XML inflates.

**C3 — route the format crates' `Package::open` to `SourceBackedPackage`.** The
direction `ADR_COMPLIANCE.md`'s "OPC source-backed reader ingress" row has already
accepted, and already underway inside `litchi-xlsx` feature-by-feature. Retention becomes B's measured figure:

| fixture | A retained (open) | **C3 measured (B open)** | factor |
| --- | ---: | ---: | ---: |
| `no_drawing_patriarch.xlsx` | 7,525,438 | 11,545 | **651.8×** |
| `saut_page.docx` | 8,352,590 | 33,148 | 252.0× |
| `ArtisticEffectSample.pptx` | 2,355,882 | 67,479 | 34.9× |
| **corpus total** | **21,081,357** | **332,753** | **63.4×** |
| 128 MiB media | 268,992,000 | 25,992 | **10,349×** |

C3 is the large win and the large project. It is tempting to read this as a
*routing* problem — callers reaching the expensive door when a cheap one exists —
but that reading does not survive contact with the source. `litchi_xlsx::Package`
*is* an `OpcPackage` newtype, and **no source-backed type has a `save(path)` to
route to**: Path B's 45 publication entry points are each a
specific guarded editor's `publish_*_to_stream`, with a documented 64-part
publication ceiling and enumerated refusals recorded in `HOTSPOTS.md`. "Make the
cheap path reachable" would mean building a general source-backed save that does
not exist, then re-expressing three crates' semantic stacks on top of it. That is
a program, not a batch.

## ADR compliance

| Question | Verdict | Citation |
| --- | --- | --- |
| Must a snapshot own its bytes? | **No** — owning is explicitly opt-in | ADR 0003:8-11, "hidden shared state"; "conversion to owned storage is explicit" |
| Must the source archive be retained? | **Yes, hard** — but "package **or source-backed object**" | ADR 0005:341-348 |
| Must decoded part blobs be retained? | **Not addressed**; lazy is the stated rule | ADR 0005:14-17; `ADR_COMPLIANCE.md` "0574: attribution only" |
| Is bounded memory mandated? | **Yes** — hierarchical budget, finite profiles | ADR 0005:19-23 |
| Is a numeric peak-byte budget mandated? | **No** — peak is tracked, not bounded | ADR 0005:56-61 |
| Does preservation require a resident source? | **Not addressed**; mechanism is raw-copy | ADR 0006:8-11; ADR 0005:39-41 |
| Is `OpcPackage` specified as owning a resident buffer? | **Not addressed** — ownership is of responsibility | ADR 0011:44-45; ADR 0002:626-628 "owns **only** OPC graph selection…" |
| Do ADRs name `SourceBackedPackage` or pick a primary model? | **No** — zero occurrences in `docs/adr/` | verified by search |

Following `ADR_COMPLIANCE.md`'s house rule in its 0575 section, the absence of an
ADR basis for
eager decoding is recorded here **as an absence of documented rationale, not as
permission**. No ADR exception is requested and none is needed, because nothing is
implemented.

## Admission gates for C2 or C3

Any future implementation must show, before it is accepted:

1. A proposed ADR for the `Part::blob` signature change (C2) or for making the
   source-backed package the format crates' primary model (C3), reviewed by a
   human. Neither is authorized by an existing record.
2. That `ReadResource::PartBytes` and `ReadResource::TotalPartBytes` still refuse
   at `open()` with the same typed error, observed value, limit, and object path —
   or an explicit, reviewed decision to relocate them.
3. Byte-identical output on the full OOXML fixture corpus against the current
   eager path, by whole-stream digest, in the shape of this record's 19/19 check.
4. That exact-source no-op publication still succeeds, and that a mutated package
   that cannot be preserved still returns `owned_source_preservation_error()`.
5. Peak-retained-byte and allocation counts on the same three axes, showing the
   predicted factors above were met and that no axis regressed.
6. That `OpcPackage` remains `Clone` and `Send + Sync`, and that clones still share
   rather than duplicate the source allocation (`package.rs:1673-1679`).

## Validation preserved

Nothing was changed, so every check keeps its position and identity by
construction. Gates run against the detached worktree of `32d25e088` with an
isolated `CARGO_TARGET_DIR`:

- `cargo test -p litchi-opc` — **664 tests across 24 binaries, zero failures**,
  1 ignored. Log retained at
  [`results/change-0581/gate-cargo-test-litchi-opc.log`](results/change-0581/gate-cargo-test-litchi-opc.log).
- `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json
  --repo-root . --mode structural` — **exits 2 both before and after**, with the
  identical single message `INVALID CLAIM REGISTRY: landed claim
  'claim-0251-xlsx-xml-borrowed' requires strict evidence verification`. This gate
  is therefore already red at `32d25e088` for an unrelated pre-existing reason and
  this record does not change its state. Both runs — the pristine worktree without
  this record, and the working tree with it and both index sections — are retained
  at
  [`results/change-0581/gate-check-perf-claims.log`](results/change-0581/gate-check-perf-claims.log).
  This record adds no registry entry, because `performance_claim: none` design
  records such as 0566 and 0577 carry none.

The two gates red at `32d25e088` — three facade `document::doc` tests and
`tools/check_example_targets.py` on three duplicate iWork example targets — are
unrelated and untouched. No workspace source file was modified, so no formatting
or lint gate could change state.

## Limitations

**The Rust snippets are excerpts, not literal transcriptions.** They preserve the
code's meaning and identifiers but reflow multi-line expressions onto one line,
elide bodies with `...`, drop some doc lines, and in the `preservation_source`
snippet add two explanatory `//` comments that are not in the source. The ADR,
`GOAL.md`, `ADR_COMPLIANCE.md`, `HOTSPOTS.md` and 0578 blockquotes *are* verbatim
apart from added bold emphasis.

No timing, cold-cache, physical-device, or cross-platform result is claimed. Every
figure is callback-ordered heap accounting from a single-threaded probe on one
machine over warm fixtures; it excludes allocator-internal fragmentation, RSS, and
the page cache holding the source file. Peak is a process-global high-water mark
over a region, so it attributes concurrent allocations from any source to that
region; the probe is single-threaded, which is why that is safe here.

**The C2 column is arithmetic, not a measurement.** It is `archive_bytes +
B_open_retained`, i.e. it assumes a lazy eager-package would retain exactly the
compressed archive plus an index the size of the source-backed package's. A real
implementation would also retain per-part metadata the probe cannot see, so the
predicted factors are an upper bound on the benefit.

**The C3 column is measured, but it is measured on a different workload.** Its
figures are what `SourceBackedPackage` retains today after an open, which is a
real number for a real path — but a C3 that carried the three format crates'
semantic stacks would pin whatever payloads those stacks actually touch, and this
probe does not exercise them. The C3 factors are therefore also an upper bound,
and a looser one than C2's: the true figure depends on a working-set question this
record does not answer.

`A open retained` is the live-heap delta across the open region, not its peak; the
open's transient peak is higher (for example 269,135,062 against 268,992,000
retained at 128 MiB). Using retained is what makes the decomposition against
0578's single-region figure exact, and it is the quantity a caller holding an open
package actually pays.

Axis 3's "unedited" saves publish through `write_topology_to_stream` with an empty
`SourceTopologyPlan` for Path B and `PackageWriter::write_to_stream` for Path A;
axes 1 and 2 edit one small part. No encrypted, signed, or malformed package was
measured, and no ODF, CFB, or iWork path was touched.

**The routing claim is read from source, not executed.** The five
`Package::open` → `OpcPackage::open_with_limits` hops across four format crates
are quoted by file and line above, but no test was run to observe a
DOCX/XLSX/PPTX/XLSB facade call arriving at `load_parts_eager`; the probe drives
`OpcPackage` directly. The `.blob()` figure is a lexical match on `.blob()`
excluding `litchi-opc`, `tests/` directories, and the excluded iWork crates. It
counts **matching lines, not occurrences** — 24 lines carry two `.blob()` each, so
the occurrence count is 957 against 950 lines — and it counts neither distinct
semantic callers nor distinct call sites; a small number may be method calls on
unrelated types of the same name. The per-crate breakdown above is on the same
matching-line basis.

Finally, this record establishes that the eager path is reached by the documented
door; it does **not** establish how often real callers use that door versus the
source-backed feature modules, which would need corpus or telemetry evidence that
does not exist here.

## Disposition

Design only. Nothing is landed and `performance_claim: none`. C2 and C3 are
recorded as candidates with predicted effect and explicit admission gates; neither
is attempted, and neither is authorized by this record.
