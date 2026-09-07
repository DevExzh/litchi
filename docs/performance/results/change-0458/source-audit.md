# ODP ordinary Snapshot/transaction/commit source audit

This is a read-only source audit of the ordinary ODP append lifecycle used by the
0457 control. It records static repetition and ownership boundaries; it does not
claim that any one boundary dominates the measured latency. No production change,
build, or test was made for this audit.

The 0457 control timer in
[`odp_existing_append.rs`](../../../../tools/perf-baseline/src/odp_existing_append.rs#L665-L685)
starts at `Snapshot::from_bytes(input)` and includes `source.transaction()`,
`transaction.add()`, `transaction.commit()`, and the sink write of
`commit.snapshot().bytes()`. Setup, the input clone, title/body, and sink
construction are outside that clock. The control p50s were about
1.9 ms for tiny, 74.7--75.7 ms for medium, and 150.7--153.8 ms for large
across the two recorded runs. The allocator control p50s were 10,613,383,
110,226,105, and 211,442,207 allocated bytes, with regional peaks above entry
of 781,342, 18,027,568, and 35,958,388 bytes respectively. Those numbers are
whole-lifecycle observations, not phase attribution.

## Ordinary path

`Snapshot` retains the exact bytes, an `OwnedPackage`, and parsed slides as
`Arc`-backed state in
[`edit.rs`](../../../../crates/litchi-odp/src/authoring/edit.rs#L219-L285).
`OwnedPackage` shares its prepared ZIP index when cloned, so the ordinary
`transaction()` reopen at
[`edit.rs`](../../../../crates/litchi-odp/src/authoring/edit.rs#L374-L403)
does not rebuild the source central directory. It does re-enter
`Presentation::from_owned_package` and then constructs a detached mutable
draft. This is semantic/XML work over the same source artifact, not a second
physical ZIP-index build.

The detached draft construction at
[`mutable.rs`](../../../../crates/litchi-odp/src/authoring/mutable.rs#L191-L254)
deep-clones the validated slide projection, runs the already accepted fused
staging metadata pass, copies styles/MIME data, and independently calls
`ContentSource::parse` on `content.xml`. `ContentSource` retains a complete
owned XML string after its scanner pass at
[`content_source/mod.rs`](../../../../crates/litchi-odp/src/codec/content_source/mod.rs#L124-L183)
and [`content_source/scanner.rs`](../../../../crates/litchi-odp/src/codec/content_source/scanner.rs#L293-L335).
That source-fragment parse and XML ownership are therefore statically present
on every ordinary transaction, even when the edit only appends one slide. The
independent source-fragment validation is intentional in accepted 0442; its
presence alone does not authorize removing it.

For the append operation, `Transaction::add` checks the edit and calls
`insert_slide` at
[`edit.rs`](../../../../crates/litchi-odp/src/authoring/edit.rs#L666-L684).
The mutable insertion also updates page metadata, effective names, and
references over the current slide collection at
[`mutable.rs`](../../../../crates/litchi-odp/src/authoring/mutable.rs#L1071-L1115).
This is an O(N) staging step, but the current evidence does not establish it as
the large-input bottleneck.

The changed-slide commit path at
[`edit.rs`](../../../../crates/litchi-odp/src/authoring/edit.rs#L2164-L2237)
has the following materialization boundary:

1. `MutablePresentation::to_bytes_bounded` generates a fresh package.
2. `OwnedPackage::from_bytes` indexes those newly serialized bytes.
3. `Snapshot::from_owned_package` reparses the candidate presentation and all
   candidate slides, then compares the readback slides with the detached draft.

`MutablePresentation::generate_content_xml` at
[`mutable.rs`](../../../../crates/litchi-odp/src/authoring/mutable.rs#L1447-L1614)
builds a body buffer, copies retained source page ranges into it, generates
changed/new pages, and then creates a second source-backed output string around
that body. `to_bytes_bounded` at
[`mutable.rs`](../../../../crates/litchi-odp/src/authoring/mutable.rs#L1649-L1703)
adds the authored XML and copies untouched auxiliary members through
`get_file`. That is a fresh logical rebuild/recompression boundary required by
the ordinary artifact contract; the source-tail publication path is a
different result contract and is not a substitute for this control.

The candidate handoff accepted in 0065 avoids one later slide projection parse
for the slide-only path, but it does not remove candidate serialization,
candidate package opening, semantic readback, compact XML checks, or patch
retention. `Commit` still keeps source and after snapshots for the reversible
patch. Their `Arc` byte/index handles are cheap to clone, although the live
source and candidate ownership extends the allocation lifetime.

## Repeated XML validation and provenance work

The most plausible avoidable cost family is the changed `content.xml` final
validation/provenance probe. The writer classifies generated content as
`AuthoredOrChanged` and calls the authored XML audit through
[`writer.rs`](../../../../crates/litchi-odf-common/src/core/writer.rs#L618-L623),
[`writer.rs`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1381-L1396),
and [`writer.rs`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1498-L1514).
The audit implementation reaches `xml_minifier::audit::verify_authored` at
[`writer.rs`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1636-L1644)
while `to_bytes_bounded` is serializing the candidate.

After the candidate has been reopened, `commit` unconditionally calls
`validate_compact_xml_parts` at
[`edit.rs`](../../../../crates/litchi-odp/src/authoring/edit.rs#L2332-L2339).
For each changed XML member with a source counterpart (including ordinary
`content.xml`), the compact validator gets the candidate payload, gets the
source payload for comparison, and attempts
`xml_splice_publication` before falling back to a full compact audit at
[`edit.rs`](../../../../crates/litchi-odp/src/authoring/edit.rs#L3882-L3957).
The splice attempt then loads the same source member inside
`XmlSourcePart::load`, decompresses/copies it again, and performs a full
well-formed scan at
[`xml_splice.rs`](../../../../crates/litchi-odf-common/src/core/xml_splice.rs#L54-L89)
and [`xml_splice.rs`](../../../../crates/litchi-odf-common/src/core/xml_splice.rs#L418-L475).
It also computes the source/candidate prefix and suffix relationship and audits
the proposed fragment through
[`edit.rs`](../../../../crates/litchi-odf-common/src/package/edit.rs#L745-L828).

The static repetition is therefore:

- generated `content.xml` is audited as authored XML during package writing;
- final compact validation reads the candidate and source member;
- the splice probe can load and scan the source member again while proving
  provenance;
- if the splice classification fails, the candidate receives another full
  compact audit.

This does **not** prove that the splice probe fails for ordinary append output.
A simple tail insertion may satisfy the splice policy and avoid the fallback
audit. Conversely, a root/style namespace change or more than one differing
region can make it fall through to the full candidate audit. The ordinary
control's splice success/failure mix is not recorded in 0457. The static fact
that the source is loaded/scanned as part of the attempt is proven; the amount
of repeated candidate auditing and its latency are still measurement questions.

A future optimization hypothesis could pass explicit provenance/audit evidence
from source-aware generation to the final validator, while retaining source
identity checks and the existing splice/fallback proof when that evidence is
absent or insufficient. This is only a measurement target at this point. It
must not become a blind skip of compact validation, source lineage, malformed
XML rejection, candidate reopen/readback, media checks, or patch semantics.

## Other static cost families

The following are visible costs, but their dominance or removability is not
established:

| Source fact | Static conclusion | What remains unproven |
| --- | --- | --- |
| Draft setup clones slides and independently builds `ContentSource` | A source-fragment scan and owned XML copy occur after the 0442 staging scan | Whether this is material relative to full slide parsing, XML audit, compression, or package copying |
| `generate_content_xml` builds a body and then a source-backed output string | There are at least two large content buffers and raw retained-page copies | Peak/live allocation contribution and whether the ordinary fresh-artifact contract permits any handoff |
| `to_bytes_bounded` reads untouched auxiliary members through `get_file` | Ordinary output logically copies/recompresses source members | Per-member time versus content generation and deflate cost on the control corpus |
| Candidate creation uses `OwnedPackage::from_bytes` followed by `Snapshot::from_owned_package` | The changed output gets a new ZIP index and semantic slide readback | Whether candidate indexing or semantic readback dominates at large sizes |
| `Commit` retains source and candidate snapshots for the reversible patch | Bytes and parsed state remain live beyond publication | Whether retention, rather than transient allocation, drives the observed regional peak |

The whole-process large control profile shows symbols consistent with these
paths: `xml_minifier::audit::verify_with_policy` at 2.21% flat samples,
`xml_splice_publication` at 0.82%, `ContentSource::parse` at 0.82%, staging
parse at 0.52%, and `MutablePresentation::to_bytes_bounded` at 0.06% self
samples. It also shows substantial quick-XML namespace/attribute work, memory
movement, and deflate work. The profile is the 0457 whole-process recording
([`top-symbols.txt`](../change-0457/profiling/r2/runs/control-large/top-symbols.txt));
it includes setup, warmups, harness/reporting, and all lifecycle phases. Flat
self percentages are therefore evidence that these routines execute, not
inclusive phase shares or causal dominance. In particular, the 0.82% splice
entry does not measure all work below its callers or establish that skipping
the probe would move the p50.

## ADR and measurement boundary

This audit is consistent with the accepted snapshot/transaction and bounded
fresh-artifact contracts in
[`ADR 0003`](../../../adr/0003-snapshots-edits-and-patches.md),
[`ADR 0005`](../../../adr/0005-io-memory-and-performance.md),
[`ADR 0023`](../../../adr/0023-odf-family-crate-split.md), and
[`ADR 0024`](../../../adr/0024-current-topology.md). It also respects the
accepted [`0060`](../../changes/0060-odp-snapshot-slide-projection-reuse.md),
[`0065`](../../changes/0065-odp-final-snapshot-handoff.md),
[`0441`](../../changes/0441-odp-shared-preservation-projection.md),
[`0442`](../../changes/0442-odp-shared-staging-traversal.md), and
[`0443`](../../changes/0443-odp-compact-fragment-frames.md) changes. No new ADR
is implied by this source audit.

The next decision should come from independent phase timings and sampled stacks
inside the existing public boundaries: draft/source setup, content generation,
writer XML audit, auxiliary copying/compression, candidate package open and
semantic readback, compact/provenance validation (including splice outcome),
and patch/retention lifetime. Until that evidence exists, the repeated
changed-`content.xml` provenance work is a plausible candidate, not a proven
dominant cost.
