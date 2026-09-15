# 0613: publication's audit of the *original* bytes is 27.6% of a source-backed XLSX save, and the memo that would reuse it can never be read, because every publication consumes its package

Status: rejected, not implemented. The implementation, its five internal tests
and its three public-contract tests are retained as a patch in the evidence
packet. `performance_claim: none` — the callgrind isolation pairs, call counts
and paired wall-clock medians below are reported as evidence, not registered as
claims. **No file under `crates/` is modified by this change.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This record answers the memo half of item **XLSX-3 / XML-5** of
[0587](0587-remaining-opportunity-survey.md) (rank 26). Base
`2d6fbeaed2083de104bbb0b52b2990ce69ac7274`; branch
`perf/0613-opc-original-audit-memo`.

## What was changed

Nothing in production. This record carries three things:

1. the **measured price of the original-bytes audit**, split out from the
   combined figure change [0528](changes/0528-xlsx-publication-attribution.md)
   reported: the audit of the *original* is **27.59%** of source-backed XLSX
   publication instructions, the audit of the *replacement* another **27.60%**,
   and the pair **55.19%** — which independently reproduces 0528's 56.2257%
   on a different corpus, at a different base, with change 0593's
   pristine-member skip already landed;
2. the **implemented memo**, built exactly as the item asks, gated, and
   measured to fire **zero** times on every reachable scenario, with the
   structural reason it can never fire;
3. the **recommendation not to implement it**, and what would have to change
   first — both of which are contract questions for human review, and both of
   which sit behind [0602](0602-xlsx-real-producer-admission-design.md)'s D0.

## The item, and what it asked

0587 ranked XLSX-3 at 26 and described two doors:

> Two new mechanisms: memoize the original-audit verdict per (source lineage,
> part name) inside `SourceBackedPackage`, or a proof-carrying replacement
> door, which needs a proposed ADR. 0528 requires both audits to stay; this
> reuses, it does not drop.

The survey's own working note (`results/change-0587/survey/xlsx.md`, R3) states
the premise the memo rests on:

> for repeat publications from the same source it is audited again each time

That premise is the subject of this record. It is false — not because the
audit is cheap, but because *there is no second publication to serve*.

## The mechanism, re-read on this base

For every replaced XML Part, source-backed publication audits the original
bytes and then the replacement:

```rust
} else if xml_minifier::audit::package::is_xml_part(
    part.partname.as_str(),
    &part.content_type,
) {
    validate_overlay_xml(part.partname.as_str(), original.as_bytes())?;
    validate_overlay_xml(part.partname.as_str(), &replacement.replacement)?;
}
```

`crates/litchi-opc/src/source_backed.rs` has five such sites after change
[0593](0593-opc-publication-pristine-members.md) moved the surrounding code:
the topology path (the one above, reached by every XLSX cell-value save), the
single-Part overlay door, the single-Part-with-relationship-removals door, and
the two bounded multi-Part doors. `litchi-xlsx` never supplies a
`SourceXmlPart` proof, so change [0519](changes/0519-opc-publication-xml-proof-reuse.md)'s
skip does not fire for XLSX and both audits always run.

The soundness argument for a memo is real. The original bytes of an admitted
Part cannot change while one source lineage lives: `SourceSnapshot::ensure_current`
refuses `OpcError::SourceChanged` before any audit runs, and the catalog is
immutable. `verify_authored` is a pure function of those bytes and the
package's own `ReadLimits`. A verdict computed once is the verdict every later
audit of the same Part in the same lineage computes.

What the argument does not supply is a *second reader*.

## Why the memo can never be read

Two independent facts, both checked against the source at this base and then
confirmed by measurement.

**1. Every publication entry point consumes its package.** `SourceBackedPackage`
has ten public publication doors — `write_topology_to_stream`,
`write_part_overlay_to_stream` and its shared/accounting variants,
`write_part_overlay_with_external_relationship_removals_to_stream`,
`write_part_overlays_to_stream` and its shared, deletion-bearing and
relationship-removal variants — and every one of them takes `self` **by
value**, as do the private `write_single_part_overlay_to_stream`,
`write_part_overlays_impl`, `write_exact_source` and the
`write_changed_overlays*` family. `litchi-xlsx`'s editors are the same shape:
`SourceBackedEditor::publish_commit_to_stream`,
`publish_multi_commit_to_stream`, `publish_patch_to_stream` and
`write_snapshot_overlay_to_stream` all consume the editor, which owns the
package. One package therefore publishes at most once. A second publication
requires a second open, and every open mints a fresh `SourceLineage(Arc::new(()))`
— deliberately, because "two source adapters may report the same caller-chosen
version token while still being different package instances; patches must
never cross that boundary."

A memo held *inside* `SourceBackedPackage`, keyed on that lineage, is therefore
destroyed at the end of the only publication that could have populated it.

**2. Within one publication every audited original is a distinct Part.**
Duplicate targets are refused before any audit: `SourceTopologyPlan::replace_part`
returns `OpcError::DuplicatePartName` when a part name already appears among
its replacements, additions or removals; `write_part_overlays_impl` performs
the same equivalence check over its sorted replacement and deletion lists; and
the single-Part doors take exactly one Part. So the memo cannot be hit twice
inside one publication either.

The two facts together mean `recall` returns `None` on every call, in every
program reachable through the public API. The measurement below confirms it.

## Measured

Host: AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`, `rustc 1.98.1`. Both legs
built `--release --locked` from the same workspace, the before leg from the
shared read-only checkout at `2d6fbeaed`. Every measured process is pinned to
CPU 9 with `taskset`. Deterministic counts first; timing last.

Selector: `xlsx_source_backed_cell_values_one_edit_save` (and its managed
sibling is unaffected for the same structural reason). One harness iteration
publishes **two** corpus shapes — `medium` and `dense-sparse` — so one
iteration is two source-backed publications. Note the harness corpora are
litchi-written and therefore compact; that is why they publish at all, and it
is the whole of this measurement's reach (see Limitations).

### Audit call counts, callgrind isolation pair

Callgrind records one `calls=` line per call site, so summing them per callee
gives the exact number of times a function ran, attributed to its caller.
Profiling `--samples 1` and `--samples 3` and differencing isolates one
iteration.

| `verify_authored` calls, per iteration | before | after |
| --- | ---: | ---: |
| from `source_backed::validate_overlay_xml` | 8 | 4 |
| from `SourceBackedPackage::validate_original_part_xml` | — | 4 |
| **total, source-backed publication** | **8** | **8** |
| from `pkgwriter::PackageWriter::validate_authored_xml` (the eager control, outside the loop) | 62 (constant) | 62 (constant) |

`is_xml_part` calls from `write_topology_to_stream` are 4 per iteration in both
legs: **two replaced XML Parts per publication, each audited twice** — once as
the original and once as the replacement.

**Memo hits: 0.** The total is unchanged because every `recall` missed. This is
the decisive count: the memo is not slow, it is never read.

### Instructions per iteration

| scope | before Ir | after Ir | delta |
| --- | ---: | ---: | ---: |
| whole case iteration | 3,159,369,076.5 | 3,159,733,758.0 | +364,681.5 (+0.0115%) |
| publication (`write_topology_to_stream`) | 188,320,847.5 | 187,896,532.0 | −424,315.5 (−0.2253%) |
| both audits | 103,695,957.0 | 103,696,710.5 | **+753.5 (+0.0007%)** |
| — original audit only | — | 51,844,180.0 | |
| — replacement audit only | — | 51,852,530.5 | |

The audit-scope delta, **+753.5 Ir per iteration**, is the only figure
causally attributable to the memo: about 188 Ir per audited original for one
uncontended lock, one miss scan, one `try_reserve` and one push. The
publication-scope and whole-program deltas have opposite signs and are two to
three orders of magnitude larger; they are code-layout noise, not effects —
the memo can only add work, never remove any, since it never answers.

The shares are the useful result:

| | share of publication Ir |
| --- | ---: |
| audit of the **original** bytes | **27.59%** |
| audit of the replacement bytes | 27.60% |
| both audits, after leg | 55.19% |
| both audits, before leg | 55.06% |

0528 measured both audits at 56.2257% of publication on its own medium and
dense-sparse profiles. 55.06% here, on a different base with 0593's pristine-member
skip already landed, reproduces it. 0587 modelled the original half at "at most
about 28% of publication"; the measured value is 27.59%.

### Paired wall clock, both directions, with the floor

Order A1 B1 B2 A2 A3 A4 (before, after, after, before, then the A/A pair), 30
samples per leg after 3 warm-up iterations, all on CPU 9 in one window. The
before medians pool A1 and A2, the after medians pool B1 and B2, and A3 against
A4 is the floor.

| shape | before p50 (ns) | after p50 (ns) | after − before | before − after | A/A floor |
| --- | ---: | ---: | ---: | ---: | ---: |
| `medium` | 5,421,559 | 5,312,024 | −2.02% | +2.06% | 2.18% |
| `dense-sparse` | 35,389,801 | 35,177,666 | −0.60% | +0.60% | 3.03% |

Both differences are smaller than the floor measured in the same window, on
both shapes and in both directions. Per-leg p50, mean, p95 and p99 for all six
legs are in `results/change-0613/timing/timing-one-edit.json`. Nothing here is
a speedup; a memo that never answers cannot produce one, and the counts above
say why. The floor is quoted as measured, not as the host's nominal p50 4%.

**All six legs on both shapes produced one output SHA-256 each** —
`9b7b66a0…` for `medium` and `6bcafbdb…` for `dense-sparse` — so the candidate
publishes byte-identical packages.

## Correctness evidence

The patch was gated as if it were landing, so that the retained implementation
is adoptable verbatim if a reachable door is ever built.

* **Five internal tests** (`crates/litchi-opc/src/original_audit_memo_tests.rs`):
  a repeated audit of one Part reproduces the acceptance; a repeated audit
  reproduces the refusal with the same `xml_minifier::audit::Error` value, the
  same `OpcError::XmlPublication` part name, the same `Display` and the same
  `Debug`; a planted verdict is the one publication reports, which proves the
  memo — not a re-run of the auditor — answered; a payload length the memo did
  not record takes a fresh audit; and each Part keeps its own verdict.
* **Three public-contract tests**
  (`crates/litchi-opc/tests/source_backed_original_audit.rs`): two publications
  from one source produce identical bytes; two publications of an original the
  audit refuses report the identical refusal; and a source that changed under
  the package is refused with `OpcError::SourceChanged`, before any recorded
  verdict is consulted.
* `cargo fmt --all --check`, `cargo clippy -p litchi-opc -p xml-minifier
  --all-targets` (workspace lints deny), `cargo test -p litchi-opc -p
  xml-minifier` — **737 passed, 0 failed, 2 ignored** — and `cargo doc -p
  litchi-opc -p xml-minifier --no-deps` all pass on the candidate.
* **Differential, downstream:** `cargo test -p litchi-xlsx -p litchi-docx -p
  litchi-pptx` — **3,629 passed, 0 failed, 33 ignored** — covering the three
  format owners that drive the five audit sites.
* **Byte identity:** identical output SHA-256 on both corpus shapes across all
  six timing legs, plus the harness's own per-iteration output, semantic and
  untouched-member digest oracles, which pass in every run above.

The patch carries one change outside `litchi-opc`: `#[derive(Clone)]` on
`xml_minifier::audit::Error`. The enum is `#[non_exhaustive]`, so a memo cannot
rebuild a refusal variant by variant from another crate; reproducing the exact
refusal requires the derive. It is additive and no behaviour depends on it.

## Validation preserved

Nothing about which bytes are published or which inputs are refused changes,
on any input — most of all because nothing is landed. Within the retained
patch: both audits stay, as 0528 requires; the replacement is audited in full
every time; the original audit keeps its auditor, its `Limits::default()`, its
call site and its `OpcError::XmlPublication { part, source }` construction, so
a refusal surfaces at the same point with the same identity and message; the
memo never observes a replacement payload; freshness is unchanged, because
`ensure_current` refuses a changed source before the memo is reached; and the
retained set is bounded by one entry per admitted Part, with an allocation
failure degrading to a fresh audit rather than to a publication failure. **What
change 0602 calls D0 — that the audit of *original* bytes refuses 94 of 95 real
producer packages — is untouched. This record does not change what the original
audit refuses.**

## What would have to change first, and why it is not worth it yet

Two doors could give the memo a reader. Both are contract questions for human
review, and both are behind D0.

**(a) A non-consuming publication door.** If a package could publish more than
once — `fn publish(&self, …)` instead of `fn publish(self, …)` — a memo keyed
on the lineage would serve the second and later saves. That moves the
publication contract itself: today, consuming `self` is what makes "the source
this package was opened on has not been published over" structural rather than
checked, and it is what lets publication drop read-ahead, retire reservations
and hand the archive to the writer. Re-entrant publication needs its own
record.

**(b) A caller-owned cross-lineage memo.** The realistic repeated-save shape —
open, edit, save, open, edit, save — crosses lineages, so a memo that survives
must be handed in by the caller: a new opaque `SourceAuditMemo` passed at open,
keyed on `(source version, part name, length, digest)` rather than on the
catalog index. That is new public API with a new key that includes a digest
the audit does not compute today, and a process-local store is exactly the
ambient state the ADRs refuse. It also trades a full XML parse for a SHA-256 of
the same bytes, which callgrind cannot price (it runs SHA-256 in software) and
which would need `perf stat` cycles.

**Both are bounded by 0602's D0 before either is worth building.** On real
producer packages the *first* original audit refuses — 94 of 95 fixtures — and
publication aborts. There is no second publication to serve, so a memo's reach
on real packages is nil until the original-bytes contract changes. On
litchi-written packages, where the audit accepts, the reach is one publication
per open, which is zero. The memo is downstream of D0, not a substitute for it,
and the 27.59% measured above is the size of D0's prize, not of the memo's.

The proof-carrying replacement door (0587's second mechanism) is untouched
here, as the brief directed: it remains a proposed-ADR item, and 0602's D1–D4
are its prerequisites.

## Limitations

* **What is not claimed.** No speedup, no regression, no allocation result, no
  claim about any scenario other than `xlsx_source_backed_cell_values_one_edit_save`
  on this host, this base, this build and these two corpus shapes.
* The 27.59% share is **measured** on the harness's litchi-written corpora,
  where the original audit *accepts*. It is **unknown** for a package whose
  original audit refuses, because publication stops at the first refusal and
  never reaches the second Part.
* Instruction counts rank work, not latency (change 0579: 1.24% of instructions
  against 6.19% of cycles). The share above is an instruction share of
  publication, not a time share of a save, and callgrind's software SHA-256 and
  per-byte `rep movsb` inflate the hashing and copying around it.
* The timing legs are **measured** and sit inside a measured floor; they
  establish "no detectable difference", not "no difference".
* "Zero memo hits" is measured on this selector and argued structurally for the
  rest of the API. The argument rests on two source facts — every publication
  door takes `self` by value, and duplicate part names are refused before any
  audit — both of which a future change could invalidate.
* The A/A floor in this window was 2.18% (`medium`) and 3.03% (`dense-sparse`)
  at p50, below the host's stated p50 4%; no p99 floor is reported.

## Retained evidence

[`results/change-0613/README.md`](results/change-0613/README.md) — the four
callgrind profiles, the extracted call counts and inclusive costs, the derived
per-iteration arithmetic, the six-leg timing report, the two extraction
scripts, the replay script, the gate tails, and the complete implementation as
`patch/0613-original-audit-memo.patch`.
