# 0729 result review: public DOC attribution and the next bounded candidate

Status: result interpretation, 2026-09-22. The frozen packet contains 72
processes and 3,600 measured lifecycles across two DOC cases and four routes.
The analyzer accepted every process, the audit matched the analysis, and all
15 negative controls were rejected as intended. The source and semantic oracle
contract is inherited from 0728. This review changes no frozen artifact,
native source, Cargo file, or probe implementation.

The detailed evidence is in
[`analysis.md`](analysis.md), [`analysis.json`](analysis.json),
[`report-summary.json`](report-summary.json), and [`audit.json`](audit.json).
The result is descriptive attribution evidence; it is not a production
speedup or allocation claim.

## What the measurements establish

The two cases behave differently under the control comparisons:

| Case | Profiled-clock whole p50 range | Profiled-clock whole mean range | `commit.Finish` p50 share | Control flags |
| --- | ---: | ---: | ---: | ---: |
| `docfloat` | 1,021.1–1,094.6 µs | 1,034.4–1,085.3 µs | 7.68–8.18% | 17 of 54 |
| `docnohf` | 107.3–109.5 µs | 108.8–111.5 µs | 13.00–13.30% | 0 of 54 |

The `docnohf` p50 and mean controls stay within the five percent review
threshold for all 54 matched comparisons. Its profiled Finish phase is a
stable, material part of the complete public lifecycle. The `docfloat` case
has 17 flags: five mean and four p50 flags for ordinary opaque → ordinary
split, two mean and four p50 flags for ordinary split → profiled empty, and no
mean flags plus two p50 flags for profiled empty → profiled clock. The broad
per-process ranges show that a small candidate effect on this larger fixture
cannot be inferred from the phase share alone.

The named semantic phases also show why the candidate seam is narrow. On
`docfloat`, final profiled public-reader validation occupies about 17.2–18.1%
of whole time and Finish about 7.7–8.2%; on `docnohf`, public-reader
validation occupies about 11.4–11.6% and Finish about 13.0–13.3%. The outer
replacement window is larger than Finish on both cases, about 20.0–22.5% for
`docfloat` and 32.1–32.8% for `docnohf`. A rendered handoff can target the
later duplicate owner render, but it does not erase the semantic replacement
staging work or the final independent validation.

The profiled/ordinary comparison includes the known strict-editor lifetime
difference: ordinary open keeps the strict editor through public-reader
validation and source retention, while profiled open drops it at the strict
owner closure. The 0729 binary enables the diagnostic feature for every
route, so this packet does not compare default-build code generation against
feature-enabled code. It also contains no new allocation, peak-memory, RSS,
or producer-corpus evidence.

`observer_clock_control_ns` is reported outside the lifecycle and is never
subtracted. The audit and negative controls protect event order, spans,
residuals, source identity, semantic witnesses, output hashes, and control
presence. These gates support the attribution labels; they do not convert the
profiled route into an ordinary allocation control.

## Recommendation

The results justify a narrowly scoped validated-render candidate as the next
engineering experiment. They do not justify landing it as a production
optimization yet. More inner phase attribution is unlikely to answer the
remaining question: the existing data already identifies the candidate seam
and bounds the possible final-render portion. The next evidence should compare
the candidate against the ordinary public route directly, with memory and
ownership gates included.

The candidate should be a DOC-owner-specific batched handoff at the existing
common-editor boundary used by `RevisionEditor::commit`. The current
`put_stream_shared_with_rendered` method handles one stream; this DOC edit can
publish WordDocument, the selected Table stream, and possibly Data together.
A future private batched equivalent can return the one rendered `Vec<u8>` that
already passed common candidate check, CFB reopen, recapture, allocation
reuse, and target discovery, while atomically installing the corresponding
candidate package state. The DOC owner can then consume those bytes for the
next publication boundary instead of asking the common editor to render the
same final candidate again.

This candidate must retain the current default `SectorLayoutPolicy::Reuse`
behavior and its fallback decisions. Switching to `Rewrite` would change the
question and could change preservation, directory layout, output size, and
the measured cost. The candidate is a handoff of a validated result, not a
sector-policy change.

## Required proof before production adoption

The handoff should remain a one-shot, bounded transfer. It must not become a
generic rendered-output cache on `ObjectEditor` or a cache shared across DOC
and PPT owners.

1. **Exact identity and freshness.** Bind the returned bytes to the exact
   source identity, current DOC revision/package generation, complete batched
   replacement paths and data, target catalog, current layout policy, and all
   relevant limits. A stream path/value key is insufficient. Any later public
   edit, stream or metadata mutation, topology change, policy/limit change,
   or source-freshness mismatch must invalidate the handoff.
2. **Atomic state and token publication.** Keep the prior `RevisionEditor` and
   common package unchanged until candidate rendering, CFB reopen, codec and
   package checks, recapture, allocation reconciliation, and discovery all
   succeed. Publish the candidate state and its one-shot rendered token
   together. A failure must leave no token describing an unpublished state.
3. **One-shot ownership.** Consume the rendered bytes immediately at the DOC
   owner boundary, or carry one generation-tagged token that can be consumed
   once. A later edit must drop it and use the current rebuild. Do not retain
   arbitrary prior renders through successive semantic operations.
4. **Explicit retained-output budget.** Account for one extra output-sized
   rendered CFB together with source bytes, captured candidate state, stream
   replacement allocations, final validation state, and transient reopen
   buffers. Charge the retained artifact against an explicit bounded
   output/candidate budget. If the budget does not fit, fall back to the
   existing validated render/finish route; a handoff budget refusal must not
   make an otherwise supported edit fail.
5. **Independent final validation.** The handoff may remove a duplicate render,
   but it must still run the final strict DOC-owner validation, independent
   public-reader validation, semantic paragraph readback, source-freshness
   checks, and reversible patch construction. Common candidate validation is
   not a substitute for the public DOC reader.
6. **No-op and preservation behavior.** Equal replacements must remain exact
   source no-ops with unchanged state and patch behavior. Changed candidates
   must preserve untouched streams, directory metadata and CLSIDs, semantic
   projections, deterministic output policy, and current Reuse/fallback
   behavior. Failed later validation must not publish a partially optimized
   result.

The likely ownership shape is a private return path from the DOC revision
commit into the outer `Edit::commit`, rather than a long-lived field copied
through every mutator. If the public edit can perform another operation after
replacement staging, the generation token must be invalidated there. The
single replacement used by this packet is evidence for the seam, not proof
that a cached result survives arbitrary multi-edit transactions.

## Evidence required for the candidate experiment

The next run should add an ordinary public control and a candidate public
route to the same fresh-process schedule, with identical source, replacement,
limits, and owner scopes. It should retain the current oracle and compare:

* exact output bytes and complete stream/path inventory;
* changed paragraph text, UTF-16 length, all survivor projections, untouched
  streams, directory metadata, CLSIDs, and source-checked forward/inverse
  patches;
* exact no-op behavior and at least one error/refusal path;
* successive-edit invalidation, candidate failure atomicity, policy fallback,
  and source-freshness behavior;
* allocation counts, peak live bytes, and retained output ownership; and
* whole lifecycle time, with the existing per-process p50/mean/p95/p99 and
  paired control-difference reporting.

The current `docnohf` controls are stable enough to use as an initial candidate
fixture. The 17 `docfloat` control flags require conservative interpretation:
repeat the paired candidate/control schedule or enlarge the process matrix
before accepting a small timing delta there. A candidate should be accepted
only when its output and ownership proofs pass and its measured change exceeds
the observed paired noise without relying on phase subtraction.

The current two fixtures do not establish producer breadth. Before a broad
DOC claim, repeat the candidate proof on representative Word and LibreOffice
artifacts, especially packages with floating objects and larger Data streams.
The 0729 phase result identifies where to test; it does not establish that a
retained rendered artifact is beneficial across those producers.

## Disposition

Proceed to a private, batched, generation-bound handoff prototype with an
explicit retained-output budget and fallback. Keep ordinary save semantics,
Reuse policy, strict/public validation, no-op behavior, and patch freshness
unchanged. Do not introduce an unbounded cache, switch layout policy, or claim
a production speedup from the 0729 phase share. The next acceptance decision
belongs to the direct candidate A/B plus allocation and ownership evidence,
not to another round of subtracting independently measured phases.
