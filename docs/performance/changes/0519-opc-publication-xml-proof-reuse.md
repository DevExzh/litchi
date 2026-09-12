# 0519: reuse an immutable OPC publication XML proof under equal limits

`performance_claim: scoped managed DOCX publication improvement`

`claim_authorized: true`

This batch applies the small OPC proof-reuse candidate at base
`45d71cb6f0cd5d003c544b61c748552a513b0245`. `SourceXmlPart` already owns a
complete immutable XML proof. The candidate reuses that proof in
`SourceXmlPart::check_for_publication` only when the destination
`ReadLimits` is exactly equal to the proof's stored limits. The [frozen plan](../results/change-0519/plan.json),
[source review](../results/change-0519/source-review.md), and
[ADR review](../results/change-0519/adr-review.md) define the contract and
measurement boundaries. OLE2/OOXML remains active; ODF is deferred and iWork
is excluded.

## Implementation and preserved boundaries

The reviewed three-file candidate changes the OPC proof owner, updates the
public addition documentation, and adds the test-only
`publication_proof_tests` module. The production branch remains after the
existing content-type, proof-lineage, source/context, and destination
`PartBytes` checks:

* equal `ReadLimits` charges exactly one payload-length `Work` amount through
  the source execution context and skips only the duplicate XML parser pass
  and its transient parser-memory reservation;
* different limits use the existing full `validate_source_xml` path, so its
  UTF-8, DTD, namespace, depth, event, attribute, and malformed-input checks
  and typed limit errors remain; and
* the final source/context fence, source and destination freshness and
  identity checks, security policy, topology/output limits, transfer
  monitoring, and DOCX semantic reparse/readback remain in their existing
  owners and order.

Both constructor-validated original proofs and fully validated edited
`finish` proofs satisfy the same invariant. Replacements still require
destination lineage/version, equivalent Part identity, and exact original
bytes. Foreign proofs remain permitted only for the established addition
path; proof reuse does not authorize a foreign destination. A Work refusal is
still typed and occurs before output. The intentional behavior change is that
an equal-policy proof hit can succeed when the duplicate parser workspace
would previously have caused a temporary memory refusal. Retained proof,
metadata, payload, topology, and output reservations are unaffected.

The candidate adds no public API, dependency, unsafe code, mutable proof
access, global cache, executor, or ambient I/O. `SourceXmlPart` remains the
owner of proof validity and `check_for_replacement` remains the owner of
replacement destination authorization.

## Evidence summary

| Lane | Scope | Recorded result |
| --- | --- | --- |
| Native same-API | 24 cases, two fresh campaigns, 30 samples/3 warmups/two internal repeats; lifecycle phases include returned Snapshot destruction in `publish_ns` | Lifecycle p50 delta ranges `r1` −22.07%…−1.88%, `r2` −26.41%…−1.43%; publication p50 ranges `r1` −80.85%…−36.62%, `r2` −78.48%…−36.31%. Whole-child RSS ranges `r1` −6.49%…+4.03%, `r2` −3.49%…+4.18%; no RSS threshold flag. The full reports retain all 48 rows and **107 current all-phase flags** (44 in `r1`, 63 in `r2`). |
| Scoped Callgrind | Six owned-source arms, two fresh method profiles per binary; publication method only, before caller Snapshot drop | All 24 profile records pass raw/annotation scope checks. Publication inclusive Ir falls about 47.3% for p128 and 77.47–77.53% for p512. The baseline retains one XML-validator call; the candidate has no positive validator row on the proof-hit profiles. |
| Allocation probe | 24 cases, two one-row fresh child repeats per baseline/candidate binary; publication region only | All 96 captures are complete with no warnings or failed allocations. Every case is 109→96 allocation calls (−11.93%), 12→6 reallocations (−50%), and 75,001→74,218 allocated bytes (−1.04%). Incremental region peak is 72,535→72,535 bytes (0%); absolute peaks remain descriptive. |
| Guards and focused tests | Existing source/output/release/Work counters plus eight new proof tests | 3,168 paired counter rows report no guard changes, including unchanged source I/O, output identity, release state, and Work. The focused receipt records 8 passed, 0 failed, 0 ignored tests. |

The full native [comparison report](../results/change-0519/candidate-comparison.md)
and [JSON](../results/change-0519/candidate-comparison.json) contain every
case's elapsed/publication p50, p95, p99, mean, RSS, bootstrap metadata,
counter drift, output identity, and guard result. The native flags remain
part of the result: `r2` includes E2E p99 increases of **63.65%** for
`p128-k1-file-batch` and **775.44%** for `p128-k1-owned-batch`, reported as
edit outliers. The additional [tail guard](../results/change-0519/tail-guard.md)
runs both workloads through A1/B1/B2/A2: eight fresh children, each with 200
samples, 10 warmups, and one repeat. Its four matched pairs improve lifecycle
p50 by 15.21–17.67% and p99 by 14.47–18.07%, with no lifecycle, publication,
or RSS adverse flag. Three short commit-phase flags remain in the owned
A2/B2 pair: p50 310→340 ns, mean 310.75→342.70 ns, and p95 340→380 ns.
The original outliers did not recur in this follow-up; their causes remain
unproven and all original 107 flags remain retained.

The [profile report](../results/change-0519/profile-comparison.md) gives the
per-repeat inclusive Ir values, direct topology owner, snapshot-owner rows,
and validator-call checks. These instructions are method-scope diagnostics;
they do not include caller Snapshot destruction, native elapsed time, or
RSS.

The [allocation report](../results/change-0519/publication-allocation-comparison.md)
records the raw rows, build custody, and all per-case values. Its baseline and
candidate probes use the same standalone manifest, lockfile, and release
defaults, so the within-probe comparison is valid. The probe workspace uses
Cargo defaults (LTO off and panic unwind), while the normal native workspace
uses LTO and panic abort; absolute allocator counts and peaks therefore do
not represent normal production-binary values. Incremental region demand is
the meaningful comparison, and allocator elapsed time is excluded from the
native timing result.

## Work, source identity, and retention

The candidate keeps one payload-length `Work` charge on an equal-limit hit;
it does not treat proof reuse as free. The native counter comparison shows no
Work change because the 0519 branch replaces parser work with the required
payload charge. Source read calls and requested/returned bytes, exact output
bytes and hashes, live memory and object gauges, and all release/input/output
guards remain equal across the paired rows. The source proof's lineage and
version remain separate from `ReadLimits`, and the final source fence still
checks the current context before transfer.

The eight focused tests cover original and derived proofs for replacement and
addition, equal-limit reuse without parser memory, unequal-limit validator
fallback, lower XML limits, typed Work refusal, content-type/lineage order,
source revision invalidation, and replacement destination identity. Existing
tests continue to cover cancellation/version fences, signed and encrypted
policy, destination XML limits, exact output/readback, topology limits, and
reservation cleanup. The full all-features OPC, DOCX, XLSX, PPTX, and XLSB
suite passed 5,011 tests, with the independent ZIP64 preservation test passing
separately: 5,012 executed tests total. All nine quality gates passed.

The optimization is retained for its repeatable median and publication
improvements, corroborated by instruction and allocation evidence. This does
not establish uniform tail improvement: original and follow-up flags remain
explicit limitations. Exact report replay, final evidence verification, and
owned temporary-directory cleanup are recorded alongside the raw captures.

This evidence covers the managed synthetic DOCX publication route and the
OPC proof owner. It makes no native Office-producer, cold-cache,
broad-provider, hardware-attribution, or parallel-scaling claim. The full
OLE2/OOXML optimization goal remains open, and historical 0499/0500
limitations remain explicit.
