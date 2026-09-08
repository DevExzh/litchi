# Change 0478: PPTX generated-name metadata spooling

`performance_claim: none; descriptive same-source storage-policy comparison`

`claim_authorized: false`

Change 0478 adds the explicit-scratch route for the public fresh PPTX
streaming writer. It targets the retained physical-member name indexes that
remain after the central-directory records themselves can be moved to
caller-owned scratch. The route is opt-in; the existing `with_options` and
`new` constructors keep their ordinary in-memory storage policy and existing
name validation behavior.

The public PPTX surface is deliberately small:

```text
StreamingPresentationScratchLimits::new(max_bytes, buffer_bytes)
StreamingPresentationWriter::with_options_and_metadata_spool(
    writer, slide_count, options, limits, spool, scratch_limits
)
StreamingPresentationWriter::with_metadata_spool(
    writer, slide_count, spool, scratch_limits
)
```

The scratch provider is caller-owned and must implement `Read + Write + Seek +
Send + Sync + 'static`. `max_bytes` bounds the serialized central-directory
records admitted to the provider; `buffer_bytes` selects the fixed replay
window. No path is opened implicitly. The provider must preserve the admitted
bytes and exclusive, coherent access to its appended range for the lifetime of
the package, including through aliases of its backing storage. Its own storage
and retention policy remain the caller's responsibility.

Before the output sink is handed to OPC, PPTX builds the complete checked
member sequence for the declared slide count. The sequence contains 37 fixed
physical members and two members per slide:

* 15 fixed non-layout members;
* one XML and one relationship member for each built-in slide layout, giving
  22 layout members in the current resource topology; and
* `ppt/slides/slideN.xml` followed by its relationship member for each of the
  `N` declared slides.

The resulting topology is exactly `37 + 2N` members, in the same order as the
existing serializer. The generated plan is checked symbolically, so its
descriptor and cursor state does not grow with `N`. ZIP proves canonical
member spelling, exact and ASCII-folded uniqueness, numeric range safety, and
component-boundary ancestor disjointness without expanding the ranges. OPC
validates the absolute PackURI literals and indexed endpoints before passing
the checked capability down. Generated mode advances the cursor only after a
member is finalized and retains no completed name set. The ordinary writer
continues to retain its compatibility indexes and accepts arbitrary names
under its existing checks.

The new route reuses the existing PPTX serializers, slide IDs, relationship
IDs, semantic limits, output-byte budget, typed progress failures, and
deterministic member bytes. The complete plan and constructor configuration are checked before
output. Sequence violations are refused at each operation boundary. Once
output has started, scratch exhaustion and provider failures retain the
existing forward-only failure and poisoning semantics. Central-directory replay remains an
explicit storage operation, so scratch quota and provider failures retain
their typed error and accepted-output behavior.

## Measurement scope

The same-source control/spool benchmark compares fresh text-only PPTX creation
through the ordinary in-memory route and the explicit-file-spool generated
route. It uses the retained small, medium, and large slide-count corpus,
separate normal and allocator-instrumented binaries, and reversed external
repeats. Each formal report has 30 measured samples after 3 warmups, runs on
CPU 2 of the recorded AMD EPYC 9R45 environment under Rust 1.98.1, and uses
the frozen no-incremental, four-job, frame-pointer/unwind-table build policy.
Each operation covers writer construction and plan creation, slide
serialization, central-directory replay, sink hashing, and spool flush/close
according to the frozen protocol. Oracle construction, endpoint snapshots,
digest finalization, and cleanup are outside the timed interval. Mean timing
intervals use t(29) = 2.045; normal and allocator timings remain separate.

The control lane has no scratch provider. The spool lane uses the caller's
configured quota and replay window. Both lanes are required to produce byte-
identical output and to pass physical and semantic reopen checks before a
measurement is retained. The comparison is descriptive and same-source; it
does not authorize a historical latency claim, a total-process memory claim,
storage durability claim, or a claim about arbitrary PPTX editing,
repackaging, append, or Part addition.

## Final measured results

The final matrix contains 24 isolated reports and 720 measured operations.
Each normal row below is one retained report; elapsed values are milliseconds,
with nearest-rank p50/p95/p99, and RSS is the contextual process maximum from
the matching `/usr/bin/time` observation. RSS includes process setup and
teardown and is not a library heap measurement. The spool extent is the
central-directory byte count; control has no scratch extent.

| Repeat | Slides | Policy | Mean ms | p50 ms | p95 ms | p99 ms | RSS KiB | Scratch bytes |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 8 | control | 0.905319000 | 0.904344000 | 0.915683000 | 0.919723000 | 4,044 | — |
| 1 | 8 | spool | 0.933504600 | 0.933944000 | 0.943084000 | 0.947064000 | 4,044 | 4,074 |
| 2 | 8 | control | 0.901271433 | 0.900583000 | 0.915664000 | 0.917114000 | 4,064 | — |
| 2 | 8 | spool | 0.931108567 | 0.929294000 | 0.945394000 | 0.961084000 | 4,052 | 4,074 |
| 1 | 256 | control | 6.344201900 | 6.339666000 | 6.455467000 | 6.455507000 | 6,808 | — |
| 1 | 256 | spool | 6.320574300 | 6.321027000 | 6.341257000 | 6.345406000 | 6,832 | 40,842 |
| 2 | 256 | control | 6.471067100 | 6.480767000 | 6.492937000 | 6.523597000 | 6,844 | — |
| 2 | 256 | spool | 6.295404733 | 6.293517000 | 6.316657000 | 6.323947000 | 6,596 | 40,842 |
| 1 | 8,192 | control | 185.536250633 | 184.195604000 | 188.825695000 | 189.065036000 | 102,860 | — |
| 1 | 8,192 | spool | 182.077139700 | 182.301259000 | 182.950661000 | 183.892064000 | 99,536 | 1,237,692 |
| 2 | 8,192 | control | 185.602817200 | 185.164615000 | 189.965465000 | 190.741678000 | 101,472 | — |
| 2 | 8,192 | spool | 181.198961100 | 182.400040000 | 182.932063000 | 183.125533000 | 99,416 | 1,237,692 |

The small eight-slide lane pays the measured spool setup cost: its spool mean
is 3.113% above control in repeat 1 and 3.311% above control in repeat 2.
The analyzer reports no positive policy flag above 5% across all 12 policy
pairs (six normal and six allocator), and no repeat-drift flag above 5%. Those
results
describe this warm, same-source operation scope; they do not authorize a
latency improvement or regression claim for the general PPTX writer.

Allocator values are identical in both external repeats; each arrow is
control → spool and each RSS pair is repeat 1 / repeat 2:

| Slides | Allocation calls | Requested bytes | Incremental peak live bytes | RSS KiB | Scratch bytes |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 8 | 850 → 617 | 499,462 → 475,312 | 435,701 → 432,436 | 4,028 / 4,060 → 4,036 / 4,028 | 4,074 |
| 256 | 9,049 → 5,825 | 1,415,242 → 861,904 | 681,819 → 432,436 | 6,812 / 6,800 → 6,844 / 6,600 | 40,842 |
| 8,192 | 278,153 → 179,677 | 31,428,173 → 13,379,518 | 8,875,252 → 432,436 | 99,524 / 101,316 → 102,880 / 99,412 | 1,237,692 |

The scoped memory gate covers 180 spool allocator samples (three sizes,
two external repeats, 30 samples each): every incremental peak is exactly
432,436 bytes, with zero failed allocations and zero live-byte exit delta.
The gate's global range is 0% against its 1% threshold. This is an
operation-scoped allocator observation for the generated spool route; it does
not bound total process RSS, provider storage, native allocations outside the
instrumented allocator, or the ordinary control route.

The explicit quota is 67,108,864 bytes and the replay window is 16,384 bytes.
Observed scratch is 4,074 / 40,842 / 1,237,692 bytes for 8 / 256 / 8,192
slides, respectively, matching the `Σ(46 + UTF-8 filename length)` identity
over the 37 fixed names and two names per slide. The provider is on tmpfs in
this environment, so scratch remains caller-selected system memory and does
not represent durable-storage latency.

## Final custody and validation

The final production source is commit `69765160b`; the final diagnostic
harness is `b16fb3962`. The copied normal and allocator binaries bind the
unchanged 7,046-file source manifest whose final digest is recorded by the
build receipts. The data-only verifier passes all 24 reports, 720 samples, output
identities, scratch arithmetic, physical-member checks, semantic PPTX checks,
and summary invariants.

The final evidence set records the full validation ledger in
`rust-validation.json`, the live source/template/binary custody in
`live-state.json`, authenticated runtime removal and cache retention in
`cleanup.json`, and portable fresh-copy/negative-probe replay in
`portable.json` with `SHA256SUMS`. The ledger retains development failures
alongside the required final gates. The final gate totals are 1,887 library
tests with 5 ignored over 103 targets, 387 harness tests with 1 ignored, 180
shared streaming tests, DOCX 4, ODF 14, ODT 14, ODS 16, ODP 12, and 21
Python evidence tests. The coordinator's final live, cleanup, and portable
receipts are the authoritative completion checks for those stages.

The final rerun includes the OPC percent-encoded-prefix guard. The two earlier
formal matrices, `preliminary/formal-before-opc-guard` and its nested
`preliminary/formal-before-cleanup-fix`, were captured before that guard and
before the cleanup fix; they remain archived for custody and are excluded from
the 24-report result above. A
separate allocator binary copy initially failed with errno 122 after a
319,131,648-byte partial destination was written; the authenticated partial
copy was removed, and `build-copy-recovery.json` records the successful
414,504,696-byte copy matching its origin at SHA-256
`6279bccb61337e572d215b7f788a670e95ea839a6d6dfe9a83e5f61fae52b632` with
the source unchanged.

The final evidence supports only the scoped operation-heap observation and
the exact correctness/storage-policy comparison above. It makes no latency,
general-memory, durability, arbitrary-name, append, Part addition,
repackaging, native/source-variant, or full non-iWork goal claim.

The remaining full-goal work includes the broader Office creation and CRUD
coverage, source/native variant breadth, scaling and worker policies, and the
non-iWork requirements tracked by the coordinator. This change covers the
fresh public PPTX text-only route and its physical metadata ownership.

The final source also rejects an OPC indexed prefix ending `%` before any
endpoint allocation: valid endpoint URIs can otherwise hide an invalid
percent-encoded name inside the range. The regression checks that refusal is
atomic and that fixed percent escapes remain accepted. Both earlier matrices
are retained separately; the final rerun binds this guard as well as the
benchmark cleanup fixes.
