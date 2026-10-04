# 0897 — allocation attribution for the restored full-cell fallback

This batch reuses the retained 0895 baseline profiles to locate allocation
work in the current producer XLSX full-cell route. It follows 0896's rejected
stream QName candidate: its direct-stream allocation savings did not produce
a demonstrated improvement in the current fallback workflow. The current
source and harness still match the profiled baseline, so these profiles can
answer the next allocation question without another workload capture.

This is an offline analysis of two existing 20-scan processes, not a new
benchmark population or native latency result. No production change is made.
The broader performance program remains active; iWork is excluded.

## Evidence and attribution scope

The fresh offline replay of the original 0895 verifier passes its complete
four-child capture, export, semantic-guard, and reduction checks. This analysis
selects only baseline children 00 and 03. It does not reuse rejected-candidate
counts to characterize current production. Source identity checks compare all
selected build paths and ADRs from 506aa2 to c6d7ba22f5 and bind the current MCE,
XLSX x14ac, test, and producer-runner source hashes to the retained manifest.

Heaptrack counts whole-process allocation events. Stack-presence regions
separate the scan closure, forced worksheet guard, remaining runner, and
remaining process. They are not temporal instrumentation of the benchmark
clock. Inlined stack order and symbol availability limit call-site precision;
all unmatched and unresolved costs remain in the denominator.

The reducer and source audit are in [the retained packet](results/change-0897/).
The original raw profiles, stacks, commands, reports, compiler identity, and
binary hashes remain in the linked 0895/0894 packets. No native speedup,
allocated-byte reduction, peak-memory improvement, or adoption is inferred
from allocation call counts.

## Restored scan-closure allocation costs

Both baseline processes have 673,663 whole-process allocation calls. The
scan-closure partition contains 518,660 calls over 20 scans, matching the
separate baseline counter of 25,933 calls per scan. Setup and other process
work contributes 154,156, remaining runner work 847, and no allocation stack
matches the forced worksheet guard. Zero guard allocations does not mean the
guard was skipped; its semantic receipts pass.

| Disjoint scan-closure group | Calls over 20 scans | Share |
| --- | ---: | ---: |
| Nearest legacy MCE owner is `start` | 343,320 | 66.19% |
| Other legacy MCE owners | 2,060 | 0.40% |
| No legacy MCE owner frame | 173,280 | 33.41% |
| **Total** | **518,660** | **100%** |

Within `start`, two export rows with identical stack text carry 166,800 and
420 calls. Their combined string-copy signature is **167,220 calls** (32.24%
of all scan-closure calls). The complete duplicate-check subgroup accounts for
**85,560 calls** (16.50%). These are stack signatures, not source-line
allocation counts. In particular, the string-copy signature cannot prove
which of the three `Mode::Emit(q.to_owned())` branches or other owned
expressions in `start` allocated, or how much a name-span change would remove.

The full disjoint `start` partition is 168,700 ToOwned-signature calls, 85,560
duplicate-check calls, 86,060 generic growth/reservation-signature calls, and
3,000 other calls. The dominant 167,220-call exact signature is contained in
the broader ToOwned subgroup; these figures must not be added. Generic growth
frames do not identify which vector grew.

A separate non-MCE stack group through `Number::new` accounts for 73,720
calls (14.21%). Its owned lexical number text serves the current lossless cell
representation; it is not classified as an avoidable temporary. The source
audit also distinguishes whole-process x14ac inclusive presence (24,615 calls)
from the earlier exclusive nearest-owner category (9,792). Those figures have
different scopes and cannot be added to the scan-closure table.

## Next bounded mechanism

The legacy MCE codec retains an owned qualified name in `Mode::Emit` so it can
write the closing tag. Both start-tag writer paths already append `<` followed
by that same name to the append-only bounded output. A candidate can retain
an output byte range instead of another name allocation, then copy that range
through the existing fallible output reservation when closing the element.

This is a source-supported hypothesis, not a measured saving. It requires no
new reader implementation, lifetime escape, cache, or validation bypass. The
implementation must preserve inactive and opaque branches, skipped and
unwrapped elements, empty-element expansion, namespace hoisting, alternate
content, report counters, and the exact `</`, name, `>` output-limit checks.
Range arithmetic must be checked; reallocation must not invalidate a retained
range or introduce unchecked slices or allocation. A range is useful only for
an emitted active element; inactive frames must retain their control behavior.

The next candidate should target this mechanism alone. Existing inline
attribute storage (0887), validated-name reuse (0890), selected metadata
admission (0894), and stream QName validation (0896) were rejected under their
own gates and are not bundled into it. It must pass MCE differential and
adversarial correctness checks, shared-format controls, and a fresh matched
native trial before adoption. Allocation counts cannot override that trial.

## Review and cleanup

The reducer reconciles both retained profiles against the prior attribution
and allocator counter. A fresh coordinator replay is byte-identical to the
frozen result. Review checks and exact helper/input hashes are retained in the
packet. The only removed artifact is that redundant replay copy, mapped by
bytes and SHA-256 to the retained canonical result. The helper-generated Python
cache is absent. No worktree or build root was created; unrelated worktrees
and the existing untracked benchmark example remain untouched.
