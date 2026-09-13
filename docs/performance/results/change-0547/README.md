# 0547: OLE2 collector attribution and bounded-cycle design

Fresh profiles confirm that visited-map bounds, address calculation, membership
testing and update are a substantial repeated cost in the CFB exact-chain
collector. They account for 20.37% of the XLS owned-source constructor's
instructions and 22.35% of few-large CFB open instructions. This justifies a
separately measured bounded-cycle candidate; no production change or speedup
is admitted in this batch.

| Profile (five timed constructor dumps) | Constructor inclusive Ir | Collector self Ir | Visited category self Ir | Visited / constructor |
| --- | ---: | ---: | ---: | ---: |
| XLS owned-source open | 11,316,345 | 5,601,140 | 2,304,960 | 20.3684% |
| CFB few-large | 10,263,764 | 5,571,945 | 2,293,760 | 22.3481% |
| CFB many-small | 13,975,077 | 764,485 | 286,720 | 2.0517% |
| CFB tiny | 244,672 | 5,200 | 1,680 | 0.6866% |

The visited category is 41.15–41.17% of collector self Ir in the two large
profiles. Sector-vector append/capacity and FAT lookup/marker/loop-state work
remain separately attributed. Entry/preparation self work is small on large
inputs; its memset and growth callees are counted separately, not hidden in
self totals. The first experiment should retain reservation and zero-fill
ordering to isolate per-step work. This category is an optimization target,
not a prediction that all its instructions or that fraction of latency can
be removed. The tiny/many-small rows show why shape guards remain necessary.

One fresh release binary is bound to the committed 0546 source. Four serial
Callgrind children retain 20 positive timed dumps, three distinct CFB setup
dumps and termination dumps. Positive incoming ancestry identifies scope;
exact instruction addresses map to the measured disassembly. Collection-off
call/jump metadata is not an operation-local call count. No native latency,
tail stability, allocation, RSS, I/O, cache, concurrency or scaling result is
claimed from this single-repeat attribution.

The independent finite graph model passes 376,264 cases, including 13,778
cycle cases, comparing exact ordered errors and successful chain order. An
immutable exact-N chain ending in ENDOFCHAIN cannot contain a duplicate.
However, terminal-only replay can turn a short-cycle refusal into declared-N
work: a 32-entry self-cycle takes 34 modeled loop positions versus two for
the current validator (17x). The checkpoint/replay model takes four (2x).
These are model step counts, not Rust timings. A Brent checkpoint strategy
has a constant-factor cycle-work argument; checked finite-width counters,
allocator behavior, scratch reuse and native malformed-input timings still
need implementation-level proof and tests.

The next candidate should preserve initial fallible allocation labels/order,
zero-fill, exact final sector order, authoritative error replay, independent
FAT/MiniFAT scratch and post-collection ownership/physical checks. It should
replace only repeated visited membership work with bounded checkpoint cycle
detection plus terminal proof. It must earn retention under fresh two-repeat
native, allocation, instruction, malformed-input and quality gates. The naive
terminal-only variant is not an acceptable unguarded production shortcut.

The complete source manifest equals the sealed 0546 candidate that passed nine
final quality commands and 1,306 tests; those are reused evidence, not new
0547 test executions. New model, attribution and strict custody replays are
separate. The owned build tree is removed after binary-hash verification and
all evidence is sealed. Accepted ADR/index hashes remain unchanged.

OLE2 and OOXML remain the active performance priorities. ODF is deferred until
that optimization goal completes; iWork remains excluded. The broader CRUD,
provider, memory and scaling requirements remain open.

See [sub-operation attribution](suboperations.json), [design proof](design-review.md), [profile analysis](profile-analysis.json), [protocol](protocol.md), and [ADR matrix](adr-compliance.md). Run `python3 -B docs/performance/results/change-0547/verify.py` for strict replay. Raw output and source-bound disassembly preserve their original whitespace.

An analysis-only finalization mistake removed an intermediate instruction report before sealing. Its original bytes are unavailable; the first suboperation attempt is explicitly superseded. The final canonical reports replay from intact raw profiles and assembly, and the rerun reproduces every numerical suboperation row. [Adaptation record](analysis-adaptation.json) preserves this limitation; no measurement was replaced or recaptured.

The [independent category review](suboperation-review.md) records shared-spill and broad entry-bucket limits. [Sealed replay fix](sealed-replay-fix.json) and its regression checks document the post-cleanup executable-presence correction; final numerical reports remain unchanged.
