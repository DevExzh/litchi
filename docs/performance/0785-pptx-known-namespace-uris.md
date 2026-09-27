# 0785 — exact known namespace URIs in PPTX capture

Retained exact namespace matching avoids repeated UTF-8 validation against six
existing namespace constants. In this warm, in-memory generated corpus, paired
process p50 improves 8.489% for large initial capture and 6.844% for the large
capture/edit/commit/apply/serialize lifecycle. Unknown ASCII and Unicode
namespace controls also improve. Operation memory and allocation counts are
unchanged. All frozen adoption criteria pass. The decision is recorded in
[the packet](results/change-0785/disposition.json).

## Mechanism and compatibility

`notes::resolved` dispatches on byte length and compares against the existing
PresentationML, DrawingML and relationship namespace literals, including Strict
variants. Equality proves that those exact bytes are valid UTF-8. Every other
bound namespace still uses the original UTF-8 decoder; unbound names and unknown
prefix errors are unchanged. There is no new cache, allocation, unsafe code,
dependency or public API.

The [release assembly](results/change-0785/assembly-review.md) shows a compiler
length-indexed jump table and a single `bcmp` for a recognized length. Equal
bytes take the static-string success path; other values keep decoding. This
supports the mechanism, but does not attribute an exact fraction of operation
cycles to the helper or establish results on another architecture.

| Constraint | Evidence |
| --- | --- |
| ADR 0001/0002/0024 ownership and topology | One private PPTX owner; unchanged dependencies; boundary gate passes. |
| ADR 0003 immutable snapshots and publication | No transaction, commit, patch or publication change; existing suites pass. |
| ADR 0005 measured performance | Frozen 15-case paired native and separate allocation matrices; all cases reported. |
| ADR 0006/0008 preservation and refusal | Independent old-resolver oracle, exhaustive substitutions, unknown URI preservation, exact text/source/output checks. |
| ADR 0010/0011 package ownership | No physical-package or archive representation change. |

All 35 recorded architecture and goal inputs remain unchanged. The
[independent static review](results/change-0785/candidate-review.md) verifies
lifetimes, exact matching and the slide-capture call path.

## Paired public workflows

Six alternating process blocks use 30 measured samples after three warmups,
pinned to CPU 12. Separate allocator processes use two blocks of three samples
without warmup. Tiny is 3 slides × 4 text boxes, medium is 12×8, and large is
100×100. Both vendor controls are 12×8 with six unknown namespace attributes
on every text element and six declarations on each slide root. Their URI byte
lengths match the six constants; ASCII values replace initial `h` with `X`, and
Unicode values use a padded `urn:vendor:é` prefix.

Fixture construction and verification remain outside the measured intervals.
`commit` stages one text edit before its clock; `lifecycle` includes capture,
staging, commit, apply and serialization. Ingress uses borrowed bytes and this
fixture takes the full writer route. The nine ordinary source/output identities
match sealed 0780 qualification; historical timings are not pooled.

Displayed milliseconds are medians of per-process nearest-rank p50 values.
Changes are medians of the six paired after/before ratios, so they need not equal
the ratio of the displayed aggregate times. Negative changes are faster.

| Shape | Operation | Before p50 ms | After p50 ms | Paired change |
| --- | --- | ---: | ---: | ---: |
| tiny | capture | 0.2460 | 0.2393 | -2.674% |
| tiny | commit | 0.2182 | 0.2141 | -1.768% |
| tiny | lifecycle | 1.4309 | 1.4174 | -0.939% |
| medium | capture | 0.4972 | 0.4711 | -4.999% |
| medium | commit | 0.3027 | 0.2984 | -1.399% |
| medium | lifecycle | 2.0539 | 2.0203 | -1.865% |
| large | capture | 21.6897 | 19.8068 | -8.489% |
| large | commit | 1.3212 | 1.2965 | -1.824% |
| large | lifecycle | 31.5973 | 29.3441 | -6.844% |
| vendor | capture | 0.5725 | 0.5478 | -4.274% |
| vendor | commit | 0.3311 | 0.3266 | -1.330% |
| vendor | lifecycle | 2.2022 | 2.1673 | -1.659% |
| unicode-vendor | capture | 0.5758 | 0.5530 | -3.987% |
| unicode-vendor | commit | 0.3322 | 0.3266 | -1.581% |
| unicode-vendor | lifecycle | 2.2130 | 2.1772 | -1.664% |

The frozen guard rejects any case with a median paired p50 ratio over 1.05 and
95% bootstrap lower bound above one, or any increase in net-live or
peak-above-entry operation bytes. At least one capture/lifecycle case must
improve at least 3% with its upper bound below one. Bootstrap uses 10,000
resamples and seed 785078. Every case and repeat remains visible; no process
is excluded or retried.
Five cases satisfy the benefit threshold: medium, large, vendor and Unicode
capture, plus large lifecycle. Large capture's after/before 95% interval is
[0.910820, 0.918528]; large lifecycle's is [0.923732, 0.947058]. No latency or
resource guard fails.
The [independent raw review](results/change-0785/results-review.md) reproduces
all 15 ratios and intervals and confirms the source, output and resource checks.

There are 22 process-spread flags over 5%: nine tail-metric flags and thirteen
RSS flags, with no p50 spread flag. Paired p99 exceeds +5% in tiny capture block
3 (+24.372%) and Unicode commit blocks 2/5 (+6.227%/+14.924%). These observations
remain in [all flags](results/change-0785/all-flags.csv); they prohibit a blanket
tail-latency improvement claim. Those families' median paired p99 changes are
−2.416% and −0.005%, respectively. Allocation metrics have no spread or paired
regression flags. Whole-process median RSS changes range from −2.098% to +1.452%;
no RSS reduction is claimed.

There are 180 native processes / 5,400 samples, 60 allocation processes / 180
samples, and 15 separate qualification samples: 255 reports / 5,595 measured
samples. Every report passes exact publication/readback checks, including every
vendor declaration and attribute. Source and output identities match across
before/after legs.

## Memory, verification and limits

All 90 paired allocation samples have equal allocation calls, allocated bytes,
net-live bytes and peak-above-entry bytes. This optimization makes no allocation
reduction claim. RSS is a whole-process gauge including setup and verification,
not the operation-local allocator peak. Full distributions, intervals and flags
are retained in the analysis and CSV tables.

Six production quality gates pass: formatting, all-feature/all-target checking,
all-feature tests, warning-denied library Clippy, warning-denied rustdoc, and
crate boundaries. Tests total 1,235 passed and 3 ignored in 85 suites, including
six new focused resolver tests. These enumerate every replacement byte at every
known-URI position, preserve original XML/invalid-prefix errors, and test borrowed
fallback pointers and same-length Unicode namespaces. Existing scanner and
end-to-end tests remain unchanged.

The probe passes 28 default-feature tests and seven real-allocator wrapper and
fixture tests. An initial mismatched test invocation combined synthetic counter
injections with live allocator instrumentation and failed; the established
separate test protocol resolves that interference without source changes.
Two pre-measurement build failures are retained: an obsolete direct dependency
in the inherited probe lock, and a missing `usize` type annotation in fixture
construction. Both were corrected before successful builds and qualification.
Neither caused a measurement retry or an excluded measurement.

This batch does not certify physical cold-cache, range-source, concurrency,
scaling, native Office producer interoperability, or comprehensive CRUD.
The generated vendor fixtures exercise fallback and preservation, but do not
represent all extension-heavy documents. Small workflows show smaller gains;
no universal 8% improvement is claimed. OLE2/OOXML remain active, ODF deferred,
and iWork excluded. The comprehensive goal remains open.

## Replay and integration

```bash
python3 -B docs/performance/results/change-0785/validate.py
python3 -B docs/performance/results/change-0785/tables.py --check
```

The packet retains frozen plans, source/lock/build identities, failed attempts,
raw reports, release assembly, independent reviews and reproduction drivers.
The seal covers 873 payload files.
Origin is `4d89ccf28d`; successful builds record `2cdbdd3c33`, after two
archive/analysis-only commits in the experiment worktree. Their production tree
is identical to origin, checked by the revision-transition receipt.
All four executable hashes were verified before removing the owned target
(2,686,542,684 file bytes). Integration is complete: `b6046fde88` was fast-forwarded onto
`feat/office-format-completeness`, including its two archive/analysis commits.
All 874 packet Git blobs (873 payloads plus the seal) matched their staged
byte hashes. The owned worktree, branch, copied workspace lock and three
reference links were then removed after exact identity checks. All preexisting
worktrees and the three unrelated local-file hashes remain unchanged. Full
sealed replay, five CSV checks and the independent raw audit pass from the
main workspace after both the original worktree and target are absent.
