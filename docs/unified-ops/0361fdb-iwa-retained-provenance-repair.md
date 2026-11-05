# 0361fdb — Compact retained IWA provenance and constructor return

- Date: 2026-10-22.
- Parent: 0361fd; source repair commit: `ef233a23eeeb868a23918a77fcdb7d784d6b60c4`.
- Status: correctness and deterministic allocation/instruction evidence recorded;
  native qualification is still unavailable. The earlier 0361fd/v10 evidence is
  historical and remains unchanged.

The retained-header cache now lives in one private `Arc<SourceProvenance>`.
That block owns the original and canonical header bytes in one allocation, the
exact parsed wire and neutral admission totals, and only the two scalar caller
budgets needed to decide whether a tighter profile can reuse those totals. It
does not copy a `Limits` value into each `ArchiveObject`. A freshly constructed
object has no provenance allocation or header comparison. Clones share the
immutable block. Public metadata rewrites clear it; validation still compares
the current canonical projection before reusing it, and stricter profiles run
the exact admission checks. Unknown fields and noncanonical source bytes are
still published byte-for-byte when the projection is unchanged.

The default `ArchiveInfo::new` path has a bounded certificate for the common
minimal shape. It checks the hard message and payload ceilings and accepts only
the fixed-cardinality neutral metadata shape whose worst-case u32 wire values
fit the default header field, nesting, memory, metadata-item and byte ceilings.
Nonminimal records and every strict profile use the exact constructor admission
path. This removes the old repeated neutral walk for the common constructor
without weakening caller-selected limits.

## Matched v11 evidence

The reduced receipt is
[`constructor-v11-repair-summary.json`](results/0361fd/constructor-v11-repair-summary.json),
24,777 bytes, SHA-256
`7e837df4ff05ee3fc6de8b293bfd31e9cc3c3c30410edcc21a317ae3a50f0f65`.
The external raw callgrind captures, allocation JSON, layout outputs and
binaries remain under
`(deleted-coord-artifact)`.
The base is `ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2`; the candidate is this
record's source commit. Both use the same probe source (SHA-256
`e2e16d88f9fa70c621599b9d6ac85fef85d2a1997810be311bcaaeae447366ad`), lock
(SHA-256 `d85b4300c58dbdf0dcc3cb8243ad4ea00cd9c29c9d1d4a77c38e51fe12f708aa`),
release LTO, panic abort and rustc 1.95.0. The callgrind binaries are
`b933b6fb18ef7713efcdb108cd2bcd06ee5bac4210f0998d544b5c092c5130cc` (candidate,
646,792 bytes) and the base identity is in the receipt. The counting binaries
are also identified there.

Callgrind uses the difference between 2,000 and 12,000 iterations divided by
10,000. Output sizes and SHA-256 digests match for every base/candidate route;
the retained output digest is
`7bc9a96f894c2361619db11643c80373d56202f07e376de6930ecdadce0146af`.

| Route | Base Ir/op | Candidate Ir/op | Candidate/base |
|---|---:|---:|---:|
| `archive_info` | 52.0050 | 377.0055 | 7.2494 |
| `parse` | 55,727.0370 | 45,711.4730 | 0.8203 |
| `retained_parse` | 91,985.6412 | 89,642.1457 | 0.9745 |
| `retained_validate` | 9,692.0076 | 9,340.0046 | 0.9637 |
| `retained_validate_tight` | 9,759.0076 | 9,408.0046 | 0.9640 |
| `retained_encoded_len` | 15,605.0046 | 9,455.0046 | 0.6059 |
| `retained_encode` | 48,879.0207 | 40,214.6201 | 0.8227 |
| `retained_encode_tight` | 48,058.3063 | 40,050.6491 | 0.8334 |

The constructor row is a mandatory admission cost and remains returned for
owner review; it is not used to waive any retained-path result. The retained
rows are below the prior base at both ordinary and exact field profiles.

The counting probe reports the first measured operation at `N=1` and repeated
work at `N=100`, with fixture setup outside the timer. For the ordinary and
exact retained validation routes, candidate allocation falls from 3,956 bytes
and 19 calls to 52 bytes and one call per operation. Retained encode falls from
20,426 bytes and 57 calls to 16,522 bytes and 39 calls. Canonical parse falls
from 22,862 bytes and 78 calls to 18,862 bytes and 60 calls; retained parse
falls from 23,376/81 to 19,682/64. `archive_info` allocates zero bytes on both
sides. `ArchiveObject::new` falls from 5,664 bytes and 27 calls to 1,760 bytes
and 9 calls. The N=100 totals are exactly the N=1 values scaled by 100 for
these routes.

The layout probe holds 1, 8, 32 and 128 clones. `ArchiveObject` is 136 bytes on
base and 112 bytes on the candidate. Canonical clone heap is
10,344/82,752/331,008/1,324,032 bytes on base and
10,320/82,560/330,240/1,320,960 on the candidate. Retained clone heap is
10,812/86,496/345,984/1,383,936 on base and exactly the canonical candidate
slope, with allocation calls unchanged at 19/145/577/2,305. The retained
provenance is therefore shared by clones rather than copied into each object.

No native ABBA block ran in this window. These deterministic results do not
qualify native latency or close the constructor return; a quiet-host run is
still required after source review.

## Correctness return and v12 constructor count

The v11 source review found one admission gap: a public mutation of
`MessageInfo.length` could bypass the default minimal certificate on the
decoded validation path. Commit `ada701ba3a` makes that path apply the typed
`MessageBytes` check and adds a constructor/decode boundary regression. Commit
`ed10b876af` then removes avoidable default-profile work: default limits skip a
second invariant walk after exact equality with `Limits::default()`, and the
minimal shape certificate aggregates all public shape fields without changing
the checks. Nondefault profiles still run `Limits::validate()` and the exact
admission path.

The corrected source was rebuilt and measured against the same base using the
external count probe. The reduced receipt is
[`constructor-v12-count-summary.json`](results/0361fd/constructor-v12-count-summary.json),
11,312 bytes, SHA-256
`a53ebbd2e1e59f3e8fdca8be3765bdd85446f4191b848af8e8094adb142cdace`.
The probe source SHA-256 is
`eb3157342b9bd2c5c9e0c5d373ad4c507987b3acf46c680246d84bae7f6efdf0` and
the lock SHA-256 is
`551fd8812350c2b13b5d1791fb8773ba4a50c71801728c3a18f6c1dca3f7f89a`.
Callgrind again uses `(Ir at 12,000 - Ir at 2,000) / 10,000`, with setup
outside the measured loop. `archive_info` is the canonical route; the
`retained_validate` route contains 64 unknown fields and a noncanonical
unknown varint in its retained header.

| Messages | Canonical base → candidate Ir/op | Ratio | Retained base → candidate Ir/op | Ratio |
|---:|---:|---:|---:|---:|
| 1 | 51.005 → 185.004 | 3.627 | 2,127.005 → 2,011.004 | 0.945 |
| 8 | 51.005 → 332.004 | 6.509 | 9,067.006 → 9,340.005 | 1.030 |
| 32 | 51.006 → 836.006 | 16.390 | 48,055.504 → 34,828.005 | 0.725 |
| 128 | 51.004 → 2,852.005 | 55.918 | 195,076.405 → 136,132.005 | 0.698 |

The callgrind call edge for `ArchiveInfo::new` itself is 119, 266, 770 and
2,786 Ir per call at those four cardinalities, an exact `98 + 21 × messages`
fit. The complete canonical route is `164 + 21 × messages` Ir/op; the base
route is a flat 51 Ir/op. The 21-instruction slope is the public shape and
length admission pass. The remaining candidate fixed work is profile equality
and constructor ownership setup; no second encoded-header pass appears.

The constructor still performs one public-field admission pass per message;
the corrected counts show that cost grows linearly with metadata cardinality.
There is no cached header generation, inventory allocation, or provenance
allocation on this route. The constructor performance remains returned for
repair and owner review. No native qualification was run.

## 0361fdc retained-validation repair

The retained validation return was traced to the canonical comparison doing
two metadata walks for every length-delimited field: it first calculated each
nested encoded length and then emitted the same fields into the comparison
probe. `CanonicalHeaderProbe` now reads the cached canonical length delimiter,
checks its exact varint spelling, and uses that length while emitting the
current public projection. The comparison remains exact, so public mutations,
unknown-field provenance and noncanonical source bytes retain their existing
selection and error behavior. The source repair also removes the duplicate
minimal structural check from `MessageInfoCollection::push_new`; strict
profiles still run the per-message check and the aggregate certificate is
checked at publication.

The reduced receipt is
[`constructor-v13-retained-stream.json`](results/0361fd/constructor-v13-retained-stream.json),
2,936 bytes. It records the external binary and callgrind-capture hashes and
the exact command inputs. At eight messages and 12,000 iterations, the
retained-validation candidate is 90,507,877 Ir versus 110,434,841 Ir for the
base (0.819×); output checksum, size and SHA-256 are unchanged. The candidate
binary predates the final `push_new`-only source simplification, which does not
touch the retained-validation route. The earlier generic `push(MessageInfo)`
end-to-end constructor experiment measured 107,689,426 Ir versus 27,558,801
base Ir and is explicitly rejected; the quiet-window constructor measurement
must use `push_new` and include its collection setup.

This is deterministic instruction evidence only. Native qualification remains
pending, and the constructor end-to-end measurement is still open.

The generic collection route now specializes `push(MessageInfo)` when the
current public record still exactly has the conventional fixed version vector
and empty optional metadata. It recomputes the current public length and
applies the caller's current limits before using the fixed-shape totals; a
mutation of length, versions or any nested field takes the existing complete
admission path. This removes the duplicate general metadata walk after
`MessageInfo::new` without trusting a cached length or public field state.
The matched generic and `push_new` end-to-end counts remain pending until the
coordinator releases a build window.

## 0361fdc constructor admission return

Commit `70930f91c3` keeps the public mutation checks above while removing work
that was repeated for every conventional message. `Limits` carries a private
profile marker in an existing bounded field, so default-profile dispatch is a
single masked-bit test without widening `Limits`. `MessageInfo::new` still
checks the current payload length and all strict profiles run their exact
structural admission. Minimal collection admission records exact encoded
message bytes as messages arrive, defers fixed field/memory/item totals to the
final aggregate, and materializes those totals before a nonminimal message can
follow a minimal prefix. The new mixed-prefix regression covers that
transition; public mutation still invalidates the certificate.

The reduced receipt is
[`constructor-v15-push-admission.json`](results/0361fd/constructor-v15-push-admission.json).
The raw captures, allocation JSON and measured binaries remain outside the
repository under
`(deleted-coord-artifact)`.
Callgrind used the matched release probe with setup inside each end-to-end
iteration, `IWA_MESSAGE_COUNT=8`, and 12,000 iterations. Both routes produced
the same checksum, output size and output digest on base and candidate.

| Route | Base Ir | Candidate Ir | Candidate/base |
|---|---:|---:|---:|
| `archive_info_e2e` (`push_new`) | 27,558,858 | 54,839,971 | 1.9899 |
| `archive_info_generic_e2e` (`push(MessageInfo)`) | 27,594,949 | 73,032,058 | 2.6466 |

The candidate still has a returned constructor cost because it performs the
required current-field admission that the historical constructor did not.
The generic allocation probe also retains the caller's input `Vec` while it
moves records into a certified collection: at N=8 and 100 iterations, peak
live bytes are 3,424 on the candidate versus 1,760 on base, while allocated
bytes are 342,500 versus 176,100. `push_new` has equal peak live bytes and
adds one 100-byte collection allocation over 100 iterations. These are
recorded returns, not performance waivers; native qualification remains
pending. The earlier 0361fd generic attempt remains preserved and rejected.

## Validation and workspace hygiene

The source passed `cargo fmt --all -- --check`,
`CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=4 cargo test --offline --locked
-p litchi-iwa-core --all-features --lib` (77 passed),
`CARGO_BUILD_JOBS=1 cargo clippy --offline --locked -p litchi-iwa-core
--all-targets --all-features -- -D warnings`, and
`cargo check --offline --locked -p litchi-iwa-archive --all-features`.
The inactive worktree target cache was removed after confirming no process
used it: 54.0 GiB reported by Cargo, 43.98 GiB recovered at the filesystem
level. Measured v11 targets and all external receipts were retained.

The staged-change size and binary checks were run before commit; no tracked
performance/report file reached 1 MiB and no `/binaries/` path was staged.

## 0361fdc v16 matched constructor return

The v16 protected-binary probe measures the complete constructor and collection
admission work inside every iteration, and records the serialized output
for each matched fixture. The reduced receipt is
[`constructor-v16-push-admission.json`](results/0361fd/constructor-v16-push-admission.json).
It records source commit `cad8523c0c`, the candidate probe source after caching
`message_count()`, and the protected baseline/candidate binary hashes. The raw
callgrind captures remain outside the repository under
`(deleted-coord-artifact)`.

Callgrind uses `(Ir at 12,000 - Ir at 2,000) / 10,000`. `archive_info` is the
original generic `Vec<MessageInfo>` route compared with `push(MessageInfo)`;
`archive_info_e2e` compares the same total construction with `push_new`; and
`archive_info_generic_e2e` includes metadata `Vec` assembly followed by the
generic `push(MessageInfo)` route. Every route produced identical serialized
output bytes and SHA-256 digests for base and candidate at all four message
counts. The semantic checks include public length/version mutation invalidation
and the minimal-prefix to nonminimal transition, so the lower work is not
obtained by trusting mutable fields.

| Route | Messages | Base Ir/op | Candidate Ir/op | Candidate/base |
|---|---:|---:|---:|---:|
| `archive_info` (`push`) | 1 | 51.0 | 794.0 | 15.567 |
| `archive_info` (`push`) | 8 | 51.0 | 2,390.5 | 46.868 |
| `archive_info` (`push`) | 32 | 51.0 | 7,325.5 | 143.620 |
| `archive_info` (`push`) | 128 | 51.0 | 26,397.0 | 517.553 |
| `archive_info_e2e` (`push_new`) | 1 | 1,061.0 | 1,730.0 | 1.631 |
| `archive_info_e2e` (`push_new`) | 8 | 2,172.0 | 4,472.0 | 2.059 |
| `archive_info_e2e` (`push_new`) | 32 | 10,328.0 | 18,435.0 | 1.785 |
| `archive_info_e2e` (`push_new`) | 128 | 51,538.0 | 81,828.0 | 1.588 |
| `archive_info_generic_e2e` (`push`) | 1 | 1,064.0 | 1,982.0 | 1.863 |
| `archive_info_generic_e2e` (`push`) | 8 | 2,175.0 | 5,992.0 | 2.755 |
| `archive_info_generic_e2e` (`push`) | 32 | 10,331.0 | 23,600.0 | 2.284 |
| `archive_info_generic_e2e` (`push`) | 128 | 51,541.0 | 101,031.0 | 1.960 |

The generic collection route remains the largest deterministic return: its
per-message admission slope is required to inspect each current public record
when callers provide an ordinary mutable `Vec<MessageInfo>`. `push_new` avoids
that record inspection while retaining exact aggregate admission, but still
returns a 1.59–2.06× instruction count in this probe. These are constructor
returns for repair review; no native latency or memory qualification is
claimed, and the quiet-host native run remains separate.


## 0361fdc v17/v18 fair raw constructor reclassification

The default-profile admission ledger repair is committed in `93b025031f`. It
keeps the exact current payload length and message count in the collection
ledger, while `push(MessageInfo)` still checks every current public field and
length. Mutable collection access still discards the certificate. The v17
callgrind run below is retained as evidence for that repair, but its direct
`archive_info` ratio compares two different operations: candidate collection
`with_capacity` plus `push` against the historical base `ArchiveInfo::new` over
a raw `Vec`. The endpoint ratio is therefore a collection-cost diagnostic,
not a same-operation regression.

The fair v18 probe adds `archive_info_raw` and `archive_info_raw_e2e` to both
endpoints. Both call `ArchiveInfo::new(identifier, Vec<MessageInfo>, limits)`
(the base adapter uses its pre-limits signature) with the same fixture and
message counts. Collection `push` and `push_new` remain separately reported
in v17; they are not silently substituted for the raw baseline.

### v17 ledger evidence

The probe source was `91f5cdac...` / `1b7f55f8...`, the base/candidate binaries
were `858bc0ca...` / `b8bfda6d...`, and the 48 `.out` captures are outside the
repository at
`(deleted-coord-artifact)`.
The sorted capture digest is
`e97b087cd9c0ad86ded97883ec60e6540b909ae97aa7969b7e7f3a499c2fcd34` and the
captures occupy 5,201,998 bytes. The complete reduced receipt is
[`constructor-v18-fair-raw.json`](results/0361fd/constructor-v18-fair-raw.json).

| Route | Messages | Base Ir/op | Candidate Ir/op | Candidate/base |
|---|---:|---:|---:|---:|
| `archive_info` | 1 | 51.0050 | 792.0055 | 15.5280 |
| `archive_info` | 8 | 51.0050 | 2374.5040 | 46.5543 |
| `archive_info` | 32 | 51.0061 | 7261.5050 | 142.3654 |
| `archive_info` | 128 | 51.0035 | 26141.0036 | 512.5335 |
| `archive_info_e2e` | 1 | 1061.0055 | 1672.0040 | 1.5759 |
| `archive_info_e2e` | 8 | 2172.0040 | 4421.0050 | 2.0354 |
| `archive_info_e2e` | 32 | 10328.0061 | 18408.0040 | 1.7823 |
| `archive_info_e2e` | 128 | 51538.0040 | 81897.0051 | 1.5891 |
| `archive_info_generic_e2e` | 1 | 1064.0055 | 1992.0036 | 1.8722 |
| `archive_info_generic_e2e` | 8 | 2175.0040 | 5988.0046 | 2.7531 |
| `archive_info_generic_e2e` | 32 | 10331.0061 | 23536.0040 | 2.2782 |
| `archive_info_generic_e2e` | 128 | 51541.0040 | 100775.0051 | 1.9552 |

Here `archive_info` is collection `push(MessageInfo)`, `archive_info_e2e`
is collection `push_new`, and `archive_info_generic_e2e` includes metadata
vector assembly plus collection `push(MessageInfo)`. The collection routes
retain exact output bytes and SHA-256. Their ratios are reported as
operation-specific evidence. The direct `archive_info` endpoint
reaches 512.5335× at 128 messages because the historical raw constructor is
constant while the bounded collection admission inspects the caller's public
records. That number is not used as a fair raw-constructor comparison.

### v18 fair raw evidence

The rebuilt probe sources are
`1758a437d86228a15da691053b85f95548c2614557cd39e47070470c03cc0812` (base)
and `f125ce58ace22bbe0da15834f2847e53f48e493b63596efd8c1e505339be7185`
(candidate). The release binaries are
`0fc6c951fbce366b6e4cae0eed9ae38851617dfc07bade1267634a75c36e2cb9` and
`32345cdd8a17ba0fb98697bcc72db6cb15f87bbeb952da4e586f5d901a2ec6e0`.
The 32 `.out` captures are outside the repository at
`(deleted-coord-artifact)`;
the sorted capture digest is
`789b3f55b095e0283498962bacbd20261b2636eb442ff205757501abcf3aa812` and
the captures occupy 3,449,352 bytes.

| Route | Messages | Base Ir/op | Candidate Ir/op | Candidate/base |
|---|---:|---:|---:|---:|
| `archive_info_raw` | 1 | 45.0050 | 136.0050 | 3.0220 |
| `archive_info_raw` | 8 | 45.0050 | 283.0040 | 6.2883 |
| `archive_info_raw` | 32 | 45.0061 | 787.0055 | 17.4866 |
| `archive_info_raw` | 128 | 45.0017 | 2803.0051 | 62.2866 |
| `archive_info_raw_e2e` | 1 | 1056.0055 | 1342.0055 | 1.2708 |
| `archive_info_raw_e2e` | 8 | 2167.0040 | 3860.0050 | 1.7813 |
| `archive_info_raw_e2e` | 32 | 10323.0061 | 17055.0050 | 1.6521 |
| `archive_info_raw_e2e` | 128 | 51533.0040 | 77376.0051 | 1.5015 |

`archive_info_raw` times a prebuilt `Vec<MessageInfo>` and the raw constructor;
`archive_info_raw_e2e` assembles the metadata vector and calls the same raw
constructor inside every measured iteration. Every fair row produced matching
checksums, serialized output lengths and output SHA-256 on both endpoints.
The direct raw route still measures 3.0220×, 6.2883×, 17.4866× and 62.2866× at
1, 8, 32 and 128 messages. This is the cost of current public-record
admission relative to the historical unchecked move and remains a returned
performance issue. The equivalent full metadata-assembly lifecycle measures
1.2708×, 1.7813×, 1.6521× and 1.5015×. These are the corrected lifecycle
comparisons; they do not waive the regression or claim native qualification.

The expected archive is parsed, validated and encoded before timing, and its
exact bytes are checked for every row. Serialization is therefore semantic
validation for this deterministic probe, not a timed native latency claim.
Native latency, heap and allocation qualification remain pending until a quiet
host is available.

## 0361fdc v22 complete authoring lifecycle repair

The retained-header correctness repair is in commit
`d56448c7e1` (`d56448c7e10251b217d973e037352675f5874980`), based on the
original source commit
`ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2`. It keeps the public mutable
`MessageInfo` contract: target-derived nested lengths are used only while the
current projection still matches, and a divergence emits the current payload
and recomputes its length prefix before falling back to canonical encoding.
The earlier constructor optimization in `f44804d2c2` remains part of this
candidate: the default minimal admission predicate widens the public `u32`
length before comparing it with the byte ceiling, and the fixed version vector
uses one extend operation. Strict profiles and nonminimal records still use
the complete admission path; public mutable fields are never trusted after
construction.

The reduced receipt is
[`constructor-v22-complete-lifecycle.json`](results/0361fd/constructor-v22-complete-lifecycle.json).
The v20 and v21 captures and measured binaries remain outside the repository at
`(deleted-coord-artifact)`.
The capture manifests are `f8142abd30ae9c0fe173d2d4756afdb5775c14d25f8999fc55bcd5290eb86f27`
for the 16 v20 output captures and
`211df51863e86e5d2881b129f86aaa3bcf6968f9cd3babae1285a28ad61a3460` for the
32 v21 output captures.

The v20 and v21 performance endpoints were measured before this correctness
repair (`5c2dc283...` candidate source); the v22 source pair above is the
explicit correctness replay. The prior performance endpoints and their
reduced manifests remain unchanged.

The complete used authoring route constructs an `ArchiveObject` and serializes
the containing `Archive` in every measured iteration. It is the relevant
end-to-end lifecycle for the conventional raw-message consumer, and every row
has identical output bytes and SHA-256 on base and candidate:

| Messages | Base Ir/op | Candidate Ir/op | Candidate/base |
|---:|---:|---:|---:|
| 1 | 10,587.0053 | 8,017.0057 | 0.7572 |
| 8 | 55,524.2542 | 39,915.5046 | 0.7189 |
| 32 | 259,512.0048 | 173,925.0052 | 0.6702 |
| 128 | 1,070,983.7656 | 717,469.9051 | 0.6699 |

The fair raw `ArchiveInfo::new(identifier, Vec<MessageInfo>, Limits)` route is
also retained as a capability test. Its direct endpoint measures the required
current-field scan against the historical unchecked move; that diagnostic is
not a whole-lifecycle score. With metadata assembly included, the same route
is 1.1426×, 1.4910×, 1.4285× and 1.3263× of base at 1, 8, 32 and 128
messages, respectively, with matching serialized output. This separates the
unavoidable public-record admission work from the used authoring lifecycle and
does not waive either route for native qualification. The assembled raw
lifecycle remains a returned performance regression at 1.1426x, 1.4910x,
1.4285x and 1.3263x of base; the correctness repair does not qualify or waive
that cost.

The matched attribution run uses the measured f448 candidate and its direct
parent (`57c2ff52...` / `f44804d2...`), with candidate binary
`9a3f63c7...` and base binary `d3ff541f...`. The additional `MessageInfo::new`
Callgrind route is 169.0040 Ir/op on base and 277.0043 on candidate, a
108.0003 Ir/op admission cost per message. The direct raw-constructor deltas
are exactly `34 + 21*N` Ir/op at 1, 8, 32 and 128 messages. The raw end-to-end
route assembles the public records and calls the constructor; it does not
encode in the timed loop. The default raw fast path therefore performs the
required public-field shape and length scan without a duplicate encoded-length
calculation.

The matched allocation probe uses base binary `df765b19...` and candidate
binary `2fc130c3...`, with 100 iterations per row. Across 80 reduced output
rows covering raw, raw end-to-end, collection, `push_new`, and full object
serialization routes, peak bytes, allocated bytes and allocation calls are
identical between endpoints, and every checksum and serialized hash matches.
The Callgrind and allocation manifests are recorded in the v22 receipt; the
captures remain outside the repository under
`(deleted-coord-artifact)`.

The collection comparison is diagnostic because its base adapter retains the
historical constructor signature. On the candidate, `push_new` is cheaper
than generic `push` at every measured message count: generic push must inspect
the public `MessageInfo` shape and length, while the certified collection then
publishes aggregate totals without another `ArchiveInfo` scan. The generic
path remains necessary for callers that can mutate records; conventional
authoring can use `push_new`. These rows and their fairness boundary are
included in the receipt and are not used to waive the fair raw regression.

The consumer audit found no production raw `Vec<MessageInfo>` constructor call
that can become `push_new`: `ArchiveObject::new` already uses its internal
single-pass conventional constructor, while the external core-type sample
deliberately mutates nested metadata and must retain the generic checked path.
The focused core-type route test passed 2/2. The candidate core unit run passed
82 tests with no failures. Each of the three retained-header tests passed on
the original `ddf2ac8a...` baseline and on candidate `d56448c7e1`; the new
packed-field-length mutation regression also passed. The former nested length
mismatch is therefore repaired rather than attributed as a baseline failure.

## Follow-up public-version mutation guard

The default-profile fast path also compares the current public `versions`
vector with `DEFAULT_MESSAGE_VERSIONS`. The earlier predicate checked only the
vector length, so a caller that changed one version value could still enter
the fixed-shape admission branch. Commit `e1da0c08f2` adds the exact value
check and a regression that mutates a version to `u32::MAX` and confirms the
constructor preserves that value while taking the complete checked path.
The test import correction is in `ceb87be1d3`.

The focused core test command was rerun sequentially against that source:
`cargo test --offline --locked --manifest-path
crates/litchi-iwa-core/Cargo.toml --lib`; all 83 tests passed, including the
retained-header and public-version mutation regressions. This is correctness
evidence only. The existing v22 Callgrind/allocation binaries predate
`e1da0c08f2`, and no native timing claim is made for the follow-up source.

## Follow-up exact matched qualification (`ceb87be1d3`)

The follow-up rebuilt both endpoints from the identical probe source
(`e2e16d88f9fa70c621599b9d6ac85fef85d2a1997810be311bcaaeae447366ad`) and
lockfile (`21e7fece767541aa3d8a189f49368dcb5a1d27a7b3ca6d137f93d4452c7d6c65`):
the original `ddf2ac8a...` core as base and the public-version guard at
`ceb87be1d3` as candidate. The reduced receipt is
[`constructor-ceb87-followup.json`](results/0361fd/constructor-ceb87-followup.json).
The complete allocation summary and Callgrind captures remain outside the
repository under
`(deleted-coord-artifact)`.

The allocation run used the same allocator-enabled probe on both endpoints,
100 iterations per route and ABBA ordering. All 15 routes had stable counters
across their repeated endpoint legs, and every equal-iteration pair had the
same checksum, encoded byte length and output SHA-256. Candidate counters were
identical for `field_path`, `field_info`, `message_info`, `archive_info`, and
`retained_encoded_len`; they were lower for the archive-object, rejection,
parse, validation and encoding routes, except for the unchanged retained
encoding peak. The exact binary hashes and all counter tuples are in the
receipt.

The matching Callgrind run used 2,000 and 12,000 iterations per endpoint and
computed `Ir/op = (Ir at 12,000 - Ir at 2,000) / 10,000`:

| Route | Base Ir/op | Candidate Ir/op | Candidate/base |
|---|---:|---:|---:|
| `field_path` | 36.0024 | 83.0036 | 2.3055 |
| `field_info` | 65.0034 | 195.0026 | 2.9999 |
| `message_info` | 168.0042 | 274.0053 | 1.6309 |
| `archive_info` | 51.0050 | 294.0051 | 5.7642 |
| `archive_object` | 13,723.9943 | 3,333.0040 | 0.2429 |
| `archive_object_reject` | 3,440.0046 | 371.0051 | 0.1079 |
| `parse` | 53,467.0360 | 48,204.4730 | 0.9016 |
| `validate` | 9,061.0046 | 3,188.0057 | 0.3518 |
| `encode` | 37,737.7921 | 31,832.9945 | 0.8435 |
| `retained_parse` | 89,533.6377 | 90,051.1457 | 1.0058 |
| `retained_validate` | 9,067.0046 | 7,194.0046 | 0.7934 |
| `retained_encoded_len` | 15,585.0046 | 7,363.0101 | 0.4724 |
| `retained_encode` | 46,978.0197 | 35,921.6190 | 0.7646 |
| `retained_validate_tight` | 9,134.0046 | 7,266.0035 | 0.7955 |
| `retained_encode_tight` | 46,157.3053 | 35,749.6491 | 0.7745 |

The higher metadata-constructor counts are the cost of checking the current
public record; the lower archive-object and wire routes show the optimized
authoring and retained-header paths. These are instruction-count results,
not native latency qualification. Native timing remains pending a verified
quiet host.

The versions-equality gap was confined to the ordinary public `Vec<MessageInfo>`
route in `ArchiveInfo::new`: its default-profile fast predicate previously
checked only `versions.len()`, so a caller-mutated version value could enter
the fixed-shape branch. `MessageInfoCollection::push` already used the exact
`is_minimal_message_info` comparison before updating its private admission
ledger, and `push_new` creates the fixed vector itself. Thus the private
certificate path did not accept a mutated version; the guard repairs the
public raw-Vec validation path and preserves mutated values by sending them
through complete admission.

## Default raw-constructor path fast bound (`bd011f99a8`)

The source-level attribution returned two fixed-constructor costs and one
per-record cost. The default `FieldPath` and minimal `FieldInfo` shapes have
only one caller-controlled dimension: path metadata count. At the hard
default ceilings, that count also proves the derived packed header, neutral
memory, field and nesting budgets. Commit `bd011f99a8` uses that proof only
for the default profile and only within the metadata ceiling; strict profiles
and over-bound inputs retain the former checked path and its error precedence.

The exact matched Callgrind rows changed from +46.0012 to +26.0012 Ir/op for
`field_path`, and from +129.9979 to +80.9992 Ir/op for `field_info`. The
`message_info` delta remains +106.0000 Ir/op: its allocator-enabled probe
still records one 12-byte fallible versions allocation per operation. The
raw `archive_info` validation must inspect every current public record and is
unchanged at +242.9990 Ir/op for the eight-message fixture.

Dynamic exact-count Callgrind separates the raw archive delta into a fixed
and per-message component:

| Messages | Base Ir/op | Candidate Ir/op | Absolute delta |
|---:|---:|---:|---:|
| 0 | 51.0061 | 87.0050 | 35.9989 |
| 1 | 51.0061 | 113.0050 | 61.9989 |
| 8 | 51.0061 | 295.0040 | 243.9979 |
| 32 | 51.0050 | 919.0066 | 868.0016 |
| 128 | 51.0046 | 3,415.0025 | 3,363.9979 |

Within measurement noise this is `36 + 26*N` Ir/op. The debug-symbol source
line accounts for 1 Ir/message at the exact versions comparison, while its
comparator helper contributes the rest of about 5 Ir/message; the eight
current shape fields cost about 6 and the message-length ceiling about 5. The
remaining roughly 10 Ir/message is loop and ownership work. A shape-only
diagnostic that intentionally omits exact version values measured about
`34 + 21*N`; it is not a qualifying candidate, but isolates the exact-value
guard at about 2 fixed plus 5 Ir/message. The public `Vec` can be mutated
before ownership transfer, so these checks cannot be replaced by the private
collection certificate. The reduced receipt is
[`constructor-ceb87-pathfast.json`](results/0361fd/constructor-ceb87-pathfast.json).

The rebuilt allocator probe covered all 15 routes with stable ABBA counters
and matching output hashes. Candidate allocation counts remain identical on
the path, field, message and archive-info routes; object, parse, validation
and encoding routes retain their earlier lower candidate totals. This is
Callgrind/allocation evidence only. Native timing remains pending a verified
quiet host.

## Inline common message versions (`5fd7624a24`)

The final constructor return migrates `MessageInfo.versions` from an always-
heap `Vec<u32>` to `SmallVec<[u32; 3]>`. The common `[1, 0, 5]` sequence is
inline; arbitrary public sequences still spill dynamically, remain mutable,
and retain their serialized representation. Every workspace consumer and test
was migrated in the same commit. The `smallvec` `union` feature is active in
the `litchi-iwa-core` feature graph.

The admission guard now requires both exact default values and non-spilled
storage before using the minimal collection ledger. This closes the equality
gap where a spilled SmallVec containing `[1, 0, 5]` could otherwise look
minimal. Spilled-memory accounting charges the full sequence length, not
`len - 3`; this preserves the existing logical-length budget contract. The
Keynote source-owned heap estimate similarly excludes inline storage while
retaining conversion scratch accounting.

The reduced receipt is
[`constructor-v24-smallvec.json`](results/0361fd/constructor-v24-smallvec.json).
The exact matched targets remain outside the repository:

| target | path | bytes | files |
|---|---|---:|---:|
| base | `(deleted-coord-artifact)` | 753,422,513 | 1,498 |
| candidate | `(deleted-coord-artifact)` | 476,841,172 | 1,128 |

The allocator binaries are 658,296 bytes (base) and 698,704 bytes
(candidate); the Callgrind binaries are 633,744 and 677,464 bytes. Their
paths, SHA-256 values, and the protected pre-SmallVec binaries are recorded in
the receipt. Existing target directories and evidence were retained; no
protected artifact was removed.

Against the exact pre-migration representation, the allocation probe removes
the 12-byte, one-call `message_info` allocation for eight common messages.
For the complete raw constructor at 128 messages, peak/allocated bytes/calls
fall from `28,160/2,816,300/13,000` to `26,624/2,662,700/200`. Callgrind
falls by 1,863,443 Ir at 12,000 iterations for the direct 128-message raw
route and by 677,683,840 Ir for raw end-to-end; the full fixed and count
tables are in the receipt and external TSV summaries. These are matched
allocation and instruction results only. Native timing remains pending a
verified quiet host.

Correctness tests passed: 84 `litchi-iwa-core` unit tests, the focused Keynote
heap-accounting test, and the selected all-target consumer check. Canonical
and retained output hashes matched the pre-migration and candidate endpoints.
