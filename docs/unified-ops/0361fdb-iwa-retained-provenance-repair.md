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
`/home/zhuhe/code/litchi-worktrees/.coord/constructor-0361f-coordinator`.
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
10,666 bytes, SHA-256
`58e9878bd595edab5ec6d2da15786cc62fd2258ebb626bc5e4284e1cfdb5f0d0`.
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

The constructor still performs one public-field admission pass per message;
the corrected counts show that cost grows linearly with metadata cardinality.
There is no cached header generation, inventory allocation, or provenance
allocation on this route. The constructor performance remains returned for
repair and owner review. No native qualification was run.

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
