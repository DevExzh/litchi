# Change 0178: immutable CFB final-planning fingerprint elision

Date: 2026-08-17

## Decision

Retain a narrow planning specialization for CFB sources opened through
`SharedOleFile::open_owned(Arc<[u8]>, SourceVersion)`. After the initial
complete source/target fingerprint and full candidate reopen, sealed owned
bytes now omit the redundant final complete fingerprint. The same rule applies
after a format-owner callback validates the exact composed splice view.

Arbitrary positional `ReadAt` sources retain the final fingerprint. They may
return a stable version token while changing bytes, so the existing generic
fence remains necessary. The owned proof is private to the CFB ingress; format
callers cannot set a boolean or otherwise mark a mutable adapter immutable.

## Preserved contract

The change does not remove the initial source/target fingerprint, source
version/length and expected-range checks, CFB candidate reopen, selected-stream
validation, format-owner semantic readback, exact no-op callback skip, retained
plan fingerprints, checked composed-view preflight, or publication validation.
Owned direct writes and atomic saves still hash the complete source and target
during emission. Atomic save still flushes and fsyncs the sibling temporary
file, renames atomically, and syncs the parent directory. Generic direct and
atomic publication mutation fences are unchanged.

Two exact counter regressions compare generic and owned planning. Their source
and target fingerprints match, their owner results match, and the generic path
performs exactly `ceil(file_size / 1 MiB)` more logical reads. The existing
dishonest stable-token owner-callback test continues to require
`SourceFingerprintChanged` on the generic path.

## Deterministic work reduction

Each effective owned plan removes one complete logical source scan and one
source/target SHA-256 digest pair:

- 16,913,408-byte CFB owned-overlay corpus: 16,913,408 bytes and 17 one-MiB
  reads;
- 16,995,840-byte XLS comments/Number corpus: 16,995,840 bytes and 17 reads,
  for both one-comment and 256-comment transactions; and
- 202,752-byte XLS RK/MulRK corpus: 202,752 bytes and one read.

These are code- and counter-derived logical `ReadAt` reductions. They are not
physical-I/O, allocation, RSS, cache-temperature, throughput, or decompression
claims.

## Verification

- two focused owned-versus-generic planning counter tests pass;
- five generic stable-token/mutation regressions pass;
- 229 remaining CFB library tests plus all CFB integration/example targets
  pass, with the unrelated pre-existing
  `detected_temp_substitution_is_not_deleted_by_cleanup` identity test recorded
  separately;
- 19 native-XLS source-backed numeric/visibility tests pass;
- CFB, OLE-common, and XLS library Clippy pass with warnings and deprecations
  denied; CFB/OLE-common rustdoc, formatting, and diff checks pass; and
- two independent final reviewers confirm the sealed provenance, generic
  mutation fence, exact read-count proof, documentation, and publication
  boundaries are safe.

The isolated sequential-writer test fails because a removed temporary file can
be recreated with an immediately reused inode on this filesystem; the failure
is present outside this planning diff and no sequential-writer code was changed
in this tranche.

## Clean release A/B/B/A

Control `d84a3d030` and candidate `1535b141e` use distinct exact release
binaries with SHA-256 `94927d403a...` and `af7a105d9f...`. Every leg is clean,
pinned to CPU 2, exposes one logical CPU, and records 20 warmups plus 500
samples for four existing source-backed XLS selectors. All four legs have the
same canonical non-timing projection SHA-256 `e2928393...`, including corpus,
output, sink, splice, fingerprint, and semantic evidence.

Positive values mean lower candidate p50 lifecycle latency:

| Workload | A1 -> B1 | B2 -> A2 | Control p50 drift | Candidate p50 drift | Decision |
|---|---:|---:|---:|---:|---|
| one comment | 33.37% | 30.45% | 1.82% | **6.29%** | latency withheld |
| all 256 comments | 32.80% | 31.43% | 1.87% | 3.95% | latency withheld: p95/p99 drift |
| one Number | 28.95% | 36.47% | **5.02%** | **6.09%** | latency withheld |
| three RK/MulRK values | 23.75% | 34.31% | **8.81%** | **6.26%** | latency withheld |

The paired directions are consistently lower, and the descriptive planning or
commit p50 reductions are larger, but every workload fails at least one of the
predeclared 5%/5%/10%/15% p50/mean/p95/p99 same-implementation stability
thresholds. The deterministic work reduction is accepted; no acceptance-grade
latency claim is made.

## Scope

This adds no selector and leaves the current matrix at 320 cases and the
historical default at 36 cases / 198 records. It does not expand XLS comment or
numeric CRUD, specialize generic/mutable sources, alter topology-changing
publication, or establish resource, physical-I/O, cold-cache, scaling, or
real-producer evidence.

Artifacts:

- [summary](../results/cfb-owned-planning-0178-summary.json)
- [manifest](../results/cfb-owned-planning-0178-manifest.json)
- raw A1/B1/B2/A2 reports listed in the manifest
