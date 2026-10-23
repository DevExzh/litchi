# 0360adc — XLSX raw aggregate admission ordering

Status: implementation candidate on `uo2/0360adc-raw-aggregate-sparse`.

Raw slicer and timeline insertion now census the current package ledger from
output-free relationship plans before capturing source XML or staging a
proposed relationship edit. The baseline check covers materialized part count
and bytes, relationship-member count and bytes, relationship events, and
relationship count. The original `RelationshipSourcePlan` remains separate
from the proposed `RelationshipsEditPlan`; the package is scanned once and
the proposed owner is applied as a delta.

The final workbook rewrite helpers admit their exact replacement size through
the same operation-local ledger before allocating the output `Vec`. Slicer
and timeline output still copies the original bytes around the exact
integration span, and the existing post-plan package mutation remains
atomic. Prepared cache payloads and relationship identity checks are
unchanged.

The ledger now uses the OPC `PartMetadata::payload_len()` projection for both
the baseline census and replacement subtraction. An owned-source package can
therefore refuse a tight raw aggregate policy from its admitted central
directory sizes without inflating deferred members. The deferred payload keeps
that checked size, bounds the ZIP decoder with it, and verifies both retained
metadata and decoded length before charging actual bytes.

The lazy OPC regression also covers a member whose central-directory and local
headers declare one more byte than the stored payload contains. Its first read
uses the ZIP reader's bounded `declared + 1` overrun probe and refuses the
decoded-size mismatch before publishing or charging the payload. The `+1` is a
bounded corruption probe; `PartBytes` remains a logical payload limit, while
the execution memory budget governs the decoder allocation. The new deferred
length field and retained-metadata check are recorded here without a
performance qualification until matched owned-open and first-read allocation
and instruction measurements are available.

The focused regression coverage includes an already over-limit package
baseline, a lazy-package baseline that remains cold under refusal, the admitted
length projection over a corrupt cold member, a declared-versus-decoded length
mismatch, and a canonical slicer edit that rejects that malformed lazy worksheet
without changing the source package. Slicer and timeline rewrites
where the new cache part fits the aggregate but the rewritten workbook does
not. Rustfmt and `git diff --check` are the static checks run for this
candidate. Cargo validation is deferred under the coordinator's build and
disk hold; no performance qualification is claimed here.

The source-only lifecycle harness was refreshed in this candidate. It compares
`0befb112d96d2409e6f40c9abfeb4813665905ad` with
`5726f03c7db2dd7dd539000b45f109f14ea8c33a` through one API and now covers
owned open, the first selected lazy read, raw slicer and timeline insertion,
typed replacement, commit/replay/inverse, stale-patch refusal, no-op, native
IDs and paths, untouched worksheet bytes, and deferred decode counters. The
existing perf-baseline normal and allocator binaries dispatch the selector;
the normal target records elapsed samples and unavailable allocation status,
while the allocator target records operation-scoped allocation samples. The
existing coordinator target root at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets` is available for a
later assigned run; this commit creates no target or evidence artifact.
