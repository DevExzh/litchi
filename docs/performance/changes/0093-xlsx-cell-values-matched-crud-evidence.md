# Change 0093: matched XLSX scalar-cell CRUD evidence

Date: 2026-08-14

The performance harness now has six opt-in XLSX scalar-cell publication
selectors:

- eager and source-backed one-cell set/save;
- eager and source-backed deterministic `ceil(1%)` set/save; and
- eager and source-backed exact-256 existing-cell batch set/save.

Each selector runs against deterministic four-sheet media-rich corpora. The
`medium` shape contains a 48-by-48 grid on each sheet. The `dense-sparse`
shape combines a dense 128-by-128 sheet with three sparse sheets. Both shapes
retain eight deterministic 512 KiB media Parts that are not edited. The source
path uses the selector-first bounded multi-sheet transaction and sequential
overlay publisher exposed by `litchi_xlsx::cell_values`.

The timed interval is deliberately comparable: the sum of open,
selector/stage/commit, and sequential publication segments to a bounded sink.
Source cache-diagnostic sampling is performed between timed segments and is
excluded from the reported duration. Full reopen, semantic
cell equality, package topology and relationship checks, raw media identity,
exact hashes, and source/materialization counters remain outside timing. Eager
outputs check semantic/topology and untouched media payload identity; source
outputs additionally compare raw local and central ZIP records for every
unselected member. The harness also runs exact no-op, clear, and physical-remove lifecycle gates for
the source-backed corpus, but those operations do not become separate timed
selectors. A CRUD-specific source-range helper keeps logical cell counts
separate from ordinary ZIP member counts.

These are selectable evidence controls, not a speedup result. No release ABBA
comparison, allocation claim, RSS claim, or materialization claim is made from
the harness alone. Default behavior remains 36 cases and 198 records; iWork is
deliberately deferred while the `iwa-*` crates are changing independently.
