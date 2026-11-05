# 0359j — W5 OLE owner sharing continuation

- Date: 2026-10-21.
- Base comparison: `0ce4bb8b61` (pre-owner migration) versus the flat-owner
  candidate `9f68596d888522f64f05f4aee526429a03487bec`.
- Status: required performance follow-up; the flat representation is rejected
  and no performance exception is accepted.

The first matched clone qualification exposed a material owner-handle
regression. The complete ABBA clone block measured candidate/baseline ratios of
`6.7134316688` for the small archive and `6.7185001338` for the large archive.
The candidate cloned seven independent `Arc` fields where the pre-wave
`Snapshot` cloned one owner state. This is a rejected result, not a route
acceptance or an allowance to call the difference noise. The runner stopped
before the four route families; the one complete block and the retained
partial attempts remain external evidence.

The raw receipt is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/matched-abba-clone-9f685`.
Its `measurements.json` is 14,943 bytes with SHA-256
`92cc8664841dbcdde3f13129bdda25e7d46a61f73f56f917e1c9b24a6e0cc7e5`.
The flat-run source and binary identities are recorded in that manifest. The
old runner's manifest SHA-256 is `f1035733fe65442ee9e39c3474188994f0d7386736e4e16565d80bae67758862`.

## COW repair

Commit `64f3e82ddb` changes the common owner to one shared `Arc<PackageState>`
with copy-on-write state, preserving shared captured topology and stream bytes
until a mutation. Commit `ad9904197c` removes the remaining default-edit and
no-op detaches: an edit with the retained policy shares the owner state, and an
unchanged commit returns a package sharing that state. Mutating edits still
capture, validate, reopen and publish a fresh state. The focused object suite
passes 41 tests; the complete `litchi-ole-common` test suite passes 227 unit
tests and all integration suites, and all-target performance-diagnostics
Clippy is clean.

The corrected current binaries were built in the external target
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/target-matched-cow-current2`
from `ad9904197c3b1dc79aba6b35c2f7dbb42a9b2190` with one capped Cargo job.
The OLE probe SHA-256 is
`c6376a020ec332d371038de0fe4cde01fd8b9421837b85ef49b7739dcf01e6db`; the
clone probe SHA-256 is
`b083114bbd54cd1381c6ac8efa21ab6327b351d1af51634b3f375fbb6c2ef083`.

## Required qualification

The next paired run must qualify the repaired owner clone and all eleven OLE
selectors in one source-matched harness:

`open`, `begin_edit`, `noop_commit`, `replace_stream`, `read`, `edit`,
`doc_tracked_revision`, `xls_cells`, `xls_cells_stage`, `doc_history`, and
`xls_history`.

Each endpoint is calibrated with the same batched span, every accepted leg is
at least 250 ms, and host snapshots are checked before, after, and during each
child. The run must retain exact-output and forward/inverse replay checks. A
separate counting-allocator capture must keep the 128-step DOC and XLS package
and patch histories alive while reporting total calls/bytes, live bytes and
peak bytes; ordinary read/edit/full XLS-cell routes must be paired as well.
The direct heaptrack invocation and any discarded overlap captures remain
external evidence. The all-eleven-route runner currently hashes to
`1a4640163d743853005581253b77fbae4618a3e23242c44eabeaed5a62787abc`.

Until that repaired paired qualification explains or removes every affected
route regression and supplies the retained-state evidence, this continuation
remains open.
