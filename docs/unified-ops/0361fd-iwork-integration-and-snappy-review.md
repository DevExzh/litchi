# 0361fd — iWork branch integration, retained Snappy review and native status

- Date: 2026-10-22. Work package WP-B (triage 2026-10-22).
- Branch: `uo2/coordinator-constructors`. Parents: 0361f/fa/fb/fc (constructors),
  0361i/ia (footnotes), 0362kb (Numbers Snappy), 0361k (Pages/Keynote Snappy).
- Status: integrated and gated. Two Keynote repairs are included. Owner routes
  improve. The constructor candidate regresses on several routes in both
  instruction counts and unqualified wall-clock, so it is not accepted.
  No protected native block qualified.

## Integration

These four merges bring the iWork branches together. No other branch's
history was rewritten:

1. Trunk `feat/unified-ops-api` (`a18593385c`). In
   `pages/section_text_style.rs` this branch's only change was `ddf2ac8ae2`.
   That commit is patch-identical to trunk's `d0e5542d0e`, so the file takes
   trunk's version, which carries `b34cb1b104`'s registry limit kinds.
   `PROGRESS.md` and the 0362j record also take trunk's text, because
   `4c15580614` supersedes the pre-trunk `61515148e1`.
2. `uo2/w8-footnote-limits`. It is based on `61515148e1`, whose content trunk
   already contains, so the merge applies only the six footnote commits.
3. `uo2/w8-numbers-snappy` (0362kb), with no conflicts.
4. `uo2/w8-snappy-policy`. Its one unique commit is `eaade0241d` (0361k).

The two pre-trunk duplicates remain in ancestry. Removing them would require
rewriting the footnote branch. Their content appears once in the tree, and
`git diff feat/unified-ops-api` shows no PPT or character-style hunk.

The 28 uncommitted Numbers files in the `uo2-w8-snappy-policy` worktree were
inspected read-only. Thirteen are identical to `uo2/w8-numbers-snappy`. The
other 15 are an earlier draft of that branch and contain no unique work:

- They map Snappy failures to `InvalidSource`/`Verification` instead of
  typed limits.
- They lack the `CompressedBytes` limit kind and its `map_core_error` mappers.
- They keep the base's cumulative per-entry ledger in pop-up
  `TransactionBudget`. 0362kb deliberately fixed that ledger.

Nothing was carried over. The worktree can be discarded.

## Retained-policy review (0361k, 0362kb)

The core API has no limit-less entry. `SnappyStream::{compress, decompress,
decompressed_len, maximum_compressed_len}` all require `SnappyLimits`, and
every `SnappyLimits` value, including `Default`, is capped by the hard
ceilings. A scan of production code outside `#[cfg(test)]` covered all 271
Snappy calls in the Pages, Keynote, Numbers and IWA crates. Every call takes
a profile derived from the retained catalog, source or open options, or a
stricter profile built from it. There are two exceptions. Both are the
Keynote frame-scratch estimates documented in 0361k
(`slide_chart_caption`, `slide_chart_value_axis`). They use the bounded core
default only to size one encoder frame's scratch buffer. The remaining
`Default::default()` uses are test fixtures in `litchi-numbers/src/package.rs`
and `litchi-iwa-archive/src/logical.rs`. `check_iwork_leaf_public_api.py` and
the glossary profiles pass.

Typing gap repaired (`a9132d52ad`): seven Keynote owners flattened every
encoder failure to `InvalidSource`:

- audio creation and its metadata writer;
- direct-drawable comments;
- media lifecycle;
- media replacement;
- physical table sort;
- soundtrack items.

Under the retained profile, a caller-selected compressed-entry or
decoded-stream ceiling is an ordinary refusal. A new private `SnappyFailure`
classifier maps core limit and allocation failures onto each owner's typed
`Limit`/`Allocation` variant. Its unit test covers exact and one-under
compressed ceilings and a decoded-stream ceiling. A decode of an
already-admitted source keeps `InvalidSource`, because a failure there means
the source is inconsistent.

0361k regression repaired (same commit): media replacement encoded the
rewritten `Index/Metadata.iwa` with the decode profile, which is tightened to
the source stream length. Any growth of the metadata was therefore refused.
The integration test
`explicit_tiny_media_byte_ceiling_rejects_lifecycle_reads_atomically` failed
on the merged base with `InvalidSource`. With typed errors it reported
`EntryBytes 594 > 592`. 0361k ran only the library suites, which is why it
missed this. The encoder now uses the package's retained profile, and all
1,167 Keynote tests pass.

One gap remains. Numbers' private `NativePopUpError::Limit` (34 sites, with
`NativeReplyError` mirroring it) has no payload. It surfaces as
`Limit { kind: WireWork, observed: 1, maximum: 0 }`. The failure stays typed,
but the kind and counts are lost. This predates 0362kb and is left as a
follow-up.

## Owner measurements (Snappy owners and footnotes)

The base is trunk `a18593385c`. The candidate is `a9132d52ad`, which has the
same library source as the branch head. The same probe
([`owner_probe.rs`](results/0361fd/owner_probe.rs)) is compiled with fat LTO
and `panic=abort` against release builds of each source snapshot. Both use
the same lock (SHA-256 `d85b4300…`) and the same rustc 1.95.0. The probe
binaries are `851a090f…` (base) and `47579bd2…` (candidate); they stay
outside Git. Output bytes are identical across endpoints on every route.

The routes are:

- `numbers_headers`: header edit on `basic.numbers`.
- `keynote_media`: content replacement on `media-replacement-native.key`,
  using `tone.wav`.
- `pages_footnote_insert`: inserts into the synthetic fixture from
  `tests/body_footnote.rs`, dumped once.
- `pages_section_name`: section rename on `basic.pages`.

| Route | Ir/op base → candidate | Ir ratio | Heap bytes/op | ABBA median (spread) |
|---|---|---|---|---|
| numbers_headers | 19,656,090 → 19,053,216 | 0.9693 | 4,179,843 → 3,419,803 | 0.9846 (0.9673–0.9873) |
| keynote_media | 81,908,477 → 81,298,157 | 0.9925 | 7,546,272 → 6,601,160 | 0.9911 (0.9601–0.9948) |
| pages_footnote_insert | 1,545,622 → 1,508,375 | 0.9759 | 183,017 → 144,649 | 0.9905 (0.9852–0.9951) |
| pages_section_name | 19,326,782 → 18,685,452 | 0.9668 | 4,189,472 → 3,392,520 | 0.9791 (0.9779–0.9841) |

How each column was measured:

- Ir/op is the callgrind difference between N=5 and N=1, divided by 4.
- Heap bytes/op comes from dhat, computed the same way. Allocation blocks per
  operation also fall on every route (25,632 → 19,579, 35,297 → 27,213,
  2,131 → 1,672 and 25,768 → 19,430).
- The ABBA column has five blocks per route on CPU 29, with every leg at least
  500 ms. Other agents' cargo/rustc jobs ran during every block (load
  4.7–7.9), so no block is clean. These are observations, not qualified
  timings.

Process peak heap drops for Numbers (−9,091 B) and the section rename
(−9,171 B). It rises by 5 B for `keynote_media` (5,909,141 → 5,909,146) and by
5 B for the footnote insert (32,684 → 32,689). The peak is the same at N=1 and
N=5. Its source was not attributed.

## Constructor matched-v5 (0361f/fa/fb/fc)

The protected binaries `3d3757b4…` (base `ddf2ac8ae2`) and `4eebb855…`
(candidate `536d7fdcca`) match the 0361fb/0361fc evidence. `litchi-iwa-core`
is unchanged from `536d7fdcca` to this head, so these binaries still
represent the integrated source.

- **Protected runner (0361fc).** No attempt started. The quiet-host wait loop
  ([`constructor-v5-wait-loop.sh`](results/0361fd/constructor-v5-wait-loop.sh))
  ran from 03:38 to 03:49. It needed 20 consecutive seconds with no
  cargo/rustc/cc/valgrind process and never found them, because other agents
  kept building. It was stopped; `native-a2` was never created, and attempt
  a1 remains the only one. The protected collector still reports zero
  qualified blocks.
- **Callgrind** ([`constructor-callgrind.sh`](results/0361fd/constructor-callgrind.sh)):
  Ir/op is the difference between N=12,000 and N=2,000, divided by 10,000.
  Outputs and checksums are identical across endpoints.
- **Tolerant ABBA**
  ([`constructor-tolerant-abba.py`](results/0361fd/constructor-tolerant-abba.py)):
  five blocks per route under the same calibration targets, with host load
  recorded rather than aborted. No block was clean, so these runs are not
  qualified.

| Route | Ir/op base → cand | Ir ratio | Tolerant ABBA median |
|---|---|---|---|
| field_path | 36 → 85 | 2.36 | 6.69 |
| field_info | 65 → 195 | 3.00 | 3.84 |
| message_info | 168 → 445 | 2.65 | 3.45 |
| archive_info | 51 → 3,428 | 67.2 | 17.76 |
| archive_object | 13,724 → 3,415 | 0.249 | 0.258 |
| archive_object_reject | 3,440 → 372 | 0.108 | 0.079 |
| parse | 53,467 → 52,606 | 0.984 | 0.971 |
| validate | 9,061 → 5,768 | 0.637 | 0.556 |
| encode | 37,738 → 37,040 | 0.982 | 0.983 |
| retained_parse | 89,534 → 102,156 | 1.141 | 1.204 |
| retained_validate | 9,067 → 18,810 | 2.075 | 2.485 |
| retained_encoded_len | 15,630 → 18,284 | 1.170 | 1.113 |
| retained_encode | 46,915 → 58,476 | 1.246 | 1.289 |
| retained_validate_tight | 9,134 → 42,986 | 4.706 | 4.347 |
| retained_encode_tight | 46,094 → 106,576 | 2.312 | 2.026 |

The instruction counts and the unqualified wall-clock times agree. The four
metadata constructors now do admission work that the base never did. In
absolute terms each call costs a few nanoseconds, but the ratio is large.
The 0361fb retained-header admission makes every retained route slower:

- Ordinary profiles: 1.11–2.5× slower.
- Exact-ceiling profiles: 2.0–4.7× slower. These take the exact wire
  preflight path by design.

`retained_validate` spends its time in four functions:

- `CanonicalHeaderProbe::varint` (27 %);
- `encoded_message_info_length` (22 %);
- `emit_packed_u32` (15 %);
- `emit_canonical_archive_info` (13 %).

These are the canonical-header comparison and admission passes.
`archive_info` spends 70 % in `encoded_message_info_length` and
`check_archive_info_constructor`. Object construction (0.25×), refusal
(0.08–0.11×) and ordinary validation (0.56–0.64×) improve.

The candidate does not meet the performance bar on 10 of 15 routes, so it is
not accepted. The repair options are in "Left open" below.

## Validation

All commands ran in this worktree with the coordinator env. Each exit status
was checked.

| Gate | Exit | Result |
|---|---|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy` on 17 iWork crates, `-p litchi-ops -p litchi -p litchi-py`, `--all-targets --all-features -D warnings` | 0 | |
| `cargo clippy -p litchi-ops --no-default-features`: no features, `pages`, `keynote`, `numbers`, all three | 0 ×5 | |
| `cargo test` on 13 IWA substrate crates, `--all-features` | 0 | 1,990 passed, 223 ignored |
| `cargo test -p litchi-pages --all-features` | 0 | 710 passed |
| `cargo test -p litchi-numbers --all-features` | 0 | 1,086 passed, 6 ignored |
| `cargo test -p litchi-keynote --all-features` | 0 | 1,167 passed, 1 ignored |
| `cargo test -p litchi-iwa --all-features` | 0 | 1,786 passed |
| `cargo test -p litchi-ops --all-features --lib` plus the six iWork targets (`iwork_batch_equivalence`, `iwork_table_cells`, `iwork_unified`, `keynote_unified`, `numbers_unified`, `pages_unified`) | 0 | 1,032 passed |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` | 0 | 14 passed, goldens unchanged |
| `cargo test -p litchi --all-features` | 0 | 209 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` (keynote, pages, numbers, iwa, iwa-core, iwa-archive) | 0 | |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py`, 15 iWork profiles | 0 ×15 | |
| `tools/check_iwork_leaf_public_api.py` | 0 | |
| Five tool unit-test modules | 0 | 1,269 tests |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `check_feature_matrix.py`, `check_crate_boundaries.py`, `non_iwork_gate.py verify` | 0 | |
| `unified_bench_matrix.py --check` | 0 | after adding rows for 0361f/fa/fb/fc/fd/i to `evidence_without_rows` |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 301 tests |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 | after `8bb7c95811`, which refreshed the lock for trunk's PPT `siphasher` dependency |

The 15 glossary profiles are iwa-text-wire, iwa-core, iwa-archive,
iwa-common, iwa-structured, iwa, numbers-minimal, numbers-full, ops-iwork,
facade-iwork, facade-sheet-text, sheet-text, pages, keynote-minimal and
keynote-full. The five tool unit-test modules are `test_officecli_corpus`,
`test_officecli_capability_delta`, `test_no_3rdparty_references`,
`test_check_crate_boundaries` and `test_non_iwork_gate`.

The full `cargo test -p litchi-ops --all-features` suite outside the iWork
targets did not run here.

## Left open

- **Constructor performance repair:** make retained-header admission cheaper,
  for example by caching the canonical-length or comparison result per header,
  or by reusing the ingress admission report. Cut the cost of the
  constructor admission. Then rerun matched-v6 with the protected runner in a
  coordinator-scheduled quiet window.
- **Protected native qualification of the owner routes:** needs a quiet host
  window.
- **`NativePopUpError::Limit`/`NativeReplyError::Limit`:** carry the kind and
  counts.
- **Peak heap:** attribute the +5 B process peak on `keynote_media` and the
  footnote insert.
