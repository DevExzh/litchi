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

## Retained-admission repair evidence

The matched-v5 table above is historical evidence for candidate `536d7fdcca`
and remains unchanged. The first repair source commit `846a0fea0b` cached the
exact wire preflight and neutral `ArchiveInfo` totals produced by `Archive::parse`.
Commit `7c2562d992` then removed the repeated minimal `MessageInfo` metadata
walk from the `ArchiveInfo` constructor path. Commit `4fe1e56411` reuses the
single canonical-header probe already required to prove that a public
`archive_info` value still matches its retained source header. It reapplies the
cached neutral dimensions and the exact raw-wire dimensions under the caller's
current profile; profiles that tighten sequence-sensitive message or metadata
budgets still use the full validator. The cache is private provenance from
parsing. Public metadata mutation is detected by the canonical comparison,
payload mutation is checked against the current message bytes, and every
metadata rewrite clears both caches. Unknown fields and noncanonical spellings
remain publishable only when the complete canonical comparison matches.

Commit `2e8b10c098` adds the retained-cache test
`retained_admission_rechecks_public_mutations_and_strict_dimensions`. It
checks each non-zero cached header dimension at one below its admitted value,
then mutates a public message type and payload and verifies the resulting
`HeaderFields` and `MessageBytes` observations. This is a test-only follow-up
to `4fe1e56411`; the release-path source measured below is exactly the latter.
The focused core suite now has 77 passing tests.

The fresh matched-v10 build compares base `ddf2ac8ae2` with candidate
`4fe1e56411`. It uses the committed
`docs/unified-ops/results/0361fb/constructor-qualify.py` builder, the same
probe and lock, release LTO, panic abort and `N=2,000`/`12,000` callgrind legs.
The external build manifest is 74,295 bytes, SHA-256
`391fe5bd7e8f2cb4abcfa42aec632e7261b2cc07a4f378806bf2b05a10046e24`; the
allocation manifest is 51,114 bytes, SHA-256
`bb0e37d2698de7ce761ab74d69d285fb01719391ef825429ba87c46542556653`.
The reduced callgrind summary is 2,735 bytes, SHA-256
`7db6407574a9d48b183d63f64b3972dcad1a755d18d51be6bdd05652b09d8bbf`.
All artifacts and binaries remain under the external
`constructor-0361f-coordinator/matched-v10` directory.

The values below are raw Ir at 2,000 and 12,000 iterations; Ir/op is their
difference divided by 10,000. Output SHA-256 and checksums match within every
row; the retained output digest is
`7bc9a96f894c2361619db11643c80373d56202f07e376de6930ecdadce0146af`.

| Route | Base Ir (2k → 12k) | Candidate Ir (2k → 12k) | Ir/op base → candidate | Ratio |
|---|---:|---:|---:|---:|
| archive_info | 1,594,977 → 2,105,018 | 3,120,540 → 11,530,584 | 51.0041 → 841.0044 | 16.4890 |
| retained_validate | 19,762,856 → 110,432,902 | 20,178,043 → 113,148,089 | 9,067.0046 → 9,297.0046 | 1.0254 |
| retained_encoded_len | 32,904,451 → 189,294,488 | 20,502,140 → 115,092,175 | 15,639.0037 → 9,459.0035 | 0.6048 |
| retained_encode | 95,458,816 → 564,609,013 | 81,837,579 → 483,092,464 | 46,915.0197 → 40,125.4885 | 0.8553 |
| retained_validate_tight | 21,233,054 → 112,573,100 | 21,640,646 → 115,290,692 | 9,134.0046 → 9,365.0046 | 1.0253 |
| retained_encode_tight | 95,153,987 → 556,097,040 | 83,338,006 → 485,462,539 | 46,094.3053 → 40,212.4533 | 0.8724 |

The archive-info constructor remains a separate mandatory-admission cost: its
ratio is 16.489× after the minimal-path repair because the base does no
neutral admission. It is not used to waive a retained-path regression. The
retained validation return is now 1.025× for both ordinary and exact profiles;
encoded length and both publication routes are below the base instruction
count. No native timing is qualified or accepted.

The v10 counting legs report zero retained bytes after each route. For 100
operations, ordinary retained validation changes from 395,600 requested bytes
and 1,900 calls (peak 3,956) to 5,200 bytes and 100 calls (peak 52). Tight
validation has the same counts. Retained publication changes from 2,042,600
bytes and 5,700 calls (peak 14,326) to 1,652,200 bytes and 3,900 calls, with
the same peak. The archive-info constructor allocates zero bytes on both
endpoints in this fixture.

The inline cache's object-size cost was measured with the external layout probe
(`layout.rs`, 5,195 bytes, SHA-256
`e93944a5a6ab2100503c4053066266b5689ebe31e392929a7b26c4d66fd5aa21`) by
cloning prepared retained and non-retained objects and holding the clones.
Peak, live and requested bytes are equal in these held sets; allocation calls
are shown in parentheses.

| Clone count | Non-retained heap base → candidate | Retained heap base → candidate |
|---:|---:|---:|
| 1 | 10,344 → 10,552 (19 → 19) | 10,812 → 11,020 (21 → 21) |
| 8 | 82,752 → 84,416 (145 → 145) | 86,496 → 88,160 (161 → 161) |
| 32 | 331,008 → 337,664 (577 → 577) | 345,984 → 352,640 (641 → 641) |
| 128 | 1,324,032 → 1,350,656 (2,305 → 2,305) | 1,383,936 → 1,410,560 (2,561 → 2,561) |

`size_of::<ArchiveObject>()` is 136 bytes on the base and 344 bytes on the
candidate, an additional 208 bytes per object. The held-clone slopes increase
by exactly 208 bytes per clone for both provenance states; the cache adds no
allocation call. The repair therefore returns retained operation work while
making the per-object inline memory cost explicit.

## Left open

- **Protected native qualification of the repaired constructor routes:** the
  matched-v10 source counts are returned, but a quiet host window is still
  required before accepting native timing.
- **Protected native qualification of the owner routes:** needs a quiet host
  window.
- **`NativePopUpError::Limit`/`NativeReplyError::Limit`:** resolved in
  [0361fe](0361fe-iwork-native-limit-payload.md); both private native errors
  now carry the typed kind and exact observed/maximum counts.
- **Peak heap:** attribute the +5 B process peak on `keynote_media` and the
  footnote insert.
