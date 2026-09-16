# Evidence: change 0657, the XLSX value-only editor's dependency rule

Change record:
[`0657-xlsx-value-editor-d4-admission.md`](../../0657-xlsx-value-editor-d4-admission.md).

Disposition: retained, implemented in `litchi-xlsx` and `litchi-opc`.
`performance_claim: none`. The change widens what the value-only editor admits,
under decision 7 of change
[0652](../../0652-owner-decisions-for-the-third-wave.md), and stops the
source-backed publication audit refusing a replacement for carrying its own
source's formatting, under decision 2 as change
[0654](../../0654-opc-original-bytes-audit-loosened.md)'s finding 5 left it. It
is **not** value-identical, and the three census legs below are what state
exactly how the admitted and publishable sets changed.

The legs are: **before**, the base checkout; **after-1**, the admission rule
alone, which admits 10 packages and publishes none; **after-2**, with the
publication audit, which publishes 9 of those 10. The after-1 files are kept
because the difference between the two after legs is the evidence that the
publication half was necessary and that it changed no admission verdict.

## Contents

| Path | What it is |
| --- | --- |
| `probe/` | The scratch probe, retained in full. It is change 0602's `xlsx-admission-probe` with two subcommands added for this change: `parts`, which asks the editor's verdict for *every* worksheet of a package rather than the first, and `roundtrip`, which runs a complete one-cell plan, commit and publication and then compares the published archive against the source member by member on uncompressed bytes. `Cargo.toml` carries path dependencies; repoint them at the checkout under test. |
| `census/census-before.tsv` | `probe census` over the 95 real fixtures on the base checkout: the editor's verdict per package through both doors, and whether a complete cycle runs. 0 admitted. |
| `census/census-after.tsv` | The same on this branch. 10 admitted through `edit_many`, 3 through `snapshot`. |
| `census/parts-before.tsv`, `census/parts-after.tsv` | The same verdict per *worksheet part*: 0 of 208 before, 15 of 208 after. |
| `census/refusals-before.tsv` | Per fixture, the before leg's verbatim first refusal, beside a deterministic structural table of the whole refused-relationship set. The verbatim message is run-dependent, because `validate_package_relationships` named whichever relationship a hash-seeded `HashMap` walk reached first; this table is what the record's before counts are read from. |
| `census/behind-first-gate-before.txt` | The second-gate census: all 95 fixtures re-censused on the base checkout with `_rels/.rels` reduced to the `officeDocument` relationship alone. Still 0 admitted — the evidence that widening the package gate alone admits nothing, and that the element, attribute and workbook-relationship gates behind it are what refuse the corpus. |
| `census/census-after2.tsv`, `census/parts-after2.tsv` | The same on the after-2 leg. Admission is identical to after-1, checked package for package across both doors: 0 differences on all 95. |
| `census/roundtrip-before.tsv`, `census/roundtrip-after.tsv` | `probe roundtrip` over the 95 real fixtures on the first two legs. No package publishes on either, so the member columns are empty; the after-1 leg's `publish-failed` column carries the ten verbatim publication refusals that motivated the publication half. |
| `census/roundtrip-after2.tsv` | The same on the after-2 leg, with the member columns populated. **9 of 10 admitted packages publish**; each differs in exactly `xl/workbook.xml` and its edited worksheet part, 0 members added, 0 removed. The tenth fails on a non-canonical `xl/_rels/workbook.xml.rels`, which is not a compactness refusal. |
| `census/roundtrip-violations.txt`, `census/roundtrip-violations2.txt` | The violation reports. None on either. |
| `census/reopen-after2.tsv` | Each published archive reopened with `SourceBackedWorkbook::open`: the edited cell read back as the value written on 9 of 9, and the worksheet's stored-extent cell count identical to the source's on 9 of 9. |
| `census/edited-part-diff2.tsv` | For each published package, the edited worksheet part's length before and after, the common prefix and suffix, and the differing middle. One contiguous range on all nine. |
| `census/roundtrip-compacted.tsv`, `census/edited-part-diff-compacted.tsv` | The interim after-1 evidence: the same comparison over the ten admitted packages copied with whitespace-only text nodes dropped, which was the only corpus that could reach publication before the publication half existed. Retained because the record's after-1 leg cites it. |
| `census/determinism-after.tsv`, `census/determinism-after2.tsv` | Five packages republished by two separate processes: byte-identical, 5 of 5 on both legs. |
| `census/admitted.txt`, `census/published.txt` | The ten admitted packages and the nine that publish. |
| `census/summary-before.txt`, `census/summary-after.txt`, `census/summary-after2.txt` | The computed tallies and the refusal-frequency tables the record quotes, per leg. Each carries the caveat that the *message* of a relationship refusal is run-dependent, because the gate walks a hash-seeded `HashMap`; the deterministic structural table is beside it. |
| `bench/summary.txt` | Both measurement rounds in full: per-leg p50, mean, p95 and p99 for every selector in every window, the per-leg `p95/p50` cleanliness verdict with the excluded legs named, the paired deltas in both directions, and the A/A floor of each round. Round A is the one the record quotes — its two legs share a byte-identical harness, so the only difference between them is this change's production code. Round B put the harness census fix in the after leg only and is the control that shows what a one-sided harness change does to the numbers. |
| `bench/counts.txt` | The callgrind isolation pairs: whole-operation instructions per operation on both legs, the changed module's inclusive share, and the publication audit's own attribution before and after. |
| `bench/legs/` | Round A's raw per-sample durations, one file per selector per leg per window, with the harness JSON beside each. |
| `bench/analyze.py`, `bench/extract_counts.py` | The scripts that computed the tables from those files, including the per-leg admission rule. |
| `bench/binaries.sha256` | The SHA-256 of each harness binary timed. |
| `gates.txt` | The tail of each gate run in the worktree. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What this change retained, what it deleted once the evidence was copied, and why. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

Base commit: `ceb0571a8`, the merged head after changes 0654, 0658, 0660, 0664
and 0655; branch `perf/0657-xlsx-value-editor-d4-admission`. The branch was
rebased twice as the wave merged — first onto `ab07e2a47` (change 0654), then
onto `ceb0571a8` — and the census was re-taken after the publication half was
added. The before leg's census was taken at `70d7768cc` and is not re-taken at
the merged head, because change 0654 changes no admission gate and its own
finding 5 records that it leaves the value editor's publication refused; that
is cited rather than re-derived.

The before leg is the shared read-only checkout
`/home/zhuhe/code/litchi-worktrees/before-70d7768cc`; its `test-data` was
verified byte-identical to this worktree's over all 95 fixtures by SHA-256
before any census ran. Both probe binaries were built `--release --locked` with
a `CARGO_TARGET_DIR` outside either worktree.

Probe binary SHA-256: before leg
`49d8d9d0037556d2b9535ee9c2e126b3e6f6c8a88fe025e4a0e6c115f35e8260`, after-1
`58b0e95ada72d8d9fccca9a0a7f9e29de6f7431aedf81a6334f6a66c9e59434e`, after-2
`8613e98bbf1ecca41a1bb11398f495e6e9ab7006fdcf3bc5943aa410d1c31ee5`. Each leg
was built from one probe source, so the censuses are column-comparable; an
earlier before-leg binary (`45881500d9…`) produced the identical tallies and is
recorded in `census/summary-before.txt`. The retained `probe/` source is the
after-2 one, which runs every leg.

Harness binary SHA-256: before leg (a detached checkout of `ceb0571a8`)
`f38bf3c867875d02441b88f7bc5887cd1eae839edaf0f0edcac10c0551caa1c9`; after leg
(this branch at `3e95c7fa2`)
`53c05e2dcafc13b42cddf0c519360398ec28b35e0357a912e4969d8aac1d5bef`; the Round B
after leg (the same commit plus the harness census fix)
`5c09c5db5d2038af281856ef8bedbeba49adbc931361490dbc5949ee6806e67f`. Both legs
were built `cargo build --release --locked` with their own `CARGO_TARGET_DIR`
and staged outside it before any leg ran. The timed after binary is `3e95c7fa2`;
the two changes made after it are a module documentation comment in
`row_visibility` and the harness census fix, neither of which runs on the
measured path, and Round B measures the second of them explicitly.

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0;
valgrind 3.26.0. Every measured process pinned with `taskset -c 12`. Seven
other agents were building and measuring on the host throughout, which is what
the A/A floor in `bench/summary.txt` measures — and why that file also carries
a per-leg `p95/p50 <= 1.05` admission rule, declared before the analysis, after
a window was observed whose end-of-window floor legs looked clean beside a
source leg at +92%.

Corpus: the 95 `.xlsx` files under `test-data/ooxml/xlsx` and
`test-data/office-interop` — change 0602's corpus and change 0587's scope — and,
for the timing legs, change 0601's producer-shaped generated corpora through
`tools/perf-baseline`.

## What is not in this packet

The whitespace-compacted scratch corpus itself and the archives published from
it. They are derived by a documented transformation from fixtures that are in
the repository, they are not the producers' files, and nothing in the record is
claimed for a file as shipped. `census/roundtrip-compacted.tsv` and
`census/edited-part-diff-compacted.tsv` are the outputs those runs produced.
