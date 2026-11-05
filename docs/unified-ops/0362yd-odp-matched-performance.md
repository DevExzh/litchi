# 0362yd — matched ODP source-owner performance evidence

This record closes the requested matched comparison for the ODP streamed source owner at `12385bc0453978b5406da808d430abae9ce77b5b`. The accepted 0358a changed-route baseline is `e4a0614f5d3a1d768346cc33a4a3b86522dda986`; `f755903f45` is the later broad 0358 record head and is retained only as a secondary attempt. Both probes use the same generated 2,000-slide archive, the same common probe source, one route per process, and one fixed executable path for allocator and Callgrind runs.

## Fixture and reproduction

The archive is 4,215,428 bytes: 591,185 bytes of `content.xml`, 2,000 slides, and one 4,194,304-byte incompressible picture. The source clone is allocated before each measured lifecycle. The small probe and adapters are tracked beside this record: [`odp_0362yd_probe.rs`](results/0362yd/odp_0362yd_probe.rs), [`route_before.rs`](results/0362yd/route_before.rs), [`route_after.rs`](results/0362yd/route_after.rs), [`run_allocator_routes.sh`](results/0362yd/run_allocator_routes.sh), and [`run_callgrind.sh`](results/0362yd/run_callgrind.sh). Their SHA-256 values and the exact build/run commands are in [`validation.json`](results/0362yd/validation.json).

The before route uses the e4a `SourceBackedPresentation` semantic owner and the old `tail_append::Edit::plan` publisher. The after route uses the 12385 source owner, `edit_tail_append().commit()`, `apply_tail_append()` for replay, and `Commit::package().write_to`. e4a predates the Patch API, so its `replay` row is the equivalent independent old publication plan; this is called out rather than presented as patch replay.

## Allocator rows

Each row is one process with the same fixed argv0 and `TMPDIR`; `peak` is peak live heap growth over the lifecycle and `allocated` is cumulative allocated bytes. `output` is the sink length and deterministic SipHash.

| Route | Before peak | After peak | Peak ratio | Before allocated | After allocated | Allocated ratio | Output bytes / SipHash |
|---|---:|---:|---:|---:|---:|---:|---|
| `open` | 693,231 | 116,001 | 0.167334× | 777,196 | 200,325 | 0.257754× | 0 / `15130871412783076140` |
| `first_slide` | 693,231 | 116,001 | 0.167334× | 3,210,706 | 2,714,155 | 0.845345× | 0 / `15130871412783076140` |
| `repeat_slide` | 1,500,637 | 990,891 | 0.660314× | 6,785,233 | 6,369,002 | 0.938656× | 0 / `15130871412783076140` |
| `first_text` | 1,500,637 | 990,891 | 0.660314× | 4,672,388 | 4,175,837 | 0.893727× | 0 / `15130871412783076140` |
| `repeat_text` | 1,500,637 | 990,891 | 0.660314× | 8,673,358 | 8,257,127 | 0.952010× | 0 / `15130871412783076140` |
| `write_text` | 693,231 | 116,001 | 0.167334× | 2,350,086 | 1,853,913 | 0.788870× | 105,778 / `18333823960124859280` |
| `repeat_write_text` | 693,231 | 116,001 | 0.167334× | 3,922,976 | 3,507,501 | 0.894092× | 211,556 / `6976564693146433600` |
| `unchanged_write` | 6,074,765 | 5,484,699 | 0.902866× | 6,782,817 | 6,205,946 | 0.914951× | 4,215,428 / `12221861870545843138` |
| `append_commit_write` | 612,845 | 612,845 | 1.000000× | 5,737,255 | 5,853,743 | 1.020304× | 4,215,645 / `5546059324926701325` |
| `replay` | 612,845 | 612,845 | 1.000000× | 5,737,255 | 5,853,759 | 1.020307× | 4,215,645 / `5546059324926701325` |

All ten output lengths and SipHashes match exactly. The streamed candidate lowers peak heap by 83.27% for open/first slide/write_text, 33.97% for repeated slide/text, and 9.72% for unchanged materialize/write. Append has the same 612,845-byte peak; cumulative allocated bytes rise 2.03% because the candidate opens the validated source owner and creates a Commit/Patch lifecycle, while the e4a route only indexed and planned the insertion. This is a 0.24% Callgrind difference below.

## Per-route Callgrind

Callgrind instruction references use the same fixed binary path, argv route, and environment as the allocator rows.

| Route | Before Ir | After Ir | Ratio | Δ Ir |
|---|---:|---:|---:|---:|
| `open` | 829,899,694 | 830,564,967 | 1.000802× | +665,273 |
| `first_slide` | 886,552,327 | 895,150,022 | 1.009698× | +8,597,695 |
| `repeat_slide` | 954,358,638 | 970,980,884 | 1.017417× | +16,622,246 |
| `first_text` | 899,166,760 | 907,983,019 | 1.009805× | +8,816,259 |
| `repeat_text` | 968,364,698 | 985,472,794 | 1.017667× | +17,108,096 |
| `write_text` | 877,896,502 | 889,723,814 | 1.013472× | +11,827,312 |
| `repeat_write_text` | 925,915,980 | 948,900,371 | 1.024823× | +22,984,391 |
| `unchanged_write` | 845,063,682 | 845,731,874 | 1.000791× | +668,192 |
| `append_commit_write` | 1,348,811,788 | 1,351,943,593 | 1.002322× | +3,131,805 |
| `replay` | 1,348,769,234 | 1,351,952,251 | 1.002360× | +3,183,017 |

The largest deterministic instruction increase is repeated `write_text` at 1.024823×; append and replay are 1.002322× and 1.002360×. These rows are diagnostics rather than native timing qualification.

## Hashes and retained evidence

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| e4a baseline probe binary | 3,048,448 | `6af07f71a4257195fb2b66eb6e34159a5b0d1159250ae3ef431a5748265e225f` |
| 12385 candidate probe binary | 3,541,224 | `37dfcc678f563552af0b07e5073e33bf62537522fe4d2a48fa5fd3bc0e4cf1d0` |
| f755 secondary probe binary | 3,048,576 | `ac4e7def67295292afd0891ec5f9c8101308d53d8a4b6b0f814cb0fb90823490` |

The binaries and Callgrind captures remain outside the repository under `/home/zhuhe/code/litchi-worktrees/.coord/odp-0362yd-evidence/`; the committed receipt records their paths, sizes, and hashes in [`artifacts.sha256`](results/0362yd/artifacts.sha256). No binary or capture is committed.

## Validation

- Both release probe builds completed with `cargo run --offline --release -p litchi-ops --features odp --example odp_0362yd`, one job, and separate owned target directories.
- The fixed-path allocator runner completed all ten routes for both binaries.
- The fixed-path Callgrind runner completed all ten routes for both binaries.
- Output lengths and SipHashes matched for every route.
- Native timing was not run; this record makes no native ABBA claim.

The preceding 0362yc source, encrypted fallback, bounded event-buffer fix, ODP tests, strict Clippy, and rustfmt checks remain recorded in [`0362yc-odp-streamed-reader-review.md`](../0362yc-odp-streamed-reader-review.md).
