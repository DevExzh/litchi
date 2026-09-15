# Evidence: change 0629, attribution of the managed DOCX facade budget test

Change record:
[`0629-facade-docx-budget-test-bisect.md`](../../0629-facade-docx-budget-test-bisect.md).

Disposition: retained, test-only correction and attribution.
`performance_claim: none`. **No library code path was modified**: the only edit
is inside the `#[cfg(test)]` module of `crates/litchi/src/document/doc.rs`.
Nothing in this packet is a timing measurement; every entry is a pass/fail
verdict, a transcript, or exact reservation arithmetic. There is therefore no
A/A floor to report — no latency, allocation or RSS figure is stated anywhere.

## Contents

| Path | What it is |
| --- | --- |
| `bisect/git-bisect-log.txt` | The replayable `git bisect log`: `git bisect start`, bad `c1d2caf85`, good `01936e4f1`, then the eight automated verdicts and the `first bad commit` line. Feed it back with `git bisect replay`. |
| `bisect/verdicts.txt` | One line per leg written by the run script — short SHA, GOOD/BAD/SKIP, subject — with the panic text appended under each BAD leg. Eleven verdicts including the two endpoints. |
| `bisect/bisect-run.sh` | The `git bisect run` script. It copies the repository's gitignored `Cargo.lock` into the worktree, runs `cargo test -p litchi --offline --features docx <test>`, and returns 0/1/125. The 125 guard matters: a commit where the test does not exist or does not compile must be skipped, not silently scored good. |
| `release/decisive-pair-and-head.txt` | `cargo test -p litchi --release --features docx --lib <test>` on three detached checkouts: `de8ee88b0` (last good, ok), `44a471069` (first bad, FAILED), `1946b964e` (current main head, FAILED). Confirms the bisect verdict is not an artefact of the `test` profile. |
| `suite/before-head-featureset.txt` | `cargo test -p litchi --offline --features xls,xlsx,xlsb,ods,opc,docx,pptx` on the **untouched** current main head `1946b964e`: `226 passed; 1 failed`, the same panic site and the same `observed 137474, limit 1154` change 0621 recorded at `1e4198321`. |
| `suite/after-head-featureset.txt` | The same command on this branch: 26 binaries, lib target `227 passed; 0 failed`, totals `313 passed, 0 failed, 7 ignored`. |
| `gates.txt` | The tail of each gate with its exit status, and the pre-existing warnings identified against change 0621's record of the same set. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge. |

## Provenance

Base commit `1946b964e852d05ea3e16735153f8adb7b7d67ec`
(`feat/office-format-completeness`, current main head); branch
`perf/0629-facade-docx-budget-test-bisect`. Host: AMD EPYC 9R45, 32 cores,
123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0, cargo 1.95.0.

Work happened in two detached `git worktree`s on disk under
`/home/zhuhe/code/litchi-worktrees/`, each with its own external
`CARGO_TARGET_DIR` under `.../targets/`: `0629` (the branch) and `0629-bisect`
(the bisect and release legs). The shared working copy at
`/home/zhuhe/code/litchi` was never built in or modified. The worktrees and
target directories were deleted once their transcripts were copied here; see
the cleanup note in `log-sections.md`. The host carried several concurrent
agents throughout, which affects wall time only: no figure in this packet is a
time.

**No measured binary exists**, so no binary sha256 is listed. Every leg is a
`cargo test` verdict on a deterministic test that builds a 1,154-byte package
in memory — no clock, no PRNG, no ambient I/O, no fixture on disk — so one run
per leg is the whole evidence.

`Cargo.lock` is gitignored in this repository. A fresh worktree therefore has
none, and `--locked` fails outright until the file is copied in; `--offline` is
used wherever a leg's manifests predate the head's lock.

## Reproducing the decisive pair

```sh
git worktree add --detach /path/good de8ee88b0
git worktree add --detach /path/bad  44a471069
for leg in good bad; do
  cp /home/zhuhe/code/litchi/Cargo.lock /path/$leg/Cargo.lock
  CARGO_TARGET_DIR=/path/target-$leg cargo test --manifest-path /path/$leg/Cargo.toml \
    -p litchi --release --offline --features docx --lib \
    managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal
done
```

`good` passes. `bad` fails with
`Memory budget exceeded in facade-managed-docx-paragraph-text: observed 137474, limit 1154`,
which is `194 * 32 + 131_072` (the parser workspace `44a471069` added to
`ensure_source_document_xml`) plus the 194-byte managed `PartData` already
charged, against a limit equal to the 1,154-byte package.

## What this packet does not establish

- No latency, allocation, peak-RSS, cold-cache, physical-device, range-source
  or cross-platform result, and no claim of any kind.
- Not that nothing else changed at the bisect commits: each leg ran one test.
- Not that the `xml_len * 32 + 131_072` workspace is correctly sized. This
  packet measures when the fence arrived and what it costs this test, not
  quick-xml's real peak working set.
- Not that changes 0588, 0591, 0592, 0593 and 0594 are correct in general —
  only that none of them caused this failure.
