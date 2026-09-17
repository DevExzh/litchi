# Evidence packet — change 0663, source-backed OLE2 sector layout

Record: [`0663-cfb-sector-layout-policy.md`](../../0663-cfb-sector-layout-policy.md).

Disposition: retained and implemented. `performance_claim: none`. The packet
proves stream and directory preservation, cutoff migration, deterministic
fallbacks, validated same-length source-backed overlays, and the source-backed
DOC route. Its timing is an engineering measurement with an A/A floor; it does
not claim the 0617 copy-through speedup.

## Contents

| path | purpose |
| --- | --- |
| `decision.json` | machine-readable disposition, scope, evidence and limits |
| `log-sections.md` | four coordinator paragraphs for the shared logs |
| `corpus.txt` | release corpus result and mutation sweep result |
| `layout-picture-doc.txt` | representative source/rewrite layout counts |
| `timing-picture-doc.txt` | three release timing windows with A/A floors |
| `doc-route-picture.txt` | before/after opened-DOC attribution check and current Reuse/Rewrite container control |
| `gates.txt` | validation commands and their results |

## Provenance

| field | value |
| --- | --- |
| base commit | `70d7768cc6dada420ede063f72c88dc99ad30383` (0652) |
| branch | `perf/0663-cfb-sector-layout-policy` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0663` |
| parent integration note | parent branch includes the newer 0659 CFB snapshot fence; this change is writer/layout and source-backed editor wiring |
| host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd 2026-03-21) |
| build | `CARGO_BUILD_JOBS=2`, `cargo test --release --locked` for measurement |
| corpus | repository `test-data/`, 214 OLE2 fixtures found, 211 parsed, 3 skipped |

## Limits

The planner declines directory shape changes, geometry changes, changed
nonzero class IDs, DIFAT layouts and any rejected invariant instead of
guessing; explicit zero class IDs are applied as clears. It retains
bounded layout state, including the packed mini-stream image, but not a second
full copy of the source artifact or regular stream payloads. Equal-length
stream edits with unchanged metadata use a validated positional overlay over
the captured source, then materialize the final output; length changes and
metadata edits use the source-layout writer. Thus the overlay avoids a second
CFB re-layout for unchanged sectors but does not claim zero output writes or
zero final materialization. The ordinary newly-authored DOC writer has no
source bytes and therefore remains from scratch. The corpus uses the
repository's fixtures and does not establish cross-platform or cold-cache
performance.
