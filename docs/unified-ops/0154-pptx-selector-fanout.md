# PPTX selector fanout beyond one native call

The unified PPTX setter previously passed every selected shape on a slide to
one native call, whose maximum is 256 replacements. Text and combined
text/geometry dispatch now split each slide into bounded native calls while
retaining one primary Edit and one final commit. Target collection, duplicate
elimination, property validation and semantic limits still precede staging.
A late native error drops all staged chunks without publishing any revision.

Four regression tests cover 1,000 selected shapes for eager and source owners,
text and mixed text/geometry, immutable inputs, exact no-ops, replay/inverse,
semantic target/byte limits, and a malformed transform in the second chunk.
The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,096 tests/doctests.
The existing focused shape-add and geometry tests also passed (eight tests).

The standalone `tools/perf-baseline/unified-pptx-shape-fanout` harness compares
one unified selector set against equivalent native bounded calls within one
primary Edit. It provides eager/source and text/text+x routes, ABBA samples,
corpus identities and output equivalence checks outside timing. Default counts
are 1, 10, 100, 1,000 and 10,000. This measures selector expansion, not a batch
of separate commands. Each native chunk still reparses the slide; neither
linear scaling nor the goal's 1,000-command batch gate is established here.
Managed geometry admission remains conservative; the 1,000-shape source tests
use the default unmanaged owner. Broader CRUD and complete acceptance remain
open.

## Paired measurement

[Machine-readable evidence](0154-pptx-selector-fanout-results.json) retains all
samples, corpus hashes, build command, source-state digest and binary SHA-256.
The release build disabled debug symbols and incremental compilation. The run
used CPU 0, nine ABBA rounds (18 samples per leg), and counts 1/10/100/1,000.
No build or test ran concurrently with measurement. The 10,000-shape matrix,
allocation/RSS evidence and separate command batching remain pending.

| Owner | Properties | Shapes | Native median ms | Unified median ms | Ratio |
|---|---|---:|---:|---:|---:|
| eager | text | 1 | 0.296 | 0.448 | 1.514 |
| eager | text_x | 1 | 0.329 | 0.477 | 1.452 |
| source | text | 1 | 0.052 | 0.055 | 1.061 |
| source | text_x | 1 | 0.080 | 0.084 | 1.047 |
| eager | text | 10 | 0.406 | 0.599 | 1.474 |
| eager | text_x | 10 | 0.585 | 0.774 | 1.322 |
| source | text | 10 | 0.262 | 0.270 | 1.029 |
| source | text_x | 10 | 0.431 | 0.437 | 1.013 |
| eager | text | 100 | 1.457 | 2.039 | 1.399 |
| eager | text_x | 100 | 2.977 | 3.570 | 1.199 |
| source | text | 100 | 2.346 | 2.365 | 1.008 |
| source | text_x | 100 | 3.855 | 3.871 | 1.004 |
| eager | text | 1000 | 31.689 | 36.060 | 1.138 |
| eager | text_x | 1000 | 83.597 | 88.396 | 1.057 |
| source | text | 1000 | 78.184 | 77.869 | 0.996 |
| source | text_x | 1000 | 130.957 | 130.316 | 0.995 |

All warm-up native/unified outputs matched exactly or by decoded ZIP members.
The eager one-shape paths exceed the 10% small-write budget (45–51% overhead).
This is an open performance finding, not accepted performance. Source samples
are close to native in this run; ratios below one are not claimed speedups.

Regenerate after the README build command (with release debug disabled):

```sh
taskset -c 0 "$BUILD/release/unified-pptx-shape-fanout-baseline" --counts 1,10,100,1000 --rounds 9
```

Facade all-target Clippy and harness Clippy passed with warnings denied.
Scoped rustfmt and diff checks passed. The artifact audit found no prohibited
evidence or compiled binary among 679 changed/untracked paths. Both required
staged-file guards printed nothing. Scoped staging failed because Git could
not create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
on the read-only filesystem; no commit was created.

After all verification processes ended, turn-owned build output, raw temporary
measurement/log files and the generated root lockfile were removed. The small
JSON evidence above retains the raw samples and provenance; no binary is kept
in the repository. No temporary worktree was created. Inherited work and the
base checkout were preserved. The full goal remains active.
