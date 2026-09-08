# Change 0472: omit owned tags for plain worksheet cells

`performance_claim: none; diagnostic whole-process allocation reduction`

`claim_authorized: false`

The retained candidate `3eac9a493` uses `Option<Tag>` for ephemeral lossless
worksheet snapshots. Only exact unprefixed `c` cells with no attributes or a
sole unqualified `r` omit owned tags. Address checking remains first; every
attribute is still checked and normalized. Rich, prefixed and metadata cells
retain the original representation. Changed cells regenerate the same `r`;
untouched cells copy original source spans. No public API, persistent cache,
unsafe code or Store handoff limit changes.

Matched whole-process Heaptrack allocation calls decrease from 32,012,510 to
27,424,990: 4,587,520 fewer calls (14.330%). This meets the frozen allocation
retention gate. Temporary allocations increase from 6,693,555 to 7,611,055;
a normalized coordinate remains temporary instead of becoming retained tag
state. Both rounded peak heap displays remain `104.38M`. These totals include
generation, expected output, warmups, verification and teardown; they are not
operation-local counts or exact peak bytes. Instrumented time and RSS are
excluded from normal comparisons. No peak-memory reduction is established.

Both roles are fresh same-path clean builds: control `a82367b52`, candidate
`3eac9a493`, Rust 1.98.1 with frame pointers and unwind tables, CPU 2 and one
worker. Both source inventories contain 6,993 files with exactly five changed
XLSX files and identical fixtures. The frozen protocol captures seven-row
100/5 ABBA, the complete 201-row 15/3 guard, and dense one-percent 5/1
Heaptrack. Existing registered latency claims still require 500 samples;
this batch adds none.

Median elapsed milliseconds, diagnostic only:

| Case | Shape | A1 | B1 | B2 | A2 |
| --- | --- | ---: | ---: | ---: | ---: |
| ppt_fresh_write_to | payload-heavy | 4.441573 | 4.881926 | 4.357172 | 4.430737 |
| xlsx_one_cell_commit_save | dense-wide | 167.128756 | 159.860859 | 157.517343 | 166.128263 |
| xlsx_one_cell_commit_save | medium | 2.419296 | 2.289932 | 2.266962 | 2.404057 |
| xlsx_one_cell_commit_save | tiny | 0.218492 | 0.209846 | 0.207046 | 0.216041 |
| xlsx_one_percent_commit_save | dense-wide | 347.600256 | 317.495379 | 317.268803 | 347.503763 |
| xlsx_one_percent_commit_save | medium | 9.575174 | 9.217749 | 9.009362 | 9.511861 |
| xlsx_one_percent_commit_save | tiny | 0.395722 | 0.376047 | 0.372207 | 0.390877 |

All six ordinary XLSX medians are lower in both matched pairs. Payload-heavy
PPT is 9.91% slower in B1/A1 but 1.66% faster in B2/A2; the adverse first
pair is retained and does not reproduce in the second pair. Normal RSS in
ABBA order is 117,216 / 117,040 / 117,088 / 117,028 KiB, with paired changes
of -0.150% and +0.051%. Full-guard RSS is 153,320 / 152,804 KiB. These small
mixed differences do not establish an RSS improvement.

The full guard compares 1,205 metrics over 201 rows and retains 86 latency
policy flags, with zero non-latency flags. An explicit five-percent trigger
finds 153 mean/p50/p95/p99 cells over 62 rows; nine rows exceed it in all four
statistics: CFB borrowed creation (wide-root), CFB shared reads (few-large,
wide-root), OPC no-op save and cached main read (many-small), OPC source open
(few-large), XLSX source first cell and sheet listing (medium), and ZIP read
(few-large). These short-run regressions are unresolved descriptive findings;
there is no uniformly regression-free or general latency claim. The retained
change is scoped to its allocation result and exercised correctness.

Six new differential tests cover plain eligibility, rich/prefixed fallback,
exact writer parity, attribute/error priority, scanner integration and the
Option<Tag> size guard. The final focused suite passes all 16 tests. Initial
raw-string fixture and duplicate-error-wording failures are retained. The
full XLSX suite passes 1,263 tests across 59 targets. Formatting, workspace
all-feature checking, XLSX Clippy/rustdoc and crate boundaries are the six
recorded gates. Evidence verification additionally checks fresh role builds,
source/binary identities, exact build/capture chronology and portable seals.

Independent evidence review found an inherited chronology check that assumed
both builds preceded A1 and indexed a two-element interval incorrectly. The
verifier now enforces the actual frozen control-build/A1/candidate-build/B1
sequence, complete capture order and non-overlap, with a regression test.
Raw measurements and protocol were not changed to satisfy the verifier.

See the [sealed bundle](../results/change-0472/README.md),
[source review](../results/change-0472/source-review.md) and
[next work](../results/change-0472/next-work.md). Prefix-rich and attribute-rich
populations have correctness coverage but no dedicated timing measurements.
No native Office, new fuzz, cold/range, streaming or scaling claim follows.
The wider non-iWork goal remains open.

All six gates and ten evidence tests pass. Live verification passes before
removing the authenticated temporary builds. An initial strict registry check
failed with filesystem quota exhaustion; its receipt is retained and the
retry follows cleanup, without changing claim definitions or evidence.

The strict registry retry passes all ten existing claims after cleanup.
