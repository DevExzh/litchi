# 0362ul165uh — Post-diagnostic build prerequisites

The first DOCX release build stops before compilation because its sparse
baseline checkout lacks `vendor/flate2/Cargo.toml`. The failed v2 attempt is
preserved. V3 materializes the exact tracked vendor and crate files in both
source checkouts: 205 vendor files and 10,918 crate files per role, with no
missing path dependencies. Baseline and candidate commits, trees, production
code, consumer overlays and locks remain unchanged.

Independent review accepts the narrow v3 change: fresh output paths and exact
source materialization, with the existing commands and guards preserved.
Root launches the four ordinary/allocator release builds serially under
20 GiB, zero swap and one Cargo job, with the toolchain path explicit. At this
checkpoint the baseline ordinary build is compiling. No native execution or
performance acceptance is claimed.

The reviewed fuzz v3 source closure reaches offline metadata validation but
stops because the local registry cache lacks locked `js-sys 0.3.106`. Inspection
finds five missing archives among 170 locked packages: that crate and the
four `wasm-bindgen` packages at `0.2.129`. The other 165 archives pass checksum
validation. V4 separates bounded, locked cache fetches from subsequent fresh
offline metadata and RTF lock validation. Its independent source review
passes; execution remains queued behind the active build. Locks and version
constraints are not changed to bypass the cache failure.

The RTF and Markdown validation handoff now uses the current reviewed gate
runner and chunk helper, an explicit toolchain path, fresh bounded targets,
and the approved workspace lock. It covers format tests, Clippy, an RTF
operations-consumer compiler check, and fresh affected glossary profiles.
Root reviews the concrete command plan and source identities. Execution is
queued; no compiler success is inferred from static review.

## Integration and cleanup

Actual three-way merge previews for all three pending leaves are conflict
free and retain the current ODF consumer and DrawingML preparation files.
DOCX integration must update the current DrawingML manifest pin to
`e98ad3e8d0367c5faba021bec5ee2fcfb48ec2f908846db31fadfae94f03483f`
in the same commit as the new probe registrations. Historical RTF pins remain
unchanged.

The two superseded fuzz source worktrees are removed after preserving their
locks; their packets, receipts and output directories remain. Cumulative
cleanup reaches 288 worktrees and 448 refs. A new concurrent
`uo2/iwa-exit-plan` worktree appears during the census and is retained for
review of its eventual changes. All seventy previously excluded office refs
are unchanged. This does not close the remaining branch dispositions.

## Evidence

Paths below are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `post-diagnostic-build-prerequisites-20261026/receipt.json` | `ce7af7deb9280bf88d465d490d34c11c926eef8013f43cd62de716dbaa8f1a35` |
| `docx-metadata-native-release-compile-9c918-v3-source-closure.json` | `57e3ed232bfb56db56d7de9a5577e1b519fd90fff204a1d3b9ef2f9c402bb0e8` |
| `docx-metadata-native-release-compile-9c918-v3-independent-review-20261026-fresh/review-receipt.json` | `10b81db7e71f0d8dc174a70b12290b032b68da5a748b823cad87a8fae8fc87fb` |
| `fuzz-lock-handoff-7c-20261026-v4-independent-review-20261026-fresh/review-receipt.json` | `e56c0611d423ef27f827673b67f59a995b4a9ada672b102ed03d1ca7881fa25a` |
| `api-glossary-post-diagnostic-20261026-v2-root-review/receipt.json` | `bc15db6dceae7f3020d0d82405f1ae2323b2e5d1578e813f48640c1497cf03af` |
| `pending-api-docx-merge-preview-20261026/receipt.json` | `4539b5f8403ddb42f1d3105c6b7246d85ee852343391aed46d5c57787bd28109` |
| `retired-fuzz-v1-v2-sources-20261026/receipt.json` | `f8ad2e9cf569a5ab7ef7d9d38c36f8bf22e3c25d6748815e90d7806fd33339d6` |

Final source selection, all 29 native groups and the fresh final all-65 gate
remain required. Nothing is pushed.
