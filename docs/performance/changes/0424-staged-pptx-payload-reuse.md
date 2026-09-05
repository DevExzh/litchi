# 0424: reuse validated staged PPTX payloads

Source-backed cross-presentation publication prepared the selected image/chart
bytes again even when its immutable plan already owned exactly those payloads.
A before profile found two eight-allocation, 16 MiB clone stacks in the media
lifecycle: initial planning and publication. The candidate preserves planning's
independent staging and shares those bytes after publication revalidates all
metadata and decoded bytes. The second large clone stack disappears.

The control is `6ca9962c7818e173538a40ed45f3a3c32cc1aa6a`; production revision
`d18bf7db4e01f35b8a936892abcb5751e7763d16` changes only low-level OPC shared
payload insertion and PPTX source-backed preparation. Both protocols were
frozen before implementation. Full preparation, source/version/lineage checks,
plan equality, candidate reread, cancellation and conservative reservations
remain. The [design and ADR matrix](../results/change-0424/design.md) and
[source review](../results/change-0424/checks/source-review.md) explain the
ownership and authority boundaries. There is no harness change or parallelism.

The matched matrix uses exact control build artifacts retained from the prior
clean build and a clean candidate build with identical flags. Sixteen fresh
CPU-2 processes retain 1,040 observations from the same plain/media corpora:
100-sample normal and 30-sample allocator lanes, two repeats per role, in
control/candidate/candidate/control order. All corpus and output gates pass.
The [result table](../results/change-0424/matched/result-table.md) preserves
every individual distribution and delta.

Media operation requested bytes fall from 117,608,609 to 100,831,137
(−14.266%), allocation calls from 13,314 to 13,306, and mean region peak live
bytes from 229,300,275 to 212,522,935 (−7.317%). Both allocator repeats agree.
Plain requests/calls are unchanged; its peak increases four bytes. Media
live-after increases 196 bytes and whole-process RSS remains effectively
unchanged. Request volume, V3 absolute operation-region live maximum,
whole-command Heaptrack totals, retained endpoints and RSS remain distinct
scopes. No physical-copy, managed-budget or post-drop reduction is inferred.

All normal repeat statistics pass the frozen drift limits, and no individual
normal timing or RSS regression crosses 5%. Plain median is +0.805% / +0.755%,
an accepted adverse diagnostic observation for the repeatable media resource
benefit. Media timing is mixed relative to control repeat drift; no release
latency claim is made. The [resource decision](../results/change-0424/resource-review.md)
retains every tradeoff and scope limit. Logical reads and output bytes remain
unchanged; there is no source-I/O or decompression reduction claim.

Validation passes 86 applicable Rust tests, including 21 OPC topology tests,
the focused reuse test, 59 public/adversarial source-backed tests and five
unchanged harness oracle tests. Warning-denied rustdoc, crate boundaries,
CRUD-index validation and all nine registered strict claim replays pass.
Strict Clippy has four pre-existing findings in three unchanged files; a
command-scoped diagnostic exemption passes and is not a clean strict gate.
Failed capture, compilation, lint and summary attempts remain retained.

The [portable bundle](../results/change-0424/README.md) contains exact source,
binary and artifact custody, individual samples, allocation stacks, deterministic
replay and report mutation guards. Standalone replay passes all 16 reports,
four traces and 104 mutation probes after original worktree/binary removal.
Explicit caller-drop snapshots, near-limit
budgets, native producers, cold/range sources, bounded scaling and broader CRUD
coverage remain open. The full non-iWork goal remains active and incomplete.
