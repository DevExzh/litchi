# 0529 XML attribute probe pilot

The runtime candidate is rejected under the frozen native timing gates.
Publication allocation instrumentation and four baseline-compatible XML tests
are retained. OLE2/OOXML remains active; ODF is deferred and iWork excluded.

See [change report](../../changes/0529-xml-attribute-probe-pilot.md),
[plan](plan.json), [comparison](comparison.json),
[numeric review](numeric-review.md), [source review](source-review.md),
[harness review](harness-review.md), and [test review](test-review.md).

`baseline/` and `candidate/` bind exact source patches/manifests, normal and
allocator release builds, chronological raw vectors and receipts. `final/`
binds restored production plus retained tests and harness. Preflight attempts
and the must-use test correction are retained. The source/harness reviews
record their original unapplied handoff state; stage manifests and the final
decision establish the later actual disposition.

Replay numeric results with `python3 -B analyze.py --stage compare --output`
followed by an external output path. Replay verification with
`python3 -B verify.py --component all --strict`; it prints to stdout and does
not overwrite sealed evidence. Do not rerun build/capture commands into these
retained paths. Reproduction requires a fresh result directory, owned target,
and the exact stage patch/manifest and pinned standalone Cargo.lock.

The final decision rejects the runtime candidate. All 14 final quality checks
pass with 2,024 successful executions. Cleanup removes every owned campaign
path; post-cleanup replay passes for source, captures, analysis and decision.
`SHA256SUMS` seals the exact file inventory; final verification is read-only.
