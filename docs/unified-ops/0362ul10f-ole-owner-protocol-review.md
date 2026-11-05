# 0362ul10f — coordinated OLE owner protocol review

The coordinator accepts source preparation packet
`89142c8518d6e8d5738a992eb8e8e6063b4bfb2e` after the returns in 0362ul10a–e.
The final one-line repair assigns the selected toolchain PATH to every child,
with a regression that changes ambient PATH after selection. Independent
review had already closed the remaining manifest, executable, compiler,
source, event, artifact, complete-fidelity and allocator-lifetime findings.
The coordinator verifies this final diff and the combined protocol tests.

The packet covers the distinct pre-owner baseline `0ce4bb8b61`, all eleven
OLE/DOC/XLS routes, small/large clone workloads and retained histories. It
keeps the 604822 property-only baseline separate. The actual Rust oracle checks
complete OLE root/storage/stream metadata, DOC paragraphs, XLS cells and opaque
stream bytes, and forward/inverse retained histories. Instrumented histories
remain live through the snapshot. The independent validator binds the actual
manifest policy, selected Cargo/rustc/PATH, exact build output locations and
ordered invocation/artifact evidence. It rejects duplicate or missing row
identities rather than accepting only their counts or marker strings.

Composition into staging is conflict-free. All nine combined Python modules
pass: 89 tests in 24.052 seconds (24.227 seconds including process setup).
The tested source index is `5d8fd9df658e2d13ef76f16168b2f82d8425e455`.
The complete command and source-bound receipt are in
[`results/0362ul10f/composed-tests.json`](results/0362ul10f/composed-tests.json).
The external 423-byte log has SHA-256
`9b887f70dd0a9dc5b55b54adf84aa09d6987dc8eb934751ceca8f4beaeb438f3`.

This accepts the preparation and validation implementation only. No Cargo
build of the new oracle, native benchmark execution, performance result or
final acceptance gate is claimed. Actual baseline/current compile and
correctness checks, quiet ABBA and independent final validation remain required.
Only redaction audit 22 is deferred.
