# Independent final review — 0916 DOCX commit-publication target reuse

## Decision

The frozen 0916 packet is approved for adoption of the exact three-file
`candidate.patch`. This approval covers the two production files and
`crates/litchi-docx/tests/source_backed_sections.rs` contained in that patch;
it does not adopt any other worktree change. Timing eligibility is evidence
for this review and does not itself change the production checkout.

## Custody and scope

The candidate is based on source revision
`a4ca85e28cdea151176534b7f8dc912253a1ca1e` and the retained 0914 archive
(171,120,640 bytes,
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`). The
independent source audit compared 9,357 regular archive files bytewise and by
mode, found no mode changes, and found exactly these three changed members:

```text
crates/litchi-docx/src/section/layout.rs
crates/litchi-docx/src/source_backed.rs
crates/litchi-docx/tests/source_backed_sections.rs
```

The patch is 10,187 bytes with SHA-256
`71f633cbe26caae0c46ba8d60d58e73e0f4bc350e7226d9cb2d568751ca6d268`.
The source and build receipts match their frozen input descriptors (3,442
bytes, `23bdc32d6e5babad675f384f767c67207ed1e4963c6d627c0aa99f2695911f21`,
and 6,020 bytes,
`196b5cace493c07e0215c7b96f069639ac3040655ac29853ff41ca8e0bff4ea8`). The
external source-tree manifest has 10,794 entries, is 1,311,966 bytes, and has
SHA-256
`1265151d60ccdf64a4aa0dd6adc4126eaf41ca894c97214c5895fa4ab715e8be`.

The candidate executable remains external at
`/home/zhuhe/litchi-performance-0916/candidate/target/release/litchi-perf-baseline`;
it is 69,973,984 bytes with SHA-256
`5857a0e06296634328e3d9f29fafbd3a60baf8faeb7bf148f99d8cd444a7297c`. The
baseline executable is 69,940,744 bytes with SHA-256
`fe56e5192c1198e831583e870beb5e1ba4731fb2077faaffdfc46044bb0ed2db`.
No executable or large raw evidence is part of this review file or the
repository packet.

The frozen native protocol, runner, and unchanged 0914 validator match their
manifest descriptors. Their hashes are respectively
`a2bf571af683795eb9bf15704e77a393097ce7d27cbae3bf4ce5ca215e825050`,
`a914f1a9664f9f7159c840dc9ed8e1dc81d8863d88c6db253096013fa87b0839`, and
`865827fa4898fb664e463759472bad2f83f9a11c446595d3852bcb6ea4793efb`.

## Native evidence

The AB/BA schedule contains eight matched pairs, 16 passed child receipts,
and 3,200 total samples at CPU 12 with 30 warmups and 200 samples per child.
The retained replay passed the same 0914 report validator for every child.
An independent read of the raw external reports reproduced the eight paired
ratios and the seed-916, 10,000-draw nearest-rank bootstrap:

```text
paired median ratio: 1.0741358800619034
95% interval:         [1.059448718604162, 1.0866573203908536]
timing eligible:      true
```

Pooled elapsed statistics were also reproduced independently:

```text
                         p50       p95       p99       samples
baseline             4,840,903 5,001,463 5,104,194       1600
candidate            4,517,335 4,607,341 4,717,232       1600
```

The whole-process maximum RSS medians were 37,646 KiB for baseline and
37,768 KiB for candidate, a 0.324% increase against the 5% review limit.
The pooled p50, p95, and p99 increase flags are all false.

All child guards passed. Across baseline and candidate, source and sink
counters were identical: 202 source reads and 8,409,488 logical source bytes;
138 ordinary-payload reads and 8,397,632 bytes; one ordinary-payload
materialization; the same cache vectors; and 2,101,446 accepted output bytes
with 75 writes and a 65,536-byte largest write. Every output has 2,101,446
bytes and SHA-256
`5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`. The
semantic reopen, preservation, inverse, refusal, source-immutability,
phase-sum, cache, and sink gates are true for all 16 children.

The independent profile replay passed all 36 retained artifacts and its
caller-chain checks. Temporary copies with a falsified native reduction and a
falsified profile binary hash were both rejected; frozen originals were not
modified.

## Quality and limits

The retained quality receipt records exit 0 for DOCX all-targets (1,914
passed, 1 ignored), doctests (78 passed, 31 ignored), formatting, Clippy with
`-D warnings`, warning-denied rustdoc, and the crate-boundary checker (65
packages, 244 declarations, 11 explicit debt entries). The design review and
source/build custody records are present, and the packaging README and
quality-layout record bind the quality-only extras outside the release-source
comparison.

The capture barrier recorded no coordinated heavy process before native
capture. Host-load observations were retained, but unrelated shared-host
services and work were uncontrolled; this packet makes no whole-machine
quiescence claim. The quality receipt likewise certifies the DOCX-focused
checks only, not the whole workspace. These are recorded protocol limits and
do not block adoption of the exact reviewed patch.

The candidate source tree and build cache may be removed after the root
coordinator completes adoption and final custody handling. The external
executable, source archive, receipts, manifests, and diagnostic artifacts
remain the reproducibility record.
