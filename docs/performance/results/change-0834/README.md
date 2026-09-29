# 0834 — aligned-source filesystem harness repair

This packet starts at `41bb4c9670`, the failed qualification retained in
0833. Production readers and writers are unchanged. The work repairs the
performance harness's treatment of private EOCD alignment padding.

The diagnostic stage is frozen separately from the repaired stage. Its fresh
release build passed. The warm PPTX selected-slide lifecycle passed, while
the cold-only request failed during `verified-prime`, before any timed cold
sample. The original error, raw read records, source snapshot, build identity,
and command receipts remain retained.

`analyze_diagnostic.py` independently reconstructed 883 reads / 169,707 bytes.
The first 881 reads belong to presentation construction. One exact 65,536-byte
tail read at offset 16,953,344 accounts for all 16,230 bytes of unselected-slide
overlap. The two query reads return 615 bytes, including the complete selected
slide's 522-byte compressed payload and zero unselected-slide/media overlap.
The reconstruction's passing status is not a passing workload classification.
The diagnostic neither proves the EOCD transformation nor admits a cold result.

The repaired stage must prove that transformation independently, preserve raw
overlap counts, and classify constructor/query reads separately. OPC cold save
oracles must retain raw route hashes while proving their comment-only
difference. Warm and advisory-cold route byte parity remain required.

Root owns all builds, tests, native runs, and Git actions. The three retained
drivers bind diagnostic, repaired, v2, and v3 source freezes. The initial
repaired source passed the full suite (569 passed, one ignored), then failed
one Clippy style lint. V2 fixes only that predicate and passes all gates,
including nine helper tests; its three OPC save requests fail because the
proof rejected descriptor-bearing members. V3 compares complete local spans,
including descriptors validated by the ZIP crate. Its twelve helper tests,
formatting/check/Clippy/rustdoc/boundary gates, release build, and all seven
qualification commands pass. The full suite is retained from before the
helper amendments, not claimed as rerun on v3.

`replay.py` validates twelve retained reports / twenty-five diagnostic or
qualification samples. V2 remains failed; v3 is admitted with seven reports /
sixteen samples. `audit.py` independently checks custody and the raw diagnostic.
The three mutation-test scripts retain 31 rejected altered claims. All four
failed workloads and the Clippy failure remain in `commands/`.

The prospective formal plan has 72 reports / 2,160 samples. `capture.py` and
`analyze.py` are unexecuted preparation; the formal baseline is deferred to a
separate batch after committing this harness repair. This packet contains no
formal samples or performance improvement claim.

The user-owned format review, API design, and matrix files are outside this
packet. Owned scratch/build roots carry `.owner-0834.json` markers; their
verified removal is recorded in `cleanup.json` before final sealing. iWork is
excluded.

Cleanup removed 6,658 files / 3,203,973,694 logical bytes and retains the exact
three binary descriptors. Post-cleanup replay, all 31 mutation checks, and
the independent audit pass. `seal.py` binds the exact owned paths and verifies
their committed Git blobs.
