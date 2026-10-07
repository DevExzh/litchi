# 0948 retained evidence — candidate rejected

See the [report](../../0948-xlsx-multi-window-audit-rejected.md). All native
latency gates pass, but both primary cases fail the frozen publication
allocation-call and requested-byte gates in all three repeats. The candidate
is **not adopted**. The production checkout is unchanged by this experiment.

This directory is a bounded review subset of the complete external packet at
`/home/zhuhe/litchi-performance-0948`. It includes the frozen plan, versioned
helpers, build/authorization/review receipts, quality composition, compact
paired results, and an unapplied candidate patch. It is not a standalone
runnable copy: raw child reports, complete source manifests/trees, binaries,
large quality records, and `analysis.json` remain external. Every committed
file is smaller than 1 MiB; no compiled executable or `binaries/` file is
committed. Build commands, source identities, binary sizes, and SHA-256 hashes
are retained in `builds/`, `build-tools/`, and the plan.

The raw reduction `analysis.json` is 5,815,204 bytes, SHA-256
`ff73db31d7374f921bc812b35b39f96ec772a9abd939fe132e3591f196334ac9`.
`reduced/summary-v1.json` preserves all paired native metrics, confidence
bounds, allocator ratios, and flags. `reduced/rejected-candidate-v7.patch`
applies cleanly to baseline `b6c5c80ffbcb60736d6c4c2f3b590838640008a9`,
but is intentionally not applied. Its source descriptor records exact final
file hashes and the test-only v5-to-v7 runtime reuse bridge.

## Reproduce the analysis

In the complete packet, use fresh output paths; helpers refuse overwriting
retained outputs:

```sh
python3 -B tools/reduce_0948_v4.py all --output /tmp/0948-analysis-replay.json
python3 -B tools/summarize_0948_v1.py --output /tmp/0948-summary-replay.json
cmp reduced/summary-v1.json /tmp/0948-summary-replay.json
```

At the original packet root, the frozen reducer independently verifies raw
report statistics/semantics, source
censuses, build/binary bindings, and terminal custody, then applies the frozen
gates. Successful tool exit with JSON `status: fail` is the expected result:
it represents a valid rejected experiment, not a reducer crash. The summary
replay is byte-identical. The v9 independent review recomputes gates from
reduced sample rows; it is not a separate raw-report parser and its historical
script uses the original packet path.

After relocation, use the additional replay adapter (provided as a post-seal
supplement) instead of invoking the frozen reducer directly:

```sh
python3 -B tools/replay_0948_v2.py all --output /tmp/0948-moved-analysis.json
```

The frozen reducer reconstructs absolute command paths from its current root,
so it rejects the historical invocation paths when moved. The adapter binds
exact reducer, plan, and root-authorization hashes, renders the original
expected argv/cwd, and leaves physical source/report/binary reads at the new
location. It does not edit captures, relax gates, or rerun workloads.

## Reproduce builds and captures

Exact original wrappers are retained:

```sh
python3 -B build-tools/build_candidate_native_v3.py
python3 -B build-tools/build_candidate_allocator_v3.py
python3 -B tools/capture_0948_v3.py qualification
python3 -B tools/capture_0948_v3.py native
python3 -B tools/capture_0948_v3.py allocator
```

These are provenance commands, **not commands to rerun inside the sealed
packet**. Source/binary/terminal-bound create-only paths already exist. A new
measurement needs a separate packet with fresh prospective bindings and the
same schedule/gates if it is to be compared. Build wrappers bind the complete
environment, target directory, Rust 1.95.0 toolchain and release overrides;
using an unqualified current Cargo build is not an equivalent regeneration.

## Corrections and limitations

The plan still declares reducer v3. Its retained qualification reduction fails
on differing source-summary schemas; correction receipt
`receipts/reducer-v4-source-census-correction-v1.json` explicitly binds v4.
V4 fully hashes both role trees and validates exact full/focused summaries.
No child is rerun and no gate changes. Earlier helper, build, and quality
failures remain external. Historical draft recovery and source-only corrections
are separately recorded; the final result must not be presented as a clean
first-attempt execution history.

Quality is 26 raw passes and 4 documented baseline failures, with a separate
4/4 focused v7 XML bridge. Full details and known Clippy/formatting drift are
in the report and quality composition. The measured runtime is source v5;
the unapplied final patch is v7, differing only in one uncompiled test file.

## Archive and cleanup

The archive is
`/home/zhuhe/litchi-performance-0948-archive/archive-0948-v1.tar.zst`,
**1,348,106,492 bytes**, SHA-256
`15a4008a6b888fed4fea50f0746d9787df32e1a2df6b164d25a7dabb19f1db95`.
Its 217,466 selected files total 7,445,754,580 bytes. The external manifest is
bound by `archive/archive-receipt-v1.json`; the archive helper and its invocation
provide regeneration provenance.

Verified extraction, relocated replay, and cleanup descriptors are recorded
in the archive supplement accompanying this review subset. The archive
preserves historical source trees, failures, raw measurements, role binaries,
and the original v3 release executable paths required by build-terminal reuse.
Only disposable build/replay intermediates are eligible for cleanup after
successful verification. Large inventories remain external and are represented
here by size, hash, and regeneration commands.

The verified post-seal bundle is
`/home/zhuhe/litchi-performance-0948-archive/postseal-supplement-v1.tar.gz`,
**15,417,955 bytes**, SHA-256
`0f493c88235c0aab8cd4edbc454d53e95541710c63e679e5ee36006e45693b3a`.
Its 59 files include the replay adapter, final independent review, preserved
78,976,976-byte qualification file-syscall trace, verification/cleanup evidence,
and later host/draft records. Every bundled file was read back and checked
against its size, mode, and digest. The supplement's `packet/` prefix maps to
the extracted main packet; `archive/` holds supplementary archival evidence.
Copy the supplemented replay helper into the extracted packet's `tools/`
directory, preserving its bytes. Extract the main archive with exact modes
(`tar --same-permissions`) into a fresh directory before invoking replay.

The main archive's external manifest is 54,217,762 bytes, SHA-256
`fe29a9f975907cad76a455f60397c7c211a9dc91805b3e8d15ab0e54ec9c21af`;
it is included inside that archive. Original archive generation used
`archive-0948-v1.py --root /home/zhuhe/litchi-performance-0948 --archive-root
/home/zhuhe/litchi-performance-0948-archive --archive-name archive-0948-v1.tar.zst`.
The Python invocation and supplement command are retained in the corresponding
helpers/receipts. Later packet additions are explicit supplements, so running
the collector against today's expanded packet is a new archive, not the old seal.

Cleanup removed 28,826 owned target files totaling 51,850,203,420 logical bytes,
while verifying both required historical executables unchanged. It also removed
all three owned replay trees and the interrupted redundant v2 tar after
preserving evidence. Logical byte totals include hard-linked scratch copies
and do not claim physical disk space reclaimed. V1 remains authoritative; the
interrupted v2 attempt and its deletion are disclosed separately. Production
and unrelated work were preserved. No source rebuild or benchmark rerun was
performed after rejection.
