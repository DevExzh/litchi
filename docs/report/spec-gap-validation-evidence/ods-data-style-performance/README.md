# ODS data-style performance smoke

This directory records a bounded correctness and allocation smoke for a
historical, pre-correction source-qualified ODS data-style candidate. It is
evidence for that frozen candidate only; it is not evidence for the current
dirty worktree, a final owner-admission implementation, or a before/after
benchmark. It makes no latency, throughput, RSS, or speedup claim.

## Scope

The harness exercises two deterministic synthetic ODS packages containing 8
and 512 automatic `number:number-style` roots. The opaque package retains a
foreign child, CDATA payload, and comment in every style. Each scale runs:

- `read_catalog`: package open plus the public `Snapshot::data_styles` read;
- `metadata_patch`: one public metadata patch, with opaque bytes checked after
  reopening the changed package member;
- `graph_insert`: one public extended-style graph insertion;
- `graph_replace`: one public same-family graph replacement against a clean
  typed fixture, because the opaque fixture is intentionally outside the
  replacement envelope;
- `exact_noop`: a default metadata patch, asserting byte-identical package
  output.

The allocator observer counts calls and requested/released bytes from the
process global allocator. Counts include package ingress and all work in the
operation. They are useful bounded observations for this harness, not an
allocator-independent library claim. `elapsed_us` is retained for future
comparison only and must not be read as a performance result without a fixed
baseline, repeated samples, and host-contended measurement policy.

## Reproduction

The harness is `harness/main.rs`, with the minimal manifest in
`harness/Cargo.toml`. From the repository root, use a separate target and a
non-`/tmp` temporary directory:

```text
TMPDIR=/var/tmp CARGO_TARGET_DIR=/var/tmp/ods-style-profile-target \
  cargo run --manifest-path docs/report/spec-gap-validation-evidence/ods-data-style-performance/harness/Cargo.toml \
  --locked --offline --quiet
```

The run writes deterministic fixture archives to
`/var/tmp/ods-style-profile-fixture-{8,512}.ods`; their hashes are recorded
below. The compact receipt is in `results.jsonl`; unmodified first and second
raw receipts are `raw-before.jsonl` and `raw-after.jsonl`, with their empty
stderr captures beside them.
The same command was run a second time. `repeat-check.txt` records the
comparison: all allocation, byte, output, and correctness fields were stable;
only the unclaimed elapsed-time observation varied.

For an exact historical-candidate replay, create a worktree at `base_head`,
apply `frozen-source.patch`, copy `frozen-Cargo.lock` over that worktree's
workspace `Cargo.lock`, and copy this directory's `harness/` directory into
the same relative path in that worktree. Restore the retained
`harness/Cargo.lock` if the copy operation creates or changes it. Then run the
command above with that worktree as the repository root. The harness manifest
uses only the relative `../../../../../crates/litchi-ods` path; it contains no
temporary checkout path. The frozen run used both retained lockfiles and
`--locked`.

## Provenance

The historical pre-correction candidate is based at
`d20f8402295c8a73a54faa25f8a424a70c092b36`. Its source was a dirty,
uncommitted review checkout. Relevant source hashes at capture time:

```text
crates/litchi-ods/src/data_style/source.rs       c56beaa6ed5e1f8cf77bf3a42006c3e2f016b890c7b1c3fdefad2a4886124ecf
crates/litchi-ods/src/data_style/mod.rs          df8c404ed5b5dfca309f4116e10239a0ee4be3cf8c4d985e42b98922cf5fdfc3
crates/litchi-ods/src/advanced.rs                cdded5bc00a8ed6b9775843ff8128ab0bf1ab227bb4b3e52cf419892a4e4ce5e
crates/litchi-ods/src/document.rs                ed61fcd31b997a3eea725a550cfc5645055aadad6e18bf0210c9e7f728a41fdc
crates/litchi-odf-common/src/core/xml_splice.rs  4a3df5effdb0b500207a6e65fce69375e93682e282dbd2239bc27ded69e3cb38
crates/litchi-odf-common/src/core/stream_xml.rs  0a69000e09a77d8cd88bf668f4d70c556741a30881df462faaba297a6ef60c2a
crates/litchi-odf-common/src/core/binding_tracker.rs bbce8561226974fe7a6dfbb9afc76b2e2a39db9c6e21f7ac182f474700dac896
```

`frozen-source.patch` contains the tracked and untracked dirty-source delta
against `base_head`; it was applied to a fresh base tree and every listed file
hash matched the candidate. `frozen-source-delta.bundle` is a compact Git
bundle of the same dirty-source files and provenance metadata. It is a delta
bundle, not a replacement for the repository's base history. The exact
candidate lockfile is retained as `frozen-Cargo.lock`.

The harness and manifest hashes are:

```text
harness/main.rs f9bc82d2e334db288166f6abf109fcd049ee7ea5f9bca35ca81dbbc689c25d91
harness/Cargo.toml 690c1bbf7fc6f977bfacec001b93f9146f383b0afcc5d500617721651bfce01c
harness/Cargo.lock 8043bb7c29e458b858249de77737f9ba5a339a6a346028df7265d44aa3d8b62e
binary 5116182e0ea7e01041d55776b2697a063d6aa8a182af7b3b6b24731dcbe40f81
fixture-8 684d12d075d4c7e63dd97c285710855c9ac207b3b7ccf73c1aebc44a8f33575f
fixture-512 5457172b1c58a3b600b9b77cbf922b28d66d60822021b67ebc7fdd96d2e35cc9
```

`binary-source-before.json` and `binary-source-after.json` bind both raw
receipts to the same source, harness, lockfile, binary, and fixture hashes;
they explicitly record that this was a repeatability check with no source
change, not an optimization comparison.

Receipt hashes: `raw-before.jsonl`
`9097d9bc0c481cfdb9fcb63a2114d5c26416b3a9f41740ec0c10f9af8eaca2dd`,
`raw-after.jsonl`
`74f72d7eeca9cc743419bca2c9cd61ccfcae982f1c4f4b5c33f69f4d804d426b`.

Archive hashes: `frozen-source.patch`
`a14b8e8c7a33f86145b70d9b4b0c713e6739e1d1753105a07380d407979b41d0`,
`frozen-source-delta.bundle`
`bb0c1192861f30639b3be8869c4ac72d263a21ef8da651378494d0d350cd4c6b`,
`frozen-Cargo.lock`
`9886ab7f4458ddcf4286fa686a420d9351f4cfc278552484c64791dd159fc282`,
and `frozen-source-manifest.json`
`51f8f27e20c2bd3a3db5c431f7c1e5ba133173d93729e70096f22b525381aa0a`.
The retained provenance manifests are
`binary-source-before.json`
`7db6ee44671eda59698203470a365930be3abd5de6cd5d3eda9eed135760abfe` and
`binary-source-after.json`
`241689876d13418b4d0f6cb950cabb3504c6c275d9b1b70b89aa2e2959c41e70`.

Toolchain at capture: `rustc 1.95.0 (59807616e 2026-04-14)`, `cargo 1.95.0
(f2d3ce0bd 2026-03-21)`, Linux `x86_64`, kernel `7.0.0-1012-aws`.

## Findings

The source-qualified catalog returned exactly 8 and 512 entries. Metadata
patching retained the opaque child and comment bytes at both scales. The
default metadata patch was an exact package no-op at both scales. Typed graph
insertion and clean typed graph replacement completed at both scales. The
observed allocation totals and output sizes are recorded without comparison to
an older implementation.

The current smoke does not establish peak live bytes, RSS, cache behavior,
range-source behavior, cold/warm distributions, concurrency, or a speedup.
Those require a separately controlled measurement run against the corrected
source and a valid frozen baseline. The historical receipts must not be used
to approve or characterize later source corrections.

The retained unified patch includes required blank context-line spaces. It is preserved byte-for-byte; whitespace checks exclude this patch artifact rather than rewriting the replay evidence.
