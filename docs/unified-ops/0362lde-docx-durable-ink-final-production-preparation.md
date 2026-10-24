# 0362lde — DOCX durable Ink final production preparation

Date: 2026-10-24. This reduced receipt rebinds the durable-Ink preparation to
the final production source. The earlier 0362ld c796 packet is preserved
unchanged and its results are not transferred to this source.

## Source and lock binding

The final production candidate is
472b6705475b073e7fe08f6e11ee17fb1e29f838, with baseline
d0406be01712496ca0d43354edf4d2fe80162553. The candidate is the clean,
read-only source root:

    /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source

The candidate tree is ebc1cee05f124c0c63977aff07a23288c19e7d00; the baseline
tree is 85e61e6307c7f31df29679f1fc46313f86f368ff. Both use the reviewed root
lock 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1,
139,606 bytes, SHA-256 equal to its filename, retained at:

    /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/locks/4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1.Cargo.lock

The fresh source-only preparation receipt is outside Git:

    /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/source-preparation.json
    13470 bytes, SHA-256 8f231da0c97436b134b7589c77e3f3df4714d9dc9fa92ef15e9b2c3e06cb2e20

It binds these newly materialized artifacts:

| endpoint | source manifest | source archive |
| --- | --- | --- |
| baseline d040 | source-base.json, 2,026,554 bytes, 7fd32b7ebad7f69b6282e38abdae2bd9cf7783e0bd5116f50a320ef768f93208 | source-base.tar, 196,956,160 bytes, 659e8f689c4542442545f5f25b698bfd64a5e31243234693bc1d8ba799a6195d |
| candidate 472b | source-candidate.json, 2,037,000 bytes, fac5cd3fe6a647ec957cd3d685c75bd123ac2d43d7ff77f8ab260419088efc0d | source-candidate.tar, 201,420,800 bytes, b6e954e2b57483407780518b4b6190ac0fac7e859d854e24e07635f7f53e805e |

The selected shared-source delta covers litchi-core, litchi-opc,
litchi-ooxml-common and litchi-drawingml: one baseline-only file, seven
candidate-only files and 274 changed files. Its canonical digest is
26667c5cbb133fc8cdc2bfcbc5cc11ed27d509ed32b865b45b8589f4f8375778.

## Actual probe, fixture and production rename binding

The final source contains the same canonical probe and runners as 0362ld:

    ink-qualify.py             14574 bytes  4bcbc6e5165befba18dc7d8b63f08d081371cebd236f26e531cce99cb21441a1
    ink-abba.py                14606 bytes  8452005f92732982f86c20cebd3262322ed0add404e51b385c959ee1355aa016
    ink-collect.py             14827 bytes  e744cdb6c9a40c8404dac75a4ca60eb9c1d6bfd3b92bfb75b5e2f735f147f03e
    ink-probe.rs               14766 bytes  73bedf6deb8d918087f113b2c1f5536660f9fdbff4e91d7c011f5bb89c6b8f92
    README.md                   2649 bytes   33f90aa185e84c6bfdfaff90e8df8f41617f4a9481009622ef5d96c1a9edf7f8
    test_ink_qualification.py 10410 bytes  dddf5fce77460b90b857875fd09f21ae15c89cb58d2a1e2fb7f89037ca829a0d

The generated fixture is unchanged: probe lines 204–222, 882 bytes,
SHA-256 7a2822b95ccf1682972b529c31b5a53aa146e8bef3f368b2fef8a64ed57eae27;
one /word/document.xml part, 1,000 paragraphs, the Word main-document
content type, the officeDocument relationship, a 4-KiB reader and an 8-KiB
digest sink. There are no external fixture files.

472b does contain a real production rename, not a textual endpoint relabel:
the OPC export ArtifactStreamPublication is now ArtifactStreamReport. The
receipt binds the changed candidate files individually:

    crates/litchi-opc/src/lib.rs                       4817  3d13cd2121306aff0fa5b832d9c77e1c3e91201c55dcd8d708474d606736cf1a
    crates/litchi-opc/src/source_backed/artifact_restore.rs 148701  8b4114bc0803c2e416964e8979978f9faa7615143ec82997bb32fdd5ea925376
    crates/litchi-docx/src/source_backed.rs           163401  d8960c6c4d4ba77249288b880f336617f8dbac9dfd22d1470d46f351faaf69b7
    crates/litchi-docx/src/source_backed/deferred.rs   38738  6a2bdb21fddf1aa1457fabd80b75ba9775e8e36994551b6889a1d224aca9b92d
    crates/litchi-docx/src/source_backed/svg_lifecycle.rs 131921 5c96fdc08a94f8db67fb80287cf254884f449b6c91eeb3900917a32a483e4fec
    crates/litchi-docx/src/source_backed/tail_append.rs 148195 9004633404fd7c722fc6cbdceab4fc549f7677beab4f813fe8f2d4fb268d2d00
    crates/litchi-docx/src/source_backed/tail_append_stream.rs 91738 595b9fd85b498d67eac5a5a5d832eadd6c0959425d552cd6734e983257f1d96b

The Ink probe does not import the renamed report type. It uses the unchanged
PackageWriter fixture boundary and the two explicit API adapters:
baseline apply_durable_ink(...).unwrap(), candidate
apply_durable_ink(...).unwrap().into_package(). This preserves the Commit
boundary, exact bytes/patch JSON, independent matching-package replay, stream
preflight, and six-route fidelity contract.

## Exact future commands

All build, target, binary, timing and raw receipt paths remain external under
/home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/. Use the
reviewed Cargo wrapper
/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin/cargo
(SHA-256 495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d),
one build job, offline/locked mode, incremental off, LITCHI_MEM_MAX=20G,
19.2-GB preflight and 20-GB hard target guards.

Build and correctness:

    PATH=/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin:$PATH \
    LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
    python3 /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/docs/unified-ops/results/0362l/ink-qualify.py build \
      /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final \
      --base d0406be01712496ca0d43354edf4d2fe80162553 \
      --candidate 472b6705475b073e7fe08f6e11ee17fb1e29f838 \
      --lock /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/locks/4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1.Cargo.lock

The builder retains the existing four endpoint commands: cargo generate-lockfile
--offline; cargo build --offline --locked --release
--message-format=json-render-diagnostics with features before;
before,allocations; no feature; and allocations. It validates all eight
output files, exact forward/inverse/no-op package bytes and patch JSON, stream
settings, source/lock/compiler identities, then protects both normal and
allocator binaries.

Allocation (not a timing claim):

    PATH=/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin:$PATH \
    LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
    python3 /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/docs/unified-ops/results/0362l/ink-qualify.py allocation \
      /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final --attempt allocation-472b-a1

It runs forward, inverse, noop, forward_write, inverse_write and noop_write
for 100 operations in base, candidate, candidate, base order, requiring equal
checksums/output and paired allocation/bytes/retained/peak counters.

Native ABBA, only after a quiet coordinator grant and successful correctness:

    PATH=/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin:$PATH \
    LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
    python3 /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/docs/unified-ops/results/0362l/ink-abba.py \
      /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final/build-base/ink-native \
      /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final/build-candidate/ink-native \
      /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final/native-472b-a1 \
      --build-manifest /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final/build.json --cpu 29

The runner calibrates each of six routes to at least 500 ms, warms twice,
then records five A,B,B,A blocks with a 150-ms minimum sample and the existing
host-overlap, descendant, output-growth, RSS, target-growth and unrelated-tool
guards. Collection is:

    python3 /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/docs/unified-ops/results/0362l/ink-collect.py \
      /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final \
      --allocation allocation-472b-a1 \
      --out /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final/evidence

## Source-only validation

Preparation invoked the canonical snapshot helper and did not modify either
endpoint or the root lock. Candidate source-only guard tests and py_compile
passed. The runner log is 1,821 bytes,
5d1846a2785d21bc42f0b8bc3405fc454590bdac210521148acb54fdb14e2097.

The collector's build-provenance validator was also run directly against its
synthetic source-only fixture (no Cargo/native). Its log is 1,096 bytes,
f2340aa2053a227b6ff1b8702f88b7ce7824877eeab40e29c929aab4d29a5110.

## Actual build and correctness preflight

The reviewed build command above was run once at the fresh external output
root:

    /home/zhuhe/code/litchi-worktrees/.coord/durable-ink-final-472b/run-final

The runner was PID 712349 (tool session 87275). Its baseline Cargo child was
PID 712779 and its candidate Cargo child was PID 741799. No runner quiet-state
guard refused the run. The command used the reviewed capped wrapper, one Cargo
job, incremental compilation off, offline mode and `--locked`; no allocation
capture, native timing, ABBA or native qualification was run.

The reduced build receipt is `build.json`, 192,372 bytes,
SHA-256 `f318000e4e92c79a08b8b1c1020a1a3da3b11615b18931dee47b0ab1cea6b915`.
It reports schema `litchi.docx-durable-ink.qualify.v2`, status `complete`, the
requested and resolved refs d040/472b, 121 equal registry identities, and
the following generated project lock for both endpoints:

    project-base/Cargo.lock and project-candidate/Cargo.lock
    30,062 bytes, SHA-256 70a7c70549244bc911a4931d76b5aed820079905d355c2cb2acd6d70cc7e7153

The materialized source locks were not changed: both are 139,606 bytes with
SHA-256 `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`,
matching the reviewed root lock.

All four builds and all four protected correctness preflights exited zero:

| endpoint | normal binary (bytes, SHA-256) | allocator binary (bytes, SHA-256) |
| --- | --- | --- |
| baseline d040 | `ink-native`, 3,929,032, `94ab36c7bd70a026da4bf8a560f02976b6d9a627a908f726369ee25cab7abf7d` | `ink-allocation`, 4,147,472, `506b9a4dfc60560d6c7b0bd7fae1b3880639d818eb6cd724578dd56b40fda7a4` |
| candidate 472b | `ink-native`, 4,018,152, `3b3e63077d82bf085cd70ce4d6de2d681164b618c98fc2f6ff5fb0601037a3e1` | `ink-allocation`, 4,248,760, `6be09c7c6c2c3ca14361c09eaa63b86c26e7b2780feed1c97151f3af2843cb79` |

Each of the four protected validation directories contains the same eight
expected outputs. The exact output hashes are forward.docx (4,249 bytes,
`735848b611c697d71c1c12c2527b6dc8f4a3b899c38d0f35f3b6d8f0fd64bcae`),
forward.json (127,509 bytes,
`a9e0acc66a795deb9586fa4445b687ae98ee9821ad6281d3db21103a921622e2`),
inverse.docx/noop.docx/source.docx (each 3,592 bytes,
`b24698282d77f5fcf2ee9f6956661496570903c36e062226573102b5be6a37a5`),
inverse.json (127,509 bytes,
`9e827557b699a69595ff9a339a2ed4426ef5f7477952f42ecdfc5b3a8df9cd55`),
noop.json (615 bytes,
`9a8f7ae53873720c88492cbd0ad3b478c65d0c4f1c7263891ac2092b56e6c441`) and
streaming.json (575 bytes,
`bf9ffb743229f9ecfcbc519b30c91789a4679a3081f3bc2c04adbe3f05ce80c3`).
The four route outputs matched byte-for-byte, including the normal versus
allocator variants. Logical target sizes were 422,298,713 bytes (baseline)
and 431,376,748 bytes (candidate), below the 20,000,000,000-byte guard; the
external run's logical file total was 1,654,189,250 bytes.

The complete reduced receipt, binaries and validation outputs remain outside
Git at the run root above. Prior c796/e157 receipts remain historical and
untouched. Native allocation capture, timing, ABBA and any acceptance claim
remain pending their separate coordinator grants.
