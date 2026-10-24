# 0362dvda — DrawingML Theme/Ink final return

- Date: 2026-10-24.
- Source repair commit: `af6844b2baceb7eb5c8eb9a0ce441001ce45a9ab`.
- Source tree: `c488dee3014f6db3606b1f83bd121c50482936ac`.
- Status: final nested-namespace and expanded-attribute review return; chart,
  text, diagram, model3d, SVG, and other raw lanes remain separately owned.

## Final source repair

Theme scheme replacement now refuses a modeled non-root element that locally
rebinds its own default or qualified name to the opposite DrawingML
conformance namespace. The root-only namespace rewrite therefore cannot leave
an unreadable nested modeled subtree. Opaque extension elements retain their
local namespace bindings. Replacement font scripts are capped at the model's
`MAX_SCRIPT_COUNT` (64).

Expanded attribute duplicate detection uses a borrowed ordered key set rather
than an O(k²) scan. Namespace identities are entity-decoded before comparison,
so equivalent bindings such as `urn:foreign` and `urn&#x3A;foreign` collide;
namespace decoding remains bounded by the existing attribute/XML limits.
Reserved XML/XMLNS bindings and the valid opaque-extension path have focused
regressions. No public aliases or compatibility spellings were introduced.

The follow-up source commit `b14ea9d225ef56887b1526e2035e97eb9345bc52`
(tree `8cd797710eedd6c5a675c9bc63c4bcb72c7b0edb`) skips this modeled-binding
guard only when the immediate parent is already an opaque foreign extension.
The valid nested `dk1`/strict-default subtree is spliced byte-for-byte and
still reads back with the original palette; modeled descendants outside opaque
subtrees retain the refusal.

## Validation provenance

The final source test ran before commit `af6844b2...` on a working tree whose
source bytes equal that commit; no source changed between the test and commit.
The pinned lock observed in the worktree was `Cargo.lock`, 139,591 bytes,
SHA-256 `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The capped shim was first in `PATH`, `LITCHI_MEM_MAX=20G`, one build job,
incremental disabled, dev/test debug disabled, offline locked dependencies,
external `TMPDIR`, and target
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target`.
The target measured 2,232,611,316 bytes after the run.

Focused precommit command (14 Theme codec tests passed):

```text
env PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_CAPPED= LITCHI_BLESS= LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp-final-nested cargo test -p litchi-drawingml theme::codec::tests --all-targets --offline --locked
```

Final all-target precommit command exited 0; its external log is
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/drawingml-final-nested.log`,
40,252 bytes, SHA-256
`e8e33f67cbf9b3732acc3c039bb41a690a87ceda4d65485ab8266384b9c5eccf`:

```text
env PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_CAPPED= LITCHI_BLESS= LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp-final-nested cargo test -p litchi-drawingml --all-targets --offline --locked
```

The follow-up focused command ran against the source that became
`b14ea9d225ef56887b1526e2035e97eb9345bc52` and passed 14/14 Theme codec
tests. Its external log is 6,985 bytes, SHA-256
`b05d38ec5d07f5d942661e8dc15eeb4be12025515dbfccc4f450a062a875641a`:

```text
env PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_CAPPED= LITCHI_BLESS= LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp-final-opaque cargo test -p litchi-drawingml theme::codec::tests --all-targets --offline --locked
```

`cargo fmt --all -- --check` and `git diff --check` passed. The staged large
evidence and binary guards were empty. The prior live fourth-harness check is
retained without relabelling: source `f5631b5260d07e6e59af17ff8181092376352e35`,
root lock SHA `52ed0cac...`, harness lock SHA
`f60da4a97f25ca90a2cd2e86d2c668c2760c68cd580e382af853d887f1641144`, target
205,682,117 bytes, log 858 bytes, SHA-256
`ccfda9726b0e184c4045d98547c5793e9e37e6832e6b9fb66ddd4ce91dca3980`.

The earlier standalone unified-chart-add locked failure remains a diagnostic,
not a pass. It exited 101 with the command below; log size was 298 bytes and
SHA-256 `408841b16b557ec63fabc2b00f3c191ca9af807b53abcfb63f326adcc7ef4d5a`:

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp cargo check --manifest-path tools/perf-baseline/unified-chart-add/Cargo.toml --offline --locked --target-dir /home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/unified-chart-target
```

The preserved-registry local-package lock repair and its successful rerun are
recorded in the earlier `0362dvb` receipt; no registry upgrade is claimed here.
No binary or evidence artifact at or above 1 MiB is committed.
