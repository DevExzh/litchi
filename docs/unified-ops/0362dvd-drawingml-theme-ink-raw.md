# 0362dvd — DrawingML Theme and Ink raw codec limits

- Date: 2026-10-24.
- Source commit: `16b3705341005f3072e25c24ebb24e935f8d8183`.
- Preceding live-harness follow-up: `996651b6b1`.
- Status: Theme and Ink raw ingress/output policy repair complete; the other
  DrawingML raw lanes remain separately owned.

## API shape and migration

The raw codecs use one required module-owned policy at every public entry:

| Surface | Required policy | Bounded dimensions |
|---|---|---|
| `theme::codec` | `theme::codec::Limits` | XML input/output and scheme replacement |
| `theme::family::codec` | `theme::family::codec::Limits` | XML bytes, nodes/depth, attributes, names and namespaces |
| `theme::family::part` | `theme::family::part::Limits` | complete Theme-part XML bytes |
| `ink::codec` | `ink::codec::Limits` | source/output, XML shape, contexts/traces/brushes, attributes, identifiers, references and trace text |
| `ink::actions` | `ink::actions::Limits` | source/output, XML shape, actions/groups, payloads, scalars and attributes |

Theme `read`, `read_override`, `replace_scheme`,
`scheme_replacement_range`, family codec read/write, and Theme-family
`Snapshot`/free-function raw operations now require those policies. InkML and
action read/write/profile operations likewise require their policy. The
action raw parser reports a real typed `PayloadBytes` error for oversized
definitions and transform/trace payload spans; decoded attribute quotas retain
a hard encoded-byte guard and do not charge namespace declarations to ordinary
attribute budgets. The immutable owner/edit policy is retained through
`Prepared`/`EditOptions`; no source fidelity or no-op replay path was dropped.

There are no public `_with_limit`/`_with_limits` twins or compatibility aliases.
The existing glossary gate rows cover the retired Theme-family spellings,
builder paths, immutable owners, and result methods; the current glossary JSON
was parsed successfully. This packet migrates 23 consumer/test/harness files
in the same source commit (10 DrawingML tests, 4 PPTX consumers, 8 XLSB
consumers/examples/tests, and 1 Ink action performance harness), in addition
to the 11 codec/owner implementation files.

This record does not claim closure for chart, text, diagram, model3d, SVG, or
other raw DrawingML lanes owned by the coordinator's parallel workers.

## Validation provenance

The workspace lock observed in this worktree was `Cargo.lock`, SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`,
139,591 bytes. Commands used the capped shim first in `PATH`, one build job,
incremental/debug disabled, offline locked dependencies, and the external
target `/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target`.
The target was 2,230,219,826 bytes after the final checks, below the 20 GB
cap; all temporary files were under
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp`.

The final DML all-target test ran before the source commit on the
staged-equivalent working tree. No source content changed between that run
and commit `16b3705341005f3072e25c24ebb24e935f8d8183`; the post-commit check
and XLSB test below ran against that committed source.

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target cargo test -p litchi-drawingml --all-features --offline --locked --all-targets --no-fail-fast
```

Result: 239 library tests and 196 integration tests passed (435 total, zero
failures). External log: `test-dml-all-targets-final2.log`, 39,666 bytes,
SHA-256 `f7a5ea7f6647d94c5d3f539eb0783c33447b5992248657e8c0f7e3a776276bc1`.

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target cargo check -p litchi-drawingml --all-features --all-targets --offline --locked
```

The post-commit DML check exited 0. External log:
`check-dml-final3.log`, 178 bytes, SHA-256
`2e06dd23598b6cb318964758ed2614e6ce1d6421cc9c723cb9e5c0da1036d794`.

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/target cargo test -p litchi-xlsb --all-features --offline --locked --lib theme --no-fail-fast
```

The post-commit XLSB Theme-filtered library suite passed 12 tests with 767
filtered. External log: `test-xlsb-theme-final2.log`, 1,712 bytes, SHA-256
`fb374afeeac202c3d3f7f7d371c659b475c404ab3de90dbeeb2de56aa0c4a8a6`.

Static checks also passed: `cargo fmt --all -- --check` (external empty log,
SHA-256 `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`),
`git diff --check` (same empty-log hash), and
`python3 -m json.tool tools/api_glossary.json` (external 1,060,555-byte log,
SHA-256 `c1746f007491d10cff1d8eeb095e6fb7aff5c9b18c7986f8f6aadd4bdea4e9bb`).
The staged evidence-size and binary guards were empty. No target, executable,
or evidence file is committed.

The earlier standalone `unified-chart-add --offline --locked` failure is not
relabelled as a pass here. Its exact failed command was:

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/tmp cargo check --manifest-path tools/perf-baseline/unified-chart-add/Cargo.toml --offline --locked --target-dir /home/zhuhe/code/litchi-worktrees/.coord/drawingml-limits/unified-chart-target
```

It exited 101 with external log `unified-chart-check.log` (298 bytes,
SHA-256 `408841b16b557ec63fabc2b00f3c191ca9af807b53abcfb63f326adcc7ef4d5a`).
The subsequent registry-preserving local-package lock repair, its final
standalone locked rerun, and all lock hashes remain in
`docs/unified-ops/results/0362dvb/drawingml-owner-policy.json`; no registry
version or checksum was upgraded.
