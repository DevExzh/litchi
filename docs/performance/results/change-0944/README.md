# 0944 retained evidence

See the [report](../../0944-xlsx-post-attribute-publication-profile.md).
`root-reduction.json` contains the four measured-owner totals, disjoint branch
counts, explicitly nested XML counts, exclusive self rows, and semantic gates.
This packet supports instruction attribution only; no optimization is adopted.

The complete external packet is `/home/zhuhe/litchi-performance-0944`.
Its verified archive is
`/home/zhuhe/litchi-performance-0944-archive/litchi-performance-0944-v1.tar.zst`
(226,729,433 bytes, SHA-256
`3f43604846b87843bbc71fcd4c8f075d57219022a1c737c472973186e85fb592`).
`archive-summary.json` records source-manifest descriptors, file counts,
verification, cleanup, and the separately retained v2 audit supplement.
Raw Callgrind dumps, source listings, source trees, and executables stay outside
Git. Every committed evidence file is smaller than 1 MiB.

## Offline replay

Extract the archive into an empty directory. It contains the original relative
layout beneath `litchi-performance-0944/`. Run:

```sh
python3 -B /path/to/litchi-performance-0944/tools/root_replay.py \
  --output /path/to/new-reduction.json
```

The output must match `root-reduction.json` byte for byte. The checker binds the
retained tool identities, so its system tools must match the frozen plan. It
does not execute the benchmark. `root-moved-replay.json` records the successful
relocation test with the original packet hidden.

The small scripts here are review copies. In the external packet, build,
capture, validation, and replay scripts live under `tools/`; the plan lives at
`docs/performance/results/change-0944/plan.json`. Retained 0845/0705/0901 helper
copies and 0900 context use the exact paths and hashes in the plan and
`retained-helper-custody.json`. Extract the archive to run this layout; do not
run the flattened review copies as though this directory were the full packet.

## Regeneration and controls

`native-build-start.json` records the exact source commit, selected source
manifest, lockfile, Cargo command, toolchain, and environment.
`native-build-terminal.json` binds the successful build to the executable's
size and SHA-256. `plan.json` freezes CPU 12, the M/D/D/M order, zero warmups,
one sample, and all collection options. The four commands are reproduced by
`capture_v2.py --index 0` through `--index 3` in a fresh, prepared external
packet; each existing child directory refuses replacement. Later children
require the first admission and preceding successful receipts.

For an intact external packet and a fresh archive destination, the retained
archive tool records and executes the regeneration command:

```sh
LITCHI_PERF_0944_ARCHIVE_GO=1 python3 \
  /home/zhuhe/litchi-performance-0944-archive/archive-0944-v1.py --archive --go
```

The tool refuses to replace existing artifacts. `--verify --go` checks the
complete extracted file set and hashes before `--cleanup --go` can remove
owned caches. Its selector recognizes only the explicit top-level
`target-native` cache; nested source names beginning with `target-` survive.
The whole-tree `source` export is redundant with the verified selected
`source-build` closure and is excluded explicitly.

Preflight controls retained here include positive semantic/raw replay and
negative malformed-dump, terminal-numbering, JSON, path, and semantic tests.
The initial adapter audit records gaps that were fixed before captures. The
original four-row audit has two labeling defects; authoritative
`independent-four-row-audit-v2.json` preserves correction provenance and
separates top-twenty self Ir from full self Ir. This supplement is outside the
archive and is retained here byte for byte. No raw data or workload changed.
