# 0380 — W2-A Numbers drawable style and text-layout slice

- Status: implementation slice prepared on `uo2/iwa-w2a`; coordinator
  integration is held until the current integrated 53-row gate completes.
  Scoped facade, host differential, static, rustdoc, ledger, boundary and
  fuzz-lock checks pass on the current branch; native acceptance and final
  integrated qualification remain pending.
- Scope: the 18 Numbers host rows for shape fill, stroke, effects and frame
  text layout, plus text-box columns and frame text layout (set, reset and
  read). Numbers shadow rows remain excluded because the host does not read a
  themed Numbers rectangle shadow reliably; endpoints, segments, image fills,
  presets, lifecycle, catalog, captions, titles and comments remain separate
  owners.
- Public destination: `litchi_numbers::sheet::drawable::style`, with
  selector-first `Package::{sheet_drawable_style,
  edit_sheet_drawable_style, apply_sheet_drawable_style}` and focused
  `Edit::{set, reset}`, `Commit`, `Patch`, `Diagnostics`, `Error`, `LimitKind`
  and `Refusal`. No native object identifiers, archives or protobuf values
  cross the public boundary.

## Implementation

The facade reuses `litchi_iwa_style::shape::{value,set,reset}` and Numbers'
existing `package::style_graph::{LeafComponents,publish}`. The adapter resolves
sheet and drawable selectors through `sheet_drawable_graph`, rejects properties
that do not apply to the selected shape/text box, refuses image fills, charges
staged capacity before work, and uses the shared exact-source artifact pair.
The candidate is reopened through the strict Numbers ingress and every final
staged property is read back before a commit is returned. Reset expectations
are derived independently from the selected style's parent-chain effective
value, checked in the working graph, and checked again after reopen; exclusive
resets must remove the selected direct slots, while shared resets must retain
an inherited child link without mutating the sibling. Staging ceilings are
reported through the scoped wire `LimitKind` values. Untouched components and
members remain under the style graph publisher's exact-source locality policy;
`Patch::inverse()` and patch-source authorization are covered by the crate
tests.

## Differential proof and evidence boundary

`crates/litchi-iwa/tests/parity/numbers_shape_style_cases.rs` compares host
routes with the leaf on a source-built `basic.numbers` fixture containing a
text box and rectangle. Set/allocation cases are Tier B because fresh style
identities and the save-token policy differ; resets, rewrites and reads are
declared Tier A where the host and leaf retain the same exact graph. The
parity cases include inverse restoration and strict leaf reopen; the scoped
run passes all 18 vectors (six reads, six resets and six sets). They remain
queued for native open/save/close/reopen evidence in
`docs/unified-ops/0363-mac-queue.md`; no E3/E4 claim is made by this slice.

The shared parity harness does not include Numbers shadow rows here: a
host-readable fixture is absent, so those rows stay host-only until an
Apple-authored fixture and a separate owner decision exist.

## Performance and verification

`numbers_w2a_cost.rs` adds ignored full-facade ABBA probes for all 18 routes.
The probes measure open, selector resolution, read or edit, commit and
serialization together. The release binary was measured under callgrind on
pinned CPU 3 in the fixed A1/B1/B2/A2 order; raw callgrind captures and a
compact summary remain outside the checkout. Wall-clock qualification was
intentionally not run, so no 5%/20-microsecond performance claim is made.
The source is commit `d1b11fed01f66dfe322cca3048289a01bbd57d1f`; the probe
source SHA-256 is `ca4306d4d601dc1894b618e253cda05e38a1a3ef2720a96a3d0ea6752131d460`.
The release test binary is 39,638,264 bytes with SHA-256
`3d3bce24be93efe0af38503108b61f26c613c2f504c1b25907cac69c24ca7737`.
The 18 raw captures total 39,853,563 bytes under
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-cost-20261028/`; the compact
per-route receipt is `summary.tsv` (2,321 bytes, SHA-256
`e0f8b9ac630508291c40d4de0eee6a60acef2feee85c4f902955d33eb5441126`) and
contains each host/leaf Ir count, delta, capture size and capture SHA-256.

Scoped verification: rustfmt, strict Numbers Clippy, 8 facade tests, 18
host-vs-leaf parity vectors, Numbers API glossary, ledger, crate boundaries,
Numbers fuzz-lock compile and Numbers rustdoc all pass. The integrated 53-row
gate, native open/save/close/reopen evidence, and final wall-clock
qualification remain coordinator-owned. The working tree contains no
untracked upstream-fixture references, generated binaries or large evidence
artifacts.
