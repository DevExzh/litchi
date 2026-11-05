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

The read proof was tightened after review: each of the six read vectors now
starts from a separately host-authored, non-default property value and
directly compares `host_read(source, kind, property)` with
`leaf_read(source, kind, property)`. The read-only `Case` still requires both
routes to preserve the source bytes exactly. A negative-control test alters a
leaf reader value and verifies that the direct comparison rejects it; the
generic harness's same-reader-across-output checks remain additional
coverage, not the cross-implementation proof.

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
`(deleted-coord-artifact)`; the compact
per-route receipt is `summary.tsv` (2,321 bytes, SHA-256
`e0f8b9ac630508291c40d4de0eee6a60acef2feee85c4f902955d33eb5441126`) and
contains each host/leaf Ir count, delta, capture size and capture SHA-256.

The release binary can be regenerated from the recorded source with this
command (the target and temporary directories are external):

```sh
cd /home/zhuhe/code/litchi-worktrees/uo2-iwa-w2a
env -u LITCHI_BLESS -u LITCHI_CAPPED \
  PATH=(deleted-coord-artifact):$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_DEV_DEBUG=0 \
  CARGO_PROFILE_TEST_DEBUG=0 CARGO_NET_OFFLINE=true RUST_TEST_THREADS=1 \
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  TMPDIR=(deleted-coord-artifact) \
  cargo test --offline --locked -p litchi-iwa --all-features \
    --test parity --release --no-run
```

The 18 captures and their per-route test logs can be regenerated without
changing the checkout as follows. `ROUTES_3` is the 18-entry list in
`numbers_w2a_cost.rs`; each selected test executes the same ABBA order
`A1(host), B1(leaf), B2(leaf), A2(host)`.

```sh
cd /home/zhuhe/code/litchi-worktrees/uo2-iwa-w2a
probe_dir=(deleted-coord-artifact)
binary=(deleted-coord-artifact)
for route in \
  shape_style_fill_set shape_style_fill_reset shape_style_fill_read \
  shape_style_stroke_set shape_style_stroke_reset shape_style_stroke_read \
  shape_style_effects_set shape_style_effects_reset shape_style_effects_read \
  shape_style_layout_set shape_style_layout_reset shape_style_layout_read \
  text_box_style_columns_set text_box_style_columns_reset text_box_style_columns_read \
  text_box_style_layout_set text_box_style_layout_reset text_box_style_layout_read
do
  W2A_ROUTE="$route" taskset -c 3 valgrind --tool=callgrind --quiet \
    --dump-instr=yes --collect-jumps=yes \
    --callgrind-out-file="$probe_dir/$route.callgrind" "$binary" \
    --ignored --exact numbers_w2a_cost::numbers_w2a_route_cost_abba_3 \
    --test-threads=1 >"$probe_dir/$route.stdout" \
    2>"$probe_dir/$route.stderr"
done
```

Instruction counts are attributed only from the direct `#[inline(never)]`
symbol rows for `w2a_host_route_3` and `w2a_leaf_route_3`; `PROGRAM TOTALS`
is not used for the comparison. The receipt used this method, removing the
thousands separators from the first field and taking the first exact symbol
row:

```sh
cd /home/zhuhe/code/litchi-worktrees/uo2-iwa-w2a
callgrind_annotate --auto=no --inclusive=yes --tree=none --show-percs=no \
  (deleted-coord-artifact) \
  | rg 'w2a_(host|leaf)_route_3|PROGRAM TOTALS'
```

Representative output (the first line is context only; the next two are the
separately attributed route symbols) is:

```text
836,112,805  PROGRAM TOTALS
258,070,899  ???:parity::numbers_w2a_cost::w2a_host_route_3 [.../parity-5e410c7ba68c5074]
150,250,623  ???:parity::numbers_w2a_cost::w2a_leaf_route_3 [.../parity-5e410c7ba68c5074]
```

The instruction-count bar passes independently: all 18 leaf Ir counts are
below their paired host rows in `summary.tsv`. Wall-clock qualification is
still pending, as are native open/save/close/reopen evidence and the final
integrated 53-row root gate.

Scoped verification: rustfmt, strict Numbers Clippy, 8 facade tests, 18
host-vs-leaf parity vectors plus the direct-reader negative control, Numbers
API glossary, ledger, crate boundaries, Numbers fuzz-lock compile and Numbers
rustdoc all pass. The integrated 53-row
gate, native open/save/close/reopen evidence, and final wall-clock
qualification remain coordinator-owned. The working tree contains no
untracked upstream-fixture references, generated binaries or large evidence
artifacts.
