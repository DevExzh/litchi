# Text-box output admission under caller limits

Source-backed text-box authoring now checks its conservative serialized output
bound against the retained slide-output policy before reserving managed memory.
Previously the reservation helper used the global patch ceiling; the caller's
tighter output limit was checked only by the later authoring helper. Under
memory pressure an impossible output therefore reported memory exhaustion first
and, when memory was available, temporarily reserved resources unnecessarily.

The retained ceiling is the minimum of caller-selected part, entry, total-part
and aggregate byte limits and the native shape-output ceiling. The same check
now runs for unmanaged edits. The existing managed transient envelope and
128 MiB operation ceiling still apply after output admission. This implements
ADR 0034's caller-selected limit policy and ADR 0031's admission discipline;
it adds no public API or capability claim.

The regression exercises each of the four physical byte dimensions separately,
with both managed and unmanaged source editors. It captures the selected slide,
fully reserves the remaining shared memory, and requires the precise output
limit error. Releasing pressure permits a smaller retry with the same shape ID
as an independent fresh edit. Memory use returns to zero after owners drop.
The complete focused native suite also covers source immutability, transported
replay/inverse, topology composition, XML dialects, cancellation and references.

Validation:

```sh
CARGO_TARGET_DIR=/tmp/litchi-unified-0152-target cargo test --offline -p litchi-pptx --features encryption --test source_primary_text_box
CARGO_TARGET_DIR=/tmp/litchi-unified-0152-target cargo clippy --offline -p litchi-pptx --features encryption --lib --test source_primary_text_box -- -D warnings
CARGO_TARGET_DIR=/tmp/litchi-unified-0152-target cargo test --offline -p litchi-ops --no-default-features --features pptx,encryption --test pptx_shape_add
rustfmt --check --edition 2024 --config skip_children=true crates/litchi-pptx/src/presentation/source.rs crates/litchi-pptx/tests/source_primary_text_box.rs
git diff --check
```

Native tests: 11 passed. Native library and focused-test Clippy passed with
warnings denied. All six facade shape-add tests passed. Formatting and diff
checks passed. These are scoped checks; the full native/facade suites from
record 0151 were not rerun for this admission-order change.

Both required staged-artifact guards printed nothing. The scoped `git add`
failed because `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
cannot be created on the read-only Git filesystem. No commit was created.
The turn-owned `/tmp/litchi-unified-0152-*` build output/logs and the generated
root `Cargo.lock` were removed. No fixture copies or temporary worktrees were
created; inherited changes remain intact.

The rewritten `feat/office-format-completeness` remains an ancestor of HEAD;
no rebase is required. The prior turn made implementation progress, and this
turn closes a verified admission gap. The full goal remains active; this record
makes no performance claim and does not establish full acceptance.

The next shape-geometry milestone needs a native combined text/geometry edit
primitive, lexical splicing of only the selected transform attributes, and a
durable proof that permits those attributes while preserving all other XML.
Separate source text and geometry calls cannot implement a mixed property map
because the focused source editor permits one atomic semantic operation.
Facade wiring and capability expansion must follow that native foundation.
This is L-PPTX-W3 in record 0015; record 0059 covers existing geometry reads.
