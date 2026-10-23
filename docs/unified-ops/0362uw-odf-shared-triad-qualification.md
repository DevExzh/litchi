# 0362uw — ODF shared-triad qualification lane

This slice adds the missing native qualification protocol for the ODC, ODI,
OTH and ODM families changed by `0362usd`. It is a separate four-route lane;
the corrected W8 109-route probe and its original-f2 baseline packet are not
modified. No release binary, timing result, allocator capture, or acceptance
claim is included here.

## Source and routes

The intended comparison is the exact shared-triad change from
`173dd5330476713714871007316794b9a3a1bdaa` to its implementing source tip
`15ab9324ef4f8c927acd6ecfc19d9707c3f886ee`. The route/probe manifest is
`tools/native_qualification/manifests/odf-shared-triad.json`.

All four routes use deterministic in-repository builders. ODC creates a chart,
ODI creates an image, and OTH creates one paragraph. ODM uses its existing
builder plus the shared timestamp-free `Structure::default_meta_xml` writer
path, because `Package::new()` intentionally has no `meta.xml` and therefore
cannot exercise `set_title`. No external corpus or untracked upstream checkout
is used.

| route | native leaf | unified facade | semantic marker |
|---|---|---|---|
| `odc_chart_title` | `replace_chart` title | `set("/", title)` | chart title |
| `odi_frame_geometry` | `set_geometry` on frame 0 | `set("/frame[1]", geometry)` | four frame lengths |
| `oth_paragraph_text` | `set_paragraph_text(Position(0))` | `set("/body/p[1]", text)` | paragraph text |
| `odm_title` | `set_title` | `set("/metadata", title)` | document title |

Every route creates a changed and an exact no-op commit on both sides. The
untimed proof reads every triad accessor, checks the source owner is unchanged,
checks patch base admission, forward replay, inverse restoration, exact no-op
replay, and repeated publication bytes. The leaf and facade outputs are also
reopened for the semantic marker. ODC, ODI and OTH currently produce byte-
identical changed outputs; ODM deliberately records separate archive digests
and requires semantic title equality, while retaining the stronger per-side
source/replay/publication proofs. This distinction is evidence, not an
acceptance waiver.

## Probe and timing contract

`crates/litchi-ops/examples/odf_shared_triad_qualification.rs` defaults to
`--mode validate` and emits
`litchi.odf.shared-triad-qualification.v1`. A family can be selected with
`--family odc|odi|oth|odm|all`. `--mode timed --iterations N` is bounded to
`1..=128` and emits three separate phases for each side:

1. `commit_only`: construct and commit the changed edit; the commit is
   black-boxed without reading diagnostics;
2. `commit_plus_diagnostics`: commit, then read `diagnostics().changed()`;
3. `publication`: commit, then serialize the resulting owner/file.

Fixture/builder setup, opening, property construction and target resolution
are outside those closures. This makes the ODM moved comparison explicit:
`commit_only` cannot hide a lazy diagnostics scan, while
`commit_plus_diagnostics` measures the accessor boundary separately. The
facade's `snapshot_bytes` is reported as the available retained-snapshot
observation; the leaf core diagnostics intentionally expose no allocation
counter, so process allocator capture remains an external runner concern.

The generic W4b/W5 native runner is not reused for this lane: it intentionally
rejects unknown historical lanes and hard-codes their workload argv. The
manifest freezes this lane's source refs, features, route operations and phase
boundaries; a later safe runner may consume it without changing the probe.

## Baseline compatibility and validation

The probe uses only the stable accessors present at the 173dd baseline:
`Commit::{package,patch,diagnostics}`, `Patch::{is_noop,is_applicable_to,
inverse}`, package `apply`, and the facade's `Commit::{file,patch,
diagnostics}` plus `Patch::apply`. A source audit against 173dd found no
post-baseline accessor in the probe. The baseline endpoint has not been built
in this source-only phase; no compatibility result is claimed until the
coordinator authorizes the external baseline build.

Current-candidate checks completed with one cargo job and the existing memory
shim (target and temporary directory outside the checkout):

```text
rustfmt --edition 2024 --check crates/litchi-ops/examples/odf_shared_triad_qualification.rs
cargo check --offline -p litchi-ops \
  --example odf_shared_triad_qualification --features odc,odi,oth,odm
cargo run --offline -p litchi-ops \
  --example odf_shared_triad_qualification --features odc,odi,oth,odm \
  -- --mode validate
```

The validation emitted all four rows; each side's changed/no-op replay and
publication checks passed, and each semantic marker matched. The run was
untimed and is not a native acceptance result. No Cargo release build, quiet
host timing, baseline build, or PROGRESS edit was performed.
