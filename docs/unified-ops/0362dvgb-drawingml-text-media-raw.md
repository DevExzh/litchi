# 0362dvgb — DrawingML raw adapter review repairs

- Date: 2026-10-24.
- Parent: 0362dvga, 0362dvg, 0362dv.
- Status: source repairs and focused validation complete; coordinator review pending.

The second independent review found two adapter-boundary issues:

- PPTX `model3d::codec::write` no longer deep-clones the scene wire before
  asking the shared writer to validate structure and output size. The shared
  writer already emits missing model and relationship namespace declarations,
  so the PPTX adapter now passes the borrowed wire directly. This preserves
  the previous emitted XML while ensuring retained opaque children are not
  copied before an immediate output refusal and avoids a redundant candidate
  buffer.
- PPTX `prepare_fragment` now maps source and completed-preparation byte
  refusals to the shared typed `DrawingML::Error::Limit` boundary, retaining
  `LimitKind::XmlBytes`, the observed source/prepared length, and the active
  adapter maximum. The public error remains content-free at display time.

Regression coverage includes typed source and prepared-size refusals and a
large retained opaque scene whose narrow output policy is refused before the
writer can publish output. The source-only packet was checked with rustfmt and
`git diff --check`; no generated artifacts, binaries, or third-party paths are
tracked.

The follow-up review found that the no-clone path suppressed the shared
writer's hard-coded `m3d:` and `r:` declarations whenever an alternate prefix
bound the same URI. The writer now checks the prefix and URI that it actually
emits, retains all existing declarations, and injects only missing bindings.
The same writer path is used by output-size preflight, so the injected bytes
are counted before an output refusal. PPTX regressions cover successful
alternate-prefix output and reparsing, plus a one-byte-narrower policy that
confirms the complete emitted size is refused.

Validation receipt for the corrected source tree (parent `8ce87b6143`, before
the follow-up commit), using `Cargo.lock` SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`,
`--locked --offline`, `CARGO_BUILD_JOBS=1`, the coordinator cargo shim, and a
20 GiB memory cap:

- `cargo test --lib -p litchi-pptx 'model3d::tests::codec_'`: 6 passed.
- `cargo test --lib -p litchi-drawingml`: 241 passed.
- `cargo test -p litchi-drawingml --test svg_blip`: 24 passed.
- `cargo check -p litchi-spreadsheet-drawing -p litchi-docx -p litchi-pptx -p litchi-xlsx`: passed.

## Live standalone SVG consumer closure

The four host checks above do not compile the tracked standalone evidence
workspaces.  A source-only audit at packet source `e1d9ab1e294552568a5bb6ca05059a690eb001d6`
found these six live SVG harness manifests.  The `baseline-source/`,
`matched/`, `selected-sources/`, and `results/` trees are retained historical
inputs or receipts and are not current consumers:

| consumer | manifest | live entrypoint(s) |
| --- | --- | --- |
| DOCX SVG lifecycle | `docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance/harness/Cargo.toml` | `harness/main.rs`, `harness/adapter.rs`, `harness/support.rs` |
| PPTX SVG lifecycle profile | `docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle-performance/harness/Cargo.toml` | `harness/main.rs` |
| PPTX SVG lifecycle evidence | `docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle/harness/Cargo.toml` | `harness/main.rs` |
| PPTX strict SVG namespace | `docs/report/spec-gap-validation-evidence/svg-strict-namespace/harness/Cargo.toml` | `harness/main.rs` |
| XLSX SVG lifecycle profile and phase profile | `docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/harness/Cargo.toml` | `harness/main.rs`, `harness/phase_main.rs`, `harness/adapter.rs`, `harness/support.rs` |
| XLSX contextual SVG source profile | `docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/harness/Cargo.toml` | `harness/main.rs`, `harness/support.rs` |

The only standalone source migration in this packet is the contextual XLSX
entrypoint.  Its tracked blob changed from
`ea01ea53027df4227bd4e3e83a56a7ec9d8c71fd` at the packet parent to
`ded954197222990dfe22d6af41262f25bb5fa56e` at the packet source.  The change
passes `Limits::DEFAULT` to SVG reads and writes and passes the contextual
output ceiling through `svg_blip::Limits`; the other five live harnesses do not
call the changed shared raw APIs directly.  No historical evidence snapshot
was relabeled or modified.

The lock provenance for the deferred checks is:

| lockfile | SHA-256 |
| --- | --- |
| `docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance/harness/Cargo.lock` | `bfa9b0faba56b190fac75e2266af959d983d17b0be40fa084833650d90e4cd6e` |
| `docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle-performance/harness/Cargo.lock` | `b34d66bff058fbcebab8bd0c83dcadef341b016a400b5b5e952c33ed73a84df9` |
| `docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle/harness/Cargo.lock` | `0114fbff7f809e9842bb7037dd4899f0f7330408a476ee3ef6d7021e501edc99` |
| `docs/report/spec-gap-validation-evidence/svg-strict-namespace/harness/Cargo.lock` | `654cd55bba9b3b7bfedda7334be9a460116150656bd31bdc1f9f5b4d45fa5b4d` |
| `docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/harness/Cargo.lock` | `12d4c700806b2b29e08ff3162e1aa9c44f49c14ae312467c0bdbc6c338caafe2` |
| `docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/harness/Cargo.lock` | `986babc16c47d8df547a914cf3b23a77807dd583425ef1463b0cc8a4983247ae` |

For the host configuration, the repository `Cargo.lock` is
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`,
`rust-toolchain.toml` is
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`, and
`.cargo/config.toml` is
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`.

No standalone Cargo compile has been performed on the packet source.  Once a
validation slot is granted, run the six manifests sequentially with the
capped shim, `LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`,
`CARGO_NET_OFFLINE=true`, `--locked --offline --all-targets`, and a fresh
external target directory per manifest.  This is a compile-closure check only;
it must not overwrite historical results or run a profile measurement.  Until
that slot is granted, the standalone closure remains pending and the four host
checks above are the only executed Cargo checks in this record.

The external generated target was cleaned after validation; no generated
target files are part of the commit.

The final review then found a valid collision case where a retained foreign
`m3d:` or `r:` declaration prevented the canonical spelling from being
redeclared safely. The writer now builds one bounded prefix plan before size
preflight: it reuses a safe existing nonempty model/relationship prefix, or
selects a fresh prefix that is absent from root, raster, and blip declarations.
The plan is propagated to model3d roots, raster/blip names, relationship
attributes, closing tags, and both counting and streaming writers. Retained
opaque bytes and their foreign bindings remain unchanged.

Final collision coverage includes PPTX round-trip fixtures with foreign
`m3d:`/`r:` bindings and opaque children, plus a shared-codec default-namespace
fixture that selects fresh prefixes and verifies the narrow output admission
boundary. With the same lock and capped offline setup:

- `cargo test --lib -p litchi-pptx 'model3d::tests::codec_'`: 7 passed.
- `cargo test --lib -p litchi-drawingml`: 242 passed.
- `cargo test -p litchi-drawingml --test svg_blip`: 24 passed.
- `cargo check -p litchi-spreadsheet-drawing -p litchi-docx -p litchi-pptx -p litchi-xlsx`: passed.
