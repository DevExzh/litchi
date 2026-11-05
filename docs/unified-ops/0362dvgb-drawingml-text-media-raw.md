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

The live consumers were migrated to current public owners and bounded APIs in
this packet.  The tracked source blobs changed from the packet base
`2f7db7ee7fa01893e9099217fc4243687d196f63` to the following final blobs:

| live source | before | after |
| --- | --- | --- |
| DOCX `harness/adapter.rs` | `32e9dec343b7d3c6c4a9c80b858ed8c07745658d` | `0a9a610d2359376009456cf247c16f183d2df20b` |
| PPTX profile `harness/main.rs` | `8b148ee11e7269a588ed7ea69e8e8f4925ab503e` | `5767ac31214619be4d05f60987808780d1e6dd8c` |
| PPTX evidence `harness/main.rs` | `e91cae72b5e71726612e0b8e7cc5e503f2696020` | `ef3d09e3a655ee75226116b9630cfd9e87aae7e5` |
| PPTX strict namespace `harness/main.rs` | `57b4832962d173dd2f470487d5fdd7863e2d84e2` | `62333b14331c070ba2e56cf891b63995026e8a2a` |
| XLSX profile `harness/adapter.rs` | `9c0d8025aefdadf3dd4128d4443f7fccca81e996` | `3a41d1871bdf7bf4d95bebb4250582c1ef4ce7e5` |
| XLSX contextual `harness/main.rs` | `ea01ea53027df4227bd4e3e83a56a7ec9d8c71fd` | `2d185547da8ccb1500af8a45c8f727b93bad2307` |

The DOCX/PPTX lifecycle consumers now use the public source package and
`source::svg` edit/commit API.  The XLSX profile now uses the public `Package`
owner, `Commit` accessors, required `SourceDrawing::scan` ordinal and
`ScanLimits`, and `SourceLimits` for OPC admission.  Contextual XLSX retains
its explicit `ScanLimits::default()` policy.  No historical evidence snapshot
was relabeled or modified.

The lock provenance for the six standalone checks is (the before hashes are
the packet-base files; the after hashes are the committed live files):

| lockfile | before SHA-256 | after SHA-256 |
| --- | --- | --- |
| `docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance/harness/Cargo.lock` | `bfa9b0faba56b190fac75e2266af959d983d17b0be40fa084833650d90e4cd6e` | `8f9ab51778fa4a33da1e07e17ee95fcfbdefaf5e2bf8bb1513585750c8ded203` |
| `docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle-performance/harness/Cargo.lock` | `b34d66bff058fbcebab8bd0c83dcadef341b016a400b5b5e952c33ed73a84df9` | `9cac33cbf5d670a90bc8a7dafe040442bb949f75bd7d540d23b696c43ddbfa20` |
| `docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle/harness/Cargo.lock` | `0114fbff7f809e9842bb7037dd4899f0f7330408a476ee3ef6d7021e501edc99` | `b19237d3a5a4a8b23ab013ef063ba75a1e6e84744551730e6ed73bf6a09342e6` |
| `docs/report/spec-gap-validation-evidence/svg-strict-namespace/harness/Cargo.lock` | `654cd55bba9b3b7bfedda7334be9a460116150656bd31bdc1f9f5b4d45fa5b4d` | `2ce6148f4398af27ace1ea9b3f96e26ca456695fdc593b48a86593b1c74a1ea2` |
| `docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/harness/Cargo.lock` | `12d4c700806b2b29e08ff3162e1aa9c44f49c14ae312467c0bdbc6c338caafe2` | `9823c4819c2294331a002ff9cb8bb8aa55bcd9c45ab4e07a64926f8b8f7943cf` |
| `docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/harness/Cargo.lock` | `986babc16c47d8df547a914cf3b23a77807dd583425ef1463b0cc8a4983247ae` | `2bf23f28af0559c82b1a61aecb16f8ac15af02c7a7e18db1a292cfca75331a1b` |

For the host configuration, the repository `Cargo.lock` is
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`,
`rust-toolchain.toml` is
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`, and
`.cargo/config.toml` is
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`.

The first `--locked --offline --all-targets` attempt for each harness was
retained externally.  Contextual XLSX, DOCX, PPTX profile, PPTX evidence, SVG
strict namespace, and XLSX profile respectively required the reviewed local
package dependency-list edges shown by their lock diffs; no registry package
`source` or `checksum` changed.  After each lock repair, the same command was
rerun with `--locked` and passed.  The final six terminal results were:

- DOCX SVG lifecycle: passed (`docx-lifecycle-locked-final3.log`).
- PPTX SVG lifecycle profile: passed (`pptx-lifecycle-profile-locked-final3.log`).
- PPTX SVG lifecycle evidence: passed (`pptx-lifecycle-locked-final3.log`).
- PPTX strict SVG namespace: passed (`svg-strict-locked-final3.log`).
- XLSX SVG lifecycle profile/phase profile: passed (`xlsx-lifecycle-locked-final3.log`).
- XLSX contextual SVG source: passed (`xlsx-contextual-locked-final3.log`).

All six used one owned external Cargo target at
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-text-media-raw/standalone-target`,
the capped shim, `LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`,
`CARGO_NET_OFFLINE=true`, `CARGO_TARGET_DIR` set to that path,
`--locked --offline --all-targets`, and no profile execution.  Its final size
was 395074325 bytes (377M `du`), below 20 GiB; no target
files, generated artifacts, binaries, or historical evidence are part of the
commit.

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
