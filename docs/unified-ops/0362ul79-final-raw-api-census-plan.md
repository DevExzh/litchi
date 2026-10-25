# 0362ul79 — final raw-public-API census execution plan

## Boundary and current state

This is a source-only execution packet prepared from assembled candidate
`0dc6c76634555aa610d31f43ca96b739f93b8720`. It is a plan, not a zero-debt
or acceptance claim. The candidate has:

| Input | Observed state |
| --- | --- |
| `tools/api_glossary.json` | 65 profile IDs covering 49 package names; SHA-256 `3654dd66117c11802f4fda75160d9bd9b909b5938b0e7553f24e911446f66407` |
| `tools/raw_api_coverage.json` | 15 declared profiles, 22 callable entries, 12 active lanes; coverage state `incomplete`; SHA-256 `1b4ddc26cc9f5ce729de222ebd64ba64d3a758104af0876621ed43ee19788f28` |
| `tools/final_api_inventory.py` | direct-library package guard is `68`; SHA-256 `82b3bb35567c4074c7eaf51da2977b603bbb432c46794b22eb1db486478bd98a` |
| direct `crates/*` library packages | 68; current package-set SHA-256 `f31dcf7d58ea9845f8f7585c077d4fcbc48f6fa3dd4c3dd2a837ce51bf23d703` |
| `rust-toolchain.toml` | pinned to Rust 1.95.0; SHA-256 `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |
| root `Cargo.lock` | ignored by `*.lock` and therefore absent from the Git tree; the verified integration working tree has the approved lock at SHA-256 `efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811` |

The `final_api_inventory.py` package count and the 68-package set are the
workspace coverage guard. The raw census must additionally cover every
glossary feature profile, not just one all-features graph per package.

The currently active rows are provisional. In particular, 0362ul44,
0362ul61–63, 0362ul65–70, 0362za, and `crypto-raw` must be closed after the
remaining source waves are merged. No active lane or active entry is accepted
by final `--require-complete` mode.

## Exact profile/input matrix

The following 65 glossary profiles are the exact commands implied by the
current glossary. `default` means omit both feature switches; otherwise the
listed switches are literal.

| Package | Profile(s) and Cargo feature switches |
| --- | --- |
| `litchi` | `facade-iwork` [`--no-default-features --features iwork`]; `facade-sheet-text` [`--no-default-features --features sheet-text`] |
| `litchi-biff` | `biff` [`--no-default-features`] |
| `litchi-cfb` | `cfb` [`--no-default-features --features write`] |
| `litchi-core` | `core-minimal` [`--no-default-features`] |
| `litchi-crypto` | `crypto` [`--no-default-features --features ooxml`] |
| `litchi-doc` | `doc-minimal` [`--no-default-features`]; `doc-full` [`--no-default-features --features formula,performance-diagnostics`] |
| `litchi-docx` | `docx-minimal` [`--no-default-features`]; `docx-full` [`--no-default-features --features automatic-fonts,encryption,sign,vba-inspection`] |
| `litchi-drawingml` | `drawingml` [`--no-default-features`] |
| `litchi-fonts` | `fonts-minimal` [`--no-default-features --features subset`]; `fonts-full` [`--no-default-features --features automatic`] |
| `litchi-imgconv` | `imgconv` [`--no-default-features`] |
| `litchi-iwa` | `iwa` [`default`] |
| `litchi-iwa-archive` | `iwa-archive` [`--features internal-iwork-source`] |
| `litchi-iwa-common` | `iwa-common` [`default`] |
| `litchi-iwa-core` | `iwa-core` [`default`] |
| `litchi-iwa-detect` | `iwa-detect` [`default`] |
| `litchi-iwa-protos` | `iwa-protos-wire-codecs` [`default`] |
| `litchi-iwa-structured` | `iwa-structured` [`default`] |
| `litchi-iwa-text-wire` | `iwa-text-wire` [`default`] |
| `litchi-keynote` | `keynote-minimal` [`default`]; `keynote-full` [`--features internal-iwork-source`] |
| `litchi-markdown` | `markdown` [`default`] |
| `litchi-numbers` | `numbers-minimal` [`default`]; `numbers-full` [`--features internal-iwork-source`] |
| `litchi-odb` | `odb` [`default`] |
| `litchi-odc` | `odc` [`default`] |
| `litchi-odf-common` | `odf-common` [`default`] |
| `litchi-odf-formula` | `odf-formula` [`default`] |
| `litchi-odg` | `odg` [`default`] |
| `litchi-odi` | `odi` [`default`] |
| `litchi-odm` | `odm` [`default`] |
| `litchi-odp` | `odp` [`default`] |
| `litchi-odraw` | `odraw` [`--no-default-features`] |
| `litchi-ods` | `ods` [`default`] |
| `litchi-odt` | `odt` [`default`] |
| `litchi-ograph` | `ograph` [`--no-default-features`] |
| `litchi-ole-common` | `ole-common` [`--no-default-features --features performance-diagnostics`] |
| `litchi-ooxml-common` | `ooxml-common-minimal` [`--no-default-features`]; `ooxml-common-full` [`--no-default-features --features encryption,vba-inspection`] |
| `litchi-opc` | `opc-minimal` [`--no-default-features`]; `opc-full` [`--no-default-features --features sign,performance-diagnostics`] |
| `litchi-ops` | `ops-rtf` [`--no-default-features --features rtf`]; `ops-odf` [`--features odt,ods,odp,oth,odm,odg,odc,odi,odb,odf-formula`]; `ops-iwork` [`--features pages,keynote,numbers`]; `ops-all` [`--no-default-features --features doc,docx,encryption,eval,formula,iwork-adapter,keynote,markdown,numbers,odb,odc,odf-formula,odf-objects,odg,odi,odm,odp,ods,odt,opc,oth,pages,ppt,pptx,rtf,test-support,xls,xlsb,xlsx,yaml`] |
| `litchi-oth` | `oth` [`default`] |
| `litchi-pages` | `pages` [`default`] |
| `litchi-ppt` | `ppt-minimal` [`--no-default-features`]; `ppt-full` [`--no-default-features --features automatic-fonts,encryption,formula,performance-diagnostics,sign,vba-inspection`] |
| `litchi-pptx` | `pptx-minimal` [`--no-default-features`]; `pptx-full` [`--no-default-features --features automatic-fonts,encryption,vba-inspection`] |
| `litchi-rtf` | `rtf-minimal` [`--no-default-features`] |
| `litchi-sheet-text` | `sheet-text` [`--no-default-features`] |
| `litchi-sign` | `sign` [`--no-default-features`] |
| `litchi-spreadsheet-drawing` | `spreadsheet-drawing-minimal` and `spreadsheet-drawing-full` [`--no-default-features`] |
| `litchi-xls` | `xls-minimal` [`--no-default-features`] |
| `litchi-xlsb` | `xlsb-minimal` [`--no-default-features`]; `xlsb-full` [`--no-default-features --features sign,vba-inspection`] |
| `litchi-xlsx` | `xlsx-minimal` [`--no-default-features`]; `xlsx-full` [`--no-default-features --features encryption,vba-inspection`] |
| `soapberry-zip` | `soapberry-zip` [`--no-default-features`] |

There are 19 direct library packages without a glossary profile. They cannot
be omitted: the checker compares declared profile packages with the 68-package
workspace set. Add one explicit custom profile for each, using these exact
feature closures:

| Custom profile ID | Package | Configuration |
| --- | --- | --- |
| `workspace-litchi-cli` | `litchi-cli` | `default_features=false`, no features |
| `workspace-litchi-codepage` | `litchi-codepage` | `default_features=false`, no features |
| `workspace-litchi-eval-full` | `litchi-eval` | `default_features=false`, `web_functions` |
| `formula-0362ul69` | `litchi-formula` | retain the existing no-feature custom profile |
| `workspace-litchi-iwa-cache` | `litchi-iwa-cache` | `default_features=false`, no features |
| `workspace-litchi-iwa-graph` | `litchi-iwa-graph` | `default_features=false`, no features |
| `workspace-litchi-iwa-index` | `litchi-iwa-index` | `default_features=false`, no features |
| `workspace-litchi-iwa-package` | `litchi-iwa-package` | `default_features=false`, no features |
| `workspace-litchi-iwa-text` | `litchi-iwa-text` | `default_features=false`, no features |
| `numbers-wire-0362ul66` | `litchi-numbers-wire` | retain the existing no-feature custom profile |
| `workspace-litchi-odf-full` | `litchi-odf` | `default_features=false`, `all` |
| `workspace-litchi-py` | `litchi-py` | `default_features=false`, no features |
| `workspace-litchi-sheet` | `litchi-sheet` | `default_features=false`, no features |
| `workspace-litchi-slide` | `litchi-slide` | `default_features=false`, no features |
| `vba-0362ul67` | `litchi-vba` | retain the existing no-feature custom profile |
| `workspace-litchi-word` | `litchi-word` | `default_features=false`, no features |
| `xldm-0362ul68` | `litchi-xldm` | retain the existing no-feature custom profile |
| `workspace-xml-minifier` | `xml-minifier` | `default_features=false`, no features |
| `workspace-xml-minifier-macros` | `xml-minifier-macros` | `default_features=false`, no features |

The final normalized input set is therefore 84 profiles: 65 glossary rows
plus these 19 custom rows. The provisional `fonts-0362za` row should be
replaced by `fonts-full` (the `automatic` feature is a superset of the
no-feature API), and `markdown-0362ul70` should be replaced by the exact
glossary `markdown` profile. Keeping either duplicate custom row requires a
separate exact profile input and must be recorded; it cannot silently count
as one of the 65 glossary profiles.

## Frozen capture procedure

After the last source wave merges, freeze a clean candidate `S` and record:

- `git rev-parse --verify S^{commit}` and a clean `git status --porcelain`;
- the root `Cargo.lock` SHA-256, `rust-toolchain.toml` SHA-256, glossary SHA,
  manifest SHA, and direct-package-set SHA;
- `rustc -Vv`, Cargo shim path/SHA, `LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`,
  `CARGO_INCREMENTAL=0`, `CARGO_NET_OFFLINE=true`, and a fresh external
  `TMPDIR`; and
- a fresh external output directory and target directory, both outside the
  checkout and not under `/tmp`.

The Git tree does not contain the ignored root lockfile, so the capture must
seed the exact approved ignored lock before any Cargo command. The current
integration working tree was verified with SHA-256
`efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811`;
the coordinator may copy that exact file (or an independently validated
approved source lock) into the frozen checkout and must record the resulting
hash. Do not generate or resolve a new lock during capture.

First run the existing all-features package guard sequentially:

```text
unset LITCHI_CAPPED
python3 tools/final_api_inventory.py \
  --root "$ROOT" \
  --output "$OUT/all-features" \
  --target-dir "$TARGET" \
  --cargo "$CARGO_SHIM" \
  --no-reuse
```

This command must enumerate exactly 68 direct library packages, keep every
JSON and receipt external, enforce the 19.2–20.0 GB target guard and 20 GB
RSS cap, and run `api_debt_inventory.py` only after all captures succeed. Its
all-features JSON is a cross-check and an input for the package-level debt
inventory; it is not permission to substitute one feature graph for a
different glossary profile.

For each of the 84 profile rows, build and copy one profile-specific JSON
immediately. The canonical command shape is:

```text
RUSTC_BOOTSTRAP=1 CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_NET_OFFLINE=true \
  "$CARGO_SHIM" rustdoc --package "$PACKAGE" --lib \
  --target-dir "$TARGET" [--no-default-features] [--features FEATURE_LIST] \
  --offline --locked -- \
  -Zunstable-options --output-format json --document-hidden-items
```

Before each invocation remove the previous
`$TARGET/doc/${PACKAGE//-/_}.json`; after success validate `index`, `paths`,
the library root, and `includes_private=false`, then copy the output to
`$OUT/rustdoc/$PROFILE.json` and record bytes/SHA-256, command, source commit,
lock/toolchain hashes, peak RSS, target pre/post size, and stdout/stderr
paths. Check source, lock, toolchain, target and process ownership before and
after every profile. A failed or missing output is a failed capture, never a
reused prior JSON.

Run the glossary policy against each of the 65 expected profile JSONs without
rebuilding them:

```text
python3 tools/check_api_glossary.py \
  --policy tools/api_glossary.json --profile "$PROFILE" \
  --json "$OUT/rustdoc/$PROFILE.json"
```

The 19 custom profiles have no glossary row by design, but still enter the
raw census and package/debt inventory. Their public API is not exempt from
disposition.

## Disposition and manifest closure

For every profile JSON, enumerate exactly
`census_profile(Surface(document, profile.crate))`. Add one manifest entry for
every `(profile, public_callable_path)` identity; duplicate identities,
unknown fields, generic dispositions, or unlisted callables are errors.

Classify each row from source evidence and the contract, not from a name
heuristic:

1. A variable raw function, record parser, or codec gets
   `required_raw_limits` and an exact `required_signature` (`inputs`, the
   limits argument, and the fully qualified nominal limits type). The checker
   must resolve the actual rustdoc nominal identity. Fixed-size input does not
   by itself turn a public raw/record parser into a scalar exemption.
2. A package owner whose limits arrive through `OpenOptions`/the retained
   owner policy gets `configured_owner` or `retained_owner`, with a source
   anchor and evidence for that policy.
3. Only an inherently bounded semantic value or fixed scalar conversion gets
   `semantic_value` or `fixed_scalar`. Source-sized allocation, raw bytes,
   records, XML/RTF/OVBA/CFB payloads and variable output remain raw-policy
   boundaries.
4. `private_removed` is allowed only when the exact callable is absent from
   every applicable final rustdoc graph and the removal is source-anchored.
   It is stale if rustdoc still exports it. No aliases, deprecations, or
   generic `out_of_scope` rows are accepted.

Every callable source anchor must be `rust_item`, with the final frozen blob,
line and declaration symbol. Only a temporary lane scope may be descriptive;
the final manifest has no lane table. Re-anchor all current 0362ul44/61–70,
0362za and `crypto-raw` rows after source freeze, then migrate their entries
to a final disposition. The final `entries` must contain no `active_lane`
rows and `lanes` must be `[]`.

## Final command and acceptance predicates

Run the tooling unit suite, then bind exactly the 84 profile IDs—no extra or
missing `--rustdoc` arguments:

```text
python3 -m unittest -v tools.test_raw_api_coverage
python3 tools/raw_api_coverage.py \
  --root "$ROOT" \
  --manifest tools/raw_api_coverage.json \
  --glossary tools/api_glossary.json \
  --require-complete \
  --rustdoc PROFILE_1="$OUT/rustdoc/PROFILE_1.json" \
  ... \
  --rustdoc PROFILE_84="$OUT/rustdoc/PROFILE_84.json" \
  --json
```

Accept only a report with all of the following empty/true: `missing_profiles`,
`missing_workspace_packages`, `profile_mismatches`, `missing_inputs`,
`extra_inputs`, `unclassified`, `stale_entries`, `signature_failures`,
`source_failures`, `current_source_failures`, `active_entries`,
`active_lanes`, and `complete=true`; the command must exit 0. Also require
the 68-package `api_debt_inventory` disposition file to have no unresolved
rows. If any row is unresolved, source drifted, a profile is stale, or a
profile input is reused without exact identity, keep coverage `incomplete`
and return the concrete row for repair.

The rustdoc JSONs, logs, target and lockfile remain external. Commit only the
small manifest/record and hashes needed to reproduce the run; do not commit
large generated rustdoc or command-log evidence.
