# 0362 — Final API inventory execution and disposition readiness

This is an execution note for the final contract audit. It records the
bounded commands and disposition rules; it is not a native, performance, or
full-gate acceptance record. Item 22 (the redaction audit) remains the only
permitted deferral.

## Source and artifact freeze

After the last source merge, capture a clean commit and keep generated JSON,
logs and reports outside the checkout:

```sh
FREEZE=$(GIT_OPTIONAL_LOCKS=0 git rev-parse HEAD)
test -z "$(GIT_OPTIONAL_LOCKS=0 git status --porcelain)"
OUT=/home/zhuhe/.cache/litchi-api-contract/$FREEZE
mkdir -p "$OUT"/{glossary,inventory,receipts,target}
```

Each receipt records `FREEZE`, the exact command, `rustc -Vv`, `cargo -V`,
feature configuration, peak RSS, and the output SHA-256. The existing
isolated wsdoc target at
`/home/zhuhe/code/litchi-worktrees/.coord/isolated-wsdoc/target` is about
4.44 GB (decimal) and is valid for the safe-53 profile set, not merely for
ordinary docs. A Cargo target cache may be reused across source commits:
Cargo fingerprints rebuild changed inputs. Reusing captured JSON or other
evidence is different and requires a receipt matching the frozen source,
toolchain, flags and output hash. Do not create a fresh multi-gigabyte target
per crate when one coordinated target fits; run the captures sequentially.
Keep `litchi-iwa-protos` isolated from other rustdoc jobs because its prior
kill was load-related.

Use the memory-capped Cargo shim and the single-job settings for every build
or rustdoc capture:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/<slug>/env.sh
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_NET_OFFLINE=true RUSTC_BOOTSTRAP=1 LITCHI_MEM_MAX=20G
export CARGO_TARGET_DIR="$OUT/target"
```

Before each Cargo command, perform the 19,200,000,000-byte decimal target
preflight; after each command enforce the 20,000,000,000-byte decimal hard
ceiling using apparent bytes:

```sh
target_bytes() {
  if test -d "$CARGO_TARGET_DIR"; then
    du -sb "$CARGO_TARGET_DIR" | awk '{print $1}'
  else
    echo 0
  fi
}
preflight_target() {
  bytes=$(target_bytes)
  test "$bytes" -lt 19200000000 || {
    echo "target preflight failed: $bytes bytes" >&2
    exit 1
  }
}
postflight_target() {
  bytes=$(target_bytes)
  test "$bytes" -lt 20000000000 || {
    echo "target hard cap failed: $bytes bytes" >&2
    exit 1
  }
}
preflight_target
```

Every Cargo invocation is explicitly `--offline --locked`; no concurrent
Cargo jobs are allowed. Call `postflight_target` immediately after each one.

## JSON inventory

For every direct library package not covered by a glossary profile, and for
every package lacking a receipt-proven all-features capture, run sequentially:

```sh
cargo rustdoc -p <package> --lib --all-features --offline --locked \
  --target-dir "$CARGO_TARGET_DIR" -- \
  -Zunstable-options --output-format json --document-hidden-items
postflight_target
```

Copy the generated crate JSON immediately to `$OUT/inventory/` and record its
crate name, source commit, command and SHA-256. The inventory is then produced
without another build:

```sh
python3 tools/api_debt_inventory.py "$OUT"/inventory/*.json \
  > "$OUT/inventory.tsv"
```

An existing JSON is reusable only when its receipt proves the exact frozen
commit, `--all-features`, `--document-hidden-items`, locked dependency graph
and matching toolchain. A feature-specific output from
`check_api_glossary.py` is not an all-features inventory substitute merely
because the file has the right rustdoc schema.

The executable wrapper enumerates the direct library manifests itself and
requires the expected 68-package count. Run it only after the final candidate
is frozen; the command below keeps all generated output external and reuses a
single existing target cache when Cargo can fingerprint it safely:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/<slug>/env.sh
FREEZE=$(GIT_OPTIONAL_LOCKS=0 git rev-parse HEAD)
OUT=/home/zhuhe/.cache/litchi-api-contract/$FREEZE
python3 tools/final_api_inventory.py \
  --root "$PWD" \
  --output "$OUT" \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/isolated-wsdoc/target \
  --expected-package-count 68
```

The wrapper writes `packages.json`, one JSON and receipt per package, external
stdout/stderr logs, `inventory.tsv`, and an initially unassigned
`disposition.tsv`. It invokes `api_debt_inventory.py` only after all 68 JSON
captures succeed. Its fake-command fixture tests are build-free:

```sh
python3 -m unittest tools.test_final_api_inventory
```

## Disposition

Every non-draft inventory row gets a table entry with `crate`, `category`,
`path`, `decision`, owner and evidence. `builder` paths are drafts only when
the rustdoc path proves that they are authoring builders. Owner mutability,
transaction/rollback/snapshot names, `*_with_limits`, missing Commit triads,
unverified replay returns, unsafe-edit variants and exhaustive errors require
source repair and a fresh inventory row; they are not accepted exceptions.

Standalone CFB and soapberry ZIP transport rows may cite the substrate scope
of `0362rd`, but a public format re-export or foreign signature remains in the
owning format's debt. Profileless public facades such as `litchi-odf`, iWork
low-level crates, `litchi-fonts` and `litchi-sign` must be scanned and
explicitly classified rather than omitted.

The following source ancestry is already verified at this base:

| Scope | Evidence | Status |
| --- | --- | --- |
| DOCX helper/raw admission | `6108636e10`, `828d9473da` | integrated |
| DOCX raw model/graph | `324ec43b6e`, `64b0628406` | integrated |
| DOCX durable Ink Commit | `3c1d25de24` | integrated |
| DOCX/ODF glossary closure | `18e811d93b` | integrated |
| Leaf design §16.4 | `3c35a08905` | integrated; no pending implementation |
| Signed-native replay | `2b594e1757` is not an ancestor | keep pending until merged and rescanned |

The old `d67468f02a` DOCX helper receipt is historical evidence, not a
pending helper chain. Final native/ABBA qualification, performance
qualification and the workspace gate remain open; this note does not claim
any of them passed.
