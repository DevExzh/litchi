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
TMPDIR="$OUT/tmp"
case "$TMPDIR" in /tmp|/tmp/*) echo "TMPDIR must be external: $TMPDIR" >&2; exit 1;; esac
mkdir -p "$OUT"/{glossary,inventory,receipts,target} "$TMPDIR"
export TMPDIR
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

Use the approved F5 capture policy and the single-job settings for every build
or rustdoc capture. The runner, cap wrapper, underlying Cargo identity,
environment and guards are the values recorded in
`tools/raw_api_coverage.json` under `coverage.capture_policy`; a PATH shim
alone is not an accepted capture boundary. A worker `env.sh` only provisions
that already-approved wrapper and must not substitute another Cargo path. The
approved `run_safe_gate.py` driver is the outer process boundary: it owns the
child process tree, records the runner/cap-wrapper identities, and supplies
the target and free-space guards. The environment below is the payload's
expected environment, not a replacement for that outer receipt:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/<slug>/env.sh
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_NET_OFFLINE=true RUSTC_BOOTSTRAP=1 LITCHI_MEM_MAX=20G
export CARGO_TARGET_DIR="$OUT/target"
export TMPDIR="$OUT/tmp"
```

Do not replace the outer driver's resource policy with ad-hoc shell guards.
The reviewed F5 boundary requires apparent target bytes below 18,000,000,000
before a command, below 19,000,000,000 while it is live, and below
20,000,000,000 after it exits, with at least 12,000,000,000 free bytes on the
filesystem. It also requires one build job, no unrelated builders, and
process ownership/global-isolation evidence. A payload receipt without those
outer checks is diagnostic only. Every Cargo invocation remains explicitly
`--offline --locked`; the driver schedules them sequentially.

## JSON inventory

The final all-features inventory is a payload selected by a frozen, reviewed
F5 driver step. The approved `run_safe_gate.py` remains the outer boundary;
`tools/final_api_inventory.py` does not replace it. The payload verifies the
resolved coordinator Cargo cap wrapper, source/lock/toolchain identity,
target observations and per-command resource receipt, while the outer driver
must additionally provide the target 18/19/20 GB, free-12 GB, ownership and
global-isolation receipt. Do not treat an ad-hoc direct `cargo rustdoc`
invocation, even when it happens to use a memory-capped shim, as a capture
receipt. The payload constructs and records the equivalent locked, offline
rustdoc command for each direct library package; do not substitute a
hand-written command for that step.

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

The executable payload enumerates the direct library manifests itself and
requires the expected 68-package count. Select it only after the final
candidate is frozen. The following is the payload argv recorded by the
reviewed F5 driver; do not run it directly from an unowned shell:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/<slug>/env.sh
FREEZE=$(GIT_OPTIONAL_LOCKS=0 git rev-parse HEAD)
OUT=/home/zhuhe/.cache/litchi-api-contract/$FREEZE
python3 tools/final_api_inventory.py \
  --root "$PWD" \
  --output "$OUT" \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/isolated-wsdoc/target
```

The wrapper creates and preflights the external `$OUT/tmp` `TMPDIR` without
deleting existing files, and rejects output paths under `/tmp`. It verifies
the resolved Cargo shim (including `systemd-run`, the
20G `MemoryMax` expression, and an empty inherited `LITCHI_CAPPED` marker),
measures each command with `/usr/bin/time -v`, and records peak RSS beside
the JSON size/SHA-256. It rechecks the full source commit, cleanliness,
lockfile and toolchain identity before and after every package and around the
final inventory. The wrapper writes `packages.json` (including the actual 68
package names), one JSON and receipt per package, external
stdout/stderr/resource logs, `inventory.tsv`, and an
initially unassigned `disposition.tsv` whose explicit column is `decision`.
It invokes `api_debt_inventory.py` only after all 68 JSON captures succeed;
diagnostic/skip paths are marked `incomplete`, never `complete`. Its
fake-command fixture tests are build-free:

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

The following source ancestry is already verified at composed target
`8225b82dde7bac272e9e77789a78751ae07aaa2a`:

| Scope | Evidence | Status |
| --- | --- | --- |
| DOCX helper/raw admission | `6108636e10`, `828d9473da` | integrated |
| DOCX raw model/graph | `324ec43b6e`, `64b0628406` | integrated |
| DOCX durable Ink Commit | `3c1d25de24` | integrated |
| DOCX/ODF glossary closure | `18e811d93b` | integrated |
| Leaf design §16.4 | `3c35a08905` | integrated; no pending implementation |
| Signed-native replay | `2b594e1757`, `7dc296655c`, `77bc2d3ac2` and `df2d948477` are ancestors | integrated; native/ABBA qualification and the workspace gate remain open |

The old `d67468f02a` DOCX helper receipt is historical evidence, not a
pending helper chain. Final native/ABBA qualification, performance
qualification and the workspace gate remain open; this note does not claim
any of them passed.
