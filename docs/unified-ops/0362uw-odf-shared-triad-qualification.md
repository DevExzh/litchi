# 0362uw — ODF shared-triad qualification lane

This slice adds the missing native qualification protocol for the ODC, ODI,
OTH and ODM families changed by `0362usd`. It is a separate four-route lane;
the corrected W8 109-route probe and its original-f2 baseline packet are not
modified. No release binary, timing result, allocator capture, or acceptance
claim is included here.

## Source and routes

The intended comparison is the exact shared-triad change from
`173dd5330476713714871007316794b9a3a1bdaa`; the candidate is deliberately
unresolved until the final integrated freeze. The historical implementation
tip `15ab9324ef4f8c927acd6ecfc19d9707c3f886ee` is provenance only and is not a
candidate or an acceptance endpoint. The route/probe manifest is
`tools/native_qualification/manifests/odf-shared-triad.json`; its runner
requires a supplied full candidate SHA and fails closed on an unresolved or
stale selection.

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
identical changed outputs; ODM records separate archive digests and performs
an ordered ZIP-member audit. The audit compares decompressed member hashes,
`meta.xml` presence/content, untouched members, compression, timestamps,
extra fields, comments, permissions and archive comments. The deterministic
ODM output currently has one explicit `archive_member_order_and_manifest_listing`
decision: untouched payload bytes and stored metadata are identical by member
name. The manifest member's listing order necessarily changes its
content/CRC/compressed size; the two physical orders and the corresponding
manifest entry orders differ. All other member records remain exact.
The audit proves each manifest lists its own physical order and rejects any
other content or archive-metadata difference. For the order-only decision, the
only permitted member-level differences are the manifest member's XML text,
content hash/length, compressed length and CRC; raw names, flags, compression
options, version fields, attributes, extra fields, timestamps, comments,
permissions, encryption and archive records must remain equal. The runner
recomputes each ordered `manifest:full-path` sequence (excluding `/`) against
the corresponding archive payload order and checks all three order flags,
rather than trusting summary booleans. In the observed pair the native
order is `mimetype,content.xml,meta.xml,META-INF/manifest.xml` and the facade
order is `mimetype,meta.xml,content.xml,META-INF/manifest.xml`. The member
audit also compares raw names, decompressed bytes and hashes, CRC/compressed
size, compression, timestamps, extra fields, comments, permissions, version
and encryption flags, plus the archive comment. This is an explicit fidelity
classification, not byte identity, semantic-only acceptance, or a waiver.

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
typed owned inputs are prepared in bounded batches of at most the fixed
128-iteration probe limit; the timed closures consume those inputs without
constructing selectors, properties or strings. The
facade's `snapshot_bytes` is reported as the available retained-snapshot
observation. `--mode allocator` runs the same real commit, diagnostics and
publication closures with the paired runner's
`odf_shared_triad_allocator` system allocator observer. It reports
operation-scoped allocation/deallocation counters for every phase. The
observer is disabled and its counters snapshotted before constructing each
JSON sample, so receipt serialization is not charged to the operation.
Allocator-instrumented elapsed time remains non-claimable; it is not a timing
or acceptance result.

The generic W4b/W5 native runner is not reused for this lane: it intentionally
rejects unknown historical lanes and hard-codes their workload argv. The
dedicated `tools/native_qualification/odf_shared_triad_runner.py` binds the
source HEAD/tree, package manifest, lock hash, probe hash, exact features,
route order, phase boundary, and validated output receipt. Its unit tests
mutate each of those contracts and verify rejection. The paired
`tools/native_qualification/odf_shared_triad_paired_runner.py` additionally
binds the preparation overlay, both endpoint trees, toolchain, binary/output
hashes, pinned CPU, live quiet-host/resource observations, and the exact
calibration plus ABBA protocol. It builds only with `--offline --locked`, one
job, an external target/TMPDIR, and fixed 19.2-GB preflight/20-GB hard limits.
The protocol is one untimed validation calibration per endpoint, three
warmups and thirty samples per leg in baseline/candidate/candidate/baseline
order, for both normal timing and allocator-observer phases. Every destination
is fresh and every stderr/stdout stream is hashed; no result is accepted when
an endpoint identity, route descriptor, phase, or source overlay drifts.

## Baseline compatibility and validation

The probe uses only the stable accessors present at the 173dd baseline:
`Commit::{package,patch,diagnostics}`, `Patch::{is_noop,is_applicable_to,
inverse}`, package `apply`, and the facade's `Commit::{file,patch,
diagnostics}` plus `Patch::apply`. A source audit against 173dd found no
post-baseline accessor in the probe. The baseline endpoint has not been built
in this source-only phase; no compatibility result is claimed until the
coordinator authorizes the external baseline build. The source-only
`tools/native_qualification/prepare_odf_shared_triad_inputs.py` command creates
that endpoint from a clean 173dd worktree by copying only the reviewed probe
adapter and a separately hashed external Cargo.lock. It verifies that the
tracked production tree is unchanged and rejects any additional overlay or a
stale `15ab` candidate. The paired runner consumes that preparation receipt;
it never transplants candidate production files into the baseline.

An earlier source-only development check was run with one cargo job and the
memory shim, but without `--locked`; it is historical debug evidence only and
is not a qualification receipt. The successful external untimed output was
`/tmp/uo2-w8-readiness-odf-shared-final.json` (15,786 bytes,
SHA-256 `37637dee1eeff82e37803a7a2579d91510af6e4f8aaed58c5436e3b21b7c89df`).
It predates the fail-closed assertions and archive audit below and must not be
reused as current evidence.

The underlying Cargo command (which the runner invokes) is:

```text
source (deleted-coord-artifact)
export CARGO_TARGET_DIR=/external/native-final/targets/odf-shared-run
export CARGO_BUILD_JOBS=1
export TMPDIR=/external/native-final/tmp/odf-shared-run
cargo run --offline --locked -p litchi-ops \
  --example odf_shared_triad_qualification --features odc,odi,oth,odm \
  -- --mode validate --family all
```

The earlier single-endpoint correctness orchestration remains useful for a
source-only sanity check. Final paired orchestration is:

```text
python3 tools/native_qualification/odf_shared_triad_paired_runner.py \
  --manifest tools/native_qualification/manifests/odf-shared-triad.json \
  --preparation "$PREPARATION" \
  --baseline-root "$BASELINE_ROOT" --candidate-root "$SOURCE_ROOT" \
  --candidate-commit "$FINAL_CANDIDATE_SHA" \
  --candidate-lock-sha256 "$CARGO_LOCK_SHA256" \
  --output "$OUT" --target-root "$TARGET_ROOT" --tmp-root "$TMP_ROOT" \
  --cpu "$CPU" --phase build

python3 tools/native_qualification/odf_shared_triad_paired_runner.py \
  --manifest tools/native_qualification/manifests/odf-shared-triad.json \
  --preparation "$PREPARATION" \
  --baseline-root "$BASELINE_ROOT" --candidate-root "$SOURCE_ROOT" \
  --candidate-commit "$FINAL_CANDIDATE_SHA" \
  --candidate-lock-sha256 "$CARGO_LOCK_SHA256" \
  --output "$OUT" --target-root "$TARGET_ROOT" --tmp-root "$TMP_ROOT" \
  --cpu "$CPU" --phase abba
```

The paired runner supplies the final selected source SHA and lock SHA, rejects
the historical `15ab` tip, verifies the baseline's exact two-file overlay and
unchanged production tree, verifies the candidate's clean source tree and
package/probe/lock identities, rejects overlapping artifacts, captures
stdout/stderr and live resource samples outside the checkout, and validates
all four rows before writing a receipt. Source closure validation rejects
sparse/skip-worktree/assume-unchanged entries, tracked symlinks and missing
HEAD paths, and hashes repository plus Cargo configuration inputs. The build
handoff journal binds the role-specific command, environment, cwd, target and
binary path/hash, toolchain, exit status, output streams and live resource
samples; probe execution never trusts a substituted receipt binary. All fresh
files reject symlink ancestors and exclusive-create existing destinations.
These commands are a future execution
recipe; this source-only slice ran no Cargo build, native probe, timing or
allocator phase.

On 2026-10-23, the runner completed that locked, untimed correctness check on
source tip `c9b352265eb833423c6852f91227a28bbf1ead02` with lock SHA
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
The external output is
`(deleted-coord-artifact)/probe.json`
(27,537 bytes, SHA-256
`11cb668cb3893de0ac9ad8858ead046e4a7dbf8a13628d3febf569d9035ac6a9`), with
receipt `receipt.json` (2,194 bytes, SHA-256
`1c3d427cee98f2e7977c13a1fbd05abdd387186110ba4342781ac2fe9fe88772`) and
validation receipt `validated-receipt.json` (1,597 bytes, SHA-256
`f639f344d52b7012036d19ab4e533882e32daa6255e9b18b0495452a544bb88c`). The
external Cargo target measured 569,618,320 bytes; stderr was 6,349 bytes. This
is a correctness-only candidate-source validation, not a baseline comparison,
release build, quiet-host timing result, final native qualification, or
acceptance claim. No PROGRESS edit is part of this lane.
