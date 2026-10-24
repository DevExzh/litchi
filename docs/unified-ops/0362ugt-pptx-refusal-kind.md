# 0362ugt — PPTX refusal classification by kind

The PPTX arm of `litchi-ops::error::Classify` had two historical exceptions
for `slide_removal`: a signed-package refusal was recognized by its operation
and exact reason to attach the signed-edit suggestion, and `FinalSlide` was
reclassified from its `InvalidValue` kind to `invalid_argument`. That made the
facade's code depend on operation/reason text instead of the shared
`RefusalKind` contract in record 0351 §2.4.

The arm now maps every `Error::Refused(Refusal)` through `refusal.kind()`.
`FinalSlide` therefore reports `invalid_value`; `SignedPackage` and
`ProtectedPresentation` retain `Protected` and report `document_protected`.
The operation and reason remain the native, static, content-free diagnostics;
they do not affect the unified code. No native PPTX refusal kind or reason was
changed.

The same change updates every active consumer of the old exception in this
commit: the error-classification regression, the final-slide eager and source
integration assertions, the capability-gap assertion and PPTX capability
golden. The active W8 error-contract wording and the native `FinalSlide`
documentation now name `invalid_value`. The historical 0356/0361 exception
sentences are superseded by this correction record; they are retained as
historical wave reports. The capability disposition remains the same gap and
follow-up; only its code changes.

The focused classifier regression iterates every current `RefusalKind` with
both the old `slide_removal`/signed reason pair and an unrelated operation and
caller-authored reason marker. It asserts the kind table and that the facade
display remains content-free.

## Validation

Validation was run from the fresh `uo2/pptx-refusal-kind` leaf using
`.coord/pptx-refusal-kind/env.sh`. The environment used the repository's
memory-capped cargo shim with `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`,
`CARGO_NET_OFFLINE=true`, and external `CARGO_TARGET_DIR` and `TMPDIR`.
The accepted source commit for every command below is
`18346e7f881fb8aa1fa1bb854a8ee0852542ecf2`. The ignored offline lock overlay
was 139592 bytes with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.

All raw logs and receipts are in
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-refusal-kind/validation-18346e7f88/`.
The raw-log evidence is:

| Check | Command result | Raw log (bytes; SHA-256) |
| --- | --- | --- |
| Kind classifier | 1 passed, 0 failed, 443 filtered | `ops-pptx-refusal-kind-18346e7f88.log` (557; `19cf47ec2c9821b165331be6b02d5cb4c5a5ec37c325b45c3636aaf91330f6c4`) |
| Slide-removal consumers | 10 passed, 0 failed | `ops-pptx-slide-remove-18346e7f88.log` (1225; `80d1091c6f8a34bf1e5a52706964d8dd6e6440b0c8f5c1e1ba6acb59a2445945`) |
| PPTX capability disposition | 1 passed, 0 failed, 2 filtered | `ops-pptx-capabilities-18346e7f88.log` (481; `ac0e7a1aec1d3e5445d9b626db1d6d6d785f8ad9c0001528e482e4ed19a31a77`) |
| Strict PPTX-only ops Clippy | exit 0, `-D warnings` | `ops-pptx-clippy-18346e7f88.log` (4592; `54628eed3bb4c2d7e7ac3c4e63e4aabf3c5931548cd0708b593ba1fe6957a8fe`) |

The commands were, respectively, the exact classifier test, the complete
`pptx_slide_remove` integration target, the exact capability disposition test,
and `cargo clippy --offline --locked -p litchi-ops --no-default-features
--features pptx --all-targets -- -D warnings`. Their receipts record the
source commit, lock SHA, UTC interval, exit status, log size and target size.
The target grew from 652071845 to 766803350 bytes and remained below the 20 GB
cap. No native PPTX suite or workspace full-gate was run for this focused
correction.

## Follow-up: registry and corpus consumers

Review found one remaining generated consumer: the PPTX `final-slide` row in
`tests/golden/registry/pptx.json` still said `invalid_argument`. The row now
matches the shared `InvalidValue` classification. The OfficeCLI capability
delta was regenerated only for that registry digest change (one line); no
other generated row changed. The W8 wording now distinguishes the typed
`SoleParagraph` arm, which remains `invalid_argument`, from the PPTX
`FinalSlide` shared-kind code, which is `invalid_value`. These corrections and
the generated table are in `d9baa98c3188a7367b7012185ae98c01305f5e31`.

The final checks ran at that commit with the same ignored 139592-byte
`Cargo.lock` overlay and SHA-256 recorded above. Their raw logs and receipts
are in
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-refusal-kind/validation-d9baa98c31/`:

| Check | Result | Raw log (bytes; SHA-256) |
| --- | --- | --- |
| PPTX registry golden | 1 passed, 0 failed, 2 filtered | `ops-pptx-registry-d9baa98c31.log` (361; `a3e39565468b47cc8339eafde69f8d395d0636de91b67a33d941f04fe6de200a`) |
| OfficeCLI capability delta (`--check`) | exit 0 | `officecli-capability-delta-check-d9baa98c31.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`) |
| OfficeCLI corpus (`--check`) | exit 0 | `officecli-corpus-check-d9baa98c31.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`) |

The registry command was the PPTX-only `registry_goldens_follow_enabled_formats`
test with `cargo test --offline --locked -p litchi-ops --no-default-features
--features pptx --test registry ...`. The static checks used the tracked
`test-data/officecli`, registry/capability goldens, corpus outcomes and
`tools/officecli_dispositions.json`; the capability command explicitly passed
`--officecli test-data/officecli`. The final target directory was
771402833 bytes. The first capability check at the pre-regeneration commit
correctly reported the table stale; it was not treated as a pass, and the
follow-up check above passed after the single-line regeneration.


## Coordinator provenance verification

The coordinator verified all seven captured logs against their adjacent
source/command/exit receipts before merge. Root Cargo, toolchain, Clippy and
Cargo configuration match the committed inputs. The absolute environment
script is `/home/zhuhe/code/litchi-worktrees/.coord/pptx-refusal-kind/env.sh`
(500 bytes, SHA-256
`acbca6ffa6ff89d4f5322bccced729a48e08a5cf50174fb1587202c533310ae5`).
The target-size cap is 20,000,000,000 bytes, separate from the `20G`
process-scope memory cap. Merge also narrows an obsolete DOC classifier
comment; no unrelated DOC behavior changes.

The exact successful test and static commands were:

```text
cargo test --offline --locked -p litchi-ops --no-default-features --features pptx --test capabilities pptx_reports_selection_and_structural_writes -- --exact --nocapture
cargo clippy --offline --locked -p litchi-ops --no-default-features --features pptx --all-targets -- -D warnings
cargo test --offline --locked -p litchi-ops --no-default-features --features pptx --lib error::tests::pptx_refusals_classify_by_kind_not_operation_or_reason -- --exact --nocapture
cargo test --offline --locked -p litchi-ops --no-default-features --features pptx --test pptx_slide_remove -- --nocapture
python3 tools/officecli_capability_delta.py --officecli test-data/officecli --check
python3 tools/officecli_corpus.py --check
cargo test --offline --locked -p litchi-ops --no-default-features --features pptx --test registry registry_goldens_follow_enabled_formats -- --exact --nocapture
```
