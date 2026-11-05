# Source-native slide removal planning

The source-backed presentation editor now exposes `plan_slide_removal(Position)`.
The resulting process-local plan removes one presentation relationship, its
slide Part and the slide's owned relationship member through native OPC topology
publication. It retains untouched payloads as lazy source transfers and preserves
presentation XML bytes outside the removed slide binding. Physical identifiers
remain behind the explicit `raw()` candidate view.

Admission requires at least two slides, one registered internal layout edge,
and exactly one incoming presentation owner. It shares the eager removal's
semantic validators and the source layout/master/theme checks. Signatures,
macros, protection, unknown physical members, positional presentation metadata,
custom-show references and unsupported dependencies remain typed refusals.
A non-clean ingress tolerance report is refused because skipped relationship
subtrees could conceal an incoming edge. This uses the existing report contract;
it does not change tolerant OPC parsing or introduce new reported defects.

The symbolic plan is bounded, acquires managed candidate/metadata memory before
allocation, checks physical XML limits before semantic parsing, and rechecks
source versions and execution policy during preparation and publication.
Formatted root XML is installed with a checked source-range deletion proof.
No eager package conversion or archive serialization occurs during planning.

Six integration tests cover removal/readback, exact lexical root preservation,
metadata references, hidden incoming edges, final-slide/dependency/signature/
macro refusals, cancellation, stale sources and untouched payload laziness.
The lazy test makes reads of a 256 KiB unrelated payload fail during planning,
then permits publication and verifies its decoded bytes exactly.

This is a one-way native prerequisite, not completion of primary source CRUD.
The primary immutable owner still needs a filtered slide catalog, topology-aware
local patches and a new durable wire version carrying removed payloads and exact
metadata images. Existing v2 permutation/text patches retain their strict key-set
contract. Facade source-removal capabilities have not been enabled prematurely.
Broader format CRUD, Python bindings and paired performance acceptance remain
required by the active full goal.

## Verification

The full native PPTX suite passed 1,589 tests/doctests, including all six new
integration tests. The combined DOCX/XLSX/PPTX/encryption facade passed 1,078.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
```

Scoped rustfmt, `git diff --check`, both staged evidence guards and the
changed/untracked compiled-artifact audit passed. The rewritten base is already
an ancestor of this branch, so no rebase was necessary.


Native PPTX and PPTX facade all-target Clippy passed with warnings denied, as
did rustdoc. Both glossary profiles passed (79 minimal and 70 full banned
names). Crate boundaries passed for 66 workspace packages, 268 internal
dependency declarations and 11 explicit debt items. The OfficeCLI delta check
and all 14 generator tests passed.

```sh
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The three temporary SVG corpus fixtures and their hashes are documented in
record 0141. They are verification inputs, not new committed artifacts.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created and no permission workaround was
attempted. Both required staged guards printed nothing.

All verification handles completed before cleanup. Temporary corpus copies,
scratch logs, generated build output/root lockfile and this turn's Python caches
were removed. No temporary worktree was created; existing worktrees and inherited
working changes were preserved. The full goal remains active and incomplete.
