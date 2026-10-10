# Exact candidate metadata for primary source removal

This increment continues the source-native removal work from records 0144–0146.
The standalone planner is not sufficient for primary CRUD: the immutable owner
must read a filtered catalog and compose removal with text/order edits, while a
transported inverse must restore parts physically absent after reopening.
Existing v2 permutation/text wire semantics must remain strict.

## Exact candidate metadata

OPC `PreparedTopology` now exposes bounded `relationships_data_for_with_limit`
and `content_types_data_with_limit` reads. `PreparedMetadataData` retains the
complete candidate and its managed reservations; source metadata retains a
managed `PartData` payload. Callers borrow exact decoded bytes rather than
receiving an unmanaged `Arc`. Acquisition fences base/donor sources and execution
contexts before and after capture. Already captured bytes remain immutable.

The handles distinguish absent relationship members from physically present
empty manifests. Changed metadata is read from the prepared overlay; appended
relationship metadata shares the publication payload through a separately
admitted index. Untouched metadata is read lazily from the source. No whole ZIP
is generated to obtain patch metadata, and unrelated payloads remain unread.
Case-equivalent owner URIs resolve through the admitted catalog spelling.

Six new regression tests cover source/changed/appended/omitted images, exact
publication equivalence, caller ceilings, managed lifetime/cancellation,
untouched-payload laziness equivalent owner names and stale-source fences. The final OPC suite passed
1,141 tests/doctests, including all 15 focused raw-metadata tests. OPC all-target
Clippy and rustdoc passed with warnings denied.


## Primary integration finding

The attempted primary-owner integration was withdrawn before publication: a
removed-key list alone does not supply a readable restored slide on an inverse
applied to a reopened source. Merely relaxing the original catalog length check
would make source positions and snapshot lookup inconsistent. The unimplemented
primary regression draft was removed rather than leaving uncompilable tests or
advertising support through facade capabilities.

The next integration must make snapshot lookup and publication consume one
prepared effective topology, with a retained declarative recipe for rebuilding
it after text/order edits. Local inverse can cancel a removal against the still
present physical Part; reopened inverse must add the missing Part and typed
relationships, then restore exact metadata images. Both cases need bounded
payload retention and complete source/target guards. A separate durable version
must carry distinct before/after topology digests and removed resource images;
existing v2 same-key-set admission must remain unchanged.

Required integration regressions remain removal/readback and source immutability,
drop rollback, order/text/removal composition, repeated removal/final-slide
refusal, local replay/inverse, durable forward replay and inverse after reopening,
full decoded member-map restoration (including `.rels` and content types),
removal of a previously edited slide, stale source and cancellation. These are
acceptance requirements, not passing evidence for this increment.


## Verification

The native PPTX suite passed all 1,589 tests/doctests after restoring the existing
primary owner. The combined DOCX/XLSX/PPTX/encryption facade passed 1,078
tests/doctests. The OfficeCLI delta check, its 14 generator tests, scoped rustfmt,
`git diff --check` and the changed/untracked artifact audit passed.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc --no-deps
python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The rewritten base is still an ancestor; no rebase was necessary. Primary source
removal, durable topology patches, broader format CRUD, Python bindings and
paired performance acceptance remain required. The full goal is active.


Crate boundaries passed for 66 workspace packages, 268 internal dependency
declarations and 11 explicit migration-debt items. Both required staged guards
printed nothing. Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem; no commit was created and no permission workaround was
attempted.

All verification handles completed before cleanup. The three temporary SVG
corpus copies (paths and hashes recorded in 0141), generated root lockfile,
build output, scratch logs/backups and this turn's Python caches were removed.
No temporary worktree was created. Inherited working changes, base checkout,
existing worktrees and tracked lockfiles were preserved.
