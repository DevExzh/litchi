# 0359l — W5 COW state reuse correction

- Date: 2026-10-21.
- Candidate source: `a84650cd70390641f7734dc1f3362bcb9cd1e025`.
- Status: allocation correction complete; native latency qualification remains
  pending and no performance exception is accepted.

The changed `Edit::commit` path now consumes the edit and normalizes its
`Arc<PackageState>` with `Arc::make_mut`. A uniquely owned state is reused for
the published package; a state still shared with a package or cloned editor
detaches before normalization. The published state resets its original and
current bytes to the committed content, rebases the captured package as its
base, clears `changed`, and records the validated layout. The borrowed
`staged()` path remains unchanged. This preserves clone isolation, exact
forward/inverse replay, and no-op state sharing while removing one state
allocation per changed publication.

The focused object regression covers a sibling package that must remain
unchanged, forward patch replay on the original package, and a follow-up
no-op edit on the published package. The complete litchi-ole-common tests and
all-target performance-diagnostics Clippy pass after this correction.

## Allocation receipt

The paired counting-allocator receipt is outside the checkout at
`(deleted-coord-artifact)`.
Its 14,592-byte directory has manifest SHA-256
`b143a9927bdc0030ca09c62d1eb2a5b53978e06039d16969663165fb2bd6af80`.
The runner SHA-256 is
`70d07e3b07bce4237536f1239be66aed23621a19b999d79b2230abd9f3c50b93`.
The identical probe source SHA-256 is
`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`, and
the lock SHA-256 is
`bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`.
The baseline binary SHA-256 is
`23a7088c3f9a6563adc3ae1de644a0d2b05fe91604824afd7517aa426c981fc0`; the
corrected candidate binary SHA-256 is
`9846f15d27c0942421e254379d71aa40ed2fcfb27fd24e5f6ee8bc1e1dbe43cc`.
The candidate was built with one capped Cargo job from the current source and
the matched lock; the baseline uses the immutable `baseline-0ce4-src-8f0ffa`
source snapshot. The exact build command was:

```text
source (deleted-coord-artifact)
export CARGO_BUILD_JOBS=1
export CARGO_TARGET_DIR=(deleted-coord-artifact)
cargo build --release --offline --locked --manifest-path (deleted-coord-artifact) --examples --features current-api,w5-counting-allocator
```

The runner reports allocation counters only; it does not report native
latency. Against the pre-wave baseline, read and edit remain identical. XLS
cells and the stage route remain 1,589 to 1,587 calls, with requested bytes
2,361,904 to 2,362,341 (+437, +0.0185%) and peak bytes +94. XLS history
remains 387,769 to 387,383 calls, with requested bytes 441,765,854 to
441,833,100 (+67,246, +0.0152%); retained live and peak bytes are unchanged.
Compared with the prior COW candidate, the extra 152-byte state allocation per
publication and the 128-step 19,456-byte history increase are gone.

The final untimed native pair is protected outside the checkout at
`(deleted-coord-artifact)`.
Its manifest SHA-256 is
`92e3b027ac33e58cadbb1d5087990336cabbfd94758487475b20f09405c989b6`.
The protected current binaries are `ole_w5_probe` SHA-256
`cb4b8fdde10d30fa1c2df44c14f4e9c9a235d4dccd2ad647d71a9cf7851d51bb` and
`ole_package_clone_probe` SHA-256
`ed1a6b54e496e64951882f46f55dee429c33989839ac01ba0f8182fb4656c2b7`.
The matched baseline binaries are retained there with SHA-256 values
`03e74e97525628b281e3f9d58664301f53bdd26841fe5af3c8cc3ebf1f586328` and
`bc1eaaa0e99c5607d9a2a2486fb44a1a3442a7e5637ffbf3f73ca8d040c67627`.

## Frozen native qualification runner

The final runner is outside the checkout at
`(deleted-coord-artifact)/run-matched-w5-a84650-final.py`.
It is mode `0555`, 21,359 bytes, and has SHA-256
`88ff6682ff2d4437e188aa77cdcff97cf159a2f96da7d1b8b2d56258292bb1a1`.
The previous runner remains preserved at
`(deleted-coord-artifact)/run-matched-w5-cow.py`
with SHA-256
`ca12ae007b4babebbc1c1302edc849dd5c27601fa7cb8d8ed458cdd474b1c37e`.

The exact final command, which must be run only after a coordinator quiet
grant and with a new output directory, is:

```text
python3 (deleted-coord-artifact) \
  --before (deleted-coord-artifact) \
  --after (deleted-coord-artifact) \
  --clone-before (deleted-coord-artifact) \
  --clone-after (deleted-coord-artifact) \
  --out (deleted-coord-artifact) \
  --before-source (deleted-coord-artifact) \
  --after-source (deleted-coord-artifact) \
  --lock (deleted-coord-artifact) \
  --before-id baseline-0ce4-src-8f0ffa \
  --after-id a84650cd70 --cpu 29 --blocks 5
```

The baseline and candidate probe sources are byte-identical at
`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`; the
matched lock is
`bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`.
The manifest records baseline features
`--no-default-features --features legacy-owner-api,legacy-property-api` and
candidate features `--features current-api`.

The runner requires exactly five ABBA blocks. It calibrates every endpoint to
500 ms and rejects any measured sample below 250 ms for each of the eleven
selectors (`open`, `begin_edit`, `noop_commit`, `replace_stream`, `read`,
`edit`, `doc_tracked_revision`, `xls_cells`, `xls_cells_stage`, `doc_history`,
and `xls_history`) plus both `small` and `large` clone sizes. A child with a
nonzero exit status or any stderr is a hard failure. Each probe performs its
expected-output, forward-replay, and inverse-replay checks before emitting
timing rows; the runner also requires equal output extents across paired legs.

The host guard snapshots before and after every child and every 100 ms while
it runs. It ignores only the runner's ancestry and the live child subtree. It
recognizes actual `test-adblock-*.py` scripts whether invoked directly or by
Python/shell, all known probe/runtime names, and any direct executable whose
basename contains `peculiar`. The only peculiar exemption is an argv whose
first word is exactly `/usr/local/bin/peculiar_shell`; no other path receives
that exemption. Native timing has not been run for this receipt.

## Remaining XLS allocation attribution

The residual XLS cells delta is 1,589 to 1,587 calls, +437 requested bytes
(+0.0185%) and +94 peak bytes. XLS history is +67,246 requested bytes
(+0.0152%) over 128 publications, or about 525 bytes per publication, with
fewer calls and unchanged retained-live and peak bytes. These are separate
from the removed 152-byte per-publication state allocation.

The source trace identifies the residual as the owner migration's
failure-atomic candidate path. `cell_values::commit` opens an
`ObjectPackage`, calls `package.edit()`, and mutates the editor while the
source `package` remains alive. The current object owner keeps the captured
state behind an `Arc`; each mutator first clones the editor, then
`Arc::make_mut` detaches `PackageState` and the captured package before the
candidate is rendered, reopened, recaptured, and discovered. Publication then
installs new `Arc` boxes for the parsed package and object catalog and retains
the already validated render behind `pending_render`. The pending render shares
the published byte allocation; it does not introduce a second rendered byte
buffer. The pre-wave editor cloned its inline candidate fields directly and
published through the rendered hand-off, so the allocation shape differs even
where the payloads remain shared.

The paired counters are consistent with that transient shape: cells are 1,589
to 1,587 calls and +437 requested bytes (+0.0185%), while 128-step XLS history
is 387,769 to 387,383 calls and +67,246 requested bytes (+0.0152%), or about
525 bytes per publication. Live bytes and peak bytes are unchanged, so this is
not retained history growth. The earlier suggestion to drop the temporary
`package` owner in XLS consumers is invalid: `put_stream_shared` clones the
editor to preserve failure atomicity, and the original editor and candidate
continue sharing their state and captured package until validation succeeds.
Dropping that outer owner cannot remove the candidate COW on the measured
route and is not a proposed repair.

A sound follow-up must keep the original editor untouched through every
fallible step. One narrowly scoped design is to detach a local captured-package
candidate, apply the mutation, render/reopen/recapture/discover it, and only
then swap the editor state. That can reduce allocations on rejected mutations,
but the successful path still needs one published state and package/catalog
replacement. Removing the measured successful-path overhead requires a larger
split between immutable owner metadata and mutable candidate state; it must be
implemented across every mutator and remeasured before acceptance. No such
repair is included in the protected `a84650cd70` binaries.

Native qualification still must cover the owner clone and all eleven OLE
selectors with the frozen runner, exact output/replay checks, host-overlap
checks, and accepted batched legs of at least 250 ms.

The first execution of the frozen command was deliberately rejected by its
host guard and is retained outside the checkout at
`(deleted-coord-artifact)` (62 MiB).
It reached four of five clone blocks and then stopped at
`clone-b04-1-after-after`. The guard found the real external
`scripts/test-adblock-ag-antiblock-program.py` process and its
`/home/zhuhe/code/peculiar/.../peculiar_shell` and
`/tmp/peculiar-runtime-.../peculiar_shell` children. The failing host snapshot
has SHA-256
`c320c25b630c60ca54b70013617216cf553c995d1572c380b29a0b735ae9c6a3`.
The run as a whole is not acceptance-complete, but its four completed clone
blocks are qualified observations and remain usable alongside a fresh run
with the same binaries and protocol. The interrupted fifth block is excluded
from timing evidence. A fresh quiet grant and fresh output directory are
required after those processes drain.

The four qualified clone ratios retained from that attempt are:

| completed block | small after/before | large after/before |
| ---: | ---: | ---: |
| 0 | 0.999886856 | 0.999462079 |
| 1 | 0.999502977 | 0.999901993 |
| 2 | 0.999891273 | 0.999710884 |
| 3 | 0.999989514 | 0.999034764 |

These rows each used three samples per leg after two warmups, with every
sample above 603 ms. They are retained as the first four of the eventual
same-binary nine-block clone distribution; the fresh run supplies five more.
