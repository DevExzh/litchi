# 0359p — W5 final state-allocation attribution

- Date: 2026-10-22.
- Candidate source: `cf3c9a4eb6c8b1d661ee57cef238de78ed67bb29`.
- Status: final-source allocator attribution recorded; the generic XLS
  allocation residual remains unresolved and native latency is unqualified.

The owner-state and captured-package reuse corrections are retained.  They
produce the recorded further reduction of 48 requested bytes and 32 peak
bytes against the preceding candidate.  The final paired count still has a
residual against the pre-wave baseline, so this improvement is not an
acceptance exception.

## Paired final count

The protected candidate binary is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/protected/candidate-base-package-edit-cf3c9a4eb6c8b1d661ee57cef238de78ed67bb29/ole_w5_probe`.
Its SHA-256 is
`a74a1044ce7ec48f4192cff43d43424547577b13287ac02a3cf12061b03024d9`, and
its size is 9,217,056 bytes.  The matched baseline is the protected binary
at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-final-20261021/protected/base-alloc/ole_w5_probe`,
SHA-256
`23a7088c3f9a6563adc3ae1de644a0d2b05fe91604824afd7517aa426c981fc0`.
The probe source and runner are the immutable 0359n inputs: their SHA-256
values are `ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`
and `70d07e3b07bce4237536f1239be66aed23621a19b999d79b2230abd9f3c50b93`.

The reduced paired receipt is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/counting-base-package-edit-v2`.
Its `MANIFEST.txt` is 3,676 bytes with SHA-256
`c8d84f7576856c788db0dcb786b7bf8096b93b2aec188bdf0964a60ee86e0bd4`.
Each route used `--warmup 0 --samples 1 --batch 1`, and the history routes
kept 129 packages and 128 patches alive through validation and the allocation
snapshot.  Native timing was not run.

| route | baseline calls / bytes | candidate calls / bytes | candidate difference |
| --- | ---: | ---: | ---: |
| read | 45 / 1,192 | 45 / 1,192 | unchanged |
| edit | 33 / 3,523 | 33 / 3,523 | unchanged |
| XLS cells | 1,589 / 2,361,904 | 1,586 / 2,362,117 | +213 bytes, +62 peak bytes |
| XLS cells stage | 1,588 / 2,249,264 | 1,585 / 2,249,477 | +213 bytes, +62 peak bytes |
| XLS history | 387,769 / 441,765,854 | 387,255 / 441,802,348 | +36,494 bytes; retained live/peak unchanged |
| DOC history | 1,609,900 / 1,830,103,867 | 1,571,904 / 1,770,453,588 | reduced; retained live/peak unchanged |

All paired routes produced the same output-byte checks.  The XLS cells and
stage residuals remain a required repair; the history byte increase is also
retained as an open allocation result rather than being called noise.

## Final-source heap attribution

A direct, one-sample heaptrack capture was used for source attribution only;
it is not a timing or acceptance measurement.  The external directory is
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/heaptrack-final-cf3c-xls-cells`
and is 996 KiB.  The compressed baseline and candidate captures have SHA-256
`2d35b32413b1c330d180549d938889f5a596e0e465ed69822e043e6528a161ac` and
`12c64fbd325f42f2e1f87a0150de8533161a48ea3be5def17b2bebd8efc42e60`.
The corrected candidate-minus-baseline report is 249 KiB with SHA-256
`fb919203a603780ef22d440d4e8e0b2184e0613f49c846ef2aa282495f18c402`.

The final source layout probe, removed after measurement, reported
`PackageState` at 120 bytes.  With the allocator's 16-byte Arc header, each
state allocation is therefore a 136-byte block.  The corrected heaptrack
size histogram contains nine candidate 136-byte blocks and nine candidate
176-byte blocks where the baseline contains nine 240-byte blocks.  Candidate
stacks place the repeated state/capture work in
`Package::edit_snapshot_with_options`,
`Snapshot::from_slice_with_limits_and_cfb`, and
`ArtifactBytes::from_source`.  The larger `ArtifactBytes::from_source`
allocations are setup and verification work in this whole-process capture;
they are not claimed as the 213-byte operation residual.

The current mutation path clones an editor for failure atomicity, then
`Arc::make_mut` detaches one `PackageState`; the consuming commit reuses the
unique state.  The capture reuse fallback is retained for shared candidates.
This explains the observed state allocation class without weakening clone
isolation, source freshness, recipient limits, or exact forward/inverse
replay.  A further representation change is not included here until a
matched operation-level attribution justifies it.

The native W5 qualification remains pending.  No timing result, memory
exception, or performance acceptance is claimed by this record.
