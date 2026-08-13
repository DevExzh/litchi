# Change 0091: XLS worksheet-visibility source-overlay evidence

This change adds four opt-in selectors to `tools/perf-baseline`:

- `xls_visibility_eager_edit_save`
- `xls_visibility_source_backed_edit_save`
- `xls_visibility_eager_batch_edit_save`
- `xls_visibility_source_backed_batch_edit_save`

They use one deterministic CFB/XLS corpus with 66 worksheet owners, eight
256 KiB incompressible opaque streams, and opaque metadata. The one-owner pair
changes worksheet position 1's `BoundSheet8.hsState` byte; the batch pair
changes exactly the 64-owner transaction bound by hiding positions 1 through
64, leaving positions 0 and 65 visible.

Each run reads the source through the harness-owned instrumented source and
records content-free source/sink counters, separate semantic staging/commit and
sequential publication durations, changed-owner/stream counts, output hashes,
and (for source-backed publication) changed spans and exact source/target
fingerprints. The default 36-case matrix is unchanged. These selectors are
correctness/baseline evidence only: no release ABBA, speedup, allocation, RSS,
peak-memory, or physical-I/O claim is made. Source counters are explicitly
limited to owned source ingress because the XLS public API does not accept a
caller-provided `ReadAt`; the source-backed path currently retains its complete
candidate snapshot.

Untimed gates reopen every worksheet, compare the complete CFB stream catalog
and opaque bytes, verify exact one-byte offset changes, eager patch replay and
inverse, source-backed fingerprints and span counts, no-op identity, the
64-owner cap refusal, and protected-source refusal. The publication sink is
bounded to 64 KiB writes and retains the complete output only for exact digest
and reopen assertions; this is not a candidate-memory bound.

Focused validation:

```text
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  xls_visibility_controls_are_deterministic_bounded_and_source_evidenced
cargo run --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --case xls_visibility_source_backed_edit_save --samples 1 --warmup 0
```
