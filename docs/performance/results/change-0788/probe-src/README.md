# 0788 RSS phase probe

This probe is an archive of `tools/perf-execution`, with the same corpus
generator, byte verification, resource limits, lifecycle, dependencies, and
release profile. Its binary is named `cached-part-memory` so it can be run
beside the frozen native execution binary.

The ordinary path leaves `LITCHI_RSS_PHASES` unset. It does not create stdin or
stdout phase handles and emits no diagnostic records. Setting
`LITCHI_RSS_PHASES=1` enables the parts-route phase handshake; other routes fail
closed. Every record is flushed as:

```text
RSS0788\tphase\tsample\tpid\n
```

The parent must reply with exactly `+\n`. EOF, a short acknowledgement, or any
other byte fails the run. Control phases use `usize::MAX` as the sample field:
`startup`, `corpus_ready`, `warmup_done`, `samples_done`, `report_written`, and
`report_dropped`. For measured parts samples, only sample zero and the final
sample are marked: `package_ready`, `after_preload`, `after_operation`,
`after_batch_drop`, and `after_package_drop`. On a fresh run `after_preload` is the
before-operation point; on a primed run it follows the verified preload. The
run's `--state` makes that distinction explicit.

Markers are outside the timed operation and are not included in native timing,
RSS, or source-call pooling. The child does not read `/proc` and does not count
allocations. Enabling the handshake deliberately perturbs the diagnostic run;
the native execution binary remains the feature-off frozen tool.

`package_ready` precedes URI allocation. `after_operation` follows the resource
snapshot and precedes byte verification. `after_package_drop` follows both
package and context drops; the local source Arc, budget root, and URI vector
remain alive until that sample returns. `report_dropped` follows corpus and
sample-report destruction, while handshake buffers remain alive.
