# Route profile postprocessing

`profile_routes.py` starts one fresh child for every tool, route, and case. The
profile child includes `/usr/bin/time`, `taskset`, route setup, the timed
samples, oracle/report work, and teardown. The profiler's own work is also in
the external observation. `summarize_route_profiles.py` preserves that scope;
it is a diagnostic reader and never starts a profiler.

Run the reader after a terminal profile attempt with:

```text
python3 docs/performance/results/change-0484/summarize_route_profiles.py \
  --profiles-dir docs/performance/results/change-0484/route-profiles/profiles1
```

The default output is `route-profiles/profiles1-summary.json`. The default
`route-profiles/profiles1-perf-script/` directory is reserved for coordinator
exports from `perf.data`. `--output` and `--perf-script-root` make both paths
explicit when an attempt is copied or archived. `--allow-incomplete` is for a
profile tree that is still receiving receipts; a completed gate should run
without it and must produce zero validation errors.

The identity chain is checked before an artifact is interpreted. The result and
started bindings must describe the expected tool × route × case inventory. Each
terminal receipt must bind to the normal `build-normal.json`, its copied binary,
the frozen `route-protocol.json`, and the exact `profile_routes.py` helper
bytes/hash. Receipt artifact paths, lengths, and SHA-256 values are checked on
disk. The route report is checked against both the selected protocol case and
route axis, including provider, replay ceiling/sync, open-count policy,
compression, input mode/storage, source contract, and sink size. A file-store
success must retain the driver's successful empty-scratch cleanup proof. A
missing or failed receipt is an error in a completed summary; an explicitly
unavailable profiler remains marked unavailable and contributes no fabricated
counter.

The `profiles1` inventory uses the owned source-input case for all three
authored routes. `file_store` describes the authored replay store and does not
measure the separate file-backed source-input axis; this postprocessor makes
no source-file behavior inference from those profiles.

`perf-stat.txt` is parsed as the raw `perf -x,` rows. Every requested event is
retained, including a numeric zero, while `<not supported>` remains
`available: false` and `value: null`. The summary retains the raw row, perf's
runtime field, and running percentage. A running percentage below 100 is
reported as multiplexed or partially scheduled; values are not silently
rescaled or compared as timing claims. The scope field is always
`whole child process; profiler overhead included`.

Cache events need the same restraint. A numeric zero is retained as the raw
counter, while the summary marks the cache interpretation unsupported for that
run and never forms a miss ratio from a zero denominator. An explicitly
unsupported LLC event stays null. Cycles, instructions, and page faults remain
available observations with their perf running-coverage fields.

`strace.log` is classified from the `-yy` path annotations emitted by the
frozen driver. For file-store profiles, `replay_path` is an exact match for the
receipt's replay directory. `setup_report_io` contains known report,
resource, stdout, and stderr paths; other known paths and pathless calls stay
in separate scopes. For `read`, `pread64`, `readv`, `write`, `pwrite64`, and
`writev`, counts, successful/error counts, returned bytes, and fixed returned
size buckets are reported. `fsync`, `fdatasync`, `unlink`, and `unlinkat` are
reported as calls and errors; their return value is not treated as payload
bytes. A deterministic or memory-store route has no explicit replay path, so
that scope is `not_applicable` rather than zero.

The existing `heaptrack_print` output is read as a whole-child artifact. The
summary presents the first ten peak-consumer blocks for inspection, but records
`top_only: false`, the complete print hash, and the tail totals for allocation
calls, temporary allocations, peak heap/RSS, and leaked bytes. Each displayed
row retains several stack symbols and the first application symbol when one is
present. The displayed rows are therefore not a claim that the rest of the
allocation report is absent, and no allocation comparison is inferred from
them.

For each `perf-record` receipt the summary emits this exact coordinator-run
command, with the profile's `perf.data` path and source binding fixed in the
manifest:

```text
perf script --header -i /absolute/path/to/perf.data --demangle \
  > /absolute/path/to/route-profiles/profiles1-perf-script/<profile-label>/perf-script.txt
```

The command manifest is written beside the JSON summary. The coordinator runs
the command outside this reader and reruns the reader. The rerun hashes the
text export and writes a sidecar `binding.json` containing the profile label,
`perf.data` hash, normal binary hash, build receipt hash, and protocol hash.
An existing sidecar with any mismatch is an error. The text itself is parsed
only after that source chain is present; if no export exists, the status stays
`pending` and the command remains available.

Stack sample counts are whole-child counts. A sample belongs to a lifecycle,
auditor, or replay group when any actual frame in that sample matches the
small configured symbol set (`run_case`/`run_iteration`, source/candidate/store
verification and scan functions, or replay/read-with-accounting functions).
The group counts are inclusive ancestry counts, so a sample can belong to more
than one group. Pairwise and all-three overlaps are retained and group shares
are explicitly non-additive. If an export has no parseable samples or no
configured symbol appears, the marker is `unavailable`; zero is never used as
a substitute for missing symbols. A command-line name/path check supplements,
but does not replace, the binary hash in the sidecar.

Rust v0 mangled frames can still contain the readable function suffixes used by
the marker matcher. The heaptrack top-consumer report covers the whole child,
so source scanning and semantic-oracle frames may be outside the timed route
operation; they remain diagnostic evidence with their actual stack ancestry. A `run_case` marker is not an exact
timed-region boundary, and absent marker samples do not prove absence of work.

The output is evidence for selecting the next measured hotspot. It does not
authorize a source-audit removal, timing comparison, speedup claim, or causal
allocation explanation. In particular, the standalone source XML audit remains
an independent validity proof; a candidate audit over `source prefix + replay +
suffix` cannot replace it for malformed-source and error-order cases.
