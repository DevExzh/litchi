# 0439 ODP existing-document append baseline

This bundle is a current-revision baseline. It measures the single selector
`odp_existing_append_lifecycle`; it does not contain a before/after role and
does not authorize a speedup or regression claim.

The frozen matrix is 12 reports: `R1` runs normal tiny, medium, large and then
allocator tiny, medium, large; `R2` runs allocator large, medium, tiny and then
normal large, medium, tiny. Each report uses one fresh process, one worker,
CPU 2, three warmups, and 30 retained samples. The three shapes are 64, 4,096,
and 8,192 slides. The intended total is 360 samples. `capture.py` writes
`runs/R1` and `runs/R2`; it rejects a nonempty destination and records every
command, source manifest, binary identity, resource log, workload log, oracle
log, and artifact hash.

After the matrix, run the two separate large normal profiles:

```sh
python3 profile.py --kind stat --repo-root /path/to/clean/worktree \
  --protocol protocol.json --custody-driver /path/to/custody.py
python3 profile.py --kind record --repo-root /path/to/clean/worktree \
  --protocol protocol.json --custody-driver /path/to/custody.py
```

`stat` records user cycles, instructions, branches, branch misses, and
L1-dcache-load-misses. `record` uses user cycles at 999 Hz with FP callchains;
it retains `perf.data`, `perf-script.txt`, and a no-children report. Both are
whole-process observations. `L1-dcache-load-misses:u` is retained as reported;
an unavailable or zero-valued counter must not be presented as validated cache
behavior.

The workload oracle is supplied by the provider harness under
`oracle/verify-report.py`. Its command must accept `--report`, `--mode`, and
`--shape`, print exactly `VALID` on success, and reject changes
to source identity, corpus identity, semantic append result, output bytes, or
sample/order metadata. The verifier is run after the profiled process, outside
the GNU-time and perf scope.

After the 12 reports and both profiles are retained, derive and later replay
the summary with:

```sh
python3 -B derive.py --write
python3 -B derive.py --check
```

The summary records raw allocator vectors in elapsed/sample order, normal
percentiles with deterministic 2,000-resample bootstrap intervals, GNU-time
RSS, repeat drift flags, perf counters, and retained self-report rows.

The workload report should expose the exact append lifecycle boundary and
separate normal elapsed vectors from allocator vectors. For the current
selector, owned input cloning, appended title/body construction, and sink setup
are outside the timer. The timer covers `Snapshot::from_bytes`,
`Snapshot::transaction`, one `transaction.add`, `transaction.commit`, and
sequential committed-snapshot `write_all` to `HashingDiscardSink`. Corpus and
expected-output construction, digest finalization, reopen/semantic/archive/
patch/member checks, report assembly, and destruction remain outside. The
sink is an opaque materialized-output observer, not a bounded-operation proof.
Allocator calls, requested bytes, operation-region peak above entry, live
delta, and process RSS are descriptive measurements; allocator elapsed time is
not a claim.

The append path is the next coverage target because it exercises an existing
document through `Snapshot::from_bytes`, a public transaction append, and
commit/publication with the package preservation gates. It should be measured
before proposing a change to the append implementation. The current streaming
creation evidence is a poor optimization gate: the 8,192-slide same-API
replay had only about a 1--2% normal p50 difference and identical allocator
vectors, with whole-process profiles unable to isolate a causal hotspot.

A separate future hypothesis is to fuse the two generated-XML validation
passes. `GeneratedXmlReader::fill_fragment` first calls
`xml_minifier::audit::verify_authored` and then reparses the same bytes in
`validate_fragment_shape`. A first-pass side-car certificate could record the
fragment-shape predicates while preserving the generic audit report; valid
fragments could skip the second `quick_xml` traversal, while invalid shape
cases fall back to the existing validator to preserve error ordering and
messages. This requires a negative corpus for multiple roots, outside-root
text/refs, declarations/comments/PI/doctype, malformed numeric references,
depth, and limits. It is a future differential experiment, not a claim from
this append baseline.
