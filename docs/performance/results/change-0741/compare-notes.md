# 0741 matched comparison runner

`compare.py` is the packet's post-capture evidence runner. It does not invoke
Cargo, build a binary, or launch a native workload. The root coordinator owns
those serial actions. The script emits the deterministic matrix with:

- six paired native repeats per case, 30 measured samples and three warmups per
  process;
- three allocator repeats per arm and case, one measured sample and no warmup;
- arm order `baseline,candidate` on even repeats and `candidate,baseline` on
  odd repeats.

There are 24 native and 12 allocator processes for the two owned PPTX cases.
The process order is recorded in `captures/manifest.json`; each receipt must
contain the matrix fields, a zero exit, a monotonic start/end interval, the
actual command, a packet-relative `report` path, and a `files` map of SHA-256
digests. Receipts must be non-overlapping. A report command must carry the
arm's recorded binary and the row's `--warmup`, `--samples`, and `--case`
values. The root scheduler can obtain the rows with:

```text
python3 docs/performance/results/change-0741/compare.py matrix
```

After both builds and all captures have been retained, run:

```text
python3 docs/performance/results/change-0741/compare.py analyze
```

The analyzer binds `build.json` to `candidate-build.json` and `source.json` to
`candidate-source.json`. It checks every recorded binary's path, size, and
SHA-256, checks source file hashes against either the captured working tree or
the recorded git revision, and checks the packet's workspace and accepted-ADR
hash bindings. Report identity, release profile, CPU 12 affinity, source
revision, instrumentation, allocator counter revision, all correctness gates,
phase cardinalities, output hash stability, and allocation fields are strict
per-process checks.

The stable projection is the complete 0739/0740 semantic projection: case,
corpus, sink, cross-copy stable fields, and the output hash identity, with the
per-sample vector checked for consistency. Baseline reports must equal
`oracle.json` exactly. Candidate reports also must equal it exactly on the first
run. This deliberately catches archive hash, archive size, sink framing,
output hash, durable patch identity, topology, and refusal changes before any
performance conclusion is considered.

The candidate's planned precompressed-entry route may use known-size fresh ZIP
framing where the legacy generated-Deflate route uses a stream descriptor.
That is a reason to inspect exact candidate diffs after qualification, not a
reason to weaken this initial projection.

If reviewed captures show that only physical compressed-entry or patch identity
fields may differ, a later run may pass `--allowlist FILE`. The file must have
`version: 1`, `arm: candidate`, `reviewed: true`, a nonempty `reviewed_by`,
and either a global `paths` list or a case-to-list `paths` map. Paths are exact
dotted projection paths; list parents may cover their indexed children. The analyzer rejects
allowlisting protected semantic, topology, or refusal paths, including corpus
shape and payload fields, logical part counts and sizes, source identity,
slide mapping, planned logical bytes, and every correctness/refusal gate. No
allowlist is part of the initial packet. Candidate physical differences are
therefore evidence to review, rather than an implicit relaxation.

For each case and lane, timing metrics (`plan_ns`, `commit_ns`,
`publication_ns`, `reopen_ns`, and `lifecycle_ns`) report process-level p50,
p95, and mean values for both arms. Paired ratio summaries compare candidate
with baseline, and a deterministic 10,000-resample bootstrap resamples whole
matched processes. Per-process p50/p95/mean changes and every allocator field
also receive an individual ±5% diagnostic flag. Allocation fields retain their
status/scope metadata and one-sample vectors; missing, added, or differently
scoped fields fail the comparison. These are matched descriptive measurements,
not a causal speedup claim.
