# Change 0465: checked default ODP append baseline

This batch adds the existing owned ODP append lifecycle to the default matrix.
It changes benchmark selection and checked corpus/coverage metadata. Production
code, generated fixture bytes, timing boundaries and preservation oracles are
unchanged. The new family is explicitly synthetic and Litchi-generated.

The frozen [protocol](protocol.json) runs 201 rows across 37 default cases and
31 corpora in each of two normal repeats. Separate allocator repeats cover
only the three ODP shapes. Each formal row retains 15 samples after three
warmups. A one-sample, zero-warmup preflight establishes the checked identity;
it is excluded from the descriptive timing summary. The normal and allocator
lanes are separate evidence streams; no speedup comparison is claimed.

[Binding](binding.json) and build receipts identify the actual binaries and
complete Rust/TOML/lock source custody. The source revision is the recorded base
commit plus the two changed Rust files archived by [source-code.json](source-code.json).
The complete custody manifest authenticates unchanged source from that base;
the selected source archive is not a copy of the whole repository. Reports
retain truthful dirty-worktree metadata. The generic clean-worktree,
distinct-revision comparison policy is unchanged and these descriptive runs
are not approved before/after comparison inputs. The expanded policy rejects
historical 198-row reports; a future comparison needs matched 201-row captures
that satisfy the existing clean-worktree and distinct-revision gates.

The checked ODP corpus contains 64/4,096/8,192 titled slides and a deterministic
64 KiB opaque package member. Each operation opens owned input, creates a
transaction, appends one slide, commits, and writes committed bytes to a
sequential hashing sink. Input cloning, strings and sink setup precede timing;
readback, preservation/patch oracles, digest finalization and drops follow it.
This is fully materialized logical append to an existing document. It provides
no bounded-memory streaming, cold/remote I/O, worker-scaling, independent-native
producer, image/rendering, or general Office compatibility claim.

The representative append category contains one measured ODP row and two
correctness-only fresh streaming rows. The index validator requires at least
one measured row for a measured category, forbids unsupported/not-applicable
rows in that category, and binds every measured row to the checked catalog and
actual full-run report. Merely adding an identity does not pass timing gates.

Original checked artifacts are retained under `inputs/`, and promoted checked
artifacts under `checked/`. `promote.py` proves all prior 198 row identities are
unchanged and checks exact Rust/Python catalog derivation before updating the
identity policy. The [source review](source-review.md) describes the unchanged
semantic contract; [native image follow-up](native-image-followup.md) records a
separate static lead without making a new application-acceptance claim.

The [summary](summary.json) retains both repeats separately. Normal ODP p50s
are 1.691887/1.687518 ms (tiny), 67.786424/67.745553 ms (medium), and
136.334131/136.843521 ms (large). The formal matrix retains 6,030 normal
samples and 90 allocator samples, including 180 ODP samples across both
instruments. Allocation totals are separate from normal timing; process RSS
includes all 201 cases in a normal run but only three ODP cases in an allocator
run, so those RSS observations are not a matched comparison.

Validation passes: 353 harness library tests (one ignored), all-feature and
all-target Clippy with warnings denied, warning-denied rustdoc, package-scoped
formatting, crate boundaries, 167 Python tests, and both official full-report
CRUD timing validators. The initial format wrapping difference and two stale
Python hash/count expectations are retained in failed receipts; latest retries
pass without changing captured measurements.

[Precleanup proof](precleanup.json), [fresh-copy portable replay](portable-verification.json),
and [five resealed negative probes](negative-probes.json) pass. The probes
reject a one-nanosecond summary mutation, a false patch gate, an omitted ODP
row, a changed checked identity count, and allocator fields in a normal row.
[Cleanup](cleanup.json) removed two owned binaries totaling 116,542,728 bytes;
the temporary task directory is absent.

```sh
python3 -B docs/performance/results/change-0465/verify.py
python3 -B docs/performance/results/change-0465/negative-probes.py
```

The verifier uses retained helper modules and needs no temporary executable.
For a new capture, rebuild the recorded source with the pinned toolchain and
use fresh output paths and new binary/source bindings; existing captures are
immutable. See [next work](next-work.md). The full non-iWork goal remains open.

A final portability correction limits comparison with the repository's current
catalog to precleanup mode. The [catalog-replacement test](portable-catalog-test.json)
proves that default portable replay accepts unchanged retained evidence beside
a later repository catalog, while live precleanup rejects that mismatch.
Prior finalization receipts are retained under `verification-history/`; the
exact binaries were restored from hash-verified build cache to repeat the live
proof and cleanup after this verifier-only change. No captures changed.
