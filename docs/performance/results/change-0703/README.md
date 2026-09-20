# 0703 bounded PPTX MCE reuse trace

This packet is a diagnostic experiment for the opened PowerPoint transaction
path at revision f69d660800af19391a01c69331c919ce8c9d114c. It has no timing or
allocation claim. The only production file touched during execution is a
temporary codec instrumentation patch; it must be restored byte for byte before
the batch is retained.

The patch instruments the existing process_ooxml default profile. For every
call it records:

* a process-local call number and the phase supplied by the driver;
* raw pointer and length plus a SHA-256 digest of the exact input slice;
* output pointer, length, ownership (borrowed or owned), capacity, and SHA-256
  digest;
* the fixed default profile label and every numeric Limits::default() bound.

The digest proves byte equality for the observed slices. Pointer and length are
kept for correlation only; pointer equality is never treated as proof of
immutable ownership or lifetime. The patch uses the already available
litchi_core::EvidenceDigest, so it adds no production dependency.

The standalone probe0703 driver starts a fresh package for each bounded
iteration. It marks setup target discovery separately, then marks open,
capture, clone, edit, commit, apply, and verify. The changed transaction
reaches commit's candidate capture. Publication is traced as its own phase
because the committed snapshot may allow the existing exact-candidate reuse
path to avoid another capture. A no-op is checked for unchanged commit state,
revision, and selected text. The two-edit workflow requires both edits to
change and verifies both published texts.

The bounded source set is:

* test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx;
* generated:12x8, a deterministic marker-free generated package.

Run each workflow (noop, one, two) once per source with one iteration. That is
six fresh process runs. A second iteration is allowed when checking
fresh-repeat behavior, but the packet does not require a matrix or timing
samples.

## Root execution

From the repository root, run the packet build driver. It checks the current
codec hash from baseline.json, applies the temporary patch, preserves the
retained standalone probe lock (generating it offline only if absent), builds under /home/zhuhe/code/litchi-target-0703,
copies the binary to /home/zhuhe/code/litchi-0703-bin, and restores the codec
in a finally block while checking the workspace lock is unchanged:

    python3 docs/performance/results/change-0703/build.py

Then run the bounded fresh-process driver. It executes two repeats of each
source/workflow pair (12 processes total), retains separate stdout/stderr and
phase receipts, refuses to overwrite an existing trace, and invokes the
validator below:

    python3 docs/performance/results/change-0703/run.py

The generated source is passed as one argument by the driver. Each stderr file
is analyzed separately because the process-local call counter resets to zero.
For a single manual diagnostic, invoke probe0703 with one source, workflow,
iteration count, and phase-file path using the same environment variable named
in the driver.

Generate the focused transaction summary after the run driver. It checks all 12
fresh runs, phase order, successful publication, and fresh-repeat equality. Its
capture-to-commit counts exclude setup target discovery and post-apply semantic
verification, so it is the primary reuse receipt:

    python3 docs/performance/results/change-0703/focused-summary.py

The focused summary reports exact capture pairs, the number of commit calls
matching capture pairs, and observed owned output capacities. The larger
trace-summary also retains setup and verify calls and groups spanning capture
and apply for audit, but its all-call totals are not a cache estimate. Retention
figures sum representative raw payloads or owned output capacities per exact
digest/profile key. They are logical accounting figures for a hypothetical
bounded cache, not allocator, RSS, or policy-budget measurements. No latency
conclusion follows.

The build driver restores the source immediately after the temporary build,
even when Cargo fails. After the run driver completes, verify the original hash
and clean source state:

    test "$(sha256sum crates/litchi-ooxml-common/src/mce/codec.rs | cut -d' ' -f1)" \
      = a5b5b0aca3ec5a392bc7ae1ea6ca482653bd0a72cb9bdd4c30b8ea9faff87bee
    git diff --check

The retained trace outputs are evidence. Only the two owned build/binary
directories and ephemeral phase files are cleanup targets. Do not describe the
instrumented build as a production performance measurement.

For reproduction, use a disposable checkout of this revision and move its
existing `trace-runs/` directory aside before `run.py`; the driver intentionally
refuses to overwrite raw traces. Fresh pointer values and receipt hashes will
differ. Recompute `focused-summary.py` and `prepare-corpus.py` to compare the
phase counts, digest pairs and byte totals against the retained results.
`artifact-hashes.json` seals the recorded packet, not newly generated receipts.

The initial successful trace run and its two probe-only Clippy findings are
retained under `initial/`. Final traces were rebuilt and rerun after replacing
two single-pattern matches with `if let`; the focused results are identical.
Generated ZIP bytes are not retained: generator source and observed XML
digests define that control, as recorded by `corpus.json`.
