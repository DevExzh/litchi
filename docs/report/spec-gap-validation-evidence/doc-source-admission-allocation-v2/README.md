# DOC source admission allocation evidence

This experiment compares `c85997bc0d93ea523731d6e7622b2300e1747255` with the same base plus the eight source changes committed in `b2e2ff8e1fdabd4ab16a67ddf077e6968bf5e668`. The measured candidate is an eight-file overlay, not the complete later commit tree. Correctness and production gates are recorded in [the source-admission batch](../doc-source-admission-v1/README.md).

The recorded runs use Rust 1.95.0, Cargo's unoptimized `dev` profile, and x86-64 Linux on an AMD EPYC 9R45 host. These are allocation-only results for deterministic synthetic fixtures. Each side has sixteen operation/profile groups, three warmups and fifteen recorded samples per group. The supplemental and portable runs reproduce all 480 original non-pointer rows and all medians.

Representative large-profile medians:

| Operation | Control allocated bytes | Candidate allocated bytes | Control allocation calls | Candidate allocation calls |
| --- | ---: | ---: | ---: | ---: |
| Common editor open and unique unchanged finish | 9,282,872 | 7,958,328 | 287 | 286 |
| DOC embedded-object snapshot open | 16,973,245 | 12,412,349 | 506 | 504 |
| DOC storage replacement and outer commit | 51,984,582 | 47,423,753 | 1,676 | 1,679 |
| DOC malformed replacement admission | 1,255,261 | 715 | 101 | 21 |

The large common source is 1,324,544 bytes; the DOC source is 2,280,448 bytes; its successful replacement is 1,059,840 bytes. Successful replacement reduces allocated bytes but adds three allocation calls. The full sixteen-group results, including small inputs, shared finish, oversized rejection and malformed snapshot admission, remain in the archived receipts and raw reports.

The harness uses `embedded_object::Snapshot` and `Transaction::replace_storage`, then commits and reopens the DOC to check the exact replacement payload, references and typed metadata. No-op bytes and source sharing are checked. The supplemental correctness run also explicitly records shared-source identity after oversized rejection. Root reran the portable driver, checked its correctness booleans and compared every non-pointer measurement field with the supplemental run.

Interpretation limits:

- Compare allocated bytes and allocation calls only. Setup allocations can be freed inside the counted region, so deallocation, net-live and peak-live fields do not define comparable ownership boundaries. No latency, throughput, RSS, syscall or release-build improvement is established.
- Common-editor counted closures include target-catalog and limits construction. Absolute totals include that caller setup; it is identical on both sides. Source and replacement byte construction are outside the counted regions. Successful DOC replacement measures replacement plus outer commit; rejection measures admission only.
- The original v2 report says six fixtures; there are eight. Its control Cargo root differs from its fixture root. The supplemental receipt and portable runner make those paths explicit. The original report's oversized shared-source claim is established by the supplemental check, not by its original correctness JSON.
- For malformed snapshot admission, `input_bytes` is the actual rejected input size; `source_bytes` retains the valid DOC fixture size as context. There are no direct common malformed-replacement groups here.
- Earlier v1 measurements used tracked-revision snapshots and cannot establish this embedded-object result. Original v2 reports are retained verbatim, including their corrected reporting caveats. The exact original command transcript was not saved; the supplemental and portable runs have recorded commands.

`evidence.tar.gz` contains eight fixture files, original and supplemental harnesses/locks/reports, the portable runner and its executed commands/results, a source patch, and a 769-file manifest covering all nine resolved path dependency crates. The outer receipt hashes every archived file. Host context was captured after the supplemental run, during the portable run. The portable run is the authoritative reproduction invocation: it uses `--locked`; the supplemental script used the copied, hashed lockfile without that flag.

To reproduce, extract the archive into a new directory. Create two detached checkouts of the recorded base commit, apply `candidate-source.patch` to the candidate checkout, then run:

```sh
python3 reproduce.py --control-source /path/to/control \
  --candidate-source /path/to/candidate --output /path/to/new-output
```

The output directory must not exist. Rust 1.95.0 must be installed; use `RUSTUP_HOME` if it is outside the default installation. The driver verifies all recorded source and fixture hashes, rewrites only local dependency paths, uses the archived Cargo lockfile, clears Rust flag overrides, and records each command. Dependency downloads may be needed on a new machine. Allocation results are scoped to the recorded build and architecture.

The separate [release latency evidence](../doc-source-admission-latency-v2/README.md)
includes a portable rebuild and root reproduction. Its large open-plus-no-op-finish
case regresses despite lower allocation totals; allocation measurements alone do
not establish a latency improvement.
