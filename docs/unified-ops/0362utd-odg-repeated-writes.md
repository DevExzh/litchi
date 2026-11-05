# 0362utd — ODG ordered repeated writes and canonical selector effects

This successor does not amend the earlier repeated-write source commit
`fca90999a7`; the selector-alias correction is its child,
`13b4e0c2e7`, whose parent is `fca90999a7`.

## Ordered payload effects

Joined-effect validation now deduplicates the durable write keys within each
individual `SubEdit`, so an intentionally ordered payload can write one
semantic target more than once.  The actual keys must be covered by the
declared write set; conservative extra declarations remain allowed.  Actual
keys are accumulated across distinct sub-edits and an overlap is refused with
the recoverable typed refusal, without consuming the plan or mutating the
owner.  Staged/package conflict checks use the same payload-derived keys.

The regression builds one durable payload that renames a page from its source
name to `Duplicate once` and then to `Duplicate twice`; it joins and commits
the sequence.  A second regression hides a page-zero write under a declared
page-one effect in a distinct sub-edit; admission refuses it while preserving
the owner and rejected plan.  The valid ordered mixed package-replacement
payload remains covered by the existing transition tests.

## Canonical selector conflict proof

Durable page and shape selectors are parsed to numeric page/shape indices when
deriving actual effect keys.  Thus accepted wire spellings such as
`page/00` and `page/0`, or `page/00/shape/00` and `page/0/shape/0`, retain their
original spelling and execution behavior but cannot bypass joined or staged
overlap checks.  Declared keys receive the same canonicalization for coverage;
unknown/custom declarations remain opaque.  The page alias regression places
canonical and leading-zero writes in distinct core-declared sub-edits, while
the shape alias regression checks a leading-zero joined write against an
ordinary staged canonical write.

## Validation

The final source commit was validated with the memory-capped cargo shim, one
job, `CARGO_INCREMENTAL=0`, debug information disabled, offline locked
resolution, and `TMPDIR=(deleted-coord-artifact)`.
The isolated target was
`(deleted-coord-artifact)`, measured at
415 MiB after the run.  External receipts are intentionally not tracked:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg -j 1` | `(deleted-coord-artifact)/u-final-check2.log` | 162 bytes | `36c58d9dd887a198d3592545144b14c3a6eb539f8121deb569529e1ca3538814` |
| `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -j 1 -- -D warnings` | `(deleted-coord-artifact)/u-final-clippy.log` | 162 bytes | `22cd5daeeb732537aa1078eb577e027bba3bb114d8b504140ff3d655acd54edd` |
| `cargo test --offline --locked --all-features -p litchi-odg --test package_semantics --test transition_review --test capability_transactions -j 1` | `(deleted-coord-artifact)/u-final-tests.log` | 5078 bytes | `ccca2182bfbbfb310b06885915941b76f0e667864c20b0e05df930107167e060` |

The final focused run passed 8 capability tests, 26 package-semantic tests,
and 21 transition-review tests (55 total).  The standalone transition rerun
also passed all 21 tests (`u-transition-tests2.log`, 2016 bytes,
SHA-256 `0716a4568be222589dbd816e9cecc08fb71f7412f93b49f24ff64570313ca647`).
`cargo fmt -p litchi-odg -- --check`, `git diff --check`, the checked API
glossary policy parse (55 profiles), and the repository artifact guards passed.
No benchmark, fixture, binary, or large repository evidence artifact was
added.
