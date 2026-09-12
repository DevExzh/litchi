# Typed XLSX error prerequisite

The public, non-exhaustive XLSX `Error` now retains two typed failures:
`ResourceLimit(litchi_core::ResourceLimit)` and
`FormControl(form_control::FormControlError)`. Callers can inspect the original
resource/observed/limit/scope or leaf error instead of parsing an `Invalid`
message. The leaf error has no dependency on the XLSX error, so this introduces
no recursive owner-error type.

Independent diff review approved the source recorded in `root-gates.json`.
Root verified it in an isolated checkout of the recorded base, with only
`crates/litchi-xlsx/src/error.rs` overlaid. This separates the prerequisite from
the unfinished pivot-data and form-control owner implementation.

Gates passed without lint suppression:

- All three `error::tests`, including the two new typed-conversion regressions.
- Strict library Clippy.
- Rustdoc with warnings denied.
- Rustfmt for the changed file.

The four compressed logs retain full gate output. This is a focused error API
gate, not approval of the unfinished owner features or a full XLSX test run.

`Cargo.lock` is untracked in the repository. The initial isolated locked command
could not run without a matching lockfile; copying the dirty workspace lockfile
also required resolution because its workspace manifest graph differed. The
isolated checkout therefore resolved dependencies with
`cargo generate-lockfile --offline`, then ran all Cargo gates with `--locked`.
The resulting exact lockfile is retained as `Cargo.lock.gz`, with its raw-byte
hash and the resolution log hash in the manifest.

To replay, check out the recorded base in a disposable directory, apply this
commit's error-file change, decompress the retained lockfile to `Cargo.lock`,
and run the recorded commands with the pinned toolchain and environment. The
owned checkout and build target may be deleted after the committed evidence and
source bytes have been verified.
