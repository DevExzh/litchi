# Independent evidence review: change 0477

This review covers the portable verifier and its mutation tests. It does not
run a build, benchmark, capture, or analysis command.

The verifier now treats the two records named by `binaries.json` as the
designated final build gates. Each gate must contain a successful, source
unchanged receipt and explicit `binary` and `original_binary`
path/size/SHA-256 records for the copied and origin executables. Its
`validation_gate` reference is content addressed and must point to an
underlying receipt whose command, environment, source snapshots, timestamps,
status, and source-unchanged result are mirrored by the wrapper. The final
source snapshots must agree with one another and with the
`source_manifest_sha256` in `binaries.json`. This is the final-build
requirement; ordinary retained validation attempts remain historical records.

The validation ledger still checks every started/final pair, preserves
non-zero development attempts, and permits a successful historical attempt to
have changed source. Both designated final build gates and required ledger
gates require exit 0 and unchanged source; retained historical attempts may
remain non-zero or source-changing. The verifier never follows the absolute
binary or temporary spool paths recorded in receipts; the build wrapper's
copy-time metadata and content-addressed gate record provide the portable
custody boundary.

`rust-validation.json` adds the final validation coverage boundary. Its
`attempts` map must enumerate every finished validation receipt with its exact
path, digest, status, and (for required gates) expected command. Each required
gate from the fixed 23-label minimum must succeed, retain unchanged source
snapshots, and match the final build source. An earlier required gate may declare the single approved
benchmark-file source exclusion when it predates the benchmark edit; the
remaining source manifest must still match the final source exactly.

The seal check requires an exact inventory of regular files below the evidence
root, excluding only `SHA256SUMS` itself. It rejects omissions, extras,
symlinks, special files, and `__pycache__` entries. The mutation test shows
why a self-consistent rewritten seal is insufficient: after a report is
changed and resealed, the retained receipt's original artifact metadata still
rejects it.

`test_evidence.py` covers resealed artifact corruption, unsuccessful designated
final gates, source-manifest drift, origin/copy binary metadata drift, and the
permitted source-changing successful historical receipt. It runs with the
standard library only and does not access external build or temporary paths.

The operation spool-path helper uses an opaque corpus vector index in its
internal temporary filename. That name is unique within the process; the
explicit outer capture directory and the report/oracle `member_count` bind the
actual 8/256/8192 case. This filename convention is recorded for review and is
not a capture blocker.

The frozen `protocol.json` intentionally binds `common.py` and `capture.py`
only; those helper hashes must remain unchanged after freeze. The other
helpers are bound by their build/gate artifacts where applicable and by the
final seal. The bundle must record the final `build.py` wrapper outputs in
`binaries.json`. The
portable verifier intentionally treats the binary paths as metadata-only
references; the wrapper remains responsible for asserting origin/copy byte and
SHA-256 equality while the build is local.
