# Independent read-only review and root disposition

Two reviewers inspected the candidate and the benchmark before formal capture.
Neither executed builds, tests, probes or scripts. Root owns all execution.

## Production review

The retained `source_ingress` flag covers borrowed/owned/reader/session opens,
clones, removal and source-backed materialization. Exact-source publication
still precedes fallback. The full publication plan and audits precede emission.
The reviewer confirmed that the bounded owned-entry wrapper wrongly poisoned
retry-safe `Interrupted` writes; the root fixed this adapter and added direct
Store/Deflate progress and payload regression coverage. Hard errors still poison.

A source-backed-specific discriminator assertion was recommended. Existing
`borrowed_materialization_preserves_opaque_xml_like_eager_unmarshal` and
`borrowed_signed_materialization_retains_publication_policy` tests exercise the
materialization seam; the explicit marker assignment remains unchanged. The
candidate's new tests cover borrowed/owned discriminator and clone/removal.
Fresh signed publication has an existing sign-feature
`signed_package_survives_zip_round_trip` regression. Supplemental sign-feature
execution was planned but not reached before rejection; no fresh-sign result
is claimed for the candidate.

The reviewer also suggested making explicit `StreamingArchiveEntry::flush`
interruptions retryable. That is outside this candidate's write-only fix:
retrying the lower flush can repeat a Sync codec call, so deterministic flush
retries need separate examination. No explicit-flush retry guarantee is added.

## Probe review

The reviewer verified the timer is exactly `PackageWriter::to_bytes` on a
prepared graph; construction, input open/edit, expected output, validation,
readback, hashing and output destruction are outside. RSS is whole-process HWM.
Cross-leg member/payload and semantic comparisons are adequate for this scope,
and opened edit controls additionally require identical complete output hashes.

Before capture the root added a fixed admission input set, original preparation
receipt binding, independent Python ZIP readback of fixture members, exact
fixture manifest/checksum validation, and a separate machine-readable adoption
decision. Thresholds now come from the frozen plan. A validation pass alone
cannot admit a candidate; at least one fresh median must improve by at least
5% with bootstrap upper endpoint below 1, and any regression/spread/size flag
requires explicit disposition. Initial script versions remain retained.

The optional probe `qualify` subcommand targets an owned opened package and is
not invoked by capture. Fresh short/Interrupted stream qualification comes from
the new production tests, not from that unused subcommand.

## Correctness rejection supersedes the provisional review

The candidate suite failed the unchanged durable paragraph-copy inverse test
at `source_backed_paragraph_copy.rs:183`. A second review traced the precise
boundary: the durable inverse stores XML and a whole-artifact source identity,
then restores the selected member with opened-package compression. The source
created by the candidate used a different protocol. Root reconstructed both
archives from the assertion log and independently decoded them with Python
ZIP: every logical member is identical, but only `word/document.xml` has a
different compressed payload. Exact physical restoration fails.

In-memory `Patch::inverse().apply` checks snapshot/XML restoration. Immediate
publication inverse can retain and copy the original artifact. Neither proves
the durable inverse after serialization and reopen. The unchanged exact-byte
test remains the gate. The initial statement that durable patch behavior would
be unaffected was disproven and is withdrawn. OPC compression changes and their
production tests were reverted; the full candidate source is retained.

A future attempt needs an explicit physical restoration representation or
proven compression provenance, with fresh-output → source-backed durable inverse
coverage before timing. Changing opened-package compression to match the new
fresh output is not authorized by the fresh-only policy.

## Final retained change review

The second reviewer confirmed the eight-line ZIP write branch and direct
Store/Deflate test cover retry progress, no double counting and full readback.
The explicit-flush limitation remains recorded. The reviewer requested root
ownership and base checks before cleanup; `close.py` now calls the original
marker/base guard while roots exist, and accepts deleted binaries only through
the successful cleanup witness. Post-commit replay checks the direct parent
and sealed files. Closure precedes cleanup, then replays after cleanup.
