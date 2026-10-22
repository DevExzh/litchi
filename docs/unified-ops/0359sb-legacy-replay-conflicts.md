# 0359sb — typed legacy raw replay conflicts

The 0359s audit incorrectly classified a valid patch-base mismatch as a
malformed input in seven public raw replay boundaries. This correction makes
each boundary return its owning crate's `PatchConflict` variant and migrates
the callers and assertions in the same change:

- OLE Property Set `SummaryInformation` and `DocumentSummaryInformation` and
  OLE1 use `litchi_ole_common::Error::PatchConflict`.
- DataSpaces uses `litchi_crypto::spaces::Error::PatchConflict`.
- DOC `RgDofr`, `Selsf`, and whole-package protection patches use
  `litchi_doc::package::Error::PatchConflict`.
- DOC maps both its direct conflict and the shared OLE conflict into
  `litchi_core::Error::PatchConflict`; litchi-ops classifies that leaf as
  `invalid_input`.

Reopened equal-content snapshots continue to replay. Parser, clone, and
structural validation failures retain their existing malformed-input variants,
so the facade still reports them as `corrupt_file`. The public body-text owner
now uses `Error::PatchConflict` for stale package patches, durable preconditions,
owner-bound auxiliary patches, and transfer plans. Its `Error::Conflict`
remains the distinct three-way semantic-overlap result.

The source changes and all consumer checks were validated by the commands and
log hashes recorded in `0359s-legacy-replay-classification.md`. The borrowed
debug target was 8.4 GiB after validation. The post-change error-layout probe
reported `litchi_ole_common::Error=48`,
`litchi_crypto::spaces::Error=40`, and `litchi_doc::Error=48`; no evidence or
binaries were added to the repository.
