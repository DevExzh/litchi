# 0732 PPT phase diagnostics implementation

The PPT slide-order owner now has an opt-in `performance-diagnostics` feature.
With that feature enabled, `Transaction::commit_profiled` reports synchronous
content-free `DiagnosticEvent` pairs through
`FnMut(DiagnosticEvent)`. The public event shape matches the established DOC
diagnostic seam:

```text
DiagnosticEvent::Started { phase }
DiagnosticEvent::Finished { phase, outcome }
DiagnosticOutcome::{Success, Error}
```

The ordinary `Transaction::commit` body is unchanged. The profiled copy keeps
the existing source and working snapshots, live-record `Vec`, pre-publication
payload map, output `Vec`, and post-publication payload map at the same owner
boundaries. Existing validation and digest expressions remain in the same
order; wrappers only report their surrounding expression result.

The fixed phase vocabulary is:

```text
DocumentCommit
BeforePayloadCapture
EmbeddedOpen
LiveDocumentRead
EmbeddedFinish
UnrelatedStreamValidation
PublicReopen
AfterPayloadCapture
ArtifactHashBefore
ArtifactHashAfter
StructuralNoOp
```

`StructuralNoOp` is emitted only when the live-document commit has an empty
patch. Formatting owners can still have changed the working package in that
route, so the name does not claim whole-artifact identity. Its two artifact
hash expressions remain individually observable and preserve the ordinary
no-op hash behavior.

The observer helper has no clock, global state, retained event buffer, or
allocation of its own. A phase that returns a typed error still emits its
matching `Finished { outcome: Error }` event. The source-consistency guard is
inside `LiveDocumentRead`, while the post-reopen expected-payload comparison
remains explicit residual work after `AfterPayloadCapture`, matching the
ordinary owner boundaries.

Unit coverage under the feature checks true-edit phase order, formatting-only
ordinary/profiled byte and patch parity, structural no-op source allocation
sharing and hash phases, and a forged live-document mismatch with balanced
events and the existing typed corruption error. An additional absent persisted-record
test reaches the actual record reader and compares ordinary/profiled error
classification and display, with a balanced failed phase. Cargo, rustfmt, and native
measurement are intentionally left to the root coordinator's serialized
qualification run.
