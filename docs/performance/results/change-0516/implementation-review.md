# 0516 emitted-output worksheet parser-feed implementation review

`performance_claim: none`

`claim_authorized: false`

`decision: blocked pending assigned fixes and differential evidence`

This is a preliminary, read-only implementation review for the OOXML lane. It
covers the candidate source represented by
`fourth-unit/source-manifest.json`, whose SHA-256 is
`9127ae1530b3892f4d0f7b4cc7e608b0e678617827bf2b4754b5fca33e04fdc9`.
No production files were edited and no build or test was run by this reviewer.
ODF and iWork are outside the scope.

## Scope and current shape

The candidate feeds the existing worksheet parser from writer-normalized events
only on the transaction path where `requires_store_verification` is true. The
compacted `Vec<u8>` remains the published-byte authority. A provisional parser
is retained only after complete output, eligibility, and budget checks; a
parser refusal falls back to the exact compacted bytes. The transaction keeps
the required order:

```text
compact -> x14ac/MCE/UTF-8 and worksheet parse -> web -> styles
         -> requested-change checks -> reopen/publication
```

The current snapshot has the important refusal boundary: once the provisional
parser becomes `None`, the event callback returns immediately, and the
post-output proof is evaluated only for a live parser. MCE/x14ac pre-input
checks and the dynamic namespace/serialization gate run before a normalized
event is consumed, so an x14ac value is not consumed with an empty extension
map.

## Findings

### Blocker: escaped shared-formula types bypass the special charge

`EventParser::try_charge_event` recognizes a shared formula only when the raw
attribute bytes are exactly `shared`. The ordinary parser decodes attributes
before classifying the formula, so `t="shar&#x65;d"` is also a shared formula
but currently does not set the expansion charge. Shared-formula resolution can
retain and translate the master formula for every member, so the provisional
bound is not established for this valid lexical form.

The assigned correction should decode the `t` value with the same reader
decoder before setting the charge, or conservatively drop the feed for every
formula carrying a `t` attribute. Add a differential test that compares the
candidate and exact compacted-byte parser for escaped `shared`, including a
large member set or formula payload.

### Blocker: provisional finalization can repeat owner work and hide errors

`WorksheetOutput::take_store` currently reduces `parser.finish(strings)` to
`Option` with `.ok()`. A shared-string callback or other finalization failure
is therefore discarded and the transaction reparses the same exact bytes,
which can invoke the shared-string owner a second time. This can duplicate
decompression/cache work and can change observable error precedence when the
owner is fallible or stateful.

The assigned finalizer correction must keep the exact parser authoritative for
the public result while avoiding an untracked second owner invocation. The
candidate-local refusal path may select fallback, but it must not silently
turn an owner/source failure into an unrelated successful retry. Add a callback
failure/once-only differential test and compare the typed error and phase with
the unchanged transaction.

### Blocker until proved: output equivalence relies on an implicit writer invariant

The feed resolves names and attributes through the source `NsReader` while
feeding the normalized `BytesStart`/`BytesEnd` events emitted by the writer.
This is sound only under the concrete invariant that the writer copies every
qualified name, namespace declaration, raw attribute value, and end name, and
changes only quote/spacing delimiters. The dynamic gate rejects unsafe raw
attribute forms and MCE/x14ac bindings, but
`output_feed_output_eligible`'s MCE-free `process_ooxml` fast path does not
structurally reread the output.

The follow-up must either state and test this writer invariant or add an exact
output event/namespace audit. The differential matrix needs nested prefix
rebinding, default namespace reset, namespaced attributes, aliases, and
empty-element scopes. Any failed proof must discard the provisional parser and
use the unchanged exact compacted bytes.

`Probe` is still observed on source events while the worksheet feed receives
normalized events. This is conservative for quote forms and is equivalent only
under the same writer invariant. It must remain on the existing full web-reader
fallback whenever the proof is ineligible; MCE/x14ac presence alone must not
become a new web rejection policy.

### Resource proof remains conditional

The candidate charges output capacity, parser records, text scratch, and
materialization estimates before retaining the feed. The final proof still
needs to cover shared-string table loading and all shared-formula text
materialization, including forms that were previously missed by raw detection.
`size_of` estimates for record slots do not by themselves bound `String`,
`Box`, hash-table, or external shared-string allocations. The separate memory
review and the assigned formula/finalizer fixes must close this before any
retention decision.

## Decoder and error-order result

Decoder divergence is **not a current blocker**. The workspace dependency is
`quick-xml = "0.41"` with default features, so its decoder is UTF-8. Both the
ordinary worksheet parser and the candidate source reader use
`NsReader::from_reader`, and `changed_observed` copies the declaration event
into the compacted output. Text and CDATA events are borrowed with their
source event decoder, so the source feed and exact output parser have the same
decoder behavior in this workspace.

If a downstream build enables quick-xml's `encoding` feature or changes the
ordinary parser to `from_str`, the feed must conservatively reject non-UTF-8
or unknown declarations, or otherwise establish the output decoder before
feeding text events. Passing only an explicit decoder argument is insufficient
because `BytesText::decode()` and `BytesCData::decode()` use the decoder stored
in the event.

Compaction remains the first fallible phase. A source XML/writer failure exits
before parser fallback. A parser-feed error or budget refusal is candidate
ineligibility and must be followed by parsing the exact compacted bytes; it is
not a new public diagnostic. The exact parser still owns x14ac capture before
MCE processing, and web finishing, style validation, requested-change checks,
reopen, and publication remain after grid parsing. The `.ok()` finalizer path
is the outstanding exception to verify because it may repeat the shared-string
owner and suppress its first error.

## Required follow-up before approval

Approval requires the assigned escaped-formula and finalizer changes, the
memory-bound proof, and differential tests covering:

* escaped formula type values and large shared-formula expansion;
* callback failure, callback invocation count, and exact typed error parity;
* namespace rebinding/default reset/qualified attributes and empty elements;
* exact compacted bytes, worksheet store, web bindings, style and requested
  changes, and reopen/publication results; and
* compaction, grid, web, style, and publication failure precedence.

Until those results are recorded against a refreshed source manifest, the
parser-feed candidate remains blocked and no performance claim is authorized.

