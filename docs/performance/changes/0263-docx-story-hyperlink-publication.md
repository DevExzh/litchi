# Change 0263: DOCX story-hyperlink publication harness

## Status

Landed in `b008784d6517fb30c9eabc318e831907c693b1dd`. This is correctness and
phase-boundary evidence for end-to-end story-hyperlink publication. It does
not establish a speedup, allocation, RSS, or physical-I/O claim.

## Scope

The opt-in selectors `docx_story_hyperlink_noop_save` and
`docx_story_hyperlink_redaction_save` use one deterministic DOCX corpus that
covers all seven relationship-reachable Word story kinds:

- main document, header, footer, footnotes, endnotes, comments, and glossary;
- 15 OPC Parts and 24 ZIP members;
- 9,900 source archive bytes;
- source archive SHA-256
  `457421e8f86ec8eb52fbe181cebe7d0821ce1e794a08142ff01a4c4e03df0cac`.

Each story contains one selected shared external hyperlink, one unselected
external hyperlink, and one media relationship. Seven deterministic media
members and one inert opaque member keep publication locality and preservation
checks independent of the selected relationship. The two selectors are
opt-in: the selectable matrix is 383 names, while the default remains 36 cases
and 198 records.

The timed interval prepares a fresh source and reserved sequential sink before
the clock, then covers open, strict target planning, commit, and sequential
publication. Source preparation, output preparation, all semantic and ZIP
oracles, and refusal probes remain outside the measured interval.

## Correctness and evidence boundaries

An independent story-XML oracle checks the story roots, hyperlink and media
references, selected/unselected link counts, and text nodes. An independent
`.rels` oracle checks namespace, relationship-member membership, relationship
IDs, targets, relationship types, and target modes for every story and the
main-story owner relationships.

The no-op selector requires exact archive bytes. The redaction selector checks
the exact changed story XML/`.rels` member set, preserves all other member
payloads, and compares untouched ZIP members using both raw local-record bytes
and central-directory records with local-header offsets normalized. Repeated
publication must be deterministic, and the measured source must retain its
exact identity throughout publication.

The harness also requires typed refusal with zero output for stale and foreign
source commits, signed packages, unknown relationship owners, and zero-output
sinks; partial sinks must fail after partial progress without being accepted as
a successful publication. These are bounded safety and preservation gates, not
claims about arbitrary DOCX extensions or producers.

## Follow-up

Any DOCX publication performance claim still requires clean release binaries,
controlled A1/B1/B2/A2 evidence, reproducible drift gates, and independently
retained resource and I/O evidence. Broader story editing, producer, signed,
extension, and structural-relationship matrices remain outside this harness.
