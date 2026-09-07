# Source and evidence boundaries

Microsoft documents p:cSld as the common slide-data container, and its example
uses that container without a name:
[CommonSlideData reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonslidedata?view=openxml-3.0.1).
The existing format helper already distinguishes missing common-slide data from
an absent/empty producer-visible name. The new admission uses that distinction:
Some(empty) is unnamed, while None remains a typed structural refusal.

Slides are selected by ordinal and checked native ID/relationship/source identity.
No slide XML name is generated or rewritten, and empty names do not enter the
nonempty-name collision index. Existing normalized collisions and ambiguous
nonempty destination names still refuse; dependency closure, exact owner graph,
MCE/unknown extension, package framing, source-version, semantic rerun, budget and
sink rules are unchanged.

The old static native-gap recommendation was insufficient: a real baseline probe
of the MSO TextFitting pair stops at a presentation extension list containing
p15:sldGuideLst, before names. The tracked POI bug62513 case stops at a trailing-ZIP
framing boundary. Neither refusal is relaxed here. The actual inventory inspects
588 local files, probes 189 direct-picture slides and records 13 enumeration
errors. Four slides reach the optional-name guard. These are external QA inputs;
fixture provenance and original producer/save-chain certainty are separate.

The name-only source epoch made all four name-guard cases reach publication, where
noncanonical presentation relationships still refused. That inventory is retained
as historical evidence, not a claim about the integrated candidate. OPC now needs
an add-only lexical splice before a validated explicit root closing boundary.
The planned scope accepts only a complete source-catalog match and understood
relationship grammar; replacements/removals and ambiguous or unknown grammar
remain refused. Canonical serialization continues on its existing output path.

The opt-in external harness has a separate correctness oracle outside API timers:
raw untouched ZIP local records and central records (masking only their relocated
local-header offsets), exact copied slide/image bytes, original metadata XML
prefix/suffix retention, and source-backed plus eager semantic reopen. Its range
adapter reports logical requests and simulates delays; it makes no physical I/O,
network, cold-cache, allocation-profile or native-application claim. Memory,
Objects and Depth are live gauges checked at zero after drop. InputBytes,
OutputBytes and Work remain reported cumulative consumed counters.

## Final source and cleanup proof boundary

The current source epoch includes the canonical relationship serializer memory
reservation work, fragment-capacity admission, and focused limit checks. The
canonical patch transport was retained as `candidate/canonical-memory.patch.txt`
and removed from `/tmp`; `checks/canonical-memory-patch-removal.json` records its
identity. Acceptance still requires the final source manifest,
all release and fuzz gate receipts, final native inventory, formal capture
receipts, and derived measurements to bind the same source epoch. Patch transport
cleanup is an operational attestation outside the verifier.

`remove-probe-example.py` is an exact cleanup helper for the retired
`native_cross_copy_probe_0454` example. It requires a passing precleanup
receipt, inventories the listed release-example and fingerprint paths, rejects
an unexpected fingerprint member, and writes
`checks/retired-example-artifacts.json` and
`checks/retired-example-removal.json`. Those receipts prove only that this
enumerated build residue was removed. The current verifier's generic cleanup
pass consumes `*-cleanup.json` receipts; the retired-example receipts therefore
remain a separately retained cleanup attestation and must be reviewed before
the final seal. No source or production claim should rely on the temporary
patch path or on the absence of an unlisted target artifact.
