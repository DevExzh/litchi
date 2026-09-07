# 0454: source-preserving unnamed PPTX slide copy

The source-backed PPTX copy planner now accepts absent or empty optional slide
names. It preserves the original slide XML and uses the existing checked slide
position, native ID, relationship, and source lineage for identity. Missing
common-slide data, duplicate nonempty destination names, and nonempty source-name
collisions still refuse.

The baseline inventory found four such name refusals among 189 direct-picture
slide probes in 588 local files. Removing that guard alone made all four reach
a second refusal: noncanonical presentation relationship XML. The OPC owner now
supports an add-only lexical insertion before a validated explicit root close.
It checks source relationships against the opened catalog, generates only the
new children, retains the complete original prefix and suffix, and reparses the
candidate against the updated catalog before output. Replacement and removal
continue to require the existing canonical-source contract.

The integrated inventory then reached a third refusal: the authored-XML
compactness audit rejected whitespace already present in the original
presentation. OPC now issues opaque source XML proofs tied to the captured
source, part identity, and bytes. PPTX uses those proofs to preserve a copied
slide and to insert a generated slide reference into the original presentation.
Only the generated fragments use the authored-XML contract; the complete result
must pass bounded XML validation. Replacement proofs must belong to the
destination snapshot. Ordinary authored additions and replacements retain their
existing publication checks.

The source XML path accepts UTF-8 XML 1.0 with literal namespace bindings.
Declarations requiring another encoding, DTDs, namespace bindings requiring
entity or whitespace normalization, and ambiguous edits refuse before output.
These restrictions keep namespace comparisons and source proofs explicit.

Source metadata and validation scratch receive reservations before allocation.
Canonical relationship output is sized before serialization, and its reservation
survives publication. Splice admission charges retained fragment capacity and
edit-list growth; an oversized caller allocation receives a typed memory refusal.
Exact no-op splices share the original source authority and storage.

This is a measured coverage prerequisite. A successful operation replacing a
typed refusal does not establish a speedup. The original inventory, intermediate
name-only results, failed checks, source identities, and fixture provenance are
retained in the [evidence bundle](../results/change-0454/).

The final native replay publishes all four previously refused copies: one
LibreOffice QA slide and three slides from a POI fixture. The other 185 outcomes
are unchanged. Both owned-byte and simulated-range pilots pass the independent
preservation oracle on the pinned QA fixture: 34 original members become 37,
with the original slide and image payloads copied exactly.

## Evidence scope

The external corpus is the unmodified, hash-pinned LibreOffice QA
`smoketest.pptx`, opened independently as source and destination. Its original
producer/save chain is unknown. This is neither an independent document pair
nor a native application roundtrip. The opt-in harness compares untouched raw
ZIP records, copied slide/image bytes, retained metadata XML, relationship
targets, and eager/source-backed semantic reopen outside the API timers.

The bytes and simulated-range cases use finite independent owner budgets,
a bounded output sink, and one worker. The range adapter models caller-supplied
short reads and delay; it does not measure physical networking or cold storage.
Input/output/work counters are cumulative, while Memory/Objects/Depth gauges
must return to zero after the owning values are dropped.

## Measurements and review

The frozen matrix contains 480 matched synthetic-control samples and 60
candidate-only native samples, each lane using 30 samples after 3 warmups.
Synthetic API median changes range from −0.18% to +1.22%; process RSS changes
stay below the 5% review threshold. The native fixture's API p50/p95/p99 is
1.857/1.891/1.896 ms for bytes and 114.608/114.759/114.854 ms for the simulated
range source. Its corresponding whole-process RSS is 13,968/13,668 KiB.

All 15 absolute timing flags remain in the
[measurements](../results/change-0454/measurements.md). Three are regressions:
the second media-rich range repeat has source-open, destination-open and combined
open p99 changes of +5.84%, +9.93% and +7.87%. The first repeat's source-open p99
moves −8.97%, and logical counters are stable. Sleep/scheduling variation is a
plausible explanation, not a demonstrated cause. These phase tails remain a
limitation; the measured capability prerequisite is retained without a speedup
claim.

Release validation passed 497 OPC, 854 PPTX and 381 harness tests, strict Clippy,
rustdoc, formatting, workspace and boundary checks. The isolated fuzz driver
passed its sanitizer build, 1,000-run smoke test, Clippy and formatting after an
ownership correction. The source amendment proves that only this standalone
fuzz target changed after production measurements.

A later Markdown renderer correction is recorded separately from the frozen
capture protocol. The bundle discloses reconstruction of receipt protocol hashes
after the capture coordinator incorrectly updated them; retained intermediate
bytes and exact inverse checks make that correction reviewable. Raw measurement
reports and timings were not changed.

Final precleanup and post-cleanup verification pass all 14 required gates and
18 formal lanes. Owned temporary binaries, fuzz build files, and seven generated
PPTX outputs were removed after their identities and oracle results were
retained. The outer verifier’s post-cleanup corrections are separately recorded;
a separate-copy replay also passes. Replay checks retained attestations without
reparsing deleted outputs.

## Architecture and remaining work

OPC retains relationship grammar, generated XML, resource admission, and physical
publication ownership. PPTX retains semantic selection and the dependency-closure
guards. No archive implementation dependency, unsafe code, hidden executor,
ambient provider, or accepted ADR change is introduced. Preserve-by-default
publication retains unsupported original bytes only where the append proof is
complete; ambiguous or unknown relationship grammar remains a typed refusal.

The full non-iWork goal remains open. Native application breadth, complete CRUD
baselines, physical cold I/O, bounded existing-document append/repackaging, and
representative bounded-worker scaling still require evidence.
