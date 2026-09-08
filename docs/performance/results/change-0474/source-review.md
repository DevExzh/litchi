# Source audit before PPTX measurement

Independent read-only reviewer pptx_streaming_review_0474 traced the public
StreamingPresentationWriter through litchi-opc and soapberry-zip. Root checked
the authoring structs, direct fragment writer and physical name maps as well.

The format writer carries scalar slide/box/text/XML counters and one active
PartWriter. Text XML is escaped directly to the part; integer formatting uses
a 20-byte stack buffer. There is no retained deck or slide XML String in this
fresh path. max_slide_xml_bytes counts emitted logical bytes and reserves the
closing suffix; it is not a requested-heap window. The active Deflate encoder
emits compressed data incrementally; this path does not use the sized writer's
CompressedScratch Vec.

Persistent package metadata grows with finalized members:

- soapberry_zip::StreamingArchiveWriter.names retains normalized member names
  in a HashSet<String>.
- ZipArchiveWriter.files retains a Vec<FileHeader>; file_names retains raw names
  in a Vec<u8> until central-directory output in finish.
- litchi_opc::PartNameSet.names retains folded keys and full PackURI strings.
- PartNameSet.descendants retains ancestor/descendant mappings. Distinct directory
  keys stabilize for this corpus, while prepare still allocates temporary
  ancestor and descendant strings for each new member.

Active PartWriter/StreamingArchiveEntry state also owns the current name and
its prepared ancestors. These transient allocations and text generation must
remain inside the measured operation, as must writer destruction.

The exact topology is 37 fixed members plus slide{i}.xml and its relationships
for each slide. The fixed count contains 11 layouts and 11 layout relationship
parts, two themes, one slide master and its relationships, one notes master and
its relationships, content types, package relationships, core/app properties,
presentation and presentation relationships, presProps, viewProps, tableStyles.
Member count alone cannot establish topology or owner relationships.

This source audit predicts growing total operation memory even with zero
retained output and a single active semantic slide. The frozen normal/allocator
matrix must establish the actual contribution before any transport redesign.
No numerical attribution to a particular map or vector follows from this
source inspection alone.

## Final harness review

The independent final review confirmed zero retained artifact state in the
returned corpus, correct timer/allocator boundaries, exact member ordering,
and opt-in dispatch. Root removed an imprecise metadata description and
changed the untimed semantic traversal to avoid repeatedly parsing the slide
reference list. The first focused run passes five tests, including valid ZIP
mutations for extra membership, text, geometry and a wrong valid layout target.

The initial valid-layout-retarget gap is closed: each slide must resolve to
slideLayout1.xml. Master, each of 11 layouts, reverse layout/master, notes
master and notes theme use exact relationship type, internal mode, wire target
and resolved member checks. The notes master reference occurs exactly once.
The public notes graph deliberately caps presentations at 4,096 slides, so the
8,192-slide preflight uses those typed OPC owner checks; smaller shapes also
invoke the full notes graph API. This does not claim full notes feature breadth.

The oracle does not explicitly compare every unrelated secondary relationship
(core/app, view/presentation properties and table styles), or reject all extra
stale relationship records. It checks the exact physical member sequence and
all authored slide semantics/owners. It is not a general PPTX validator or a
native Office round-trip test. The target-part `authored_part_bytes` value is
preflight presentation.xml identity, not a measured sum of slide XML counters.

## ADR scope

| Accepted constraints | Application to this batch |
| --- | --- |
| 0001, 0004 | Existing public semantic writer remains the measured API. |
| 0002, 0010, 0011, 0024 | Only the separate harness gains code; physical/ZIP oracle helpers stay in the harness, without new production dependencies. |
| 0003, 0006 | Fresh creation is explicitly separated from preservation/edit/commit. Existing production contracts are unchanged; mutations test the oracle. |
| 0005 | Caller-owned sequential output, finite writer limits, no retained archive output, actual operation allocation observations; growing metadata remains explicit. |
| 0008 | Focused and complete relevant checks are retained with source hashes and exact commands. No broader support/completion claim follows. |

All 30 previously read ADR/README hashes are unchanged; see adr-refresh.json.
