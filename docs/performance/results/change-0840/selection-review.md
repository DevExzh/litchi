# Selection and scope

The previous goal turn made progress: commit f43040e1da adopted cached Part
scheduling with fresh evidence. Current HEAD and the three unrelated files were
checked again; all 35 previously read normative input hashes remain unchanged.
The broad goal is not complete. OLE2/OOXML remain the priority, ODF is deferred,
and iWork is excluded under the standing owner decisions.

Root found a concrete remaining work-elimination opportunity in the fresh CFB
writer. `OleWriter::write_to` allocates large streams before the ministream, but
emits the ministream before large streams. On an initially empty Cursor<Vec>,
writing at the later ministream offset initializes the gap across the large
streams before their payload overwrites it. Change 0753 identified destination
zero-fill/page-fault work as substantial in the remaining DOC payload-heavy
writer. This is historical motivation, not a current speedup estimate.

The candidate only moves the ministream emission block after the large stream
loop. It leaves physical allocation, offsets, headers, FAT/DIFAT, reserved-hole
clearing, padding, the source-layout route and flush behavior unchanged. No
compression policy or durable wire changes are needed. Exact output comparison,
public DOC readback and independent CFB parsing must qualify both legs before
comparative capture. A counting seekable sink measures skipped gaps separately
from timing. Native and allocator builds remain separate; allocation growth and
whole-child RSS are independent guards because changing write order can change
Vec growth.

ADR 0001 priorities, 0002/0024 ownership, 0005 explicit I/O and evidence,
0006 validation/preservation, and 0026 directory metadata remain binding. This
change introduces neither concurrency nor unsafe code. Interrupted/short writes,
nonzero reused sinks, version-3/version-4 layouts and typed sink errors receive
focused tests. Existing affected-owner gates remain required.

The frozen nine-case plan requires a practical DOC payload workflow benefit,
not only a low-level synthetic win. Historical timings are not pooled. The
probe excludes input generation, verification and returned output destruction
from the measured operation; report that boundary explicitly rather than
claiming direct comparability to 0753's older timer.

Read-only alternative review found no materially promising source-local,
byte-transparent Deflate change. The measured replacement has the same 4 MiB
length as its original, so a length shortcut does not help; dependency-level
codec dispatch remains a separate hypothesis. Opt-in weaker durability already
exists and was measured in 0821; it is not a new candidate.

Mixed delayed batches have a large scaling opportunity, but grouping a
below-floor member with larger work changes the task unit in accepted ADR 0031
§7. That is not authorized by the existing physical read-coalescing work. The
current serial fallback stays intact. The compatible CFB emission change can
proceed without requesting a decision on that independent policy question.

Independent ROI review confirmed this is a non-duplicate target in the measured
legacy path: DOC hands WordDocument, 1Table, Data and metadata streams to the
fresh CFB writer. The reviewer found no hard blocker for the minimal swap,
provided the planned source-layout route remains untouched. A failed sink may
observe a different partial-byte sequence because write order intentionally
changes; the preserved contract is typed error propagation, complete successful
bytes, and flush behavior, not an identical failed write trace. DIFAT-before-FAT
allocation but FAT-before-DIFAT emission is a smaller distinct remaining gap and
is deliberately outside this production patch.
