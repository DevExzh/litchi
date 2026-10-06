# 0920 admission history — candidate rejected

Final disposition: the frozen comparison completed all 192 native and 96
allocator children, and both replay paths passed. Allocation and peak-live
gates pass, but all four FloatingPictures latency confidence lower bounds
fail the predeclared `1/1.05` guardrail. Production adoption is rejected.
See [the change record](../../0920-cfb-output-reservation-rejected.md) and
[comparison-reduced.json](comparison-reduced.json). No selective recapture or
statistical-gate change was made.

The text below is the earlier admission history. Its references to uncommitted
work, pending review, and zero performance children apply only to that stage.
[capture-readiness.json](capture-readiness.json) supersedes those pending
items at the pre-capture boundary. Historical controls remain retained;
[capacity-controls-final.json](capacity-controls-final.json) and
[qualification-replay-final.json](qualification-replay-final.json) bind the
final runner/protocol. The earlier synthetic tooling validation binds its
then-current runner; the real final replays validate the completed packet.


The current CFB candidate and comparison tooling are **not admitted** to
performance capture or production adoption. No comparison manifest has been
frozen and no comparative performance children have run.

`docs/GOAL.md` explicitly requires:

> Before editing code, read docs/adr/README.md and every accepted ADR.

Both code authors reported that they cannot certify this prerequisite was met.
The CFB author reviewed relevant constraints but not every accepted ADR; the
comparison author reported truncated ADR reads. Relevant-only review and
truncated output do not satisfy the instruction. Existing source and test
receipts remain evidence of the prototype; they do not prove process compliance.
This finding does not imply that the passing tests failed, nor can a later
read retroactively repair the original editing chronology.

Current corrective work is to read accepted ADR decisions and applicable
amendments before further code edits, with bounded untruncated reads. The
user explicitly excludes iWork from this goal; clearly iWork-only material
need not be considered. Review the existing source against those constraints
and repair concrete defects. The goal does not require reconstructing identical
source solely to repeat its editing chronology. Existing source/build evidence
may be reused after exact identity and scope checks; it does not establish
that the original authors met their read prerequisite. Preserve the archived
prototype packet and bind any new source, tooling, checks and review before
capture. No ADR semantic exception or user approval is being requested.

The prototype has 48 passing baseline and 48 passing candidate functional
qualifications, with matching paired semantic identities and permitted capacity
relations. The release configurations and 157 compiler/profile/feature/flag
fingerprint records match. CFB tests passed; additional shared-owner checks
are in progress. These are scoped prototype facts, not adoption or speedup
claims. Qualifying processes are excluded from comparison statistics.

The comparison author additionally reported an unresolved prototype defect:
`build_source_binding()` applies a SHA-256 validator to a 40-character Git
revision. The tooling repair must use the correct revision validation and test
actual source/build receipt binding, copied-packet replay, capacity controls
and the entire planned reduction size before freezing. Existing control
receipts bind their recorded script versions; later prototype edits do not
inherit those results automatically.

The read-only Git limitation was historical and has cleared. On 2026-10-06,
0919 was finalized in `13be448613`; 0920 remains an uncommitted candidate while
source review, tooling repairs and shared-owner quality checks proceed. No
comparative capture has started. The broad non-iWork goal remains open.
