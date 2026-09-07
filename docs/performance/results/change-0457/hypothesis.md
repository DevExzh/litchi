# Existing ODP tail append: bounded publication dependencies

The owned control remains `odp_existing_append_lifecycle`: open an existing
64/4,096/8,192-slide archive, append one plain title/body slide, commit, and
publish its bytes. The 0439 evidence identified growth in retained XML and
candidate buffers; this batch first recaptures that unchanged control after
the rebase. It does not treat a new baseline as an optimization comparison.

The current implementation has two independent document-sized allocations:
`SourceBackedPresentation` retains complete `content.xml`, while the common
source publisher materializes the original and replacement and prepares a
complete regenerated ZIP member. Removing either alone does not establish a
bounded append lifecycle. The work therefore needs all of the following:

1. A common ODF callback reader that verifies a member without materializing it.
2. A finite XML event window that refuses oversized tokens before parser growth.
3. A format-owned tail insertion proof and individually audited generated slide.
4. Candidate XML validation before any output, followed by bounded replay into
   a sequential ZIP sink, with every untouched member preserved.

The ZIP dependency measures a generated member in one pass and emits it in a
second. It reuses the existing sized header and ZIP64 layout planner. Replay
must compare decoded and compressed sizes, CRC, and strong digests; callback
errors, ignored sink failures, partial writes, and nondeterminism cannot become
successful archives. Compression must perform no implicit writes during Drop.
The extra pass and digest work are explicit CPU costs to measure, not a presumed
speed improvement.

The anticipated memory benefit is a flat payload-working-set slope as slide
count grows, with bounded parser/compressor buffers and archive metadata. The
source provider's retained bytes, fixture construction, and full independent
oracle buffers must remain separately attributed. This does not imply flat
memory for unlimited archive member counts, unlimited XML depth, or oversized
individual text tokens.

Source versions, signature/encryption refusals, finite hierarchical budgets,
cancellation, exact accepted-output progress, unknown source XML, and untouched
ZIP framing remain mandatory. A specialized publication plan does not by itself
implement the ordinary snapshot/commit/reversible-patch lifecycle. Any such gap
must stay explicit in the benchmark and goal audit.

The immediately preceding turn verified the requested rebase and completed the
already-running baseline build. It made no production performance change. The
safe next action is to capture the retained baseline and implement these
dependencies; there is no external blocker. The full non-iWork goal is active.
