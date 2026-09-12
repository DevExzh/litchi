# Root review: storage change and competing scanner work

The current scanner has one primary-span consumer: `copy_without` in
`write_cell`. Both callers already have the owning `RowSlot`: ordinary row
rewriting and the 0525 replacement-row provenance writer. A row arena is
therefore a coherent private ownership change; it does not need a semantic
parser handoff or a new public API.

The arena must append a span at exactly the current two event sites: an empty
primary child, or the close of a nonempty primary child. Cell start records the
current arena length; cell close records the resulting range. Empty cells
record an empty range at the current length. Empty rows own an empty arena.
Each range belongs to its row, not the preceding row or a global worksheet
array. Only payload replacement consumes the range; style-only edits continue
to copy the original body. Multiple and repeated payload spans must remain in
order, including gaps containing unknown markup. Treating them as one bounding
span would delete those gaps and violate preservation.

The source scanner's row/cell/formula errors and event/depth limits remain in
the same event order. Checked range lookup belongs in the existing nonempty
payload-replacement branch. It must return a typed internal refusal on a
broken private range; it must not silently fall back to copying or removing a
guessed body. This does not replace the complete output validator or the
independent actual-output semantic readback introduced before this batch.

A row arena avoids the current per-cell vector lifetime and shrink-to-box
operation. It also adds row storage and changes allocation size/growth
patterns. Allocation counts, allocated bytes, live peaks and whole-child RSS
must therefore be measured separately. Existing inclusive allocator edges
contain inlining and shared callers; their complete costs cannot automatically
be assigned to primary spans or treated as removable latency.

## Other source observations

* `Scanner::start_cell` still includes cell address parsing and `cell_tag`.
  The rejected 0522 combined scan already targeted that work. Do not relaunch
  it on the strength of its current instruction share.
* The scanner computes a resolved namespace for `Event::End` and then ignores
  it. The pinned quick-xml `resolve_event` resolves the end-name prefix without
  mutating the resolver. Deferring this lookup to start/empty events is a
  separate possible follow-up; `NsReader::read_event` must still maintain
  namespace scope and raw closing-name checks. The value-only XML validator
  **uses** the resolved end namespace, so it cannot receive that shortcut.
  This observation is not a measured speedup or permission to skip validation.
* Repeated `QName::local_name` calls and shared namespace helpers remain visible
  in the scanner profile. Source-level repetition does not prove that all of
  those calls survive optimization or can be removed without changing guards.
  Keep such tuning separate from the row-storage candidate.
* Store merge remains a secondary owner. Do not move to its smaller measured
  share merely because sorting or cloning appears easier to change.

The next action is a fresh baseline and bounded native/allocator pilot for the
reviewed arena patch, with thresholds fixed before capture. Production remains
unchanged until correctness and useful end-to-end results justify retention.
