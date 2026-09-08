# Borrowed events in changed SpreadsheetML compaction

The production change removes `Event::into_owned()` in the slice-backed
`NsReader` loop. Events borrow the immutable input slice. Every branch finishes
using the event before another read; `Writer::write_event` writes synchronously
and retains no borrow. A source comment records that lifetime assumption.

The full compaction pass remains. Start attributes are still checked first for
`xml:space` and again for normalized emission; empty tags keep their checked
attribute traversal. End names retain their UTF-8 conversion. Text, CDATA,
comments, declarations, processing instructions, doctypes and general references
follow the same writer path. XML parser errors and checked-attribute errors
occur in the same order. No bounds, validation phases, output reservations,
public APIs, or Store retention thresholds change.

An independent reviewer checked quick-xml 0.41.0 slice-reader, namespace resolver
and writer lifetimes. The differential test reference retains the former owned
event loop. It compares exact bytes on success and both debug error variants
and display messages on failure. Mixed markup is explicitly required to succeed;
malformed inputs compare the existing acceptance/refusal behavior without
silently tightening the previous contract.

The hypothesis is fewer temporary event allocations and copies. The 0468 CPU
profile places the explicit ownership conversion at only 0.348% of whole-process
sample weight. No source-level argument establishes an end-to-end speedup;
matched timing and whole-process allocation measurements decide retention.
