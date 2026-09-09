# 0487 implementation and contract review

The production change appends each short replay read into the unused tail of
the existing OPC adapter buffer after the parser has consumed the preceding
slice. A full window or logical fragment end flushes that contiguous prefix.
The separate EOF probe still authenticates replay completion before advancing
to the source suffix. Source prefix/suffix and fixed-payload buffering are
unchanged. No allocation, public API, dependency, or source-version policy is
added.

The callback sequence remains freshness/context check, one provider read,
source-first post-callback check, and the existing result/context/length checks.
A debug assertion records the append precondition `buffer_start == buffer_len`.
The parser sees only the newly appended tail; repeated `fill_buf` calls cannot
advance the provider before consumption. Digests and candidate length process
the contiguous consumed range once at the existing flush boundary. Work remains
charged for each parser consumption.

Independent review found no retained-prefix overwrite or EOF/phase-transition
problem. It confirmed the public partial-output contract requires exact
accepted-byte reporting, without requiring publication after every provider
read. Mutation or cancellation during a subsequent callback can leave an
earlier consumed prefix unpublished; this is deliberate and prevents output
after authorization has failed. Ordinary provider errors flush the prior prefix
once unless source, cancellation, or sink failure takes precedence.

Six added adapter tests cover one-byte reads and exact source-version/sink-write
counts, no read-ahead, ordinary provider error, retained-prefix short-sink
failure, callback-injected source/cancellation/error combinations, and Work-limit
prefix handling. The boundary matrix includes 3/4/7-byte payloads at a four-byte
window and 128 bytes at a sixteen-byte window. Source-version probes equal
`4 + 2 * payload_bytes + 3 * window_flushes` in that one-byte-reader fixture;
checks still scale with callbacks as required. All 16 adapter tests pass.
A public Stored/Deflate regression also exercises a provider I/O failure during
publication and requires the reported count to equal actual sink acceptance.

The ADR obligations and pre-change hypothesis are in `hypothesis.md`.
The independent review corrected an early expected-probe-count assertion and
moved fault injection into the provider callback before the draft was applied.
The unrelated Keynote formatting change remains outside this batch.
