# Change 0295: RTF owned file-open parsing

Date: 2026-08-27

Status: Accepted bounded ownership behavior

`performance_claim: none`

## Decision

RTF transport parsing is split into a private source-free transport phase and
two source-retention paths. Borrowed constructors such as `from_bytes` and
`parse_bytes` continue to make a fallible independent copy of the caller's
transport bytes after successful parsing. `open_with_limits` consumes the
single `Vec<u8>` produced by the bounded file reader, parses from a borrow of
that allocation, and moves the same allocation into the exact-source writer
fast path. The native file-open path therefore removes exactly one
encoded-source-sized clone that previously existed between file reading and
source retention.

For compressed RTF, the moved retained source is always the original LZFu or
MELA frame. Decompression still produces a separate bounded temporary parse
buffer, and all existing source, decompressed-size, lexer, parser, binary, and
model detachment limits and errors remain unchanged. A no-op or otherwise
source-preserving writer continues to emit the original bytes exactly;
non-default mutations continue through the existing canonical writer path.

## Scope and claim boundary

The full encoded source remains materialized in the document for exact-source
preservation. Borrowed constructors still clone, and the unified `litchi`
facade's RTF detection path remains outside this ownership optimization and
continues to clone as required by its API. Compressed inputs retain both the
original frame and the temporary decompressed parse buffer while parsing.

This record makes no claim about RSS, allocation counts beyond the specified
clone boundary, latency, throughput, physical I/O, decompression work, or
general whole-operation memory use. It also does not change the public API or
the writer's canonicalization behavior after edits.

Focused inline tests cover native and compressed file-open exactness,
borrowed-caller independence, compressed and source limits, malformed
compressed input, and canonical output after a non-default edit.
