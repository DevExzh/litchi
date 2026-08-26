# Change 0296: unified RTF owned-byte parse handoff

## Status

Implemented as a deterministic ownership and correctness change. No
performance claim is made.

## Production scope

The RTF model now exposes additive consuming constructors for owned transport
bytes, with and without an explicit parse limit profile. The immutable RTF
facade exposes matching constructors. Both delegate to the existing bounded
transport parser and move the original `Vec<u8>` into exact-source retention
only after parsing succeeds.

The unified `litchi::Document` route now passes `DetectedFormat::Rtf` bytes to
that consuming model constructor. Detection already owns and moves the buffer;
the handoff no longer creates a borrowed-parser preserved-source clone. Error
mapping, feature gates, ZIP/OLE2 precedence, and the raw `DocumentImpl::Rtf`
variant remain unchanged.

Plain, literal CP-1252, LZFu, and stored MELA transports remain covered. The
compressed frame is retained as the exact source while its bounded
decompressed parser buffer remains temporary. Borrowed parsing keeps its
independent-copy contract. Detection remains responsible only for recognizing
and moving the source; it does not parse RTF. The preceding 0295 relationship
accounting work is outside this change and is not coupled to the RTF handoff.

## Claim boundary

`performance_claim: none`. This change does not claim an allocator-count,
copy-byte, RSS, latency, throughput, physical-I/O, or zero-copy result. It
does not remove compressed-input decompression or the temporary decoded parser
buffer, and it does not broaden RTF semantic, edit, publication, or
real-producer coverage.
