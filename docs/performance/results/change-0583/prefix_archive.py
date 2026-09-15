#!/usr/bin/env python3
"""Reconstruct change 0580's `archive.rs` from the change 0583 working tree.

    prefix_archive.py <fixed archive.rs> <output path>

Change 0583's whole production change is two hunks in
`crates/soapberry-zip/src/archive.rs`: a new `neighbour_payload_length`, and a
`local_span_bound_from_fixed` split into a descriptor branch and an exact one.
Undoing both yields the file change 0582 measured, whose sha256 is asserted
below -- so the "pre-fix" tree this reproduces is the same code change 0582's
"after" build ran, not an approximation of it.

Pairing that file with change 0583's `office.rs`, which carries the six new
tests and no production change, is what `prechange-tests.txt` records.
"""

import hashlib
import pathlib
import sys

EXPECTED = "df9ed280e21112cbd99fd5bb373f50b676c7d592e0b15117ecff6e89bc673b27"

FIXED_DOC = """/// The fixed header carries both variable-region lengths, so the only open term
/// is the payload. A record with no declared data descriptor gets an exact span
/// end built from the larger of its two declared payload lengths (see
/// [`neighbour_payload_length`]). A record that declares one keeps the central
/// payload length: that value is where the descriptor is read from, not merely a
/// threshold, so it must stay where the record's own framing puts it."""

PRIOR_DOC = """/// The payload end is exact: the fixed header carries both variable-region
/// lengths, and the strict paths always take the payload size from the central
/// record. Only a declared data descriptor leaves the span end open."""

FIXED_BODY = """    let variable_end = entry
        .local_header_offset
        .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
        .and_then(|offset| offset.checked_add(variable_length))
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    if entry.has_data_descriptor {
        let payload_end = variable_end
            .checked_add(entry.compressed_size)
            .ok_or_else(|| Error::from(ErrorKind::Eof))?;
        return Ok(LocalSpanBound::Descriptor {
            payload_end,
            local_zip64_sentinel: file_header.compressed_size == u32::MAX
                || file_header.uncompressed_size == u32::MAX,
        });
    }
    let payload_end = variable_end
        .checked_add(neighbour_payload_length(&file_header, entry))
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    Ok(LocalSpanBound::Exact(payload_end))
}"""

PRIOR_BODY = """    let payload_end = entry
        .local_header_offset
        .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
        .and_then(|offset| offset.checked_add(variable_length))
        .and_then(|offset| offset.checked_add(entry.compressed_size))
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    Ok(if entry.has_data_descriptor {
        LocalSpanBound::Descriptor {
            payload_end,
            local_zip64_sentinel: file_header.compressed_size == u32::MAX
                || file_header.uncompressed_size == u32::MAX,
        }
    } else {
        LocalSpanBound::Exact(payload_end)
    })
}"""

HELPER_START = "/// The payload length one neighbouring record's *exact* span bound must"
HELPER_END = "/// Derive one record's span bound from its 30-byte fixed local header."


def main(source, destination):
    src = pathlib.Path(source).read_text()
    start, end = src.index(HELPER_START), src.index(HELPER_END)
    src = src[:start] + src[end:]
    for old, new in ((FIXED_DOC, PRIOR_DOC), (FIXED_BODY, PRIOR_BODY)):
        if src.count(old) != 1:
            raise SystemExit(f"expected exactly one occurrence of:\n{old}")
        src = src.replace(old, new)
    pathlib.Path(destination).write_text(src)
    digest = hashlib.sha256(src.encode()).hexdigest()
    verdict = "matches change 0582" if digest == EXPECTED else f"DOES NOT MATCH {EXPECTED}"
    print(f"pre-fix archive.rs sha256 {digest} ({verdict})")
    return 0 if digest == EXPECTED else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
