# 0905 parity capture review

The retained baseline and candidate captures are byte identical. Both gzip
files are 1,416,383 bytes with SHA-256
`4e1cc9e25cf600e4afa43a5e3e1649a05a8344d6dd92de67b234e1843b709904`; their
decompressed JSONL is 74,265,040 bytes with SHA-256
`567535be0da4b7165e65ed9d923fb1a9f15b39ffb5451133a876a6c9d22b0bc3`. The two
streams contain the same 125,954 lines: one header, 125,952 records, and one
footer. Parsed JSON objects also compare equal, and both receipts agree with
these sizes and hashes.

The 125,952 records cover 11,520 case IDs from 0 through 11,519. The case
generator covers cardinalities 0, 1, 2, 31, 32, and 33 (1,920 cases each), all
128 styles 0 through 127 (90 cases each), all 15 resource profiles (768 cases
each), and all eight replacement kinds (1,440 cases each). Default-profile
cases produce ten records; nondefault profiles produce eleven because they add
the constrained source call. Every case has one source input hash and one
replacement hash, and no case has an unexpected API set.

The source-policy slice calls return 3,766 reports and 7,754 errors under
default limits, then 2,030 reports and 8,722 errors under constrained source
profiles. Authored slice calls return 2,281 reports and 9,239 errors; compact
slice calls return 2,617 reports and 8,903 errors. Each authored reader chunk
shape (`one`, `uneven`, and `prefix`) repeats 2,281 reports and the same error
breakdown. Each compact reader shape repeats 2,617 reports and its matching
breakdown. The JSON companion contains the exact `Report`/`Error` variant and
`Resource` limit counts for every policy and chunk shape.

Replacement proofs contain 706 `Identical`, 308 `Window`, and 611 `Complete`
successes. The remaining records are 9,239 original-side errors and 656
replacement-side errors. This covers replacement limit retotaling, fallback,
and both failure sides in the generated matrix.

The review recomputed every embedded record digest. The source uses the seed
`0xcbf29ce484222325`, one NUL separator, and multiplier
`0x1000_0000_01b3`; all 125,952 baseline and candidate digests match that
implemented recurrence. This multiplier is custom; the method file's label
“FNV-1a” should not be read as the standard FNV-64 prime. The authoritative
parity checks remain the raw gzip and decompressed-byte hashes above.

One generator claim needs qualification. `parity-example-v2.rs` bounds style
to 0..127, while its declaration branch tests `style & 0x80 != 0`, so no XML
declaration case is generated. BOM, malformed UTF-8, malformed attribute and
text, namespace, `xml:space`, cardinality-boundary, limit, reader-chunk, and
replacement branches remain represented by the recorded dimensions. The
public API has no source-policy streaming reader, so authored and compact are
the complete reader policy set here. Finally, differential equality establishes
that the two builds behaved identically on this corpus; it is not an
independent semantic oracle and carries no allocator or performance result.
