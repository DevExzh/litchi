# 0522 cell-reference candidate design review

The candidate is recorded in [`cell-reference-candidate.patch`](cell-reference-candidate.patch). It is an artifact only; it has not been applied to the production checkout. The proposed source was formatted in a private temporary copy and the patch was generated from that copy. No build, test, or capture was run for this review.

## Scope and measured owner

The retained 0521 candidate Callgrind fourth dumps identify the layout scanner as a material part of the source-backed edit/save path. The profile owner is `litchi_xlsx::cell_values::source::MultiSourceEdit::commit`. Callgrind Ir is a diagnostic instruction count, not elapsed time or allocation count. The collection-off call metadata in these dumps includes setup and readback activity, so it cannot establish timed event counts or allocation counts; the separate allocator lane is authoritative for allocation measurements.

Evidence is in [`candidate/profile-r2-dense-sparse.inclusive.txt`](../change-0521/candidate/profile-r2-dense-sparse.inclusive.txt), [`candidate/profile-r2-dense-sparse.self.txt`](../change-0521/candidate/profile-r2-dense-sparse.self.txt), [`candidate/profile-r2-medium.inclusive.txt`](../change-0521/candidate/profile-r2-medium.inclusive.txt), and [`candidate/profile-r2-medium.self.txt`](../change-0521/candidate/profile-r2-medium.self.txt). The dense fourth dump reports 377,135,098 total Ir, with 137,508,563 in `rewrite`'s `scan_with_limit` edge. The selected scanner edges include:

- `Scanner::start_cell`: 42,197,278 Ir.
- `Scanner::cell_address`: 26,834,782 Ir.
- `unqualified_attribute_value` under `cell_address`: 21,695,460 Ir.
- `cell_tag`: 14,401,728 Ir; its attribute iterator accounts for 7,234,027 Ir.
- `parse_a1`: 3,219,445 Ir.

The medium fourth dump reports 197,089,142 total Ir and 71,001,069 Ir in `scan_with_limit`. Its corresponding `start_cell`, `cell_address`, `unqualified_attribute_value`, `cell_tag`, and `parse_a1` costs are 20,978,620, 13,147,503, 10,616,903, 7,333,453, and 1,532,544 Ir, respectively. These figures establish redundant work; they do not predict the candidate's final latency or allocation result.

## Proposed operation

The patch adds a private `cell_reference` helper in [`scan.rs`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs). It performs one checked `element.attributes()` iteration and returns the decoded unqualified `r` value as `Option<Cow<'a, str>>` together with a `plain` proof bit:

- `plain` starts as the raw element-name test `element.name().as_ref() == b"c"`.
- After every successfully parsed attribute, `plain` is ANDed with the raw key test `attribute.key.as_ref() == b"r"`.
- The unqualified reference match remains the same `prefix().is_none()` and `local_name() == b"r"` test used by the shared helper.
- The matching value uses the same `Explicit1_0` decoder and returns a `Cow` instead of forcing `.into_owned()`.

`Scanner::cell_address` consumes this result, retaining its existing row lookup, A1 parse, row-mismatch error, inferred-column path, `last_column` update, and `Address::at` validation. It returns `(Address, plain)`. The start-cell and empty-cell handlers call the existing `cell_tag` only when `plain` is false. All other tag paths remain unchanged.

## Error and ordering proof

The old cell path is `cell_address` followed by `cell_tag`. The first operation is the shared `unqualified_attribute_value(element, b"r", decoder)` loop. `cell_reference` preserves that loop's observable behavior exactly:

1. `element.attributes()` keeps quick-xml's checked iterator, including malformed-attribute and duplicate-name errors. Iterator errors map to `XmlError::Malformed(error.to_string())` exactly as before.
2. Attributes are examined in source order. Unknown attributes are not UTF-8 name checked or value decoded during the reference pass, exactly as before. This is what lets a reference decode/coordinate error precede a later tag-attribute error.
3. An unqualified `r` is decoded and normalized at the same encounter point with the same decoder/version. Decode errors map to `XmlError::Malformed(error.to_string())` exactly as before.
4. A second matching unqualified `r` takes the explicit `XmlError::Invalid("duplicate XML attribute 'r'")` branch, with the same `String::from_utf8_lossy` message. quick-xml's checked iterator still reports duplicate raw keys before this branch where it did previously.
5. `parse_a1`, row comparison, inferred-column bounds, pending-row mutation, and typed address bounds run in their original order. A borrowed `Cow` does not alter the decoded lexical value passed to `parse_a1` or the row error formatter.

On a successful call, the proof bit can be true only for raw `c` with zero attributes or one successfully parsed attribute whose raw key is exactly `r`: the checked iterator rejects duplicate raw keys, and the bit is false for every other raw name or attribute key. Those are exactly the cases in which the existing `cell_tag` returns `None`. Therefore skipping `cell_tag` in this branch cannot skip an unknown attribute validation, a namespace-bearing tag, or a retained lexical attribute.

When the proof bit is false, the original `cell_tag(element, decoder)?` call remains after `cell_address`, so its UTF-8 checks, normalized value checks, source-order errors, namespace/prefix retention, and owned `Tag` representation are unchanged. Prefixed cell names, prefixed attributes, `xmlns` declarations, style/type/metadata attributes, and future attributes all take this path. Namespace routing and `scan_guard` execute before this helper as before.

The patch does not change XML event/depth limits, namespace resolution, source spans, writer behavior, semantic parsing, full worksheet validation, staged readback, or any 0514/0516 fusion. The compact branch already discarded the original tag representation; the writer reconstructs the same compact cell tag, so no lossless lexical bytes are newly discarded.

## Remaining guards before adoption

The artifact still needs the planned candidate campaign after a baseline is sealed. The required guards are differential error and output checks for empty/inferred cells, encoded and whitespace-padded references, mismatched rows, invalid references, malformed names/values, duplicate attributes, truncated starts, prefixed cells, `xmlns`/future attributes, vendor-extension input, managed-cell paths, and noncompact cells. The normal XLSX correctness/resource oracles and the allocation and Callgrind stages must run unchanged. In particular, the profile must establish the actual whole-commit gain and allocation effect; the 0521 counts alone are attribution evidence, not an acceptance result.
