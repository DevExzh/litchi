# 0425 codec chunk migration source review

This is a read-only review of the codec and image-conversion changes in the
working tree against baseline `340cc91ae`. The reviewed diff changes 33 fixed
chunk iteration sites across the crypto, fonts, image-conversion, and
OLE-common owners from `chunks_exact`/`chunks_exact_mut` to
`as_chunks::<N>().0.iter()`/`as_chunks_mut::<N>().0.iter_mut()`. No Cargo,
build, test, profiler, or CPU command was run for this review.

## Result

No migration-specific correctness blocker was found. For every changed site,
the new iterator visits the same complete chunks as the old exact-chunk
iterator. The tuple remainder is intentionally ignored in both forms. The
review found no changed use of `remainder()`, no moved padding write, and no
changed row or stride calculation.

The two remaining dynamic exact-chunk paths are significant guard points:

- [`render_text_element`](../../../../../crates/litchi-imgconv/src/emf/svg/converter.rs:1896)
  keeps `chunks_exact(stride)` because `stride` is runtime-selected (`1` or
  `2`). Its incomplete suffix behavior is unchanged from baseline.
- [`bezier_path`](../../../../../crates/litchi-imgconv/src/emfplus/renderer.rs:908)
  still ignores a point suffix that is not a complete three-point curve, as it
  did with `chunks_exact(3)`. The playback path normally supplies point arrays
  from a counted record, but this helper has no local multiple-of-three check;
  strict malformed-input rejection would be a separate hardening change.

## Crypto

- [`validate_provider_name`](../../../../../crates/litchi-crypto/src/ooxml/standard.rs:254)
  still rejects input shorter than one UTF-16 unit or with an odd byte count
  before iterating the body. `body_end = len - 2` is therefore even, the final
  two-byte terminator check is unchanged, and the new typed pairs cover the
  entire body. Invalid UTF-16 and interior terminator checks remain in the
  same order.
- [`crypt_blocks`](../../../../../crates/litchi-crypto/src/ooxml/standard.rs:365)
  retains the `len.is_multiple_of(BLOCK)` guard immediately before the mutable
  fixed-size iteration. AES-128 still mutates every 16-byte block and rejects
  a partial block before any mutation. `encrypt_package` continues to round
  up, zero-fill the package buffer, copy the clear payload, and encrypt the
  complete ciphertext range at lines 298-323; `decrypt_package` keeps its
  length check, decrypt, copy, and truncate sequence at lines 327-357.
- [`SliceReader::unicode_lpp4`](../../../../../crates/litchi-crypto/src/spaces.rs:948)
  retains the even-length check before decoding and retains the separately
  consumed, zero-validated four-byte padding at lines 964-967. The CLSID
  decoder at lines 1490-1530 still validates exact field lengths (`4` and
  `12` bytes) before its pair loops, so no hexadecimal digit or byte is lost.

The crypto replacements therefore preserve AES block boundaries, encrypted
package padding, UTF-16 unit boundaries, and DataSpaces padding validation.

## Fonts

- [`Name::units`](../../../../../crates/litchi-fonts/src/embedding/powerpoint.rs:86)
  is reached only through the private `Name { bytes }` construction in the
  validated EOT view. `parse_with` rejects odd name sizes at lines 157-163,
  checks the name range, and checks each inter-name padding word at lines
  168-193. The borrowed iterator consequently sees all name units.
- The EOT validation loop at lines 174-184 keeps malformed-surrogate
  rejection, and the sfnt name path at lines 715-726 keeps its odd-byte check
  before the big-endian pair loop. Neither path changes allocation or padding
  behavior.

The public `Name::bytes` accessor can expose the encoded bytes, but it cannot
construct a `Name` with an odd slice; the field remains private and all parser
construction passes the even-length checks above.

## Image conversion

- [`EmfSvgConverter::poly_record`](../../../../../crates/litchi-imgconv/src/emf/svg/converter.rs:893)
  checks `(points.len() - index) % 3 == 0` before the new three-point chunks.
  A malformed Bezier count still returns a parse error rather than dropping a
  suffix.
- Unicode EMF text at [`decode_text`](../../../../../crates/litchi-imgconv/src/emf/svg/converter.rs:1947)
  derives its byte range as `count * 2`; the range lookup proves the complete
  even slice before decoding. `parse_font` iterates a fixed 64-byte face-name
  field, so its 32 UTF-16 units are exact.
- EMF+ font data at [`decode_font`](../../../../../crates/litchi-imgconv/src/emfplus/objects.rs:633),
  draw-string data at [`draw_string`](../../../../../crates/litchi-imgconv/src/emfplus/playback.rs:1199),
  and driver-string glyphs at lines 1228-1253 all derive their byte ranges
  from a character/glyph count multiplied by two. The existing bounded range
  checks remain before iteration; no UTF-16 tail can be newly consumed or
  discarded.
- WMF checksum changes in [`WmfPlaceableHeader`](../../../../../crates/litchi-imgconv/src/wmf/parser.rs:92),
  its test builders at lines 524-651, the conformance builder at
  [`metafile_conformance.rs`](../../../../../crates/litchi-imgconv/tests/metafile_conformance.rs:671),
  and the generated placeable header at [`wmf_with_header`](../../../../../crates/litchi-imgconv/src/codec.rs:1022)
  all operate on a fixed 20-byte prefix. The checksum input remains exactly
  ten little-endian words.
- [`demultiply_rgba`](../../../../../crates/litchi-imgconv/src/raster.rs:167)
  now uses mutable four-byte chunks but retains the old full-pixel mutation:
  alpha-zero pixels are cleared and partially transparent pixels are
  unpremultiplied with the same rounding and clamp. The caller allocates the
  exact `width * height * 4` buffer through [`pixel_buffer_len`](../../../../../crates/litchi-imgconv/src/raster.rs:273)
  before rendering, so there is no row tail or padding byte for this helper to
  skip. The JPEG `rgba_to_rgb` path and all row/stride behavior are unchanged.

The only image-conversion follow-up is the retained pre-existing suffix policy
in `bezier_path` and dynamic text `dx` handling described above. This diff does
not broaden either policy.

## OLE-common and UTF-16/padding behavior

- [`normalize_encoding`](../../../../../crates/litchi-ole-common/src/custom_xml/xml.rs:112)
  still rejects odd UTF-16 XML byte lengths before collecting units and still
  validates decoded Unicode XML characters afterward.
- Binary property-set readers keep count-derived even lengths. The LPWSTR
  reader at [`composite.rs`](../../../../../crates/litchi-ole-common/src/property_set/codec/binary/composite.rs:207),
  code-page and Unicode readers at [`wire.rs`](../../../../../crates/litchi-ole-common/src/property_set/codec/binary/wire.rs:194),
  and the shared [`decode_utf16`](../../../../../crates/litchi-ole-common/src/property_set/codec/binary/wire.rs:256)
  still derive byte counts from UTF-16 units or reject odd lengths before
  decoding. Terminator search, invalid-surrogate errors, and allocation sizing
  are unchanged.
- The binary reader still applies `align4` at lines 223 and 252. Its
  top-level retained-padding policy and vector zero-filler policy remain in
  [`ValueReader::align4`](../../../../../crates/litchi-ole-common/src/property_set/codec/binary/wire.rs:88);
  the migration does not mutate or validate those bytes differently.
- User-defined hyperlink strings at [`codec.rs`](../../../../../crates/litchi-ole-common/src/property_set/user_defined/codec.rs:56)
  still derive `byte_count = units * 2`, require a terminal NUL, reject
  interior NULs and malformed surrogates, and consume the same relative
  padding at lines 97-99. Its standalone decoder retains the odd-byte and
  empty-payload checks at lines 409-417.
- Smart-tag UTF-16 PBString data at [`smart_tags/codec.rs`](../../../../../crates/litchi-ole-common/src/smart_tags/codec.rs:47)
  uses `count * 2`, while toolbar `WString` wire input rejects odd payloads at
  lines 55-61 before [`encoded_units`](../../../../../crates/litchi-ole-common/src/toolbar/text_icon.rs:361)
  is called. Construction from Rust units also always emits two bytes per
  unit. These replacements preserve the existing NUL and surrogate checks.

The OLE-common migration therefore leaves terminators, UTF-16 boundaries,
alignment, retained nonzero padding, and serialized zero-padding behavior
unchanged.

## Security and verification boundary

`as_chunks::<N>().0` is deliberately a full-chunk view; it does not inspect
the tuple remainder. Every changed wire-facing call site has an even/fixed
length or count-derived precondition, except the two existing rendering
helpers called out above. A future strictness change should add an explicit
remainder error before those helpers rather than relying on a chunk API.

This review did not run the parent agent's serialized Cargo gates. The
conclusions are source-diff conclusions and should be paired with the planned
codec, malformed-input, UTF-16, AES alignment, WMF checksum, and raster pixel
coverage once the frozen worktree is tested by the root agent.
