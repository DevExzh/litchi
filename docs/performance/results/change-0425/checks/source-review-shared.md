# 0425 shared source review

This is a read-only review of the shared/root-owned 0425 source changes against
`340cc91ae`. The reviewed set is the current diff in `litchi-core`,
`litchi-cfb`, `litchi-docx`, `litchi-odraw`, `litchi-odt`, `litchi-ograph`,
`litchi-rtf`, `litchi-vba`, and `soapberry-zip`. No Cargo command, test,
profiler, or CPU workload was run by this reviewer.

## Findings

No correctness blocker was found. The production changes are bounded
`chunks_exact`/`chunks_exact_mut` array-view conversions, with one equivalent
MD4 block-loop rewrite and two removals of impossible slice-pattern fallbacks.
The ZIP and SOAPberry changes are test assertion spelling only.

`litchi-cfb` keeps an exact-input precondition at every new fixed-width view.
FAT and MiniFAT sector buffers are always a 512- or 4096-byte sector, hence
divisible by four (`crates/litchi-cfb/src/file.rs:851-858, 917-922`). Directory
validation rejects an empty or non-128-byte-aligned stream before iterating
128-byte entries (`file.rs:1173-1187`), and directory-name parsing rejects
lengths outside 2..=64 or odd lengths before viewing UTF-16 pairs
(`file.rs:1326-1335`). The standalone UTF-16 decoder likewise rejects odd
input before `as_chunks::<2>()` (`file.rs:2885-2905`). The allocation-validation
fixture premise keeps its historical whole-entry scan and does not alter the
reader. There is no dropped remainder, new allocation, or changed error path.

The scalar core hex decoder still receives only the even-length, whitespace-
filtered string checked by `decode` (`crates/litchi-core/src/hex.rs:49-60`);
the fixed 64-byte `BlobId` parser checks exact ASCII length before taking 32
pairs (`crates/litchi-core/src/patch.rs:90-103`). DOCX `fixed_hex` performs the
same exact `N * 2` and ASCII-hexdigit checks before its array view
(`crates/litchi-docx/src/font/codec.rs:649-657`). Their nibble validation,
lowercase policy, result capacity, and failure ordering remain unchanged.

The OfficeArt changes retain the malformed-name checks before all views.
FBSE names reject odd lengths and out-of-record extents before the UTF-16
prefix view, and a nonempty name must already have a two-byte NUL terminator
(`crates/litchi-odraw/src/image/codec.rs:254-287`). `image::Name` is private to
that validated construction path, so `Name::to_string`'s `len - 2` slice remains
backed by the same terminator invariant (`image/model.rs:526-552`). Picture
property names reject empty, over-limit, and odd byte strings before building
the unit vector and require one final NUL before `Name::from_raw`
(`prop/picture/validation.rs:20-44, 73-92`); `Name::text` therefore has the
same checked invariant (`prop/picture/model.rs:137-158`). These edits preserve
the existing bounded `Vec` materialization in property validation and text
decoding; they do not turn any path into an infallible allocation.

The MD4 rewrite in `crates/litchi-odraw/src/image/write.rs:528-561` is
equivalent for every input length. `data.as_chunks::<64>()` processes each full
block and leaves the exact remainder; the zeroed 128-byte stack tail receives
the `0x80` byte and little-endian bit length at offset 56 or 120, and the 64/
128-byte padded slice is itself an exact 64-byte view. The digest and 64-byte
compression block are fixed-size arrays, so their 4-byte views are exact. It
uses the same stack storage and no heap allocation, preserves wrapping length
semantics, and does not change digest or writer error ordering. Existing
`md4_matches_rfc_vectors` and `writes_two_uid_alternate_jpeg_losslessly` cover
the digest and writer path; explicit 55/56-byte boundary vectors would be a
useful future regression, but their absence is not a source-level blocker
because the padding extent is proved by the fixed remainder bound.

The ODT, OGraph, RTF, and VBA conversions all retain their complete-pair
guards. ODT hex decoding checks even length before taking pairs
(`crates/litchi-odt/src/transaction.rs:5305-5319`). OGraph first checks the
checked count/width product and requires `data.len() == header + content`, so
the wide chart string view cannot have a remainder
(`crates/litchi-ograph/src/chart/codec/records/text.rs:55-105`). Its removed
`match ... _ => 0` arm was unreachable under that equality check; the
little-endian unit conversion, UTF-16 error, output reservation, and error
order are unchanged. RTF PANOSE checks exactly 20 ASCII hexadecimal bytes
before zipping ten pairs (`crates/litchi-rtf/src/codec/parser/codec/resources.rs:641-664`),
and writereservhash checks nonempty, even, ASCII-hex input before decoding
(`codec/styles.rs:742-766`). Picture-payload and tail-append hex decoders
perform their odd-length, size-limit, and fallible-reservation checks before
the pair loops (`edit/picture_payload.rs:1268-1290`; `tail_append.rs:5079-5109`).
VBA UTF-16 decoding retains its even-length check, exact string capacity,
surrogate rejection, and NUL rejection (`crates/litchi-vba/src/dir.rs:768-787`).
The old RTF PANOSE slice-pattern error and OGraph short-pair fallback were
unreachable after their preceding exact-width checks, so removing them does
not weaken malformed-input handling.

The two SOAPberry edits only replace `result.err().expect(...)` with
`result.expect_err(...)` and `assert_eq!(..., false)` with `assert!(!...)`
(`crates/soapberry-zip/src/office.rs:11778-11799`; `preserve.rs:3116-3130`).
They do not alter production code, preservation behavior, or error values.

These conclusions fit ADR 0005's bounded positional parsing and explicit
allocation/resource accounting, ADR 0006's fail-closed validation and
preservation contract, ADR 0024's ownership of CFB, ODraw, OGraph, DOCX, ODT,
RTF, and VBA codec boundaries, and ADR 0026's rule that shared OLE owners
reuse the validated CFB directory rather than duplicate or weaken its checks.

## Focused existing coverage

If the frozen source is rebased or a focused gate is useful, the highest-value
existing selectors are:

- CFB: `directory_name_decoder_checks_utf16_extents_without_panicking`,
  `opens_version_3_files_with_an_uninitialized_stream_size_high_word`,
  `opens_files_whose_length_is_not_a_whole_number_of_sectors`, the
  `reusable_chain_scratch_*` tests, `fat_stream_chain_errors_remain_typed_and_ordered`,
  `minifat_stream_chain_errors_remain_typed_and_ordered`, and the real-world
  `opens_the_complete_legacy_office_compound_file_corpus` gate. These exercise
  directory/name boundaries, fixed-sector decoding, and error ordering.
- ODraw: `image::write::tests::md4_matches_rfc_vectors`,
  `image::write::tests::writes_two_uid_alternate_jpeg_losslessly`, and the
  picture tests `decodes_picture_name_and_retains_reserved_flags`,
  `validates_utf16_bounds_and_flag_dependencies`, and
  `edits_only_modeled_picture_values_and_preserves_opaque_neighbors`.
- OGraph and RTF: the OGraph `framing` tests
  `malformed_frames_remain_typed_biff_errors_at_their_wire_offset` and
  `biff_record_bounds_are_enforced_without_a_second_ograph_frame_policy`,
  together with RTF `rejects_malformed_extended_font_metadata`,
  `rejects_invalid_hash_and_active_or_oversized_payloads`, and
  `selected_payload_splice_preserves_hex_layout_metadata_and_other_bytes`.
- Core and format callers: the existing core hex/patch parser suites,
  DOCX `package_delegates_fonts_to_the_canonical_docx_owner`, ODT packaged
  transaction tests such as
  `packaged_transaction_is_source_checked_reversible_and_exact_for_noop`,
  and the VBA DIR/project fixture suite. These retain the caller-level
  malformed-input, source-checking, and round-trip coverage around the small
  array-view conversions.

No additional production change is required for this mechanical batch. The
only worthwhile follow-up is an explicit MD4 padding-boundary vector around
55/56 and 63/64 bytes if maintainers want direct regression coverage; it is a
test-strengthening suggestion, not a discovered defect.

## Follow-up source review

The small post-review changes also contain no blocker. In
`crates/litchi-docx/src/source_backed/story_text.rs:2246-2312`,
`decode_attribute_value` now assigns the named-entity or numeric replacement
through one expression. The numeric branch still encodes into the same local
four-byte buffer, and the resulting slice is consumed before the branch's
buffer goes out of scope. The checked decoded-length pass, UTF-8/entity error
checks, per-replacement fallible reservation, and append order are unchanged;
there is no added allocation or changed limit behavior.

The RTF `logical_tail_append` test helper's
`std::iter::once(b'"').chain(field.iter().copied())` constructs the same byte
marker as the prior one-element-array iterator. It only changes test setup;
the JSON mutation search and delimiter checks are unchanged. The five CFB
utility views in `tools/perf-baseline/src/lib.rs:25431-25556` retain complete
input boundaries: the directory is assembled from full 512/4096-byte sectors,
each directory entry is 128 bytes, directory names are checked to be 2..=64
and even before the UTF-16 view, and a Unicode BoundSheet slice is obtained at
exactly `name_length * 2` bytes. The helper does not silently lose a partial
entry or pair, and its error/selection behavior is unchanged.

The ODF detector test changes replace `matches!(..., None)` with
`...expect(...).is_none()` while retaining all cursor-position, snapshot-read,
and index-build assertions (`crates/litchi-odf-common/src/detect.rs:1503-2095`).
The font tests use struct update syntax with `..Default::default()` for the
same limits and preserve `mut` only where a later boundary is changed
(`crates/litchi-fonts/src/embedding/mod.rs:282-316` and
`embedding/powerpoint.rs:889-1090`). The image-conversion tests make the same
mechanical substitution for `DibLimits`, `DeviceContext`, and the WMF fixture;
the explicitly assigned fields and all default fields are unchanged
(`crates/litchi-imgconv/src/dib/mod.rs:1261-1288`,
`emf/svg/converter.rs:2469-2490`, `emf/svg/state.rs:619-626`, and
`wmf/parser.rs:659-674`).

For completeness, the adjacent image-conversion array views retain their local
width proofs: EMF Bezier curves check a multiple-of-three remainder, Unicode
text and EMF+ text/glyph paths derive even byte lengths by checked multiplication
before slicing, the fixed WMF checksum covers a fixed 20-byte prefix, and the
rasterizer retains the prior behavior of ignoring a non-four-byte pixel tail.
No test initializer or these helper conversions changes production ownership,
limits, or error ordering.
