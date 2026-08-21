#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::pages_section_codec::{
    DecodeError, DecodeOptions, SectionReferenceSnapshot, decode_pagination,
    decode_section_settings, decode_section_settings_with_report,
};
use litchi_iwa_protos::tp::SectionArchive;
use litchi_iwa_protos::tsp::Reference;
use prost::Message as _;

// Keep libFuzzer's source and both strict/Buffa passes finite. Inputs are
// skipped rather than truncated so every entry point always sees one
// unchanged caller-owned source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_NAME_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;

#[derive(Clone, Copy)]
struct SourceRange {
    start: usize,
    end: usize,
}

impl SourceRange {
    fn new(source: &[u8]) -> Self {
        let start = source.as_ptr() as usize;
        let end = start
            .checked_add(source.len())
            .expect("bounded source pointer range");
        Self { start, end }
    }

    fn assert_borrowed(self, bytes: &str) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start
            .checked_add(bytes.len())
            .expect("bounded borrowed-name pointer range");
        assert!(
            start >= self.start && end <= self.end,
            "decoded section name did not borrow from the source"
        );
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise(&source);
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn options(max_message_bytes: usize, recursion_limit: u32) -> DecodeOptions {
    DecodeOptions::new(max_message_bytes, recursion_limit)
        .with_max_fields(MAX_FIELDS)
        .with_max_work_bytes(MAX_WORK_BYTES)
        .with_max_name_bytes(MAX_NAME_BYTES)
}

fn exercise(source: &[u8]) {
    let before = source.to_vec();
    let source_range = SourceRange::new(source);

    let pagination = decode_pagination(source, options(MAX_INPUT_BYTES, MAX_RECURSION));
    assert_eq!(
        source,
        before.as_slice(),
        "pagination decode modified source"
    );
    match pagination {
        Ok(snapshot) => {
            black_box(snapshot);
            // Pagination intentionally leaves the other selected fields
            // opaque. It can therefore succeed for a payload rejected by the
            // stricter aggregate projection (for example invalid field 26
            // UTF-8), but its selected scalars must still match Prost.
            let archive = SectionArchive::decode(source).unwrap_or_else(|error| {
                panic!("pagination acceptance disagreed with Prost: {error}")
            });
            assert_eq!(snapshot.section_start_kind, archive.section_start_kind);
            assert_eq!(
                snapshot.section_page_number_kind,
                archive.section_page_number_kind
            );
            assert_eq!(
                snapshot.section_page_number_start,
                archive.section_page_number_start
            );
        },
        Err(error) => observe_error(error),
    }

    let aggregate =
        decode_section_settings_with_report(source, options(MAX_INPUT_BYTES, MAX_RECURSION));
    assert_eq!(
        source,
        before.as_slice(),
        "aggregate decode modified source"
    );

    let (aggregate_ok, has_nonempty_name) = match aggregate {
        Ok((snapshot, report)) => {
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            assert!(report.name_bytes() <= MAX_NAME_BYTES);
            if let Some(name) = snapshot.name() {
                source_range.assert_borrowed(name);
            }
            let has_nonempty_name = snapshot.name().is_some_and(|name| !name.is_empty());
            let archive = SectionArchive::decode(source).unwrap_or_else(|error| {
                panic!("aggregate acceptance disagreed with Prost: {error}")
            });
            assert_snapshot_matches(snapshot, &archive);
            black_box((snapshot, report));
            (true, has_nonempty_name)
        },
        Err(error) => {
            observe_error(error);
            (false, false)
        },
    };

    let scalar = decode_section_settings(source, options(MAX_INPUT_BYTES, MAX_RECURSION));
    assert_eq!(
        source,
        before.as_slice(),
        "aggregate scalar decode modified source"
    );
    match (aggregate_ok, scalar) {
        (true, Ok(snapshot)) => {
            if let Some(name) = snapshot.name() {
                source_range.assert_borrowed(name);
            }
            black_box(snapshot);
        },
        (true, Err(error)) => panic!("report and scalar aggregate paths disagreed: {error}"),
        (false, Ok(snapshot)) => {
            panic!("report and scalar aggregate paths disagreed: {snapshot:?}")
        },
        (false, Err(error)) => observe_error(error),
    }

    exercise_limit_profiles(source, aggregate_ok, has_nonempty_name);
    assert_eq!(source, before.as_slice(), "limit probes modified source");
}

fn exercise_limit_profiles(source: &[u8], aggregate_ok: bool, has_nonempty_name: bool) {
    // Validation rejects these profiles before touching source bytes, while
    // still proving the public error is typed as a finite resource failure.
    expect_limit(
        decode_pagination(source, options(usize::MAX, 1)),
        "configured byte ceiling",
    );
    expect_limit(
        decode_section_settings(source, options(usize::MAX, 1)),
        "configured byte ceiling",
    );
    expect_limit(
        decode_pagination(source, options(MAX_INPUT_BYTES, 0)),
        "zero recursion ceiling",
    );
    expect_limit(
        decode_section_settings(source, options(MAX_INPUT_BYTES, MAX_RECURSION + 1)),
        "excessive recursion ceiling",
    );

    if !aggregate_ok || source.is_empty() {
        return;
    }

    // A valid, non-empty source reaches the first strict field under a zero
    // field budget. Invalid sources are intentionally observed rather than
    // asserted here: they may fail on wire shape before the selected budget.
    let field_limited = options(MAX_INPUT_BYTES, MAX_RECURSION).with_max_fields(0);
    observe_optional_limit(
        decode_section_settings(source, field_limited),
        "field ceiling",
    );

    // A successful aggregate source performs two full-source work charges;
    // one byte below that exact amount must reject on the second pass.
    let work_limit = source.len().saturating_mul(2).saturating_sub(1);
    let work_limited = options(MAX_INPUT_BYTES, MAX_RECURSION).with_max_work_bytes(work_limit);
    observe_optional_limit(
        decode_section_settings(source, work_limited),
        "work ceiling",
    );

    if has_nonempty_name {
        let name_limited = options(MAX_INPUT_BYTES, MAX_RECURSION).with_max_name_bytes(0);
        observe_optional_limit(
            decode_section_settings(source, name_limited),
            "name-byte ceiling",
        );
    }

    let bytes_limited = options(source.len() - 1, MAX_RECURSION);
    expect_limit(
        decode_section_settings(source, bytes_limited),
        "input byte ceiling",
    );
}

fn assert_snapshot_matches(
    snapshot: litchi_iwa_protos::pages_section_codec::SectionSettingsSnapshot<'_>,
    archive: &SectionArchive,
) {
    assert_eq!(
        snapshot.inherit_previous_header_footer(),
        archive.inherit_previous_header_footer
    );
    assert_eq!(
        snapshot.section_template_first_page_different(),
        archive.section_template_first_page_different
    );
    assert_eq!(
        snapshot.section_template_even_odd_pages_different(),
        archive.section_template_even_odd_pages_different
    );
    assert_eq!(snapshot.section_start_kind(), archive.section_start_kind);
    assert_eq!(
        snapshot.section_page_number_kind(),
        archive.section_page_number_kind
    );
    assert_eq!(
        snapshot.section_page_number_start(),
        archive.section_page_number_start
    );
    assert_reference_matches(
        snapshot.first_section_template_page(),
        archive.first_section_template_page.as_ref(),
    );
    assert_reference_matches(
        snapshot.even_section_template_page(),
        archive.even_section_template_page.as_ref(),
    );
    assert_reference_matches(
        snapshot.odd_section_template_page(),
        archive.odd_section_template_page.as_ref(),
    );
    assert_eq!(snapshot.name(), archive.name.as_deref());
    assert_eq!(
        snapshot.section_template_first_page_hides_header_footer(),
        archive.section_template_first_page_hides_header_footer
    );
    assert_reference_matches(
        snapshot.user_defined_guide_storage(),
        archive.user_defined_guide_storage.as_ref(),
    );
}

fn assert_reference_matches(
    snapshot: Option<SectionReferenceSnapshot>,
    reference: Option<&Reference>,
) {
    match (snapshot, reference) {
        (None, None) => {},
        (Some(snapshot), Some(reference)) => {
            assert_eq!(snapshot.identifier().get(), reference.identifier);
            assert_eq!(snapshot.deprecated_type(), reference.deprecated_type);
            assert_eq!(
                snapshot.deprecated_is_external(),
                reference.deprecated_is_external
            );
        },
        _ => panic!("strict and Prost reference presence differed"),
    }
}

fn expect_limit<T>(result: Result<T, DecodeError>, context: &str) {
    let Err(error) = result else {
        panic!("{context} unexpectedly accepted");
    };
    assert!(
        error.resource_limit().is_some(),
        "{context} returned a non-resource error: {error}"
    );
    observe_error(error);
}

fn observe_optional_limit<T>(result: Result<T, DecodeError>, context: &str) {
    if let Err(error) = result {
        if error.resource_limit().is_none() {
            black_box(context);
        }
        observe_error(error);
    }
}

fn observe_error(error: DecodeError) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
    black_box(error.resource_limit());
}
