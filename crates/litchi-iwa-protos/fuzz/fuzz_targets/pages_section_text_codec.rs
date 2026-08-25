#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::ops::Range;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::{tsp::Reference, tswp::StorageArchive};
use litchi_iwa_text_wire::{
    RewriteBehavior, RewriteLimits, StorageRewriteExecutionLimits, decode_storage_with_limits,
    prepare_storage_text_rewrite_with_behavior_and_limits, validate_storage_with_limits,
};
use prost::Message as _;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_NESTING: usize = 8;
const MAX_FRAGMENTS: usize = 2 * 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_TABLE_ENTRIES: usize = 8 * 1024;
const MAX_REFERENCES: usize = 8 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_REWRITE_WORK: usize = 512 * 1024;

fuzz_target!(|data: &[u8]| {
    if let Some(source) = normalize_input(data) {
        exercise_raw(&source);
    }

    let source = generated_storage(data);
    exercise_valid_storage(&source, data);
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

fn limits() -> RewriteLimits {
    RewriteLimits::new(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_NESTING,
        MAX_FRAGMENTS,
        MAX_TEXT_BYTES,
        MAX_TABLE_ENTRIES,
        MAX_REFERENCES,
        MAX_OUTPUT_BYTES,
        MAX_REWRITE_WORK,
    )
    .unwrap_or_else(|error| unreachable!("valid section-text wire limits: {error}"))
}

fn exercise_raw(source: &[u8]) {
    let before = source.to_vec();
    observe(validate_storage_with_limits(source, limits()));
    assert_eq!(source, before.as_slice(), "validation modified its source");
    observe(decode_storage_with_limits(source, limits()));
    assert_eq!(source, before.as_slice(), "decode modified its source");

    let replacement = replacement(source);
    let range = 0..usize::from(source.first().copied().unwrap_or_default() & 7);
    observe(prepare_storage_text_rewrite_with_behavior_and_limits(
        source,
        range,
        &replacement,
        RewriteBehavior::PreserveOnEqualText,
        limits(),
    ));
    assert_eq!(source, before.as_slice(), "planning modified its source");
}

fn exercise_valid_storage(source: &[u8], data: &[u8]) {
    let validation = match validate_storage_with_limits(source, limits()) {
        Ok(validation) => validation,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    black_box((
        validation.storage_kind(),
        validation.utf8_len(),
        validation.utf16_len(),
        validation.fragments(),
        validation.fields(),
        validation.table_entries(),
        validation.reference_occurrences(),
        validation.validation_work(),
        validation.has_unknown_wire_fields(),
    ));

    let decoded = match decode_storage_with_limits(source, limits()) {
        Ok(decoded) => decoded,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(decoded.validation(), validation);
    let text = decoded.storage().text().to_owned();
    let range = command_range(&text, data);
    let selected = utf16_slice(&text, range.clone()).unwrap_or_default();
    let replacement = replacement(data);
    let mode = data.first().copied().unwrap_or_default() % 4;
    let (range, replacement) = match mode {
        // An empty insertion with empty replacement is an exact no-op.
        0 => (0..0, String::new()),
        // A normal Unicode replacement exercises UTF-16 offsets and table
        // shifts without relying on arbitrary bytes being valid UTF-8.
        1 => (range, replacement),
        // Clear the complete storage, then let the inverse insert it again.
        2 => (0..text.encode_utf16().count(), String::new()),
        // Equal selected text follows the source-preserving no-op behavior.
        _ => (range, selected.to_owned()),
    };
    exercise_rewrite(source, &text, range.clone(), &replacement, data);
    exercise_limit_rejection(source, range, &replacement);
}

fn exercise_rewrite(
    source: &[u8],
    text: &str,
    range: Range<usize>,
    replacement: &str,
    data: &[u8],
) {
    let before = source.to_vec();
    let behavior = if data.get(1).copied().unwrap_or_default() & 1 == 0 {
        RewriteBehavior::PreserveOnEqualText
    } else {
        RewriteBehavior::ReplaceSelection
    };
    let prepared = match prepare_storage_text_rewrite_with_behavior_and_limits(
        source,
        range.clone(),
        replacement,
        behavior,
        limits(),
    ) {
        Ok(prepared) => prepared,
        Err(error) => {
            observe_error(error);
            assert_eq!(source, before.as_slice(), "failed planning changed source");
            return;
        },
    };
    let requirements = prepared.execution_requirements();
    let output = match prepared.execute(StorageRewriteExecutionLimits {
        max_output_bytes: requirements.output_bytes(),
        max_retained_elements: requirements.retained_elements(),
        max_retained_bytes: requirements.retained_bytes(),
        max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
        max_allocations: requirements.allocations(),
        max_work: requirements.work(),
    }) {
        Ok(output) => output,
        Err(error) => {
            observe_error(error);
            assert_eq!(source, before.as_slice(), "failed execution changed source");
            return;
        },
    };
    assert_eq!(source, before.as_slice(), "rewrite changed its source");
    assert_eq!(output.changed(), output.bytes() != source);
    if !output.changed() {
        assert_eq!(output.bytes(), source);
    }
    black_box(output.execution_report());

    if !output.changed() {
        return;
    }

    // Replay the exact inverse using the semantic text selected before the
    // forward edit. This exercises source-preserving unknown fields and every
    // positional table that the strict text wire planner updates.
    let replacement_units = replacement.encode_utf16().count();
    let inverse_range = range.start..range.start.saturating_add(replacement_units);
    let inverse_text = utf16_slice(text, range).unwrap_or_default();
    let inverse = match prepare_storage_text_rewrite_with_behavior_and_limits(
        output.bytes(),
        inverse_range,
        inverse_text,
        behavior,
        limits(),
    ) {
        Ok(inverse) => inverse,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let requirements = inverse.execution_requirements();
    match inverse.execute(StorageRewriteExecutionLimits {
        max_output_bytes: requirements.output_bytes(),
        max_retained_elements: requirements.retained_elements(),
        max_retained_bytes: requirements.retained_bytes(),
        max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
        max_allocations: requirements.allocations(),
        max_work: requirements.work(),
    }) {
        Ok(restored) => {
            assert_eq!(
                restored.bytes(),
                source,
                "text rewrite inverse was not exact"
            );
            black_box(restored.execution_report());
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_limit_rejection(source: &[u8], range: Range<usize>, replacement: &str) {
    let Ok(prepared) = prepare_storage_text_rewrite_with_behavior_and_limits(
        source,
        range,
        replacement,
        RewriteBehavior::PreserveOnEqualText,
        limits(),
    ) else {
        return;
    };
    let requirements = prepared.execution_requirements();
    let mut constrained = StorageRewriteExecutionLimits {
        max_output_bytes: requirements.output_bytes().saturating_sub(1),
        max_retained_elements: requirements.retained_elements(),
        max_retained_bytes: requirements.retained_bytes(),
        max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
        max_allocations: requirements.allocations(),
        max_work: requirements.work(),
    };
    observe(prepared.execute(constrained));

    let Ok(prepared) = prepare_storage_text_rewrite_with_behavior_and_limits(
        source,
        0..0,
        "",
        RewriteBehavior::PreserveOnEqualText,
        limits(),
    ) else {
        return;
    };
    let requirements = prepared.execution_requirements();
    constrained.max_output_bytes = requirements.output_bytes();
    constrained.max_retained_elements = requirements.retained_elements().saturating_sub(1);
    observe(prepared.execute(constrained));
}

fn generated_storage(data: &[u8]) -> Vec<u8> {
    let suffix = data
        .iter()
        .copied()
        .take(32)
        .map(|byte| char::from(b'a' + byte % 26))
        .collect::<String>();
    let first = format!("A😀{suffix}");
    let second = format!("東京 section {suffix}");
    let boundary = u32::try_from(first.encode_utf16().count())
        .unwrap_or_else(|error| unreachable!("bounded generated text length: {error}"));
    let storage = StorageArchive {
        text: vec![first, second],
        table_section: Some(litchi_iwa_protos::tswp::ObjectAttributeTable {
            entries: vec![
                litchi_iwa_protos::tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: None,
                },
                litchi_iwa_protos::tswp::object_attribute_table::ObjectAttribute {
                    character_index: boundary,
                    object: Some(Reference {
                        identifier: 101,
                        ..Reference::default()
                    }),
                },
            ],
        }),
        ..StorageArchive::default()
    };
    let mut source = storage.encode_to_vec();
    append_varint_field(
        &mut source,
        99,
        u64::from(data.first().copied().unwrap_or_default()),
    );
    append_key(&mut source, 100, 3);
    append_varint_field(&mut source, 99, 7);
    append_key(&mut source, 100, 4);
    source
}

fn command_range(text: &str, data: &[u8]) -> Range<usize> {
    let length = text.encode_utf16().count();
    let start = usize::from(data.get(2).copied().unwrap_or_default()) % length.saturating_add(1);
    let end = usize::from(data.get(3).copied().unwrap_or_default()) % length.saturating_add(1);
    if start <= end { start..end } else { end..start }
}

fn utf16_slice(text: &str, range: Range<usize>) -> Option<&str> {
    let start = utf16_byte_index(text, range.start)?;
    let end = utf16_byte_index(text, range.end)?;
    text.get(start..end)
}

fn utf16_byte_index(text: &str, target: usize) -> Option<usize> {
    if target == 0 {
        return Some(0);
    }
    let mut units = 0usize;
    for (byte_index, character) in text.char_indices() {
        if units == target {
            return Some(byte_index);
        }
        units = units.checked_add(character.len_utf16())?;
        if units > target {
            return None;
        }
    }
    (units == target).then_some(text.len())
}

fn replacement(data: &[u8]) -> String {
    let suffix = data
        .iter()
        .copied()
        .skip(4)
        .take(48)
        .map(|byte| char::from(b'a' + byte % 26))
        .collect::<String>();
    format!("fuzz 😀 {suffix} 東京")
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_key(output, number, 0);
    append_varint(output, value);
}

fn append_key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    append_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn observe<T>(result: Result<T, impl Debug + Display>) {
    if let Err(error) = result {
        black_box((error.to_string(), format!("{error:?}")));
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box((error.to_string(), format!("{error:?}")));
}
