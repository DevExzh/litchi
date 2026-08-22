#![no_main]

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::pages_movie_caption_codec::{
    CaptionInfoWrite, DecodeError, DecodeOptions, decode_caption_info,
    rewrite_caption_info, rewrite_caption_info_with_report,
};

// Keep the source and every strict/rewrite pass within one finite profile.
// Inputs are skipped rather than truncated, so all checks see the same
// caller-owned byte slice.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_RECURSION: u32 = 64;

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);

    // Keep fixed malformed and limit recipes independent of arbitrary bytes.
    // This makes each failure mode reachable even when the campaign corpus is
    // still finding the required inheritance chain.
    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(|| {
        exercise_known_malformed();
        exercise_limit_guards();
    });
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

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    )
    .with_max_output_bytes(MAX_OUTPUT_BYTES.max(source.len()))
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let decoded = decode_caption_info(source, decode_options);
    assert_eq!(source, before.as_slice(), "decode modified its source");

    let Ok(snapshot) = decoded else {
        if let Err(error) = decoded {
            observe_error(error);
        }
        return;
    };

    black_box(snapshot.parent_identifier());
    if let Some(kind) = snapshot.child_info_kind() {
        black_box(kind);
    }
    black_box((
        snapshot.deprecated_storage_identifier(),
        snapshot.owned_storage_identifier(),
        snapshot.is_text_box(),
        snapshot.style_identifier(),
        snapshot.placement_identifier(),
    ));

    // A matching map must be an exact byte-preserving no-op, including all
    // unknown fields and the source's selected-field ordering.
    let parent = snapshot.parent_identifier();
    let no_op = [(parent, parent)];
    let no_op_succeeded = exercise_rewrite(source, &no_op, snapshot, decode_options, false);

    // Rewrite one required edge to a wide canonical varint. The callback-free
    // codec must preserve every other selected and unknown span, and the
    // source must remain unchanged on both success and refusal.
    let replacement = replacement_identifier(data, parent);
    let remap = [(parent, replacement)];
    exercise_rewrite(source, &remap, snapshot, decode_options, replacement != parent);
    assert_eq!(source, before.as_slice(), "rewrite modified its source");

    // Keep the convenience alias on the same semantic path as the reporting
    // entry point for successful no-op writes.
    if no_op_succeeded {
        let alias = rewrite_caption_info(source, CaptionInfoWrite::new(&no_op), decode_options)
            .unwrap_or_else(|error| panic!("no-op rewrite alias failed: {error}"));
        assert_eq!(alias, source);
    }
}

fn replacement_identifier(data: &[u8], original: u64) -> u64 {
    let mut replacement = 0x1_0000_0000_u64
        | u64::from(data.first().copied().unwrap_or_default())
        | (u64::from(data.get(1).copied().unwrap_or_default()) << 8);
    if replacement == original {
        replacement = original.wrapping_add(1);
    }
    replacement
}

fn exercise_rewrite(
    source: &[u8],
    remap: &[(u64, u64)],
    snapshot: litchi_iwa_protos::pages_movie_caption_codec::CaptionInfoSnapshot,
    decode_options: DecodeOptions,
    expected_change: bool,
) -> bool {
    let before = source.to_vec();
    let result = rewrite_caption_info_with_report(
        source,
        CaptionInfoWrite::new(remap),
        decode_options,
    );
    assert_eq!(source, before.as_slice(), "rewrite modified its source");

    let Ok((output, report)) = result else {
        if let Err(error) = result {
            // Near finite source ceilings a changed rewrite may correctly
            // refuse its aggregate work or candidate-output budget.
            observe_error(error);
        }
        return false;
    };

    assert_eq!(report.input_bytes(), source.len());
    assert_eq!(report.output_bytes(), output.len());
    assert_eq!(report.references_rewritten() > 0, expected_change);
    assert!(output.len() <= MAX_OUTPUT_BYTES.max(source.len()));

    let readback = decode_caption_info(&output, options(&output))
        .unwrap_or_else(|error| panic!("caption-info rewrite readback failed: {error}"));
    if expected_change {
        assert_ne!(readback.parent_identifier(), snapshot.parent_identifier());
    } else {
        assert_eq!(output, source);
        assert_eq!(readback, snapshot);
    }
    black_box(report);
    true
}

fn observe_error(error: DecodeError) {
    // Error text deliberately contains no source bytes; format it to keep all
    // typed failure paths exercised without retaining attacker-controlled data.
    black_box(error.to_string());
}

fn exercise_known_malformed() {
    let recipes: &[&[u8]] = &[
        &[],
        &[0x0a],
        &[0x0a, 0x00],
        // Wrong wire type for childInfoKind.
        &[0x1a, 0x01, 0x01],
        // Non-canonical bool scalar in the selected shape-info envelope.
        &[
            0x0a, 0x22, 0x0a, 0x10, 0x0a, 0x07, 0x12, 0x05, 0x08, 0x0b, 0xd0, 0x05,
            0x07, 0x12, 0x05, 0x08, 0x16, 0xd0, 0x05, 0x07, 0x12, 0x05, 0x08, 0x21,
            0xd0, 0x05, 0x07, 0x22, 0x05, 0x08, 0x2c, 0xd0, 0x05, 0x07, 0x30, 0x02,
            0x12, 0x05, 0x08, 0x37, 0xd0, 0x05, 0x07, 0x18, 0x01,
        ],
        // Duplicate singular childInfoKind.
        &[
            0x0a, 0x22, 0x0a, 0x10, 0x0a, 0x07, 0x12, 0x05, 0x08, 0x0b, 0xd0, 0x05,
            0x07, 0x12, 0x05, 0x08, 0x16, 0xd0, 0x05, 0x07, 0x12, 0x05, 0x08, 0x21,
            0xd0, 0x05, 0x07, 0x22, 0x05, 0x08, 0x2c, 0xd0, 0x05, 0x07, 0x30, 0x01,
            0x12, 0x05, 0x08, 0x37, 0xd0, 0x05, 0x07, 0x18, 0x01, 0x18, 0x02,
        ],
    ];
    for recipe in recipes {
        let source = recipe.to_vec();
        let before = source.clone();
        assert!(decode_caption_info(&source, options(&source)).is_err());
        assert_eq!(source, before.as_slice(), "malformed decode modified source");
        let remap = [(1_u64, 2_u64)];
        let rewritten = rewrite_caption_info_with_report(
            &source,
            CaptionInfoWrite::new(&remap),
            options(&source),
        );
        assert!(rewritten.is_err());
        assert_eq!(source, before.as_slice(), "malformed rewrite modified source");
    }
}

fn exercise_limit_guards() {
    let source = full_caption_info();
    let bytes = DecodeOptions::new(
        source.len().saturating_sub(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
    );
    assert!(decode_caption_info(&source, bytes)
        .expect_err("byte limit was not enforced")
        .message_byte_limit_values()
        .is_some());

    let fields = DecodeOptions::new(source.len(), 1, MAX_WORK_BYTES, MAX_RECURSION);
    assert!(decode_caption_info(&source, fields)
        .expect_err("field limit was not enforced")
        .field_limit_values()
        .is_some());

    let work = DecodeOptions::new(source.len(), MAX_FIELDS, source.len(), MAX_RECURSION);
    assert!(decode_caption_info(&source, work)
        .expect_err("work limit was not enforced")
        .work_limit_values()
        .is_some());

    let nesting = DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, 0);
    assert!(decode_caption_info(&source, nesting)
        .expect_err("recursion limit was not enforced")
        .recursion_limit_values()
        .is_some());

    let output = DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, MAX_RECURSION)
        .with_max_output_bytes(1);
    let remap = [(11_u64, 1_u64 << 40)];
    let before = source.clone();
    let error = rewrite_caption_info_with_report(&source, CaptionInfoWrite::new(&remap), output)
        .expect_err("output limit was not enforced");
    assert!(error.output_limit_values().is_some());
    assert_eq!(source, before);
}

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return output;
        }
    }
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    [varint(u64::from(number) << 3), varint(value)].concat()
}

fn field_bytes(number: u32, value: &[u8]) -> Vec<u8> {
    [
        varint((u64::from(number) << 3) | 2),
        varint(value.len() as u64),
        value.to_vec(),
    ]
    .concat()
}

fn reference(identifier: u64) -> Vec<u8> {
    [field_varint(1, identifier), field_varint(90, 7)].concat()
}

fn full_caption_info() -> Vec<u8> {
    let drawable = field_bytes(2, &reference(11));
    let shape = field_bytes(1, &drawable) // ShapeInfoArchive.super -> ShapeArchive
        .into_iter()
        .chain(field_bytes(2, &reference(22)))
        .collect::<Vec<_>>();
    let shape_info = [
        field_bytes(1, &shape),
        field_bytes(2, &reference(33)),
        field_bytes(4, &reference(44)),
        field_varint(6, 1),
    ]
    .concat();
    [
        field_bytes(1, &shape_info),
        field_bytes(2, &reference(55)),
        field_varint(3, 1),
    ]
    .concat()
}
