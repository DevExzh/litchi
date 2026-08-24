#![no_main]

use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_chart_caption_codec::{
    ChartCaptionWrite, DecodeError, DecodeOptions, decode_chart_caption,
    decode_chart_caption_identifier, rewrite_chart_caption, rewrite_chart_caption_with_report,
};

// Keep every strict pass and every candidate output finite. Oversized inputs
// are skipped rather than truncated, so the codec always sees one unchanged
// caller-owned source slice.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_RECURSION: u32 = 64;

// These recipes deliberately include valid unknown groups and overlong values
// on unknown fields as well as malformed selected fields. They are kept here,
// rather than borrowing another target's corpus, so chart-caption campaigns
// remain independent of the existing Pages footnote corpus.
const FIXED_CASES: &[&[u8]] = &[
    // Canonical super -> drawable -> caption -> reference(identifier=7).
    &[0x0a, 0x04, 0x5a, 0x02, 0x08, 0x07],
    // Unknown root scalar with a non-canonical value, before the selected edge.
    &[0x98, 0x06, 0x81, 0x00, 0x0a, 0x04, 0x5a, 0x02, 0x08, 0x07],
    // Unknown group and unknown overlong scalar nested in the group.
    &[
        0x9b, 0x03, 0xa0, 0x03, 0x81, 0x00, 0x9c, 0x03, 0x0a, 0x04, 0x5a, 0x02, 0x08, 0x07,
    ],
    // Known optional reference fields plus an unknown group, all valid.
    &[
        0x0a, 0x0c, 0x5a, 0x0a, 0x10, 0x01, 0x08, 0x07, 0x18, 0x01, 0x9b, 0x03, 0x9c, 0x03,
    ],
    // The same selected edge with nested unknown group and overlong scalar.
    &[
        0x0a, 0x10, 0x5a, 0x0e, 0x98, 0x06, 0x81, 0x00, 0x10, 0x01, 0x08, 0x07, 0x18, 0x01, 0x9b,
        0x03, 0x9c, 0x03,
    ],
    // Missing required identifier.
    &[0x0a, 0x02, 0x5a, 0x00],
    // Wrong wire type for the selected caption edge.
    &[0x0a, 0x02, 0x58, 0x01],
    // Duplicate selected super envelope.
    &[0x0a, 0x00, 0x0a, 0x00],
    // Non-canonical value on the selected deprecated type field.
    &[0x0a, 0x07, 0x5a, 0x05, 0x08, 0x07, 0x10, 0x81, 0x00],
    // Unterminated unknown group.
    &[0x9b, 0x03, 0x08, 0x01, 0x0a, 0x04, 0x5a, 0x02, 0x08, 0x07],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_source(&source, data);

    // Mutating a bounded copy gives the campaign a second path without ever
    // allowing a codec operation to retain or modify libFuzzer's input.
    if !source.is_empty() {
        let mut mutated = source.clone();
        let index = usize::from(data.first().copied().unwrap_or_default()) % mutated.len();
        mutated[index] ^= 0xff;
        exercise_source(&mutated, data);
    }

    static GUARDS: OnceLock<()> = OnceLock::new();
    GUARDS.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-chart-caption");
        }
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
    let decoded = decode_chart_caption(source, decode_options);
    if source != before.as_slice() {
        return;
    }

    let Ok(snapshot) = decoded else {
        if let Err(error) = decoded {
            observe_error(error);
        }
        return;
    };

    let identifier = snapshot.caption_identifier();
    black_box((snapshot.has_drawable(), identifier));
    if decode_chart_caption_identifier(source, decode_options) != Ok(identifier) {
        return;
    }
    let Some(identifier) = identifier else {
        return;
    };

    // First exercise the exact no-op path, then choose a wide replacement so
    // both same-width and length-prefix-changing rewrites are reachable.
    exercise_rewrite(source, identifier, decode_options, &before);
    let replacement = replacement_identifier(data, identifier);
    exercise_rewrite(source, replacement, decode_options, &before);
}

fn replacement_identifier(data: &[u8], original: u64) -> u64 {
    let candidate = 0x1_0000_0000_u64
        | u64::from(data.first().copied().unwrap_or_default())
        | (u64::from(data.get(1).copied().unwrap_or_default()) << 8)
        | (u64::from(data.get(2).copied().unwrap_or_default()) << 16);
    if candidate == original {
        original.wrapping_add(1)
    } else {
        candidate
    }
}

fn exercise_rewrite(source: &[u8], identifier: u64, decode_options: DecodeOptions, before: &[u8]) {
    let result = rewrite_chart_caption_with_report(
        source,
        ChartCaptionWrite::new(identifier),
        decode_options,
    );
    if source != before {
        return;
    }

    let Ok((output, report)) = result else {
        if let Err(error) = result {
            observe_error(error);
        }
        return;
    };

    if report.input_bytes() != source.len()
        || report.output_bytes() != output.len()
        || output.len() > MAX_OUTPUT_BYTES.max(source.len())
    {
        return;
    }
    let readback_options = options(&output);
    match decode_chart_caption(&output, readback_options) {
        Ok(snapshot) if snapshot.caption_identifier() == Some(identifier) => {
            black_box(report);
        },
        Ok(_) | Err(_) => return,
    }
    if Some(identifier)
        == decode_chart_caption_identifier(source, options(source))
            .ok()
            .flatten()
    {
        if output != source {
            return;
        }
    }

    // Keep the convenience alias on the same strict path for successful
    // rewrites. A refusal is valid under the finite profile and is observed.
    match rewrite_chart_caption(source, ChartCaptionWrite::new(identifier), decode_options) {
        Ok(alias) if alias == output => {},
        Ok(_) | Err(_) => {},
    }
}

fn exercise_limit_guards() {
    let source = FIXED_CASES[0];
    let finite = options(source);

    let _ = decode_chart_caption(
        source,
        DecodeOptions::new(
            source.len().saturating_sub(1),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
        ),
    )
    .map_err(observe_error);
    let _ = decode_chart_caption(
        source,
        DecodeOptions::new(source.len(), 1, MAX_WORK_BYTES, MAX_RECURSION),
    )
    .map_err(observe_error);
    let _ = decode_chart_caption(
        source,
        DecodeOptions::new(source.len(), MAX_FIELDS, 1, MAX_RECURSION),
    )
    .map_err(observe_error);
    let _ = decode_chart_caption(
        source,
        DecodeOptions::new(source.len(), MAX_FIELDS, MAX_WORK_BYTES, 0),
    )
    .map_err(observe_error);

    let _ = rewrite_chart_caption_with_report(
        source,
        ChartCaptionWrite::new(0x1_0000_0000),
        finite.with_max_output_bytes(source.len().saturating_sub(1)),
    )
    .map_err(|error| observe_error(error));
}

fn observe_error(error: DecodeError) {
    // Error formatting is intentionally source-free; black_box keeps all
    // typed failure paths live without retaining attacker-controlled bytes.
    black_box(error.to_string());
}
