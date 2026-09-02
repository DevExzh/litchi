//! Focused public-boundary coverage for the Numbers Fraction format codec.
//!
//! The wire payload is deliberately built without generated protobuf values.
//! Fraction owns only the native type and accuracy fields; unknown records
//! remain source-authoritative, while field 20 is accepted only in its
//! canonical, explicitly-false form.

use litchi_iwa_protos::numbers_table_cell_fraction_format_codec as codec;

fn options(source: &[u8]) -> codec::DecodeOptions {
    let bytes = source.len().max(64);
    codec::DecodeOptions::new(
        bytes.saturating_mul(2),
        bytes.saturating_mul(4),
        bytes.saturating_mul(16),
        bytes.saturating_mul(64),
        64,
        bytes.saturating_mul(2),
        bytes.saturating_mul(2),
        bytes.saturating_mul(2),
    )
}

fn push_varint_value(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = u8::try_from(value & 0x7f).expect("seven bits fit in a byte");
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn push_varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    push_varint_value(output, u64::from(field) << 3);
    push_varint_value(output, value);
}

fn fraction_payload(accuracy: u32) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint_field(
        &mut output,
        1,
        u64::from(codec::NATIVE_FRACTION_FORMAT_TYPE),
    );
    push_varint_field(&mut output, 11, u64::from(accuracy));
    output
}

fn with_replacement_flag(mut source: Vec<u8>, value: &[u8]) -> Vec<u8> {
    source.extend_from_slice(&[0xa0, 0x01]);
    source.extend_from_slice(value);
    source
}

fn contains_bytes(source: &[u8], needle: &[u8]) -> bool {
    source.windows(needle.len()).any(|window| window == needle)
}

#[test]
fn scalar_and_reported_reads_have_lazy_parity_for_all_signed_accuracies() {
    let accuracies = [
        (
            codec::FractionAccuracy::UpToOneDigit,
            codec::NATIVE_FRACTION_UP_TO_ONE_DIGIT,
        ),
        (
            codec::FractionAccuracy::UpToTwoDigits,
            codec::NATIVE_FRACTION_UP_TO_TWO_DIGITS,
        ),
        (
            codec::FractionAccuracy::UpToThreeDigits,
            codec::NATIVE_FRACTION_UP_TO_THREE_DIGITS,
        ),
        (
            codec::FractionAccuracy::Halves,
            codec::NATIVE_FRACTION_HALVES,
        ),
        (
            codec::FractionAccuracy::Quarters,
            codec::NATIVE_FRACTION_QUARTERS,
        ),
        (
            codec::FractionAccuracy::Eighths,
            codec::NATIVE_FRACTION_EIGHTHS,
        ),
        (
            codec::FractionAccuracy::Sixteenths,
            codec::NATIVE_FRACTION_SIXTEENTHS,
        ),
        (
            codec::FractionAccuracy::Tenths,
            codec::NATIVE_FRACTION_TENTHS,
        ),
        (
            codec::FractionAccuracy::Hundredths,
            codec::NATIVE_FRACTION_HUNDREDTHS,
        ),
    ];

    for (expected, native) in accuracies {
        let source = fraction_payload(native);
        let scalar = codec::decode_fraction_format(&source, options(&source)).expect("fraction");
        let (reported, report) =
            codec::decode_fraction_format_with_report(&source, options(&source))
                .expect("reported fraction");

        assert_eq!(scalar.raw(), source.as_slice());
        assert_eq!(scalar.raw().as_ptr(), source.as_ptr());
        assert_eq!(scalar.format_type(), codec::NATIVE_FRACTION_FORMAT_TYPE);
        assert_eq!(scalar.fraction_accuracy(), native);
        assert_eq!(scalar.accuracy(), Some(expected));
        assert_eq!(scalar.requires_fraction_replacement(), None);

        assert_eq!(reported.raw(), scalar.raw());
        assert_eq!(reported.format_type(), scalar.format_type());
        assert_eq!(reported.fraction_accuracy(), scalar.fraction_accuracy());
        assert_eq!(reported.accuracy(), scalar.accuracy());
        assert_eq!(reported.requires_fraction_replacement(), None);
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.fields(), 2);
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.allocations(), 0);
    }
}

#[test]
fn unknown_fraction_records_are_preserved_by_source_rewrite() {
    let mut source = fraction_payload(codec::NATIVE_FRACTION_UP_TO_THREE_DIGITS);
    let unknown_scalar = [0xf0, 0x02, 0x81, 0x00];
    let unknown_bytes = [0xfa, 0x02, 0x03, 0xca, 0xfe, 0xba];
    let unknown_group = [0x83, 0x03, 0x08, 0x07, 0x84, 0x03];
    source.extend_from_slice(&unknown_scalar);
    source.extend_from_slice(&unknown_bytes);
    source.extend_from_slice(&unknown_group);

    let before = source.clone();
    let output = codec::rewrite_fraction_format(
        &source,
        codec::FractionFormatWrite::from_accuracy(codec::FractionAccuracy::UpToTwoDigits),
        options(&source),
    )
    .expect("source-preserving rewrite");

    assert_eq!(source, before);
    assert!(contains_bytes(output.bytes(), &unknown_scalar));
    assert!(contains_bytes(output.bytes(), &unknown_bytes));
    assert!(contains_bytes(output.bytes(), &unknown_group));
    assert!(!contains_bytes(output.bytes(), &[0xa0, 0x01, 0x00]));

    let snapshot = codec::decode_fraction_format(output.bytes(), options(output.bytes()))
        .expect("rewritten fraction");
    assert_eq!(
        snapshot.accuracy(),
        Some(codec::FractionAccuracy::UpToTwoDigits)
    );
    assert_eq!(snapshot.requires_fraction_replacement(), None);
}

#[test]
fn replacement_false_is_preserved_but_not_synthesized_and_true_is_rejected() {
    let source = with_replacement_flag(
        fraction_payload(codec::NATIVE_FRACTION_UP_TO_THREE_DIGITS),
        &[0x00],
    );
    let snapshot = codec::decode_fraction_format(&source, options(&source))
        .expect("canonical false replacement flag");
    assert_eq!(snapshot.requires_fraction_replacement(), Some(false));

    let output = codec::rewrite_fraction_format(
        &source,
        codec::FractionFormatWrite::from_accuracy(codec::FractionAccuracy::UpToTwoDigits),
        options(&source),
    )
    .expect("false replacement flag is source-preserved");
    assert!(contains_bytes(output.bytes(), &[0xa0, 0x01, 0x00]));
    let rewritten = codec::decode_fraction_format(output.bytes(), options(output.bytes()))
        .expect("rewritten false replacement flag");
    assert_eq!(rewritten.requires_fraction_replacement(), Some(false));

    let canonical = codec::canonical_fraction_format(
        codec::FractionFormatWrite::from_accuracy(codec::FractionAccuracy::Halves),
        options(&source),
    )
    .expect("canonical fraction");
    assert_eq!(
        canonical.bytes(),
        fraction_payload(codec::NATIVE_FRACTION_HALVES)
    );
    assert!(!contains_bytes(canonical.bytes(), &[0xa0, 0x01, 0x00]));

    let true_flag =
        with_replacement_flag(fraction_payload(codec::NATIVE_FRACTION_EIGHTHS), &[0x01]);
    let noncanonical_false = with_replacement_flag(
        fraction_payload(codec::NATIVE_FRACTION_EIGHTHS),
        &[0x80, 0x00],
    );
    let duplicate_false = {
        let mut value =
            with_replacement_flag(fraction_payload(codec::NATIVE_FRACTION_EIGHTHS), &[0x00]);
        value.extend_from_slice(&[0xa0, 0x01, 0x00]);
        value
    };
    for malformed in [true_flag, noncanonical_false, duplicate_false] {
        let before = malformed.clone();
        assert!(codec::decode_fraction_format(&malformed, options(&malformed)).is_err());
        assert!(
            codec::decode_fraction_format_with_report(&malformed, options(&malformed)).is_err()
        );
        assert_eq!(malformed, before);
    }

    let wrong_wire = {
        let mut value = fraction_payload(codec::NATIVE_FRACTION_EIGHTHS);
        value.extend_from_slice(&[0xa2, 0x01, 0x00]);
        value
    };
    assert!(codec::decode_fraction_format(&wrong_wire, options(&wrong_wire)).is_err());
}
