//! Focused invariants for the binary Property Set facade.

use super::super::super::model::{
    Array, CodePage, Dimension, Guid, SUMMARY_INFORMATION_FMTID, Scalar, Section, Stream, Value,
    Vector, VersionedStream,
};
use super::wire::{ByteSink, CountingWriter};

fn counted_length(length: usize) -> usize {
    let mut counter = CountingWriter::new();
    counter
        .reserve(length, "test counted serialized bytes")
        .expect("counting sink must represent the test length");
    // `reserve` is the allocation-free counting preflight. The sink only
    // advances its length when bytes are appended, so retain the requested
    // synthetic length for layout-boundary tests without materializing it.
    length
}

#[test]
fn facade_round_trips_typed_property_values() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section.add(2, Value::I4(42)).expect("valid property value");

    let bytes = Stream::new(section)
        .to_bytes()
        .expect("serializable property set");
    let parsed = Stream::parse(&bytes).expect("parseable property set");

    assert_eq!(parsed.sections[0].property(2), Some(&Value::I4(42)));
}

#[test]
fn versioned_stream_round_trip_uses_the_typed_inert_selector() {
    let version_guid = Guid::from_bytes([
        0x10, 0x32, 0x54, 0x76, 0x98, 0xba, 0xdc, 0xfe, 0x01, 0x23, 0x45, 0x67, 0x89, 0xab, 0xcd,
        0xef,
    ]);
    let value = VersionedStream::new(version_guid, 42).expect("normal property identifier");
    assert_eq!(value.stream_name(), "prop42");

    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::Utf16Le);
    section
        .add(42, Value::VersionedStream(value))
        .expect("versioned stream property should be accepted");

    let bytes = Stream::new(section)
        .to_bytes()
        .expect("versioned stream property should serialize");
    let parsed = super::parse_non_simple_stream(&bytes)
        .expect("versioned stream property should parse in non-simple mode");
    let Some(Value::VersionedStream(parsed)) = parsed.sections[0].property(42) else {
        panic!("expected typed versioned stream property");
    };
    assert_eq!(parsed.version_guid(), version_guid);
    assert_eq!(parsed.stream_name(), "prop42");
}

#[test]
fn versioned_stream_rejects_special_property_identifiers() {
    let version_guid = Guid::from_bytes([0xabu8; 16]);
    assert!(VersionedStream::new(version_guid, 0).is_err());
    assert!(VersionedStream::new(version_guid, 1).is_err());
}

#[test]
fn array_values_round_trip_with_row_major_dimensions() {
    let array = Array::new(
        Scalar::I4,
        vec![Dimension::new(2, 0), Dimension::new(3, 1)],
        vec![
            Value::I4(1),
            Value::I4(2),
            Value::I4(3),
            Value::I4(4),
            Value::I4(5),
            Value::I4(6),
        ],
    )
    .expect("array should validate");
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section
        .add(2, Value::Array(array))
        .expect("array property should be accepted");

    let bytes = Stream::new(section)
        .to_bytes()
        .expect("array property set should serialize");
    let parsed = Stream::parse(&bytes).expect("array property set should parse");
    let Value::Array(array) = parsed.sections[0].property(2).expect("array property") else {
        panic!("expected an array value");
    };
    assert_eq!(array.scalar(), Scalar::I4);
    assert_eq!(
        array.dimensions(),
        [Dimension::new(2, 0), Dimension::new(3, 1)]
    );
    assert_eq!(array.value(4), Some(&Value::I4(5)));
}

#[test]
fn variant_array_round_trip_preserves_element_types() {
    let array = Array::variant(
        vec![Dimension::new(2, 0)],
        vec![Value::I4(42), Value::Bstr("two".into())],
    )
    .expect("variant array should validate");
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section
        .add(2, Value::Array(array))
        .expect("variant array property should be accepted");

    let bytes = Stream::new(section)
        .to_bytes()
        .expect("variant array property set should serialize");
    let parsed = Stream::parse(&bytes).expect("variant array property set should parse");
    assert_eq!(
        parsed.sections[0].property(2),
        Some(&Value::Array(
            Array::variant(
                vec![Dimension::new(2, 0)],
                vec![Value::I4(42), Value::Bstr("two".into())]
            )
            .expect("expected valid variant array")
        ))
    );
}

#[test]
fn homogeneous_vector_round_trip_uses_its_scalar_type() {
    let vector = Vector::new(Scalar::UI2, vec![Value::UI2(7), Value::UI2(11)])
        .expect("homogeneous vector should validate");
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section
        .add(2, Value::Vector(vector))
        .expect("vector property should be accepted");

    let bytes = Stream::new(section)
        .to_bytes()
        .expect("vector property set should serialize");
    let parsed = Stream::parse(&bytes).expect("vector property set should parse");
    let Value::Vector(vector) = parsed.sections[0].property(2).expect("vector property") else {
        panic!("expected a vector value");
    };
    assert_eq!(vector.scalar(), Scalar::UI2);
    assert_eq!(vector.values(), [Value::UI2(7), Value::UI2(11)]);
}

#[test]
fn narrow_vectors_and_arrays_pad_the_scalar_sequence_once() {
    let vector = Vector::new(Scalar::I1, vec![Value::I1(-3), Value::I1(7), Value::I1(9)])
        .expect("narrow vector should validate");
    let array = Array::new(
        Scalar::UI2,
        vec![Dimension::new(3, 0)],
        vec![Value::UI2(1), Value::UI2(2), Value::UI2(3)],
    )
    .expect("narrow array should validate");
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section
        .add(2, Value::Vector(vector))
        .expect("narrow vector should be accepted");
    section
        .add(3, Value::Array(array))
        .expect("narrow array should be accepted");

    let mut stream = Stream::new(section);
    stream.version = Stream::VERSION_1;
    let bytes = stream.to_bytes().expect("narrow values should serialize");
    assert_eq!(
        stream
            .to_bytes_with_limit(bytes.len() as u64)
            .expect("bounded narrow values should serialize"),
        bytes
    );
    let parsed = Stream::parse(&bytes).expect("narrow values should parse");
    assert_eq!(
        parsed.sections[0].property(2),
        Some(&Value::Vector(
            Vector::new(Scalar::I1, vec![Value::I1(-3), Value::I1(7), Value::I1(9)])
                .expect("expected valid vector")
        ))
    );
    assert_eq!(
        parsed.sections[0].property(3),
        Some(&Value::Array(
            Array::new(
                Scalar::UI2,
                vec![Dimension::new(3, 0)],
                vec![Value::UI2(1), Value::UI2(2), Value::UI2(3)]
            )
            .expect("expected valid array")
        ))
    );
}

#[test]
fn arrays_reject_wrong_shape_and_scalar_values() {
    assert!(Array::new(Scalar::I4, vec![Dimension::new(2, 0)], vec![Value::I4(1)]).is_err());
    assert!(
        Array::new(
            Scalar::I4,
            vec![Dimension::new(1, 0)],
            vec![Value::Lpwstr("wrong".into())]
        )
        .is_err()
    );
    assert!(Array::new(Scalar::I8, vec![Dimension::new(1, 0)], vec![Value::I8(1)]).is_err());
    assert!(Array::variant(vec![Dimension::new(1, 0)], vec![Value::I8(1)]).is_err());
}

#[test]
fn bounded_serialization_accepts_exact_stream_bytes_and_rejects_one_under() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section
        .add(2, Value::I4(42))
        .expect("ordinary value should be accepted");
    let stream = Stream::new(section);
    let canonical = stream.to_bytes().expect("canonical stream");

    assert_eq!(
        stream
            .to_bytes_with_limit(canonical.len() as u64)
            .expect("exact stream ceiling"),
        canonical
    );
    assert!(matches!(
        stream.to_bytes_with_limit((canonical.len() - 1) as u64),
        Err(litchi_cfb::OleError::LimitExceeded { .. })
    ));
    assert!(matches!(
        stream.to_bytes_with_limit(0),
        Err(litchi_cfb::OleError::LimitExceeded { .. })
    ));
}

#[test]
fn bounded_serialization_rejects_giant_blob_before_output_materialization() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section
        .add(2, Value::Blob(vec![0x5a; 4096]))
        .expect("blob should be accepted");
    let stream = Stream::new(section);

    assert!(matches!(
        stream.to_bytes_with_limit(128),
        Err(litchi_cfb::OleError::LimitExceeded { maximum: 128, .. })
    ));
}

#[test]
fn bounded_serialization_preserves_ansi_and_unicode_boundaries() {
    let mut ansi_section = Section::new(SUMMARY_INFORMATION_FMTID);
    ansi_section.set_page(CodePage::WINDOWS_1252);
    ansi_section
        .add(2, Value::Lpstr("café".into()))
        .expect("ANSI value should be accepted");
    let ansi_stream = Stream::new(ansi_section);
    let ansi_bytes = ansi_stream.to_bytes().expect("ANSI stream");
    assert_eq!(
        ansi_stream
            .to_bytes_with_limit(ansi_bytes.len() as u64)
            .expect("exact ANSI ceiling"),
        ansi_bytes
    );
    assert!(matches!(
        ansi_stream.to_bytes_with_limit((ansi_bytes.len() - 1) as u64),
        Err(litchi_cfb::OleError::LimitExceeded { .. })
    ));

    let mut giant_ansi_section = Section::new(SUMMARY_INFORMATION_FMTID);
    giant_ansi_section.set_page(CodePage::WINDOWS_1252);
    giant_ansi_section
        .add(2, Value::Lpstr("a".repeat(4096)))
        .expect("large ANSI value should be accepted");
    assert!(matches!(
        Stream::new(giant_ansi_section).to_bytes_with_limit(128),
        Err(litchi_cfb::OleError::LimitExceeded { maximum: 128, .. })
    ));

    let mut unicode_section = Section::new(SUMMARY_INFORMATION_FMTID);
    unicode_section.set_page(CodePage::Utf16Le);
    unicode_section
        .add(2, Value::Lpwstr("雪だるま🌨".into()))
        .expect("Unicode value should be accepted");
    let unicode_stream = Stream::new(unicode_section);
    let unicode_bytes = unicode_stream.to_bytes().expect("Unicode stream");
    assert_eq!(
        unicode_stream
            .to_bytes_with_limit(unicode_bytes.len() as u64)
            .expect("exact Unicode ceiling"),
        unicode_bytes
    );
    assert!(matches!(
        unicode_stream.to_bytes_with_limit((unicode_bytes.len() - 1) as u64),
        Err(litchi_cfb::OleError::LimitExceeded { .. })
    ));
}

#[test]
fn bounded_serialization_preserves_composites_and_opaque_values() {
    let mut section = Section::new(crate::property_set::DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::Utf16Le);
    section
        .add(
            crate::property_set::PID_DOC_PARTS,
            Value::DocParts(
                crate::property_set::DocParts::unicode(vec!["Résumé".into(), "Résumé/雪".into()])
                    .expect("document parts should be constructible"),
            ),
        )
        .expect("composite should be accepted");
    section
        .add(
            42,
            Value::Unknown {
                variant_type: 0x7777,
                data: vec![1, 2, 3, 4, 5],
            },
        )
        .expect("opaque value should be accepted");
    let stream = Stream::new(section);
    let canonical = stream.to_bytes().expect("composite stream");
    assert_eq!(
        stream
            .to_bytes_with_limit(canonical.len() as u64)
            .expect("exact composite ceiling"),
        canonical
    );
    assert!(matches!(
        stream.to_bytes_with_limit((canonical.len() - 1) as u64),
        Err(litchi_cfb::OleError::LimitExceeded { .. })
    ));

    let mut giant_section = Section::new(crate::property_set::DOCUMENT_SUMMARY_INFORMATION_FMTID);
    giant_section.set_page(CodePage::Utf16Le);
    giant_section
        .add(
            crate::property_set::PID_DOC_PARTS,
            Value::DocParts(
                crate::property_set::DocParts::unicode(vec!["x".repeat(4096)])
                    .expect("large document part should be constructible"),
            ),
        )
        .expect("large composite should be accepted");
    assert!(matches!(
        Stream::new(giant_section).to_bytes_with_limit(128),
        Err(litchi_cfb::OleError::LimitExceeded { maximum: 128, .. })
    ));
}

#[test]
fn bounded_preflight_rejects_unrepresentable_wire_offsets_without_materializing() {
    let over_u32 = counted_length(u32::MAX as usize);

    assert!(matches!(
        super::validation::serialized_section_layout_len(1, 8, [Ok(over_u32)]),
        Err(litchi_cfb::OleError::InvalidFormat(message))
            if message.contains("Section length exceeds u32")
    ));
    assert!(matches!(
        super::validation::serialized_section_layout_len(2, 24, [Ok(over_u32), Ok(0)]),
        Err(litchi_cfb::OleError::InvalidFormat(message))
            if message.contains("Property offset exceeds u32")
    ));
    assert!(matches!(
        super::validation::serialized_stream_layout_len(2, 48, [Ok(over_u32), Ok(0)]),
        Err(litchi_cfb::OleError::InvalidFormat(message))
            if message.contains("Section offset exceeds u32")
    ));
}
