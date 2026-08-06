//! Focused invariants for the binary Property Set facade.

use super::super::super::model::{
    Array, CodePage, Dimension, Guid, SUMMARY_INFORMATION_FMTID, Scalar, Section, Stream, Value,
    Vector, VersionedStream,
};

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
    let parsed = Stream::parse(&bytes).expect("versioned stream property should parse");
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
