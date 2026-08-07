use super::{
    EMFPLUS_COMMENT_IDENTIFIER, EMR_COMMENT, EmfPlusRecordIter, EmfPlusStreamValidator,
    ObjectRecordFlags, ObjectType, ParserLimits, RecordFlags, RecordType,
    extract_emfplus_comment_body, extract_emfplus_comment_record, try_extract_emfplus_comment_body,
    try_extract_emfplus_comment_record, validate_complete_stream,
};

fn limits() -> ParserLimits {
    ParserLimits {
        max_bytes: 4096,
        max_records: 128,
        max_object_slots: 64,
    }
}

fn record(record_type: u16, flags: u16, data: &[u8]) -> Vec<u8> {
    assert_eq!(data.len() % 4, 0);
    let data_size = u32::try_from(data.len())
        .unwrap_or_else(|error| panic!("test record is too large: {error}"));
    let size = data_size + 12;
    let mut bytes = Vec::with_capacity(size as usize);
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&flags.to_le_bytes());
    bytes.extend_from_slice(&size.to_le_bytes());
    bytes.extend_from_slice(&data_size.to_le_bytes());
    bytes.extend_from_slice(data);
    bytes
}

fn header() -> Vec<u8> {
    record(RecordType::Header.raw(), 0, &[0; 16])
}

fn eof() -> Vec<u8> {
    record(RecordType::EndOfFile.raw(), 0, &[])
}

fn complete_stream() -> Vec<u8> {
    let mut bytes = header();
    bytes.extend_from_slice(&record(RecordType::Clear.raw(), 0, &[0; 4]));
    bytes.extend_from_slice(&eof());
    bytes
}

fn comment(payload: &[u8]) -> Vec<u8> {
    let data_size = 4usize + payload.len();
    let unpadded = 12usize + data_size;
    let padding = (4 - unpadded % 4) % 4;
    let size = unpadded + padding;
    let mut bytes = Vec::with_capacity(size);
    bytes.extend_from_slice(&EMR_COMMENT.to_le_bytes());
    let size_u32 =
        u32::try_from(size).unwrap_or_else(|error| panic!("test comment is too large: {error}"));
    let data_size_u32 = u32::try_from(data_size)
        .unwrap_or_else(|error| panic!("test comment data is too large: {error}"));
    bytes.extend_from_slice(&size_u32.to_le_bytes());
    bytes.extend_from_slice(&data_size_u32.to_le_bytes());
    bytes.extend_from_slice(&EMFPLUS_COMMENT_IDENTIFIER.to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes.resize(size, 0);
    bytes
}

fn first_record(data: &[u8]) -> super::EmfPlusRecord<'_> {
    let mut iter = match EmfPlusRecordIter::new(data, limits()) {
        Ok(value) => value,
        Err(error) => panic!("iterator construction failed: {error}"),
    };
    match iter.next() {
        Some(Ok(value)) => value,
        Some(Err(error)) => panic!("record parse failed: {error}"),
        None => panic!("record was missing"),
    }
}

fn assert_first_error(data: &[u8]) {
    let Ok(mut iter) = EmfPlusRecordIter::new(data, limits()) else {
        return;
    };
    assert!(matches!(iter.next(), Some(Err(_))));
    assert!(iter.next().is_none());
}

#[test]
fn frames_multiple_records_without_copying() {
    let bytes = complete_stream();
    let iter = match EmfPlusRecordIter::new(&bytes, limits()) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    let records: Vec<_> = iter
        .map(|result| match result {
            Ok(value) => value,
            Err(error) => panic!("unexpected error: {error}"),
        })
        .collect();
    assert_eq!(records.len(), 3);
    assert_eq!(records[0].offset, 0);
    assert_eq!(records[0].header.record_type, RecordType::Header);
    assert_eq!(records[0].data.len(), 16);
    assert_eq!(records[1].offset, 28);
    assert_eq!(records[1].bytes().len(), 16);
    assert_eq!(records[2].header.record_type, RecordType::EndOfFile);
}

#[test]
fn recognizes_every_specified_record_identifier() {
    assert_eq!(RecordType::ALL.len(), 0x403A - 0x4001 + 1);
    for (index, expected) in RecordType::ALL.iter().copied().enumerate() {
        let index_u16 = u16::try_from(index)
            .unwrap_or_else(|error| panic!("record index does not fit u16: {error}"));
        let raw = 0x4001u16 + index_u16;
        assert_eq!(expected.raw(), raw);
        assert_eq!(RecordType::from_raw(raw), Some(expected));
        let bytes = record(raw, 0, &[]);
        assert_eq!(first_record(&bytes).header.record_type, expected);
    }
    assert_eq!(RecordType::from_raw(0x4000), None);
    assert_eq!(RecordType::from_raw(0x403B), None);
}

#[test]
fn rejects_unknown_record_identifiers() {
    assert_first_error(&record(0x4000, 0, &[]));
    assert_first_error(&record(0x403B, 0, &[]));
    assert_first_error(&record(0xFFFF, 0, &[]));
}

#[test]
fn rejects_every_truncated_header_length() {
    let full = eof();
    for length in 1..12 {
        assert_first_error(&full[..length]);
    }
}

#[test]
fn empty_payload_is_an_empty_iterator_but_not_a_complete_stream() {
    let mut iter = match EmfPlusRecordIter::new(&[], limits()) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(iter.next().is_none());
    assert!(validate_complete_stream(&[], limits()).is_err());
}

#[test]
fn rejects_bad_record_sizes_and_alignment() {
    let mut undersized = eof();
    undersized[4..8].copy_from_slice(&8u32.to_le_bytes());
    assert_first_error(&undersized);

    let mut misaligned_size = eof();
    misaligned_size[4..8].copy_from_slice(&13u32.to_le_bytes());
    misaligned_size[8..12].copy_from_slice(&1u32.to_le_bytes());
    misaligned_size.push(0);
    assert_first_error(&misaligned_size);

    let mut mismatch = eof();
    mismatch[4..8].copy_from_slice(&16u32.to_le_bytes());
    assert_first_error(&mismatch);

    let mut overflow = eof();
    overflow[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
    overflow[8..12].copy_from_slice(&(u32::MAX - 3).to_le_bytes());
    assert_first_error(&overflow);
}

#[test]
fn rejects_records_extending_past_payload_and_trailing_garbage() {
    let mut truncated_data = record(RecordType::Clear.raw(), 0, &[0; 8]);
    truncated_data.truncate(16);
    assert_first_error(&truncated_data);

    let mut trailing = eof();
    trailing.extend_from_slice(&[1, 2, 3, 4]);
    let mut iter = match EmfPlusRecordIter::new(&trailing, limits()) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(matches!(iter.next(), Some(Ok(_))));
    assert!(matches!(iter.next(), Some(Err(_))));
}

#[test]
fn enforces_iterator_byte_and_record_limits() {
    let tiny_bytes = ParserLimits {
        max_bytes: 12,
        ..limits()
    };
    assert!(
        EmfPlusRecordIter::new(&record(RecordType::Clear.raw(), 0, &[0; 4]), tiny_bytes).is_err()
    );

    let one_record = ParserLimits {
        max_records: 1,
        ..limits()
    };
    let mut bytes = eof();
    bytes.extend_from_slice(&eof());
    let mut iter = match EmfPlusRecordIter::new(&bytes, one_record) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(matches!(iter.next(), Some(Ok(_))));
    assert!(matches!(iter.next(), Some(Err(_))));
}

#[test]
fn validates_parser_limit_configuration() {
    assert!(
        ParserLimits {
            max_bytes: 11,
            ..limits()
        }
        .validate()
        .is_err()
    );
    assert!(
        ParserLimits {
            max_records: 0,
            ..limits()
        }
        .validate()
        .is_err()
    );
    assert!(
        ParserLimits {
            max_object_slots: 0,
            ..limits()
        }
        .validate()
        .is_err()
    );
    assert!(
        ParserLimits {
            max_object_slots: 65,
            ..limits()
        }
        .validate()
        .is_err()
    );
}

#[test]
fn extracts_full_comment_and_existing_parser_body_layout() {
    let payload = complete_stream();
    let bytes = comment(&payload);
    let full = extract_emfplus_comment_record(&bytes, limits());
    assert!(matches!(full, Ok(value) if value == payload));
    let body = extract_emfplus_comment_body(&bytes[8..], limits());
    assert!(matches!(body, Ok(value) if value == payload));
}

#[test]
fn distinguishes_other_records_and_comment_signatures() {
    let payload = eof();
    let mut bytes = comment(&payload);
    bytes[0..4].copy_from_slice(&1u32.to_le_bytes());
    assert!(matches!(
        try_extract_emfplus_comment_record(&bytes, limits()),
        Ok(None)
    ));
    assert!(extract_emfplus_comment_record(&bytes, limits()).is_err());

    let mut bytes = comment(&payload);
    bytes[12..16].copy_from_slice(&0x4349_4447u32.to_le_bytes());
    assert!(matches!(
        try_extract_emfplus_comment_body(&bytes[8..], limits()),
        Ok(None)
    ));
    assert!(extract_emfplus_comment_body(&bytes[8..], limits()).is_err());
}

#[test]
fn rejects_truncated_and_inconsistent_comment_envelopes() {
    assert!(extract_emfplus_comment_record(&[0; 7], limits()).is_err());
    assert!(extract_emfplus_comment_body(&[0; 7], limits()).is_err());

    let mut bytes = comment(&eof());
    bytes.truncate(bytes.len() - 1);
    assert!(extract_emfplus_comment_record(&bytes, limits()).is_err());

    let mut bytes = comment(&eof());
    bytes[4..8].copy_from_slice(&18u32.to_le_bytes());
    assert!(extract_emfplus_comment_record(&bytes, limits()).is_err());

    let mut bytes = comment(&eof());
    bytes[8..12].copy_from_slice(&3u32.to_le_bytes());
    assert!(extract_emfplus_comment_record(&bytes, limits()).is_err());

    let mut bytes = comment(&eof());
    bytes[8..12].copy_from_slice(&0xFFFF_FFFCu32.to_le_bytes());
    assert!(extract_emfplus_comment_record(&bytes, limits()).is_err());
}

#[test]
fn rejects_empty_or_resource_excessive_emfplus_comment_payloads() {
    let empty = comment(&[]);
    assert!(extract_emfplus_comment_record(&empty, limits()).is_err());

    let bytes = comment(&record(RecordType::Comment.raw(), 0, &[0; 4]));
    let small = ParserLimits {
        max_bytes: 12,
        ..limits()
    };
    assert!(extract_emfplus_comment_record(&bytes, small).is_err());
}

#[test]
fn validates_header_and_eof_semantics() {
    assert!(validate_complete_stream(&complete_stream(), limits()).is_ok());
    assert!(validate_complete_stream(&eof(), limits()).is_err());
    assert!(validate_complete_stream(&header(), limits()).is_err());

    let mut duplicate_header = header();
    duplicate_header.extend_from_slice(&header());
    duplicate_header.extend_from_slice(&eof());
    assert!(validate_complete_stream(&duplicate_header, limits()).is_err());

    let mut after_eof = header();
    after_eof.extend_from_slice(&eof());
    after_eof.extend_from_slice(&record(RecordType::Clear.raw(), 0, &[0; 4]));
    assert!(validate_complete_stream(&after_eof, limits()).is_err());

    let mut bad_eof = header();
    bad_eof.extend_from_slice(&record(RecordType::EndOfFile.raw(), 1, &[]));
    assert!(validate_complete_stream(&bad_eof, limits()).is_err());
}

#[test]
fn rejects_bad_header_body_and_reserved_records() {
    let mut short_header = record(RecordType::Header.raw(), 0, &[0; 12]);
    short_header.extend_from_slice(&eof());
    assert!(validate_complete_stream(&short_header, limits()).is_err());

    for reserved in [
        RecordType::MultiFormatStart,
        RecordType::MultiFormatSection,
        RecordType::MultiFormatEnd,
    ] {
        let mut bytes = header();
        bytes.extend_from_slice(&record(reserved.raw(), 0, &[]));
        bytes.extend_from_slice(&eof());
        assert!(validate_complete_stream(&bytes, limits()).is_err());
    }
}

#[test]
fn validates_a_stream_split_across_comment_payloads() {
    let mut validator = match EmfPlusStreamValidator::new(limits()) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(matches!(validator.push_payload(&header()), Ok(1)));
    assert!(matches!(
        validator.push_payload(&record(RecordType::Clear.raw(), 0, &[0; 4])),
        Ok(1)
    ));
    assert!(matches!(validator.push_payload(&eof()), Ok(1)));
    assert_eq!(validator.records_seen(), 3);
    assert_eq!(validator.bytes_seen(), 56);
    assert!(validator.finish().is_ok());
}

#[test]
fn validates_object_flag_helpers_and_slot_limits() {
    let flags = RecordFlags::new(0x8503);
    assert_eq!(flags.low_byte(), 3);
    assert_eq!(flags.high_byte(), 0x85);
    let decoded = ObjectRecordFlags::parse(flags, limits());
    assert!(matches!(decoded, Ok(value)
        if value.continued
            && value.object_type == ObjectType::Image
            && value.object_id.get() == 3));

    assert!(ObjectRecordFlags::parse(RecordFlags::new(0x013F), limits()).is_ok());
    assert!(ObjectRecordFlags::parse(RecordFlags::new(0x0140), limits()).is_err());
    assert!(ObjectRecordFlags::parse(RecordFlags::new(0x0000), limits()).is_err());
    assert!(ObjectRecordFlags::parse(RecordFlags::new(0x0A00), limits()).is_err());

    let fewer_slots = ParserLimits {
        max_object_slots: 4,
        ..limits()
    };
    assert!(RecordFlags::new(3).object_id(fewer_slots).is_ok());
    assert!(RecordFlags::new(4).object_id(fewer_slots).is_err());
}

#[test]
fn object_fragment_strips_total_size_only_on_continued_records() {
    let bytes = record(RecordType::Object.raw(), 0x8102, &[8, 0, 0, 0, 1, 2, 3, 4]);
    let fragment = first_record(&bytes).object_fragment(limits());
    assert!(matches!(fragment, Ok(value)
        if value.total_object_size == Some(8) && value.data == [1, 2, 3, 4]));

    let bytes = record(RecordType::Object.raw(), 0x0102, &[1, 2, 3, 4]);
    let fragment = first_record(&bytes).object_fragment(limits());
    assert!(matches!(fragment, Ok(value)
        if value.total_object_size.is_none() && value.data == [1, 2, 3, 4]));

    assert!(first_record(&eof()).object_fragment(limits()).is_err());
    let missing_total = record(RecordType::Object.raw(), 0x8102, &[]);
    assert!(
        first_record(&missing_total)
            .object_fragment(limits())
            .is_err()
    );
}

#[test]
fn accepts_object_continuation_across_comment_boundaries() {
    let first = record(RecordType::Object.raw(), 0x8307, &[8, 0, 0, 0, 1, 2, 3, 4]);
    let final_part = record(RecordType::Object.raw(), 0x0307, &[5, 6, 7, 8]);
    let mut validator = match EmfPlusStreamValidator::new(limits()) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(validator.push_payload(&header()).is_ok());
    assert!(validator.push_payload(&first).is_ok());
    assert!(validator.push_payload(&final_part).is_ok());
    assert!(validator.push_payload(&eof()).is_ok());
    assert!(validator.finish().is_ok());
}

#[test]
fn accepts_multiple_continued_object_fragments() {
    let mut bytes = header();
    bytes.extend_from_slice(&record(
        RecordType::Object.raw(),
        0x8201,
        &[12, 0, 0, 0, 1, 2, 3, 4],
    ));
    bytes.extend_from_slice(&record(
        RecordType::Object.raw(),
        0x8201,
        &[12, 0, 0, 0, 5, 6, 7, 8],
    ));
    bytes.extend_from_slice(&record(RecordType::Object.raw(), 0x0201, &[9, 10, 11, 12]));
    bytes.extend_from_slice(&eof());
    assert!(validate_complete_stream(&bytes, limits()).is_ok());
}

#[test]
fn rejects_interrupted_or_unfinished_continuation() {
    let continued = record(RecordType::Object.raw(), 0x8101, &[8, 0, 0, 0, 1, 2, 3, 4]);
    let mut interrupted = header();
    interrupted.extend_from_slice(&continued);
    interrupted.extend_from_slice(&record(RecordType::Clear.raw(), 0, &[0; 4]));
    interrupted.extend_from_slice(&eof());
    assert!(validate_complete_stream(&interrupted, limits()).is_err());

    let mut unfinished = header();
    unfinished.extend_from_slice(&continued);
    assert!(validate_complete_stream(&unfinished, limits()).is_err());
}

#[test]
fn rejects_continuation_identity_and_total_size_changes() {
    let first = record(RecordType::Object.raw(), 0x8101, &[12, 0, 0, 0, 1, 2, 3, 4]);
    for changed in [
        record(RecordType::Object.raw(), 0x8102, &[12, 0, 0, 0, 5, 6, 7, 8]),
        record(RecordType::Object.raw(), 0x8201, &[12, 0, 0, 0, 5, 6, 7, 8]),
        record(RecordType::Object.raw(), 0x8101, &[16, 0, 0, 0, 5, 6, 7, 8]),
    ] {
        let mut bytes = header();
        bytes.extend_from_slice(&first);
        bytes.extend_from_slice(&changed);
        assert!(validate_complete_stream(&bytes, limits()).is_err());
    }
}

#[test]
fn accepts_continuation_flag_on_final_bytes() {
    let mut bytes = header();
    bytes.extend_from_slice(&record(
        RecordType::Object.raw(),
        0x8101,
        &[4, 0, 0, 0, 1, 2, 3, 4],
    ));
    bytes.extend_from_slice(&eof());
    assert!(validate_complete_stream(&bytes, limits()).is_ok());
}

#[test]
fn rejects_short_and_overlong_continuation_endings() {
    let cases = [
        (
            record(RecordType::Object.raw(), 0x8101, &[12, 0, 0, 0, 1, 2, 3, 4]),
            record(RecordType::Object.raw(), 0x0101, &[5, 6, 7, 8]),
        ),
        (
            record(RecordType::Object.raw(), 0x8101, &[8, 0, 0, 0, 1, 2, 3, 4]),
            record(RecordType::Object.raw(), 0x0101, &[0; 8]),
        ),
    ];
    for (first, final_part) in cases {
        let mut bytes = header();
        bytes.extend_from_slice(&first);
        bytes.extend_from_slice(&final_part);
        bytes.extend_from_slice(&eof());
        assert!(validate_complete_stream(&bytes, limits()).is_err());
    }
}

#[test]
fn continuation_total_and_cumulative_stream_obey_byte_limit() {
    let constrained = ParserLimits {
        max_bytes: 39,
        ..limits()
    };
    let mut validator = match EmfPlusStreamValidator::new(constrained) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(validator.push_payload(&header()).is_ok());
    assert!(validator.push_payload(&eof()).is_err());

    let total_too_large = record(
        RecordType::Object.raw(),
        0x8101,
        &[0x04, 0x10, 0x00, 0x00, 1, 2, 3, 4],
    );
    let mut validator = match EmfPlusStreamValidator::new(limits()) {
        Ok(value) => value,
        Err(error) => panic!("unexpected error: {error}"),
    };
    assert!(validator.push_payload(&header()).is_ok());
    assert!(validator.push_payload(&total_too_large).is_err());
}
