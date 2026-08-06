use super::codec::{filetime_to_date, filetime_to_duration, parse_typed_property};
use super::*;
use litchi_cfb::OleFile;
use litchi_cfb::consts::*;
use std::io::Cursor;
use std::path::Path;

fn summary_property_stream() -> Vec<u8> {
    let mut data = vec![0u8; 96];
    data[0..2].copy_from_slice(&0xfffeu16.to_le_bytes());
    data[24..28].copy_from_slice(&1u32.to_le_bytes());
    data[44..48].copy_from_slice(&48u32.to_le_bytes());
    data[48..52].copy_from_slice(&48u32.to_le_bytes());
    data[52..56].copy_from_slice(&2u32.to_le_bytes());
    data[56..60].copy_from_slice(&1u32.to_le_bytes());
    data[60..64].copy_from_slice(&24u32.to_le_bytes());
    data[64..68].copy_from_slice(&2u32.to_le_bytes());
    data[68..72].copy_from_slice(&32u32.to_le_bytes());
    data[72..74].copy_from_slice(&VT_I2.to_le_bytes());
    data[76..78].copy_from_slice(&65001u16.to_le_bytes());
    data[80..82].copy_from_slice(&VT_LPSTR.to_le_bytes());
    data[84..88].copy_from_slice(&6u32.to_le_bytes());
    data[88..94].copy_from_slice(b"Hello\0");
    data
}

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../")
            .join(path),
    )
    .unwrap()
}

#[test]
fn parses_typed_stream_and_unsigned_codepage() {
    let stream = Stream::parse(&summary_property_stream()).unwrap();
    let section = &stream.sections[0];
    assert_eq!(section.page().map(CodePage::id), Some(65001));
    assert_eq!(section.property(2), Some(&Value::Lpstr("Hello".into())));
}

#[test]
fn version_one_round_trips_versioned_numeric_properties() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section.add(2, Value::I1(-7)).unwrap();
    section.add(3, Value::Int(-42)).unwrap();
    section.add(4, Value::UInt(42)).unwrap();

    let version_zero = Stream::new(section.clone());
    assert!(version_zero.to_bytes().is_err());

    let mut version_one = Stream::new(section);
    version_one.version = Stream::VERSION_1;
    let bytes = version_one.to_bytes().unwrap();
    let parsed = Stream::parse(&bytes).unwrap();
    assert_eq!(parsed.version, Stream::VERSION_1);
    assert_eq!(parsed.sections[0].property(2), Some(&Value::I1(-7)));
    assert_eq!(parsed.sections[0].property(3), Some(&Value::Int(-42)));
    assert_eq!(parsed.sections[0].property(4), Some(&Value::UInt(42)));

    let mut downgraded = bytes;
    downgraded[2..4].copy_from_slice(&Stream::VERSION_0.to_le_bytes());
    assert!(Stream::parse(&downgraded).is_err());
}

#[test]
fn validates_the_version_one_behavior_property() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section.add(PID_BEHAVIOR, Value::UI4(1)).unwrap();
    let mut stream = Stream::new(section);
    stream.version = Stream::VERSION_1;
    let parsed = Stream::parse(&stream.to_bytes().unwrap()).unwrap();
    assert_eq!(
        parsed.sections[0].property(PID_BEHAVIOR),
        Some(&Value::UI4(1))
    );

    let mut invalid_behavior = Section::new(SUMMARY_INFORMATION_FMTID);
    invalid_behavior.add(PID_BEHAVIOR, Value::I4(1)).unwrap();
    let mut invalid_stream = Stream::new(invalid_behavior);
    invalid_stream.version = Stream::VERSION_1;
    assert!(invalid_stream.to_bytes().is_err());
}

#[test]
fn property_set_pages_are_checked_before_mutation() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    assert!(section.set_page_id(1201).is_err());
    assert_eq!(section.page(), None);
    assert_eq!(section.property(PID_CODEPAGE), None);
    assert!(section.add(PID_CODEPAGE, Value::I2(1201)).is_err());

    section.set_page(CodePage::WINDOWS_1252);
    assert_eq!(section.page(), Some(CodePage::WINDOWS_1252));
    assert_eq!(section.property(PID_CODEPAGE), Some(&Value::I2(1252)));
    assert_eq!(section.clear_page(), Some(CodePage::WINDOWS_1252));
    assert_eq!(section.property(PID_CODEPAGE), None);
}

#[test]
fn rejects_duplicate_offsets_and_truncated_values() {
    let mut duplicate = summary_property_stream();
    duplicate[68..72].copy_from_slice(&24u32.to_le_bytes());
    assert!(Stream::parse(&duplicate).is_err());

    let mut truncated = summary_property_stream();
    truncated[84..88].copy_from_slice(&u32::MAX.to_le_bytes());
    assert!(Stream::parse(&truncated).is_err());
}

#[test]
fn parses_variant_vectors_and_preserves_unknown_values() {
    let mut vector = Vec::new();
    vector.extend_from_slice(&(VT_VECTOR | VT_VARIANT).to_le_bytes());
    vector.extend_from_slice(&0u16.to_le_bytes());
    vector.extend_from_slice(&2u32.to_le_bytes());
    vector.extend_from_slice(&VT_I4.to_le_bytes());
    vector.extend_from_slice(&0u16.to_le_bytes());
    vector.extend_from_slice(&42i32.to_le_bytes());
    vector.extend_from_slice(&VT_BOOL.to_le_bytes());
    vector.extend_from_slice(&0u16.to_le_bytes());
    vector.extend_from_slice(&(-1i16).to_le_bytes());
    vector.extend_from_slice(&0u16.to_le_bytes());
    assert_eq!(
        parse_typed_property(&vector, DEFAULT_CODEPAGE, 0).unwrap(),
        Value::Vector(
            Vector::variant(vec![Value::I4(42), Value::Bool(true)])
                .expect("variant vector should validate"),
        )
    );

    let unknown = [0x34, 0x12, 0, 0, 1, 2, 3, 4];
    assert_eq!(
        parse_typed_property(&unknown, DEFAULT_CODEPAGE, 0).unwrap(),
        Value::Unknown {
            variant_type: 0x1234,
            data: vec![1, 2, 3, 4]
        }
    );
}

#[test]
fn reads_apache_poi_named_custom_properties() {
    let bytes = fixture("test-data/poi/test-data/hpsf/TestMickey.doc");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("sample title"));
    assert_eq!(metadata.subject.as_deref(), Some("sample subject"));
    assert_eq!(metadata.author.as_deref(), Some("Miroslav Obradovic"));
    assert_eq!(metadata.manager.as_deref(), Some("sample manager"));
    assert_eq!(metadata.company.as_deref(), Some("sample company"));
    assert_eq!(
        metadata.custom_properties.get("Client"),
        Some(&Value::Lpstr("sample client".into()))
    );
    assert_eq!(
        metadata.custom_properties.get("Division"),
        Some(&Value::Lpstr("sample division".into()))
    );
}

#[test]
fn reads_apache_poi_two_section_unicode_properties() {
    let bytes = fixture("test-data/poi/test-data/hpsf/TestUnicode.xls");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let stream =
        PropertySetReader::property_set_stream(&mut ole, &["\u{0005}DocumentSummaryInformation"])
            .unwrap();
    assert_eq!(stream.sections.len(), 2);
    let custom = &stream.sections[1];
    assert_eq!(custom.page(), Some(CodePage::Utf16Le));
    assert_eq!(custom.property(2), Some(&Value::I4(-96_070_278)));
    assert_eq!(
        custom.property(3),
        Some(&Value::Lpwstr("MCon_Info zu Office bei Schreiner".into()))
    );
    assert_eq!(
        custom.property(4),
        Some(&Value::Lpwstr("petrovitsch@schreiner-online.de".into()))
    );
    assert_eq!(
        custom.property(5),
        Some(&Value::Lpwstr("Petrovitsch, Wilhelm".into()))
    );
}

#[test]
fn projects_existing_document_properties_fixture() {
    let bytes = fixture("test-data/ole/doc/documentProperties.doc");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("This is document title"));
    assert_eq!(
        metadata.subject.as_deref(),
        Some("This is document subject")
    );
    assert_eq!(metadata.author.as_deref(), Some("Sergey Vladimirov"));
    assert_eq!(metadata.revision_number.as_deref(), Some("0"));
    assert_eq!(
        metadata.create_time.map(|value| value.timestamp()),
        Some(1_309_939_357)
    );
}

#[test]
fn reads_non_dword_and_zero_length_property_fixtures() {
    let bytes = fixture("test-data/poi/test-data/hpsf/TestNon4ByteBoundary.doc");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(
        metadata.creating_application.as_deref(),
        Some("Microsoft Word 10.0")
    );
    assert_eq!(metadata.title.as_deref(), Some(""));
    assert_eq!(metadata.author.as_deref(), Some(""));
    assert_eq!(metadata.company.as_deref(), Some("Cour de Justice"));

    let bytes = fixture("test-data/poi/test-data/hpsf/TestZeroLengthCodePage.mpp");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.creating_application.as_deref(), Some("MSProject"));
    assert_eq!(metadata.title.as_deref(), Some("project1"));
    assert_eq!(metadata.author.as_deref(), Some("Jon Iles"));
    assert_eq!(metadata.company.as_deref(), Some(""));
    let stream =
        PropertySetReader::property_set_stream(&mut ole, &["\u{0005}DocumentSummaryInformation"])
            .unwrap();
    assert_eq!(stream.sections.len(), 2);
}

#[test]
fn reads_undefined_filetime_and_word90_fixtures() {
    let bytes = fixture("test-data/poi/test-data/hpsf/TestBug52117.doc");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.last_printed_time, None);
    assert_eq!(
        metadata.edit_time.map(|value| value.num_milliseconds()),
        Some(180_000)
    );

    let bytes = fixture("test-data/poi/test-data/hpsf/TestGermanWord90.doc");
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let stream =
        PropertySetReader::property_set_stream(&mut ole, &["\u{0005}SummaryInformation"]).unwrap();
    assert_eq!(stream.sections[0].properties.len(), 17);
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Titel"));
    assert_eq!(metadata.author.as_deref(), Some("Rainer Klute (Autor)"));
    assert_eq!(metadata.subject.as_deref(), Some("Thema"));
}

#[test]
fn rejects_unicode_and_vector_allocation_overflows() {
    let mut unicode = Vec::new();
    unicode.extend_from_slice(&VT_LPWSTR.to_le_bytes());
    unicode.extend_from_slice(&0u16.to_le_bytes());
    unicode.extend_from_slice(&0x4000_0001u32.to_le_bytes());
    assert!(parse_typed_property(&unicode, DEFAULT_CODEPAGE, 0).is_err());

    let mut vector = Vec::new();
    vector.extend_from_slice(&(VT_VECTOR | VT_UI1).to_le_bytes());
    vector.extend_from_slice(&0u16.to_le_bytes());
    vector.extend_from_slice(&u32::MAX.to_le_bytes());
    assert!(parse_typed_property(&vector, DEFAULT_CODEPAGE, 0).is_err());
}

#[test]
fn filetime_conversion_is_checked() {
    const UNIX_EPOCH_FILETIME: u64 = 116_444_736_000_000_000;
    assert_eq!(
        filetime_to_date(UNIX_EPOCH_FILETIME).map(|date| date.timestamp()),
        Some(0)
    );
    assert_eq!(
        filetime_to_duration(10).map(|value| value.num_nanoseconds()),
        Some(Some(1000))
    );
    assert!(filetime_to_date(u64::MAX).is_none());
    assert!(filetime_to_duration(u64::MAX).is_none());
}
