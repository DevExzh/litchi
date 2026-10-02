//! Default-repair fixtures for the property-set tolerance contract.
//!
//! Each defect class has a hand-built minimal stream that carries the
//! deviation and must parse by default with the repair recorded; the same
//! stream without the deviation must parse with a clean report.

use super::*;
use crate::leniency::{Defect, PropertyDefect, StreamDefect, StringDefect, ToleranceReport};
use litchi_cfb::OleFile;
use litchi_cfb::consts::*;
use litchi_cfb::writer::OleWriter;
use std::io::Cursor;

const PID_TITLE: u32 = 2;
const PID_SUBJECT: u32 = 3;

fn u16le(value: u16) -> [u8; 2] {
    value.to_le_bytes()
}

fn u32le(value: u32) -> [u8; 4] {
    value.to_le_bytes()
}

/// A typed value: variant type, reserved word, body.
fn typed(variant_type: u16, body: &[u8]) -> Vec<u8> {
    let mut value = Vec::new();
    value.extend_from_slice(&u16le(variant_type));
    value.extend_from_slice(&u16le(0));
    value.extend_from_slice(body);
    value
}

fn codepage_property(page: u16) -> (u32, Vec<u8>) {
    let mut body = u16le(page).to_vec();
    body.extend_from_slice(&[0, 0]);
    (1, typed(VT_I2, &body))
}

/// A VT_LPSTR body with an explicit byte count (terminator included or not
/// as the caller writes it), padded to a DWORD.
fn lpstr(raw: &[u8]) -> Vec<u8> {
    let mut body = u32le(u32::try_from(raw.len()).unwrap()).to_vec();
    body.extend_from_slice(raw);
    while body.len() % 4 != 0 {
        body.push(0);
    }
    typed(VT_LPSTR, &body)
}

/// A VT_LPWSTR body from UTF-16 units (terminator as the caller writes it).
fn lpwstr(units: &[u16]) -> Vec<u8> {
    let mut body = u32le(u32::try_from(units.len()).unwrap()).to_vec();
    for unit in units {
        body.extend_from_slice(&u16le(*unit));
    }
    while body.len() % 4 != 0 {
        body.push(0);
    }
    typed(VT_LPWSTR, &body)
}

/// One section body (size, count, descriptors, values) for `properties`.
fn section_body(properties: &[(u32, Vec<u8>)]) -> Vec<u8> {
    let table_end = 8 + properties.len() * 8;
    let mut values = Vec::new();
    let mut descriptors = Vec::new();
    for (identifier, value) in properties {
        descriptors.extend_from_slice(&u32le(*identifier));
        descriptors.extend_from_slice(&u32le(u32::try_from(table_end + values.len()).unwrap()));
        values.extend_from_slice(value);
    }
    let size = table_end + values.len();
    let mut section = u32le(u32::try_from(size).unwrap()).to_vec();
    section.extend_from_slice(&u32le(u32::try_from(properties.len()).unwrap()));
    section.extend_from_slice(&descriptors);
    section.extend_from_slice(&values);
    section
}

/// A complete one-section `SummaryInformation` stream.
fn summary_stream(properties: &[(u32, Vec<u8>)]) -> Vec<u8> {
    one_section_stream(SUMMARY_INFORMATION_FMTID, properties)
}

/// A complete one-section stream for `format_identifier`.
fn one_section_stream(format_identifier: Guid, properties: &[(u32, Vec<u8>)]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&u16le(0xfffe));
    data.extend_from_slice(&u16le(0));
    data.extend_from_slice(&u32le(0));
    data.extend_from_slice(&[0; 16]);
    data.extend_from_slice(&u32le(1));
    data.extend_from_slice(format_identifier.as_bytes());
    data.extend_from_slice(&u32le(48));
    data.extend_from_slice(&section_body(properties));
    data
}

fn recorded(report: &ToleranceReport) -> Vec<(Defect, u32)> {
    report
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

fn section_of(stream: &Stream) -> &Section {
    stream.section(SUMMARY_INFORMATION_FMTID).unwrap()
}

/// A compound file holding `stream` as `\x05SummaryInformation`.
fn compound_file(stream: &[u8]) -> Vec<u8> {
    let name = Binding::SummaryInformation.name();
    let mut writer = OleWriter::new();
    writer.create_stream(&[name.as_str()], stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

// --- F7.8: strict MBCS and UTF-16 decode ---

#[test]
fn malformed_mbcs_and_utf16_strings_decode_lossily_and_are_recorded() {
    let clean = summary_stream(&[
        codepage_property(932),
        (PID_TITLE, lpstr(b"Title\0")),
        (PID_SUBJECT, lpwstr(&[0x41, 0x42, 0])),
    ]);
    let parsed = Stream::parse(&clean).unwrap();
    assert!(parsed.tolerance_report().is_clean());

    // A Shift-JIS lead byte truncated by the terminator, and an unpaired
    // UTF-16 surrogate.
    let faulted = summary_stream(&[
        codepage_property(932),
        (PID_TITLE, lpstr(b"Ti\x82\0")),
        (PID_SUBJECT, lpwstr(&[0xd800, 0x41, 0])),
    ]);
    let parsed = Stream::parse(&faulted).unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [
            (StringDefect::LossyDecode.into(), PID_TITLE),
            (StringDefect::LossyDecode.into(), PID_SUBJECT),
        ]
    );
    let section = section_of(&parsed);
    assert_eq!(
        section.property(PID_TITLE),
        Some(&Value::Lpstr("Ti\u{fffd}".into()))
    );
    assert_eq!(
        section.property(PID_SUBJECT),
        Some(&Value::Lpwstr("\u{fffd}A".into()))
    );

    // Reading through the compound file yields the same metadata, but the
    // editor refuses to write the substituted text back as the source.
    let bytes = compound_file(&faulted);
    let mut ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Ti\u{fffd}"));
    let mut editor = Editor::new(bytes).unwrap();
    assert!(
        editor
            .property_set(Binding::SummaryInformation)
            .unwrap()
            .is_some()
    );
    assert!(
        editor
            .update(Binding::SummaryInformation, |_| Ok(()))
            .is_err()
    );
    let mut clean_editor = Editor::new(compound_file(&clean)).unwrap();
    clean_editor
        .update(Binding::SummaryInformation, |_| Ok(()))
        .unwrap();
}

// --- F7.9: PID 1 code page outside the supported table ---

#[test]
fn unsupported_codepage_falls_back_to_windows_1252_and_is_recorded() {
    let clean = summary_stream(&[codepage_property(1252), (PID_TITLE, lpstr(b"Caf\xe9\0"))]);
    let parsed = Stream::parse(&clean).unwrap();
    assert!(parsed.tolerance_report().is_clean());

    for page in [28591u16, 20127, 437] {
        let faulted = summary_stream(&[codepage_property(page), (PID_TITLE, lpstr(b"Caf\xe9\0"))]);
        let parsed = Stream::parse(&faulted).unwrap();
        assert_eq!(
            recorded(parsed.tolerance_report()),
            [(StringDefect::UnsupportedCodePage.into(), u32::from(page))],
            "code page {page}"
        );
        let section = section_of(&parsed);
        assert_eq!(section.page(), None);
        assert_eq!(section.property(1), None);
        assert_eq!(
            section.property(PID_TITLE),
            Some(&Value::Lpstr("Caf\u{e9}".into()))
        );

        let mut editor = Editor::new(compound_file(&faulted)).unwrap();
        assert!(
            editor
                .update(Binding::SummaryInformation, |_| Ok(()))
                .is_err()
        );
    }
}

// --- F7.10: strings not NUL-terminated within their declared size ---

#[test]
fn unterminated_strings_are_read_to_their_declared_size_and_recorded() {
    let faulted = summary_stream(&[
        codepage_property(1252),
        (PID_TITLE, lpstr(b"Hello")),
        (PID_SUBJECT, lpwstr(&[0x41, 0x42])),
    ]);
    let parsed = Stream::parse(&faulted).unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [
            (StringDefect::MissingTerminator.into(), PID_TITLE),
            (StringDefect::MissingTerminator.into(), PID_SUBJECT),
        ]
    );
    let section = section_of(&parsed);
    assert_eq!(
        section.property(PID_TITLE),
        Some(&Value::Lpstr("Hello".into()))
    );
    assert_eq!(
        section.property(PID_SUBJECT),
        Some(&Value::Lpwstr("AB".into()))
    );

    // The repair keeps the content, so the editor may rewrite the stream;
    // the rewrite is canonical and therefore conformant.
    let mut editor = Editor::new(compound_file(&faulted)).unwrap();
    editor
        .update(Binding::SummaryInformation, |_| Ok(()))
        .unwrap();
    let rewritten = editor.finish().unwrap();
    let mut ole = OleFile::open(Cursor::new(rewritten)).unwrap();
    let reparsed = PropertySetReader::property_set(&mut ole, Binding::SummaryInformation).unwrap();
    assert!(reparsed.tolerance_report().is_clean());
    assert_eq!(reparsed, parsed);
}

#[test]
fn unterminated_dictionary_names_are_read_to_their_declared_length() {
    let custom = Guid::from_bytes([0x5a; 16]);
    let dictionary = |name: &[u8]| {
        let mut body = u32le(1).to_vec();
        body.extend_from_slice(&u32le(PID_TITLE));
        body.extend_from_slice(&u32le(u32::try_from(name.len()).unwrap()));
        body.extend_from_slice(name);
        while body.len() % 4 != 0 {
            body.push(0);
        }
        (0, body)
    };
    let stream = |name: &[u8]| {
        one_section_stream(
            custom,
            &[
                dictionary(name),
                codepage_property(1252),
                (PID_TITLE, typed(VT_I4, &u32le(7))),
            ],
        )
    };
    let clean = Stream::parse(&stream(b"Name\0")).unwrap();
    assert!(clean.tolerance_report().is_clean());

    let parsed = Stream::parse(&stream(b"Name")).unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [(StringDefect::MissingTerminator.into(), 0)]
    );
    let section = parsed.section(custom).unwrap();
    assert_eq!(section.property_name(PID_TITLE), Some("Name"));
    assert_eq!(parsed, clean);
}

// --- F7.11: one malformed property fails the stream; nested padding ---

#[test]
fn nested_reserved_words_and_trailing_padding_are_ignored_and_recorded() {
    // A VT_VECTOR | VT_VARIANT of one VT_I4 whose nested reserved word is
    // nonzero, then a VT_I4 property followed by nonzero trailing bytes.
    let vector = |reserved: u16| {
        let mut body = u32le(1).to_vec();
        body.extend_from_slice(&u16le(VT_I4));
        body.extend_from_slice(&u16le(reserved));
        body.extend_from_slice(&u32le(7));
        typed(VT_VECTOR | VT_VARIANT, &body)
    };
    let padded = |tail: [u8; 4]| {
        let mut value = typed(VT_I4, &u32le(9));
        value.extend_from_slice(&tail);
        value
    };
    let clean = summary_stream(&[
        codepage_property(1252),
        (0x20, vector(0)),
        (PID_SUBJECT, padded([0; 4])),
    ]);
    let clean = Stream::parse(&clean).unwrap();
    assert!(clean.tolerance_report().is_clean());

    let faulted = summary_stream(&[
        codepage_property(1252),
        (0x20, vector(0xbeef)),
        (PID_SUBJECT, padded([1, 2, 3, 4])),
    ]);
    let parsed = Stream::parse(&faulted).unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [
            (PropertyDefect::NonZeroPadding.into(), 0x20),
            (PropertyDefect::NonZeroPadding.into(), PID_SUBJECT),
        ]
    );
    assert_eq!(parsed, clean);
    assert!(parsed.tolerance_report().preserves_content());
}

#[test]
fn one_malformed_property_is_dropped_and_the_rest_stay_readable() {
    // PID 3 claims a 4 GiB string; PID 2 and PID 4 are well formed. PID 3
    // also has a dictionary name, which is dropped with it.
    let custom = Guid::from_bytes([0x6b; 16]);
    let dictionary = {
        let mut body = u32le(2).to_vec();
        for (pid, name) in [(PID_TITLE, b"Two\0"), (PID_SUBJECT, b"Bad\0")] {
            body.extend_from_slice(&u32le(pid));
            body.extend_from_slice(&u32le(4));
            body.extend_from_slice(name);
        }
        (0, body)
    };
    let mut overrun = typed(VT_LPSTR, &u32le(u32::MAX));
    overrun.extend_from_slice(b"xx\0\0");
    let parsed = Stream::parse(&one_section_stream(
        custom,
        &[
            dictionary,
            codepage_property(1252),
            (PID_TITLE, lpstr(b"kept\0")),
            (PID_SUBJECT, overrun),
            (4, typed(VT_I4, &u32le(5))),
        ],
    ))
    .unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [(PropertyDefect::MalformedValue.into(), PID_SUBJECT)]
    );
    let section = parsed.section(custom).unwrap();
    assert_eq!(
        section.property(PID_TITLE),
        Some(&Value::Lpstr("kept".into()))
    );
    assert_eq!(section.property(PID_SUBJECT), None);
    assert_eq!(section.property(4), Some(&Value::I4(5)));
    assert_eq!(section.property_name(PID_TITLE), Some("Two"));
    assert_eq!(section.property_name(PID_SUBJECT), None);
    assert!(!parsed.tolerance_report().preserves_content());

    // A dropped property is content loss: the editor will not write the
    // degraded section back over the source.
    let summary = summary_stream(&[
        codepage_property(1252),
        (PID_TITLE, lpstr(b"kept\0")),
        (PID_SUBJECT, {
            let mut value = typed(VT_LPSTR, &u32le(u32::MAX));
            value.extend_from_slice(b"xx\0\0");
            value
        }),
    ]);
    let mut ole = OleFile::open(Cursor::new(compound_file(&summary))).unwrap();
    let metadata = PropertySetReader::get_metadata(&mut ole).unwrap();
    assert_eq!(metadata.title.as_deref(), Some("kept"));
    assert_eq!(metadata.subject, None);
    let mut editor = Editor::new(compound_file(&summary)).unwrap();
    assert!(
        editor
            .update(Binding::SummaryInformation, |_| Ok(()))
            .is_err()
    );
}

// --- F7.12: stream header exactness ---

#[test]
fn section_size_overrunning_its_implied_range_is_clamped_and_recorded() {
    let properties = [codepage_property(1252), (PID_TITLE, lpstr(b"Sized\0"))];
    let clean = summary_stream(&properties);
    let expected = Stream::parse(&clean).unwrap();
    assert!(expected.tolerance_report().is_clean());

    // A padding-aware writer rounded the section size past the stream end.
    let mut rounded = clean.clone();
    let size = u32::from_le_bytes(rounded[48..52].try_into().unwrap());
    rounded[48..52].copy_from_slice(&u32le(size + 64));
    let parsed = Stream::parse(&rounded).unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [(StreamDefect::SectionSizeMismatch.into(), 0)]
    );
    assert_eq!(parsed, expected);
    assert!(parsed.tolerance_report().preserves_content());
}

#[test]
fn first_section_size_overrunning_the_second_is_clamped_to_it() {
    let first = section_body(&[codepage_property(1252), (PID_TITLE, lpstr(b"One\0"))]);
    let second_id = Guid::from_bytes([0x7c; 16]);
    let second = section_body(&[
        codepage_property(1252),
        (PID_TITLE, typed(VT_I4, &u32le(2))),
    ]);
    let build = |first_size: u32| {
        let mut data = Vec::new();
        data.extend_from_slice(&u16le(0xfffe));
        data.extend_from_slice(&u16le(0));
        data.extend_from_slice(&u32le(0));
        data.extend_from_slice(&[0; 16]);
        data.extend_from_slice(&u32le(2));
        data.extend_from_slice(SUMMARY_INFORMATION_FMTID.as_bytes());
        data.extend_from_slice(&u32le(68));
        data.extend_from_slice(second_id.as_bytes());
        data.extend_from_slice(&u32le(u32::try_from(68 + first.len()).unwrap()));
        let mut body = first.clone();
        body[0..4].copy_from_slice(&u32le(first_size));
        data.extend_from_slice(&body);
        data.extend_from_slice(&second);
        data
    };
    let exact = u32::try_from(first.len()).unwrap();
    let clean = Stream::parse(&build(exact)).unwrap();
    assert!(clean.tolerance_report().is_clean());

    let parsed = Stream::parse(&build(exact + 16)).unwrap();
    assert_eq!(
        recorded(parsed.tolerance_report()),
        [(StreamDefect::SectionSizeMismatch.into(), 0)]
    );
    assert_eq!(parsed, clean);
    assert_eq!(
        parsed.section(second_id).unwrap().property(PID_TITLE),
        Some(&Value::I4(2))
    );
}

/// Version and section-count deviations stay errors: POI rejects a nonzero
/// format word outright, and a newer version or a third section may carry
/// semantics this grammar cannot represent, so repairing them would be a
/// judgment call (the audit's opt-in class), not a default repair.
#[test]
fn unsupported_versions_and_section_counts_stay_errors() {
    let clean = summary_stream(&[codepage_property(1252)]);
    let mut version = clean.clone();
    version[2..4].copy_from_slice(&u16le(2));
    assert!(Stream::parse(&version).is_err());
    let mut count = clean;
    count[24..28].copy_from_slice(&u32le(3));
    assert!(Stream::parse(&count).is_err());
}
