#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "contract assertions intentionally panic on reader errors"
)]

//! Change 0765: `litchi_core::xml::ReaderOrigin` models how the `quick-xml`
//! build this workspace links reports positions for an input that begins with
//! a UTF-8 byte-order mark. Every OOXML crate converts reader positions to
//! byte offsets through it, so this test pins the model to the real reader: if
//! a `quick-xml` upgrade or a feature change (for example enabling
//! `encoding`, which also consumes UTF-16 marks) moved the positions, the
//! offsets would silently address the wrong bytes, and this test fails first.

use std::io::BufReader;

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::reader::{NsReader, Reader};

const BOM: &[u8] = b"\xEF\xBB\xBF";

/// One reader event: the kind and the raw input bytes the positions cover.
fn slices(input: &[u8]) -> Vec<(&'static str, Vec<u8>)> {
    let origin = ReaderOrigin::of(input);
    let mut reader = Reader::from_reader(input);
    let mut events = Vec::new();
    loop {
        let start = origin.offset(reader.buffer_position()).unwrap();
        let event = reader.read_event().unwrap();
        let end = origin.offset(reader.buffer_position()).unwrap();
        let kind = match event {
            Event::Eof => break,
            Event::Start(_) => "start",
            Event::End(_) => "end",
            Event::Empty(_) => "empty",
            Event::Text(_) => "text",
            Event::Decl(_) => "decl",
            Event::Comment(_) => "comment",
            _ => "other",
        };
        events.push((kind, input[start..end].to_vec()));
    }
    events
}

#[test]
fn slice_reader_positions_address_marked_input_through_the_origin() {
    let plain = b"<?xml version=\"1.0\"?>\r\n<a x=\"1\"><!--c--><b/>t</a>";
    let mut marked = BOM.to_vec();
    marked.extend_from_slice(plain);
    assert_eq!(ReaderOrigin::of(&marked).skipped(), 3);
    assert_eq!(ReaderOrigin::of(plain).skipped(), 0);
    let expected = slices(plain);
    assert_eq!(expected[0], ("decl", b"<?xml version=\"1.0\"?>".to_vec()));
    assert_eq!(slices(&marked), expected);

    // The mark is exactly the bytes before the first event.
    let origin = ReaderOrigin::of(&marked);
    let mut reader = Reader::from_reader(marked.as_slice());
    assert_eq!(reader.buffer_position(), 0);
    let _ = reader.read_event().unwrap();
    assert_eq!(
        &marked[..origin.offset(0).unwrap()],
        BOM,
        "only the mark precedes position zero"
    );
}

#[test]
fn a_second_mark_is_character_data_not_origin() {
    let mut doubled = BOM.to_vec();
    doubled.extend_from_slice(BOM);
    doubled.extend_from_slice(b"<a/>");
    assert_eq!(
        slices(&doubled),
        vec![("text", BOM.to_vec()), ("empty", b"<a/>".to_vec())]
    );
}

#[test]
fn utf16_marks_are_not_consumed_by_this_build() {
    // Without quick-xml's `encoding` feature a UTF-16 mark is ordinary input;
    // the origin must stay zero so offsets still address the raw bytes.
    for input in [&b"\xFF\xFE<a/>"[..], b"\xFE\xFF<a/>"] {
        assert_eq!(ReaderOrigin::of(input).skipped(), 0);
        let events = slices(input);
        assert_eq!(events[0].1[..2], input[..2], "{input:?}");
    }
}

#[test]
fn error_positions_address_marked_input_through_the_origin() {
    let plain = b"<a><b></c></a>";
    let mut marked = BOM.to_vec();
    marked.extend_from_slice(plain);
    let error_offset = |input: &[u8]| {
        let origin = ReaderOrigin::of(input);
        let mut reader = Reader::from_reader(input);
        reader.config_mut().check_end_names = true;
        loop {
            match reader.read_event() {
                Ok(Event::Eof) => panic!("mismatched end tag must fail"),
                Ok(_) => {},
                Err(_) => return origin.offset(reader.error_position()).unwrap(),
            }
        }
    };
    let plain_offset = error_offset(plain);
    assert_eq!(&plain[plain_offset..plain_offset + 3], b"</c");
    assert_eq!(error_offset(&marked), plain_offset + BOM.len());
}

#[test]
fn namespace_and_buffered_readers_share_the_slice_readers_origin() {
    let plain = b"<p:a xmlns:p=\"urn:p\"><p:b>x</p:b></p:a>";
    let mut marked = BOM.to_vec();
    marked.extend_from_slice(plain);
    let origin = ReaderOrigin::of(&marked);

    let expected = slices(plain);

    let mut namespaced = NsReader::from_reader(marked.as_slice());
    // The smallest buffer whose first fill holds the whole mark.
    let mut buffered = Reader::from_reader(BufReader::with_capacity(3, marked.as_slice()));
    let mut buffer = Vec::new();
    let mut seen = Vec::new();
    loop {
        let start = origin.offset(namespaced.buffer_position()).unwrap();
        let (_, event) = namespaced.read_resolved_event().unwrap();
        let end = origin.offset(namespaced.buffer_position()).unwrap();
        let buffered_start = origin.offset(buffered.buffer_position()).unwrap();
        let buffered_event = buffered.read_event_into(&mut buffer).unwrap();
        let buffered_end = origin.offset(buffered.buffer_position()).unwrap();
        assert_eq!((start, end), (buffered_start, buffered_end));
        if matches!(event, Event::Eof) {
            assert!(matches!(buffered_event, Event::Eof));
            break;
        }
        seen.push(marked[start..end].to_vec());
        buffer.clear();
    }
    let expected: Vec<Vec<u8>> = expected.into_iter().map(|(_, bytes)| bytes).collect();
    assert_eq!(seen, expected);
}

#[test]
fn a_buffered_reader_removes_only_a_mark_its_first_fill_holds() {
    // `ReaderOrigin::of` describes slice readers, and buffered readers whose
    // first fill holds the complete mark. A buffered reader that sees only
    // part of the mark in its first fill keeps all of it as character data,
    // with positions that are already byte offsets. Streaming callers must
    // present the whole mark in the first fill or remove it themselves.
    let mut marked = BOM.to_vec();
    marked.extend_from_slice(b"<a/>");
    let mut reader = Reader::from_reader(BufReader::with_capacity(2, marked.as_slice()));
    let mut buffer = Vec::new();
    let event = reader.read_event_into(&mut buffer).unwrap();
    assert!(matches!(event, Event::Text(ref text) if text.as_ref() == BOM));
    assert_eq!(reader.buffer_position(), 3);
}
