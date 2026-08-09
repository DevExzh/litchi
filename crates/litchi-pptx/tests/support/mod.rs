use quick_xml::Reader;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};

/// Remove element-only formatting whitespace from a valid positive XML fixture.
///
/// Keep malformed and whitespace-sensitive inputs out of this helper: those
/// belong in raw parser/ZIP tests, not in packages sent through the production
/// writer.
pub(crate) fn compact_xml_fixture(source: &[u8]) -> Vec<u8> {
    assert!(
        !source
            .windows(b"xml:space".len())
            .any(|window| window == b"xml:space"),
        "whitespace-sensitive XML fixtures must be compacted explicitly"
    );

    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut output = Vec::with_capacity(source.len());
    loop {
        let event = reader
            .read_event()
            .expect("positive XML fixture must parse");
        match event {
            Event::Start(element) => {
                write_compact_start(&mut output, &element, reader.decoder(), false)
            },
            Event::Empty(element) => {
                write_compact_start(&mut output, &element, reader.decoder(), true)
            },
            Event::End(element) => {
                output.extend_from_slice(b"</");
                output.extend_from_slice(element.name().as_ref());
                output.push(b'>');
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Text(text) => output.extend_from_slice(text.as_ref()),
            Event::CData(data) => {
                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(data.as_ref());
                output.extend_from_slice(b"]]");
                output.push(b'>');
            },
            Event::GeneralRef(reference) => {
                output.push(b'&');
                output.extend_from_slice(reference.as_ref());
                output.push(b';');
            },
            Event::Decl(declaration) => {
                output.extend_from_slice(b"<?");
                output.extend_from_slice(declaration.as_ref());
                output.extend_from_slice(b"?>");
            },
            Event::PI(instruction) => {
                output.extend_from_slice(b"<?");
                output.extend_from_slice(instruction.as_ref());
                output.extend_from_slice(b"?>");
            },
            Event::Comment(comment) => {
                output.extend_from_slice(b"<!--");
                output.extend_from_slice(comment.as_ref());
                output.extend_from_slice(b"-->");
            },
            Event::Eof => break,
            Event::DocType(_) => panic!("positive package-writer fixtures must not contain DTDs"),
        }
    }
    output
}

fn write_compact_start(
    output: &mut Vec<u8>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    empty: bool,
) {
    output.push(b'<');
    output.extend_from_slice(element.name().as_ref());
    for attribute in element.attributes() {
        let attribute = attribute.expect("positive XML fixture attributes must parse");
        output.push(b' ');
        output.extend_from_slice(attribute.key.as_ref());
        output.extend_from_slice(b"=\"");
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .expect("positive XML fixture attributes must decode");
        output.extend_from_slice(quick_xml::escape::escape(value.as_ref()).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
}
