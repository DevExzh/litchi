//! Crate-private XML markup helpers shared by the tolerant readers.

use std::borrow::Cow;

use litchi_core::xml::ReaderOrigin;
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::{Error, Result};

/// `xml` without its processing instructions, and whether it had any.
///
/// Processing instructions are legal XML 1.0 (§2.6) that carry no OOXML
/// content (audit finding 6.3), but the shared markup-compatibility
/// processor and several feature codecs refuse them. A tolerant reader runs
/// its semantic parse over this output and records
/// [`crate::MarkupDefect::ProcessingInstruction`] when the flag is set; the
/// source bytes themselves are never changed. The XML declaration is not a
/// processing instruction and is kept.
///
/// A part without `<?` past its declaration is returned borrowed without
/// being parsed.
///
/// # Errors
///
/// Returns an XML error if the scan meets ill-formed XML, and an allocation
/// error if the filtered copy cannot be reserved.
pub(crate) fn without_processing_instructions(xml: &[u8]) -> Result<(Cow<'_, [u8]>, bool)> {
    if !may_contain_processing_instruction(xml) {
        return Ok((Cow::Borrowed(xml), false));
    }
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut filtered: Option<Vec<u8>> = None;
    let mut copied = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = position(&origin, &reader)?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let is_instruction = matches!(event, Event::PI(_));
        let finished = matches!(event, Event::Eof);
        if is_instruction {
            let after = position(&origin, &reader)?;
            let output = match filtered.as_mut() {
                Some(output) => output,
                None => {
                    let mut output = Vec::new();
                    output
                        .try_reserve_exact(xml.len())
                        .map_err(|source| Error::Allocation {
                            resource: "processing-instruction-free XML",
                            source,
                        })?;
                    filtered.insert(output)
                },
            };
            output.extend_from_slice(&xml[copied..before]);
            copied = after;
        }
        if finished {
            break;
        }
        buffer.clear();
    }
    Ok(match filtered {
        Some(mut output) => {
            output.extend_from_slice(&xml[copied..]);
            (Cow::Owned(output), true)
        },
        None => (Cow::Borrowed(xml), false),
    })
}

fn position(origin: &ReaderOrigin, reader: &Reader<&[u8]>) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("processing-instruction offset does not fit usize".into()))
}

/// Whether `xml` has a `<?` other than a leading XML declaration.
fn may_contain_processing_instruction(xml: &[u8]) -> bool {
    let mut search = 0;
    while let Some(offset) = xml[search..].windows(2).position(|pair| pair == b"<?") {
        let at = search + offset;
        let is_declaration = xml[..at]
            .iter()
            .all(|byte| byte.is_ascii_whitespace() || matches!(byte, 0xEF | 0xBB | 0xBF))
            && xml[at..].starts_with(b"<?xml")
            && xml
                .get(at + 5)
                .is_some_and(|byte| byte.is_ascii_whitespace() || *byte == b'?');
        if !is_declaration {
            return true;
        }
        search = at + 2;
    }
    false
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::without_processing_instructions;

    #[test]
    fn declaration_only_parts_are_borrowed() {
        let xml = br#"<?xml version="1.0"?><a x="1"><b/></a>"#;
        let (output, found) = without_processing_instructions(xml).unwrap();
        assert!(!found);
        assert!(matches!(output, std::borrow::Cow::Borrowed(_)));
    }

    #[test]
    fn instructions_are_removed_and_reported() {
        let xml = br#"<?xml version="1.0"?><?pub a?><a><?x?>t<!--<?no?>--><![CDATA[<?no?>]]></a>"#;
        let (output, found) = without_processing_instructions(xml).unwrap();
        assert!(found);
        assert_eq!(
            output.as_ref(),
            br#"<?xml version="1.0"?><a>t<!--<?no?>--><![CDATA[<?no?>]]></a>"#
        );
        let (output, found) = without_processing_instructions(b"<a><!--<?no?>--></a>").unwrap();
        assert!(!found);
        assert_eq!(output.as_ref(), b"<a><!--<?no?>--></a>");
    }
}
