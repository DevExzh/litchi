//! PPT `ClientTextbox` and text-interaction record assembly.

use litchi_odraw::write::Header;
use zerocopy::IntoBytes;

use super::super::Error;

/// Builds a plain-text `ClientTextbox` record.
pub(crate) fn build_client_textbox(text: &str, text_type: u32) -> Result<Vec<u8>, Error> {
    build_client_textbox_with_interactions(text, text_type, &[])
}

pub(crate) fn build_client_textbox_with_interactions(
    text: &str,
    text_type: u32,
    interactions: &[crate::TextInteraction],
) -> Result<Vec<u8>, Error> {
    let mut record = Vec::new();
    append_client_textbox_with_interactions(&mut record, text, text_type, interactions)?;
    Ok(record)
}

/// Bytes of a plain-text `ClientTextbox` other than its text atom's body: the
/// `OfficeArt` header, `TextHeaderAtom`, the text atom's header, and
/// `StyleTextPropAtom`.
const PLAIN_TEXTBOX_FIXED_BYTES: usize = 8 + (8 + 4) + 8 + (8 + 18);

/// Appends a plain-text `ClientTextbox` record to `output` in one pass.
///
/// The record is written in place and its container length patched once its
/// payload is complete, so the text is copied once, straight into `output`.
/// ASCII text becomes a `TextBytesAtom` and is counted by its byte length
/// without decoding; other text is encoded once into a `TextCharsAtom` and
/// counted from the encoded length. The bytes, the refusals and their order
/// are those of building each atom separately: the text is encoded before it
/// is counted, and the count is checked before the style atom and the text
/// interactions. On error `output` holds a partial record and must be
/// discarded.
pub(crate) fn append_client_textbox_with_interactions(
    output: &mut Vec<u8>,
    text: &str,
    text_type: u32,
    interactions: &[crate::TextInteraction],
) -> Result<(), Error> {
    use crate::writer::records::{InPlaceRecord, RecordHeader, record_type as ppt_rt};

    let too_large = || {
        Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox text exceeds the PPT size limit",
        )
    };
    let ascii = text.is_ascii();
    // Exact for ASCII; for other text an upper bound, because a UTF-8 byte
    // never yields more than one UTF-16 code unit.
    let text_bytes = if ascii {
        text.len()
    } else {
        text.len().saturating_mul(2)
    };
    output.reserve(PLAIN_TEXTBOX_FIXED_BYTES.saturating_add(text_bytes));

    let start = output.len();
    output.extend_from_slice(&[0; 8]);

    RecordHeader::new(0, 0, ppt_rt::TEXT_HEADER_ATOM, 4).write(output)?;
    output.extend_from_slice(&text_type.to_le_bytes());

    let text_units = if ascii {
        let atom = InPlaceRecord::begin(output, 0, 0, ppt_rt::TEXT_BYTES_ATOM);
        output.extend_from_slice(text.as_bytes());
        atom.finish(output)?;
        text.len()
    } else {
        let atom = InPlaceRecord::begin(output, 0, 0, ppt_rt::TEXT_CHARS_ATOM);
        let body_start = output.len();
        for unit in text.encode_utf16() {
            output.extend_from_slice(&unit.to_le_bytes());
        }
        let units = (output.len() - body_start) / 2;
        atom.finish(output)?;
        units
    };
    let text_units = u32::try_from(text_units).map_err(|_err| too_large())?;
    let char_count = text_units.checked_add(1).ok_or_else(too_large)?;

    RecordHeader::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM, 18).write(output)?;
    output.extend_from_slice(&char_count.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&char_count.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    append_text_interactions(
        output,
        text_units,
        interactions,
        crate::TextInteractionLimits::default(),
    )?;

    let content = u32::try_from(output.len() - start - 8).map_err(|_err| too_large())?;
    let header = Header::new(0x0F, 0, 0xF00D, content);
    output
        .get_mut(start..start + 8)
        .ok_or_else(|| std::io::Error::other("ClientTextbox header is outside its output"))?
        .copy_from_slice(header.as_bytes());
    Ok(())
}

#[cfg(test)]
pub(crate) fn build_client_textbox_formatted(
    paragraphs: &[crate::writer::text_format::Paragraph],
    text_type: u32,
) -> Result<Vec<u8>, Error> {
    build_client_textbox_formatted_with_interactions(paragraphs, text_type, &[])
}

pub(super) fn build_client_textbox_formatted_with_interactions(
    paragraphs: &[crate::writer::text_format::Paragraph],
    text_type: u32,
    interactions: &[crate::TextInteraction],
) -> Result<Vec<u8>, Error> {
    use crate::writer::records::{RecordBuilder, record_type as ppt_rt};
    use crate::writer::text_format::TextPropsBuilder;

    let mut result = Vec::new();
    let mut ppt_content = Vec::new();

    let mut text_header = RecordBuilder::new(0, 0, ppt_rt::TEXT_HEADER_ATOM);
    text_header.write_data(&text_type.to_le_bytes());
    ppt_content.extend_from_slice(&text_header.build()?);

    let mut builder = TextPropsBuilder::new();
    for para in paragraphs {
        builder.add_paragraph(para.clone());
    }

    let text_chars = builder.build_text_chars();
    let text_units = u32::try_from(text_chars.len() / 2).map_err(|_err| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox text exceeds the PPT size limit",
        )
    })?;
    let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_CHARS_ATOM);
    text_atom.write_data(&text_chars);
    ppt_content.extend_from_slice(&text_atom.build()?);

    let style_data = builder.build_style_text_prop()?;
    let mut style_atom = RecordBuilder::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM);
    style_atom.write_data(&style_data);
    ppt_content.extend_from_slice(&style_atom.build()?);
    append_text_interactions(
        &mut ppt_content,
        text_units,
        interactions,
        crate::TextInteractionLimits::default(),
    )?;

    let header = Header::new(
        0x0F,
        0,
        0xF00D,
        u32::try_from(ppt_content.len()).map_err(|_err| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "ClientTextbox text exceeds the PPT size limit",
            )
        })?,
    );
    result.extend_from_slice(header.as_bytes());
    result.extend_from_slice(&ppt_content);
    Ok(result)
}

fn append_text_interactions(
    output: &mut Vec<u8>,
    text_units: u32,
    interactions: &[crate::TextInteraction],
    limits: crate::TextInteractionLimits,
) -> Result<(), Error> {
    if interactions.len() > limits.max_interactions {
        return Err(Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox exceeds the text interaction count limit",
        ));
    }
    for interaction in interactions {
        output.extend_from_slice(
            &interaction
                .to_bytes_for_text(text_units, limits)
                .map_err(|error| std::io::Error::other(error.to_string()))?,
        );
    }
    Ok(())
}
