//! PPT ClientTextbox and text-interaction record assembly.

use litchi_odraw::write::Header;
use zerocopy::IntoBytes;

use super::super::Error;

/// Builds a plain-text ClientTextbox record.
pub(crate) fn build_client_textbox(text: &str, text_type: u32) -> Result<Vec<u8>, Error> {
    build_client_textbox_with_interactions(text, text_type, &[])
}

pub(crate) fn build_client_textbox_with_interactions(
    text: &str,
    text_type: u32,
    interactions: &[crate::TextInteraction],
) -> Result<Vec<u8>, Error> {
    use crate::writer::records::{RecordBuilder, record_type as ppt_rt};

    let mut result = Vec::new();
    let mut ppt_content = Vec::new();

    let mut text_header = RecordBuilder::new(0, 0, ppt_rt::TEXT_HEADER_ATOM);
    text_header.write_data(&text_type.to_le_bytes());
    ppt_content.extend_from_slice(&text_header.build()?);

    if text.is_ascii() {
        let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_BYTES_ATOM);
        text_atom.write_data(text.as_bytes());
        ppt_content.extend_from_slice(&text_atom.build()?);
    } else {
        let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_CHARS_ATOM);
        for ch in text.encode_utf16() {
            text_atom.write_data(&ch.to_le_bytes());
        }
        ppt_content.extend_from_slice(&text_atom.build()?);
    }

    let too_large = || {
        Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox text exceeds the PPT size limit",
        )
    };
    let text_units = u32::try_from(text.encode_utf16().count()).map_err(|_| too_large())?;
    let char_count = text_units.checked_add(1).ok_or_else(too_large)?;
    let mut style_atom = RecordBuilder::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM);
    style_atom.write_data(&char_count.to_le_bytes());
    style_atom.write_data(&0u16.to_le_bytes());
    style_atom.write_data(&0u32.to_le_bytes());
    style_atom.write_data(&char_count.to_le_bytes());
    style_atom.write_data(&0u32.to_le_bytes());
    ppt_content.extend_from_slice(&style_atom.build()?);
    append_text_interactions(
        &mut ppt_content,
        text_units,
        interactions,
        crate::TextInteractionLimits::default(),
    )?;

    let header = Header::new(0x0F, 0, 0xF00D, ppt_content.len() as u32);
    result.extend_from_slice(header.as_bytes());
    result.extend_from_slice(&ppt_content);
    Ok(result)
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
    let text_units = u32::try_from(text_chars.len() / 2).map_err(|_| {
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

    let header = Header::new(0x0F, 0, 0xF00D, ppt_content.len() as u32);
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
