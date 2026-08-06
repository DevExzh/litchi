use super::{CharacterFormatting, HeaderFooterParagraph, WriteError, utf16_code_unit_len};

pub(in crate::writer::core) const MAX_HEADER_FOOTER_PARAGRAPHS: usize = 65_535;
pub(in crate::writer::core) const MAX_HEADER_FOOTER_RUNS: usize = 65_535;
pub(in crate::writer::core) const MAX_HEADER_FIELD_DEPTH: usize = 128;

#[derive(Default)]
pub(in crate::writer::core) struct HeaderFieldState {
    pub(in crate::writer::core) separator_seen: Vec<bool>,
}

impl HeaderFieldState {
    pub(in crate::writer::core) fn observe(
        &mut self,
        character: char,
        formatting: &CharacterFormatting,
    ) -> Result<bool, WriteError> {
        if !matches!(character as u32, 0x0013..=0x0015) {
            return Ok(false);
        }
        if formatting.special != Some(true) {
            return Err(WriteError::InvalidData(
                "DOC header/footer field marker requires fSpec formatting".to_string(),
            ));
        }
        match character as u32 {
            0x0013 => {
                if self.separator_seen.len() >= MAX_HEADER_FIELD_DEPTH {
                    return Err(WriteError::InvalidData(
                        "DOC header/footer field nesting exceeds the limit".to_string(),
                    ));
                }
                self.separator_seen.push(false);
            },
            0x0014 => {
                let seen = self.separator_seen.last_mut().ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC header/footer field separator has no begin marker".to_string(),
                    )
                })?;
                if *seen {
                    return Err(WriteError::InvalidData(
                        "DOC header/footer field has duplicate separators".to_string(),
                    ));
                }
                *seen = true;
            },
            0x0015 => {
                self.separator_seen.pop().ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC header/footer field end has no begin marker".to_string(),
                    )
                })?;
            },
            _ => unreachable!(),
        }
        Ok(true)
    }

    pub(in crate::writer::core) fn finish(self) -> Result<(), WriteError> {
        if self.separator_seen.is_empty() {
            Ok(())
        } else {
            Err(WriteError::InvalidData(
                "DOC header/footer field is not terminated within its story".to_string(),
            ))
        }
    }
}

pub(in crate::writer::core) fn checked_text_fc(
    text_fc_start: u32,
    stream_length: usize,
) -> Result<u32, WriteError> {
    let stream_length = u32::try_from(stream_length).map_err(|_| {
        WriteError::InvalidData("DOC text stream exceeds 32-bit FC space".to_string())
    })?;
    text_fc_start
        .checked_add(stream_length)
        .ok_or_else(|| WriteError::InvalidData("DOC text stream FC range overflows".to_string()))
}

pub(in crate::writer::core) fn validate_header_footer_paragraphs(
    paragraphs: &[HeaderFooterParagraph],
) -> Result<(), WriteError> {
    if paragraphs.is_empty() {
        return Err(WriteError::InvalidData(
            "DOC header/footer story requires at least one paragraph".to_string(),
        ));
    }
    if paragraphs.len() > MAX_HEADER_FOOTER_PARAGRAPHS {
        return Err(WriteError::InvalidData(
            "DOC header/footer story exceeds the paragraph limit".to_string(),
        ));
    }

    let mut run_count = 0usize;
    let mut character_count = 1u32; // Inter-story guard paragraph mark.
    let mut field_state = HeaderFieldState::default();
    for paragraph in paragraphs {
        run_count = run_count.checked_add(paragraph.runs.len()).ok_or_else(|| {
            WriteError::InvalidData("DOC header/footer run count overflows".to_string())
        })?;
        if run_count > MAX_HEADER_FOOTER_RUNS {
            return Err(WriteError::InvalidData(
                "DOC header/footer story exceeds the run limit".to_string(),
            ));
        }
        for (text, formatting) in &paragraph.runs {
            if text.contains('\r') {
                return Err(WriteError::InvalidData(
                    "DOC header/footer run contains an embedded paragraph mark".to_string(),
                ));
            }
            for character in text.chars() {
                field_state.observe(character, formatting)?;
            }
            character_count = character_count
                .checked_add(utf16_code_unit_len(text)?)
                .ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC header/footer story CP range overflows".to_string(),
                    )
                })?;
        }
        character_count = character_count.checked_add(1).ok_or_else(|| {
            WriteError::InvalidData("DOC header/footer story CP range overflows".to_string())
        })?;
    }
    if character_count >= 0x7FFF_FFFF {
        return Err(WriteError::InvalidData(
            "DOC header/footer story exceeds the MS-DOC CP limit".to_string(),
        ));
    }
    field_state.finish()
}
