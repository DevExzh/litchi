//! Section table (PLCF of SEDs) generation for DOC files
//!
//! The section table defines document sections and page layout.
//! Based on Microsoft's "[MS-DOC]" specification Section 2.9.245 and
//! Apache POI's SectionTable implementation.

use crate::section::borders;
use crate::section::borders::Borders;
use crate::section::columns::{Layout, WireView};

/// Generate SEPX (Section Properties) structure with optional first/odd-even header/footer flags
///
/// - When `first_page_header` is true, emits `sprmSFTitlePage` to enable different first page.
/// - When `grpf_ihdt` != 0, emits `sprmSGprfIhdt` to declare which headers/footers exist in this section.
///   Bits follow LibreOffice nsHdFtFlags/Word semantics:
///   0x01=HeaderEven, 0x02=HeaderOdd, 0x04=FooterEven, 0x08=FooterOdd, 0x10=HeaderFirst, 0x20=FooterFirst
pub fn generate_sepx(first_page_header: bool, grpf_ihdt: u8) -> Vec<u8> {
    generate_sepx_with_revision(first_page_header, grpf_ihdt, None)
}

/// Generate SEPX properties with an optional section property revision mark.
pub(crate) fn generate_sepx_with_revision(
    first_page_header: bool,
    grpf_ihdt: u8,
    revision: Option<(u16, u32)>,
) -> Vec<u8> {
    generate_sepx_with_properties(
        first_page_header,
        grpf_ihdt,
        revision,
        None,
        false,
        crate::TextFlow::HorizontalNonAsian,
        None,
    )
    .expect("default section properties are valid")
}

/// Generate Word 97+ SEPX properties for the writer's single section.
pub(crate) fn generate_sepx_with_properties(
    first_page_header: bool,
    grpf_ihdt: u8,
    revision: Option<(u16, u32)>,
    columns: Option<&Layout>,
    right_to_left: bool,
    text_flow: crate::TextFlow,
    page_borders: Option<&Borders>,
) -> Result<Vec<u8>, String> {
    let mut grpprl: Vec<u8> = Vec::with_capacity(8);
    if first_page_header {
        // sprmSFTitlePage (u16 opcode) + 1-byte operand (1)
        grpprl.extend_from_slice(&crate::sprm_operations::SPRM_S_F_TITLE_PAGE.to_le_bytes());
        grpprl.push(1u8);
    }
    if grpf_ihdt != 0 {
        // sprmSGprfIhdt (u16 opcode) + 1-byte operand (bitfield)
        grpprl.extend_from_slice(&crate::sprm_operations::SPRM_S_GPRF_IHDT.to_le_bytes());
        grpprl.push(grpf_ihdt);
    }
    if let Some(columns) = columns {
        columns.validate().map_err(|error| error.to_string())?;
        let count_minus_one = columns
            .count()
            .checked_sub(1)
            .and_then(|count| u16::try_from(count).ok())
            .ok_or_else(|| {
                "validated section column count does not fit the SEPX field".to_string()
            })?;
        push_word(
            &mut grpprl,
            crate::sprm_operations::SPRM_S_C_COLUMNS,
            count_minus_one,
        );
        match columns.wire_view() {
            WireView::Even {
                spacing_twips,
                line_between,
                ..
            } => {
                push_bool(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_F_EVENLY_SPACED,
                    true,
                );
                push_word(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_DXA_COLUMNS,
                    spacing_twips,
                );
                push_bool(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_L_BETWEEN,
                    line_between,
                );
            },
            WireView::Unequal {
                columns,
                line_between,
            } => {
                push_bool(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_F_EVENLY_SPACED,
                    false,
                );
                for (index, column) in columns.iter().enumerate() {
                    push_indexed_twips(
                        &mut grpprl,
                        crate::sprm_operations::SPRM_S_DXA_COL_WIDTH,
                        index,
                        column.width_twips(),
                    )?;
                    if let Some(spacing) = column.spacing_after_twips() {
                        push_indexed_twips(
                            &mut grpprl,
                            crate::sprm_operations::SPRM_S_DXA_COL_SPACING,
                            index,
                            spacing,
                        )?;
                    }
                }
                push_bool(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_L_BETWEEN,
                    line_between,
                );
            },
        }
    }
    if right_to_left {
        push_bool(&mut grpprl, crate::sprm_operations::SPRM_S_F_BIDI, true);
    }
    if let Some(page_borders) = page_borders {
        borders::encode_sepx(&mut grpprl, page_borders).map_err(|error| error.to_string())?;
    }
    if text_flow != crate::TextFlow::HorizontalNonAsian {
        let value = match text_flow {
            crate::TextFlow::HorizontalNonAsian => 0,
            crate::TextFlow::TopToBottomAsian => 1,
            crate::TextFlow::BottomToTop => 2,
            crate::TextFlow::TopToBottomNonAsian => 3,
            crate::TextFlow::HorizontalAsian => 4,
            crate::TextFlow::VerticalNonAsian => 5,
        };
        push_word(&mut grpprl, crate::sprm_operations::SPRM_S_TEXT_FLOW, value);
    }
    if let Some((author_index, timestamp)) = revision {
        // sprmSPropRMark + PropRMarkOperand(cb=7, active, ibstshort, DTTM)
        grpprl.extend_from_slice(&0xD243u16.to_le_bytes());
        grpprl.push(7);
        grpprl.push(1);
        grpprl.extend_from_slice(&author_index.to_le_bytes());
        grpprl.extend_from_slice(&timestamp.to_le_bytes());
    }
    let size = grpprl.len() as u16;
    let mut sepx = Vec::with_capacity(2 + grpprl.len());
    sepx.extend_from_slice(&size.to_le_bytes());
    sepx.extend_from_slice(&grpprl);
    Ok(sepx)
}

fn push_bool(output: &mut Vec<u8>, opcode: u16, value: bool) {
    output.extend_from_slice(&opcode.to_le_bytes());
    output.push(u8::from(value));
}

fn push_word(output: &mut Vec<u8>, opcode: u16, value: u16) {
    output.extend_from_slice(&opcode.to_le_bytes());
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_indexed_twips(
    output: &mut Vec<u8>,
    opcode: u16,
    index: usize,
    value: u16,
) -> Result<(), String> {
    let index = u8::try_from(index)
        .map_err(|_| "validated section column index does not fit the SEPX field".to_string())?;
    output.extend_from_slice(&opcode.to_le_bytes());
    output.push(index);
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

/// Generate minimal SEPX (Section Properties) structure (no section SPRMs)
#[inline]
pub fn generate_minimal_sepx() -> Vec<u8> {
    generate_sepx(false, 0)
}

/// Generate section table (PLCF of SEDs)
///
/// Creates a single section covering the entire document
///
/// # Arguments
///
/// * `text_length` - Total length of document text in characters  
/// * `sepx_offset` - Offset in WordDocument stream where SEPX was written
pub fn generate_section_table(text_length: u32, sepx_offset: u32) -> Vec<u8> {
    let mut plcfsed = Vec::new();

    // PLCF structure (Apache POI's PlexOfCps):
    // - Array of n+1 CPs (character positions)
    // - Array of n data elements (SEDs)

    // We have 1 section, so we need 2 CPs

    // CP[0] = 0 (start of document)
    plcfsed.extend_from_slice(&0u32.to_le_bytes());

    // CP[1] = text_length (end of document)
    plcfsed.extend_from_slice(&text_length.to_le_bytes());

    // SED (Section Descriptor) - 12 bytes (POI's SectionDescriptor.toByteArray())

    // fn (short) - used internally by Word - 0 for new documents
    plcfsed.extend_from_slice(&0u16.to_le_bytes());

    // fcSepx (int) - CRITICAL: Must point to SEPX in WordDocument stream
    // Apache POI sets this to the offset where SEPX was written (line 195)
    plcfsed.extend_from_slice(&sepx_offset.to_le_bytes());

    // fnMpr (short) - used internally - 0
    plcfsed.extend_from_slice(&0u16.to_le_bytes());

    // fcMpr (int) - Mac print record offset - 0
    plcfsed.extend_from_slice(&0u32.to_le_bytes());

    plcfsed
}

#[cfg(test)]
mod tests {
    use super::generate_sepx_with_properties;

    #[test]
    fn page_border_default_is_omitted_and_nondefault_is_canonical() {
        let default = crate::section::borders::Borders::default();
        assert_eq!(
            generate_sepx_with_properties(
                false,
                0,
                None,
                None,
                false,
                crate::TextFlow::HorizontalNonAsian,
                Some(&default),
            )
            .unwrap(),
            [0, 0]
        );

        let borders = crate::section::borders::Borders {
            top: Some(crate::section::borders::Border {
                style: crate::section::borders::Style::Single,
                width_eighth_points: 8,
                color: crate::section::borders::Color::Red,
                spacing_points: 3,
                shadow: true,
                frame: false,
            }),
            apply_to: crate::section::borders::ApplyTo::FirstPage,
            ..default
        };
        assert_eq!(
            generate_sepx_with_properties(
                false,
                0,
                None,
                None,
                false,
                crate::TextFlow::HorizontalNonAsian,
                Some(&borders),
            )
            .unwrap(),
            [10, 0, 0x2B, 0x70, 8, 1, 6, 0x23, 0x2F, 0x52, 1, 0]
        );
    }
}
