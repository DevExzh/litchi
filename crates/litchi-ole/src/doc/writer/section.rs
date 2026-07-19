//! Section table (PLCF of SEDs) generation for DOC files
//!
//! The section table defines document sections and page layout.
//! Based on Microsoft's "[MS-DOC]" specification Section 2.9.245 and
//! Apache POI's SectionTable implementation.

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
        crate::doc::SectionTextFlow::HorizontalNonAsian,
        None,
    )
    .expect("default section properties are valid")
}

/// Generate Word 97+ SEPX properties for the writer's single section.
pub(crate) fn generate_sepx_with_properties(
    first_page_header: bool,
    grpf_ihdt: u8,
    revision: Option<(u16, u32)>,
    columns: Option<&crate::doc::SectionColumnLayout>,
    right_to_left: bool,
    text_flow: crate::doc::SectionTextFlow,
    page_borders: Option<&crate::doc::SectionPageBorders>,
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
        let count_minus_one = u16::try_from(columns.count() - 1)
            .expect("validated section column count fits u16");
        push_word(
            &mut grpprl,
            crate::sprm_operations::SPRM_S_C_COLUMNS,
            count_minus_one,
        );
        match columns {
            crate::doc::SectionColumnLayout::Even {
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
                    *spacing_twips,
                );
                push_bool(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_L_BETWEEN,
                    *line_between,
                );
            },
            crate::doc::SectionColumnLayout::Unequal {
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
                        column.width_twips,
                    );
                    if let Some(spacing) = column.spacing_after_twips {
                        push_indexed_twips(
                            &mut grpprl,
                            crate::sprm_operations::SPRM_S_DXA_COL_SPACING,
                            index,
                            spacing,
                        );
                    }
                }
                push_bool(
                    &mut grpprl,
                    crate::sprm_operations::SPRM_S_L_BETWEEN,
                    *line_between,
                );
            },
        }
    }
    if right_to_left {
        push_bool(
            &mut grpprl,
            crate::sprm_operations::SPRM_S_F_BIDI,
            true,
        );
    }
    if let Some(page_borders) = page_borders {
        page_borders.validate().map_err(|error| error.to_string())?;
        for (opcode, border) in [
            (
                crate::sprm_operations::SPRM_S_BRC_TOP80,
                page_borders.top,
            ),
            (
                crate::sprm_operations::SPRM_S_BRC_LEFT80,
                page_borders.left,
            ),
            (
                crate::sprm_operations::SPRM_S_BRC_BOTTOM80,
                page_borders.bottom,
            ),
            (
                crate::sprm_operations::SPRM_S_BRC_RIGHT80,
                page_borders.right,
            ),
        ] {
            if let Some(border) = border {
                push_page_border(&mut grpprl, opcode, border);
            }
        }
        if page_borders.apply_to != crate::doc::SectionPageBorderApplyTo::AllPages
            || page_borders.depth != crate::doc::SectionPageBorderDepth::InFront
            || page_borders.offset_from != crate::doc::SectionPageBorderOffsetFrom::Text
        {
            let apply_to = match page_borders.apply_to {
                crate::doc::SectionPageBorderApplyTo::AllPages => 0,
                crate::doc::SectionPageBorderApplyTo::FirstPage => 1,
                crate::doc::SectionPageBorderApplyTo::AllButFirstPage => 2,
            };
            let depth = match page_borders.depth {
                crate::doc::SectionPageBorderDepth::InFront => 0,
                crate::doc::SectionPageBorderDepth::Behind => 1,
            };
            let offset_from = match page_borders.offset_from {
                crate::doc::SectionPageBorderOffsetFrom::Text => 0,
                crate::doc::SectionPageBorderOffsetFrom::PageEdge => 1,
            };
            grpprl.extend_from_slice(&crate::sprm_operations::SPRM_S_PGB_PROP.to_le_bytes());
            grpprl.push(apply_to | depth << 3 | offset_from << 5);
            grpprl.push(0);
        }
    }
    if text_flow != crate::doc::SectionTextFlow::HorizontalNonAsian {
        let value = match text_flow {
            crate::doc::SectionTextFlow::HorizontalNonAsian => 0,
            crate::doc::SectionTextFlow::TopToBottomAsian => 1,
            crate::doc::SectionTextFlow::BottomToTop => 2,
            crate::doc::SectionTextFlow::TopToBottomNonAsian => 3,
            crate::doc::SectionTextFlow::HorizontalAsian => 4,
            crate::doc::SectionTextFlow::VerticalNonAsian => 5,
        };
        push_word(
            &mut grpprl,
            crate::sprm_operations::SPRM_S_TEXT_FLOW,
            value,
        );
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

fn push_indexed_twips(output: &mut Vec<u8>, opcode: u16, index: usize, value: u16) {
    output.extend_from_slice(&opcode.to_le_bytes());
    output.push(u8::try_from(index).expect("validated section column index fits u8"));
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_page_border(
    output: &mut Vec<u8>,
    opcode: u16,
    border: crate::doc::SectionPageBorder,
) {
    let style = match border.style {
        crate::doc::SectionPageBorderStyle::Single => 0x01,
        crate::doc::SectionPageBorderStyle::Double => 0x03,
        crate::doc::SectionPageBorderStyle::Thick => 0x05,
        crate::doc::SectionPageBorderStyle::Dotted => 0x06,
        crate::doc::SectionPageBorderStyle::Dashed => 0x07,
        crate::doc::SectionPageBorderStyle::DotDash => 0x08,
        crate::doc::SectionPageBorderStyle::DotDotDash => 0x09,
        crate::doc::SectionPageBorderStyle::Triple => 0x0A,
        crate::doc::SectionPageBorderStyle::ThinThickSmallGap => 0x0B,
        crate::doc::SectionPageBorderStyle::ThickThinSmallGap => 0x0C,
        crate::doc::SectionPageBorderStyle::ThinThickThinSmallGap => 0x0D,
        crate::doc::SectionPageBorderStyle::ThinThickMediumGap => 0x0E,
        crate::doc::SectionPageBorderStyle::ThickThinMediumGap => 0x0F,
        crate::doc::SectionPageBorderStyle::ThinThickThinMediumGap => 0x10,
        crate::doc::SectionPageBorderStyle::ThinThickLargeGap => 0x11,
        crate::doc::SectionPageBorderStyle::ThickThinLargeGap => 0x12,
        crate::doc::SectionPageBorderStyle::ThinThickThinLargeGap => 0x13,
        crate::doc::SectionPageBorderStyle::Wave => 0x14,
        crate::doc::SectionPageBorderStyle::DoubleWave => 0x15,
        crate::doc::SectionPageBorderStyle::DashSmallGap => 0x16,
        crate::doc::SectionPageBorderStyle::DashDotStroked => 0x17,
        crate::doc::SectionPageBorderStyle::ThreeDEmboss => 0x18,
        crate::doc::SectionPageBorderStyle::ThreeDEngrave => 0x19,
        crate::doc::SectionPageBorderStyle::Art(art) => art.code(),
    };
    let color = match border.color {
        crate::doc::SectionPageBorderColor::Automatic => 0x00,
        crate::doc::SectionPageBorderColor::Black => 0x01,
        crate::doc::SectionPageBorderColor::Blue => 0x02,
        crate::doc::SectionPageBorderColor::Cyan => 0x03,
        crate::doc::SectionPageBorderColor::Green => 0x04,
        crate::doc::SectionPageBorderColor::Magenta => 0x05,
        crate::doc::SectionPageBorderColor::Red => 0x06,
        crate::doc::SectionPageBorderColor::Yellow => 0x07,
        crate::doc::SectionPageBorderColor::White => 0x08,
        crate::doc::SectionPageBorderColor::DarkBlue => 0x09,
        crate::doc::SectionPageBorderColor::DarkCyan => 0x0A,
        crate::doc::SectionPageBorderColor::DarkGreen => 0x0B,
        crate::doc::SectionPageBorderColor::DarkMagenta => 0x0C,
        crate::doc::SectionPageBorderColor::DarkRed => 0x0D,
        crate::doc::SectionPageBorderColor::DarkYellow => 0x0E,
        crate::doc::SectionPageBorderColor::DarkGray => 0x0F,
        crate::doc::SectionPageBorderColor::LightGray => 0x10,
    };
    output.extend_from_slice(&opcode.to_le_bytes());
    output.extend_from_slice(&[
        border.width_eighth_points,
        style,
        color,
        border.spacing_points | u8::from(border.shadow) << 5 | u8::from(border.frame) << 6,
    ]);
}

#[cfg(test)]
mod tests {
    use super::generate_sepx_with_properties;

    #[test]
    fn page_border_default_is_omitted_and_nondefault_is_canonical() {
        let default = crate::doc::SectionPageBorders::default();
        assert_eq!(
            generate_sepx_with_properties(
                false,
                0,
                None,
                None,
                false,
                crate::doc::SectionTextFlow::HorizontalNonAsian,
                Some(&default),
            )
            .unwrap(),
            [0, 0]
        );

        let borders = crate::doc::SectionPageBorders {
            top: Some(crate::doc::SectionPageBorder {
                style: crate::doc::SectionPageBorderStyle::Single,
                width_eighth_points: 8,
                color: crate::doc::SectionPageBorderColor::Red,
                spacing_points: 3,
                shadow: true,
                frame: false,
            }),
            apply_to: crate::doc::SectionPageBorderApplyTo::FirstPage,
            ..default
        };
        assert_eq!(
            generate_sepx_with_properties(
                false,
                0,
                None,
                None,
                false,
                crate::doc::SectionTextFlow::HorizontalNonAsian,
                Some(&borders),
            )
            .unwrap(),
            [10, 0, 0x2B, 0x70, 8, 1, 6, 0x23, 0x2F, 0x52, 1, 0]
        );
    }
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
