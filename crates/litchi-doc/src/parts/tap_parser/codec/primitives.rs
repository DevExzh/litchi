//! Primitive TAP operand decoders.

use super::prelude::*;

pub(in crate::parts::tap_parser) fn read_byte(data: &[u8], offset: usize) -> BinaryResult<u8> {
    if offset >= data.len() {
        return Err(litchi_core::binary::BinaryError::InsufficientData {
            expected: offset + 1,
            available: data.len(),
        });
    }
    Ok(data[offset])
}

/// Convert a binary operand error into the DOC package error surface.
#[inline]
pub(in crate::parts::tap_parser) fn binary_to_doc_result<T>(result: BinaryResult<T>) -> Result<T> {
    result.map_err(|e| PackageError::InvalidFormat(format!("Binary read error: {}", e)))
}

impl<'arena> TapParser<'arena> {
    pub(in crate::parts::tap_parser) fn parse_bool8(sprm: &Sprm, name: &str) -> Result<bool> {
        let operand = sprm.operand_bytes();
        if operand.len() != 1 || !matches!(operand[0], 0 | 1) {
            return Err(PackageError::Corrupted(format!(
                "{name} must contain one Boolean8 value"
            )));
        }
        Ok(operand[0] != 0)
    }

    pub(in crate::parts::tap_parser) fn parse_bool16(sprm: &Sprm, name: &str) -> Result<bool> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(format!(
                "{name} must contain one Bool16 value"
            )));
        }
        match binary_to_doc_result(read_u16_le(operand, 0))? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(PackageError::Corrupted(format!(
                "{name} contains an invalid Bool16 value"
            ))),
        }
    }

    pub(in crate::parts::tap_parser) fn parse_horizontal_position(
        sprm: &Sprm,
    ) -> Result<TableHorizontalPosition> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(
                "sprmTDxaAbs operand must contain 2 bytes".to_string(),
            ));
        }
        let stored = binary_to_doc_result(read_i16_le(operand, 0))?;
        Ok(match stored {
            0 => TableHorizontalPosition::Left,
            -4 => TableHorizontalPosition::Center,
            -8 => TableHorizontalPosition::Right,
            -12 => TableHorizontalPosition::Inside,
            -16 => TableHorizontalPosition::Outside,
            _ => {
                let offset = i32::from(stored) - 1;
                if !(-31_679..=31_681).contains(&offset) {
                    return Err(PackageError::Corrupted(
                        "sprmTDxaAbs is outside the XAS_plusOne range".to_string(),
                    ));
                }
                TableHorizontalPosition::Offset(offset as i16)
            },
        })
    }

    pub(in crate::parts::tap_parser) fn parse_vertical_position(
        sprm: &Sprm,
    ) -> Result<TableVerticalPosition> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(
                "sprmTDyaAbs operand must contain 2 bytes".to_string(),
            ));
        }
        let stored = binary_to_doc_result(read_i16_le(operand, 0))?;
        Ok(match stored {
            0 => TableVerticalPosition::Inline,
            -4 => TableVerticalPosition::Top,
            -8 => TableVerticalPosition::Center,
            -12 => TableVerticalPosition::Bottom,
            -16 => TableVerticalPosition::Inside,
            -20 => TableVerticalPosition::Outside,
            _ => {
                let offset = i32::from(stored) - 1;
                if !(-31_679..=31_681).contains(&offset) {
                    return Err(PackageError::Corrupted(
                        "sprmTDyaAbs is outside the YAS_plusOne range".to_string(),
                    ));
                }
                TableVerticalPosition::Offset(offset as i16)
            },
        })
    }

    pub(in crate::parts::tap_parser) fn parse_wrap_distance(
        sprm: &Sprm,
        name: &str,
    ) -> Result<u16> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(format!(
                "{name} operand must contain 2 bytes"
            )));
        }
        let value = binary_to_doc_result(read_u16_le(operand, 0))?;
        if value > 31_680 {
            return Err(PackageError::Corrupted(format!(
                "{name} is outside its nonnegative distance range"
            )));
        }
        Ok(value)
    }

    pub(in crate::parts::tap_parser) fn parse_fts_width(
        sprm: &Sprm,
        usage: WidthUsage,
    ) -> Result<Option<TableWidth>> {
        let operand = sprm.operand_bytes();
        if operand.len() != 3 {
            return Err(PackageError::Corrupted(
                "DOC preferred-width operand must contain 3 bytes".to_string(),
            ));
        }
        let raw_value = binary_to_doc_result(read_u16_le(operand, 1))?;
        let signed_value = i16::from_le_bytes([operand[1], operand[2]]);
        let invalid = || {
            PackageError::Corrupted(
                "DOC preferred-width operand has invalid units or value".to_string(),
            )
        };
        Ok(match (usage, operand[0]) {
            (WidthUsage::Table | WidthUsage::Indent, 0) if raw_value == 0 => None,
            (WidthUsage::TablePart, 0) => None,
            (_, 1) if raw_value == 0 => Some(TableWidth {
                value: 0,
                width_type: WidthType::Auto,
            }),
            (WidthUsage::Table, 2) if raw_value <= 30_000 => Some(TableWidth {
                value: raw_value as i16,
                width_type: WidthType::Percentage,
            }),
            (WidthUsage::TablePart, 2) if raw_value <= 5_000 => Some(TableWidth {
                value: raw_value as i16,
                width_type: WidthType::Percentage,
            }),
            (WidthUsage::Table | WidthUsage::TablePart, 3) if raw_value <= 31_680 => {
                Some(TableWidth {
                    value: raw_value as i16,
                    width_type: WidthType::Twips,
                })
            },
            (WidthUsage::Indent, 3) if (-31_560..=31_680).contains(&signed_value) => {
                Some(TableWidth {
                    value: signed_value,
                    width_type: WidthType::Twips,
                })
            },
            _ => return Err(invalid()),
        })
    }

    pub(in crate::parts::tap_parser) fn parse_justification(
        sprm: &Sprm,
        name: &str,
    ) -> Result<TableJustification> {
        let operand = sprm.operand_bytes();
        if operand.len() != 2 {
            return Err(PackageError::Corrupted(format!(
                "{name} operand must contain 2 bytes"
            )));
        }
        match binary_to_doc_result(read_u16_le(operand, 0))? {
            0 => Ok(TableJustification::Left),
            1 => Ok(TableJustification::Center),
            2 => Ok(TableJustification::Right),
            _ => Err(PackageError::Corrupted(format!(
                "{name} contains an invalid justification"
            ))),
        }
    }

    pub(in crate::parts::tap_parser) fn parse_byte(sprm: &Sprm, name: &str) -> Result<u8> {
        let operand = sprm.operand_bytes();
        if operand.len() != 1 {
            return Err(PackageError::Corrupted(format!(
                "{name} operand must contain exactly 1 byte"
            )));
        }
        Ok(operand[0])
    }

    pub(in crate::parts::tap_parser) fn parse_band_size(sprm: &Sprm, name: &str) -> Result<u8> {
        let size = Self::parse_byte(sprm, name)?;
        if !(1..=3).contains(&size) {
            return Err(PackageError::Corrupted(format!(
                "{name} band size must be in 1..=3"
            )));
        }
        Ok(size)
    }
    /// Convert ico (color index) to RGB.
    ///
    /// Based on POI's color index mapping.

    pub(in crate::parts::tap_parser) fn ico_to_rgb(ico: u8) -> Option<(u8, u8, u8)> {
        Some(match ico {
            0 => return None,
            1 => (0, 0, 0),
            2 => (0, 0, 255),
            3 => (0, 255, 255),
            4 => (0, 255, 0),
            5 => (255, 0, 255),
            6 => (255, 0, 0),
            7 => (255, 255, 0),
            8 => (255, 255, 255),
            9 => (0, 0, 128),
            10 => (0, 128, 128),
            11 => (0, 128, 0),
            12 => (128, 0, 128),
            13 => (128, 0, 0),
            14 => (128, 128, 0),
            15 => (128, 128, 128),
            16 => (192, 192, 192),
            _ => return None,
        })
    }
    pub(in crate::parts::tap_parser) fn parse_colorref(
        bytes: &[u8],
    ) -> Result<Option<(u8, u8, u8)>> {
        if bytes.len() != 4 || !matches!(bytes[3], 0x00 | 0xFF) {
            return Err(PackageError::Corrupted(
                "DOC COLORREF has an invalid automatic-color flag".to_string(),
            ));
        }
        Ok((bytes[3] == 0).then_some((bytes[0], bytes[1], bytes[2])))
    }
}
