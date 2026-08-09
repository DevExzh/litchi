//! TAP semantic invariants.

use super::prelude::*;

impl TapParser<'_> {
    /// Validate the property list nested inside a `CNFOperand`.
    ///
    /// [MS-DOC] `UpxTapx` lists table-property SPRMs that cannot occur in a
    /// table-style property list. A conditional list inherits those
    /// restrictions, with the conditional border-style exceptions handled by
    /// operations 0x7F..=0x84 (and the diagonal style operations naturally
    /// remaining representable by the typed model).
    pub(in crate::parts::tap_parser) fn validate_conditional_sprms(sprms: &[Sprm]) -> Result<()> {
        for sprm in sprms {
            if !Self::is_tap_sprm(sprm.opcode) {
                return Err(PackageError::Corrupted(format!(
                    "sprmTCnf grpprl contains non-table SPRM {:#06x}",
                    sprm.opcode
                )));
            }
            let operation = get_sprm_operation(sprm.opcode);
            if matches!(
                operation,
                0x01
                    | 0x08
                    | 0x09
                    | 0x0C
                    | 0x12
                    | 0x16
                    | 0x18
                    | 0x19
                    | 0x1A..=0x1D
                    | 0x20..=0x25
                    | 0x29
                    | 0x2B..=0x2C
                    | 0x2F
                    | 0x32
                    | 0x35..=0x36
                    | 0x39
                    | 0x42
                    | 0x60
                    | 0x62
                    | 0x64..=0x65
                    | 0x69
                    | 0x70..=0x72
            ) {
                return Err(PackageError::Corrupted(format!(
                    "sprmTCnf grpprl contains disallowed table SPRM {:#06x}",
                    sprm.opcode
                )));
            }
        }
        Ok(())
    }

    pub(in crate::parts::tap_parser) fn validate_preferred_indent(
        tap: &TableProperties,
    ) -> Result<()> {
        let Some(TableWidth {
            value: indent,
            width_type: WidthType::Twips,
        }) = tap.preferred_indent
        else {
            return Ok(());
        };
        let table_width = match tap.preferred_width {
            Some(TableWidth {
                value,
                width_type: WidthType::Twips,
            }) => i32::from(value),
            _ => {
                i32::from(tap.cell_boundaries.last().copied().unwrap_or(0))
                    - i32::from(tap.cell_boundaries.first().copied().unwrap_or(0))
            },
        };
        if i32::from(indent) + table_width > 31_680 {
            return Err(PackageError::Corrupted(
                "DOC preferred table indent places the right edge beyond 31680 twips".to_string(),
            ));
        }
        Ok(())
    }
}
