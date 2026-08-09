//! TAP input framing and semantic parse entry points.

use super::prelude::*;

impl TapParser<'_> {
    pub(in crate::parts::tap_parser) fn parse_tap_context(
        &self,
        grpprl: &[u8],
        inside_conditional: bool,
        stylesheet: Option<&StyleSheet>,
    ) -> Result<TableProperties> {
        // Parse all SPRMs using arena for temporary storage
        let sprms = parse_sprms(grpprl)?;
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != grpprl.len() {
            return Err(PackageError::Corrupted(
                "TAP grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }
        if inside_conditional {
            Self::validate_conditional_sprms(&sprms)?;
        }

        // Find sprmTDefTable (0xD608 / operation 0x08) to initialize TAP
        let mut tap = self.find_and_init_tap(&sprms)?;

        // Apply each TAP-type SPRM to the table properties
        for sprm in sprms {
            if Self::is_tap_sprm(sprm.opcode) {
                self.apply_sprm_to_tap(&mut tap, &sprm, grpprl, inside_conditional, stylesheet)?;
            }
        }

        Self::validate_preferred_indent(&tap)?;

        Ok(tap)
    }

    /// Find sprmTDefTable and initialize TAP structure.
    ///
    /// This SPRM defines the basic table structure including cell count
    /// and cell boundaries.

    pub(in crate::parts::tap_parser) fn find_and_init_tap(
        &self,
        sprms: &[Sprm],
    ) -> Result<TableProperties> {
        for sprm in sprms {
            if sprm.opcode == 0xD608 {
                // Found sprmTDefTable
                // The shared decoder removes the long-SPRM size field, so the
                // first operand byte is itcMac.
                if let Some(cell_count) = sprm.operand_byte() {
                    let cell_count = cell_count as usize;
                    if cell_count > 63 {
                        return Err(PackageError::Corrupted(
                            "sprmTDefTable contains more than 63 columns".to_string(),
                        ));
                    }
                    return Ok(TableProperties::with_cell_count(cell_count));
                }
            }
        }

        // No table definition found - use default with 1 cell
        Ok(TableProperties::with_cell_count(1))
    }

    /// Check if a SPRM is a TAP (table) SPRM.
    ///
    /// TAP SPRMs have type 5 (bits 10-12 of opcode).

    pub(in crate::parts::tap_parser) fn is_tap_sprm(opcode: u16) -> bool {
        ((opcode >> 10) & 0x07) == 5
    }
}
