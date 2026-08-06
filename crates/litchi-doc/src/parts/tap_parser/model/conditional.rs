//! Conditional table-style property exceptions from [MS-DOC].

use super::prelude::*;

impl<'arena> TapParser<'arena> {
    pub(in crate::parts::tap_parser) fn parse_conditional_formatting(
        &self,
        tap: &mut TableProperties,
        sprm: &Sprm,
    ) -> Result<()> {
        let operand = sprm.operand_bytes();
        let (condition_offset, grpprl_offset) =
            if operand.len() >= 3 && usize::from(operand[0]) == operand.len() - 1 {
                // [MS-DOC] CNFOperand.cb excludes itself and therefore covers
                // cnfc plus the nested grpprl.
                if operand[0] < 2 {
                    return Err(PackageError::Corrupted(
                        "sprmTCnf cb must include the 2-byte condition".to_string(),
                    ));
                }
                (1, 3)
            } else {
                // Keep accepting the historical in-tree shape while producers
                // migrate to the exact CNFOperand layout.
                if operand.len() < 2 {
                    return Err(PackageError::Corrupted(
                        "sprmTCnf must contain a 2-byte condition".to_string(),
                    ));
                }
                (0, 2)
            };
        let code = binary_to_doc_result(read_u16_le(operand, condition_offset))?;
        let condition = TableStyleCondition::from_code(code).ok_or_else(|| {
            PackageError::Corrupted(format!("sprmTCnf contains invalid condition {code:#06x}"))
        })?;
        let raw_grpprl = operand[grpprl_offset..].to_vec();
        let nested = self.parse_conditional_tap(&raw_grpprl)?;
        tap.conditional_formats.push(TableConditionalFormatting {
            condition,
            properties: nested.style_defaults,
            raw_grpprl,
        });
        Ok(())
    }
}
