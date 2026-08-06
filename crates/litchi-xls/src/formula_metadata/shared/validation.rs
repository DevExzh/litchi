//! Structural validation for the BIFF8 `ShrFmla` owner.

use crate::{Error, Result};

use super::model::Owner;

pub(crate) const RECORD_TYPE: u16 = 0x04BC;
pub(crate) const FIXED_PAYLOAD_SIZE: usize = 10;
pub(crate) const MAX_RECORD_PAYLOAD: usize = 8_224;
pub(crate) const MAX_FORMULA_BYTES: usize = MAX_RECORD_PAYLOAD - FIXED_PAYLOAD_SIZE;

pub(crate) fn validate(owner: &Owner) -> Result<()> {
    if !owner.range().contains(owner.anchor()) {
        return Err(invalid("shared-formula anchor is outside its RefU range"));
    }
    if owner.participants().is_empty() {
        return Err(invalid("shared-formula cUse must include its anchor cell"));
    }
    if !owner.is_participant(owner.anchor()) {
        return Err(invalid("shared-formula participants omit the anchor cell"));
    }
    if owner
        .participants()
        .windows(2)
        .any(|cells| cells[0] >= cells[1])
    {
        return Err(invalid(
            "shared-formula participants must be strictly ordered",
        ));
    }
    if owner
        .participants()
        .iter()
        .any(|cell| !owner.range().contains(*cell) || *cell < owner.anchor())
    {
        return Err(invalid(
            "shared-formula participants must be in RefU and not precede the anchor",
        ));
    }
    if owner.participants().len() > usize::from(u8::MAX) {
        return Err(invalid("shared-formula cUse exceeds the BIFF8 limit"));
    }
    if owner.tokens().is_empty() {
        return Err(invalid("ShrFmla shared parsed formula cannot be empty"));
    }
    if owner.tokens().len() > MAX_FORMULA_BYTES {
        return Err(invalid(format!(
            "ShrFmla shared parsed formula exceeds the BIFF8 limit of {MAX_FORMULA_BYTES} bytes"
        )));
    }
    if owner
        .tokens()
        .first()
        .is_some_and(|opcode| opcode & 0x7F == 0x01)
    {
        return Err(invalid(
            "ShrFmla shared parsed formula cannot begin with PtgExp",
        ));
    }
    Ok(())
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: RECORD_TYPE,
        message: message.into(),
    }
}
