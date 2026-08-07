//! Structural validation for Formula-record metadata and token bounds.

use crate::{Error, Result};

use super::Metadata;

pub(crate) const FORMULA_RECORD_TYPE: u16 = 0x0006;
pub(crate) const FORMULA_FIXED_SIZE: usize = 22;
pub(crate) const MAX_FORMULA_PAYLOAD: usize = 8_224;
const VALID_FLAGS: u16 = 0x002D;

pub(crate) fn decode_flags(flags: u16, tokens: &[u8]) -> Result<Metadata> {
    if flags & !VALID_FLAGS != 0 {
        return Err(invalid("Formula flags contain reserved bits"));
    }
    let metadata = Metadata::from_wire(flags, 0);
    if metadata.shared_formula() && !is_ptg_exp(tokens) {
        return Err(invalid(
            "shared Formula metadata requires a leading PtgExp token",
        ));
    }
    Ok(metadata)
}

pub(crate) fn encode_flags(metadata: &Metadata, tokens: &[u8]) -> Result<u16> {
    if tokens.is_empty() {
        return Err(Error::InvalidFormula(
            "Formula token stream cannot be empty".to_string(),
        ));
    }
    validate_for_write(metadata)?;
    Ok(metadata.wire_flags())
}

pub(crate) fn validate_for_write(metadata: &Metadata) -> Result<()> {
    if let Some(owner) = metadata.array_owner() {
        if metadata.shared_formula() || metadata.shared_owner().is_some() {
            return Err(Error::InvalidData(
                "Formula metadata cannot own both Array and ShrFmla records".to_string(),
            ));
        }
        return owner.validate();
    }

    match (metadata.shared_formula(), metadata.shared_owner()) {
        (true, Some(owner)) => owner.validate(),
        (true, None) => Err(Error::UnsupportedFeature(
            "shared Formula authoring requires a ShrFmla owner".to_string(),
        )),
        (false, Some(_)) => Err(Error::InvalidData(
            "a shared-formula owner requires fShrFmla metadata".to_string(),
        )),
        (false, None) => Ok(()),
    }
}

pub(crate) const fn is_ptg_exp(tokens: &[u8]) -> bool {
    tokens.len() == 5 && tokens[0] == 0x01 && tokens[4] == 0
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: FORMULA_RECORD_TYPE,
        message: message.into(),
    }
}
