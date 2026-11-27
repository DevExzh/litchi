//! Validation of bounded MS-DOC numbering wire values.

use super::model::{ListAlignment, ListFollowCharacter, NumberFormat};
use crate::package::{Error as PackageError, Result};

/// Prevent malformed table counts from causing unbounded allocations.
pub(super) const MAX_ITEMS: usize = 1 << 20;

pub(super) fn count(value: usize, context: &str) -> Result<()> {
    if value > MAX_ITEMS {
        return Err(PackageError::InvalidFormat(format!(
            "{context} exceeds the bounded item limit"
        )));
    }
    Ok(())
}

pub(super) fn level(value: u8) -> Result<()> {
    if value > 8 {
        return Err(PackageError::InvalidFormat(
            "list level index exceeds 8".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn number_format(raw: u8) -> Result<NumberFormat> {
    let value = NumberFormat::try_from(raw).map_err(|invalid| {
        PackageError::InvalidFormat(format!("LVLF has invalid MSONFC value {invalid:#04x}"))
    })?;
    if matches!(
        value,
        NumberFormat::Hex
            | NumberFormat::Chicago
            | NumberFormat::DecimalHalfWidth
            | NumberFormat::DecimalFullWidth2
    ) {
        return Err(PackageError::InvalidFormat(format!(
            "LVLF forbids MSONFC value {raw:#04x}"
        )));
    }
    Ok(value)
}

pub(super) fn start_at(value: NumberFormat, start_at: u32) -> Result<()> {
    if value != NumberFormat::Bullet && value != NumberFormat::None && start_at > 0x7FFF {
        return Err(PackageError::InvalidFormat(format!(
            "LVLF start value {start_at} exceeds 32767"
        )));
    }
    Ok(())
}

pub(super) fn alignment(raw: u8) -> Result<ListAlignment> {
    ListAlignment::try_from(raw).map_err(|invalid| {
        PackageError::InvalidFormat(format!("LVLF has invalid alignment {invalid}"))
    })
}

pub(super) fn follow_character(raw: u8) -> Result<ListFollowCharacter> {
    ListFollowCharacter::try_from(raw).map_err(|invalid| {
        PackageError::InvalidFormat(format!("LVLF has invalid follow character {invalid}"))
    })
}

pub(super) fn override_count(value: u8) -> Result<()> {
    if value > 9 {
        return Err(PackageError::InvalidFormat(
            "LFO override count exceeds 9".to_string(),
        ));
    }
    Ok(())
}
