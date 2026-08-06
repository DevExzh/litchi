//! Wire and relationship validation for command-bar metadata.

use super::model::*;
use crate::package::{Error as PackageError, Result};

/// A count cap that is large enough for real documents while preventing a
/// malformed FIB payload from requesting an unbounded allocation.
pub(super) const MAX_ITEMS: usize = 1 << 20;

pub(super) fn validate_command_bars(value: &CommandBars<'_>) -> Result<()> {
    if value.version != 0xFF {
        return Err(corrupted("Tcg.nTcgVer must be 0xFF"));
    }
    if value.terminator != 0x40 {
        return Err(corrupted("Tcg255 chTerminator must be 0x40"));
    }
    if value.entries.len() > MAX_ITEMS {
        return Err(corrupted("Tcg255 record count exceeds the bounded limit"));
    }
    for entry in &value.entries {
        match entry {
            Entry::MacroCommands(commands) => commands.validate()?,
            Entry::AllocatedCommands(commands) => commands.validate()?,
            Entry::KeyMaps(maps) => maps.validate()?,
            Entry::Toolbar(wrapper) => wrapper.validate()?,
        }
    }
    Ok(())
}

pub(super) fn validate_macro_commands(value: &MacroCommands) -> Result<()> {
    validate_count(value.commands.len(), "PlfMcd iMac")?;
    for command in &value.commands {
        if command.reserved1 != 0x56 {
            return Err(corrupted("Mcd reserved1 must be 0x56"));
        }
        if command.reserved2 != 0 {
            return Err(corrupted("Mcd reserved2 must be zero"));
        }
        if command.reserved3 != 0xFFFF {
            return Err(corrupted("Mcd reserved3 must be 0xFFFF"));
        }
        if command.reserved5 != 0 {
            return Err(corrupted("Mcd reserved5 must be zero"));
        }
    }
    Ok(())
}

pub(super) fn validate_allocated_commands(value: &AllocatedCommands) -> Result<()> {
    validate_count(value.commands.len(), "PlfAcd iMac")?;
    for command in &value.commands {
        if command.command > 0x1FFF {
            return Err(corrupted("Acd fciBasedOn exceeds 13 bits"));
        }
        if command.flags & 0x01 == 0 {
            return Err(corrupted("Acd reserved flag bit must be set"));
        }
        let free = command.flags & 0x02 != 0;
        let referenced = command.flags & 0x04 != 0;
        if free == referenced {
            return Err(corrupted("Acd fFree and fRef have an invalid combination"));
        }
    }
    Ok(())
}

pub(super) fn validate_key_maps(value: &KeyMaps) -> Result<()> {
    validate_count(value.entries.len(), "PlfKme iMac")?;
    for entry in &value.entries {
        if entry.reserved1 != 0 || entry.reserved2 != 0 {
            return Err(corrupted("Kme reserved words must be zero"));
        }
        if matches!(entry.action, Action::Character) && entry.parameter > u32::from(u16::MAX) {
            return Err(corrupted(
                "Kme ktChar parameter exceeds a Unicode character",
            ));
        }
    }
    Ok(())
}

pub(super) fn validate_toolbar_wrapper(value: &ToolbarWrapper<'_>) -> Result<()> {
    if value.reserved1 != 0x12
        || value.reserved2 != 0
        || value.reserved3 != 0x07
        || value.reserved4 != 0x0006
        || value.reserved5 != 0x000C
    {
        return Err(corrupted("CTBWRAPPER reserved fields have invalid values"));
    }
    if value.cb_tbd != 0x0012 {
        return Err(corrupted("CTBWRAPPER cbTBD must be 0x0012"));
    }
    validate_count(value.customizations.len(), "CTBWRAPPER cCust")?;
    if value.customizations.is_empty() {
        return Err(corrupted("CTBWRAPPER cCust must be greater than zero"));
    }
    if value.customizations.len() > usize::from(i16::MAX as u16) {
        return Err(corrupted("CTBWRAPPER cCust exceeds i16::MAX"));
    }
    if value.toolbar_controls.len() > usize::try_from(i32::MAX).unwrap_or(usize::MAX) {
        return Err(corrupted("CTBWRAPPER cbDTBC exceeds i32::MAX"));
    }
    for customization in &value.customizations {
        validate_customization(customization, value.customizations.len())?;
    }
    Ok(())
}

pub(super) fn validate_customization(
    value: &Customization<'_>,
    customization_count: usize,
) -> Result<()> {
    if value.reserved != 0 {
        return Err(corrupted("Customization reserved field must be zero"));
    }
    match &value.data {
        CustomizationData::Toolbar(toolbar) => {
            if value.toolbar_id != 0 || value.delta_count != 0 {
                return Err(corrupted(
                    "Customization CTB entry must have zero toolbar id and delta count",
                ));
            }
            validate_toolbar(toolbar, customization_count)?;
        },
        CustomizationData::Deltas(deltas) => {
            if value.toolbar_id == 0 {
                return Err(corrupted("Customization delta entry must name a toolbar"));
            }
            if usize::from(value.delta_count) != deltas.len() {
                return Err(corrupted(
                    "Customization ctbds does not match its delta count",
                ));
            }
            validate_count(deltas.len(), "Customization ctbds")?;
            for delta in deltas {
                validate_delta(delta)?;
            }
        },
    }
    Ok(())
}

pub(super) fn validate_toolbar(toolbar: &Toolbar<'_>, customization_count: usize) -> Result<()> {
    if toolbar.toolbar_index < 0 {
        return Err(corrupted("CTB iWCTB must be nonnegative"));
    }
    if usize::try_from(toolbar.toolbar_index).ok() >= Some(customization_count) {
        return Err(corrupted("CTB iWCTB is outside rCustomizations"));
    }
    if toolbar.control_count != 0 || toolbar.header.control_count() != 0 {
        return Err(corrupted(
            "CTB toolbar controls are unsupported unless both counts are zero",
        ));
    }
    if toolbar.reserved != 0 {
        return Err(corrupted("CTB reserved field must be zero"));
    }
    if toolbar.name.len() > usize::from(u16::MAX) {
        return Err(corrupted("CTB name exceeds u16::MAX characters"));
    }
    Ok(())
}

fn validate_delta(delta: &ToolbarDelta) -> Result<()> {
    if delta.reserved_flags != 0 {
        return Err(corrupted("TBDelta reserved flags must be zero"));
    }
    if !matches!(delta.operation, Operation::Insert) && delta.at_end {
        return Err(corrupted("TBDelta fAtEnd is only valid for insertion"));
    }
    if delta.operation != Operation::Modify && !delta.on_disk() {
        return Err(corrupted("TBDelta change/insert records must be on disk"));
    }
    if !delta.on_disk() && (delta.file_offset != 0 || delta.control_size != 0) {
        return Err(corrupted(
            "TBDelta fc and cbTBC must be zero when fOnDisk is clear",
        ));
    }
    Ok(())
}

pub(super) fn validate_count(count: usize, field: &str) -> Result<()> {
    if count > MAX_ITEMS {
        return Err(corrupted(format!(
            "{field} count exceeds the bounded limit"
        )));
    }
    Ok(())
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
