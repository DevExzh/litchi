//! Wire and relationship validation for command-bar metadata.

use super::model::*;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};

/// A count cap that is large enough for real documents while preventing a
/// malformed FIB payload from requesting an unbounded allocation.
pub(super) const MAX_ITEMS: usize = 1 << 20;
pub(super) const COMMAND_POINTER_BASE: usize = 154;
pub(super) const FIB_INDEX_CMDS: usize = 24;

/// Validate the Word profile required by the package editor.
pub(super) fn package_fib(fib: &FileInformationBlock) -> Result<()> {
    if fib.version() < WORD_97_NFIB {
        return Err(PackageError::UnsupportedVersion {
            nfib: fib.version(),
            name: fib.version_name(),
        });
    }
    if fib.is_encrypted() {
        return Err(corrupted(
            "encrypted DOC packages cannot be edited by the command-bar owner",
        ));
    }
    if fib.table_pointer_count().is_none() {
        return Err(corrupted(
            "WordDocument FIB table-pointer array is truncated",
        ));
    }
    if fib.table_pointer_count().unwrap_or(0) <= FIB_INDEX_CMDS {
        return Err(corrupted("WordDocument FIB does not expose fcCmds/lcbCmds"));
    }
    Ok(())
}

/// Locate the `fcCmds/lcbCmds` pair in the WordDocument stream.
pub(super) fn pointer_location(fib: &FileInformationBlock) -> Result<usize> {
    package_fib(fib)?;
    let offset = COMMAND_POINTER_BASE
        .checked_add(
            FIB_INDEX_CMDS
                .checked_mul(8)
                .ok_or_else(|| corrupted("fcCmds pointer index overflows"))?,
        )
        .ok_or_else(|| corrupted("fcCmds pointer offset overflows"))?;
    let end = offset
        .checked_add(8)
        .ok_or_else(|| corrupted("fcCmds pointer range overflows"))?;
    if end > fib.raw_data().len() {
        return Err(corrupted(
            "WordDocument FIB does not contain fcCmds/lcbCmds",
        ));
    }
    Ok(offset)
}

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
    let mut command_strings = None;
    let mut macro_names = None;
    for entry in &value.entries {
        match entry {
            Entry::MacroCommands(commands) => commands.validate()?,
            Entry::AllocatedCommands(commands) => commands.validate()?,
            Entry::KeyMaps(maps) => maps.validate()?,
            Entry::CommandStrings(strings) => {
                if command_strings.replace(strings).is_some() {
                    return Err(corrupted(
                        "Tcg255 contains more than one TcgSttbf command string table",
                    ));
                }
                strings.validate()?;
            },
            Entry::MacroNames(names) => {
                if macro_names.replace(names).is_some() {
                    return Err(corrupted("Tcg255 contains more than one MacroNames table"));
                }
                names.validate()?;
            },
            Entry::Toolbar(wrapper) => wrapper.validate()?,
        }
    }
    validate_command_references(value, command_strings, macro_names)?;
    Ok(())
}

fn validate_command_references(
    value: &CommandBars<'_>,
    command_strings: Option<&CommandStrings<'_>>,
    macro_names: Option<&MacroNames<'_>>,
) -> Result<()> {
    let macro_commands = value.entries.iter().find_map(|entry| match entry {
        Entry::MacroCommands(commands) => Some(commands),
        _ => None,
    });
    let allocated_commands = value.entries.iter().find_map(|entry| match entry {
        Entry::AllocatedCommands(commands) => Some(commands),
        _ => None,
    });

    if let Some(commands) = macro_commands.filter(|commands| !commands.commands.is_empty()) {
        let names = macro_names.ok_or_else(|| {
            corrupted("PlfMcd requires a MacroNames table for its macro-name indexes")
        })?;
        let strings = command_strings.ok_or_else(|| {
            corrupted("PlfMcd requires a TcgSttbf table for its command-name indexes")
        })?;
        for command in &commands.commands {
            if !names
                .names
                .iter()
                .any(|name| name.index == command.macro_name_index)
            {
                return Err(corrupted(format!(
                    "Mcd ibst {} has no MacroNames entry",
                    command.macro_name_index
                )));
            }
            if usize::from(command.command_name_index) >= strings.strings.len() {
                return Err(corrupted(format!(
                    "Mcd ibstName {} is outside TcgSttbf",
                    command.command_name_index
                )));
            }
        }
    }

    if let Some(commands) = allocated_commands.filter(|commands| !commands.commands.is_empty()) {
        let strings = command_strings.ok_or_else(|| {
            corrupted("PlfAcd requires a TcgSttbf table for its argument indexes")
        })?;
        for command in &commands.commands {
            if usize::from(command.argument_index) >= strings.strings.len() {
                return Err(corrupted(format!(
                    "Acd ibst {} is outside TcgSttbf",
                    command.argument_index
                )));
            }
        }
    }
    Ok(())
}

pub(super) fn validate_command_strings(value: &CommandStrings<'_>) -> Result<()> {
    if value.strings.len() > usize::from(u16::MAX) {
        return Err(corrupted("TcgSttbfCore cData exceeds u16::MAX"));
    }
    value.strings.iter().try_for_each(CommandString::validate)
}

pub(super) fn validate_macro_names(value: &MacroNames<'_>) -> Result<()> {
    if value.names.len() > usize::from(u16::MAX) {
        return Err(corrupted("MacroNames iMac exceeds u16::MAX"));
    }
    let mut indexes = std::collections::HashSet::with_capacity(value.names.len());
    for name in &value.names {
        name.validate()?;
        if !indexes.insert(name.index) {
            return Err(corrupted(format!("MacroNames repeats ibst {}", name.index)));
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
    validate_count(value.delta_controls.len(), "CTBWRAPPER rtbdc")?;
    for control in &value.delta_controls {
        control.validate()?;
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

pub(super) fn validate_control(value: &Control<'_>) -> Result<()> {
    if matches!(
        value.header.control_type(),
        litchi_ole_common::toolbar::ControlType::Unknown(_)
    ) {
        return Err(corrupted("TBCHeader has an unsupported control type"));
    }
    let control_id = value.header.control_id();
    let command_required = control_id != 0x0001 && control_id != 0x1051;
    if command_required != value.command.is_some() {
        return Err(corrupted(
            "TBC Cid presence does not match the toolbar control identifier",
        ));
    }
    if let Some(command) = value.command {
        if !command.is_fci() && !command.is_allocated() {
            return Err(corrupted("TBC Cid has an unsupported Cmt value"));
        }
    }

    if matches!(
        value.header.control_type(),
        litchi_ole_common::toolbar::ControlType::ActiveX
    ) {
        if value.data.is_some() {
            return Err(corrupted("ActiveX TBC controls must not contain TBCData"));
        }
        return Ok(());
    }

    let Some(data) = value.data.as_ref() else {
        return Err(corrupted("non-ActiveX TBC controls require TBCData"));
    };
    let general_flags = data.general().flags();
    if (general_flags.save_text()
        || general_flags.save_misc_ui_strings()
        || general_flags.save_misc_custom())
        && !value.header.specifics().save_ui_strings()
    {
        return Err(corrupted(
            "TBCGeneralInfo strings require TBCSFlags.fSaveUIStrings",
        ));
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
    if toolbar.control_count < 0 {
        return Err(corrupted("CTB cCtls must be nonnegative"));
    }
    if usize::try_from(toolbar.control_count).ok() != Some(toolbar.controls.len()) {
        return Err(corrupted("CTB cCtls does not match the rTBC control count"));
    }
    validate_count(toolbar.controls.len(), "CTB rTBC")?;
    for control in &toolbar.controls {
        control.validate()?;
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
