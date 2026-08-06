//! Semantic models for the bounded DOC command-bar records.

use crate::package::Result;
use litchi_ole_common::toolbar::Header;
use std::borrow::Cow;

/// A lossless DOC `Xst` string (MS-DOC 2.9.353).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XString<'a> {
    pub(super) encoded: Cow<'a, [u8]>,
}

impl<'a> XString<'a> {
    /// Construct an `Xst` from UTF-8 text.
    pub fn new(value: &str) -> Result<Self> {
        Self::from_units(&value.encode_utf16().collect::<Vec<_>>())
    }

    /// Construct an `Xst` from UTF-16 code units without normalizing them.
    pub fn from_units(units: &[u16]) -> Result<Self> {
        let byte_len = units
            .len()
            .checked_mul(2)
            .ok_or_else(|| corrupted("Xst byte length overflows"))?;
        if units.len() > usize::from(u16::MAX) {
            return Err(corrupted("Xst character count exceeds u16::MAX"));
        }
        let mut encoded = Vec::with_capacity(byte_len);
        for unit in units {
            encoded.extend_from_slice(&unit.to_le_bytes());
        }
        Ok(Self {
            encoded: Cow::Owned(encoded),
        })
    }

    pub(super) fn from_wire(encoded: &'a [u8]) -> Result<Self> {
        if encoded.len() % 2 != 0 {
            return Err(corrupted("Xst payload has an odd byte count"));
        }
        Ok(Self {
            encoded: Cow::Borrowed(encoded),
        })
    }

    /// Number of UTF-16 code units in this string.
    pub fn len(&self) -> usize {
        self.encoded.len() / 2
    }

    /// Whether the string is empty.
    pub fn is_empty(&self) -> bool {
        self.encoded.is_empty()
    }

    /// Exact UTF-16LE payload without the `cch` prefix.
    pub fn encoded_bytes(&self) -> &[u8] {
        &self.encoded
    }

    /// Iterate over the original UTF-16 code units.
    pub fn units(&self) -> impl Iterator<Item = u16> + '_ {
        self.encoded
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
    }

    /// Decode the string for display without changing its stored code units.
    pub fn text(&self) -> String {
        String::from_utf16_lossy(&self.units().collect::<Vec<_>>())
    }

    /// Copy a borrowed string into an owned representation.
    pub fn into_owned(self) -> XString<'static> {
        XString {
            encoded: Cow::Owned(self.encoded.into_owned()),
        }
    }
}

/// A macro-command descriptor from `PlfMcd` (MS-DOC 2.9.154/2.9.202).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MacroCommand {
    pub(super) reserved1: u8,
    pub(super) reserved2: u8,
    pub(super) macro_name_index: u16,
    pub(super) command_name_index: u16,
    pub(super) reserved3: u16,
    pub(super) reserved4: u32,
    pub(super) reserved5: u32,
    pub(super) reserved6: u32,
    pub(super) reserved7: u32,
}

impl MacroCommand {
    /// Construct a descriptor while retaining all ignored wire fields.
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        macro_name_index: u16,
        command_name_index: u16,
        reserved1: u8,
        reserved2: u8,
        reserved3: u16,
        reserved4: u32,
        reserved5: u32,
        reserved6: u32,
        reserved7: u32,
    ) -> Self {
        Self {
            reserved1,
            reserved2,
            macro_name_index,
            command_name_index,
            reserved3,
            reserved4,
            reserved5,
            reserved6,
            reserved7,
        }
    }

    pub const fn macro_name_index(self) -> u16 {
        self.macro_name_index
    }

    pub const fn command_name_index(self) -> u16 {
        self.command_name_index
    }

    pub const fn reserved1(self) -> u8 {
        self.reserved1
    }

    pub const fn reserved2(self) -> u8 {
        self.reserved2
    }

    pub const fn reserved3(self) -> u16 {
        self.reserved3
    }

    pub const fn reserved4(self) -> u32 {
        self.reserved4
    }

    pub const fn reserved5(self) -> u32 {
        self.reserved5
    }

    pub const fn reserved6(self) -> u32 {
        self.reserved6
    }

    pub const fn reserved7(self) -> u32 {
        self.reserved7
    }
}

/// A `PlfMcd` collection of inert macro descriptors.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroCommands {
    pub(super) commands: Vec<MacroCommand>,
}

impl MacroCommands {
    pub fn new(commands: Vec<MacroCommand>) -> Result<Self> {
        let value = Self { commands };
        value.validate()?;
        Ok(value)
    }

    pub fn commands(&self) -> &[MacroCommand] {
        &self.commands
    }

    pub fn commands_mut(&mut self) -> &mut Vec<MacroCommand> {
        &mut self.commands
    }

    pub fn validate(&self) -> Result<()> {
        super::validation::validate_macro_commands(self)
    }
}

/// An allocated command descriptor from `Acd` (MS-DOC 2.9.1).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AllocatedCommand {
    pub(super) argument_index: u16,
    pub(super) command: u16,
    pub(super) flags: u8,
}

impl AllocatedCommand {
    /// Construct an allocated command from its raw 13-bit command and flags.
    pub const fn new(argument_index: u16, command: u16, flags: u8) -> Self {
        Self {
            argument_index,
            command,
            flags,
        }
    }

    pub const fn argument_index(self) -> u16 {
        self.argument_index
    }

    pub const fn command(self) -> u16 {
        self.command
    }

    /// The raw A/B/C flag bits, including the required reserved bit.
    pub const fn flags(self) -> u8 {
        self.flags
    }

    pub const fn is_free(self) -> bool {
        self.flags & 0x02 != 0
    }

    pub const fn is_referenced(self) -> bool {
        self.flags & 0x04 != 0
    }
}

/// A `PlfAcd` collection of inert allocated-command descriptors.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AllocatedCommands {
    pub(super) commands: Vec<AllocatedCommand>,
}

impl AllocatedCommands {
    pub fn new(commands: Vec<AllocatedCommand>) -> Result<Self> {
        let value = Self { commands };
        value.validate()?;
        Ok(value)
    }

    pub fn commands(&self) -> &[AllocatedCommand] {
        &self.commands
    }

    pub fn commands_mut(&mut self) -> &mut Vec<AllocatedCommand> {
        &mut self.commands
    }

    pub fn validate(&self) -> Result<()> {
        super::validation::validate_allocated_commands(self)
    }
}

/// Whether a `PlfKme` contains valid or mismatched-keyboard entries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum KeyMapKind {
    Regular,
    Mismatched,
}

impl KeyMapKind {
    pub const fn raw(self) -> u8 {
        match self {
            Self::Regular => 3,
            Self::Mismatched => 4,
        }
    }

    pub(super) fn from_raw(value: u8) -> Result<Self> {
        match value {
            3 => Ok(Self::Regular),
            4 => Ok(Self::Mismatched),
            _ => Err(corrupted(format!("PlfKme tag must be 3 or 4, got {value}"))),
        }
    }
}

/// The action kind stored in a `Kme` entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Action {
    Command,
    Character,
    Default,
    Unknown(u16),
}

impl Action {
    pub const fn raw(self) -> u16 {
        match self {
            Self::Command => 0,
            Self::Character => 1,
            Self::Default => 3,
            Self::Unknown(value) => value,
        }
    }

    pub(super) const fn from_raw(value: u16) -> Self {
        match value {
            0 => Self::Command,
            1 => Self::Character,
            3 => Self::Default,
            value => Self::Unknown(value),
        }
    }
}

/// One fixed-size keyboard mapping entry from `Kme`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct KeyMap {
    pub(super) reserved1: u16,
    pub(super) reserved2: u16,
    pub(super) primary_key: u16,
    pub(super) secondary_key: u16,
    pub(super) action: Action,
    pub(super) parameter: u32,
}

impl KeyMap {
    pub const fn new(
        primary_key: u16,
        secondary_key: u16,
        action: Action,
        parameter: u32,
        reserved1: u16,
        reserved2: u16,
    ) -> Self {
        Self {
            reserved1,
            reserved2,
            primary_key,
            secondary_key,
            action,
            parameter,
        }
    }

    pub const fn primary_key(self) -> u16 {
        self.primary_key
    }

    pub const fn secondary_key(self) -> u16 {
        self.secondary_key
    }

    pub const fn action(self) -> Action {
        self.action
    }

    pub const fn parameter(self) -> u32 {
        self.parameter
    }

    pub const fn reserved1(self) -> u16 {
        self.reserved1
    }

    pub const fn reserved2(self) -> u16 {
        self.reserved2
    }
}

/// A `PlfKme` collection of inert keyboard mappings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct KeyMaps {
    pub(super) kind: KeyMapKind,
    pub(super) entries: Vec<KeyMap>,
}

impl KeyMaps {
    pub fn new(kind: KeyMapKind, entries: Vec<KeyMap>) -> Result<Self> {
        let value = Self { kind, entries };
        value.validate()?;
        Ok(value)
    }

    pub const fn kind(&self) -> KeyMapKind {
        self.kind
    }

    pub fn entries(&self) -> &[KeyMap] {
        &self.entries
    }

    pub fn entries_mut(&mut self) -> &mut Vec<KeyMap> {
        &mut self.entries
    }

    pub fn validate(&self) -> Result<()> {
        super::validation::validate_key_maps(self)
    }
}

/// The operation encoded by a `TBDelta`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operation {
    Change,
    Insert,
    Modify,
}

impl Operation {
    pub const fn raw(self) -> u8 {
        match self {
            Self::Change => 0,
            Self::Insert => 1,
            Self::Modify => 2,
        }
    }

    pub(super) fn from_raw(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::Change),
            1 => Ok(Self::Insert),
            2 => Ok(Self::Modify),
            _ => Err(corrupted(format!("TBDelta operation {value} is invalid"))),
        }
    }
}

/// A bounded, inert `TBDelta` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ToolbarDelta {
    pub(super) operation: Operation,
    pub(super) at_end: bool,
    pub(super) reserved_flags: u8,
    pub(super) control_index: u8,
    pub(super) next_command: u32,
    pub(super) command: u32,
    pub(super) file_offset: u32,
    pub(super) state: u16,
    pub(super) control_size: u16,
}

impl ToolbarDelta {
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        operation: Operation,
        at_end: bool,
        reserved_flags: u8,
        control_index: u8,
        next_command: u32,
        command: u32,
        file_offset: u32,
        state: u16,
        control_size: u16,
    ) -> Self {
        Self {
            operation,
            at_end,
            reserved_flags,
            control_index,
            next_command,
            command,
            file_offset,
            state,
            control_size,
        }
    }

    pub const fn operation(self) -> Operation {
        self.operation
    }

    pub const fn at_end(self) -> bool {
        self.at_end
    }

    pub const fn reserved_flags(self) -> u8 {
        self.reserved_flags
    }

    pub const fn control_index(self) -> u8 {
        self.control_index
    }

    pub const fn next_command(self) -> u32 {
        self.next_command
    }

    pub const fn command(self) -> u32 {
        self.command
    }

    pub const fn file_offset(self) -> u32 {
        self.file_offset
    }

    pub const fn state(self) -> u16 {
        self.state
    }

    pub const fn on_disk(self) -> bool {
        self.state & 0x0001 != 0
    }

    pub const fn toolbar_index(self) -> u16 {
        (self.state >> 1) & 0x1FFF
    }

    pub const fn dead(self) -> bool {
        self.state & 0x8000 != 0
    }

    pub const fn control_size(self) -> u16 {
        self.control_size
    }
}

/// A custom-toolbar definition or a toolbar-delta array.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CustomizationData<'a> {
    Toolbar(Toolbar<'a>),
    Deltas(Vec<ToolbarDelta>),
}

/// One `Customization` entry in a `CTBWRAPPER`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Customization<'a> {
    pub(super) toolbar_id: u32,
    pub(super) reserved: u16,
    pub(super) delta_count: u16,
    pub(super) data: CustomizationData<'a>,
}

impl<'a> Customization<'a> {
    pub fn toolbar(toolbar: Toolbar<'a>) -> Result<Self> {
        let value = Self {
            toolbar_id: 0,
            reserved: 0,
            delta_count: 0,
            data: CustomizationData::Toolbar(toolbar),
        };
        value.validate(1)?;
        Ok(value)
    }

    pub fn deltas(toolbar_id: u32, deltas: Vec<ToolbarDelta>) -> Result<Self> {
        let delta_count = u16::try_from(deltas.len())
            .map_err(|_| corrupted("Customization delta count exceeds u16::MAX"))?;
        let value = Self {
            toolbar_id,
            reserved: 0,
            delta_count,
            data: CustomizationData::Deltas(deltas),
        };
        value.validate(1)?;
        Ok(value)
    }

    pub const fn toolbar_id(&self) -> u32 {
        self.toolbar_id
    }

    pub const fn reserved(&self) -> u16 {
        self.reserved
    }

    pub const fn delta_count(&self) -> u16 {
        self.delta_count
    }

    pub const fn data(&self) -> &CustomizationData<'a> {
        &self.data
    }

    pub fn data_mut(&mut self) -> &mut CustomizationData<'a> {
        &mut self.data
    }

    pub(super) fn validate(&self, customization_count: usize) -> Result<()> {
        super::validation::validate_customization(self, customization_count)
    }
}

/// A custom toolbar (`CTB`) with controls intentionally bounded to zero.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Toolbar<'a> {
    pub(super) name: XString<'a>,
    pub(super) header: Header<'a>,
    pub(super) visual_data: [u8; 100],
    pub(super) toolbar_index: i32,
    pub(super) reserved: u16,
    pub(super) unused: u16,
    pub(super) control_count: i32,
}

impl<'a> Toolbar<'a> {
    pub fn new(
        name: XString<'a>,
        header: Header<'a>,
        visual_data: [u8; 100],
        toolbar_index: i32,
        reserved: u16,
        unused: u16,
    ) -> Result<Self> {
        let value = Self {
            name,
            header,
            visual_data,
            toolbar_index,
            reserved,
            unused,
            control_count: 0,
        };
        value.validate(1)?;
        Ok(value)
    }

    pub const fn name(&self) -> &XString<'a> {
        &self.name
    }

    pub const fn header(&self) -> &Header<'a> {
        &self.header
    }

    pub const fn visual_data(&self) -> &[u8; 100] {
        &self.visual_data
    }

    pub const fn toolbar_index(&self) -> i32 {
        self.toolbar_index
    }

    pub const fn reserved(&self) -> u16 {
        self.reserved
    }

    pub const fn unused(&self) -> u16 {
        self.unused
    }

    pub const fn control_count(&self) -> i32 {
        self.control_count
    }

    pub(super) fn validate(&self, customization_count: usize) -> Result<()> {
        super::validation::validate_toolbar(self, customization_count)
    }
}

/// A `CTBWRAPPER` containing customizations and lossless opaque TBC bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToolbarWrapper<'a> {
    pub(super) reserved1: u8,
    pub(super) reserved2: u16,
    pub(super) reserved3: u8,
    pub(super) reserved4: u16,
    pub(super) reserved5: u16,
    pub(super) cb_tbd: u16,
    pub(super) toolbar_controls: Cow<'a, [u8]>,
    pub(super) customizations: Vec<Customization<'a>>,
}

impl<'a> ToolbarWrapper<'a> {
    pub fn new(customizations: Vec<Customization<'a>>) -> Result<Self> {
        let value = Self {
            reserved1: 0x12,
            reserved2: 0,
            reserved3: 0x07,
            reserved4: 0x0006,
            reserved5: 0x000C,
            cb_tbd: 0x0012,
            toolbar_controls: Cow::Owned(Vec::new()),
            customizations,
        };
        value.validate().map(|()| value)
    }

    pub fn toolbar_controls(&self) -> &[u8] {
        &self.toolbar_controls
    }

    pub fn customizations(&self) -> &[Customization<'a>] {
        &self.customizations
    }

    pub fn customizations_mut(&mut self) -> &mut Vec<Customization<'a>> {
        &mut self.customizations
    }

    pub const fn customization_count(&self) -> usize {
        self.customizations.len()
    }

    pub fn validate(&self) -> Result<()> {
        super::validation::validate_toolbar_wrapper(self)
    }
}

/// One type-tagged `Tcg255` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Entry<'a> {
    MacroCommands(MacroCommands),
    AllocatedCommands(AllocatedCommands),
    KeyMaps(KeyMaps),
    Toolbar(ToolbarWrapper<'a>),
}

/// Command-related customizations addressed by `fcCmds`/`lcbCmds`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommandBars<'a> {
    pub(super) version: u8,
    pub(super) entries: Vec<Entry<'a>>,
    pub(super) terminator: u8,
}

impl<'a> CommandBars<'a> {
    pub fn new(entries: Vec<Entry<'a>>) -> Result<Self> {
        let value = Self {
            version: 0xFF,
            entries,
            terminator: 0x40,
        };
        value.validate()?;
        Ok(value)
    }

    pub const fn version(&self) -> u8 {
        self.version
    }

    pub fn entries(&self) -> &[Entry<'a>] {
        &self.entries
    }

    pub fn entries_mut(&mut self) -> &mut Vec<Entry<'a>> {
        &mut self.entries
    }

    pub const fn terminator(&self) -> u8 {
        self.terminator
    }

    pub fn validate(&self) -> Result<()> {
        super::validation::validate_command_bars(self)
    }
}

fn corrupted(message: impl Into<String>) -> crate::package::Error {
    crate::package::Error::Corrupted(message.into())
}
