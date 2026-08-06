//! Semantic DOC toolbar-control objects.

use crate::package::Result;
use litchi_ole_common::toolbar::{ControlHeader, ControlType, Data};

/// A four-byte DOC command identifier (Cid).
///
/// DOC toolbar controls only permit built-in (cmtFci) and allocated
/// (cmtAllocated) command identifiers. The raw value is retained so the
/// command can be edited without losing its wire representation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CommandId {
    pub(in crate::parts::command_bars) raw: u32,
}

impl CommandId {
    /// Construct a command identifier accepted by a DOC toolbar control.
    pub fn new(raw: u32) -> Result<Self> {
        let kind = (raw & 0x07) as u8;
        if kind != 0x01 && kind != 0x03 {
            return Err(corrupted("Cid command type must be cmtFci or cmtAllocated"));
        }
        Ok(Self { raw })
    }

    /// Return the exact four-byte command identifier.
    pub const fn raw(self) -> u32 {
        self.raw
    }

    /// Return the low three-bit Cmt value.
    pub const fn command_type(self) -> u8 {
        (self.raw & 0x07) as u8
    }

    /// Whether this identifies a built-in FCI command.
    pub const fn is_fci(self) -> bool {
        self.command_type() == 0x01
    }

    /// Whether this identifies an allocated command.
    pub const fn is_allocated(self) -> bool {
        self.command_type() == 0x03
    }
}

/// One DOC toolbar control (TBC).
///
/// The fixed header and shared TBCGeneralInfo are typed by
/// litchi-ole-common; the format-specific tail remains bounded and
/// lossless in Data::specific.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Control<'a> {
    pub(in crate::parts::command_bars) header: ControlHeader,
    pub(in crate::parts::command_bars) command: Option<CommandId>,
    pub(in crate::parts::command_bars) data: Option<Data<'a>>,
}

impl<'a> Control<'a> {
    /// Construct a validated toolbar control.
    pub fn new(
        header: ControlHeader,
        command: Option<CommandId>,
        data: Option<Data<'a>>,
    ) -> Result<Self> {
        let value = Self {
            header,
            command,
            data,
        };
        value.validate()?;
        Ok(value)
    }

    /// Return the shared toolbar-control header.
    pub const fn header(&self) -> &ControlHeader {
        &self.header
    }

    /// Return the optional DOC command identifier.
    pub const fn command(&self) -> Option<CommandId> {
        self.command
    }

    /// Return typed common metadata and its retained specific tail.
    pub const fn data(&self) -> Option<&Data<'a>> {
        self.data.as_ref()
    }

    /// Return the control type from its shared header.
    pub const fn control_type(&self) -> ControlType {
        self.header.control_type()
    }

    /// Move borrowed wire data into an owned control.
    pub fn into_owned(self) -> Control<'static> {
        Control {
            header: self.header,
            command: self.command,
            data: self.data.map(Data::into_owned),
        }
    }

    pub(in crate::parts::command_bars) fn validate(&self) -> Result<()> {
        super::super::validation::validate_control(self)
    }
}

fn corrupted(message: impl Into<String>) -> crate::package::Error {
    crate::package::Error::Corrupted(message.into())
}
