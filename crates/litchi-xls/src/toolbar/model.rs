//! Semantic models for the bounded XLS Office Toolbars stream.

use crate::{Error, Result};
use litchi_ole_common::toolbar::{ControlHeader, Data, ExtraInfo, GeneralInfo, Header};

use super::validation;

/// The fixed byte length of the three XLS toolbar visual records.
pub const VISUAL_DATA_LEN: usize = 60;

/// The application-specific custom-toolbar identifier required by `[MS-XLS]`.
pub const APPLICATION_TOOLBAR_ID: i32 = 0x0000_0FFF;

/// Opaque, lossless `TBVisualData[3]` bytes.
///
/// `[MS-OSHARED]` defines each view as a 20-byte structure.  XLS does not
/// provide a discriminator for the optional array, so the bounded owner keeps
/// those bytes intact instead of interpreting docking behavior or UI state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct VisualData {
    bytes: [u8; VISUAL_DATA_LEN],
}

impl VisualData {
    /// Construct a visual-data array without normalizing any bytes.
    #[must_use]
    pub const fn new(bytes: [u8; VISUAL_DATA_LEN]) -> Self {
        Self { bytes }
    }

    /// Return the exact serialized visual-data bytes.
    #[must_use]
    pub const fn bytes(&self) -> &[u8; VISUAL_DATA_LEN] {
        &self.bytes
    }

    /// Return one of the three 20-byte view records, if its index is valid.
    #[must_use]
    pub fn view(&self, index: usize) -> Option<&[u8; 20]> {
        let start = index.checked_mul(20)?;
        let end = start.checked_add(20)?;
        self.bytes.get(start..end)?.try_into().ok()
    }
}

/// The `[MS-XLS]` `CTBS` toolbar-set header.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToolbarSet {
    signature: u8,
    version: u8,
    reserved1: u16,
    reserved2: u16,
    reserved3: u16,
    toolbar_count: u16,
    view_count: u16,
    active_view: u16,
}

impl ToolbarSet {
    /// Construct the canonical XLS toolbar-set header.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(toolbar_count: u16, active_view: u16) -> Result<Self> {
        Self::from_parts(0x01, 0x01, 0, 0, 0, toolbar_count, 0x0003, active_view)
    }

    /// Construct a header while retaining all reserved wire fields.
    #[allow(
        clippy::too_many_arguments,
        reason = "arguments map positionally to BIFF fields"
    )]
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_parts(
        signature: u8,
        version: u8,
        reserved1: u16,
        reserved2: u16,
        reserved3: u16,
        toolbar_count: u16,
        view_count: u16,
        active_view: u16,
    ) -> Result<Self> {
        let value = Self {
            signature,
            version,
            reserved1,
            reserved2,
            reserved3,
            toolbar_count,
            view_count,
            active_view,
        };
        value.validate()?;
        Ok(value)
    }

    #[must_use]
    pub const fn signature(self) -> u8 {
        self.signature
    }

    #[must_use]
    pub const fn version(self) -> u8 {
        self.version
    }

    #[must_use]
    pub const fn reserved1(self) -> u16 {
        self.reserved1
    }

    #[must_use]
    pub const fn reserved2(self) -> u16 {
        self.reserved2
    }

    #[must_use]
    pub const fn reserved3(self) -> u16 {
        self.reserved3
    }

    #[must_use]
    pub const fn toolbar_count(self) -> u16 {
        self.toolbar_count
    }

    #[must_use]
    pub const fn view_count(self) -> u16 {
        self.view_count
    }

    #[must_use]
    pub const fn active_view(self) -> u16 {
        self.active_view
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        validation::validate_toolbar_set(self)
    }
}

/// The four-byte `[MS-XLS]` `TBCCmd` prefix.
///
/// The command identifier is intentionally retained as a signed value and
/// the command-type byte is validated only against the wire-level reserved
/// bits. The command tables themselves belong to `[MS-CTXLS]`, so this owner
/// does not silently reinterpret an identifier it cannot classify.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Command {
    bytes: [u8; 4],
}

impl Command {
    /// Construct a command while enforcing the reserved-bit rules from
    /// `[MS-XLS]` section 2.6.5.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(command_id: i16, command_type: u8, hide_drawing: bool) -> Result<Self> {
        if command_type > 0x1F {
            return Err(validation::invalid("TBCCmd command type exceeds five bits"));
        }
        let id = command_id.to_le_bytes();
        let mut flags = (command_type & 0x1F) << 2;
        if hide_drawing {
            flags |= 1;
        }
        Self::from_bytes([id[0], id[1], flags, 0])
    }

    /// Decode a command without losing any of its four wire bytes.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_bytes(bytes: [u8; 4]) -> Result<Self> {
        let value = Self { bytes };
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized command bytes.
    #[must_use]
    pub const fn bytes(self) -> [u8; 4] {
        self.bytes
    }

    /// Return the signed command identifier.
    #[must_use]
    pub const fn command_id(self) -> i16 {
        i16::from_le_bytes([self.bytes[0], self.bytes[1]])
    }

    /// Return the command type from the packed command flags.
    #[must_use]
    pub const fn command_type(self) -> u8 {
        (self.bytes[2] >> 2) & 0x1F
    }

    /// Return the `fHideDrawing` bit.
    #[must_use]
    pub const fn hide_drawing(self) -> bool {
        self.bytes[2] & 1 != 0
    }

    fn validate(self) -> Result<()> {
        if self.bytes[2] & 0x02 != 0 {
            return Err(validation::invalid("TBCCmd reserved1 must be zero"));
        }
        if self.bytes[2] & 0x80 != 0 {
            return Err(validation::invalid("TBCCmd reserved2 must be zero"));
        }
        if self.bytes[3] != 0 {
            return Err(validation::invalid("TBCCmd reserved3 must be zero"));
        }
        if !matches!(
            self.command_type(),
            0x00 | 0x01 | 0x02 | 0x03 | 0x05 | 0x07 | 0x08 | 0x10 | 0x14
        ) {
            return Err(validation::invalid(format!(
                "TBCCmd command type 0x{:02X} is not defined",
                self.command_type()
            )));
        }
        if self.hide_drawing() && !matches!(self.command_type(), 0x10 | 0x14) {
            return Err(validation::invalid(
                "TBCCmd fHideDrawing requires a drawing command type",
            ));
        }
        Ok(())
    }
}

/// A bounded, inert toolbar control.
///
/// `ActiveX` (`tct = 0x16`) controls retain their historical fixed-header
/// representation. Every other control retains the shared `TBCData` general
/// metadata and its type-specific tail, plus an optional `TBCCmd`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Control<'a> {
    header: ControlHeader,
    command: Option<Command>,
    data: Option<Data<'a>>,
}

impl<'a> Control<'a> {
    /// Construct a fixed-header `ActiveX` control with no variable data.
    ///
    /// The `'static` result keeps the existing constructor ergonomic for
    /// programmatically-created `ActiveX` controls while decoded data can
    /// continue borrowing the source XCB stream.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(header: ControlHeader) -> Result<Control<'static>> {
        Control::<'static>::from_parts(header, None, None)
    }

    /// Construct a control from its optional command and shared data.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_parts(
        header: ControlHeader,
        command: Option<Command>,
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

    #[must_use]
    pub const fn header(&self) -> &ControlHeader {
        &self.header
    }

    /// Return the optional `[MS-XLS]` command prefix.
    #[must_use]
    pub const fn command(&self) -> Option<Command> {
        self.command
    }

    /// Return the shared `TBCData`, when this is not an `ActiveX` control.
    #[must_use]
    pub const fn data(&self) -> Option<&Data<'a>> {
        self.data.as_ref()
    }

    /// Return the shared general metadata, when present.
    pub fn general(&self) -> Option<&GeneralInfo<'a>> {
        self.data.as_ref().map(Data::general)
    }

    /// Return the optional shared `TBCExtraInfo` metadata directly.
    pub fn extra(&self) -> Option<&ExtraInfo<'a>> {
        self.general().and_then(GeneralInfo::extra)
    }

    /// Return whether this is the fixed-header `ActiveX` control form.
    #[must_use]
    pub const fn is_active_x(&self) -> bool {
        self.header.control_type().raw() == 0x16
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        validation::validate_control(self)
    }

    pub(super) fn from_decoded(
        header: ControlHeader,
        command: Option<Command>,
        data: Option<Data<'a>>,
    ) -> Result<Self> {
        Self::from_parts(header, command, data)
    }

    /// Move this decoded control into an owned representation.
    pub fn into_owned(self) -> Control<'static> {
        Control {
            header: self.header,
            command: self.command,
            data: self.data.map(Data::into_owned),
        }
    }
}

/// A `[MS-XLS]` `CTB` custom toolbar.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Toolbar<'a> {
    header: Header<'a>,
    visual_data: Option<VisualData>,
    application_id: i32,
    controls: Vec<Control<'a>>,
}

impl<'a> Toolbar<'a> {
    /// Construct an empty custom toolbar with no optional visual records.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(header: Header<'a>, application_id: i32) -> Result<Self> {
        Self::from_parts(header, None, application_id, Vec::new())
    }

    /// Construct a toolbar from its typed metadata and fixed-header controls.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_parts(
        header: Header<'a>,
        visual_data: Option<VisualData>,
        application_id: i32,
        controls: Vec<Control<'a>>,
    ) -> Result<Self> {
        let value = Self {
            header,
            visual_data,
            application_id,
            controls,
        };
        value.validate()?;
        Ok(value)
    }

    /// Add or replace the optional lossless 60-byte visual-data array.
    #[must_use]
    pub fn with_visual_data(mut self, visual_data: VisualData) -> Self {
        self.visual_data = Some(visual_data);
        self
    }

    /// Replace the fixed-header control list.
    #[must_use]
    pub fn with_controls(mut self, controls: Vec<Control<'a>>) -> Self {
        self.controls = controls;
        self
    }

    #[must_use]
    pub const fn header(&self) -> &Header<'a> {
        &self.header
    }

    #[must_use]
    pub const fn visual_data(&self) -> Option<&VisualData> {
        self.visual_data.as_ref()
    }

    #[must_use]
    pub const fn application_id(&self) -> i32 {
        self.application_id
    }

    #[must_use]
    pub fn controls(&self) -> &[Control<'a>] {
        &self.controls
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        validation::validate_toolbar(self)
    }

    /// Move this decoded toolbar into an owned representation.
    #[must_use]
    pub fn into_owned(self) -> Toolbar<'static> {
        Toolbar {
            header: self.header.into_owned(),
            visual_data: self.visual_data,
            application_id: self.application_id,
            controls: self.controls.into_iter().map(Control::into_owned).collect(),
        }
    }
}

/// The single `[MS-XLS]` `CTBWRAPPER` stored in an `XCB` stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Wrapper<'a> {
    toolbar_set: ToolbarSet,
    toolbars: Vec<Toolbar<'a>>,
}

impl<'a> Wrapper<'a> {
    /// Construct a canonical wrapper with normal view active.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(toolbars: Vec<Toolbar<'a>>) -> Result<Self> {
        let toolbar_count = u16::try_from(toolbars.len()).map_err(|_error| {
            Error::InvalidData("XCB toolbar count exceeds u16::MAX".to_string())
        })?;
        let toolbar_set = ToolbarSet::new(toolbar_count, 0)?;
        Self::from_parts(toolbar_set, toolbars)
    }

    /// Construct a wrapper while retaining the exact `CTBS` header.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_parts(toolbar_set: ToolbarSet, toolbars: Vec<Toolbar<'a>>) -> Result<Self> {
        let value = Self {
            toolbar_set,
            toolbars,
        };
        value.validate()?;
        Ok(value)
    }

    #[must_use]
    pub const fn toolbar_set(&self) -> &ToolbarSet {
        &self.toolbar_set
    }

    #[must_use]
    pub fn toolbars(&self) -> &[Toolbar<'a>] {
        &self.toolbars
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        validation::validate_wrapper(self)
    }

    /// Move this decoded wrapper into an owned representation.
    pub fn into_owned(self) -> Wrapper<'static> {
        Wrapper {
            toolbar_set: self.toolbar_set,
            toolbars: self.toolbars.into_iter().map(Toolbar::into_owned).collect(),
        }
    }
}
