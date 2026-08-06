//! Semantic models for the bounded XLS Office Toolbars stream.

use crate::{Error, Result};
use litchi_ole_common::toolbar::{ControlHeader, Header};

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
    pub const fn new(bytes: [u8; VISUAL_DATA_LEN]) -> Self {
        Self { bytes }
    }

    /// Return the exact serialized visual-data bytes.
    pub const fn bytes(&self) -> &[u8; VISUAL_DATA_LEN] {
        &self.bytes
    }

    /// Return one of the three 20-byte view records, if its index is valid.
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
    pub fn new(toolbar_count: u16, active_view: u16) -> Result<Self> {
        Self::from_parts(0x01, 0x01, 0, 0, 0, toolbar_count, 0x0003, active_view)
    }

    /// Construct a header while retaining all reserved wire fields.
    #[allow(clippy::too_many_arguments)]
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

    pub const fn signature(self) -> u8 {
        self.signature
    }

    pub const fn version(self) -> u8 {
        self.version
    }

    pub const fn reserved1(self) -> u16 {
        self.reserved1
    }

    pub const fn reserved2(self) -> u16 {
        self.reserved2
    }

    pub const fn reserved3(self) -> u16 {
        self.reserved3
    }

    pub const fn toolbar_count(self) -> u16 {
        self.toolbar_count
    }

    pub const fn view_count(self) -> u16 {
        self.view_count
    }

    pub const fn active_view(self) -> u16 {
        self.active_view
    }

    pub fn validate(&self) -> Result<()> {
        validation::validate_toolbar_set(self)
    }
}

/// A bounded, inert toolbar control.
///
/// Only `ActiveX` (`tct = 0x16`) controls are representable because that
/// control type has no `TBCData` payload.  The shared control header retains
/// all flags, undefined bits, and future values exactly as decoded.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Control {
    header: ControlHeader,
}

impl Control {
    /// Construct a fixed-header control with no variable `TBCData`.
    pub fn new(header: ControlHeader) -> Result<Self> {
        let value = Self { header };
        value.validate()?;
        Ok(value)
    }

    pub const fn header(&self) -> &ControlHeader {
        &self.header
    }

    /// Return whether this is the fixed-header ActiveX control form.
    pub const fn is_active_x(&self) -> bool {
        self.header.control_type().raw() == 0x16
    }

    pub fn validate(&self) -> Result<()> {
        validation::validate_control(self)
    }

    pub(super) fn from_decoded(header: ControlHeader) -> Result<Self> {
        Self::new(header)
    }
}

/// A `[MS-XLS]` `CTB` custom toolbar.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Toolbar<'a> {
    header: Header<'a>,
    visual_data: Option<VisualData>,
    application_id: i32,
    controls: Vec<Control>,
}

impl<'a> Toolbar<'a> {
    /// Construct an empty custom toolbar with no optional visual records.
    pub fn new(header: Header<'a>, application_id: i32) -> Result<Self> {
        Self::from_parts(header, None, application_id, Vec::new())
    }

    /// Construct a toolbar from its typed metadata and fixed-header controls.
    pub fn from_parts(
        header: Header<'a>,
        visual_data: Option<VisualData>,
        application_id: i32,
        controls: Vec<Control>,
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
    pub fn with_visual_data(mut self, visual_data: VisualData) -> Self {
        self.visual_data = Some(visual_data);
        self
    }

    /// Replace the fixed-header control list.
    pub fn with_controls(mut self, controls: Vec<Control>) -> Self {
        self.controls = controls;
        self
    }

    pub const fn header(&self) -> &Header<'a> {
        &self.header
    }

    pub const fn visual_data(&self) -> Option<&VisualData> {
        self.visual_data.as_ref()
    }

    pub const fn application_id(&self) -> i32 {
        self.application_id
    }

    pub fn controls(&self) -> &[Control] {
        &self.controls
    }

    pub fn validate(&self) -> Result<()> {
        validation::validate_toolbar(self)
    }

    /// Move this decoded toolbar into an owned representation.
    pub fn into_owned(self) -> Toolbar<'static> {
        Toolbar {
            header: self.header.into_owned(),
            visual_data: self.visual_data,
            application_id: self.application_id,
            controls: self.controls,
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
    pub fn new(toolbars: Vec<Toolbar<'a>>) -> Result<Self> {
        let toolbar_count = u16::try_from(toolbars.len())
            .map_err(|_| Error::InvalidData("XCB toolbar count exceeds u16::MAX".to_string()))?;
        let toolbar_set = ToolbarSet::new(toolbar_count, 0)?;
        Self::from_parts(toolbar_set, toolbars)
    }

    /// Construct a wrapper while retaining the exact `CTBS` header.
    pub fn from_parts(toolbar_set: ToolbarSet, toolbars: Vec<Toolbar<'a>>) -> Result<Self> {
        let value = Self {
            toolbar_set,
            toolbars,
        };
        value.validate()?;
        Ok(value)
    }

    pub const fn toolbar_set(&self) -> &ToolbarSet {
        &self.toolbar_set
    }

    pub fn toolbars(&self) -> &[Toolbar<'a>] {
        &self.toolbars
    }

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
