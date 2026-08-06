use super::{ControlFlags, ControlType, Error, Flags, Restrictions, SpecificFlags, WString};

/// Optional width and height in a `ControlHeader`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Dimensions {
    width: u16,
    height: u16,
}

impl Dimensions {
    /// Construct a pair of unsigned pixel dimensions.
    pub const fn new(width: u16, height: u16) -> Self {
        Self { width, height }
    }

    /// Return the width in pixels.
    pub const fn width(self) -> u16 {
        self.width
    }

    /// Return the height in pixels.
    pub const fn height(self) -> u16 {
        self.height
    }
}

/// The fixed and optional fields of `TBCHeader`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlHeader {
    control_type: ControlType,
    control_id: u16,
    flags: ControlFlags,
    specifics: SpecificFlags,
    priority: u8,
    dimensions: Option<Dimensions>,
}

impl ControlHeader {
    /// Construct a validated toolbar-control header.
    pub fn new(
        control_type: ControlType,
        control_id: u16,
        flags: ControlFlags,
        specifics: SpecificFlags,
        priority: u8,
        dimensions: Option<Dimensions>,
    ) -> Result<Self, Error> {
        let value = Self {
            control_type,
            control_id,
            flags,
            specifics,
            priority,
            dimensions,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn from_decoded(
        control_type: ControlType,
        control_id: u16,
        flags: ControlFlags,
        specifics: SpecificFlags,
        priority: u8,
        dimensions: Option<Dimensions>,
    ) -> Self {
        Self {
            control_type,
            control_id,
            flags,
            specifics,
            priority,
            dimensions,
        }
    }

    /// Return the toolbar-control type.
    pub const fn control_type(&self) -> ControlType {
        self.control_type
    }

    /// Return the format-specific toolbar-control identifier.
    pub const fn control_id(&self) -> u16 {
        self.control_id
    }

    /// Return the general toolbar-control flags.
    pub const fn flags(&self) -> ControlFlags {
        self.flags
    }

    /// Return the toolbar-control settings flags.
    pub const fn specifics(&self) -> SpecificFlags {
        self.specifics
    }

    /// Return the drop and wrap priority.
    pub const fn priority(&self) -> u8 {
        self.priority
    }

    /// Return optional saved dimensions.
    pub const fn dimensions(&self) -> Option<Dimensions> {
        self.dimensions
    }

    pub(crate) fn validate(&self) -> Result<(), Error> {
        self.control_type.validate()?;
        self.flags.validate()?;
        self.specifics.validate()?;
        if self.priority > 7 {
            return Err(Error::invalid("TBCHeader priority exceeds 7"));
        }
        if self.flags.save_dimensions() != self.dimensions.is_some() {
            return Err(Error::invalid("TBCHeader dimensions must match fSaveDxy"));
        }
        Ok(())
    }
}

/// A `TB` toolbar header and its name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header<'a> {
    control_count: i16,
    restrictions: Restrictions,
    rows_default: u16,
    flags: Flags,
    name: WString<'a>,
}

impl<'a> Header<'a> {
    /// Construct a validated toolbar header.
    pub fn new(
        control_count: i16,
        restrictions: Restrictions,
        rows_default: u16,
        flags: Flags,
        name: WString<'a>,
    ) -> Result<Self, Error> {
        let value = Self {
            control_count,
            restrictions,
            rows_default,
            flags,
            name,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn from_decoded(
        control_count: i16,
        restrictions: Restrictions,
        rows_default: u16,
        flags: Flags,
        name: WString<'a>,
    ) -> Self {
        Self {
            control_count,
            restrictions,
            rows_default,
            flags,
            name,
        }
    }

    /// Return the signed `cCL` field exactly as decoded.
    pub const fn control_count(&self) -> i16 {
        self.control_count
    }

    /// Return toolbar type and restriction flags.
    pub const fn restrictions(&self) -> Restrictions {
        self.restrictions
    }

    /// Return the preferred row count exactly as decoded.
    pub const fn rows_default(&self) -> u16 {
        self.rows_default
    }

    /// Return toolbar flags.
    pub const fn flags(&self) -> Flags {
        self.flags
    }

    /// Return the borrowed or owned toolbar name.
    pub const fn name(&self) -> &WString<'a> {
        &self.name
    }

    /// Move a decoded toolbar header into an owned representation.
    ///
    /// This is used by format facades whose compound-file stream buffer is
    /// shorter-lived than the public workbook/document object.
    pub fn into_owned(self) -> Header<'static> {
        Header {
            control_count: self.control_count,
            restrictions: self.restrictions,
            rows_default: self.rows_default,
            flags: self.flags,
            name: self.name.into_owned(),
        }
    }

    pub(crate) fn validate(&self) -> Result<(), Error> {
        if self.control_count < 0 {
            return Err(Error::invalid("TB cCL cannot be negative"));
        }
        if self.rows_default > u16::from(u8::MAX) {
            return Err(Error::invalid("TB cRowsDefault exceeds 255"));
        }
        self.restrictions.validate()?;
        self.flags.validate()?;
        Ok(())
    }
}
