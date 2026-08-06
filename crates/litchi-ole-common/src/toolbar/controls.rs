use std::borrow::Cow;

use super::{Error, ExtraInfo, GeneralFlags, WString};

/// Variable general metadata from `TBCGeneralInfo`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeneralInfo<'a> {
    flags: GeneralFlags,
    custom_text: Option<WString<'a>>,
    description: Option<WString<'a>>,
    tooltip: Option<WString<'a>>,
    extra: Option<ExtraInfo<'a>>,
}

impl<'a> GeneralInfo<'a> {
    /// Construct `TBCGeneralInfo` while enforcing its flag-controlled fields.
    pub fn new(
        flags: GeneralFlags,
        custom_text: Option<WString<'a>>,
        description: Option<WString<'a>>,
        tooltip: Option<WString<'a>>,
        extra: Option<ExtraInfo<'a>>,
    ) -> Result<Self, Error> {
        let value = Self {
            flags,
            custom_text,
            description,
            tooltip,
            extra,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn from_decoded(
        flags: GeneralFlags,
        custom_text: Option<WString<'a>>,
        description: Option<WString<'a>>,
        tooltip: Option<WString<'a>>,
        extra: Option<ExtraInfo<'a>>,
    ) -> Self {
        Self {
            flags,
            custom_text,
            description,
            tooltip,
            extra,
        }
    }

    pub const fn flags(&self) -> GeneralFlags {
        self.flags
    }

    pub const fn custom_text(&self) -> Option<&WString<'a>> {
        self.custom_text.as_ref()
    }

    pub const fn description(&self) -> Option<&WString<'a>> {
        self.description.as_ref()
    }

    pub const fn tooltip(&self) -> Option<&WString<'a>> {
        self.tooltip.as_ref()
    }

    pub const fn extra(&self) -> Option<&ExtraInfo<'a>> {
        self.extra.as_ref()
    }

    pub fn into_owned(self) -> GeneralInfo<'static> {
        GeneralInfo {
            flags: self.flags,
            custom_text: self.custom_text.map(WString::into_owned),
            description: self.description.map(WString::into_owned),
            tooltip: self.tooltip.map(WString::into_owned),
            extra: self.extra.map(ExtraInfo::into_owned),
        }
    }

    pub(crate) fn validate(&self) -> Result<(), Error> {
        if self.flags.save_text() != self.custom_text.is_some() {
            return Err(Error::invalid(
                "TBCGeneralInfo custom text must match fSaveText",
            ));
        }
        if self.flags.save_misc_ui_strings()
            != (self.description.is_some() && self.tooltip.is_some())
        {
            return Err(Error::invalid(
                "TBCGeneralInfo UI strings must match fSaveMiscUIStrings",
            ));
        }
        if self.flags.save_misc_custom() != self.extra.is_some() {
            return Err(Error::invalid(
                "TBCGeneralInfo extra info must match fSaveMiscCustom",
            ));
        }
        if let Some(extra) = &self.extra {
            extra.validate()?;
        }
        Ok(())
    }
}

/// Lossless `TBCData` payload.
///
/// The shared layer owns the flag-controlled general metadata and retains the
/// format-specific tail as bytes. XLS/DOC/PPT owners know the surrounding
/// record boundaries and can project that tail into their contextual models
/// without copying or making the common crate depend on a host format.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Data<'a> {
    general: GeneralInfo<'a>,
    specific: Cow<'a, [u8]>,
}

impl<'a> Data<'a> {
    /// Construct a payload from typed common metadata and an opaque specific tail.
    pub fn new(
        general: GeneralInfo<'a>,
        specific: impl Into<Cow<'a, [u8]>>,
    ) -> Result<Self, Error> {
        general.validate()?;
        Ok(Self {
            general,
            specific: specific.into(),
        })
    }

    pub(crate) fn from_decoded(general: GeneralInfo<'a>, specific: &'a [u8]) -> Self {
        Self {
            general,
            specific: Cow::Borrowed(specific),
        }
    }

    pub const fn general(&self) -> &GeneralInfo<'a> {
        &self.general
    }

    pub fn specific(&self) -> &[u8] {
        &self.specific
    }

    pub fn into_owned(self) -> Data<'static> {
        Data {
            general: self.general.into_owned(),
            specific: Cow::Owned(self.specific.into_owned()),
        }
    }
}

/// The toolbar-control types listed by `TBCHeader.tct`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ControlType {
    /// A push button.
    Button,
    /// An edit control.
    Edit,
    /// A drop-down control.
    DropDown,
    /// A combo box.
    ComboBox,
    /// A split drop-down.
    SplitDropDown,
    /// An OCX drop-down.
    OcxDropDown,
    /// A graphic drop-down.
    GraphicDropDown,
    /// A popup.
    Popup,
    /// A button popup.
    ButtonPopup,
    /// A split button popup.
    SplitButtonPopup,
    /// A split button MRU popup.
    SplitButtonMruPopup,
    /// A label.
    Label,
    /// An expanding grid.
    ExpandingGrid,
    /// A grid.
    Grid,
    /// A gauge.
    Gauge,
    /// A graphic combo.
    GraphicCombo,
    /// A pane.
    Pane,
    /// An ActiveX control.
    ActiveX,
    /// An unrecognized future wire value.
    Unknown(u8),
}

impl ControlType {
    pub const fn raw(self) -> u8 {
        match self {
            Self::Button => 0x01,
            Self::Edit => 0x02,
            Self::DropDown => 0x03,
            Self::ComboBox => 0x04,
            Self::SplitDropDown => 0x06,
            Self::OcxDropDown => 0x07,
            Self::GraphicDropDown => 0x09,
            Self::Popup => 0x0A,
            Self::ButtonPopup => 0x0C,
            Self::SplitButtonPopup => 0x0D,
            Self::SplitButtonMruPopup => 0x0E,
            Self::Label => 0x0F,
            Self::ExpandingGrid => 0x10,
            Self::Grid => 0x12,
            Self::Gauge => 0x13,
            Self::GraphicCombo => 0x14,
            Self::Pane => 0x15,
            Self::ActiveX => 0x16,
            Self::Unknown(value) => value,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value {
            0x01 => Self::Button,
            0x02 => Self::Edit,
            0x03 => Self::DropDown,
            0x04 => Self::ComboBox,
            0x06 => Self::SplitDropDown,
            0x07 => Self::OcxDropDown,
            0x09 => Self::GraphicDropDown,
            0x0A => Self::Popup,
            0x0C => Self::ButtonPopup,
            0x0D => Self::SplitButtonPopup,
            0x0E => Self::SplitButtonMruPopup,
            0x0F => Self::Label,
            0x10 => Self::ExpandingGrid,
            0x12 => Self::Grid,
            0x13 => Self::Gauge,
            0x14 => Self::GraphicCombo,
            0x15 => Self::Pane,
            0x16 => Self::ActiveX,
            value => Self::Unknown(value),
        }
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if matches!(self, Self::Unknown(_)) {
            return Err(Error::invalid(format!(
                "unsupported toolbar-control type 0x{:02X}",
                self.raw()
            )));
        }
        Ok(())
    }
}

impl TryFrom<u8> for ControlType {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let kind = Self::from_raw(value);
        kind.validate()?;
        Ok(kind)
    }
}
