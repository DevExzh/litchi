use super::{Error, WString};

/// OLE host/server merge mode from `TBCExtraInfo.tbcu`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeMode {
    /// The control is not used during OLE merging.
    Neither,
    /// The control is used in OLE server mode.
    Server,
    /// The control is used in OLE host mode.
    Host,
    /// The control is used in both OLE host and server modes.
    Both,
    /// The producer did not provide a value; consumers use `Server`.
    Unresolved,
    /// An unrecognized future wire value.
    Unknown(u8),
}

impl MergeMode {
    pub const fn raw(self) -> u8 {
        match self {
            Self::Neither => 0,
            Self::Server => 1,
            Self::Host => 2,
            Self::Both => 3,
            Self::Unresolved => 0xFF,
            Self::Unknown(value) => value,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value {
            0 => Self::Neither,
            1 => Self::Server,
            2 => Self::Host,
            3 => Self::Both,
            0xFF => Self::Unresolved,
            value => Self::Unknown(value),
        }
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if matches!(self, Self::Unknown(_)) {
            return Err(Error::invalid("TBCExtraInfo tbcu has an invalid value"));
        }
        Ok(())
    }
}

impl TryFrom<u8> for MergeMode {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let mode = Self::from_raw(value);
        mode.validate()?;
        Ok(mode)
    }
}

/// OLE menu merge group from `TBCExtraInfo.tbmg`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MenuMerge {
    /// File menu group.
    File,
    /// Edit menu group.
    Edit,
    /// Container menu group.
    Container,
    /// Object menu group.
    Object,
    /// Window menu group.
    Window,
    /// Help menu group.
    Help,
    /// The control is not placed in an OLE menu group.
    None,
    /// An unrecognized future wire value.
    Unknown(u8),
}

impl MenuMerge {
    pub const fn raw(self) -> u8 {
        match self {
            Self::File => 0,
            Self::Edit => 1,
            Self::Container => 2,
            Self::Object => 3,
            Self::Window => 4,
            Self::Help => 5,
            Self::None => 0xFF,
            Self::Unknown(value) => value,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value {
            0 => Self::File,
            1 => Self::Edit,
            2 => Self::Container,
            3 => Self::Object,
            4 => Self::Window,
            5 => Self::Help,
            0xFF => Self::None,
            value => Self::Unknown(value),
        }
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if matches!(self, Self::Unknown(_)) {
            return Err(Error::invalid("TBCExtraInfo tbmg has an invalid value"));
        }
        Ok(())
    }
}

impl TryFrom<u8> for MenuMerge {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let merge = Self::from_raw(value);
        merge.validate()?;
        Ok(merge)
    }
}

/// Extra command and OLE-merging metadata from `TBCExtraInfo`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtraInfo<'a> {
    help_file: WString<'a>,
    help_context: i32,
    tag: WString<'a>,
    on_action: WString<'a>,
    param: WString<'a>,
    merge: MergeMode,
    menu_merge: MenuMerge,
}

impl<'a> ExtraInfo<'a> {
    /// Construct validated extra command metadata.
    #[allow(clippy::too_many_arguments, reason = "matches the wire fields")]
    pub fn new(
        help_file: WString<'a>,
        help_context: i32,
        tag: WString<'a>,
        on_action: WString<'a>,
        param: WString<'a>,
        merge: MergeMode,
        menu_merge: MenuMerge,
    ) -> Result<Self, Error> {
        let value = Self {
            help_file,
            help_context,
            tag,
            on_action,
            param,
            merge,
            menu_merge,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn from_decoded(
        help_file: WString<'a>,
        help_context: i32,
        tag: WString<'a>,
        on_action: WString<'a>,
        param: WString<'a>,
        merge: u8,
        menu_merge: u8,
    ) -> Self {
        Self {
            help_file,
            help_context,
            tag,
            on_action,
            param,
            merge: MergeMode::from_raw(merge),
            menu_merge: MenuMerge::from_raw(menu_merge),
        }
    }

    pub const fn help_file(&self) -> &WString<'a> {
        &self.help_file
    }

    pub const fn help_context(&self) -> i32 {
        self.help_context
    }

    pub const fn tag(&self) -> &WString<'a> {
        &self.tag
    }

    pub const fn on_action(&self) -> &WString<'a> {
        &self.on_action
    }

    pub const fn param(&self) -> &WString<'a> {
        &self.param
    }

    /// Return the raw `tbcu` OLE host/server merge value.
    pub const fn merge(&self) -> MergeMode {
        self.merge
    }

    /// Return the raw `tbmg` OLE menu merge value.
    pub const fn menu_merge(&self) -> MenuMerge {
        self.menu_merge
    }

    pub fn into_owned(self) -> ExtraInfo<'static> {
        ExtraInfo {
            help_file: self.help_file.into_owned(),
            help_context: self.help_context,
            tag: self.tag.into_owned(),
            on_action: self.on_action.into_owned(),
            param: self.param.into_owned(),
            merge: self.merge,
            menu_merge: self.menu_merge,
        }
    }

    pub(crate) fn validate(&self) -> Result<(), Error> {
        self.merge.validate()?;
        self.menu_merge.validate()?;
        Ok(())
    }
}
