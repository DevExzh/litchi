#[path = "controls.rs"]
mod controls;
#[path = "flags.rs"]
mod flags;
#[path = "header.rs"]
mod header;
#[path = "merge.rs"]
mod merge;
#[path = "restrictions.rs"]
mod restrictions;
#[path = "text_icon.rs"]
mod text_icon;

use std::fmt;

pub use self::controls::{ControlType, Data, GeneralInfo};
pub use self::flags::{ControlFlags, Flags, GeneralFlags, SpecificFlags};
pub use self::header::{ControlHeader, Dimensions, Header};
pub use self::merge::{ExtraInfo, MenuMerge, MergeMode};
pub use self::restrictions::{Restrictions, Type};
pub use self::text_icon::{ButtonFlags, ButtonState, HyperlinkType, TextIcon, WString};

/// A malformed or semantically invalid [MS-OSHARED] toolbar structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// The input ended before the named field was complete.
    Truncated(&'static str),
    /// The structure violates a wire or semantic invariant.
    Invalid(String),
}

impl Error {
    pub(crate) fn invalid(message: impl Into<String>) -> Self {
        Self::Invalid(message.into())
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Truncated(field) => write!(formatter, "truncated toolbar {field}"),
            Self::Invalid(message) => write!(formatter, "invalid toolbar structure: {message}"),
        }
    }
}

impl std::error::Error for Error {}
