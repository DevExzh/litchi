//! IWA-native adapters for the dependency-free character-style vocabulary.

use litchi_iwa_text::character::{
    Error as CharacterError, TextCapitalization, TextCharacterSpacing,
};
use litchi_iwa_text::character::{TextLigatures, TextScript, TextStrikethrough, TextUnderline};

use crate::{Error, Result};

impl From<CharacterError> for Error {
    fn from(error: CharacterError) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

/// Native integer conversion for the closed iWork character-style enums.
pub(crate) trait NativeTextValue: Sized {
    /// Convert the semantic value to its native protobuf integer.
    fn native_value(self) -> i32;

    /// Decode a native protobuf integer.
    fn from_native_value(value: i32) -> Result<Self>;
}

impl NativeTextValue for TextUnderline {
    fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Double => 2,
            Self::Wavy => 3,
        }
    }

    fn from_native_value(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Single),
            2 => Ok(Self::Double),
            3 => Ok(Self::Wavy),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native iWork underline type {value}"
            ))),
        }
    }
}

impl NativeTextValue for TextStrikethrough {
    fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Double => 2,
            Self::Triple => 3,
        }
    }

    fn from_native_value(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Single),
            2 => Ok(Self::Double),
            3 => Ok(Self::Triple),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native iWork strikethrough type {value}"
            ))),
        }
    }
}

impl NativeTextValue for TextScript {
    fn native_value(self) -> i32 {
        match self {
            Self::Normal => 0,
            Self::Superscript => 1,
            Self::Subscript => 2,
        }
    }

    fn from_native_value(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::Normal),
            1 => Ok(Self::Superscript),
            2 => Ok(Self::Subscript),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native iWork text script {value}"
            ))),
        }
    }
}

impl NativeTextValue for TextLigatures {
    fn native_value(self) -> i32 {
        match self {
            Self::RequiredOnly => 0,
            Self::Standard => 1,
            Self::All => 2,
        }
    }

    fn from_native_value(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::RequiredOnly),
            1 => Ok(Self::Standard),
            2 => Ok(Self::All),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native iWork ligature policy {value}"
            ))),
        }
    }
}

/// Native protobuf conversion for iWork capitalization fields.
pub(crate) trait NativeTextCapitalization: Sized {
    /// Convert to the native capitalization integer.
    fn native_value(self) -> i32;

    /// Return the optional native linguistic-boundaries flag.
    fn uses_linguistics(self) -> Option<bool>;

    /// Return the number of native override fields authored by this value.
    fn native_override_count(self) -> u32;

    /// Decode the native capitalization integer and linguistic flag.
    fn from_native_value(value: i32, uses_linguistics: Option<bool>) -> Result<Self>;
}

impl NativeTextCapitalization for TextCapitalization {
    fn native_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::AllCaps => 1,
            Self::SmallCaps => 2,
            Self::TitleCase | Self::StartCase => 3,
        }
    }

    fn uses_linguistics(self) -> Option<bool> {
        match self {
            Self::TitleCase => Some(true),
            _ => None,
        }
    }

    fn native_override_count(self) -> u32 {
        match self {
            Self::TitleCase => 2,
            _ => 1,
        }
    }

    fn from_native_value(value: i32, uses_linguistics: Option<bool>) -> Result<Self> {
        match (value, uses_linguistics.unwrap_or(false)) {
            (0, false) => Ok(Self::None),
            (1, false) => Ok(Self::AllCaps),
            (2, false) => Ok(Self::SmallCaps),
            (3, true) => Ok(Self::TitleCase),
            (3, false) => Ok(Self::StartCase),
            (0..=2, true) => Err(Error::InvalidFormat(
                "native iWork linguistic capitalization is not title case".to_owned(),
            )),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native iWork capitalization type {value}"
            ))),
        }
    }
}

/// Native protobuf conversion for iWork character tracking fields.
pub(crate) trait NativeTextCharacterSpacing: Sized {
    /// Convert to the native tracking ratio.
    fn native_ratio(self) -> f32;

    /// Decode a native tracking ratio.
    fn from_native_ratio(ratio: f32) -> Result<Self>;
}

impl NativeTextCharacterSpacing for TextCharacterSpacing {
    fn native_ratio(self) -> f32 {
        self.percent() / 100.0
    }

    fn from_native_ratio(ratio: f32) -> Result<Self> {
        if !ratio.is_finite()
            || !(TextCharacterSpacing::MINIMUM_PERCENT / 100.0
                ..=TextCharacterSpacing::MAXIMUM_PERCENT / 100.0)
                .contains(&ratio)
        {
            return Err(Error::InvalidFormat(
                "native iWork character spacing must be finite and between -0.4 and 4.0".to_owned(),
            ));
        }
        Ok(TextCharacterSpacing::from_percent(ratio * 100.0)?)
    }
}
