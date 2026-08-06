//! Strict public value types for native iWork Date & Time smart fields.

use crate::{Error, Result};

use litchi_iwa_text::position::TextRange;

const MAX_DATE_TIME_FORMAT_BYTES: usize = 4 * 1_024;
const MAX_DATE_TIME_LOCALE_BYTES: usize = 512;
const MAX_DATE_TIME_DISPLAY_BYTES: usize = 16 * 1_024;

/// Identifier of a native Date & Time smart-field object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextDateTimeFieldId(u64);

impl TextDateTimeFieldId {
    /// Construct an identifier obtained from a previously read field.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork Date & Time field object identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying package object identifier.
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// ICU-style date/time format string stored by iWork.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextDateTimeFormat(Box<str>);

impl TextDateTimeFormat {
    /// Validate and construct a native ICU-style format string.
    pub fn new(format: impl Into<Box<str>>) -> Result<Self> {
        let format = format.into();
        validate_nonempty_text(&format, MAX_DATE_TIME_FORMAT_BYTES, "Date & Time format")?;
        Ok(Self(format))
    }

    /// Return the format string exactly as stored by iWork.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextDateTimeFormat {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Locale identifier used by iWork's date formatter, such as `en_US`.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextDateTimeLocaleIdentifier(Box<str>);

impl TextDateTimeLocaleIdentifier {
    /// Validate and construct an iWork locale identifier.
    pub fn new(locale: impl Into<Box<str>>) -> Result<Self> {
        let locale = locale.into();
        validate_nonempty_text(
            &locale,
            MAX_DATE_TIME_LOCALE_BYTES,
            "Date & Time locale identifier",
        )?;
        if locale.trim() != locale.as_ref() {
            return Err(Error::ParseError(
                "iWork Date & Time locale identifier cannot have surrounding whitespace".to_owned(),
            ));
        }
        Ok(Self(locale))
    }

    /// Return the locale identifier exactly as stored by iWork.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextDateTimeLocaleIdentifier {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Visible text inserted for a Date & Time field.
///
/// Formatting is intentionally explicit: litchi-iwa stores the exact string
/// supplied by the caller and does not pretend to implement Apple's locale
/// formatter.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextDateTimeDisplayText(Box<str>);

impl TextDateTimeDisplayText {
    /// Validate exact visible text for insertion with a Date & Time field.
    pub fn new(text: impl Into<Box<str>>) -> Result<Self> {
        let text = text.into();
        validate_nonempty_text(
            &text,
            MAX_DATE_TIME_DISPLAY_BYTES,
            "Date & Time display text",
        )?;
        Ok(Self(text))
    }

    /// Return the exact visible text.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextDateTimeDisplayText {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Seconds from Apple's 2001-01-01 00:00:00 UTC reference date.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextDateTimeInstant(f64);

impl TextDateTimeInstant {
    /// Construct an instant from finite Apple reference-date seconds.
    pub fn from_reference_date_seconds(seconds: f64) -> Result<Self> {
        if !seconds.is_finite() {
            return Err(Error::ParseError(
                "iWork Date & Time instant must be finite".to_owned(),
            ));
        }
        Ok(Self(seconds))
    }

    /// Return seconds from Apple's UTC reference date.
    pub const fn reference_date_seconds(self) -> f64 {
        self.0
    }
}

/// Native date or time style selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextDateTimeFormatterStyle {
    None,
    Short,
    Medium,
    Long,
    Full,
    Unknown(i32),
}

impl TextDateTimeFormatterStyle {
    pub(crate) const fn from_raw(raw: i32) -> Self {
        match raw {
            0 => Self::None,
            1 => Self::Short,
            2 => Self::Medium,
            3 => Self::Long,
            4 => Self::Full,
            unknown => Self::Unknown(unknown),
        }
    }

    pub(crate) const fn as_raw(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Short => 1,
            Self::Medium => 2,
            Self::Long => 3,
            Self::Full => 4,
            Self::Unknown(raw) => raw,
        }
    }
}

/// Native refresh policy for a Date & Time field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextDateTimeUpdatePlan {
    Never,
    Automatic,
    Once,
    Unknown(i32),
}

impl TextDateTimeUpdatePlan {
    pub(crate) const fn from_raw(raw: i32) -> Self {
        match raw {
            0 => Self::Never,
            1 => Self::Automatic,
            2 => Self::Once,
            unknown => Self::Unknown(unknown),
        }
    }

    pub(crate) const fn as_raw(self) -> i32 {
        match self {
            Self::Never => 0,
            Self::Automatic => 1,
            Self::Once => 2,
            Self::Unknown(raw) => raw,
        }
    }
}

/// Lossless writable payload of a native Date & Time field.
#[derive(Debug, Clone, PartialEq)]
pub struct TextDateTimeFieldSettings {
    /// Optional ICU-style format string.
    pub format: Option<TextDateTimeFormat>,
    /// Optional locale identifier used by the formatter.
    pub locale_identifier: Option<TextDateTimeLocaleIdentifier>,
    /// Optional native date-style selector.
    pub date_style: Option<TextDateTimeFormatterStyle>,
    /// Optional native time-style selector.
    pub time_style: Option<TextDateTimeFormatterStyle>,
    /// Optional native refresh policy.
    pub update_plan: Option<TextDateTimeUpdatePlan>,
    /// Optional native flag requesting a display refresh.
    pub needs_update: Option<bool>,
    /// Optional instant in Apple reference-date seconds.
    pub instant: Option<TextDateTimeInstant>,
}

impl TextDateTimeFieldSettings {
    /// Construct a fixed field with explicit formatter metadata.
    pub fn fixed(
        format: TextDateTimeFormat,
        locale_identifier: TextDateTimeLocaleIdentifier,
        instant: TextDateTimeInstant,
    ) -> Self {
        Self {
            format: Some(format),
            locale_identifier: Some(locale_identifier),
            date_style: Some(TextDateTimeFormatterStyle::None),
            time_style: Some(TextDateTimeFormatterStyle::None),
            update_plan: Some(TextDateTimeUpdatePlan::Never),
            needs_update: Some(false),
            instant: Some(instant),
        }
    }

    /// Replace both native formatter style selectors.
    pub const fn with_styles(
        mut self,
        date_style: TextDateTimeFormatterStyle,
        time_style: TextDateTimeFormatterStyle,
    ) -> Self {
        self.date_style = Some(date_style);
        self.time_style = Some(time_style);
        self
    }

    /// Replace the native refresh policy.
    pub const fn with_update_plan(mut self, update_plan: TextDateTimeUpdatePlan) -> Self {
        self.update_plan = Some(update_plan);
        self
    }
}

/// One native Date & Time field attached to a nonempty UTF-16 range.
#[derive(Debug, Clone, PartialEq)]
pub struct TextDateTimeField {
    /// Native smart-field object identifier.
    pub id: TextDateTimeFieldId,
    /// Nonempty UTF-16 text range covered by the field.
    pub range: TextRange,
    /// Losslessly decoded native formatter payload.
    pub settings: TextDateTimeFieldSettings,
}

impl TextDateTimeField {
    pub(crate) fn new(
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Self {
        Self {
            id,
            range,
            settings,
        }
    }
}

fn validate_nonempty_text(value: &str, maximum: usize, label: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::ParseError(format!("iWork {label} cannot be empty")));
    }
    if value.len() > maximum {
        return Err(Error::ParseError(format!(
            "iWork {label} exceeds {maximum} bytes"
        )));
    }
    if value.chars().any(char::is_control) {
        return Err(Error::ParseError(format!(
            "iWork {label} cannot contain control characters"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn public_scalars_are_strict_and_lossless() {
        assert!(TextDateTimeFieldId::from_object_id(0).is_err());
        assert!(TextDateTimeFormat::new("").is_err());
        assert!(TextDateTimeLocaleIdentifier::new(" en_US").is_err());
        assert!(TextDateTimeDisplayText::new("Friday\nJuly").is_err());
        assert!(TextDateTimeInstant::from_reference_date_seconds(f64::NAN).is_err());
        for raw in [-7, 0, 1, 2, 3, 4, 9] {
            assert_eq!(TextDateTimeFormatterStyle::from_raw(raw).as_raw(), raw);
            assert_eq!(TextDateTimeUpdatePlan::from_raw(raw).as_raw(), raw);
        }
    }
}
