//! Archive-free semantic values for iWork Date & Time smart fields.

const MAX_FORMAT_BYTES: usize = 4 * 1_024;
const MAX_LOCALE_IDENTIFIER_BYTES: usize = 512;
const MAX_DISPLAY_TEXT_BYTES: usize = 16 * 1_024;

/// Why a Date & Time semantic value was rejected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A format string was empty.
    FormatEmpty,
    /// A format string exceeded [`Format::MAX_BYTES`].
    FormatTooLong,
    /// A format string contained a Unicode control character.
    FormatControlCharacter,
    /// A locale identifier was empty.
    LocaleIdentifierEmpty,
    /// A locale identifier exceeded [`LocaleIdentifier::MAX_BYTES`].
    LocaleIdentifierTooLong,
    /// A locale identifier had surrounding whitespace.
    LocaleIdentifierSurroundingWhitespace,
    /// A locale identifier contained a Unicode control character.
    LocaleIdentifierControlCharacter,
    /// Display text was empty.
    DisplayTextEmpty,
    /// Display text exceeded [`DisplayText::MAX_BYTES`].
    DisplayTextTooLong,
    /// Display text contained a Unicode control character.
    DisplayTextControlCharacter,
    /// An instant was NaN or infinite.
    InstantNotFinite,
    /// A known formatter style was represented by `Unknown`.
    NonCanonicalFormatterStyle,
    /// A known update plan was represented by `Unknown`.
    NonCanonicalUpdatePlan,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::FormatEmpty => "iWork Date & Time format cannot be empty",
            Self::FormatTooLong => "iWork Date & Time format exceeds 4096 bytes",
            Self::FormatControlCharacter => {
                "iWork Date & Time format cannot contain control characters"
            },
            Self::LocaleIdentifierEmpty => "iWork Date & Time locale identifier cannot be empty",
            Self::LocaleIdentifierTooLong => {
                "iWork Date & Time locale identifier exceeds 512 bytes"
            },
            Self::LocaleIdentifierSurroundingWhitespace => {
                "iWork Date & Time locale identifier cannot have surrounding whitespace"
            },
            Self::LocaleIdentifierControlCharacter => {
                "iWork Date & Time locale identifier cannot contain control characters"
            },
            Self::DisplayTextEmpty => "iWork Date & Time display text cannot be empty",
            Self::DisplayTextTooLong => "iWork Date & Time display text exceeds 16384 bytes",
            Self::DisplayTextControlCharacter => {
                "iWork Date & Time display text cannot contain control characters"
            },
            Self::InstantNotFinite => "iWork Date & Time instant must be finite",
            Self::NonCanonicalFormatterStyle => {
                "iWork Date & Time formatter style must use its named variant for known values"
            },
            Self::NonCanonicalUpdatePlan => {
                "iWork Date & Time update plan must use its named variant for known values"
            },
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for Date & Time semantic values.
pub type Result<T> = std::result::Result<T, Error>;

/// Validated ICU-style date/time format text.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Format(Box<str>);

impl Format {
    /// Maximum UTF-8 bytes retained by one format string.
    pub const MAX_BYTES: usize = MAX_FORMAT_BYTES;

    /// Validate and own a borrowed format string.
    ///
    /// Validation happens before allocation for borrowed input.
    ///
    /// # Errors
    ///
    /// Returns the corresponding [`Error`] variant when the value is empty,
    /// too long, or contains a control character.
    pub fn new(value: &str) -> Result<Self> {
        validate_text(value, TextKind::Format)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain an existing boxed string without reallocating it.
    ///
    /// # Errors
    ///
    /// Returns the corresponding [`Error`] variant when the value is empty,
    /// too long, or contains a control character.
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_text(&value, TextKind::Format)?;
        Ok(Self(value))
    }

    /// Borrow the exact format text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the format as an owned `String`.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for Format {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for Format {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        validate_text(&value, TextKind::Format)?;
        Ok(Self(value.into_boxed_str()))
    }
}

impl TryFrom<Box<str>> for Format {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        validate_text(&value, TextKind::Format)?;
        Ok(Self(value))
    }
}

/// Validated locale identifier used by the iWork date formatter.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct LocaleIdentifier(Box<str>);

impl LocaleIdentifier {
    /// Maximum UTF-8 bytes retained by one locale identifier.
    pub const MAX_BYTES: usize = MAX_LOCALE_IDENTIFIER_BYTES;

    /// Validate and own a borrowed locale identifier.
    ///
    /// Validation happens before allocation for borrowed input.
    ///
    /// # Errors
    ///
    /// Returns the corresponding [`Error`] variant when the value is empty,
    /// too long, surrounded by whitespace, or contains a control character.
    pub fn new(value: &str) -> Result<Self> {
        validate_text(value, TextKind::LocaleIdentifier)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain an existing boxed string without reallocating it.
    ///
    /// # Errors
    ///
    /// Returns the corresponding [`Error`] variant when the value is empty,
    /// too long, surrounded by whitespace, or contains a control character.
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_text(&value, TextKind::LocaleIdentifier)?;
        Ok(Self(value))
    }

    /// Borrow the exact locale identifier.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the locale identifier as an owned `String`.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for LocaleIdentifier {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for LocaleIdentifier {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        validate_text(&value, TextKind::LocaleIdentifier)?;
        Ok(Self(value.into_boxed_str()))
    }
}

impl TryFrom<Box<str>> for LocaleIdentifier {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        validate_text(&value, TextKind::LocaleIdentifier)?;
        Ok(Self(value))
    }
}

/// Validated visible text inserted for a Date & Time field.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct DisplayText(Box<str>);

impl DisplayText {
    /// Maximum UTF-8 bytes retained by one display-text value.
    pub const MAX_BYTES: usize = MAX_DISPLAY_TEXT_BYTES;

    /// Validate and own borrowed display text.
    ///
    /// Validation happens before allocation for borrowed input.
    ///
    /// # Errors
    ///
    /// Returns the corresponding [`Error`] variant when the value is empty,
    /// too long, or contains a control character.
    pub fn new(value: &str) -> Result<Self> {
        validate_text(value, TextKind::DisplayText)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain an existing boxed string without reallocating it.
    ///
    /// # Errors
    ///
    /// Returns the corresponding [`Error`] variant when the value is empty,
    /// too long, or contains a control character.
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_text(&value, TextKind::DisplayText)?;
        Ok(Self(value))
    }

    /// Borrow the exact visible text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the display text as an owned `String`.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for DisplayText {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for DisplayText {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        validate_text(&value, TextKind::DisplayText)?;
        Ok(Self(value.into_boxed_str()))
    }
}

impl TryFrom<Box<str>> for DisplayText {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        validate_text(&value, TextKind::DisplayText)?;
        Ok(Self(value))
    }
}

/// Seconds from Apple's `2001-01-01 00:00:00 UTC` reference date.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct Instant(f64);

impl Instant {
    /// Construct an instant from finite Apple reference-date seconds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InstantNotFinite`] for NaN or infinite input.
    pub fn from_reference_date_seconds(seconds: f64) -> Result<Self> {
        if !seconds.is_finite() {
            return Err(Error::InstantNotFinite);
        }
        Ok(Self(seconds))
    }

    /// Return seconds from Apple's UTC reference date.
    #[must_use]
    pub const fn reference_date_seconds(self) -> f64 {
        self.0
    }
}

/// Native date or time formatter style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FormatterStyle {
    /// No style override.
    None,
    /// Short native style.
    Short,
    /// Medium native style.
    Medium,
    /// Long native style.
    Long,
    /// Full native style.
    Full,
    /// A style introduced by a newer iWork release.
    Unknown(i32),
}

impl FormatterStyle {
    /// Decode a native style discriminant losslessly.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            0 => Self::None,
            1 => Self::Short,
            2 => Self::Medium,
            3 => Self::Long,
            4 => Self::Full,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the native style discriminant.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Short => 1,
            Self::Medium => 2,
            Self::Long => 3,
            Self::Full => 4,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether this value uses a named variant for a known value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(self, Self::Unknown(0..=4))
    }
}

/// Native refresh policy for a Date & Time field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum UpdatePlan {
    /// Never refresh the field automatically.
    Never,
    /// Refresh the field automatically.
    Automatic,
    /// Refresh the field once.
    Once,
    /// A plan introduced by a newer iWork release.
    Unknown(i32),
}

impl UpdatePlan {
    /// Decode a native update-plan discriminant losslessly.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            0 => Self::Never,
            1 => Self::Automatic,
            2 => Self::Once,
            unknown => Self::Unknown(unknown),
        }
    }

    /// Return the native update-plan discriminant.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Never => 0,
            Self::Automatic => 1,
            Self::Once => 2,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether this value uses a named variant for a known value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(self, Self::Unknown(0..=2))
    }
}

/// Lossless, validated semantic payload of a Date & Time field.
#[derive(Debug, Clone, PartialEq)]
pub struct Settings {
    format: Option<Format>,
    locale_identifier: Option<LocaleIdentifier>,
    date_style: Option<FormatterStyle>,
    time_style: Option<FormatterStyle>,
    update_plan: Option<UpdatePlan>,
    needs_update: Option<bool>,
    instant: Option<Instant>,
}

impl Settings {
    /// Construct settings from optional native values after validation.
    ///
    /// # Errors
    ///
    /// Returns a non-canonical-value error when a known formatter style or
    /// update plan is represented by its `Unknown` variant.
    pub fn new(
        format: Option<Format>,
        locale_identifier: Option<LocaleIdentifier>,
        date_style: Option<FormatterStyle>,
        time_style: Option<FormatterStyle>,
        update_plan: Option<UpdatePlan>,
        needs_update: Option<bool>,
        instant: Option<Instant>,
    ) -> Result<Self> {
        let settings = Self {
            format,
            locale_identifier,
            date_style,
            time_style,
            update_plan,
            needs_update,
            instant,
        };
        match settings.validate() {
            Ok(()) => Ok(settings),
            Err(error) => Err(error),
        }
    }

    /// Construct settings with every optional native field absent.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            format: None,
            locale_identifier: None,
            date_style: None,
            time_style: None,
            update_plan: None,
            needs_update: None,
            instant: None,
        }
    }

    /// Construct the conventional fixed Date & Time field settings.
    #[must_use]
    pub const fn fixed(
        format: Format,
        locale_identifier: LocaleIdentifier,
        instant: Instant,
    ) -> Self {
        Self {
            format: Some(format),
            locale_identifier: Some(locale_identifier),
            date_style: Some(FormatterStyle::None),
            time_style: Some(FormatterStyle::None),
            update_plan: Some(UpdatePlan::Never),
            needs_update: Some(false),
            instant: Some(instant),
        }
    }

    /// Return the optional ICU-style format.
    #[must_use]
    pub const fn format(&self) -> Option<&Format> {
        self.format.as_ref()
    }

    /// Return the optional locale identifier.
    #[must_use]
    pub const fn locale_identifier(&self) -> Option<&LocaleIdentifier> {
        self.locale_identifier.as_ref()
    }

    /// Return the optional date style.
    #[must_use]
    pub const fn date_style(&self) -> Option<FormatterStyle> {
        self.date_style
    }

    /// Return the optional time style.
    #[must_use]
    pub const fn time_style(&self) -> Option<FormatterStyle> {
        self.time_style
    }

    /// Return the optional refresh policy.
    #[must_use]
    pub const fn update_plan(&self) -> Option<UpdatePlan> {
        self.update_plan
    }

    /// Return the optional display-refresh flag.
    #[must_use]
    pub const fn needs_update(&self) -> Option<bool> {
        self.needs_update
    }

    /// Return the optional Apple-reference-date instant.
    #[must_use]
    pub const fn instant(&self) -> Option<Instant> {
        self.instant
    }

    /// Replace both formatter styles, rejecting non-canonical known values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalFormatterStyle`] when either style uses
    /// `Unknown` for a known native value.
    pub fn with_styles(
        mut self,
        date_style: FormatterStyle,
        time_style: FormatterStyle,
    ) -> Result<Self> {
        self.date_style = Some(date_style);
        self.time_style = Some(time_style);
        match self.validate() {
            Ok(()) => Ok(self),
            Err(error) => Err(error),
        }
    }

    /// Replace the refresh policy, rejecting a non-canonical known value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalUpdatePlan`] when `update_plan` uses
    /// `Unknown` for a known native value.
    pub fn with_update_plan(mut self, update_plan: UpdatePlan) -> Result<Self> {
        self.update_plan = Some(update_plan);
        match self.validate() {
            Ok(()) => Ok(self),
            Err(error) => Err(error),
        }
    }

    /// Validate all semantic invariants before crossing into an adapter.
    ///
    /// # Errors
    ///
    /// Returns a non-canonical-value error when a known formatter style or
    /// update plan is represented by its `Unknown` variant.
    pub const fn validate(&self) -> Result<()> {
        if let Some(style) = self.date_style
            && !style.is_canonical()
        {
            return Err(Error::NonCanonicalFormatterStyle);
        }
        if let Some(style) = self.time_style
            && !style.is_canonical()
        {
            return Err(Error::NonCanonicalFormatterStyle);
        }
        if let Some(plan) = self.update_plan
            && !plan.is_canonical()
        {
            return Err(Error::NonCanonicalUpdatePlan);
        }
        Ok(())
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self::empty()
    }
}

#[derive(Clone, Copy)]
enum TextKind {
    Format,
    LocaleIdentifier,
    DisplayText,
}

fn validate_text(value: &str, kind: TextKind) -> Result<()> {
    let (maximum, empty, too_long, control) = match kind {
        TextKind::Format => (
            MAX_FORMAT_BYTES,
            Error::FormatEmpty,
            Error::FormatTooLong,
            Error::FormatControlCharacter,
        ),
        TextKind::LocaleIdentifier => (
            MAX_LOCALE_IDENTIFIER_BYTES,
            Error::LocaleIdentifierEmpty,
            Error::LocaleIdentifierTooLong,
            Error::LocaleIdentifierControlCharacter,
        ),
        TextKind::DisplayText => (
            MAX_DISPLAY_TEXT_BYTES,
            Error::DisplayTextEmpty,
            Error::DisplayTextTooLong,
            Error::DisplayTextControlCharacter,
        ),
    };
    if value.is_empty() {
        return Err(empty);
    }
    if value.len() > maximum {
        return Err(too_long);
    }
    if value.chars().any(char::is_control) {
        return Err(control);
    }
    if matches!(kind, TextKind::LocaleIdentifier) && value.trim() != value {
        return Err(Error::LocaleIdentifierSurroundingWhitespace);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn text_values_validate_before_owning_and_preserve_owned_capacity() {
        assert_eq!(Format::new(""), Err(Error::FormatEmpty));
        assert_eq!(
            Format::new(&"x".repeat(Format::MAX_BYTES + 1)),
            Err(Error::FormatTooLong)
        );
        assert_eq!(
            LocaleIdentifier::new(" en_US"),
            Err(Error::LocaleIdentifierSurroundingWhitespace)
        );
        assert_eq!(
            DisplayText::new("Friday\nJuly"),
            Err(Error::DisplayTextControlCharacter)
        );

        let owned = String::from("EEEE, MMMM d, y");
        let pointer = owned.as_ptr();
        let format = Format::try_from(owned).unwrap();
        assert_eq!(format.as_str().as_ptr(), pointer);
    }

    #[test]
    fn text_values_reject_controls_and_are_bounded() {
        assert_eq!(
            LocaleIdentifier::new("en\0US"),
            Err(Error::LocaleIdentifierControlCharacter)
        );
        assert_eq!(
            DisplayText::try_from("x".repeat(DisplayText::MAX_BYTES + 1)),
            Err(Error::DisplayTextTooLong)
        );
        assert_eq!(
            Format::try_from(Box::<str>::from("\t")),
            Err(Error::FormatControlCharacter)
        );
    }

    #[test]
    fn instant_and_native_enums_are_strict_and_lossless() {
        assert_eq!(
            Instant::from_reference_date_seconds(f64::NAN),
            Err(Error::InstantNotFinite)
        );
        assert_eq!(
            Instant::from_reference_date_seconds(f64::INFINITY),
            Err(Error::InstantNotFinite)
        );
        for raw in [i32::MIN, -1, 0, 1, 2, 3, 4, 9, i32::MAX] {
            assert_eq!(FormatterStyle::from_raw(raw).as_raw(), raw);
        }
        for raw in [i32::MIN, -1, 0, 1, 2, 3, i32::MAX] {
            assert_eq!(UpdatePlan::from_raw(raw).as_raw(), raw);
        }
        assert!(!FormatterStyle::Unknown(0).is_canonical());
        assert!(!UpdatePlan::Unknown(2).is_canonical());
        assert!(FormatterStyle::Unknown(9).is_canonical());
        assert!(UpdatePlan::Unknown(9).is_canonical());
    }

    #[test]
    fn settings_are_private_field_compositions_with_checked_builders() {
        let settings = Settings::fixed(
            Format::new("EEEE, MMMM d, y").unwrap(),
            LocaleIdentifier::new("en_US").unwrap(),
            Instant::from_reference_date_seconds(805_965_335.0).unwrap(),
        )
        .with_styles(FormatterStyle::Full, FormatterStyle::None)
        .unwrap()
        .with_update_plan(UpdatePlan::Automatic)
        .unwrap();
        assert_eq!(
            settings.format().map(Format::as_str),
            Some("EEEE, MMMM d, y")
        );
        assert_eq!(
            settings.locale_identifier().map(LocaleIdentifier::as_str),
            Some("en_US")
        );
        assert_eq!(settings.date_style(), Some(FormatterStyle::Full));
        assert_eq!(settings.time_style(), Some(FormatterStyle::None));
        assert_eq!(settings.update_plan(), Some(UpdatePlan::Automatic));
        assert_eq!(settings.needs_update(), Some(false));
        assert_eq!(
            settings.instant().map(Instant::reference_date_seconds),
            Some(805_965_335.0)
        );
        assert_eq!(
            Settings::new(
                None,
                None,
                Some(FormatterStyle::Unknown(4)),
                None,
                None,
                None,
                None,
            ),
            Err(Error::NonCanonicalFormatterStyle)
        );
        assert_eq!(
            Settings::new(
                None,
                None,
                None,
                None,
                Some(UpdatePlan::Unknown(1)),
                None,
                None,
            ),
            Err(Error::NonCanonicalUpdatePlan)
        );
    }
}
