//! Archive-free semantic values for native chart reference lines.

const MAXIMUM_NAME_BYTES: usize = 1_024;
const SHOW_NAME: u8 = 0b01;
const SHOW_VALUE: u8 = 0b10;

const NATIVE_MINIMUM: i32 = 1;
const NATIVE_MAXIMUM: i32 = 2;
const NATIVE_AVERAGE: i32 = 3;
const NATIVE_MEDIAN: i32 = 4;
const NATIVE_CUSTOM: i32 = 5;

/// Construction failures for chart reference-line values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A custom reference-line value was not finite.
    #[error("chart reference-line value must be finite")]
    NonFiniteValue,
    /// A reference-line name was empty.
    #[error("chart reference-line name must not be empty")]
    EmptyName,
    /// A reference-line name contains a control character.
    #[error("chart reference-line name contains control character U+{character:04X}")]
    ControlCharacter { character: u32 },
    /// A reference-line name exceeded the bounded semantic budget.
    #[error("chart reference-line name uses {bytes} UTF-8 bytes, maximum is {maximum}")]
    NameTooLong { bytes: usize, maximum: usize },
    /// A known native type was supplied through the lossless unsupported form.
    #[error("known chart reference-line type {native_type} must use its named representation")]
    KnownTypeAsUnsupported { native_type: i32 },
}

/// Result type for reference-line value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A finite custom value used to position a chart reference line.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Value(f64);

impl Value {
    /// Validate and construct a custom reference-line value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFiniteValue`] for NaN and infinities.
    #[must_use = "use the validated value or handle the construction error"]
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::NonFiniteValue);
        }
        Ok(Self(value))
    }

    /// Return the finite numeric value.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for Value {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::new(value)
    }
}

/// Native calculation used to position one chart reference line.
#[derive(Clone, Copy, Debug, PartialEq)]
#[non_exhaustive]
pub enum Kind {
    /// The minimum value in the plotted data.
    Minimum,
    /// The maximum value in the plotted data.
    Maximum,
    /// The average value in the plotted data.
    Average,
    /// The median value in the plotted data.
    Median,
    /// A caller-supplied finite value.
    Custom(Value),
    /// A future native type retained without interpretation.
    Unsupported(Unknown),
}

/// A checked, lossless representation of an unrecognized native kind.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Unknown {
    native_type: i32,
    custom_value: Option<Value>,
}

impl Unknown {
    /// Construct a future native kind without permitting known values to be
    /// smuggled through the lossless representation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownTypeAsUnsupported`] when `native_type` is one of
    /// the known native identifiers.
    #[must_use = "use the validated unknown kind or handle the construction error"]
    pub fn new(native_type: i32, custom_value: Option<Value>) -> Result<Self> {
        if matches!(
            native_type,
            NATIVE_MINIMUM | NATIVE_MAXIMUM | NATIVE_AVERAGE | NATIVE_MEDIAN | NATIVE_CUSTOM
        ) {
            return Err(Error::KnownTypeAsUnsupported { native_type });
        }
        Ok(Self {
            native_type,
            custom_value,
        })
    }

    /// Return the unrecognized native calculation identifier.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.native_type
    }

    /// Return the optional native custom-value payload.
    #[must_use]
    pub const fn custom_value(self) -> Option<Value> {
        self.custom_value
    }
}

impl Kind {
    /// Construct and validate an unsupported native calculation identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownTypeAsUnsupported`] when `native_type` is one of
    /// the known identifiers.
    #[must_use = "use the validated kind or handle the construction error"]
    pub fn unsupported(native_type: i32, custom_value: Option<Value>) -> Result<Self> {
        Ok(Self::Unsupported(Unknown::new(native_type, custom_value)?))
    }

    /// Return the native calculation identifier.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Minimum => NATIVE_MINIMUM,
            Self::Maximum => NATIVE_MAXIMUM,
            Self::Average => NATIVE_AVERAGE,
            Self::Median => NATIVE_MEDIAN,
            Self::Custom(_) => NATIVE_CUSTOM,
            Self::Unsupported(unknown) => unknown.native_value(),
        }
    }

    /// Return the optional custom-value payload represented by this kind.
    #[must_use]
    pub const fn custom_value(self) -> Option<Value> {
        match self {
            Self::Custom(value) => Some(value),
            Self::Unsupported(unknown) => unknown.custom_value(),
            Self::Minimum | Self::Maximum | Self::Average | Self::Median => None,
        }
    }

    /// Return the native default label for this kind.
    #[must_use]
    pub const fn default_name(self) -> &'static str {
        match self {
            Self::Minimum => "Minimum",
            Self::Maximum => "Maximum",
            Self::Average => "Average",
            Self::Median => "Median",
            Self::Custom(_) => "Custom",
            Self::Unsupported(_) => "Reference Line",
        }
    }

    /// Return whether the native default displays the numeric value.
    #[must_use]
    pub const fn default_shows_value(self) -> bool {
        matches!(self, Self::Custom(_))
    }

    fn validate(self) -> Result<()> {
        if let Self::Unsupported(unknown) = self {
            Unknown::new(unknown.native_type, unknown.custom_value)?;
        }
        Ok(())
    }
}

/// Complete semantic configuration for one native chart reference line.
///
/// Default names are borrowed from [`Kind`] and therefore do not allocate.
/// Explicit names are bounded and stored in one boxed string. The two
/// visibility settings share one byte.
#[derive(Clone, Debug, PartialEq)]
pub struct Line {
    kind: Kind,
    name: Option<Box<str>>,
    visibility: u8,
}

impl Line {
    /// Create a minimum reference line with native defaults.
    #[must_use]
    pub fn minimum() -> Self {
        Self::from_known_kind(Kind::Minimum)
    }

    /// Create a maximum reference line with native defaults.
    #[must_use]
    pub fn maximum() -> Self {
        Self::from_known_kind(Kind::Maximum)
    }

    /// Create an average reference line with native defaults.
    #[must_use]
    pub fn average() -> Self {
        Self::from_known_kind(Kind::Average)
    }

    /// Create a median reference line with native defaults.
    #[must_use]
    pub fn median() -> Self {
        Self::from_known_kind(Kind::Median)
    }

    /// Create a custom reference line with native defaults.
    #[must_use]
    pub fn custom(value: Value) -> Self {
        Self::from_known_kind(Kind::Custom(value))
    }

    /// Create a reference line from a validated kind.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownTypeAsUnsupported`] for an invalid lossless
    /// unsupported representation.
    #[must_use = "use the validated line or handle the construction error"]
    pub fn from_kind(kind: Kind) -> Result<Self> {
        kind.validate()?;
        Ok(Self::from_known_kind(kind))
    }

    /// Create a reference line while preserving an unrecognized native kind.
    ///
    /// # Errors
    ///
    /// Returns [`Error::KnownTypeAsUnsupported`] when the identifier is known.
    #[must_use = "use the validated line or handle the construction error"]
    pub fn unsupported(native_type: i32, custom_value: Option<Value>) -> Result<Self> {
        Self::from_kind(Kind::unsupported(native_type, custom_value)?)
    }

    /// Return the native calculation kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the effective label, including the native default when no
    /// explicit label was supplied.
    #[must_use]
    pub fn name(&self) -> &str {
        self.name
            .as_deref()
            .unwrap_or_else(|| self.kind.default_name())
    }

    /// Return the explicitly stored label, if one was supplied.
    #[must_use]
    pub fn explicit_name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Whether iWork renders the line's label.
    #[must_use]
    pub const fn shows_name(&self) -> bool {
        self.visibility & SHOW_NAME != 0
    }

    /// Whether iWork renders the line's numeric value.
    #[must_use]
    pub const fn shows_value(&self) -> bool {
        self.visibility & SHOW_VALUE != 0
    }

    /// Replace the label with a bounded explicit name.
    ///
    /// The native default name is normalized back to the allocation-free
    /// default representation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyName`] for an empty name or
    /// [`Error::NameTooLong`] above [`Self::MAXIMUM_NAME_BYTES`].
    #[must_use = "use the updated line or handle the name validation error"]
    pub fn try_with_name(mut self, new_name: impl AsRef<str>) -> Result<Self> {
        let name = new_name.as_ref();
        if name.is_empty() {
            return Err(Error::EmptyName);
        }
        if let Some(character) = name.chars().find(|character| character.is_control()) {
            return Err(Error::ControlCharacter {
                character: u32::from(character),
            });
        }
        if name.len() > Self::MAXIMUM_NAME_BYTES {
            return Err(Error::NameTooLong {
                bytes: name.len(),
                maximum: Self::MAXIMUM_NAME_BYTES,
            });
        }
        self.name = (name != self.kind.default_name()).then(|| name.into());
        Ok(self)
    }

    /// Maximum UTF-8 bytes accepted for an explicit reference-line name.
    pub const MAXIMUM_NAME_BYTES: usize = MAXIMUM_NAME_BYTES;

    /// Set whether iWork renders the line's label.
    #[must_use]
    pub const fn with_name_visibility(mut self, visible: bool) -> Self {
        if visible {
            self.visibility |= SHOW_NAME;
        } else {
            self.visibility &= !SHOW_NAME;
        }
        self
    }

    /// Set whether iWork renders the line's numeric value.
    #[must_use]
    pub const fn with_value_visibility(mut self, visible: bool) -> Self {
        if visible {
            self.visibility |= SHOW_VALUE;
        } else {
            self.visibility &= !SHOW_VALUE;
        }
        self
    }

    /// Validate the semantic kind and its bounded state.
    ///
    /// # Errors
    ///
    /// Returns a typed validation error for an invalid kind.
    pub fn validate(&self) -> Result<()> {
        self.kind.validate()
    }

    fn from_known_kind(kind: Kind) -> Self {
        let mut visibility = SHOW_NAME;
        if kind.default_shows_value() {
            visibility |= SHOW_VALUE;
        }
        Self {
            kind,
            name: None,
            visibility,
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "Common semantic tests use expect for fixed valid fixtures"
)]
mod tests {
    use std::mem::size_of;

    use super::{Error, Kind, Line, Value};

    #[test]
    fn values_are_finite_and_compact() {
        assert_eq!(size_of::<Value>(), size_of::<f64>());
        assert!(Value::new(17.5).is_ok());
        assert_eq!(Value::new(f64::NAN), Err(Error::NonFiniteValue));
        assert_eq!(Value::new(f64::INFINITY), Err(Error::NonFiniteValue));
    }

    #[test]
    fn default_names_do_not_allocate_and_visibility_is_packed() {
        assert!(size_of::<Line>() <= 48);
        let line = Line::average();
        assert_eq!(line.name(), "Average");
        assert!(line.shows_name());
        assert!(!line.shows_value());
        let custom = Line::custom(Value::new(17.5).expect("finite test value"));
        assert!(custom.shows_value());
    }

    #[test]
    fn explicit_names_are_bounded_and_default_names_normalize() {
        let line = Line::average()
            .try_with_name("Middle")
            .expect("valid reference-line name");
        assert_eq!(line.name(), "Middle");
        let normalized = line
            .try_with_name("Average")
            .expect("native default name is valid");
        assert_eq!(normalized, Line::average());
        assert_eq!(Line::average().try_with_name(""), Err(Error::EmptyName));
        assert_eq!(
            Line::average().try_with_name("Middle\n"),
            Err(Error::ControlCharacter { character: 0x0a })
        );
        assert!(
            Line::average()
                .try_with_name("x".repeat(Line::MAXIMUM_NAME_BYTES))
                .is_ok()
        );
        assert!(matches!(
            Line::average().try_with_name("x".repeat(Line::MAXIMUM_NAME_BYTES + 1)),
            Err(Error::NameTooLong { .. })
        ));
    }

    #[test]
    fn unsupported_kinds_are_lossless_but_known_values_are_strict() {
        let kind = Kind::unsupported(-7, Some(Value::new(2.0).expect("finite test value")))
            .expect("future native value");
        assert_eq!(kind.native_value(), -7);
        assert_eq!(kind.custom_value().map(Value::value), Some(2.0));
        assert_eq!(
            Kind::unsupported(1, None),
            Err(Error::KnownTypeAsUnsupported { native_type: 1 })
        );
        assert!(Line::from_kind(kind).is_ok());
    }
}
