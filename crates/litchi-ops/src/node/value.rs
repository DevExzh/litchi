//! Typed node-view values (record 0010 §2.3): what one `format` entry holds,
//! and how each value is written.
//!
//! Values keep OfficeCLI's JSON types: booleans as booleans, counts and
//! indices as numbers, positions and extents as unit-qualified strings,
//! colours as `#RRGGBB` or scheme references. Lengths and colours are the
//! canonical [`crate::value`] types, so a value prints exactly the way the
//! coercion that reads it in M2 parses it. A [`Length`] is stored in EMU; the
//! property's [`ValueType`] picks how it is written (`"2.54cm"` for a
//! position, `"11pt"` for a font size), so the node view needs no per-node
//! strings.

use std::borrow::Cow;
use std::fmt;

use serde::ser::{Serialize, Serializer};

use crate::value::{Color, Length, Scalar, ValueType};

/// A `format` key: the OfficeCLI property name, which is also the key `set`
/// accepts in M2 (`bold`, `font.latin`, `x`, `cachedValue`).
///
/// ```
/// use litchi_ops::node::Key;
///
/// let key: Key = "font.latin";
/// assert_eq!(key, "font.latin");
/// ```
pub type Key = &'static str;

/// One typed value of a node's `format` map.
///
/// The variants mirror the JSON types OfficeCLI writes: [`Value::Bool`] is a
/// JSON boolean, [`Value::Int`] and [`Value::Number`] are JSON numbers, and
/// every other variant is a JSON string. The text of a value is its
/// [`Scalar`]'s canonical text ([`Value::as_scalar`]), which is the form `set`
/// accepts in M2 and the form query comparison reads; a [`Length`] prints in
/// OfficeCLI's centimetre-first form there, while the node view writes it in
/// the style its key's [`ValueType`] calls for.
///
/// ```
/// use litchi_ops::node::Value;
/// use litchi_ops::value::{Color, Length};
///
/// assert_eq!(Value::from(true).to_string(), "true");
/// assert_eq!(Value::Int(3).to_string(), "3");
/// assert_eq!(Value::from("Heading1").to_string(), "Heading1");
/// assert_eq!(Value::from(Length::from_emu(914_400)).to_string(), "2.54cm");
/// assert_eq!(Value::from(Color::rgb(0x1F, 0x4E, 0x79)).to_string(), "#1F4E79");
/// assert_eq!(Value::number(f64::NAN), None);
/// ```
#[derive(Clone, PartialEq)]
#[non_exhaustive]
pub enum Value<'a> {
    /// A flag. An explicitly-off flag is `false`; an unset one is absent.
    Bool(bool),
    /// A count, index or identifier.
    Int(i64),
    /// A finite measurement without a unit. Build it with [`Value::number`],
    /// which refuses NaN and infinities; a non-finite value is refused during serialization. In JSON an integral value is written without a fractional
    /// part, as OfficeCLI's `System.Text.Json` writes doubles (`1`, not
    /// `1.0`); as text it is .NET's shortest round-trip form
    /// ([`DotNetDouble`](crate::value::DotNetDouble)).
    Number(f64),
    /// Free text: a style id, a name, a formula.
    Text(Cow<'a, str>),
    /// A token from a closed vocabulary (`single`, `center`).
    Token(&'static str),
    /// A length in EMU.
    Length(Length),
    /// A colour.
    Color(Color),
    /// A list of text items, a JSON array of strings (the metadata
    /// keywords, record 0354).
    List(Cow<'a, [String]>),
}

impl<'a> Value<'a> {
    /// A [`Value::Number`], or `None` when `value` is NaN or infinite.
    ///
    /// ```
    /// use litchi_ops::node::Value;
    ///
    /// assert_eq!(Value::number(1.5), Some(Value::Number(1.5)));
    /// assert_eq!(Value::number(f64::INFINITY), None);
    /// ```
    #[must_use]
    pub fn number(value: f64) -> Option<Self> {
        value.is_finite().then_some(Self::Number(value))
    }

    /// An owned [`Value::List`] of `items`, or `None` for an empty list: an
    /// empty list is an absent property, as the keyword list's `Vec` is.
    #[cfg_attr(
        not(any(
            feature = "docx",
            feature = "pptx",
            feature = "xls",
            feature = "xlsx",
            feature = "xlsb",
            feature = "odt",
            feature = "ods",
            feature = "odp",
            feature = "odi"
        )),
        allow(dead_code, reason = "only the metadata adapters read keyword lists")
    )]
    pub(crate) fn owned_list(items: &[String]) -> Option<Value<'static>> {
        (!items.is_empty()).then(|| Value::List(Cow::Owned(items.to_vec())))
    }

    /// The same value, borrowing its text instead of cloning it. Never
    /// allocates.
    ///
    /// ```
    /// use std::borrow::Cow;
    ///
    /// use litchi_ops::node::Value;
    ///
    /// let owned = Value::Text(Cow::Owned("Normal".to_owned()));
    /// assert!(matches!(owned.borrowed(), Value::Text(Cow::Borrowed("Normal"))));
    /// ```
    #[must_use]
    pub fn borrowed(&self) -> Value<'_> {
        match self {
            Self::Bool(value) => Value::Bool(*value),
            Self::Int(value) => Value::Int(*value),
            Self::Number(value) => Value::Number(*value),
            Self::Text(text) => Value::Text(Cow::Borrowed(text)),
            Self::Token(token) => Value::Token(token),
            Self::Length(length) => Value::Length(*length),
            Self::Color(color) => Value::Color(*color),
            Self::List(items) => Value::List(Cow::Borrowed(items)),
        }
    }

    /// The same value with owned text, independent of the node it came from.
    ///
    /// ```
    /// use litchi_ops::node::Value;
    ///
    /// let text = String::from("Title");
    /// let owned: Value<'static> = Value::from(text.as_str()).into_owned();
    /// drop(text);
    /// assert_eq!(owned.to_string(), "Title");
    /// ```
    #[must_use]
    pub fn into_owned(self) -> Value<'static> {
        match self {
            Self::Bool(value) => Value::Bool(value),
            Self::Int(value) => Value::Int(value),
            Self::Number(value) => Value::Number(value),
            Self::Text(text) => Value::Text(Cow::Owned(text.into_owned())),
            Self::Token(token) => Value::Token(token),
            Self::Length(length) => Value::Length(length),
            Self::Color(color) => Value::Color(color),
            Self::List(items) => Value::List(Cow::Owned(items.into_owned())),
        }
    }

    /// The text of a [`Value::Text`] or [`Value::Token`].
    ///
    /// ```
    /// use litchi_ops::node::Value;
    ///
    /// assert_eq!(Value::Token("center").as_str(), Some("center"));
    /// assert_eq!(Value::Bool(true).as_str(), None);
    /// ```
    #[must_use]
    pub fn as_str(&self) -> Option<&str> {
        match self {
            Self::Text(text) => Some(text),
            Self::Token(token) => Some(token),
            _ => None,
        }
    }

    /// The value as a borrowed [`Scalar`]: the input of query comparison
    /// ([`Test::eval`](crate::value::Test::eval)) and of every canonical
    /// rendering. Never allocates.
    ///
    /// ```
    /// use std::borrow::Cow;
    ///
    /// use litchi_ops::node::Value;
    /// use litchi_ops::value::Scalar;
    ///
    /// let style = Value::Text(Cow::Owned("Heading1".to_owned()));
    /// assert_eq!(style.as_scalar(), Scalar::Text("Heading1"));
    /// assert_eq!(Value::Int(3).as_scalar(), Scalar::Int(3));
    /// ```
    #[must_use]
    pub fn as_scalar(&self) -> Scalar<'_> {
        match self {
            Self::Bool(value) => Scalar::Bool(*value),
            Self::Int(value) => Scalar::Int(*value),
            Self::Number(value) => Scalar::Number(*value),
            Self::Text(text) => Scalar::Text(text),
            Self::Token(token) => Scalar::Token(token),
            Self::Length(length) => Scalar::Length(*length),
            Self::Color(color) => Scalar::Color(*color),
            Self::List(items) => Scalar::List(items),
        }
    }

    /// Writes the value's canonical text as a property of type `ty` prints
    /// it (a length in the type's style); `None` prints lengths
    /// centimetre-first.
    pub(crate) fn fmt_typed(
        &self,
        ty: Option<ValueType>,
        f: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        fmt::Display::fmt(&self.as_scalar().display(ty), f)
    }

    /// Serializes the value with its JSON type, rendering a length the way a
    /// property of type `ty` prints it. Strings are streamed into the
    /// serializer, never collected.
    pub(crate) fn serialize_typed<S: Serializer>(
        &self,
        ty: Option<ValueType>,
        serializer: S,
    ) -> Result<S::Ok, S::Error> {
        match self {
            Self::Bool(value) => serializer.serialize_bool(*value),
            Self::Int(value) => serializer.serialize_i64(*value),
            Self::Number(value) if !value.is_finite() => {
                Err(serde::ser::Error::custom("node number must be finite"))
            },
            Self::Number(value) => match integral(*value) {
                Some(value) => serializer.serialize_i64(value),
                None => serializer.serialize_f64(*value),
            },
            Self::Text(text) => serializer.serialize_str(text),
            Self::Token(token) => serializer.serialize_str(token),
            Self::Length(_) | Self::Color(_) => {
                let scalar = self.as_scalar();
                serializer.collect_str(&scalar.display(ty))
            },
            Self::List(items) => serializer.collect_seq(items.iter()),
        }
    }
}

impl fmt::Display for Value<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.fmt_typed(None, f)
    }
}

/// Serializes the value with its JSON type; a length is written in the
/// centimetre-first style. Inside a node view, the key's own style applies.
impl Serialize for Value<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        self.serialize_typed(None, serializer)
    }
}

impl From<bool> for Value<'_> {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<i64> for Value<'_> {
    fn from(value: i64) -> Self {
        Self::Int(value)
    }
}

impl From<i32> for Value<'_> {
    fn from(value: i32) -> Self {
        Self::Int(value.into())
    }
}

impl From<u32> for Value<'_> {
    fn from(value: u32) -> Self {
        Self::Int(value.into())
    }
}

impl<'a> From<&'a str> for Value<'a> {
    fn from(value: &'a str) -> Self {
        Self::Text(Cow::Borrowed(value))
    }
}

impl From<String> for Value<'_> {
    fn from(value: String) -> Self {
        Self::Text(Cow::Owned(value))
    }
}

impl<'a> From<Cow<'a, str>> for Value<'a> {
    fn from(value: Cow<'a, str>) -> Self {
        Self::Text(value)
    }
}

impl From<Vec<String>> for Value<'_> {
    fn from(items: Vec<String>) -> Self {
        Self::List(Cow::Owned(items))
    }
}

impl<'a> From<&'a [String]> for Value<'a> {
    fn from(items: &'a [String]) -> Self {
        Self::List(Cow::Borrowed(items))
    }
}

impl From<Length> for Value<'_> {
    fn from(value: Length) -> Self {
        Self::Length(value)
    }
}

impl From<Color> for Value<'_> {
    fn from(value: Color) -> Self {
        Self::Color(value)
    }
}

/// A coerced value as a node value: what a leaf reader that produces
/// [`Scalar`]s hands the node view. A [`Scalar::Number`] that is not finite
/// stays a [`Value::Number`] and is refused during serialization.
///
/// ```
/// use litchi_ops::node::Value;
/// use litchi_ops::value::{Scalar, ValueType};
///
/// let scalar = ValueType::Bool.coerce("on")?;
/// assert_eq!(Value::from(scalar), Value::Bool(true));
/// assert_eq!(Value::from(Scalar::Text("Normal")).as_str(), Some("Normal"));
/// # Ok::<(), litchi_ops::Error>(())
/// ```
impl<'a> From<Scalar<'a>> for Value<'a> {
    fn from(value: Scalar<'a>) -> Self {
        match value {
            Scalar::Bool(value) => Self::Bool(value),
            Scalar::Int(value) => Self::Int(value),
            Scalar::Number(value) => Self::Number(value),
            Scalar::Text(text) => Self::Text(Cow::Borrowed(text)),
            Scalar::Token(token) => Self::Token(token),
            Scalar::Length(length) => Self::Length(length),
            Scalar::Color(color) => Self::Color(color),
            Scalar::List(items) => Self::List(Cow::Borrowed(items)),
        }
    }
}

/// The integer an integral finite double equals, when it is exactly
/// representable (|value| ≤ 2^53).
fn integral(value: f64) -> Option<i64> {
    const LIMIT: f64 = 9_007_199_254_740_992.0; // 2^53
    // Integral and within ±2^53, so the cast is exact.
    (value.is_finite() && value.fract() == 0.0 && value.abs() <= LIMIT).then_some(value as i64)
}

impl fmt::Debug for Value<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Value").finish_non_exhaustive()
    }
}
