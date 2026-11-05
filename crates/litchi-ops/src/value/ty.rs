//! Value types and borrowed typed values.

use std::fmt;

use super::boolean::parse_bool;
use super::color::{Color, ColorContext};
use super::error::ValueError;
use super::length::{Length, LengthSpec, LengthStyle};
use super::number::{DotNetDouble, parse_integer, parse_number};
use super::token::Tokens;

/// The scalar type of a property: how its text is coerced and how its value
/// prints. A property registry entry (M2) and a node view key (M1) carry one;
/// a query compiles each literal against the key's type once.
///
/// ```
/// use litchi_ops::value::{LengthSpec, Scalar, ValueType};
///
/// let width = ValueType::Length(LengthSpec::EXTENT);
/// assert_eq!(width.schema_type(), "length");
/// let value = width.coerce("1in")?;
/// assert_eq!(value.display(Some(width)).to_string(), "2.54cm");
/// assert!(matches!(ValueType::Bool.coerce("on")?, Scalar::Bool(true)));
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ValueType {
    /// A boolean ([`parse_bool`](super::parse_bool)).
    Bool,
    /// A 64-bit integer ([`parse_integer`](super::parse_integer)).
    Int,
    /// A finite number ([`parse_number`](super::parse_number)).
    Number,
    /// Free text, taken as given.
    Text,
    /// Text, boolean, 64-bit integer or finite number, preserving the supplied scalar kind.
    /// Text input stays literal; it is never inferred as a number or formula.
    Scalar,
    /// A closed vocabulary.
    Token(&'static Tokens),
    /// A length, read and printed per the [`LengthSpec`].
    Length(LengthSpec),
    /// A color in a [`ColorContext`].
    Color(ColorContext),
    /// A list of text items (the metadata keywords, record 0354). Text
    /// coerces to itself: the format splits it at `;` and `,`.
    List,
}

impl ValueType {
    /// The type name OfficeCLI's help schemas use: `bool`, `int`, `number`,
    /// `string`, `enum`, `length`, `font-size` or `color`; `scalar` describes
    /// the typed text/boolean/integer/number union.
    #[must_use]
    pub const fn schema_type(self) -> &'static str {
        match self {
            Self::Bool => "bool",
            Self::Int => "int",
            Self::Number => "number",
            Self::Text => "string",
            Self::Scalar => "scalar",
            Self::Token(_) => "enum",
            Self::Length(spec) => match spec.context {
                super::LengthContext::FontSize => "font-size",
                _ => "length",
            },
            Self::Color(_) => "color",
            Self::List => "list",
        }
    }

    /// Coerces `text` to a value of this type. Allocation-free.
    ///
    /// # Errors
    ///
    /// Returns the type's `invalid_value` [`ValueError`] when the text does
    /// not coerce.
    pub fn coerce(self, text: &str) -> Result<Scalar<'_>, ValueError<'_>> {
        Ok(match self {
            Self::Bool => Scalar::Bool(parse_bool(text)?),
            Self::Int => Scalar::Int(parse_integer(text)?),
            Self::Number => Scalar::Number(parse_number(text)?),
            Self::Text | Self::Scalar | Self::List => Scalar::Text(text),
            Self::Token(tokens) => Scalar::Token(tokens.parse(text)?),
            Self::Length(spec) => Scalar::Length(Length::parse_in(text, spec.context)?),
            Self::Color(context) => Scalar::Color(Color::parse(text, context)?),
        })
    }

    /// The length style values of this type print in; [`LengthStyle::Auto`]
    /// for types that are not lengths.
    pub(super) const fn length_style(this: Option<Self>) -> LengthStyle {
        match this {
            Some(Self::Length(spec)) => spec.style,
            _ => LengthStyle::Auto,
        }
    }
}

/// A typed value, borrowing its text: the input of query comparison and
/// the output of coercion.
///
/// [`Display`](fmt::Display) prints the canonical form with default length
/// rendering; [`Scalar::display`] prints it the way a property of a given
/// [`ValueType`] prints it.
///
/// ```
/// use litchi_ops::value::{Length, LengthSpec, Scalar, ValueType};
///
/// let inch = Scalar::Length(Length::from_points(72));
/// assert_eq!(inch.to_string(), "2.54cm");
/// let as_font_size = inch.display(Some(ValueType::Length(LengthSpec::FONT_SIZE)));
/// assert_eq!(as_font_size.to_string(), "72pt");
/// assert_eq!(Scalar::Number(1e20).to_string(), "1E+20");
/// ```
#[derive(Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum Scalar<'a> {
    /// A boolean: `true` or `false`.
    Bool(bool),
    /// An integer.
    Int(i64),
    /// A finite number, printed as .NET prints doubles.
    Number(f64),
    /// Text.
    Text(&'a str),
    /// A canonical token of a closed vocabulary.
    Token(&'static str),
    /// A length.
    Length(Length),
    /// A color.
    Color(Color),
    /// A list of text items, printed joined with `, `.
    List(&'a [String]),
}

impl Scalar<'_> {
    /// The canonical text of this value as a property of type `ty` prints it
    /// (a length in the type's [`LengthStyle`]). Allocation-free.
    #[must_use]
    pub fn display(&self, ty: Option<ValueType>) -> ScalarDisplay<'_> {
        ScalarDisplay {
            scalar: self,
            style: ValueType::length_style(ty),
        }
    }
}

impl fmt::Display for Scalar<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.display(None).fmt(formatter)
    }
}

/// The canonical text of a [`Scalar`] in a property's rendering.
///
/// ```
/// use litchi_ops::value::Scalar;
///
/// assert_eq!(Scalar::Bool(true).display(None).to_string(), "true");
/// ```
#[derive(Debug, Clone, Copy)]
pub struct ScalarDisplay<'s> {
    scalar: &'s Scalar<'s>,
    style: LengthStyle,
}

impl fmt::Display for ScalarDisplay<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match *self.scalar {
            Scalar::Bool(value) => formatter.write_str(if value { "true" } else { "false" }),
            Scalar::Int(value) => write!(formatter, "{value}"),
            Scalar::Number(value) => DotNetDouble(value).fmt(formatter),
            Scalar::Text(text) => formatter.write_str(text),
            Scalar::Token(token) => formatter.write_str(token),
            Scalar::Length(length) => length.display(self.style).fmt(formatter),
            Scalar::Color(color) => color.fmt(formatter),
            Scalar::List(items) => {
                for (index, item) in items.iter().enumerate() {
                    if index > 0 {
                        formatter.write_str(", ")?;
                    }
                    formatter.write_str(item)?;
                }
                Ok(())
            },
        }
    }
}

impl fmt::Debug for Scalar<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Scalar").finish_non_exhaustive()
    }
}
