//! Typed coercion errors.
//!
//! A [`ValueError`] is `Copy` and borrows the rejected text, so building one
//! allocates nothing: probing parses (a query comparison trying whether a
//! cell's text is a length) fail without cost. The message and the "did you
//! mean" suggestion are computed only when the error is displayed or
//! converted into the unified [`Error`].

use std::borrow::Cow;
use std::fmt;

use crate::error::{Code, Error};

use super::color::{self, ColorContext};
use super::suggest;
use super::token::Tokens;

/// What a coercion was reading. It names the value in messages and selects
/// the suggestion vocabulary.
///
/// ```
/// use litchi_ops::value::{Length, LengthContext, What};
///
/// let error = Length::parse_in("-1", LengthContext::LineWidth).unwrap_err();
/// assert_eq!(error.what(), What::LineWidth);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum What {
    /// Unsigned Word spacing in twips.
    WordSpacing,
    /// A number ([`parse_number`](super::parse_number)).
    Number,
    /// An integer ([`parse_integer`](super::parse_integer)).
    Integer,
    /// A boolean ([`parse_bool`](super::parse_bool)).
    Boolean,
    /// A position or extent ([`LengthContext::Emu`](super::LengthContext::Emu),
    /// [`LengthContext::Extent`](super::LengthContext::Extent)).
    Length,
    /// A line width ([`LengthContext::LineWidth`](super::LengthContext::LineWidth)).
    LineWidth,
    /// A font size ([`LengthContext::FontSize`](super::LengthContext::FontSize)).
    FontSize,
    /// A color in the given context.
    Color(ColorContext),
    /// An opacity ([`parse_opacity`](super::parse_opacity)).
    Opacity,
    /// A value of a closed vocabulary.
    Token(&'static Tokens),
}

/// Why a coercion failed.
///
/// ```
/// use litchi_ops::value::{ErrorKind, Length};
///
/// assert_eq!(Length::parse("2em").unwrap_err().kind(), ErrorKind::UnsupportedUnit { unit: "em" });
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ErrorKind {
    /// The text is empty or white space only.
    Empty,
    /// The text is not in any accepted form.
    Syntax,
    /// The number is not finite.
    NotFinite,
    /// The value is outside the accepted range.
    OutOfRange,
    /// A unit suffix with no number before it.
    MissingNumber {
        /// The unit suffix, in its canonical spelling.
        unit: &'static str,
    },
    /// The text before a unit suffix is not a number of the required form.
    BadNumber {
        /// The unit suffix, in its canonical spelling.
        unit: &'static str,
    },
    /// A font- or viewport-relative unit that needs context a coercion does
    /// not have (`em`, `rem`, `ex`, `vw`, `vh`).
    UnsupportedUnit {
        /// The unit suffix, lower case.
        unit: &'static str,
    },
    /// The length does not fit in 64 bits of EMU.
    Overflow,
    /// A negative value where only non-negative values are accepted.
    Negative,
    /// A value above 2 147 483 647 EMU, the 32-bit maximum of OOXML extents
    /// and line widths.
    TooLarge,
    /// A font size written with a comma.
    Comma,
    /// A font size below 0.5 pt (one half-point).
    TooSmall,
    /// An RGB component outside 0 to 255, or an `rgb()` percentage outside
    /// 0% to 100%.
    Component,
    /// An alpha outside 0 to 1, or a percentage alpha outside 0% to 100%.
    Alpha,
    /// An `hsl()` hue that is not a finite number of degrees.
    Hue,
    /// An `hsl()` saturation or lightness that is not a percentage from 0% to
    /// 100%.
    HslPercent,
    /// `auto` where the color is a DrawingML color, which has no automatic
    /// value.
    Auto,
    /// A `+` suffix without a transform name (`accent1+75`).
    TransformName,
    /// A transform name without a value (`accent1+lumMod`).
    TransformValue,
    /// A transform name outside `lumMod`, `lumOff`, `satMod`, `satOff`,
    /// `hueMod`, `hueOff`, `shade`, `tint` and `alpha`.
    UnknownTransform,
    /// A transform value that is not an integer.
    TransformNumber,
    /// A transform value outside its range.
    TransformRange,
    /// More transforms than a [`Color`](super::Color) holds.
    TooManyTransforms,
    /// A `+tint` suffix whose percentage is not a number.
    Tint,
    /// A `+tint` percentage outside -100 to 100.
    TintRange,
    /// Transform suffixes in a context that has none (spreadsheet cells take
    /// only `+tint`; Word text colors take none).
    NoTransforms,
    /// A Word theme tail (`;themeColor=…`), which the color coercion does not
    /// read.
    ThemeTail,
    /// A word outside a closed vocabulary.
    UnknownToken,
    /// An opacity between 1 and 2, which reads both as a fraction and as a
    /// percentage.
    Ambiguous,
}

/// A rejected value: what was being read, why it failed, and the text.
///
/// Converting it into [`Error`] (with `?` or [`ValueError::into_error`])
/// gives code `invalid_value`, the message, and a suggestion when one is
/// known.
///
/// ```
/// use litchi_ops::Code;
/// use litchi_ops::value::{Color, ColorContext};
///
/// let error = Color::parse("reed", ColorContext::Drawing).unwrap_err();
/// assert_eq!(error.code(), Code::InvalidValue);
/// assert_eq!(error.suggestion().as_deref(), Some("did you mean 'red'?"));
/// let error = litchi_ops::Error::from(error);
/// assert_eq!(error.code(), Code::InvalidValue);
/// ```
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ValueError<'a> {
    what: What,
    kind: ErrorKind,
    input: &'a str,
    part: &'a str,
}

impl<'a> ValueError<'a> {
    pub(super) const fn new(what: What, kind: ErrorKind, input: &'a str) -> Self {
        Self {
            what,
            kind,
            input,
            part: input,
        }
    }

    /// Names the part of the input at fault (a component, a transform).
    pub(super) const fn at(mut self, part: &'a str) -> Self {
        self.part = part;
        self
    }

    /// What was being read.
    #[must_use]
    pub const fn what(&self) -> What {
        self.what
    }

    /// Why it failed.
    #[must_use]
    pub const fn kind(&self) -> ErrorKind {
        self.kind
    }

    /// The unified error code: always [`Code::InvalidValue`].
    #[must_use]
    pub const fn code(&self) -> Code {
        Code::InvalidValue
    }

    /// How the caller can fix the value, when there is a known way. This is
    /// computed on demand (edit distance over the vocabulary) and may
    /// allocate.
    #[must_use]
    pub fn suggestion(&self) -> Option<Cow<'static, str>> {
        match (self.what, self.kind) {
            (What::Color(context), ErrorKind::Syntax) => {
                if let Some(hint) = color::context_hint(context, self.input) {
                    return Some(hint.into());
                }
                let near = suggest::nearest_keyed(self.input, color::names(context))?;
                Some(format!("did you mean '{near}'?").into())
            },
            (What::Color(_), ErrorKind::UnknownTransform) => {
                let name = transform_name(self.part);
                let near = suggest::nearest(name, color::TRANSFORM_NAMES.iter().copied());
                Some(match near {
                    Some(near) => format!("did you mean '{near}'?").into(),
                    None => TRANSFORM_LIST.into(),
                })
            },
            (What::Color(_), ErrorKind::TransformName | ErrorKind::TransformValue) => {
                Some(TRANSFORM_LIST.into())
            },
            (What::Color(_), ErrorKind::Auto) => {
                Some("'auto' applies to Word text colors; use a hex, named or theme color".into())
            },
            (What::Color(ColorContext::Spreadsheet), ErrorKind::NoTransforms) => {
                Some("cell colors accept a '+tint<percent>' suffix, e.g. 'accent1+tint40'".into())
            },
            (What::Color(ColorContext::Word), ErrorKind::NoTransforms | ErrorKind::ThemeTail) => {
                Some("give a hex, named or theme color such as 'accent1' or 'auto'".into())
            },
            (What::Token(tokens), ErrorKind::UnknownToken) => {
                let near = suggest::nearest_keyed(self.input, tokens.spellings())?;
                Some(format!("did you mean '{near}'?").into())
            },
            (What::Boolean, ErrorKind::Syntax) => {
                let near =
                    suggest::nearest(self.input, ["true", "false", "yes", "no", "on", "off"])?;
                Some(format!("did you mean '{near}'?").into())
            },
            (What::Length | What::LineWidth, ErrorKind::UnsupportedUnit { .. }) => {
                Some("use cm, mm, in, pt, pc, px, Q or emu".into())
            },
            (What::Length, ErrorKind::Syntax) if looks_decimal(self.input) => Some(
                "add a unit such as 'cm' or 'pt'; a bare number is a whole number of EMU".into(),
            ),
            _ => None,
        }
    }

    /// Converts into the unified error: code `invalid_value`, this error's
    /// message, and its suggestion.
    #[must_use]
    pub fn into_error(self) -> Error {
        let error = Error::new(Code::InvalidValue, self.to_string());
        match self.suggestion() {
            Some(suggestion) => error.with_suggestion(suggestion),
            None => error,
        }
    }

    /// Converts into the unified error, naming the property the value was
    /// given for (`invalid 'opacity' value: …`).
    #[must_use]
    pub fn into_error_for(self, property: &'static str) -> Error {
        let error = Error::new(
            Code::InvalidValue,
            format!("invalid '{property}' value: {self}"),
        );
        match self.suggestion() {
            Some(suggestion) => error.with_suggestion(suggestion),
            None => error,
        }
    }
}

impl From<ValueError<'_>> for Error {
    fn from(error: ValueError<'_>) -> Self {
        error.into_error()
    }
}

impl std::error::Error for ValueError<'_> {}

const TRANSFORM_LIST: &str =
    "color transforms are lumMod, lumOff, satMod, satOff, hueMod, hueOff, shade, tint and alpha";

/// The leading run of ASCII letters of a transform token.
fn transform_name(token: &str) -> &str {
    let end = token
        .bytes()
        .position(|byte| !byte.is_ascii_alphabetic())
        .unwrap_or(token.len());
    &token[..end]
}

/// `true` for a plain decimal such as `1.5`, which is not an EMU count.
fn looks_decimal(text: &str) -> bool {
    let text = text.trim();
    text.contains('.')
        && text
            .trim_start_matches(['+', '-'])
            .bytes()
            .all(|byte| byte.is_ascii_digit() || byte == b'.')
}

impl fmt::Display for ValueError<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let input = self.input;
        let part = self.part;
        match (self.what, self.kind) {
            (What::WordSpacing, _) => f.write_str("invalid Word spacing: expected a finite nonnegative value within unsigned 32-bit twips, optionally in pt, cm or in"),
            (What::Number, ErrorKind::Empty) => f.write_str("invalid number: the value is empty"),
            (What::Number, ErrorKind::NotFinite) => {
                write!(f, "invalid number: it must be finite")
            },
            (What::Number, _) => write!(
                f,
                "invalid number: expected a finite decimal such as 12, -1.5 or 2e3"
            ),
            (What::Integer, ErrorKind::Empty) => f.write_str("invalid integer: the value is empty"),
            (What::Integer, _) => write!(f, "invalid integer: expected whole digits"),
            (What::Boolean, _) => write!(
                f,
                "invalid boolean: expected true/false, yes/no, 1/0 or on/off"
            ),
            (What::Length | What::LineWidth, kind) => {
                let noun = if self.what == What::LineWidth {
                    "line width"
                } else {
                    "length"
                };
                write_length_message(f, noun, input, part, kind)
            },
            (What::FontSize, kind) => write_font_size_message(f, input, kind),
            (What::Color(_), kind) => write_color_message(f, input, part, kind),
            (What::Opacity, ErrorKind::Ambiguous) => write!(
                f,
                "invalid opacity: values between 1 and 2 are ambiguous; give 0.0-1.0 or 2-100 as a percentage"
            ),
            (What::Opacity, _) => write!(
                f,
                "invalid opacity: expected 0.0-1.0, or 2-100 as a percentage"
            ),
            (What::Token(tokens), _) => {
                write!(f, "invalid {}: expected one of ", tokens.name())?;
                for (index, value) in tokens.values().iter().enumerate() {
                    if index > 0 {
                        f.write_str(", ")?;
                    }
                    f.write_str(value)?;
                }
                Ok(())
            },
        }
    }
}

fn write_length_message(
    f: &mut fmt::Formatter<'_>,
    noun: &str,
    _input: &str,
    _part: &str,
    kind: ErrorKind,
) -> fmt::Result {
    match kind {
        ErrorKind::Empty => write!(f, "invalid {noun}: the value is empty"),
        ErrorKind::MissingNumber { unit } => {
            write!(f, "invalid {noun}: no number before the '{unit}' unit")
        },
        ErrorKind::BadNumber { unit: "emu" } => write!(
            f,
            "invalid {noun}: the component before the 'emu' unit is not a whole number"
        ),
        ErrorKind::BadNumber { unit } => write!(
            f,
            "invalid {noun}: the component before the '{unit}' unit is not a finite number"
        ),
        ErrorKind::UnsupportedUnit { unit } => write!(
            f,
            "invalid {noun}: the unit '{unit}' needs font or viewport context"
        ),
        ErrorKind::NotFinite => write!(f, "invalid {noun}: it must be finite"),
        ErrorKind::Overflow => write!(f, "invalid {noun}: it exceeds the EMU range"),
        ErrorKind::Negative => write!(f, "invalid {noun}: it must not be negative"),
        ErrorKind::TooLarge => write!(
            f,
            "invalid {noun}: it exceeds the maximum of 2147483647 EMU"
        ),
        _ => write!(
            f,
            "invalid {noun}: expected a number with an optional unit (cm, mm, in, pt, pc, px, Q, emu)"
        ),
    }
}

fn write_font_size_message(
    f: &mut fmt::Formatter<'_>,
    _input: &str,
    kind: ErrorKind,
) -> fmt::Result {
    match kind {
        ErrorKind::Comma => write!(
            f,
            "invalid font size: use '.' as the decimal separator, not ','"
        ),
        ErrorKind::Negative => {
            write!(f, "invalid font size: it must be greater than 0")
        },
        ErrorKind::TooSmall => {
            write!(f, "invalid font size: the minimum is 0.5pt")
        },
        ErrorKind::TooLarge => {
            write!(f, "invalid font size: the maximum is 4000pt")
        },
        _ => write!(
            f,
            "invalid font size: expected a finite number of points such as 12, 10.5 or 14pt"
        ),
    }
}

fn write_color_message(
    f: &mut fmt::Formatter<'_>,
    _input: &str,
    part: &str,
    kind: ErrorKind,
) -> fmt::Result {
    match kind {
        ErrorKind::Empty => f.write_str("invalid color: the value is empty"),
        ErrorKind::Component => write!(
            f,
            "invalid color: RGB component the component must be 0-255 or 0%-100%"
        ),
        ErrorKind::Alpha => write!(
            f,
            "invalid color: alpha the component must be 0-1 (e.g. 0.5) or 0%-100%"
        ),
        ErrorKind::Hue => write!(
            f,
            "invalid color: hue the component must be a finite number of degrees"
        ),
        ErrorKind::HslPercent => write!(
            f,
            "invalid color: saturation and lightness (the component) must be percentages from 0% to 100%"
        ),
        ErrorKind::Auto => write!(f, "invalid color: DrawingML colors have no 'auto'"),
        ErrorKind::TransformName => write!(
            f,
            "invalid color: the suffix a transform has no transform name"
        ),
        ErrorKind::TransformValue => {
            write!(f, "invalid color: the transform the component has no value")
        },
        ErrorKind::UnknownTransform => write!(f, "invalid color: unknown color transform"),
        ErrorKind::TransformNumber => write!(
            f,
            "invalid color: the value of the component must be an integer"
        ),
        ErrorKind::TransformRange => {
            let (percent, raw) = color::transform_range_text(transform_name(part));
            write!(
                f,
                "invalid color: the component is out of range ({percent} as a percentage, {raw} after '=')"
            )
        },
        ErrorKind::TooManyTransforms => write!(
            f,
            "invalid color: more than {} color transforms",
            super::Color::MAX_TRANSFORMS
        ),
        ErrorKind::Tint => write!(
            f,
            "invalid color: the tint the component is not a percentage such as 'tint40' or 'tint-25'"
        ),
        ErrorKind::TintRange => write!(
            f,
            "invalid color: the tint percentage must be between -100 and 100"
        ),
        ErrorKind::NoTransforms => {
            write!(f, "invalid color: this color takes no a transform suffix")
        },
        ErrorKind::ThemeTail => write!(
            f,
            "invalid color: theme tails such as ';themeColor=...' are not color values"
        ),
        _ => write!(
            f,
            "invalid color: expected hex (F00, FF0000, #RRGGBBAA, AARRGGBB, #RGBA), a named color (red), rgb()/rgba()/hsl()/hsla(), or a theme color (accent1)"
        ),
    }
}

impl fmt::Debug for ValueError<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("ValueError").finish_non_exhaustive()
    }
}
