//! Value coercions: numbers, booleans, lengths, colors and closed
//! vocabularies, their canonical text, and the typed comparison query
//! predicates run on.
//!
//! The module is shared by query compilation (record 0010 §5.2) and the
//! property registry (design §6): a query literal and a `set` value go
//! through the same coercion, so what a property prints is what it accepts.
//! The rules are OfficeCLI's (`EmuConverter`, `ParseHelpers`,
//! `DrawingColorBuilder`, `AttributeFilter`), with the departures recorded in
//! the unified-ops change record for this module.
//!
//! | Value | Read with | Canonical text |
//! |---|---|---|
//! | Length | [`Length::parse_in`] and a [`LengthContext`] | [`Length::display`] with a [`LengthStyle`]: `2.54cm`, `960pt`, `11pt`, `1800emu` |
//! | Color | [`Color::parse`] and a [`ColorContext`] | `#RRGGBB`, `#RRGGBBAA`, `accent1+lumMod75`, `lt1+tint40`, `auto` |
//! | Boolean | [`parse_bool`] | `true`, `false` |
//! | Number | [`parse_number`], [`parse_integer`], [`parse_numeric_text`] | [`DotNetDouble`], [`Decimal`] |
//! | Opacity | [`parse_opacity`] | a fraction from 0 to 1 |
//! | Enumeration | [`Tokens::parse`] | the vocabulary's canonical token |
//!
//! [`ValueType`] names a property's scalar type and coerces text to a
//! [`Scalar`]; [`Test`] compiles one query condition and evaluates it
//! against scalars.
//!
//! # Allocation
//!
//! Length, color, boolean and token parses borrow their input without allocating; the
//! canonical text is written straight into the caller's formatter. Grouped numeric
//! text longer than 128 bytes uses an allocating normalization fallback. Tables
//! are `static` and searched by binary search. A failed parse is a
//! [`ValueError`]: `Copy`, borrowing the input, and allocation-free to build;
//! its message and "did you mean" suggestion are computed only when it is
//! displayed or converted into the unified [`Error`](crate::Error) (code
//! `invalid_value`).
//!
//! # Example
//!
//! ```
//! use litchi_ops::value::{Color, ColorContext, Length, LengthContext, LengthStyle, parse_bool};
//!
//! let x = Length::parse_in("1in", LengthContext::Emu)?;
//! assert_eq!(x.emu(), 914_400);
//! assert_eq!(x.to_string(), "2.54cm");
//! assert_eq!(Length::parse_in("0.75", LengthContext::LineWidth)?.display(LengthStyle::LineWidth).to_string(), "0.75pt");
//!
//! let fill = Color::parse("accent1+lumMod75", ColorContext::Drawing)?;
//! assert_eq!(fill.to_string(), "accent1+lumMod75");
//! assert_eq!(Color::parse("hsl(120, 100%, 25%)", ColorContext::Drawing)?.to_string(), "#008000");
//!
//! assert!(parse_bool("on")?);
//!
//! let error = litchi_ops::Error::from(Color::parse("bleu", ColorContext::Drawing).unwrap_err());
//! assert_eq!(error.code(), litchi_ops::Code::InvalidValue);
//! assert_eq!(error.suggestion(), Some("did you mean 'blue'?"));
//! # Ok::<(), litchi_ops::Error>(())
//! ```

mod boolean;
mod color;
pub(crate) mod compare;
mod error;
mod length;
mod number;
pub(crate) mod suggest;
pub(crate) mod text;
mod token;
mod ty;

#[cfg(test)]
mod tests;

pub use boolean::{parse_bool, parse_opacity};
pub use color::{
    Base, Color, ColorContext, Rgba, Scheme, System, ThemeSlot, Transform, TransformKind,
};
pub use compare::{Test, Warning};
pub use error::{ErrorKind, ValueError, What};
pub use length::{Length, LengthContext, LengthDisplay, LengthSpec, LengthStyle};
pub use number::{Decimal, DotNetDouble, parse_integer, parse_number, parse_numeric_text};
pub use token::Tokens;
pub use ty::{Scalar, ScalarDisplay, ValueType};
