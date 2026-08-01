//! Checked worksheet-grid defaults and typographic layout values.

use std::num::NonZeroU64;

use bitflags::bitflags;
use thiserror::Error;

use crate::outline::Outline;

/// SpreadsheetML's effective base-column-width default.
pub const DEFAULT_BASE_WIDTH: u8 = 8;

/// Checked default row height in points.
///
/// Unlike an explicit [`crate::row::Height`], the `sheetFormatPr` schema does
/// not impose Excel's 409-point per-row authoring limit. The semantic facade
/// still rejects negative and non-finite values before they enter a plan.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Height(NonZeroU64);

impl Height {
    /// Validate a finite, non-negative default row height.
    pub const fn new(value: f64) -> std::result::Result<Self, HeightError> {
        if !(value >= 0.0 && value.is_finite()) {
            return Err(HeightError { value });
        }
        let normalized = if value == 0.0 { 0.0 } else { value };
        let Some(encoded) = normalized.to_bits().checked_add(1) else {
            return Err(HeightError { value });
        };
        let Some(encoded) = NonZeroU64::new(encoded) else {
            return Err(HeightError { value });
        };
        Ok(Self(encoded))
    }

    /// Return the height in points.
    pub const fn get(self) -> f64 {
        f64::from_bits(self.0.get() - 1)
    }
}

/// Invalid worksheet default-row height.
#[derive(Debug, Clone, Copy, PartialEq, Error)]
#[error("default row height {value} must be finite and non-negative")]
pub struct HeightError {
    value: f64,
}

impl HeightError {
    /// Rejected numeric value.
    pub const fn value(self) -> f64 {
        self.value
    }
}

/// Convenient checked-or-raw input for [`Height`].
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum HeightAt {
    /// A value that has already been checked.
    Checked(Height),
    /// A raw point value checked when resolved.
    Value(f64),
}

impl HeightAt {
    /// Resolve this input into a checked height.
    pub const fn resolve(self) -> std::result::Result<Height, HeightError> {
        match self {
            Self::Checked(value) => Ok(value),
            Self::Value(value) => Height::new(value),
        }
    }
}

impl From<Height> for HeightAt {
    fn from(value: Height) -> Self {
        Self::Checked(value)
    }
}

impl From<f64> for HeightAt {
    fn from(value: f64) -> Self {
        Self::Value(value)
    }
}

/// Checked worksheet default-column width in character units.
///
/// This is intentionally distinct from [`crate::column::Width`]: Office
/// permits `sheetFormatPr/defaultColWidth` values below 65,536 while an
/// explicit column record is limited to 255.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Width(NonZeroU64);

impl Width {
    /// Validate Office's finite `0..65536` default-column-width domain.
    pub const fn new(value: f64) -> std::result::Result<Self, WidthError> {
        if !(value >= 0.0 && value < 65_536.0) {
            return Err(WidthError { value });
        }
        let normalized = if value == 0.0 { 0.0 } else { value };
        let Some(encoded) = normalized.to_bits().checked_add(1) else {
            return Err(WidthError { value });
        };
        let Some(encoded) = NonZeroU64::new(encoded) else {
            return Err(WidthError { value });
        };
        Ok(Self(encoded))
    }

    /// Return the width in SpreadsheetML character units.
    pub const fn get(self) -> f64 {
        f64::from_bits(self.0.get() - 1)
    }
}

/// Invalid worksheet default-column width.
#[derive(Debug, Clone, Copy, PartialEq, Error)]
#[error("default column width {value} is outside Office's finite range 0..65536")]
pub struct WidthError {
    value: f64,
}

impl WidthError {
    /// Rejected numeric value.
    pub const fn value(self) -> f64 {
        self.value
    }
}

/// Convenient checked-or-raw input for [`Width`].
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum WidthAt {
    /// A value that has already been checked.
    Checked(Width),
    /// A raw character-unit value checked when resolved.
    Value(f64),
}

impl WidthAt {
    /// Resolve this input into a checked width.
    pub const fn resolve(self) -> std::result::Result<Width, WidthError> {
        match self {
            Self::Checked(value) => Ok(value),
            Self::Value(value) => Width::new(value),
        }
    }
}

impl From<Width> for WidthAt {
    fn from(value: Width) -> Self {
        Self::Checked(value)
    }
}

impl From<f64> for WidthAt {
    fn from(value: f64) -> Self {
        Self::Value(value)
    }
}

/// Checked typographic descent in pixels at 100% worksheet zoom.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Descent(NonZeroU64);

impl Descent {
    /// Validate a finite, non-negative typographic descent.
    pub const fn new(value: f64) -> std::result::Result<Self, DescentError> {
        if !(value >= 0.0 && value.is_finite()) {
            return Err(DescentError { value });
        }
        let normalized = if value == 0.0 { 0.0 } else { value };
        let Some(encoded) = normalized.to_bits().checked_add(1) else {
            return Err(DescentError { value });
        };
        let Some(encoded) = NonZeroU64::new(encoded) else {
            return Err(DescentError { value });
        };
        Ok(Self(encoded))
    }

    /// Return the descent in pixels at 100% worksheet zoom.
    pub const fn get(self) -> f64 {
        f64::from_bits(self.0.get() - 1)
    }
}

/// Invalid typographic descent.
#[derive(Debug, Clone, Copy, PartialEq, Error)]
#[error("typographic descent {value} must be finite and non-negative")]
pub struct DescentError {
    value: f64,
}

impl DescentError {
    /// Rejected numeric value.
    pub const fn value(self) -> f64 {
        self.value
    }
}

/// Convenient checked-or-raw input for [`Descent`].
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum DescentAt {
    /// A value that has already been checked.
    Checked(Descent),
    /// A raw pixel value checked when resolved.
    Value(f64),
}

impl DescentAt {
    /// Resolve this input into a checked descent.
    pub const fn resolve(self) -> std::result::Result<Descent, DescentError> {
        match self {
            Self::Checked(value) => Ok(value),
            Self::Value(value) => Descent::new(value),
        }
    }
}

impl From<Descent> for DescentAt {
    fn from(value: Descent) -> Self {
        Self::Checked(value)
    }
}

impl From<f64> for DescentAt {
    fn from(value: f64) -> Self {
        Self::Value(value)
    }
}

macro_rules! numeric_inputs {
    ($target:ty; $($input:ty),+ $(,)?) => {
        $(
            impl From<$input> for $target {
                fn from(value: $input) -> Self {
                    Self::Value(f64::from(value))
                }
            }
        )+
    };
}

numeric_inputs!(HeightAt; f32, u8, u16, u32, i8, i16, i32);
numeric_inputs!(WidthAt; f32, u8, u16, u32, i8, i16, i32);
numeric_inputs!(DescentAt; f32, u8, u16, u32, i8, i16, i32);

bitflags! {
    /// Independently composable worksheet-default facets.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Fields: u8 {
        const BASE_WIDTH = 1 << 0;
        const WIDTH = 1 << 1;
        const HEIGHT = 1 << 2;
        const HIDDEN = 1 << 3;
        const THICK_TOP = 1 << 4;
        const THICK_BOTTOM = 1 << 5;
        const DESCENT = 1 << 6;
    }
}

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub(crate) struct Flags: u8 {
        const CUSTOM_HEIGHT = 1 << 0;
        const HIDDEN = 1 << 1;
        const THICK_TOP = 1 << 2;
        const THICK_BOTTOM = 1 << 3;
    }
}

/// Complete modeled `sheetFormatPr` state.
///
/// Optional attribute presence is retained so reversible edits can distinguish
/// producer defaults from explicitly stored equivalent values. Accessors
/// expose effective behavior and keep that lexical detail out of ordinary use.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Defaults {
    pub(crate) base_width: Option<u8>,
    pub(crate) width: Option<Width>,
    pub(crate) height: Height,
    pub(crate) descent: Option<Descent>,
    pub(crate) row_outline: Option<Outline>,
    pub(crate) column_outline: Option<Outline>,
    pub(crate) flags: Flags,
    pub(crate) present: Flags,
}

impl Defaults {
    /// Effective base column width, defaulting to eight characters.
    pub const fn base_width(&self) -> u8 {
        match self.base_width {
            Some(value) => value,
            None => DEFAULT_BASE_WIDTH,
        }
    }

    /// Explicit producer-stored base width, if present.
    pub const fn stored_base_width(&self) -> Option<u8> {
        self.base_width
    }

    /// Manually stored default column width, if present.
    pub const fn width(&self) -> Option<Width> {
        self.width
    }

    /// Required default row height.
    pub const fn height(&self) -> Height {
        self.height
    }

    /// Whether the default height is custom.
    ///
    /// Microsoft specifies `dyDescent` as making this true even when the core
    /// `customHeight` attribute is absent or explicitly false.
    pub const fn custom_height(&self) -> bool {
        self.flags.contains(Flags::CUSTOM_HEIGHT) || self.descent.is_some()
    }

    /// Whether rows without an explicit record are hidden by default.
    pub const fn hidden(&self) -> bool {
        self.flags.contains(Flags::HIDDEN)
    }

    /// Whether default rows request a thick top edge.
    pub const fn thick_top(&self) -> bool {
        self.flags.contains(Flags::THICK_TOP)
    }

    /// Whether default rows request a thick bottom edge.
    pub const fn thick_bottom(&self) -> bool {
        self.flags.contains(Flags::THICK_BOTTOM)
    }

    /// Producer-reported highest row outline level.
    pub const fn row_outline(&self) -> Outline {
        match self.row_outline {
            Some(value) => value,
            None => Outline::NONE,
        }
    }

    /// Producer-reported highest column outline level.
    pub const fn column_outline(&self) -> Outline {
        match self.column_outline {
            Some(value) => value,
            None => Outline::NONE,
        }
    }

    /// Typographic descent in pixels at 100% worksheet zoom.
    pub const fn descent(&self) -> Option<Descent> {
        self.descent
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numeric_domains_are_const_checked_and_normalize_negative_zero() {
        const HEIGHT: std::result::Result<Height, HeightError> = Height::new(15.0);
        const WIDTH: std::result::Result<Width, WidthError> = Width::new(65_535.5);
        const DESCENT: std::result::Result<Descent, DescentError> = Descent::new(0.25);
        assert_eq!(HEIGHT.map(Height::get), Ok(15.0));
        assert_eq!(WIDTH.map(Width::get), Ok(65_535.5));
        assert_eq!(DESCENT.map(Descent::get), Ok(0.25));
        assert_eq!(Height::new(-0.0).map(Height::get), Ok(0.0));
        assert_eq!(Width::new(-0.0).map(Width::get), Ok(0.0));
        assert_eq!(Descent::new(-0.0).map(Descent::get), Ok(0.0));
        assert!(Height::new(f64::NAN).is_err());
        assert!(Width::new(65_536.0).is_err());
        assert!(Width::new(f64::INFINITY).is_err());
        assert!(Descent::new(-0.1).is_err());
        assert_eq!(std::mem::size_of::<Option<Height>>(), 8);
        assert_eq!(std::mem::size_of::<Option<Width>>(), 8);
        assert_eq!(std::mem::size_of::<Option<Descent>>(), 8);
    }
}
