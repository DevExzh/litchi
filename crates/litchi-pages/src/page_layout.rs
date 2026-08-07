//! Validated, archive-free Pages page-layout values.

use std::fmt;

use thiserror::Error;

const PAGE_WIDTH: usize = 0;
const PAGE_HEIGHT: usize = 1;
const LEFT_MARGIN: usize = 2;
const RIGHT_MARGIN: usize = 3;
const TOP_MARGIN: usize = 4;
const BOTTOM_MARGIN: usize = 5;
const HEADER_MARGIN: usize = 6;
const FOOTER_MARGIN: usize = 7;
const PAGE_SCALE: usize = 8;
const PAGE_WIDTH_PRESENT: u16 = 1 << PAGE_WIDTH;
const PAGE_HEIGHT_PRESENT: u16 = 1 << PAGE_HEIGHT;
const LEFT_MARGIN_PRESENT: u16 = 1 << LEFT_MARGIN;
const RIGHT_MARGIN_PRESENT: u16 = 1 << RIGHT_MARGIN;
const TOP_MARGIN_PRESENT: u16 = 1 << TOP_MARGIN;
const BOTTOM_MARGIN_PRESENT: u16 = 1 << BOTTOM_MARGIN;
const HEADER_MARGIN_PRESENT: u16 = 1 << HEADER_MARGIN;
const FOOTER_MARGIN_PRESENT: u16 = 1 << FOOTER_MARGIN;
const PAGE_SCALE_PRESENT: u16 = 1 << PAGE_SCALE;
const ORIENTATION_PRESENT: u16 = 1 << 9;
const VERTICAL_BODY_PRESENT: u16 = 1 << 10;
const RAW_ORIENTATION_PORTRAIT: u32 = 0;
const RAW_ORIENTATION_LANDSCAPE: u32 = 1;

/// A named page-layout scalar used by typed validation errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Field {
    /// Physical page width.
    PageWidth,
    /// Physical page height.
    PageHeight,
    /// Left body margin.
    LeftMargin,
    /// Right body margin.
    RightMargin,
    /// Top body margin.
    TopMargin,
    /// Bottom body margin.
    BottomMargin,
    /// Header margin.
    HeaderMargin,
    /// Footer margin.
    FooterMargin,
    /// Page scale.
    PageScale,
}

impl fmt::Display for Field {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PageWidth => "page width",
            Self::PageHeight => "page height",
            Self::LeftMargin => "left margin",
            Self::RightMargin => "right margin",
            Self::TopMargin => "top margin",
            Self::BottomMargin => "bottom margin",
            Self::HeaderMargin => "header margin",
            Self::FooterMargin => "footer margin",
            Self::PageScale => "page scale",
        })
    }
}

/// Validation failures for Pages page-layout values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A positive scalar was NaN or infinite.
    #[error("Pages {field} must be finite and greater than zero")]
    NonPositive { field: Field },
    /// A margin scalar was NaN or infinite.
    #[error("Pages {field} must be finite and non-negative")]
    Negative { field: Field },
    /// Horizontal margins consume the complete page width.
    #[error("Pages horizontal margins must leave positive body width")]
    HorizontalMargins,
    /// Vertical margins consume the complete page height.
    #[error("Pages vertical margins must leave positive body height")]
    VerticalMargins,
    /// A known native orientation was represented by `Unknown`.
    #[error("Pages orientation must use its canonical variant for a known value")]
    NonCanonicalOrientation,
}

/// Result type for page-layout construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Physical orientation of a Pages document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Orientation {
    /// Portrait page orientation.
    Portrait,
    /// Landscape page orientation.
    Landscape,
    /// A value written by a newer Pages version.
    Unknown(u32),
}

impl Orientation {
    /// Decode a lossless native orientation value.
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            RAW_ORIENTATION_PORTRAIT => Self::Portrait,
            RAW_ORIENTATION_LANDSCAPE => Self::Landscape,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown orientation without shadowing a named variant.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalOrientation`] when `raw` is already
    /// assigned to a named native value.
    pub const fn unknown(raw: u32) -> Result<Self> {
        match raw {
            RAW_ORIENTATION_PORTRAIT | RAW_ORIENTATION_LANDSCAPE => {
                Err(Error::NonCanonicalOrientation)
            },
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the lossless native orientation value.
    #[must_use]
    pub const fn as_raw(self) -> u32 {
        match self {
            Self::Portrait => RAW_ORIENTATION_PORTRAIT,
            Self::Landscape => RAW_ORIENTATION_LANDSCAPE,
            Self::Unknown(raw) => raw,
        }
    }

    /// Return whether an `Unknown` value is canonical for its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(RAW_ORIENTATION_PORTRAIT | RAW_ORIENTATION_LANDSCAPE)
        )
    }
}

/// Validated page dimensions, margins, scale, and orientation.
///
/// Native field presence is retained in a compact bitset. All scalar values
/// are private and can enter the value only through finite, range-checked
/// construction or setters.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Layout {
    values: [f32; 9],
    orientation: u32,
    present: u16,
    vertical_body: u8,
}

impl Layout {
    /// Construct a page layout while retaining each native field's presence.
    ///
    /// # Errors
    ///
    /// Returns a typed error when a scalar is non-finite, outside its allowed
    /// range, when margins consume a page, or when an orientation is not
    /// canonical.
    #[allow(
        clippy::too_many_arguments,
        reason = "The constructor mirrors the complete native page-layout record."
    )]
    pub fn new(
        page_width: Option<f32>,
        page_height: Option<f32>,
        left_margin: Option<f32>,
        right_margin: Option<f32>,
        top_margin: Option<f32>,
        bottom_margin: Option<f32>,
        header_margin: Option<f32>,
        footer_margin: Option<f32>,
        page_scale: Option<f32>,
        orientation: Option<Orientation>,
        lays_out_body_vertically: Option<bool>,
    ) -> Result<Self> {
        if orientation.is_some_and(|candidate| !candidate.is_canonical()) {
            return Err(Error::NonCanonicalOrientation);
        }
        let mut layout = Self {
            values: [0.0; 9],
            orientation: orientation.map_or(RAW_ORIENTATION_PORTRAIT, Orientation::as_raw),
            present: 0,
            vertical_body: u8::from(lays_out_body_vertically.unwrap_or(false)),
        };
        layout.assign(PAGE_WIDTH, PAGE_WIDTH_PRESENT, page_width);
        layout.assign(PAGE_HEIGHT, PAGE_HEIGHT_PRESENT, page_height);
        layout.assign(LEFT_MARGIN, LEFT_MARGIN_PRESENT, left_margin);
        layout.assign(RIGHT_MARGIN, RIGHT_MARGIN_PRESENT, right_margin);
        layout.assign(TOP_MARGIN, TOP_MARGIN_PRESENT, top_margin);
        layout.assign(BOTTOM_MARGIN, BOTTOM_MARGIN_PRESENT, bottom_margin);
        layout.assign(HEADER_MARGIN, HEADER_MARGIN_PRESENT, header_margin);
        layout.assign(FOOTER_MARGIN, FOOTER_MARGIN_PRESENT, footer_margin);
        layout.assign(PAGE_SCALE, PAGE_SCALE_PRESENT, page_scale);
        if orientation.is_some() {
            layout.present |= ORIENTATION_PRESENT;
        }
        if lays_out_body_vertically.is_some() {
            layout.present |= VERTICAL_BODY_PRESENT;
        }
        layout.validate()?;
        Ok(layout)
    }

    /// Return an empty layout with all native fields absent.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            values: [0.0; 9],
            orientation: RAW_ORIENTATION_PORTRAIT,
            present: 0,
            vertical_body: 0,
        }
    }

    /// Return the optional page width.
    #[must_use]
    pub const fn page_width(self) -> Option<f32> {
        self.value(PAGE_WIDTH, PAGE_WIDTH_PRESENT)
    }

    /// Return the optional page height.
    #[must_use]
    pub const fn page_height(self) -> Option<f32> {
        self.value(PAGE_HEIGHT, PAGE_HEIGHT_PRESENT)
    }

    /// Return the optional left margin.
    #[must_use]
    pub const fn left_margin(self) -> Option<f32> {
        self.value(LEFT_MARGIN, LEFT_MARGIN_PRESENT)
    }

    /// Return the optional right margin.
    #[must_use]
    pub const fn right_margin(self) -> Option<f32> {
        self.value(RIGHT_MARGIN, RIGHT_MARGIN_PRESENT)
    }

    /// Return the optional top margin.
    #[must_use]
    pub const fn top_margin(self) -> Option<f32> {
        self.value(TOP_MARGIN, TOP_MARGIN_PRESENT)
    }

    /// Return the optional bottom margin.
    #[must_use]
    pub const fn bottom_margin(self) -> Option<f32> {
        self.value(BOTTOM_MARGIN, BOTTOM_MARGIN_PRESENT)
    }

    /// Return the optional header margin.
    #[must_use]
    pub const fn header_margin(self) -> Option<f32> {
        self.value(HEADER_MARGIN, HEADER_MARGIN_PRESENT)
    }

    /// Return the optional footer margin.
    #[must_use]
    pub const fn footer_margin(self) -> Option<f32> {
        self.value(FOOTER_MARGIN, FOOTER_MARGIN_PRESENT)
    }

    /// Return the optional page scale.
    #[must_use]
    pub const fn page_scale(self) -> Option<f32> {
        self.value(PAGE_SCALE, PAGE_SCALE_PRESENT)
    }

    /// Return the optional physical orientation.
    #[must_use]
    pub const fn orientation(self) -> Option<Orientation> {
        if self.present & ORIENTATION_PRESENT != 0 {
            Some(Orientation::from_raw(self.orientation))
        } else {
            None
        }
    }

    /// Return the optional vertical body-layout flag.
    #[must_use]
    pub const fn lays_out_body_vertically(self) -> Option<bool> {
        if self.present & VERTICAL_BODY_PRESENT != 0 {
            Some(self.vertical_body != 0)
        } else {
            None
        }
    }

    /// Replace the optional page width.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is invalid or makes the page margins
    /// consume the available width.
    pub fn set_page_width(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(PAGE_WIDTH, PAGE_WIDTH_PRESENT, value)
    }

    /// Replace the optional page height.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is invalid or makes the page margins
    /// consume the available height.
    pub fn set_page_height(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(PAGE_HEIGHT, PAGE_HEIGHT_PRESENT, value)
    }

    /// Replace the optional left margin.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is negative or consumes the page width.
    pub fn set_left_margin(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(LEFT_MARGIN, LEFT_MARGIN_PRESENT, value)
    }

    /// Replace the optional right margin.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is negative or consumes the page width.
    pub fn set_right_margin(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(RIGHT_MARGIN, RIGHT_MARGIN_PRESENT, value)
    }

    /// Replace the optional top margin.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is negative or consumes the page height.
    pub fn set_top_margin(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(TOP_MARGIN, TOP_MARGIN_PRESENT, value)
    }

    /// Replace the optional bottom margin.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is negative or consumes the page height.
    pub fn set_bottom_margin(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(BOTTOM_MARGIN, BOTTOM_MARGIN_PRESENT, value)
    }

    /// Replace the optional header margin.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is negative.
    pub fn set_header_margin(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(HEADER_MARGIN, HEADER_MARGIN_PRESENT, value)
    }

    /// Replace the optional footer margin.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is negative.
    pub fn set_footer_margin(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(FOOTER_MARGIN, FOOTER_MARGIN_PRESENT, value)
    }

    /// Replace the optional page scale.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is invalid.
    pub fn set_page_scale(&mut self, value: Option<f32>) -> Result<()> {
        self.set_value(PAGE_SCALE, PAGE_SCALE_PRESENT, value)
    }

    /// Replace the optional physical orientation.
    ///
    /// # Errors
    ///
    /// Returns an error when an `Unknown` value uses a known native
    /// discriminant.
    pub fn set_orientation(&mut self, value: Option<Orientation>) -> Result<()> {
        if value.is_some_and(|candidate| !candidate.is_canonical()) {
            return Err(Error::NonCanonicalOrientation);
        }
        let before = *self;
        self.orientation = value.map_or(RAW_ORIENTATION_PORTRAIT, Orientation::as_raw);
        if value.is_some() {
            self.present |= ORIENTATION_PRESENT;
        } else {
            self.present &= !ORIENTATION_PRESENT;
        }
        if let Err(error) = self.validate() {
            *self = before;
            return Err(error);
        }
        Ok(())
    }

    /// Replace the optional vertical body-layout flag.
    pub fn set_lays_out_body_vertically(&mut self, value: Option<bool>) {
        self.vertical_body = u8::from(value.unwrap_or(false));
        if value.is_some() {
            self.present |= VERTICAL_BODY_PRESENT;
        } else {
            self.present &= !VERTICAL_BODY_PRESENT;
        }
    }

    /// Validate all stored scalar and cross-field invariants.
    ///
    /// # Errors
    ///
    /// Returns an error when a stored scalar is invalid, margins consume the
    /// page, or an orientation is not canonical.
    pub fn validate(self) -> Result<()> {
        for (index, bit, field, positive) in [
            (PAGE_WIDTH, PAGE_WIDTH_PRESENT, Field::PageWidth, true),
            (PAGE_HEIGHT, PAGE_HEIGHT_PRESENT, Field::PageHeight, true),
            (LEFT_MARGIN, LEFT_MARGIN_PRESENT, Field::LeftMargin, false),
            (
                RIGHT_MARGIN,
                RIGHT_MARGIN_PRESENT,
                Field::RightMargin,
                false,
            ),
            (TOP_MARGIN, TOP_MARGIN_PRESENT, Field::TopMargin, false),
            (
                BOTTOM_MARGIN,
                BOTTOM_MARGIN_PRESENT,
                Field::BottomMargin,
                false,
            ),
            (
                HEADER_MARGIN,
                HEADER_MARGIN_PRESENT,
                Field::HeaderMargin,
                false,
            ),
            (
                FOOTER_MARGIN,
                FOOTER_MARGIN_PRESENT,
                Field::FooterMargin,
                false,
            ),
            (PAGE_SCALE, PAGE_SCALE_PRESENT, Field::PageScale, true),
        ] {
            if self.present & bit == 0 {
                continue;
            }
            let value = self.values[index];
            if !value.is_finite() {
                return Err(if positive {
                    Error::NonPositive { field }
                } else {
                    Error::Negative { field }
                });
            }
            if positive {
                if value <= 0.0 {
                    return Err(Error::NonPositive { field });
                }
            } else if value < 0.0 {
                return Err(Error::Negative { field });
            }
        }

        if self
            .orientation()
            .is_some_and(|orientation| !orientation.is_canonical())
        {
            return Err(Error::NonCanonicalOrientation);
        }
        if let Some(width) = self.page_width()
            && !leaves_positive_space(width, self.left_margin(), self.right_margin())
        {
            return Err(Error::HorizontalMargins);
        }
        if let Some(height) = self.page_height()
            && !leaves_positive_space(height, self.top_margin(), self.bottom_margin())
        {
            return Err(Error::VerticalMargins);
        }
        Ok(())
    }

    fn assign(&mut self, index: usize, bit: u16, value: Option<f32>) {
        if let Some(candidate) = value {
            self.values[index] = candidate;
            self.present |= bit;
        }
    }

    fn set_value(&mut self, index: usize, bit: u16, value: Option<f32>) -> Result<()> {
        let before = *self;
        if let Some(candidate) = value {
            self.values[index] = candidate;
            self.present |= bit;
        } else {
            self.values[index] = 0.0;
            self.present &= !bit;
        }
        if let Err(error) = self.validate() {
            *self = before;
            return Err(error);
        }
        Ok(())
    }

    const fn value(self, index: usize, bit: u16) -> Option<f32> {
        if self.present & bit != 0 {
            Some(self.values[index])
        } else {
            None
        }
    }
}

impl Default for Layout {
    fn default() -> Self {
        Self::empty()
    }
}

/// Check the positive body-space invariant without adding two `f32` values.
///
/// An absent native margin has the scalar default of zero. Checking each
/// margin against the page dimension first keeps the subtraction finite and
/// avoids accepting an overflowing sum as a valid layout.
fn leaves_positive_space(page: f32, first_margin: Option<f32>, second_margin: Option<f32>) -> bool {
    let first = first_margin.unwrap_or(0.0);
    let second = second_margin.unwrap_or(0.0);
    first < page && second < page && first < page - second
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Error, Layout, Orientation};

    #[test]
    fn layout_is_compact_and_retains_presence() {
        assert_eq!(size_of::<Layout>(), 44);
        let layout = Layout::new(
            Some(612.0),
            Some(792.0),
            Some(72.0),
            Some(72.0),
            Some(54.0),
            Some(54.0),
            None,
            None,
            Some(1.0),
            Some(Orientation::Landscape),
            Some(false),
        )
        .unwrap_or_else(|error| panic!("valid layout: {error}"));
        assert_eq!(layout.page_width(), Some(612.0));
        assert_eq!(layout.header_margin(), None);
        assert_eq!(layout.orientation(), Some(Orientation::Landscape));
        assert_eq!(layout.lays_out_body_vertically(), Some(false));
    }

    #[test]
    fn layout_rejects_invalid_scalars_margins_and_unknown_aliases() {
        assert!(matches!(
            Layout::new(
                Some(f32::NAN),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            Err(Error::NonPositive { .. })
        ));
        assert!(matches!(
            Layout::new(
                Some(100.0),
                None,
                Some(60.0),
                Some(40.0),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            Err(Error::HorizontalMargins)
        ));
        assert!(matches!(
            Layout::new(
                Some(100.0),
                None,
                Some(100.0),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            Err(Error::HorizontalMargins)
        ));
        assert!(matches!(
            Layout::new(
                Some(f32::MAX),
                None,
                Some(f32::MAX),
                Some(f32::MAX),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            Err(Error::HorizontalMargins)
        ));
        assert_eq!(Orientation::unknown(0), Err(Error::NonCanonicalOrientation));
        assert!(!Orientation::Unknown(1).is_canonical());
    }

    #[test]
    fn rejected_margin_setters_leave_the_previous_layout_unchanged() {
        let mut layout = Layout::new(
            Some(100.0),
            None,
            Some(10.0),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        )
        .unwrap_or_else(|error| panic!("valid layout: {error}"));

        assert_eq!(
            layout.set_left_margin(Some(100.0)),
            Err(Error::HorizontalMargins)
        );
        assert_eq!(layout.left_margin(), Some(10.0));
    }
}
