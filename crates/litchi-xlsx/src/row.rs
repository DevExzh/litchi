//! Borrowed worksheet row views and checked layout properties.

use std::slice;

use bitflags::bitflags;
use litchi_sheet::Row as Index;
use thiserror::Error;

use crate::layout::Descent;
pub use crate::outline::{Outline, OutlineAt, OutlineError};
use crate::style::StyleState;

/// Checked SpreadsheetML row height in points.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Height(u64);

impl Height {
    /// Validate Excel's finite `0..=409` point row-height domain.
    pub const fn new(value: f64) -> Result<Self, HeightError> {
        if !(value >= 0.0 && value <= 409.0) {
            return Err(HeightError { value });
        }
        let normalized = if value == 0.0 { 0.0 } else { value };
        Ok(Self(normalized.to_bits()))
    }

    /// Return the height in points.
    pub const fn get(self) -> f64 {
        f64::from_bits(self.0)
    }
}

/// Invalid SpreadsheetML row height.
#[derive(Debug, Clone, Copy, PartialEq, Error)]
#[error("row height {value} is outside Excel's finite range 0..=409 points")]
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
    /// A height that has already been checked.
    Checked(Height),
    /// A raw point value validated when resolved.
    Value(f64),
}

impl HeightAt {
    /// Resolve this input into a checked height.
    pub const fn resolve(self) -> Result<Height, HeightError> {
        match self {
            Self::Checked(height) => Ok(height),
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

macro_rules! height_inputs {
    ($($input:ty),+ $(,)?) => {
        $(
            impl From<$input> for HeightAt {
                fn from(value: $input) -> Self {
                    Self::Value(f64::from(value))
                }
            }
        )+
    };
}

height_inputs!(f32, u8, u16, u32, i8, i16, i32);

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub(crate) struct Flags: u8 {
        const HIDDEN = 1 << 0;
        const CUSTOM_HEIGHT = 1 << 1;
        const COLLAPSED = 1 << 2;
        const THICK_TOP = 1 << 3;
        const THICK_BOTTOM = 1 << 4;
        const PHONETIC = 1 << 5;
        const CUSTOM_FORMAT = 1 << 6;
    }
}

/// Complete modeled properties of one stored SpreadsheetML row record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Properties {
    pub(crate) height: Option<Height>,
    pub(crate) descent: Option<Descent>,
    pub(crate) style: Option<u32>,
    pub(crate) outline: Outline,
    pub(crate) flags: Flags,
}

/// Complete semantic state of one stored row record.
///
/// Physical shared-style indexes remain hidden behind [`StyleState`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Props {
    pub(crate) height: Option<Height>,
    pub(crate) descent: Option<Descent>,
    pub(crate) style: StyleState,
    pub(crate) outline: Outline,
    pub(crate) flags: Flags,
}

impl Props {
    /// Producer-stored row height, if present.
    pub const fn height(&self) -> Option<Height> {
        self.height
    }

    /// Producer-stored typographic descent at 100% worksheet zoom.
    pub const fn descent(&self) -> Option<Descent> {
        self.descent
    }

    /// Exact local shared-style state without a physical style index.
    pub const fn style(&self) -> &StyleState {
        &self.style
    }

    /// Checked outline level.
    pub const fn outline(&self) -> Outline {
        self.outline
    }

    /// Whether the row is hidden.
    pub const fn hidden(&self) -> bool {
        self.flags.contains(Flags::HIDDEN)
    }

    /// Whether the row height is explicitly customized.
    pub const fn custom_height(&self) -> bool {
        self.flags.contains(Flags::CUSTOM_HEIGHT) || self.descent.is_some()
    }

    /// Whether the row's outline is stored in its collapsed state.
    pub const fn collapsed(&self) -> bool {
        self.flags.contains(Flags::COLLAPSED)
    }

    /// Whether the row requests a thick top edge.
    pub const fn thick_top(&self) -> bool {
        self.flags.contains(Flags::THICK_TOP)
    }

    /// Whether the row requests a thick bottom edge.
    pub const fn thick_bottom(&self) -> bool {
        self.flags.contains(Flags::THICK_BOTTOM)
    }

    /// Whether phonetic information is shown by default.
    pub const fn phonetic(&self) -> bool {
        self.flags.contains(Flags::PHONETIC)
    }

    /// Whether the producer explicitly applies the row style.
    pub const fn custom_format(&self) -> bool {
        self.flags.contains(Flags::CUSTOM_FORMAT)
    }

    pub(crate) fn rebind_style(&mut self, lineage: &std::sync::Arc<crate::style::StyleLineage>) {
        self.style.rebind(lineage);
    }

    pub(crate) const fn uses_shared_style(&self) -> bool {
        matches!(&self.style, StyleState::Shared(_))
    }
}

/// Row-record state captured before or after a patch change.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    /// No explicit row record exists.
    Missing,
    /// One complete row record exists.
    Stored(Props),
}

impl State {
    pub(crate) fn rebind_style(&mut self, lineage: &std::sync::Arc<crate::style::StyleLineage>) {
        if let Self::Stored(properties) = self {
            properties.rebind_style(lineage);
        }
    }

    pub(crate) const fn uses_shared_style(&self) -> bool {
        matches!(self, Self::Stored(properties) if properties.uses_shared_style())
    }
}

/// One stored SpreadsheetML row record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Stored {
    pub(crate) index: Index,
    pub(crate) properties: Properties,
}

impl Stored {
    pub(crate) const fn new(index: Index, properties: Properties) -> Self {
        Self { index, properties }
    }
}

/// Borrowed view of one logical worksheet row.
///
/// Every checked grid row has a view. [`Self::stored`] distinguishes an
/// explicit SpreadsheetML row record from an implicit default row.
#[derive(Debug, Clone, Copy)]
pub struct Row<'a> {
    index: Index,
    stored: Option<&'a Stored>,
}

impl<'a> Row<'a> {
    pub(crate) const fn new(index: Index, stored: Option<&'a Stored>) -> Self {
        Self { index, stored }
    }

    /// Checked zero-based row coordinate.
    pub const fn index(self) -> Index {
        self.index
    }

    /// Whether the worksheet contains an explicit row record here.
    pub const fn stored(self) -> bool {
        self.stored.is_some()
    }

    /// Producer-stored row height, if present.
    pub const fn height(self) -> Option<Height> {
        match self.stored {
            Some(row) => row.properties.height,
            None => None,
        }
    }

    /// Producer-stored typographic descent at 100% worksheet zoom.
    pub const fn descent(self) -> Option<Descent> {
        match self.stored {
            Some(row) => row.properties.descent,
            None => None,
        }
    }

    /// Whether the row is explicitly hidden.
    pub const fn hidden(self) -> bool {
        match self.stored {
            Some(row) => row.properties.flags.contains(Flags::HIDDEN),
            None => false,
        }
    }

    /// Whether the producer stored a custom-height marker.
    pub const fn custom_height(self) -> bool {
        match self.stored {
            Some(row) => {
                row.properties.flags.contains(Flags::CUSTOM_HEIGHT)
                    || row.properties.descent.is_some()
            },
            None => false,
        }
    }

    /// Effective checked row outline level.
    pub const fn outline(self) -> Outline {
        match self.stored {
            Some(row) => row.properties.outline,
            None => Outline::NONE,
        }
    }

    /// Whether the row's outline is stored in its collapsed state.
    pub const fn collapsed(self) -> bool {
        match self.stored {
            Some(row) => row.properties.flags.contains(Flags::COLLAPSED),
            None => false,
        }
    }

    /// Whether the row requests a thick top edge.
    pub const fn thick_top(self) -> bool {
        match self.stored {
            Some(row) => row.properties.flags.contains(Flags::THICK_TOP),
            None => false,
        }
    }

    /// Whether the row requests a thick bottom edge.
    pub const fn thick_bottom(self) -> bool {
        match self.stored {
            Some(row) => row.properties.flags.contains(Flags::THICK_BOTTOM),
            None => false,
        }
    }

    /// Whether phonetic information is shown by default in this row.
    pub const fn phonetic(self) -> bool {
        match self.stored {
            Some(row) => row.properties.flags.contains(Flags::PHONETIC),
            None => false,
        }
    }

    /// Whether the producer explicitly applies the row style.
    pub const fn custom_format(self) -> bool {
        match self.stored {
            Some(row) => row.properties.flags.contains(Flags::CUSTOM_FORMAT),
            None => false,
        }
    }
}

/// Lazy borrowed traversal of explicit worksheet row records.
#[derive(Debug, Clone)]
pub struct Rows<'a> {
    inner: slice::Iter<'a, Stored>,
}

impl<'a> Rows<'a> {
    pub(crate) fn new(rows: &'a [Stored]) -> Self {
        Self { inner: rows.iter() }
    }
}

impl<'a> Iterator for Rows<'a> {
    type Item = Row<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        self.inner
            .next()
            .map(|stored| Row::new(stored.index, Some(stored)))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.inner.size_hint()
    }
}

impl DoubleEndedIterator for Rows<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.inner
            .next_back()
            .map(|stored| Row::new(stored.index, Some(stored)))
    }
}

impl ExactSizeIterator for Rows<'_> {}
impl std::iter::FusedIterator for Rows<'_> {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn heights_are_const_checked_before_entering_row_state() {
        const VALID: Result<Height, HeightError> = Height::new(30.0);
        assert_eq!(VALID.map(Height::get), Ok(30.0));
        assert_eq!(HeightAt::from(24).resolve().map(Height::get), Ok(24.0));
        assert_eq!(Height::new(-0.0).map(Height::get), Ok(0.0));
        assert_eq!(Height::new(409.0).map(Height::get), Ok(409.0));
        assert!(Height::new(-0.1).is_err());
        assert!(Height::new(f64::NAN).is_err());
        assert!(Height::new(f64::INFINITY).is_err());
        assert!(Height::new(409.1).is_err());
    }
}
