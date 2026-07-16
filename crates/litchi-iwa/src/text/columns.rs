//! Strictly typed text-frame column layouts shared by Pages, Numbers, and Keynote.

use std::num::NonZeroU32;

use crate::protobuf::tswp;
use crate::{Error, Result};

/// Validated, non-zero number of equal-width text columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextColumnCount(NonZeroU32);

impl TextColumnCount {
    pub const ONE: Self = Self(NonZeroU32::MIN);

    pub fn new(count: u32) -> Result<Self> {
        NonZeroU32::new(count).map(Self).ok_or_else(|| {
            Error::ParseError("iWork text columns must have a non-zero count".into())
        })
    }

    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

impl TryFrom<u32> for TextColumnCount {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<TextColumnCount> for u32 {
    fn from(value: TextColumnCount) -> Self {
        value.get()
    }
}

/// Finite, non-negative distance between adjacent columns, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextColumnGap(f32);

impl TextColumnGap {
    pub const ZERO: Self = Self(0.0);

    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::ParseError(
                "iWork text column gap must be finite and non-negative".into(),
            ));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Finite, positive width of one variable-width column, in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct TextColumnWidth(f32);

impl TextColumnWidth {
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(Error::ParseError(
                "iWork text column width must be finite and positive".into(),
            ));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Equal-width column layout with an optional explicit inter-column gap.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct EqualTextColumns {
    count: TextColumnCount,
    gap: Option<TextColumnGap>,
}

impl EqualTextColumns {
    pub const fn new(count: TextColumnCount, gap: Option<TextColumnGap>) -> Self {
        Self { count, gap }
    }

    pub const fn count(self) -> TextColumnCount {
        self.count
    }

    /// Explicit gap, or `None` for the native automatic gap.
    pub const fn gap(self) -> Option<TextColumnGap> {
        self.gap
    }
}

/// One variable-width column after the first, including its preceding gap.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FollowingTextColumn {
    gap: TextColumnGap,
    width: TextColumnWidth,
}

impl FollowingTextColumn {
    pub const fn new(gap: TextColumnGap, width: TextColumnWidth) -> Self {
        Self { gap, width }
    }

    pub const fn gap(self) -> TextColumnGap {
        self.gap
    }

    pub const fn width(self) -> TextColumnWidth {
        self.width
    }
}

/// Variable-width column layout in native first-plus-following order.
#[derive(Debug, Clone, PartialEq)]
pub struct VariableTextColumns {
    first_width: TextColumnWidth,
    following: Box<[FollowingTextColumn]>,
}

impl VariableTextColumns {
    pub fn new(first_width: TextColumnWidth, following: Vec<FollowingTextColumn>) -> Result<Self> {
        if following.is_empty() {
            return Err(Error::ParseError(
                "variable-width iWork text columns require at least two columns".into(),
            ));
        }
        Ok(Self {
            first_width,
            following: following.into_boxed_slice(),
        })
    }

    pub const fn first_width(&self) -> TextColumnWidth {
        self.first_width
    }

    pub fn following(&self) -> &[FollowingTextColumn] {
        &self.following
    }

    pub fn count(&self) -> usize {
        self.following.len() + 1
    }
}

/// Native equal-width or variable-width text-frame columns.
#[derive(Debug, Clone, PartialEq)]
pub enum TextColumns {
    Equal(EqualTextColumns),
    Variable(VariableTextColumns),
}

impl TextColumns {
    pub const fn equal(count: TextColumnCount, gap: Option<TextColumnGap>) -> Self {
        Self::Equal(EqualTextColumns::new(count, gap))
    }

    pub fn variable(
        first_width: TextColumnWidth,
        following: Vec<FollowingTextColumn>,
    ) -> Result<Self> {
        Ok(Self::Variable(VariableTextColumns::new(
            first_width,
            following,
        )?))
    }

    pub fn count(&self) -> usize {
        match self {
            Self::Equal(columns) => columns.count().get() as usize,
            Self::Variable(columns) => columns.count(),
        }
    }
}

impl Default for TextColumns {
    fn default() -> Self {
        Self::equal(TextColumnCount::ONE, None)
    }
}

pub(crate) fn from_native(native: &tswp::ColumnsArchive) -> Result<TextColumns> {
    match (&native.equal_columns, &native.non_equal_columns) {
        (Some(_), Some(_)) => Err(Error::InvalidFormat(
            "iWork text columns are both equal-width and variable-width".into(),
        )),
        (Some(equal), None) => {
            let count = equal.count.ok_or_else(|| {
                Error::InvalidFormat("iWork equal text columns have no count".into())
            })?;
            Ok(TextColumns::equal(
                TextColumnCount::new(count).map_err(|_| {
                    Error::InvalidFormat("iWork equal text columns have a zero count".into())
                })?,
                equal
                    .gap
                    .map(TextColumnGap::from_points)
                    .transpose()
                    .map_err(|_| {
                        Error::InvalidFormat("iWork equal text columns have an invalid gap".into())
                    })?,
            ))
        },
        (None, Some(variable)) => {
            let first_width = native_width(variable.first)?;
            let following = variable
                .following
                .iter()
                .map(|column| {
                    Ok(FollowingTextColumn::new(
                        native_gap(column.gap)?,
                        native_width(column.width)?,
                    ))
                })
                .collect::<Result<Vec<_>>>()?;
            VariableTextColumns::new(first_width, following)
                .map(TextColumns::Variable)
                .map_err(|_| {
                    Error::InvalidFormat(
                        "variable-width iWork text columns contain only one column".into(),
                    )
                })
        },
        (None, None) => Err(Error::InvalidFormat(
            "iWork text columns have no layout".into(),
        )),
    }
}

pub(crate) fn to_native(columns: &TextColumns) -> tswp::ColumnsArchive {
    match columns {
        TextColumns::Equal(equal) => tswp::ColumnsArchive {
            equal_columns: Some(tswp::columns_archive::EqualColumnsArchive {
                count: Some(equal.count().get()),
                gap: equal.gap().map(TextColumnGap::points),
            }),
            non_equal_columns: None,
        },
        TextColumns::Variable(variable) => tswp::ColumnsArchive {
            equal_columns: None,
            non_equal_columns: Some(tswp::columns_archive::NonEqualColumnsArchive {
                first: variable.first_width().points(),
                following: variable
                    .following()
                    .iter()
                    .map(|column| {
                        tswp::columns_archive::non_equal_columns_archive::GapWidthArchive {
                            gap: column.gap().points(),
                            width: column.width().points(),
                        }
                    })
                    .collect(),
            }),
        },
    }
}

fn native_gap(value: f32) -> Result<TextColumnGap> {
    TextColumnGap::from_points(value)
        .map_err(|_| Error::InvalidFormat("iWork text columns have an invalid gap".into()))
}

fn native_width(value: f32) -> Result<TextColumnWidth> {
    TextColumnWidth::from_points(value)
        .map_err(|_| Error::InvalidFormat("iWork text columns have an invalid width".into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalars_reject_invalid_values() {
        assert!(TextColumnCount::new(0).is_err());
        assert!(TextColumnGap::from_points(-0.01).is_err());
        assert!(TextColumnGap::from_points(f32::NAN).is_err());
        assert!(TextColumnWidth::from_points(0.0).is_err());
        assert!(TextColumnWidth::from_points(f32::INFINITY).is_err());
    }

    #[test]
    fn equal_and_variable_columns_round_trip() {
        let equal = TextColumns::equal(
            TextColumnCount::new(3).unwrap(),
            Some(TextColumnGap::from_points(12.0).unwrap()),
        );
        assert_eq!(from_native(&to_native(&equal)).unwrap(), equal);

        let variable = TextColumns::variable(
            TextColumnWidth::from_points(72.0).unwrap(),
            vec![FollowingTextColumn::new(
                TextColumnGap::from_points(8.0).unwrap(),
                TextColumnWidth::from_points(96.0).unwrap(),
            )],
        )
        .unwrap();
        assert_eq!(from_native(&to_native(&variable)).unwrap(), variable);
    }

    #[test]
    fn malformed_native_columns_are_rejected() {
        assert!(from_native(&tswp::ColumnsArchive::default()).is_err());
        assert!(
            from_native(&tswp::ColumnsArchive {
                equal_columns: Some(tswp::columns_archive::EqualColumnsArchive {
                    count: Some(0),
                    gap: None,
                }),
                non_equal_columns: None,
            })
            .is_err()
        );
    }
}
