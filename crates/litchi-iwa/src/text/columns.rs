//! Native/protobuf adaptation for archive-free text-column layouts.

use litchi_iwa_text::columns::{
    Columns, Count, Following, Gap, MAX_VARIABLE_COLUMNS, Variable, Width,
};

use crate::protobuf::tswp;
use crate::{Error, Result};

pub(crate) fn from_native(native: &tswp::ColumnsArchive) -> Result<Columns> {
    match (&native.equal_columns, &native.non_equal_columns) {
        (Some(_), Some(_)) => Err(Error::InvalidFormat(
            "iWork text columns are both equal-width and variable-width".into(),
        )),
        (Some(equal), None) => {
            let count = equal.count.ok_or_else(|| {
                Error::InvalidFormat("iWork equal text columns have no count".into())
            })?;
            Ok(Columns::equal(
                Count::new(count).map_err(|error| match error {
                    litchi_iwa_text::columns::Error::ZeroCount => {
                        Error::InvalidFormat("iWork equal text columns have a zero count".into())
                    },
                    litchi_iwa_text::columns::Error::TooManyColumns => Error::InvalidFormat(
                        "iWork equal text columns exceed the supported column limit".into(),
                    ),
                    _ => Error::InvalidFormat(
                        "iWork equal text columns contain an invalid count".into(),
                    ),
                })?,
                equal.gap.map(Gap::from_points).transpose().map_err(|_| {
                    Error::InvalidFormat("iWork equal text columns have an invalid gap".into())
                })?,
            ))
        },
        (None, Some(variable)) => {
            if variable.following.len().saturating_add(1) > MAX_VARIABLE_COLUMNS {
                return Err(Error::InvalidFormat(
                    "iWork variable text columns exceed the supported column limit".into(),
                ));
            }
            let first_width = native_width(variable.first)?;
            let following = variable
                .following
                .iter()
                .map(|column| {
                    Ok(Following::new(
                        native_gap(column.gap)?,
                        native_width(column.width)?,
                    ))
                })
                .collect::<Result<Vec<_>>>()?;
            Variable::new(first_width, following)
                .map(Columns::Variable)
                .map_err(|error| match error {
                    litchi_iwa_text::columns::Error::MissingFollowing => Error::InvalidFormat(
                        "variable-width iWork text columns contain only one column".into(),
                    ),
                    litchi_iwa_text::columns::Error::TooManyColumns => Error::InvalidFormat(
                        "iWork equal text columns exceed the supported column limit".into(),
                    ),
                    litchi_iwa_text::columns::Error::TooManyVariableColumns => {
                        Error::InvalidFormat(
                            "iWork variable text columns exceed the supported column limit".into(),
                        )
                    },
                    _ => Error::InvalidFormat(
                        "iWork variable text columns contain invalid layout values".into(),
                    ),
                })
        },
        (None, None) => Err(Error::InvalidFormat(
            "iWork text columns have no layout".into(),
        )),
    }
}

pub(crate) fn to_native(columns: &Columns) -> tswp::ColumnsArchive {
    match columns {
        Columns::Equal(equal) => tswp::ColumnsArchive {
            equal_columns: Some(tswp::columns_archive::EqualColumnsArchive {
                count: Some(equal.count().get()),
                gap: equal.gap().map(Gap::points),
            }),
            non_equal_columns: None,
        },
        Columns::Variable(variable) => tswp::ColumnsArchive {
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

fn native_gap(value: f32) -> Result<Gap> {
    Gap::from_points(value)
        .map_err(|_| Error::InvalidFormat("iWork text columns have an invalid gap".into()))
}

fn native_width(value: f32) -> Result<Width> {
    Width::from_points(value)
        .map_err(|_| Error::InvalidFormat("iWork text columns have an invalid width".into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_text::columns::MAX_COLUMNS;

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
        assert!(
            from_native(&tswp::ColumnsArchive {
                equal_columns: Some(tswp::columns_archive::EqualColumnsArchive {
                    count: Some(MAX_COLUMNS + 1),
                    gap: None,
                }),
                non_equal_columns: None,
            })
            .is_err()
        );
    }

    #[test]
    fn native_layouts_round_trip_through_the_adapter() {
        let equal = Columns::equal(
            Count::new(3).unwrap(),
            Some(Gap::from_points(12.0).unwrap()),
        );
        assert_eq!(from_native(&to_native(&equal)).unwrap(), equal);

        let variable = Columns::variable(
            Width::from_points(72.0).unwrap(),
            vec![Following::new(
                Gap::from_points(8.0).unwrap(),
                Width::from_points(96.0).unwrap(),
            )],
        )
        .unwrap();
        assert_eq!(from_native(&to_native(&variable)).unwrap(), variable);
    }
}
