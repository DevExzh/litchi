//! Native merged-cell range semantics for table-axis mutations.

use super::*;

/// A physical table axis participating in a topology mutation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum MergeAxis {
    Row,
    Column,
}

/// The native outcome for one merged region after deleting a cell band.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum MergeDeletion {
    /// The region remains a native merge with these coordinates.
    Retain(IWorkTableCellRegion),
    /// The deletion removes the merge formula, either because every merged
    /// cell was deleted or because the remaining region is one cell.
    Remove,
}

/// A merge anchor that must move into a surviving cell before the physical
/// row or column is removed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct MergeAnchorRelocation {
    pub(crate) source_row: usize,
    pub(crate) source_column: usize,
    pub(crate) destination_row: usize,
    pub(crate) destination_column: usize,
}

/// Apply Numbers' native merge behavior for a one-cell axis insertion.
///
/// Inserting at a merge's leading boundary moves the entire region. Inserting
/// strictly inside the region expands it by one cell along that axis. An
/// insertion immediately after the region leaves it unchanged.
pub(super) fn region_after_insertion(
    region: IWorkTableCellRegion,
    axis: MergeAxis,
    insertion: usize,
) -> Result<IWorkTableCellRegion> {
    let (start, end, count) = match axis {
        MergeAxis::Row => (region.row(), region.end_row(), region.row_count()),
        MergeAxis::Column => (region.column(), region.end_column(), region.column_count()),
    };
    let (new_start, new_count) = if insertion <= start {
        (
            start.checked_add(1).ok_or_else(|| {
                Error::ParseError("iWork merged-cell coordinate overflow".to_owned())
            })?,
            count,
        )
    } else if insertion <= end {
        (
            start,
            count
                .checked_add(1)
                .ok_or_else(|| Error::ParseError("iWork merged-cell span overflow".to_owned()))?,
        )
    } else {
        (start, count)
    };

    match axis {
        MergeAxis::Row => {
            IWorkTableCellRegion::new(new_start, region.column(), new_count, region.column_count())
        },
        MergeAxis::Column => {
            IWorkTableCellRegion::new(region.row(), new_start, region.row_count(), new_count)
        },
    }
}

/// Apply Numbers' native merge behavior for a one-cell axis deletion.
///
/// A deleted band before a merge shifts it toward the origin. A deleted band
/// through a merge contracts it along that axis. Once only one table cell
/// remains, native iWork removes the merge formula altogether.
pub(super) fn region_after_deletion(
    region: IWorkTableCellRegion,
    axis: MergeAxis,
    deletion: usize,
) -> Result<MergeDeletion> {
    let (start, end, count, other_count) = match axis {
        MergeAxis::Row => (
            region.row(),
            region.end_row(),
            region.row_count(),
            region.column_count(),
        ),
        MergeAxis::Column => (
            region.column(),
            region.end_column(),
            region.column_count(),
            region.row_count(),
        ),
    };

    if deletion < start {
        let shifted_start = start.checked_sub(1).ok_or_else(|| {
            Error::ParseError("iWork merged-cell coordinate underflow".to_owned())
        })?;
        return match axis {
            MergeAxis::Row => {
                IWorkTableCellRegion::new(shifted_start, region.column(), count, other_count)
                    .map(MergeDeletion::Retain)
            },
            MergeAxis::Column => {
                IWorkTableCellRegion::new(region.row(), shifted_start, other_count, count)
                    .map(MergeDeletion::Retain)
            },
        };
    }
    if deletion > end {
        return Ok(MergeDeletion::Retain(region));
    }

    if count == 1 {
        return Ok(MergeDeletion::Remove);
    }
    let contracted_count = count - 1;
    if contracted_count == 1 && other_count == 1 {
        return Ok(MergeDeletion::Remove);
    }
    match axis {
        MergeAxis::Row => {
            IWorkTableCellRegion::new(region.row(), region.column(), contracted_count, other_count)
                .map(MergeDeletion::Retain)
        },
        MergeAxis::Column => {
            IWorkTableCellRegion::new(region.row(), region.column(), other_count, contracted_count)
                .map(MergeDeletion::Retain)
        },
    }
}

/// Return the anchor-cell relocation required before deleting a leading merge
/// boundary. A region that loses its only row or column has no survivor.
pub(super) fn anchor_relocation_after_deletion(
    region: IWorkTableCellRegion,
    axis: MergeAxis,
    deletion: usize,
) -> Result<Option<MergeAnchorRelocation>> {
    let (start, count) = match axis {
        MergeAxis::Row => (region.row(), region.row_count()),
        MergeAxis::Column => (region.column(), region.column_count()),
    };
    if deletion != start || count == 1 {
        return Ok(None);
    }
    let next = start
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork merged-cell coordinate overflow".to_owned()))?;
    Ok(Some(match axis {
        MergeAxis::Row => MergeAnchorRelocation {
            source_row: region.row(),
            source_column: region.column(),
            destination_row: next,
            destination_column: region.column(),
        },
        MergeAxis::Column => MergeAnchorRelocation {
            source_row: region.row(),
            source_column: region.column(),
            destination_row: region.row(),
            destination_column: next,
        },
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn insertion_moves_or_expands_regions_at_native_boundaries() {
        let region = IWorkTableCellRegion::new(3, 4, 2, 3).unwrap();

        assert_eq!(
            region_after_insertion(region, MergeAxis::Row, 3).unwrap(),
            IWorkTableCellRegion::new(4, 4, 2, 3).unwrap()
        );
        assert_eq!(
            region_after_insertion(region, MergeAxis::Row, 4).unwrap(),
            IWorkTableCellRegion::new(3, 4, 3, 3).unwrap()
        );
        assert_eq!(
            region_after_insertion(region, MergeAxis::Row, 5).unwrap(),
            region
        );
        assert_eq!(
            region_after_insertion(region, MergeAxis::Column, 5).unwrap(),
            IWorkTableCellRegion::new(3, 4, 2, 4).unwrap()
        );
    }

    #[test]
    fn deletion_shifts_contracts_and_removes_regions_at_native_boundaries() {
        let region = IWorkTableCellRegion::new(3, 4, 2, 3).unwrap();

        assert_eq!(
            region_after_deletion(region, MergeAxis::Row, 2).unwrap(),
            MergeDeletion::Retain(IWorkTableCellRegion::new(2, 4, 2, 3).unwrap())
        );
        assert_eq!(
            region_after_deletion(region, MergeAxis::Row, 3).unwrap(),
            MergeDeletion::Retain(IWorkTableCellRegion::new(3, 4, 1, 3).unwrap())
        );
        assert_eq!(
            region_after_deletion(region, MergeAxis::Column, 5).unwrap(),
            MergeDeletion::Retain(IWorkTableCellRegion::new(3, 4, 2, 2).unwrap())
        );
        assert_eq!(
            region_after_deletion(region, MergeAxis::Column, 7).unwrap(),
            MergeDeletion::Retain(region)
        );
        assert_eq!(
            region_after_deletion(
                IWorkTableCellRegion::new(1, 1, 1, 2).unwrap(),
                MergeAxis::Column,
                1,
            )
            .unwrap(),
            MergeDeletion::Remove
        );
        assert_eq!(
            region_after_deletion(
                IWorkTableCellRegion::new(1, 1, 1, 2).unwrap(),
                MergeAxis::Row,
                1,
            )
            .unwrap(),
            MergeDeletion::Remove
        );
    }

    #[test]
    fn deletion_relocates_only_surviving_leading_anchors() {
        let region = IWorkTableCellRegion::new(3, 4, 2, 3).unwrap();
        assert_eq!(
            anchor_relocation_after_deletion(region, MergeAxis::Row, 3).unwrap(),
            Some(MergeAnchorRelocation {
                source_row: 3,
                source_column: 4,
                destination_row: 4,
                destination_column: 4,
            })
        );
        assert_eq!(
            anchor_relocation_after_deletion(region, MergeAxis::Column, 4).unwrap(),
            Some(MergeAnchorRelocation {
                source_row: 3,
                source_column: 4,
                destination_row: 3,
                destination_column: 5,
            })
        );
        assert_eq!(
            anchor_relocation_after_deletion(region, MergeAxis::Row, 4).unwrap(),
            None
        );
        assert_eq!(
            anchor_relocation_after_deletion(
                IWorkTableCellRegion::new(3, 4, 1, 3).unwrap(),
                MergeAxis::Row,
                3,
            )
            .unwrap(),
            None
        );
    }
}
