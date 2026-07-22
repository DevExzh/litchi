//! Native merged-cell range semantics for table-axis insertion.

use super::*;

/// The physical table axis receiving one inserted cell band.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum MergeAxis {
    Row,
    Column,
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
}
