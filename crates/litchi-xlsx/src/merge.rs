//! Checked merged ranges and compact lookup support.

use std::cmp::Reverse;
use std::collections::{BTreeMap, BinaryHeap};
use std::iter::FusedIterator;

use litchi_sheet::{Cell as Address, Rect};

use crate::error::{Result, allocation, invalid};

/// One valid merged-range membership transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Change {
    /// The range was absent and is now merged.
    Add,
    /// The range was merged and is now absent.
    Remove,
}

impl Change {
    /// Membership before this transition.
    #[must_use]
    pub const fn before(self) -> bool {
        matches!(self, Self::Remove)
    }

    /// Membership after this transition.
    #[must_use]
    pub const fn after(self) -> bool {
        matches!(self, Self::Add)
    }

    pub(crate) const fn inverse(self) -> Self {
        match self {
            Self::Add => Self::Remove,
            Self::Remove => Self::Add,
        }
    }
}

/// Borrowed merged ranges in deterministic worksheet order.
#[derive(Debug, Clone)]
pub struct Merges<'a> {
    ranges: std::slice::Iter<'a, Rect>,
}

impl<'a> Merges<'a> {
    pub(crate) fn new(ranges: &'a [Rect]) -> Self {
        Self {
            ranges: ranges.iter(),
        }
    }
}

impl Iterator for Merges<'_> {
    type Item = Rect;

    fn next(&mut self) -> Option<Self::Item> {
        self.ranges.next().copied()
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.ranges.size_hint()
    }
}

impl DoubleEndedIterator for Merges<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.ranges.next_back().copied()
    }
}

impl ExactSizeIterator for Merges<'_> {}
impl FusedIterator for Merges<'_> {}

const NONE: u32 = u32::MAX;

#[derive(Debug, Clone, Copy)]
struct Node {
    center: u32,
    spans_start: u32,
    spans_len: u32,
    lower: u32,
    upper: u32,
}

/// Validated static interval tree over non-overlapping merged ranges.
#[derive(Debug)]
pub(crate) struct Index {
    ranges: Box<[Rect]>,
    nodes: Box<[Node]>,
    spans: Box<[u32]>,
    root: u32,
}

impl Default for Index {
    fn default() -> Self {
        Self {
            ranges: Box::new([]),
            nodes: Box::new([]),
            spans: Box::new([]),
            root: NONE,
        }
    }
}

impl Index {
    pub(crate) fn new(mut ranges: Vec<Rect>) -> Result<Self> {
        ranges.sort_unstable_by_key(key);
        validate_non_overlapping(&ranges)?;
        let count = u32::try_from(ranges.len())
            .ok()
            .filter(|count| *count != NONE)
            .ok_or_else(|| invalid("merged-range count exceeds the compact index domain"))?;
        let mut indices = Vec::new();
        indices
            .try_reserve_exact(ranges.len())
            .map_err(|source| allocation("merged-range indexes", source))?;
        indices.extend(0..count);
        let mut nodes = Vec::new();
        nodes
            .try_reserve_exact(ranges.len())
            .map_err(|source| allocation("merged-range tree", source))?;
        let mut spans = Vec::new();
        spans
            .try_reserve_exact(ranges.len())
            .map_err(|source| allocation("merged-range spans", source))?;
        let root = build(&ranges, indices, &mut nodes, &mut spans)?;
        Ok(Self {
            ranges: ranges.into_boxed_slice(),
            nodes: nodes.into_boxed_slice(),
            spans: spans.into_boxed_slice(),
            root,
        })
    }

    pub(crate) fn containing(&self, address: Address) -> Option<Rect> {
        let row = address.row().get();
        let column = address.column().get();
        let mut current = self.root;
        while current != NONE {
            let node = *self.nodes.get(usize::try_from(current).ok()?)?;
            let start = usize::try_from(node.spans_start).ok()?;
            let end = start.checked_add(usize::try_from(node.spans_len).ok()?)?;
            let spans = self.spans.get(start..end)?;
            let candidate = spans.partition_point(|index| {
                usize::try_from(*index)
                    .ok()
                    .and_then(|index| self.ranges.get(index))
                    .is_some_and(|range| range.start().column().get() <= column)
            });
            if let Some(index) = candidate.checked_sub(1).and_then(|index| spans.get(index)) {
                let range = *self.ranges.get(usize::try_from(*index).ok()?)?;
                if range.contains(address) {
                    return Some(range);
                }
            }
            current = match row.cmp(&node.center) {
                std::cmp::Ordering::Less => node.lower,
                std::cmp::Ordering::Greater => node.upper,
                std::cmp::Ordering::Equal => NONE,
            };
        }
        None
    }

    pub(crate) fn as_slice(&self) -> &[Rect] {
        &self.ranges
    }

    pub(crate) fn iter(&self) -> Merges<'_> {
        Merges::new(&self.ranges)
    }
}

fn build(
    ranges: &[Rect],
    indices: Vec<u32>,
    nodes: &mut Vec<Node>,
    spans: &mut Vec<u32>,
) -> Result<u32> {
    if indices.is_empty() {
        return Ok(NONE);
    }
    let median = *indices
        .get(indices.len() / 2)
        .ok_or_else(|| invalid("merged-range tree lost its median"))?;
    let center = ranges
        .get(usize::try_from(median).map_err(|_| invalid("merge index does not fit usize"))?)
        .ok_or_else(|| invalid("merged-range tree median escaped its source"))?
        .start()
        .row()
        .get();
    let (mut lower_count, mut spanning_count, mut upper_count) = (0usize, 0usize, 0usize);
    for index in &indices {
        let range = ranges
            .get(usize::try_from(*index).map_err(|_| invalid("merge index does not fit usize"))?)
            .ok_or_else(|| invalid("merged-range tree index escaped its source"))?;
        if range.end().0 <= center {
            lower_count = lower_count.saturating_add(1);
        } else if range.start().row().get() > center {
            upper_count = upper_count.saturating_add(1);
        } else {
            spanning_count = spanning_count.saturating_add(1);
        }
    }
    let mut lower = Vec::new();
    lower
        .try_reserve_exact(lower_count)
        .map_err(|source| allocation("lower merge tree", source))?;
    let mut spanning = Vec::new();
    spanning
        .try_reserve_exact(spanning_count)
        .map_err(|source| allocation("spanning merge tree", source))?;
    let mut upper = Vec::new();
    upper
        .try_reserve_exact(upper_count)
        .map_err(|source| allocation("upper merge tree", source))?;
    for index in indices {
        let range = ranges
            .get(usize::try_from(index).map_err(|_| invalid("merge index does not fit usize"))?)
            .ok_or_else(|| invalid("merged-range tree index escaped its source"))?;
        if range.end().0 <= center {
            lower.push(index);
        } else if range.start().row().get() > center {
            upper.push(index);
        } else {
            spanning.push(index);
        }
    }
    spanning.sort_unstable_by_key(|index| {
        usize::try_from(*index)
            .ok()
            .and_then(|index| ranges.get(index))
            .map_or((u32::MAX, u32::MAX), |range| {
                (range.start().column().get(), range.end().1)
            })
    });

    let position = u32::try_from(nodes.len())
        .ok()
        .filter(|position| *position != NONE)
        .ok_or_else(|| invalid("merged-range tree exceeds its compact node domain"))?;
    let spans_start = u32::try_from(spans.len())
        .map_err(|_| invalid("merged-range span offset does not fit u32"))?;
    let spans_len = u32::try_from(spanning.len())
        .map_err(|_| invalid("merged-range span length does not fit u32"))?;
    spans.extend(spanning);
    nodes.push(Node {
        center,
        spans_start,
        spans_len,
        lower: NONE,
        upper: NONE,
    });
    let lower = build(ranges, lower, nodes, spans)?;
    let upper = build(ranges, upper, nodes, spans)?;
    let node = nodes
        .get_mut(
            usize::try_from(position)
                .map_err(|_| invalid("merge node index does not fit usize"))?,
        )
        .ok_or_else(|| invalid("merged-range tree node disappeared during construction"))?;
    node.lower = lower;
    node.upper = upper;
    Ok(position)
}

pub(crate) fn overlaps(left: Rect, right: Rect) -> bool {
    let (left_end_row, left_end_column) = left.end();
    let (right_end_row, right_end_column) = right.end();
    left.start().row().get() < right_end_row
        && right.start().row().get() < left_end_row
        && left.start().column().get() < right_end_column
        && right.start().column().get() < left_end_column
}

fn key(range: &Rect) -> (u32, u32, u32, u32) {
    let (end_row, end_column) = range.end();
    (
        range.start().row().get(),
        range.start().column().get(),
        end_row,
        end_column,
    )
}

fn validate_non_overlapping(ranges: &[Rect]) -> Result<()> {
    let mut active = BTreeMap::<u32, (u32, usize)>::new();
    let mut expiry = BinaryHeap::<Reverse<(u32, u32, usize)>>::new();
    for (index, range) in ranges.iter().copied().enumerate() {
        let start_row = range.start().row().get();
        while expiry
            .peek()
            .is_some_and(|Reverse((end_row, _, _))| *end_row <= start_row)
        {
            let Some(Reverse((_, start_column, expired))) = expiry.pop() else {
                return Err(invalid("merged-range expiry index became inconsistent"));
            };
            if active
                .get(&start_column)
                .is_some_and(|(_, live)| *live == expired)
            {
                active.remove(&start_column);
            }
        }

        let start_column = range.start().column().get();
        let (end_row, end_column) = range.end();
        if let Some((_, (active_end, active_index))) = active.range(..end_column).next_back()
            && *active_end > start_column
        {
            let existing = ranges
                .get(*active_index)
                .copied()
                .ok_or_else(|| invalid("merged-range overlap index escaped its source"))?;
            return Err(invalid(format!(
                "overlapping merged ranges {existing} and {range}"
            )));
        }
        if active.insert(start_column, (end_column, index)).is_some() {
            return Err(invalid(format!(
                "overlapping merged ranges share start column {start_column}"
            )));
        }
        expiry.push(Reverse((end_row, start_column, index)));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_and_finds_sparse_vertical_and_horizontal_merges() {
        let index = Index::new(vec![
            Rect::from_a1("A1:A100").expect("range"),
            Rect::from_a1("C2:F2").expect("range"),
            Rect::from_a1("B101:D102").expect("range"),
        ])
        .expect("non-overlapping merges");
        assert_eq!(
            index.containing(Address::from_a1("A50").expect("address")),
            Some(Rect::from_a1("A1:A100").expect("range"))
        );
        assert_eq!(
            index.containing(Address::from_a1("E2").expect("address")),
            Some(Rect::from_a1("C2:F2").expect("range"))
        );
        assert!(
            index
                .containing(Address::from_a1("B50").expect("address"))
                .is_none()
        );
    }

    #[test]
    fn rejects_two_dimensional_overlap_without_expanding_ranges() {
        let result = Index::new(vec![
            Rect::from_a1("A1:C3").expect("range"),
            Rect::from_a1("C3:D4").expect("range"),
        ]);
        assert!(result.is_err());
    }

    #[test]
    fn interval_tree_finds_many_tall_ranges_without_expanding_rows() {
        let mut ranges = Vec::new();
        for column in (0..2_000).step_by(2) {
            ranges.push(
                Rect::at(0, column, litchi_sheet::ROWS, column + 1).expect("tall merged range"),
            );
        }
        let index = Index::new(ranges).expect("non-overlapping tall ranges");
        assert_eq!(index.nodes.len(), 1, "same-row spans share one tree node");
        assert_eq!(index.spans.len(), 1_000);
        assert!(
            index
                .containing(Address::at(litchi_sheet::ROWS - 1, 1_998).expect("covered"))
                .is_some()
        );
        assert!(
            index
                .containing(Address::at(litchi_sheet::ROWS - 1, 1_999).expect("gap"))
                .is_none()
        );
    }

    #[test]
    fn interval_tree_matches_a_sparse_grid_oracle() {
        let mut ranges = Vec::new();
        for row in 0..40 {
            for column in 0..20 {
                ranges.push(
                    Rect::at(row * 3, column * 3, row * 3 + 2, column * 3 + 2)
                        .expect("spaced merge"),
                );
            }
        }
        ranges.reverse();
        let index = Index::new(ranges.clone()).expect("non-overlapping sparse grid");
        for row in 0..120 {
            for column in 0..60 {
                let address = Address::at(row, column).expect("grid address");
                let expected = ranges.iter().copied().find(|range| range.contains(address));
                assert_eq!(index.containing(address), expected, "at {address}");
            }
        }
    }
}
