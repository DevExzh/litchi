//! Borrowed worksheet column views and bounded interval resolution.

use bitflags::bitflags;
use litchi_sheet::{COLUMNS, Column as Index};
use thiserror::Error;

use crate::error::{Result, allocation, invalid};
pub use crate::outline::{Outline, OutlineAt, OutlineError};
use crate::style::StyleState;

/// Checked SpreadsheetML column width in character units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Width(u64);

impl Width {
    /// Validate the Office column-width range `0..=255`.
    pub const fn new(value: f64) -> std::result::Result<Self, WidthError> {
        if !(value >= 0.0 && value <= 255.0) {
            return Err(WidthError { value });
        }
        let normalized = if value == 0.0 { 0.0 } else { value };
        Ok(Self(normalized.to_bits()))
    }

    /// Return the width in SpreadsheetML character units.
    pub const fn get(self) -> f64 {
        f64::from_bits(self.0)
    }
}

/// Invalid SpreadsheetML column width.
#[derive(Debug, Clone, Copy, PartialEq, Error)]
#[error("column width {value} is outside the finite Office range 0..=255")]
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
    /// A width that has already been checked.
    Checked(Width),
    /// A raw width validated when resolved.
    Value(f64),
}

impl WidthAt {
    /// Resolve this input into a checked width.
    pub const fn resolve(self) -> std::result::Result<Width, WidthError> {
        match self {
            Self::Checked(width) => Ok(width),
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

macro_rules! width_inputs {
    ($($input:ty),+ $(,)?) => {
        $(
            impl From<$input> for WidthAt {
                fn from(value: $input) -> Self {
                    Self::Value(f64::from(value))
                }
            }
        )+
    };
}

width_inputs!(f32, u8, u16, u32, i8, i16, i32);

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub(crate) struct Flags: u8 {
        const HIDDEN = 1 << 0;
        const BEST_FIT = 1 << 1;
        const CUSTOM_WIDTH = 1 << 2;
        const PHONETIC = 1 << 3;
        const COLLAPSED = 1 << 4;
    }
}

/// Complete effective state of one SpreadsheetML column-property record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Properties {
    pub(crate) width: Option<Width>,
    pub(crate) style: Option<u32>,
    pub(crate) outline: Outline,
    pub(crate) flags: Flags,
}

/// Complete modeled state of one effective column-property record.
///
/// Physical shared-style indexes remain hidden behind [`StyleState`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Props {
    pub(crate) width: Option<Width>,
    pub(crate) style: StyleState,
    pub(crate) outline: Outline,
    pub(crate) flags: Flags,
}

impl Props {
    /// Producer-stored width, if the record has one.
    pub const fn width(&self) -> Option<Width> {
        self.width
    }

    /// Exact local shared-style state without a physical style index.
    pub const fn style(&self) -> &StyleState {
        &self.style
    }

    /// Checked outline level.
    pub const fn outline(&self) -> Outline {
        self.outline
    }

    /// Whether the column is hidden.
    pub const fn hidden(&self) -> bool {
        self.flags.contains(Flags::HIDDEN)
    }

    /// Whether the producer marked the width as best-fit.
    pub const fn best_fit(&self) -> bool {
        self.flags.contains(Flags::BEST_FIT)
    }

    /// Whether the width is explicitly customized.
    pub const fn custom_width(&self) -> bool {
        self.flags.contains(Flags::CUSTOM_WIDTH)
    }

    /// Whether phonetic information is shown by default.
    pub const fn phonetic(&self) -> bool {
        self.flags.contains(Flags::PHONETIC)
    }

    /// Whether the outline is stored in its collapsed state.
    pub const fn collapsed(&self) -> bool {
        self.flags.contains(Flags::COLLAPSED)
    }

    pub(crate) fn rebind_style(&mut self, lineage: &std::sync::Arc<crate::style::StyleLineage>) {
        self.style.rebind(lineage);
    }

    pub(crate) const fn uses_shared_style(&self) -> bool {
        matches!(&self.style, StyleState::Shared(_))
    }
}

/// Column-property record state captured before or after a patch change.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    /// No explicit column-property record applies.
    Missing,
    /// One complete effective record applies.
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

/// One disjoint effective SpreadsheetML column-property range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Stored {
    pub(crate) first: Index,
    pub(crate) last: Index,
    pub(crate) properties: Properties,
}

impl Stored {
    const fn contains(self, index: Index) -> bool {
        self.first.get() <= index.get() && index.get() <= self.last.get()
    }

    const fn len(self) -> usize {
        (self.last.get() - self.first.get() + 1) as usize
    }
}

/// Borrowed view of one logical worksheet column.
///
/// Every checked grid column has a view. [`Self::stored`] distinguishes an
/// explicit effective `<col>` record from an implicit default column.
#[derive(Debug, Clone, Copy)]
pub struct Column<'a> {
    index: Index,
    stored: Option<&'a Stored>,
}

impl<'a> Column<'a> {
    pub(crate) const fn new(index: Index, stored: Option<&'a Stored>) -> Self {
        Self { index, stored }
    }

    /// Checked zero-based column coordinate.
    pub const fn index(self) -> Index {
        self.index
    }

    /// Whether an explicit column-property record applies here.
    pub const fn stored(self) -> bool {
        self.stored.is_some()
    }

    /// Whether the effective column-property record hides this column.
    pub const fn hidden(self) -> bool {
        match self.stored {
            Some(column) => column.properties.flags.contains(Flags::HIDDEN),
            None => false,
        }
    }

    /// Producer-stored width in character units, if present.
    pub const fn width(self) -> Option<Width> {
        match self.stored {
            Some(column) => column.properties.width,
            None => None,
        }
    }

    /// Whether the producer marked this column as best-fit.
    pub const fn best_fit(self) -> bool {
        match self.stored {
            Some(column) => column.properties.flags.contains(Flags::BEST_FIT),
            None => false,
        }
    }

    /// Whether the producer stored a custom width flag.
    pub const fn custom_width(self) -> bool {
        match self.stored {
            Some(column) => column.properties.flags.contains(Flags::CUSTOM_WIDTH),
            None => false,
        }
    }

    /// Whether phonetic information is shown by default in this column.
    pub const fn phonetic(self) -> bool {
        match self.stored {
            Some(column) => column.properties.flags.contains(Flags::PHONETIC),
            None => false,
        }
    }

    /// Effective checked column outline level.
    pub const fn outline(self) -> Outline {
        match self.stored {
            Some(column) => column.properties.outline,
            None => Outline::NONE,
        }
    }

    /// Whether the affected outline is stored in the collapsed state.
    pub const fn collapsed(self) -> bool {
        match self.stored {
            Some(column) => column.properties.flags.contains(Flags::COLLAPSED),
            None => false,
        }
    }
}

/// Lazy borrowed traversal of logical columns with explicit property records.
#[derive(Debug, Clone)]
pub struct Columns<'a> {
    ranges: &'a [Stored],
    front_range: usize,
    front: Option<Index>,
    back_range: usize,
    back: Option<Index>,
    remaining: usize,
}

impl<'a> Columns<'a> {
    pub(crate) fn new(ranges: &'a [Stored]) -> Self {
        let remaining = ranges.iter().map(|range| range.len()).sum();
        Self {
            ranges,
            front_range: 0,
            front: ranges.first().map(|range| range.first),
            back_range: ranges.len().saturating_sub(1),
            back: ranges.last().map(|range| range.last),
            remaining,
        }
    }

    fn advance_front(&mut self, stored: &Stored) {
        let Some(front) = self.front else {
            self.remaining = 0;
            return;
        };
        if front == stored.last {
            self.front_range = self.front_range.saturating_add(1);
            self.front = self.ranges.get(self.front_range).map(|range| range.first);
        } else {
            self.front = front.next();
        }
    }

    fn advance_back(&mut self, stored: &Stored) {
        let Some(back) = self.back else {
            self.remaining = 0;
            return;
        };
        if back == stored.first {
            self.back_range = self.back_range.saturating_sub(1);
            self.back = self.ranges.get(self.back_range).map(|range| range.last);
        } else {
            self.back = back.previous();
        }
    }
}

impl<'a> Iterator for Columns<'a> {
    type Item = Column<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.remaining == 0 {
            return None;
        }
        let stored = self.ranges.get(self.front_range)?;
        let index = self.front?;
        self.remaining -= 1;
        self.advance_front(stored);
        Some(Column::new(index, Some(stored)))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (self.remaining, Some(self.remaining))
    }
}

impl DoubleEndedIterator for Columns<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        if self.remaining == 0 {
            return None;
        }
        let stored = self.ranges.get(self.back_range)?;
        let index = self.back?;
        self.remaining -= 1;
        self.advance_back(stored);
        Some(Column::new(index, Some(stored)))
    }
}

impl ExactSizeIterator for Columns<'_> {}
impl std::iter::FusedIterator for Columns<'_> {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Node<T> {
    Unset,
    Value(T),
    Split,
}

/// Fixed-grid lazy range assignment map.
///
/// Assignments are `O(log COLUMNS)` even for a full-width record. This keeps
/// maliciously overlapping `<col>` records from multiplying parser work by
/// the width of every range.
#[derive(Debug)]
pub(crate) struct Assignments<T> {
    nodes: Vec<Node<T>>,
}

impl<T> Assignments<T>
where
    T: Copy + Eq,
{
    pub(crate) fn new() -> Result<Self> {
        let capacity = (COLUMNS as usize)
            .checked_mul(2)
            .ok_or_else(|| invalid("column interval-map capacity overflow"))?;
        let mut nodes = Vec::new();
        nodes
            .try_reserve_exact(capacity)
            .map_err(|source| allocation("column interval map", source))?;
        nodes.resize(capacity, Node::Unset);
        Ok(Self { nodes })
    }

    pub(crate) fn assign(&mut self, first: Index, last: Index, value: T) {
        self.assign_node(
            1,
            0,
            COLUMNS,
            first.get(),
            last.get().saturating_add(1),
            value,
        );
    }

    pub(crate) fn get(&self, index: Index) -> Option<T> {
        let mut node = 1usize;
        let mut first = 0u32;
        let mut last = COLUMNS;
        loop {
            match self.nodes.get(node).copied()? {
                Node::Unset => return None,
                Node::Value(value) => return Some(value),
                Node::Split => {
                    let middle = first + (last - first) / 2;
                    if index.get() < middle {
                        node = node.saturating_mul(2);
                        last = middle;
                    } else {
                        node = node.saturating_mul(2).saturating_add(1);
                        first = middle;
                    }
                },
            }
        }
    }

    pub(crate) fn into_ranges(self) -> Result<Box<[Assigned<T>]>> {
        let mut raw = Vec::new();
        self.collect(1, 0, COLUMNS, &mut raw)?;
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(raw.len())
            .map_err(|source| allocation("column ranges", source))?;
        for (first, last, value) in raw {
            ranges.push(Assigned {
                first: Index::new(first)?,
                last: Index::new(last - 1)?,
                value,
            });
        }
        Ok(ranges.into_boxed_slice())
    }

    fn assign_node(
        &mut self,
        node: usize,
        first: u32,
        last: u32,
        assigned_first: u32,
        assigned_last: u32,
        value: T,
    ) {
        if assigned_first <= first && last <= assigned_last {
            if let Some(slot) = self.nodes.get_mut(node) {
                *slot = Node::Value(value);
            }
            return;
        }
        let middle = first + (last - first) / 2;
        let left = node.saturating_mul(2);
        let right = left.saturating_add(1);
        let current = self.nodes.get(node).copied().unwrap_or(Node::Unset);
        if current != Node::Split {
            if let Some(slot) = self.nodes.get_mut(left) {
                *slot = current;
            }
            if let Some(slot) = self.nodes.get_mut(right) {
                *slot = current;
            }
        }
        if assigned_first < middle {
            self.assign_node(left, first, middle, assigned_first, assigned_last, value);
        }
        if middle < assigned_last {
            self.assign_node(right, middle, last, assigned_first, assigned_last, value);
        }
        let left_value = self.nodes.get(left).copied().unwrap_or(Node::Unset);
        let right_value = self.nodes.get(right).copied().unwrap_or(Node::Unset);
        if let Some(slot) = self.nodes.get_mut(node) {
            *slot = if left_value == right_value {
                left_value
            } else {
                Node::Split
            };
        }
    }

    fn collect(
        &self,
        node: usize,
        first: u32,
        last: u32,
        output: &mut Vec<(u32, u32, T)>,
    ) -> Result<()> {
        match self.nodes.get(node).copied().unwrap_or(Node::Unset) {
            Node::Unset => Ok(()),
            Node::Value(value) => push_range(output, first, last, value),
            Node::Split => {
                let middle = first + (last - first) / 2;
                self.collect(node.saturating_mul(2), first, middle, output)?;
                self.collect(
                    node.saturating_mul(2).saturating_add(1),
                    middle,
                    last,
                    output,
                )
            },
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Assigned<T> {
    pub(crate) first: Index,
    pub(crate) last: Index,
    pub(crate) value: T,
}

pub(crate) fn resolve(assignments: Option<Assignments<Properties>>) -> Result<Box<[Stored]>> {
    let Some(assignments) = assignments else {
        return Ok(Box::new([]));
    };
    let ranges = assignments.into_ranges()?;
    let mut stored = Vec::new();
    stored
        .try_reserve_exact(ranges.len())
        .map_err(|source| allocation("effective column ranges", source))?;
    for range in ranges {
        stored.push(Stored {
            first: range.first,
            last: range.last,
            properties: range.value,
        });
    }
    Ok(stored.into_boxed_slice())
}

fn push_range<T: Copy + Eq>(
    output: &mut Vec<(u32, u32, T)>,
    first: u32,
    last: u32,
    value: T,
) -> Result<()> {
    if let Some((_, previous_last, previous_value)) = output.last_mut()
        && *previous_last == first
        && *previous_value == value
    {
        *previous_last = last;
        return Ok(());
    }
    output
        .try_reserve(1)
        .map_err(|source| allocation("column interval map", source))?;
    output.push((first, last, value));
    Ok(())
}

pub(crate) fn entry(ranges: &[Stored], index: Index) -> Option<&Stored> {
    let position = ranges.partition_point(|range| range.last < index);
    ranges.get(position).filter(|range| range.contains(index))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn widths_and_outlines_are_checked_before_they_enter_layout_state() {
        const WIDTH: std::result::Result<Width, WidthError> = Width::new(24.0);
        const OUTLINE: std::result::Result<Outline, OutlineError> = OutlineAt::Level(2).resolve();
        assert_eq!(WIDTH.map(Width::get), Ok(24.0));
        assert_eq!(OUTLINE.map(Outline::get), Ok(2));
        assert_eq!(Width::new(-0.0).map(Width::get), Ok(0.0));
        assert_eq!(WidthAt::from(18.5).resolve().map(Width::get), Ok(18.5));
        assert_eq!(WidthAt::from(24).resolve().map(Width::get), Ok(24.0));
        assert!(Width::new(f64::NAN).is_err());
        assert!(Width::new(f64::INFINITY).is_err());
        assert!(Width::new(255.1).is_err());
        assert_eq!(Outline::new(7).map(Outline::get), Ok(7));
        assert_eq!(OutlineAt::from(3_i32).resolve().map(Outline::get), Ok(3));
        assert_eq!(OutlineAt::from(-1_i32).resolve().unwrap_err().value(), -1);
        assert_eq!(OutlineAt::from(8_u32).resolve().unwrap_err().value(), 8);
    }

    #[test]
    fn later_ranges_replace_the_complete_effective_column_record() {
        let hidden = Properties {
            width: Width::new(20.0).ok(),
            style: None,
            outline: Outline::NONE,
            flags: Flags::HIDDEN | Flags::CUSTOM_WIDTH,
        };
        let visible = Properties {
            width: Width::new(10.0).ok(),
            style: None,
            outline: Outline::NONE,
            flags: Flags::CUSTOM_WIDTH,
        };
        let mut assignments = Assignments::new().expect("interval map");
        assignments.assign(Index::new(1).expect("B"), Index::new(3).expect("D"), hidden);
        assignments.assign(
            Index::new(2).expect("C"),
            Index::new(2).expect("C"),
            visible,
        );

        let stored = resolve(Some(assignments)).expect("stored ranges");
        assert_eq!(stored.len(), 3);
        assert!(
            Column::new(
                Index::new(1).expect("B"),
                entry(&stored, Index::new(1).expect("B"))
            )
            .hidden()
        );
        let column_c = Column::new(
            Index::new(2).expect("C"),
            entry(&stored, Index::new(2).expect("C")),
        );
        assert!(!column_c.hidden());
        assert_eq!(column_c.width().map(Width::get), Some(10.0));
        assert!(
            Column::new(
                Index::new(3).expect("D"),
                entry(&stored, Index::new(3).expect("D"))
            )
            .hidden()
        );
        assert!(entry(&stored, Index::new(4).expect("E")).is_none());

        let mut columns = Columns::new(&stored);
        assert_eq!(columns.len(), 3);
        assert_eq!(columns.next().map(Column::index), Index::new(1).ok());
        assert_eq!(columns.next_back().map(Column::index), Index::new(3).ok());
        assert_eq!(columns.next().map(Column::index), Index::new(2).ok());
        assert!(columns.next().is_none());
    }

    #[test]
    fn bounded_assignment_tree_matches_a_naive_grid_oracle() {
        let mut assignments = Assignments::new().expect("interval map");
        let mut expected = vec![None; COLUMNS as usize];
        for step in 0..1_000u32 {
            let first = step.wrapping_mul(7_919) % COLUMNS;
            let width = step.wrapping_mul(104_729) % 257;
            let last = first.saturating_add(width).min(COLUMNS - 1);
            let value = (step % 11) as u8;
            assignments.assign(
                Index::new(first).expect("first"),
                Index::new(last).expect("last"),
                value,
            );
            expected[first as usize..=last as usize].fill(Some(value));
        }
        for (raw, expected) in expected.iter().copied().enumerate() {
            let index = Index::new(raw as u32).expect("grid index");
            assert_eq!(assignments.get(index), expected, "column {raw}");
        }

        let ranges = assignments.into_ranges().expect("compact ranges");
        let mut actual = vec![None; COLUMNS as usize];
        for range in ranges {
            actual[range.first.get() as usize..=range.last.get() as usize].fill(Some(range.value));
        }
        assert_eq!(actual, expected);
    }
}
