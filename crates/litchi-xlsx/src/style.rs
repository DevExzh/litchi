//! Shared cell-format handles without native SpreadsheetML identifiers.

/// Applied cell-format views used by worksheet-facing APIs.
pub mod format;
/// The parsed `styles.xml` resource graph.
pub mod stylesheet;
/// The DrawingML theme resource associated with a workbook.
pub mod theme;

use std::fmt;
use std::hash::{Hash, Hasher};
use std::ops::Range;
use std::sync::Arc;

use crate::error::Result;
use crate::workbook::Inner;

/// Opaque style identity used by semantic patch states.
///
/// The physical SpreadsheetML index remains private. Keys are meaningful only
/// within the patch lineage that produced them.
#[derive(Debug)]
pub(crate) struct StyleLineage;

#[derive(Clone)]
pub struct StyleKey {
    raw: u32,
    lineage: Arc<StyleLineage>,
}

impl StyleKey {
    pub(crate) fn new(raw: u32, lineage: Arc<StyleLineage>) -> Self {
        Self { raw, lineage }
    }
}

impl PartialEq for StyleKey {
    fn eq(&self, other: &Self) -> bool {
        self.raw == other.raw && Arc::ptr_eq(&self.lineage, &other.lineage)
    }
}

impl Eq for StyleKey {}

impl Hash for StyleKey {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.raw.hash(state);
        Arc::as_ptr(&self.lineage).hash(state);
    }
}

impl fmt::Debug for StyleKey {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("StyleKey(..)")
    }
}

/// Exact local style state for a stored cell.
///
/// `Default` means the cell has no local `s` attribute. It remains distinct
/// from an explicit reference to the workbook's base shared style.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum LocalStyle {
    Default,
    Shared(Style),
}

/// Exact local style identity recorded in a semantic patch state.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum StyleState {
    Default,
    Shared(StyleKey),
}

impl StyleState {
    pub(crate) fn rebind(&mut self, lineage: &Arc<StyleLineage>) {
        if let Self::Shared(key) = self {
            key.lineage = Arc::clone(lineage);
        }
    }
}

/// Lazy, immutable view of the workbook's shared cell formats.
#[derive(Clone)]
pub struct Styles {
    owner: Arc<Inner>,
    len: u32,
}

impl Styles {
    pub(crate) fn new(owner: Arc<Inner>, len: u32) -> Self {
        Self { owner, len }
    }

    /// Number of shared cell formats.
    pub const fn len(&self) -> usize {
        self.len as usize
    }

    pub const fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Base shared cell format used by cells without an explicit local style.
    pub fn base(&self) -> Option<Style> {
        self.get(0)
    }

    /// Checked zero-based physical-position lookup for diagnostics and import.
    ///
    /// Copying a style obtained from [`crate::Worksheet::style`] is the preferred
    /// semantic entry point.
    pub fn get(&self, position: usize) -> Option<Style> {
        let key = u32::try_from(position).ok().filter(|key| *key < self.len)?;
        Some(Style {
            owner: Arc::clone(&self.owner),
            raw: key,
        })
    }

    /// Resolve an opaque key when it belongs to this shared-style table.
    pub fn find(&self, key: &StyleKey) -> Option<Style> {
        (key.raw < self.len && Arc::ptr_eq(&key.lineage, &self.owner.style_lineage)).then(|| {
            Style {
                owner: Arc::clone(&self.owner),
                raw: key.raw,
            }
        })
    }

    pub fn iter(&self) -> StylesIter {
        StylesIter {
            owner: Arc::clone(&self.owner),
            remaining: 0..self.len,
        }
    }
}

impl fmt::Debug for Styles {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Styles")
            .field("len", &self.len)
            .finish_non_exhaustive()
    }
}

impl IntoIterator for &Styles {
    type Item = Style;
    type IntoIter = StylesIter;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Iterator over cheap shared-style handles.
#[derive(Debug, Clone)]
pub struct StylesIter {
    owner: Arc<Inner>,
    remaining: Range<u32>,
}

impl Iterator for StylesIter {
    type Item = Style;

    fn next(&mut self) -> Option<Self::Item> {
        self.remaining.next().map(|key| Style {
            owner: Arc::clone(&self.owner),
            raw: key,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.remaining.size_hint()
    }
}

impl DoubleEndedIterator for StylesIter {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.remaining.next_back().map(|key| Style {
            owner: Arc::clone(&self.owner),
            raw: key,
        })
    }
}

impl ExactSizeIterator for StylesIter {}
impl std::iter::FusedIterator for StylesIter {}

/// Lineage-checked handle to one immutable shared cell format.
#[derive(Clone)]
pub struct Style {
    pub(crate) owner: Arc<Inner>,
    raw: u32,
}

impl Style {
    pub(crate) fn from_raw(owner: Arc<Inner>, key: u32) -> Self {
        Self { owner, raw: key }
    }

    pub(crate) const fn raw(&self) -> u32 {
        self.raw
    }

    /// Opaque, lineage-checked identity for maps and semantic patch states.
    pub fn key(&self) -> StyleKey {
        StyleKey::new(self.raw, Arc::clone(&self.owner.style_lineage))
    }

    /// Whether two handles name the same format in the same shared-style lineage.
    ///
    /// This remains true across descendant snapshots whose shared-style table
    /// is byte-for-byte unchanged. Use [`Self::same_workbook`] when exact
    /// snapshot identity matters.
    pub fn same(&self, other: &Self) -> bool {
        self.raw == other.raw && Arc::ptr_eq(&self.owner.style_lineage, &other.owner.style_lineage)
    }

    /// Whether two handles belong to the same immutable workbook snapshot.
    pub fn same_workbook(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.owner, &other.owner)
    }

    /// Count stored worksheet cells whose effective format is this resource.
    ///
    /// The base style includes stored cells with no explicit local style.
    /// Row/column defaults and unstored grid positions are intentionally not
    /// approximated.
    pub fn fan_out(&self) -> Result<usize> {
        self.owner.style_fan_out(self.raw)
    }
}

impl PartialEq for Style {
    fn eq(&self, other: &Self) -> bool {
        self.same(other)
    }
}

impl Eq for Style {}

impl Hash for Style {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.raw.hash(state);
        Arc::as_ptr(&self.owner.style_lineage).hash(state);
    }
}

impl fmt::Debug for Style {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Style").finish_non_exhaustive()
    }
}
