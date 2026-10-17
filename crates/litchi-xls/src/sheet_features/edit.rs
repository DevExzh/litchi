//! The worksheet-feature transaction (record 0327, ADR 0034 scoped family).
//!
//! [`Edit`] stages record-level changes to a working copy of the Workbook
//! stream; [`Edit::commit`] publishes them as a new [`Package`] and an exact
//! [`Patch`] (both package images). Every record a change does not touch
//! keeps its bytes; an edit that changes nothing commits to the source
//! package itself.

use std::fmt;
use std::sync::Arc;

use super::encode::{self, HyperlinkSpec, Rule12, StyleSpec, ValidationSpec};
use super::read::{self, Features};
use super::stream::{
    self, AUTO_FILTER, AUTO_FILTER_INFO, AUTO_FILTER12, CF, CF_EX, CF12, COND_FMT, COND_FMT12,
    CONTINUE_FRT12, DV, DVAL, FILTER_MODE, Globals, HLINK, HLINK_TOOLTIP, HORIZONTAL_PAGE_BREAKS,
    Sheet, VERTICAL_PAGE_BREAKS, u16_at, u32_at,
};
use super::{Area, FeatureSet};
use crate::Package;

pub use super::PageBreak;
pub use encode::{HyperlinkSpec as Hyperlink, StyleSpec as Style, ValidationSpec as Validation};
pub use objects::{
    Anchor, DrawingChange, PictureSpec as Picture, ShapeGeometry, ShapeSpec as Shape,
};

#[path = "objects.rs"]
mod objects;

/// A worksheet-feature edit limit dimension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Staged operations.
    Operations,
    /// Shapes in one worksheet drawing cluster.
    Shapes,
}

/// Typed errors of the worksheet-feature transaction. Messages carry no
/// authored content.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The package could not be read, written or republished.
    Package(crate::package::Error),
    /// The source records are malformed, or a value is outside BIFF8 bounds.
    Records(crate::Error),
    /// An authored value is invalid.
    Invalid(&'static str),
    /// The source has a shape this editor does not rewrite.
    Unsupported(&'static str),
    /// The addressed feature does not exist.
    Missing(&'static str),
    /// A bound was exceeded.
    Limit {
        /// The dimension.
        kind: LimitKind,
        /// The observed amount.
        observed: usize,
        /// The maximum.
        maximum: usize,
    },
    /// A patch's source does not match the package receiving it.
    PatchConflict,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Records(error) => error.fmt(formatter),
            Self::Invalid(message) | Self::Unsupported(message) | Self::Missing(message) => {
                formatter.write_str(message)
            },
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "XLS worksheet-feature limit {kind:?}: {observed} exceeds {maximum}"
            ),
            Self::PatchConflict => formatter.write_str(
                "the worksheet-feature patch source does not match the package receiving it",
            ),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Records(error) => Some(error),
            _ => None,
        }
    }
}

impl From<crate::package::Error> for Error {
    fn from(error: crate::package::Error) -> Self {
        Self::Package(error)
    }
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        match error {
            crate::Error::UnsafeEdit(_) => {
                Self::Unsupported("the source records are not rewritten")
            },
            other => Self::Records(other),
        }
    }
}

/// The result type of the family.
pub type Result<T> = std::result::Result<T, Error>;

/// Content-free facts about one publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operations: usize,
}

impl Diagnostics {
    /// Whether the package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// The number of staged operations.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operations
    }
}

/// The exact package images before and after an edit.
#[derive(Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .finish()
    }
}

impl Patch {
    /// The source package image.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// The published package image.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the images are equal.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        Arc::ptr_eq(&self.before, &self.after) || self.before == self.after
    }

    /// The patch that undoes this one.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }

    /// The bytes the patch retains.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        if Arc::ptr_eq(&self.before, &self.after) {
            self.before.len()
        } else {
            self.before.len().saturating_add(self.after.len())
        }
    }
}

/// A published edit.
#[derive(Debug, Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The published package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The publication facts.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// The published package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The published package and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// The staged features of one package.
pub struct Edit {
    base: Package,
    source: Arc<[u8]>,
    stream: Vec<u8>,
    operations: usize,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operations", &self.operations)
            .finish_non_exhaustive()
    }
}

const MAX_OPERATIONS: usize = 4_096;

/// The full extent of a row break (`max` column) and of a column break.
const LAST_COLUMN: u16 = 255;
const LAST_ROW: u16 = 65_535;

impl Edit {
    pub(super) fn new(base: &Package) -> Result<Self> {
        if base.is_encrypted() {
            return Err(Error::Package(crate::package::Error::EncryptionPolicy));
        }
        let source = base.workbook_stream()?;
        Globals::parse(&source)?;
        Ok(Self {
            base: base.clone(),
            stream: source.to_vec(),
            source,
            operations: 0,
        })
    }

    fn step(&mut self) -> Result<()> {
        self.operations += 1;
        if self.operations > MAX_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Operations,
                observed: self.operations,
                maximum: MAX_OPERATIONS,
            });
        }
        Ok(())
    }

    fn sheet(&self, sheet: usize) -> Result<(Globals, Sheet)> {
        let globals = Globals::parse(&self.stream)?;
        let start = globals.worksheet(sheet)?;
        let layout = Sheet::parse(&self.stream, start)?;
        Ok((globals, layout))
    }

    /// The staged features of the worksheet at `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is not a worksheet or its records
    /// are malformed.
    pub fn features(&self, sheet: usize, select: FeatureSet) -> Result<Features> {
        let (globals, layout) = self.sheet(sheet)?;
        Ok(read::read(&self.stream, &globals, &layout, sheet, select)?)
    }

    /// Replace records `[first, end)` of the current sheet layout with
    /// `bytes`.
    fn replace(&mut self, layout: &Sheet, first: usize, end: usize, bytes: &[u8]) -> Result<()> {
        let start = layout.records[first].start;
        let stop = if end > first {
            layout.records[end - 1].end
        } else {
            start
        };
        stream::splice(&mut self.stream, start, stop, bytes)?;
        Ok(())
    }

    /// Insert `bytes` before record `at` of the current layout.
    fn insert(&mut self, layout: &Sheet, at: usize, bytes: &[u8]) -> Result<()> {
        self.replace(layout, at, at, bytes)
    }

    // -----------------------------------------------------------------------
    // Page breaks

    /// Replace the manual row breaks of `sheet`; an empty list removes them.
    /// A break whose `max` is 0 spans every column.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for more than 1,026 breaks.
    pub fn set_row_breaks(&mut self, sheet: usize, breaks: &[PageBreak]) -> Result<()> {
        self.set_breaks(sheet, HORIZONTAL_PAGE_BREAKS, breaks, LAST_COLUMN)
    }

    /// Replace the manual column breaks of `sheet`; an empty list removes
    /// them. A break whose `max` is 0 spans every row.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for more than 1,026 breaks or a column
    /// outside the grid.
    pub fn set_column_breaks(&mut self, sheet: usize, breaks: &[PageBreak]) -> Result<()> {
        if breaks.iter().any(|item| item.index > 255) {
            return Err(Error::Invalid("a column break is outside the BIFF8 grid"));
        }
        self.set_breaks(sheet, VERTICAL_PAGE_BREAKS, breaks, LAST_ROW)
    }

    fn set_breaks(
        &mut self,
        sheet: usize,
        kind: u16,
        breaks: &[PageBreak],
        last: u16,
    ) -> Result<()> {
        self.step()?;
        let mut breaks: Vec<PageBreak> = breaks
            .iter()
            .map(|item| PageBreak {
                index: item.index,
                min: item.min,
                max: if item.max == 0 { last } else { item.max },
            })
            .collect();
        breaks.sort_by_key(|item| item.index);
        breaks.dedup_by_key(|item| item.index);
        if breaks
            .iter()
            .any(|item| item.min > item.max || item.index == 0)
        {
            return Err(Error::Invalid(
                "a page break span is reversed or before the first cell",
            ));
        }
        let (_, layout) = self.sheet(sheet)?;
        let existing = layout.records[..layout.dimensions]
            .iter()
            .position(|record| record.kind == kind);
        let bytes = if breaks.is_empty() {
            Vec::new()
        } else {
            encode::page_breaks(kind, &breaks)
                .map_err(|_| Error::Invalid("too many page breaks"))?
        };
        match existing {
            Some(index) => self.replace(&layout, index, index + 1, &bytes),
            None if bytes.is_empty() => Ok(()),
            None => {
                let at = layout.breaks_insert(kind)?;
                self.insert(&layout, at, &bytes)
            },
        }
    }

    // -----------------------------------------------------------------------
    // Hyperlinks

    /// Each hyperlink of the layout: its first record and its end.
    fn hyperlink_units(layout: &Sheet) -> Vec<(usize, usize)> {
        let mut units = Vec::new();
        for (index, record) in layout.tail() {
            if record.kind == HLINK {
                let end = if layout
                    .records
                    .get(index + 1)
                    .is_some_and(|next| next.kind == HLINK_TOOLTIP)
                {
                    index + 2
                } else {
                    index + 1
                };
                units.push((index, end));
            }
        }
        units
    }

    /// Add a hyperlink after the existing ones; returns its position.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for a link without a target or a value
    /// outside its BIFF8 bound.
    pub fn add_hyperlink(&mut self, sheet: usize, link: &HyperlinkSpec) -> Result<usize> {
        self.step()?;
        let bytes = encode::hyperlink(link)
            .map_err(|_| Error::Invalid("the hyperlink cannot be encoded"))?;
        let (_, layout) = self.sheet(sheet)?;
        let count = Self::hyperlink_units(&layout).len();
        let at = layout.hyperlinks_insert();
        self.insert(&layout, at, &bytes)?;
        Ok(count)
    }

    /// Replace the hyperlink at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such hyperlink.
    pub fn set_hyperlink(
        &mut self,
        sheet: usize,
        index: usize,
        link: &HyperlinkSpec,
    ) -> Result<()> {
        self.step()?;
        let bytes = encode::hyperlink(link)
            .map_err(|_| Error::Invalid("the hyperlink cannot be encoded"))?;
        let (_, layout) = self.sheet(sheet)?;
        let (first, end) = *Self::hyperlink_units(&layout)
            .get(index)
            .ok_or(Error::Missing("the hyperlink does not exist"))?;
        self.replace(&layout, first, end, &bytes)
    }

    /// Remove the hyperlink at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such hyperlink.
    pub fn remove_hyperlink(&mut self, sheet: usize, index: usize) -> Result<()> {
        self.step()?;
        let (_, layout) = self.sheet(sheet)?;
        let (first, end) = *Self::hyperlink_units(&layout)
            .get(index)
            .ok_or(Error::Missing("the hyperlink does not exist"))?;
        self.replace(&layout, first, end, &[])
    }

    // -----------------------------------------------------------------------
    // Data validations

    fn validation_records(layout: &Sheet) -> (Option<usize>, Vec<usize>) {
        let dval = layout
            .tail()
            .find(|(_, record)| record.kind == DVAL)
            .map(|(index, _)| index);
        let rules = layout
            .tail()
            .filter(|(_, record)| record.kind == DV)
            .map(|(index, _)| index)
            .collect();
        (dval, rules)
    }

    fn dval_count(&mut self, layout: &Sheet, dval: usize, count: usize) -> Result<()> {
        let record = layout.records[dval];
        let mut payload = record.payload(&self.stream).to_vec();
        if payload.len() < 18 {
            return Err(Error::Unsupported("the DVal record is malformed"));
        }
        payload[14..18].copy_from_slice(&u32::try_from(count).unwrap_or(u32::MAX).to_le_bytes());
        let bytes = stream::record(DVAL, &payload)?;
        self.replace(layout, dval, dval + 1, &bytes)
    }

    /// Add a data validation after the existing ones; returns its position.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for a validation outside its BIFF8 bounds.
    pub fn add_validation(&mut self, sheet: usize, spec: &ValidationSpec) -> Result<usize> {
        self.step()?;
        let bytes = encode::validation(spec)
            .map_err(|_| Error::Invalid("the validation cannot be encoded"))?;
        let (_, layout) = self.sheet(sheet)?;
        let (dval, rules) = Self::validation_records(&layout);
        let position = rules.len();
        match (dval, rules.last()) {
            (Some(dval), last) => {
                let at = last.map_or(dval + 1, |last| last + 1);
                self.insert(&layout, at, &bytes)?;
                let (_, layout) = self.sheet(sheet)?;
                self.dval_count(&layout, dval, position + 1)?;
            },
            (None, _) => {
                let mut block = encode::dval(1)?;
                block.extend(bytes);
                let at = layout.validations_insert();
                self.insert(&layout, at, &block)?;
            },
        }
        Ok(position)
    }

    /// Replace the data validation at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such validation.
    pub fn set_validation(
        &mut self,
        sheet: usize,
        index: usize,
        spec: &ValidationSpec,
    ) -> Result<()> {
        self.step()?;
        let bytes = encode::validation(spec)
            .map_err(|_| Error::Invalid("the validation cannot be encoded"))?;
        let (_, layout) = self.sheet(sheet)?;
        let (_, rules) = Self::validation_records(&layout);
        let at = *rules
            .get(index)
            .ok_or(Error::Missing("the validation does not exist"))?;
        self.replace(&layout, at, at + 1, &bytes)
    }

    /// Remove the data validation at `index`; the last one takes its `DVal`
    /// with it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such validation.
    pub fn remove_validation(&mut self, sheet: usize, index: usize) -> Result<()> {
        self.step()?;
        let (_, layout) = self.sheet(sheet)?;
        let (dval, rules) = Self::validation_records(&layout);
        let at = *rules
            .get(index)
            .ok_or(Error::Missing("the validation does not exist"))?;
        let dval = dval.ok_or(Error::Unsupported("a Dv record has no DVal"))?;
        if rules.len() == 1 {
            if at != dval + 1 {
                return Err(Error::Unsupported("the DVal block is not contiguous"));
            }
            return self.replace(&layout, dval, at + 1, &[]);
        }
        self.replace(&layout, at, at + 1, &[])?;
        let (_, layout) = self.sheet(sheet)?;
        self.dval_count(&layout, dval, rules.len() - 1)
    }

    // -----------------------------------------------------------------------
    // Conditional formats

    /// Each conditional format of the layout: its header record, its end,
    /// whether it is a `CondFmt12`, and its identifier.
    fn format_units(&self, layout: &Sheet) -> Vec<(usize, usize, bool, u16)> {
        let mut units = Vec::new();
        let records = &layout.records;
        for (index, record) in layout.tail() {
            let future = match record.kind {
                COND_FMT => false,
                COND_FMT12 => true,
                _ => continue,
            };
            let rule = if future { CF12 } else { CF };
            let mut end = index + 1;
            while records.get(end).is_some_and(|next| next.kind == rule) {
                end += 1;
            }
            let payload = record.payload(&self.stream);
            let body = if future {
                &payload[payload.len().min(12)..]
            } else {
                payload
            };
            let identifier = u16_at(body, 2) >> 1;
            units.push((index, end, future, identifier));
        }
        units
    }

    fn next_format_identifier(&self, layout: &Sheet) -> u16 {
        self.format_units(layout)
            .iter()
            .map(|unit| unit.3)
            .max()
            .map_or(1, |max| max.saturating_add(1))
    }

    /// The next free CF12/CFEx priority of the layout.
    pub(super) fn next_priority(&self, layout: &Sheet) -> u16 {
        let mut max = 0u16;
        let mut classic = 0u16;
        for (_, record) in layout.tail() {
            let payload = record.payload(&self.stream);
            match record.kind {
                CF => classic = classic.saturating_add(1),
                CF12 => {
                    if let Some(priority) = cf12_priority(payload) {
                        max = max.max(priority);
                    }
                },
                CF_EX if payload.len() >= 24 => {
                    // CFEx: frtRefHeaderU (12), fIsCF12 (4), nID (2), then
                    // CFExNonCF12 (icf, cp, icfTemplate, ipriority, …).
                    if u32_at(payload, 12) == 0 && payload.len() >= 24 {
                        max = max.max(u16_at(payload, 22));
                    }
                },
                _ => {},
            }
        }
        max.max(classic).saturating_add(1)
    }

    /// Add a conditional format with one rule after the existing ones;
    /// returns its position.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for a format outside its BIFF8 bounds.
    pub fn add_format(&mut self, sheet: usize, format: &Format) -> Result<usize> {
        self.step()?;
        let (_, layout) = self.sheet(sheet)?;
        let identifier = self.next_format_identifier(&layout);
        let priority = self.next_priority(&layout);
        let bytes = format.encode(identifier, priority)?;
        let position = self.format_units(&layout).len();
        let at = layout.formats_insert();
        self.insert(&layout, at, &bytes)?;
        Ok(position)
    }

    /// Replace the conditional format at `index` with `format` (one rule);
    /// its identifier and priority are kept.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such format and
    /// [`Error::Unsupported`] when it holds more than one rule or a `CFEx`
    /// extends it.
    pub fn set_format(&mut self, sheet: usize, index: usize, format: &Format) -> Result<()> {
        self.step()?;
        let (_, layout) = self.sheet(sheet)?;
        let (first, end, future, identifier) = *self
            .format_units(&layout)
            .get(index)
            .ok_or(Error::Missing("the conditional format does not exist"))?;
        if end - first != 2 {
            return Err(Error::Unsupported(
                "a conditional format with several rules is replaced only as a whole",
            ));
        }
        if !future && self.extended(&layout, identifier) {
            return Err(Error::Unsupported(
                "a conditional format a CFEx record extends is not rewritten",
            ));
        }
        let priority = if future {
            cf12_priority(layout.records[first + 1].payload(&self.stream))
                .unwrap_or_else(|| self.next_priority(&layout))
        } else {
            self.next_priority(&layout)
        };
        let bytes = format.encode(identifier, priority)?;
        self.replace(&layout, first, end, &bytes)
    }

    /// Change only the ranges of the conditional format at `index`; its rules
    /// keep their bytes.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such format.
    pub fn set_format_ranges(&mut self, sheet: usize, index: usize, ranges: &[Area]) -> Result<()> {
        self.step()?;
        let (_, layout) = self.sheet(sheet)?;
        let (first, end, future, identifier) = *self
            .format_units(&layout)
            .get(index)
            .ok_or(Error::Missing("the conditional format does not exist"))?;
        let rules = u16::try_from(end - first - 1).unwrap_or(u16::MAX);
        let header = layout.records[first].payload(&self.stream);
        let tough = if future {
            header.get(14)
        } else {
            header.get(2)
        }
        .is_some_and(|flags| flags & 1 != 0);
        let body = encode::cond_fmt_header(ranges, rules, identifier, tough)
            .map_err(|_| Error::Invalid("the conditional-format ranges are invalid"))?;
        let bytes = if future {
            let mut payload = header[..4].to_vec();
            payload.extend_from_slice(&body[4..12]);
            payload.extend(body);
            stream::record(COND_FMT12, &payload)?
        } else {
            stream::record(COND_FMT, &body)?
        };
        self.replace(&layout, first, first + 1, &bytes)
    }

    fn extended(&self, layout: &Sheet, identifier: u16) -> bool {
        layout.tail().any(|(_, record)| {
            record.kind == CF_EX && u16_at(record.payload(&self.stream), 16) == identifier
        })
    }

    /// Remove the conditional format at `index` with the `CFEx` records that
    /// extend it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Missing`] when there is no such format.
    pub fn remove_format(&mut self, sheet: usize, index: usize) -> Result<()> {
        self.step()?;
        let (_, layout) = self.sheet(sheet)?;
        let (first, end, future, identifier) = *self
            .format_units(&layout)
            .get(index)
            .ok_or(Error::Missing("the conditional format does not exist"))?;
        // CFEx records of a classic format (and a CF12 that follows one) go
        // first, from the end, so the format's own indexes stay valid.
        if !future {
            let mut extensions: Vec<(usize, usize)> = Vec::new();
            for (at, record) in layout.tail() {
                if record.kind == CF_EX && u16_at(record.payload(&self.stream), 16) == identifier {
                    let is_cf12 = u32_at(record.payload(&self.stream), 12) != 0;
                    let stop = if is_cf12
                        && layout
                            .records
                            .get(at + 1)
                            .is_some_and(|next| next.kind == CF12)
                    {
                        at + 2
                    } else {
                        at + 1
                    };
                    extensions.push((at, stop));
                }
            }
            for (at, stop) in extensions.into_iter().rev() {
                let (_, layout) = self.sheet(sheet)?;
                self.replace(&layout, at, stop, &[])?;
            }
        }
        let (_, layout) = self.sheet(sheet)?;
        self.replace(&layout, first, end, &[])
    }

    // -----------------------------------------------------------------------
    // Auto-filter

    /// Set the auto-filter of `sheet` to `area`, or remove it with `None`.
    /// A new filter writes the sheet's `_FilterDatabase` name (and the
    /// internal `SupBook`/`ExternSheet` it references when the workbook has
    /// none), `AutoFilterInfo` and one drop-down object per column.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for a range outside the grid and
    /// [`Error::Unsupported`] for source layouts the editor does not rewrite.
    pub fn set_auto_filter(&mut self, sheet: usize, area: Option<Area>) -> Result<()> {
        self.step()?;
        if let Some(area) = area {
            if area.last_column > 255 || area.last_column - area.first_column > 255 {
                return Err(Error::Invalid(
                    "the auto-filter range is outside the BIFF8 grid",
                ));
            }
        }
        let (globals, layout) = self.sheet(sheet)?;
        let current = layout
            .find(AUTO_FILTER_INFO)
            .map(|_| read::filter_database(&self.stream, &globals, sheet).map(|(_, area)| area));
        if current == Some(area.map(Some).unwrap_or(None)) && area.is_some() {
            return Ok(());
        }
        // Drop the existing filter records and drop-down objects.
        if current.is_some() {
            self.remove_dropdowns(sheet)?;
            let (_, layout) = self.sheet(sheet)?;
            let doomed: Vec<usize> = layout
                .records
                .iter()
                .enumerate()
                .take(layout.dimensions)
                .filter(|(_, record)| {
                    matches!(
                        record.kind,
                        AUTO_FILTER_INFO
                            | AUTO_FILTER
                            | AUTO_FILTER12
                            | CONTINUE_FRT12
                            | FILTER_MODE
                    )
                })
                .map(|(index, _)| index)
                .collect();
            for index in doomed.into_iter().rev() {
                let (_, layout) = self.sheet(sheet)?;
                self.replace(&layout, index, index + 1, &[])?;
            }
        }
        let Some(area) = area else {
            return Ok(());
        };
        self.put_filter_database(sheet, area)?;
        let (_, layout) = self.sheet(sheet)?;
        let columns = area.last_column - area.first_column + 1;
        let bytes = encode::auto_filter_info(columns)?;
        self.insert(&layout, layout.dimensions, &bytes)?;
        self.add_dropdowns(sheet, area)
    }

    /// Write or update the `_FilterDatabase` name of `sheet`.
    fn put_filter_database(&mut self, sheet: usize, area: Area) -> Result<()> {
        let globals = Globals::parse(&self.stream)?;
        let xti = self.internal_xti(sheet)?;
        let globals = if xti.1 {
            Globals::parse(&self.stream)?
        } else {
            globals
        };
        let xti = xti.0;
        let name = encode::filter_database(sheet, xti, area)?;
        if let Some((index, _)) = read::filter_database(&self.stream, &globals, sheet) {
            let record = globals.records[index];
            stream::splice(&mut self.stream, record.start, record.end, &name)?;
            return Ok(());
        }
        let at = globals
            .records
            .iter()
            .rposition(|record| matches!(record.kind, stream::LBL | stream::EXTERN_SHEET))
            .map(|index| globals.records[index].end)
            .ok_or(Error::Unsupported(
                "the workbook globals have no ExternSheet",
            ))?;
        stream::splice(&mut self.stream, at, at, &name)?;
        Ok(())
    }

    /// The `ExternSheet` index of the internal reference to `sheet`, added
    /// (with the internal `SupBook` and `ExternSheet`) when missing; the
    /// flag says whether the globals changed.
    fn internal_xti(&mut self, sheet: usize) -> Result<(u16, bool)> {
        let globals = Globals::parse(&self.stream)?;
        let sheet_u16 = u16::try_from(sheet)
            .map_err(|_| Error::Invalid("the sheet position exceeds 16 bits"))?;
        let books: Vec<usize> = globals
            .records
            .iter()
            .enumerate()
            .filter(|(_, record)| record.kind == stream::SUP_BOOK)
            .map(|(index, _)| index)
            .collect();
        let internal = books
            .iter()
            .position(|index| u16_at(globals.records[*index].payload(&self.stream), 2) == 0x0401);
        let extern_sheet = globals.first(stream::EXTERN_SHEET);
        match (internal, extern_sheet) {
            (Some(book), Some(index)) => {
                let record = globals.records[index];
                let payload = record.payload(&self.stream);
                if globals
                    .records
                    .get(index + 1)
                    .is_some_and(|next| next.kind == stream::CONTINUE)
                {
                    return Err(Error::Unsupported(
                        "a continued ExternSheet is not rewritten",
                    ));
                }
                let count = usize::from(u16_at(payload, 0));
                let book = u16::try_from(book).unwrap_or(u16::MAX);
                for xti in 0..count {
                    let at = 2 + xti * 6;
                    if u16_at(payload, at) == book
                        && u16_at(payload, at + 2) == sheet_u16
                        && u16_at(payload, at + 4) == sheet_u16
                    {
                        return Ok((u16::try_from(xti).unwrap_or(u16::MAX), false));
                    }
                }
                if count >= 1_370 {
                    return Err(Error::Unsupported("the ExternSheet record is full"));
                }
                let mut grown = payload.to_vec();
                grown[0..2].copy_from_slice(&u16::try_from(count + 1).unwrap_or(0).to_le_bytes());
                grown.extend_from_slice(&book.to_le_bytes());
                grown.extend_from_slice(&sheet_u16.to_le_bytes());
                grown.extend_from_slice(&sheet_u16.to_le_bytes());
                let bytes = stream::record(stream::EXTERN_SHEET, &grown)?;
                stream::splice(&mut self.stream, record.start, record.end, &bytes)?;
                Ok((u16::try_from(count).unwrap_or(u16::MAX), true))
            },
            (None, None) => {
                // An internal SupBook and an ExternSheet with one XTI, before
                // the first defined name or the records that follow names.
                let tabs = u16::try_from(globals.bounds.len()).unwrap_or(u16::MAX);
                let mut bytes = stream::record(
                    stream::SUP_BOOK,
                    &[tabs.to_le_bytes(), 0x0401u16.to_le_bytes()].concat(),
                )?;
                let mut xti = Vec::with_capacity(8);
                xti.extend_from_slice(&1u16.to_le_bytes());
                xti.extend_from_slice(&0u16.to_le_bytes());
                xti.extend_from_slice(&sheet_u16.to_le_bytes());
                xti.extend_from_slice(&sheet_u16.to_le_bytes());
                bytes.extend(stream::record(stream::EXTERN_SHEET, &xti)?);
                let following = [
                    stream::LBL,
                    0x0813, // RTD
                    0x01c1, // RecalcId
                    0x0866, // HFPicture
                    stream::MSO_DRAWING_GROUP,
                    stream::SST,
                    stream::EXT_SST,
                    stream::EOF,
                ];
                let at = globals
                    .records
                    .iter()
                    .position(|record| following.contains(&record.kind))
                    .map(|index| globals.records[index].start)
                    .ok_or(Error::Unsupported("the workbook globals have no EOF"))?;
                stream::splice(&mut self.stream, at, at, &bytes)?;
                Ok((0, true))
            },
            _ => Err(Error::Unsupported(
                "a workbook with only external references takes no new filter name",
            )),
        }
    }

    // -----------------------------------------------------------------------
    // Publication

    /// Publish the staged changes.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited stream cannot be written into the
    /// package or the result does not reopen.
    pub fn commit(self) -> Result<Commit> {
        let operations = self.operations;
        if self.stream.as_slice() == &*self.source {
            let bytes: Arc<[u8]> = Arc::from(self.base.to_bytes()?);
            return Ok(Commit {
                package: self.base,
                patch: Patch {
                    before: Arc::clone(&bytes),
                    after: bytes,
                },
                diagnostics: Diagnostics {
                    changed: false,
                    operations,
                },
            });
        }
        let before: Arc<[u8]> = Arc::from(self.base.to_bytes()?);
        let name = workbook_stream_name(&before)?;
        let mut editor = litchi_ole_common::object::Editor::open(
            before.to_vec(),
            litchi_ole_common::object::Targets::default(),
            litchi_ole_common::object::Limits::default(),
        )
        .map_err(|error| Error::Records(error.into()))?;
        let rendered = editor
            .put_stream_shared_with_rendered(&[name.to_owned()], Arc::from(self.stream))
            .map_err(|error| Error::Records(error.into()))?;
        let after: Arc<[u8]> = Arc::from(rendered.as_slice());
        let package = crate::package::republish(&self.base, rendered)?;
        Ok(Commit {
            package,
            patch: Patch { before, after },
            diagnostics: Diagnostics {
                changed: true,
                operations,
            },
        })
    }
}

fn workbook_stream_name(bytes: &[u8]) -> Result<&'static str> {
    let ole = litchi_cfb::OleFile::open(std::io::Cursor::new(bytes))
        .map_err(|error| Error::Records(crate::Error::Cfb(error)))?;
    Ok(if ole.exists(&["Workbook"]) {
        "Workbook"
    } else {
        "Book"
    })
}

fn cf12_priority(payload: &[u8]) -> Option<u16> {
    // Walk the CF12 layout to `ipriority`.
    let cce1 = usize::from(u16_at(payload, 14));
    let cce2 = usize::from(u16_at(payload, 16));
    let cb = u32_at(payload, 18) as usize;
    let mut at = 22 + if cb == 0 { 2 } else { cb };
    at += cce1 + cce2;
    let active = usize::from(u16_at(payload, at));
    at += 2 + active + 1;
    (at + 2 <= payload.len()).then(|| u16_at(payload, at))
}

/// Compile a conditional-format or validation formula (text without a
/// leading `=`) into BIFF8 tokens: same-sheet references, numbers, strings,
/// operators and the built-in worksheet functions (not `CALL`,
/// `REGISTER.ID` or `EXEC`; add-in functions such as `ISODD` are not
/// built-in in BIFF8).
///
/// # Errors
///
/// Returns [`Error::Invalid`] for a formula the tokenizer cannot encode.
pub fn compile_formula(formula: &str) -> Result<Vec<u8>> {
    let tokens = crate::writer::FormulaTokenizer::extended()
        .tokenize(formula)
        .map_err(|_| Error::Invalid("the formula cannot be encoded as BIFF8 tokens"))?;
    crate::writer::formula::encode_ptg_tokens_by_arity(&tokens)
        .map_err(|_| Error::Invalid("the formula cannot be encoded as BIFF8 tokens"))
}

/// Apply a patch to the exact package it was made from.
pub(super) fn apply(package: &Package, patch: &Patch) -> Result<Package> {
    let bytes = package.to_bytes()?;
    if bytes.as_slice() != patch.before() {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(package.clone());
    }
    Ok(crate::package::republish(package, patch.after().to_vec())?)
}

/// One conditional format with one rule.
#[derive(Debug, Clone, PartialEq)]
pub struct Format {
    /// The ranges it applies to.
    pub ranges: Vec<Area>,
    /// The rule.
    pub rule: FormatRule,
    /// The formatting.
    pub style: StyleSpec,
    /// `fStopIfTrue` (Office 2007 rules only).
    pub stop_if_true: bool,
}

/// The rule of a [`Format`]. Formulas are BIFF8 token streams.
#[derive(Debug, Clone, PartialEq)]
pub enum FormatRule {
    /// A classic cell-value comparison (`cp` 1–8).
    CellIs {
        /// `cp`.
        operator: u8,
        /// The operands.
        formulas: [Vec<u8>; 2],
    },
    /// A classic formula rule.
    Expression(Vec<u8>),
    /// An Office 2007 rule with a template (`icfTemplate`), its parameters
    /// and, for a formula template, its formula.
    Template {
        /// `icfTemplate`.
        template: u16,
        /// `rgbTemplateParms`.
        params: [u8; 16],
        /// The formula of the template.
        formula: Vec<u8>,
    },
    /// A top/bottom N filter (`ct` 5).
    TopN {
        /// Top (not bottom).
        top: bool,
        /// A percentage.
        percent: bool,
        /// The rank.
        rank: u16,
    },
    /// A color scale: thresholds (`cfvoType`, number, formula) and colors.
    ColorScale(Vec<(Threshold, u32)>),
    /// A data bar.
    DataBar {
        /// The bar color.
        color: u32,
        /// The minimum.
        min: Threshold,
        /// The maximum.
        max: Threshold,
        /// `iPercentMin`.
        min_length: u8,
        /// `iPercentMax`.
        max_length: u8,
        /// `fShowValue`.
        show_value: bool,
    },
    /// An icon set.
    IconSet {
        /// `iIconSet`.
        set: u8,
        /// The thresholds after the first icon.
        thresholds: Vec<Threshold>,
        /// `fReverse`.
        reverse: bool,
        /// Whether the value is shown.
        show_value: bool,
    },
}

/// A threshold to encode: `cfvoType` with a number or formula tokens.
#[derive(Debug, Clone, PartialEq)]
pub struct Threshold {
    /// `cfvoType`.
    pub kind: u8,
    /// The number.
    pub number: Option<f64>,
    /// The formula tokens.
    pub formula: Option<Vec<u8>>,
}

impl Threshold {
    fn encode(&self, out: &mut Vec<u8>) {
        encode::cfvo(out, self.kind, self.number, self.formula.as_deref());
    }
}

impl Format {
    fn encode(&self, identifier: u16, priority: u16) -> Result<Vec<u8>> {
        let invalid = |_| Error::Invalid("the conditional format cannot be encoded");
        match &self.rule {
            FormatRule::CellIs { operator, formulas } if !self.stop_if_true => {
                encode::classic_format(
                    &self.ranges,
                    identifier,
                    1,
                    *operator,
                    [&formulas[0], &formulas[1]],
                    &self.style,
                )
                .map_err(invalid)
            },
            FormatRule::Expression(formula) if !self.stop_if_true => {
                encode::classic_format(&self.ranges, identifier, 2, 0, [formula, &[]], &self.style)
                    .map_err(invalid)
            },
            rule => {
                let mut cf12 = Rule12 {
                    ct: 2,
                    cp: 0,
                    formulas: [Vec::new(), Vec::new()],
                    style: self.style.clone(),
                    stop_if_true: self.stop_if_true,
                    priority,
                    template: 1,
                    params: [0; 16],
                    payload: Vec::new(),
                };
                match rule {
                    FormatRule::CellIs { operator, formulas } => {
                        cf12.ct = 1;
                        cf12.cp = *operator;
                        cf12.template = 0;
                        cf12.formulas = formulas.clone();
                    },
                    FormatRule::Expression(formula) => cf12.formulas[0] = formula.clone(),
                    FormatRule::Template {
                        template,
                        params,
                        formula,
                    } => {
                        cf12.template = *template;
                        cf12.params = *params;
                        cf12.formulas[0] = formula.clone();
                    },
                    FormatRule::TopN { top, percent, rank } => {
                        cf12.ct = 5;
                        cf12.template = 5;
                        let flags = u8::from(*top) | (u8::from(*percent) << 1);
                        cf12.params[0] = flags;
                        cf12.params[1..3].copy_from_slice(&rank.to_le_bytes());
                        cf12.payload.extend_from_slice(&4u16.to_le_bytes());
                        cf12.payload.push(0);
                        cf12.payload.push(flags);
                        cf12.payload.extend_from_slice(&rank.to_le_bytes());
                    },
                    FormatRule::ColorScale(points) => {
                        if !(2..=3).contains(&points.len()) {
                            return Err(Error::Invalid("a color scale has two or three points"));
                        }
                        cf12.ct = 3;
                        cf12.template = 2;
                        cf12.style = StyleSpec::default();
                        cf12.stop_if_true = false;
                        let count = points.len() as u8;
                        cf12.payload
                            .extend_from_slice(&[0, 0, 0, count, count, 0x03]);
                        let domain = |index: usize| -> f64 {
                            match (index, points.len()) {
                                (0, _) => 0.0,
                                (1, 3) => 0.5,
                                _ => 1.0,
                            }
                        };
                        for (index, (threshold, _)) in points.iter().enumerate() {
                            threshold.encode(&mut cf12.payload);
                            cf12.payload.extend_from_slice(&domain(index).to_le_bytes());
                        }
                        for (index, (_, color)) in points.iter().enumerate() {
                            cf12.payload.extend_from_slice(&domain(index).to_le_bytes());
                            encode::cf_color(&mut cf12.payload, *color);
                        }
                    },
                    FormatRule::DataBar {
                        color,
                        min,
                        max,
                        min_length,
                        max_length,
                        show_value,
                    } => {
                        if min_length >= max_length || *max_length > 100 {
                            return Err(Error::Invalid(
                                "a data bar's lengths must rise to at most 100",
                            ));
                        }
                        cf12.ct = 4;
                        cf12.template = 3;
                        cf12.style = StyleSpec::default();
                        cf12.stop_if_true = false;
                        cf12.payload.extend_from_slice(&[
                            0,
                            0,
                            0,
                            if *show_value { 2 } else { 0 },
                            *min_length,
                            *max_length,
                        ]);
                        encode::cf_color(&mut cf12.payload, *color);
                        min.encode(&mut cf12.payload);
                        max.encode(&mut cf12.payload);
                    },
                    FormatRule::IconSet {
                        set,
                        thresholds,
                        reverse,
                        show_value,
                    } => {
                        cf12.ct = 6;
                        cf12.template = 4;
                        cf12.style = StyleSpec::default();
                        cf12.stop_if_true = false;
                        let count = u8::try_from(thresholds.len()).unwrap_or(0);
                        let flags = u8::from(!*show_value) | (u8::from(*reverse) << 2);
                        cf12.payload
                            .extend_from_slice(&[0, 0, 0, count, *set, flags]);
                        for threshold in thresholds {
                            threshold.encode(&mut cf12.payload);
                            // fEqual: the first icon's threshold is >=.
                            cf12.payload.push(1);
                            cf12.payload.extend_from_slice(&[0; 4]);
                        }
                    },
                }
                encode::future_format(&self.ranges, identifier, &cf12).map_err(invalid)
            },
        }
    }
}
