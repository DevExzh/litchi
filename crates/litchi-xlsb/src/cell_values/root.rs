//! Durable, source-checked workbook transactions for sheet structure and cells.

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::unnecessary_unwrap,
    reason = "new workbook transaction code uses checked wire conversions and explicit validation"
)]

use super::{CellFormula, Reference, TransferLimits, Value};
use crate::Workbook;
use crate::package::error::{Error, Result};
use crate::raw::{Header, Limits as RawLimits, Records, Writer, kind};
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{BlobPart, PackURI, Part};
use std::collections::BTreeMap;
use std::sync::Arc;

const MAGIC: &[u8; 8] = b"LCXBWRP1";
const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// Typed resources for one newly authored cell XF.
#[derive(Debug, Clone)]
pub struct AuthoredStyle {
    /// Font interned into `BrtBeginFonts`.
    pub font: crate::styles::Font,
    /// Pattern fill interned into `BrtBeginFills`.
    pub fill: crate::styles::Fill,
    /// Border interned into `BrtBeginBorders`.
    pub border: crate::styles::Border,
    /// Optional custom number-format code, interned by semantic code equality.
    pub number_format: Option<String>,
    /// Optional compact alignment fields stored in the new cell XF.
    pub alignment: Option<crate::styles::Alignment>,
}

impl Default for AuthoredStyle {
    fn default() -> Self {
        Self {
            font: crate::styles::Font::default(),
            fill: crate::styles::Fill::default(),
            border: crate::styles::Border::default(),
            number_format: None,
            alignment: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    RenameSheet {
        sheet: usize,
        name: String,
    },
    TransferCell {
        source: Arc<[u8]>,
        source_sheet: usize,
        source_reference: Reference,
        target_sheet: usize,
        target_reference: Reference,
    },
    AddImage {
        sheet: usize,
        image: ImagePlan,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ImagePlan {
    data: Arc<[u8]>,
    format: crate::package::drawing_image::ImageFormat,
    from_col: u32,
    from_col_offset: i64,
    from_row: u32,
    from_row_offset: i64,
    to_col: u32,
    to_col_offset: i64,
    to_row: u32,
    to_row_offset: i64,
    description: Option<String>,
}

impl ImagePlan {
    fn from_image(image: &crate::writer::Image) -> Self {
        let anchor = image.anchor();
        Self {
            data: Arc::from(image.data()),
            format: image.format(),
            from_col: anchor.from_col,
            from_col_offset: anchor.from_col_offset,
            from_row: anchor.from_row,
            from_row_offset: anchor.from_row_offset,
            to_col: anchor.to_col,
            to_col_offset: anchor.to_col_offset,
            to_row: anchor.to_row,
            to_row_offset: anchor.to_row_offset,
            description: image.description().map(str::to_string),
        }
    }

    fn image(&self) -> Result<crate::writer::Image> {
        let anchor = crate::chart::Anchor::with_offsets(
            self.from_col,
            self.from_col_offset,
            self.from_row,
            self.from_row_offset,
            self.to_col,
            self.to_col_offset,
            self.to_row,
            self.to_row_offset,
        );
        let image = crate::writer::Image::new(Arc::clone(&self.data), self.format, anchor)?;
        match &self.description {
            Some(description) => image.with_description(description.clone()),
            None => Ok(image),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum Selection {
    Sheet(usize),
    Cell(usize, Reference),
    Drawing(usize),
}

impl Operation {
    const fn selection(&self) -> Selection {
        match self {
            Self::RenameSheet { sheet, .. } => Selection::Sheet(*sheet),
            Self::TransferCell {
                target_sheet,
                target_reference,
                ..
            } => Selection::Cell(*target_sheet, *target_reference),
            Self::AddImage { sheet, .. } => Selection::Drawing(*sheet),
        }
    }
}

/// A detached workbook transaction planned against one exact package image.
#[derive(Debug, Clone)]
pub struct WorkbookEdit {
    before: Arc<[u8]>,
    operations: Vec<Operation>,
    limits: TransferLimits,
}

impl WorkbookEdit {
    pub(crate) fn new(workbook: &Workbook, limits: TransferLimits) -> Result<Self> {
        validate_limits(limits)?;
        Ok(Self {
            before: Arc::from(workbook_bytes(workbook)?),
            operations: Vec::new(),
            limits,
        })
    }

    /// Plan a sheet rename without mutating the source workbook.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, invalid/duplicate name, or a
    /// transaction that exceeds its finite change policy.
    pub fn rename_sheet(&mut self, sheet: usize, name: String) -> Result<()> {
        let workbook = workbook_from_bytes(&self.before)?;
        validate_sheet_name(&workbook, sheet, &name)?;
        self.stage(Operation::RenameSheet { sheet, name })
    }

    /// Plan dependency-complete transfer of one stored cell from another XLSB.
    ///
    /// Cell XF resources, rich-string fonts, and shared-string entries are
    /// interned into the target during commit. Formula tokens remain inert and
    /// must validate against the target workbook during complete readback.
    /// Neither workbook is mutated while this operation is staged.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, absent source cell, occupied
    /// target coordinate, unsupported resource topology, or exceeded limits.
    pub fn transfer_cell(
        &mut self,
        source: &Workbook,
        source_sheet: usize,
        source_reference: Reference,
        target_sheet: usize,
        target_reference: Reference,
    ) -> Result<()> {
        let source = Arc::from(workbook_bytes(source)?);
        let operation = Operation::TransferCell {
            source,
            source_sheet,
            source_reference,
            target_sheet,
            target_reference,
        };
        let mut candidate = self.operations.clone();
        candidate.push(operation.clone());
        let _validated_candidate = replay(&self.before, &candidate)?;
        self.stage(operation)
    }

    /// Author a new SST-backed cell and its complete typed style resources.
    ///
    /// Formatting and phonetic runs are normalized to the authored font so no
    /// caller-supplied workbook index can escape into the candidate.
    pub fn insert_shared_string(
        &mut self,
        sheet: usize,
        reference: Reference,
        string: crate::package::SharedString,
        style: &AuthoredStyle,
    ) -> Result<()> {
        self.insert_authored_string(sheet, reference, string, style, true)
    }

    /// Author a new inline `BrtCellRString` and its typed style resources.
    pub fn insert_rich_string(
        &mut self,
        sheet: usize,
        reference: Reference,
        string: crate::package::SharedString,
        style: &AuthoredStyle,
    ) -> Result<()> {
        self.insert_authored_string(sheet, reference, string, style, false)
    }

    /// Author a numeric formula cell with inert tokens and typed style closure.
    ///
    /// Formula name, sheet, external-link, and table indexes are resolved
    /// against the target during detached planning and complete readback.
    pub fn insert_formula_number(
        &mut self,
        sheet: usize,
        reference: Reference,
        cache: f64,
        formula: CellFormula,
        style: &AuthoredStyle,
    ) -> Result<()> {
        self.insert_formula(
            sheet,
            reference,
            Value::FormulaNumberCache(cache),
            formula,
            style,
        )
    }

    /// Author a formula cell with any supported inert cached-result family.
    ///
    /// `cache` must be one of the four `Value::Formula*Cache` variants. The
    /// detached candidate validates the cache, formula tokens, target context,
    /// and typed style closure before the operation can be staged.
    pub fn insert_formula(
        &mut self,
        sheet: usize,
        reference: Reference,
        cache: Value,
        formula: CellFormula,
        style: &AuthoredStyle,
    ) -> Result<()> {
        if !cache.is_formula_cache() {
            return Err(Error::InvalidFormat(
                "formula authoring requires a Formula*Cache value".to_string(),
            ));
        }
        let mut donor = workbook_from_bytes(&self.before)?;
        let style_index = author_style(&mut donor, style)?;
        insert_candidate_cell(
            &mut donor,
            sheet,
            reference,
            style_index,
            cache,
            Some(formula),
        )?;
        self.transfer_cell(&donor, sheet, reference, sheet, reference)
    }

    /// Add one dependency-closed embedded image to a worksheet.
    ///
    /// The image payload, DrawingML part, relationships, and binary
    /// `BrtDrawing` link are replayed as one semantic root operation. For an
    /// existing standard drawing, the new anchor is inserted before the root
    /// close while all prior XML bytes, anchors, and relationships are retained.
    pub fn insert_image(&mut self, sheet: usize, image: crate::writer::Image) -> Result<()> {
        let operation = Operation::AddImage {
            sheet,
            image: ImagePlan::from_image(&image),
        };
        let mut candidate = self.operations.clone();
        candidate.push(operation.clone());
        let _validated_candidate = replay(&self.before, &candidate)?;
        self.stage(operation)
    }

    /// Transfer one decoded embedded-image resource into a new target anchor.
    ///
    /// The source image bytes, declared format, and alternative text are
    /// preserved. The caller supplies the target anchor because worksheet
    /// dimensions are target-owned semantic context.
    pub fn transfer_image(
        &mut self,
        source: &Workbook,
        source_sheet: usize,
        image_index: usize,
        target_sheet: usize,
        target_anchor: crate::chart::Anchor,
    ) -> Result<()> {
        let drawing = source.sheet_drawing(source_sheet).ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "source sheet {source_sheet} has no decoded drawing"
            ))
        })?;
        let embedded = drawing.images.get(image_index).ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "source sheet {source_sheet} has no embedded image {image_index}"
            ))
        })?;
        let image =
            crate::writer::Image::new(Arc::clone(&embedded.data), embedded.format, target_anchor)?;
        let image = match &embedded.description {
            Some(description) => image.with_description(description.clone())?,
            None => image,
        };
        self.insert_image(target_sheet, image)
    }

    fn insert_authored_string(
        &mut self,
        sheet: usize,
        reference: Reference,
        mut string: crate::package::SharedString,
        style: &AuthoredStyle,
        shared: bool,
    ) -> Result<()> {
        let mut donor = workbook_from_bytes(&self.before)?;
        let style_index = author_style(&mut donor, style)?;
        let font_id = donor
            .styles()
            .get_cell_format(usize::try_from(style_index.get()).map_err(|error| {
                Error::InvalidFormat(format!("authored style index overflow: {error}"))
            })?)
            .ok_or_else(|| Error::InvalidFormat("authored cell XF is absent".to_string()))?
            .font_id;
        let font_id = u16::try_from(font_id).map_err(|error| {
            Error::InvalidFormat(format!("authored font index exceeds u16: {error}"))
        })?;
        for run in &mut string.runs {
            run.font_id = font_id;
        }
        if let Some(phonetic) = &mut string.phonetic {
            phonetic.font_id = font_id;
        }
        let value = if shared {
            let mut package = donor.package.clone();
            let index = super::resources::intern_shared_string_for_new_cell(&mut package, &string)?;
            donor = Workbook::from_opc_package(package)?;
            Value::SharedStringIndex(index)
        } else {
            Value::RichString(string)
        };
        insert_candidate_cell(&mut donor, sheet, reference, style_index, value, None)?;
        self.transfer_cell(&donor, sheet, reference, sheet, reference)
    }

    /// Validate all staged operations and produce an immutable durable patch.
    ///
    /// # Errors
    ///
    /// Returns an error when replay, complete workbook readback, or bounds fail.
    pub fn commit(self) -> Result<WorkbookCommit> {
        let after = Arc::from(replay(&self.before, &self.operations)?);
        let patch = WorkbookPatch {
            before: self.before,
            after,
            operations: self.operations,
        };
        let _bounded_encoding = patch.to_bytes(self.limits)?;
        Ok(WorkbookCommit { patch })
    }

    fn stage(&mut self, operation: Operation) -> Result<()> {
        if self.operations.len() >= self.limits.changes() {
            return Err(Error::InvalidLength {
                expected: self.limits.changes(),
                found: self.operations.len().saturating_add(1),
            });
        }
        let selection = operation.selection();
        if self
            .operations
            .iter()
            .any(|existing| existing.selection() == selection)
        {
            return Err(Error::UnsupportedFeature(format!(
                "workbook transaction selects {selection:?} more than once"
            )));
        }
        self.operations.push(operation);
        Ok(())
    }
}

/// A validated workbook transaction ready for atomic publication.
#[derive(Debug, Clone)]
pub struct WorkbookCommit {
    patch: WorkbookPatch,
}

impl WorkbookCommit {
    /// Exact-source durable patch carried by this commit.
    #[must_use]
    pub const fn patch(&self) -> &WorkbookPatch {
        &self.patch
    }
}

/// An exact before/after workbook patch with semantic replay operations.
#[derive(Debug, Clone)]
pub struct WorkbookPatch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    operations: Vec<Operation>,
}

impl WorkbookPatch {
    /// Canonical exact package image required before application.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Canonical validated package image published by application.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Number of semantic workbook operations retained for durable replay.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.operations.len()
    }

    /// Apply this patch atomically after checking the exact package source.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale workbook or failed complete readback.
    pub fn apply(&self, workbook: &mut Workbook) -> Result<()> {
        if workbook_bytes(workbook)?.as_slice() != self.before.as_ref() {
            return Err(Error::UnsupportedFeature(
                "workbook patch source is stale".to_string(),
            ));
        }
        let candidate = validated_workbook(&self.after)?;
        *workbook = candidate;
        Ok(())
    }

    /// Construct the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            operations: Vec::new(),
        }
    }

    /// Perform a non-mutating semantic three-way merge on a common base.
    ///
    /// Identical overlapping operations coalesce. Divergent sheet/cell
    /// selections are reported without producing a partial patch.
    ///
    /// # Errors
    ///
    /// Returns an error for distinct exact bases or failed bounded replay.
    pub fn merge_three_way(&self, other: &Self) -> Result<WorkbookMergeOutcome> {
        if self.before != other.before {
            return Err(Error::UnsupportedFeature(
                "workbook patches do not share an exact base".to_string(),
            ));
        }
        if (self.operations.is_empty() && self.before != self.after)
            || (other.operations.is_empty() && other.before != other.after)
        {
            return Err(Error::UnsupportedFeature(
                "exact-image inverse patches do not carry forward semantic merge operations"
                    .to_string(),
            ));
        }
        let mut merged = self.operations.clone();
        let mut positions = merged
            .iter()
            .enumerate()
            .map(|(index, operation)| (operation.selection(), index))
            .collect::<BTreeMap<_, _>>();
        let mut conflicts = Vec::new();
        for right in &other.operations {
            if let Some(index) = positions.get(&right.selection()).copied() {
                if merged[index] != *right {
                    conflicts.push(WorkbookMergeConflict {
                        selection: right.selection(),
                    });
                }
            } else {
                positions.insert(right.selection(), merged.len());
                merged.push(right.clone());
            }
        }
        if !conflicts.is_empty() {
            return Ok(WorkbookMergeOutcome {
                patch: None,
                conflicts,
            });
        }
        let after = Arc::from(replay(&self.before, &merged)?);
        Ok(WorkbookMergeOutcome {
            patch: Some(Self {
                before: Arc::clone(&self.before),
                after,
                operations: merged,
            }),
            conflicts,
        })
    }

    /// Encode a deterministic, bounded, versioned workbook patch.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch exceeds the finite transfer policy.
    pub fn to_bytes(&self, limits: TransferLimits) -> Result<Vec<u8>> {
        validate_limits(limits)?;
        if self.operations.len() > limits.changes() {
            return Err(Error::InvalidLength {
                expected: limits.changes(),
                found: self.operations.len(),
            });
        }
        let mut encoded_operations = Vec::new();
        for operation in &self.operations {
            encode_operation(operation, &mut encoded_operations)?;
        }
        let total = MAGIC
            .len()
            .checked_add(24)
            .and_then(|size| size.checked_add(self.before.len()))
            .and_then(|size| size.checked_add(self.after.len()))
            .and_then(|size| size.checked_add(encoded_operations.len()))
            .ok_or(Error::CapacityOverflow {
                resource: "durable workbook patch bytes",
            })?;
        if total > limits.bytes() {
            return Err(Error::InvalidLength {
                expected: limits.bytes(),
                found: total,
            });
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(total)
            .map_err(|source| Error::Allocation {
                resource: "durable workbook patch bytes",
                source,
            })?;
        output.extend_from_slice(MAGIC);
        push_usize(&mut output, self.before.len())?;
        push_usize(&mut output, self.after.len())?;
        push_usize(&mut output, self.operations.len())?;
        output.extend_from_slice(&self.before);
        output.extend_from_slice(&self.after);
        output.extend_from_slice(&encoded_operations);
        Ok(output)
    }

    /// Decode, replay, and fully validate a durable workbook patch.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed input, exceeded limits, invalid package
    /// images, or forward semantic operations inconsistent with the stored
    /// result. Exact-image inverse patches intentionally carry no forward
    /// operations and validate both package endpoints instead.
    pub fn from_bytes(data: &[u8], limits: TransferLimits) -> Result<Self> {
        validate_limits(limits)?;
        if data.len() > limits.bytes() {
            return Err(Error::InvalidLength {
                expected: limits.bytes(),
                found: data.len(),
            });
        }
        require(data, 32)?;
        if data.get(..8) != Some(MAGIC.as_slice()) {
            return Err(Error::InvalidFormat(
                "unknown durable workbook patch version".to_string(),
            ));
        }
        let before_len = read_usize(data, 8)?;
        let after_len = read_usize(data, 16)?;
        let operation_count = read_usize(data, 24)?;
        if operation_count > limits.changes() {
            return Err(Error::InvalidLength {
                expected: limits.changes(),
                found: operation_count,
            });
        }
        let before_end = 32usize
            .checked_add(before_len)
            .ok_or(Error::CapacityOverflow {
                resource: "workbook patch before image",
            })?;
        let after_end = before_end
            .checked_add(after_len)
            .ok_or(Error::CapacityOverflow {
                resource: "workbook patch after image",
            })?;
        require(data, after_end)?;
        let mut offset = after_end;
        let mut operations = Vec::new();
        operations
            .try_reserve_exact(operation_count)
            .map_err(|source| Error::Allocation {
                resource: "workbook patch operations",
                source,
            })?;
        for _ in 0..operation_count {
            operations.push(decode_operation(data, &mut offset)?);
        }
        if offset != data.len() {
            return Err(Error::InvalidLength {
                expected: offset,
                found: data.len(),
            });
        }
        let before = Arc::from(data[32..before_end].to_vec());
        let after = Arc::from(data[before_end..after_end].to_vec());
        let _validated_before = validated_workbook(&before)?;
        let _validated_after = validated_workbook(&after)?;
        if !operations.is_empty() && replay(&before, &operations)?.as_slice() != after.as_ref() {
            return Err(Error::InvalidFormat(
                "durable workbook patch operations do not reconstruct its after image".to_string(),
            ));
        }
        Ok(Self {
            before,
            after,
            operations,
        })
    }

    fn retained_bytes(&self) -> Result<usize> {
        let mut bytes =
            self.before
                .len()
                .checked_add(self.after.len())
                .ok_or(Error::CapacityOverflow {
                    resource: "workbook history bytes",
                })?;
        for operation in &self.operations {
            let operation_bytes = match operation {
                Operation::RenameSheet { name, .. } => 17usize.checked_add(name.len()),
                Operation::TransferCell { source, .. } => 41usize.checked_add(source.len()),
                Operation::AddImage { image, .. } => {
                    74usize.checked_add(image.data.len()).and_then(|size| {
                        size.checked_add(image.description.as_ref().map_or(0, String::len))
                    })
                },
            }
            .ok_or(Error::CapacityOverflow {
                resource: "workbook history operation bytes",
            })?;
            bytes = bytes
                .checked_add(operation_bytes)
                .ok_or(Error::CapacityOverflow {
                    resource: "workbook history bytes",
                })?;
        }
        Ok(bytes)
    }
}

/// One divergent structural/cell selection in a workbook merge.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkbookMergeConflict {
    selection: Selection,
}

impl WorkbookMergeConflict {
    /// Sheet index selected by both divergent operations.
    #[must_use]
    pub const fn sheet_index(&self) -> usize {
        match self.selection {
            Selection::Sheet(sheet) | Selection::Cell(sheet, _) | Selection::Drawing(sheet) => {
                sheet
            },
        }
    }

    /// Conflicting cell coordinate, or `None` for a sheet-level conflict.
    #[must_use]
    pub const fn cell_reference(&self) -> Option<Reference> {
        match self.selection {
            Selection::Sheet(_) | Selection::Drawing(_) => None,
            Selection::Cell(_, reference) => Some(reference),
        }
    }
}

/// Atomic result of a workbook-level three-way merge.
#[derive(Debug, Clone)]
pub struct WorkbookMergeOutcome {
    patch: Option<WorkbookPatch>,
    conflicts: Vec<WorkbookMergeConflict>,
}

impl WorkbookMergeOutcome {
    /// Merged patch, present exactly when there are no conflicts.
    #[must_use]
    pub const fn patch(&self) -> Option<&WorkbookPatch> {
        self.patch.as_ref()
    }

    /// All divergent selections.
    #[must_use]
    pub fn conflicts(&self) -> &[WorkbookMergeConflict] {
        &self.conflicts
    }
}

/// Bounded exact-source workbook undo/redo history.
#[derive(Debug, Clone)]
pub struct WorkbookHistory {
    entries: Vec<WorkbookPatch>,
    cursor: usize,
    retained_bytes: usize,
    limits: TransferLimits,
}

impl WorkbookHistory {
    /// Construct an empty bounded workbook history.
    pub fn new(limits: TransferLimits) -> Result<Self> {
        validate_limits(limits)?;
        Ok(Self {
            entries: Vec::new(),
            cursor: 0,
            retained_bytes: 0,
            limits,
        })
    }

    /// Number of retained undo/redo entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether no entries are retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Retain one contiguous committed patch and evict the oldest entries.
    pub fn push(&mut self, patch: WorkbookPatch) -> Result<()> {
        let bytes = patch.retained_bytes()?;
        if bytes > self.limits.history_bytes() {
            return Err(Error::InvalidLength {
                expected: self.limits.history_bytes(),
                found: bytes,
            });
        }
        let expected = if self.cursor == 0 {
            self.entries.first().map(|entry| Arc::clone(&entry.before))
        } else {
            self.entries
                .get(self.cursor.saturating_sub(1))
                .map(|entry| Arc::clone(&entry.after))
        };
        if expected.is_some_and(|source| source.as_ref() != patch.before.as_ref()) {
            return Err(Error::UnsupportedFeature(
                "workbook history patch does not continue its exact source tip".to_string(),
            ));
        }
        while self.entries.len() > self.cursor {
            if let Some(removed) = self.entries.pop() {
                self.retained_bytes = self
                    .retained_bytes
                    .saturating_sub(removed.retained_bytes()?);
            }
        }
        self.retained_bytes =
            self.retained_bytes
                .checked_add(bytes)
                .ok_or(Error::CapacityOverflow {
                    resource: "workbook history bytes",
                })?;
        self.entries.push(patch);
        self.cursor = self.entries.len();
        while self.entries.len() > self.limits.history_entries()
            || self.retained_bytes > self.limits.history_bytes()
        {
            let removed = self.entries.remove(0);
            self.retained_bytes = self
                .retained_bytes
                .saturating_sub(removed.retained_bytes()?);
            self.cursor = self.cursor.saturating_sub(1);
        }
        Ok(())
    }

    /// Apply one exact inverse at the current history tip.
    pub fn undo(&mut self, workbook: &mut Workbook) -> Result<()> {
        let index = self.cursor.checked_sub(1).ok_or_else(|| {
            Error::UnsupportedFeature("workbook history has no undo entry".to_string())
        })?;
        self.entries[index].inverse().apply(workbook)?;
        self.cursor = index;
        Ok(())
    }

    /// Reapply one entry on the redo side.
    pub fn redo(&mut self, workbook: &mut Workbook) -> Result<()> {
        let patch = self.entries.get(self.cursor).ok_or_else(|| {
            Error::UnsupportedFeature("workbook history has no redo entry".to_string())
        })?;
        patch.apply(workbook)?;
        self.cursor = self.cursor.saturating_add(1);
        Ok(())
    }
}

fn replay(before: &[u8], operations: &[Operation]) -> Result<Vec<u8>> {
    let mut workbook = validated_workbook(before)?;
    for operation in operations {
        match operation {
            Operation::RenameSheet { sheet, name } => rename_sheet(&mut workbook, *sheet, name)?,
            Operation::TransferCell {
                source,
                source_sheet,
                source_reference,
                target_sheet,
                target_reference,
            } => transfer_cell(
                &mut workbook,
                source,
                *source_sheet,
                *source_reference,
                *target_sheet,
                *target_reference,
            )?,
            Operation::AddImage { sheet, image } => add_image(&mut workbook, *sheet, image)?,
        }
    }
    validate_all_worksheets(&workbook)?;
    workbook_bytes(&workbook)
}

fn author_style(workbook: &mut Workbook, style: &AuthoredStyle) -> Result<super::StyleIndex> {
    let plan = super::resources::plan_style(style)?;
    let mut package = workbook.package.clone();
    let index = super::resources::intern_style_plan(&mut package, &plan)?;
    package.unsign();
    *workbook = Workbook::from_opc_package(package)?;
    Ok(index)
}

fn insert_candidate_cell(
    workbook: &mut Workbook,
    sheet: usize,
    reference: Reference,
    style: super::StyleIndex,
    value: Value,
    formula: Option<CellFormula>,
) -> Result<()> {
    let uri = workbook.worksheet_uri(sheet)?;
    let mut package = workbook.package.clone();
    let mut edit = super::workbook::read(&package, &uri)?.edit();
    if let Some(formula) = formula {
        edit.insert_formula(reference, style, value, formula)?;
    } else {
        edit.insert(reference, style, value)?;
    }
    let commit = edit.commit()?;
    let _snapshot = super::workbook::apply(&mut package, &uri, &commit)?;
    *workbook = Workbook::from_opc_package(package)?;
    Ok(())
}

fn add_image(workbook: &mut Workbook, sheet: usize, plan: &ImagePlan) -> Result<()> {
    let worksheet_uri = workbook.worksheet_uri(sheet)?;
    let worksheet_source = workbook.package.get_part(&worksheet_uri)?.blob().to_vec();
    let mut drawing_rel_id = None;
    for item in Records::new(&worksheet_source) {
        let record = item?;
        if record.kind() == kind::DRAWING {
            if drawing_rel_id.is_some() {
                return Err(Error::UnsupportedFeature(
                    "worksheet contains multiple BrtDrawing records".to_string(),
                ));
            }
            let mut cursor = crate::raw::Cursor::new(record.payload(), "BrtDrawing");
            drawing_rel_id = Some(cursor.read_wide_string()?);
        }
    }
    let image = plan.image()?;
    if let Some(drawing_rel_id) = drawing_rel_id {
        return append_image(
            workbook,
            sheet,
            &worksheet_uri,
            &drawing_rel_id,
            plan,
            &image,
        );
    }
    if workbook
        .package
        .get_part(&worksheet_uri)?
        .rels()
        .iter()
        .any(|relationship| {
            matches!(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::DRAWING
                    | litchi_opc::constants::relationship_type::STRICT_DRAWING
            )
        })
    {
        return Err(Error::UnsupportedFeature(
            "worksheet has a Drawing relationship without BrtDrawing ownership".to_string(),
        ));
    }
    let search_end = u32::try_from(workbook.package.iter_parts().count().saturating_add(1))
        .map_err(|error| {
            Error::InvalidFormat(format!("drawing part search bound exceeds u32: {error}"))
        })?;
    let index = (1_u32..=search_end)
        .find(|index| {
            let drawing = PackURI::new(format!("/xl/drawings/drawing{index}.xml"));
            let media = PackURI::new(format!(
                "/xl/media/image{index}.{}",
                plan.format.extension()
            ));
            drawing.is_ok_and(|uri| !workbook.package.contains_part(&uri))
                && media.is_ok_and(|uri| !workbook.package.contains_part(&uri))
        })
        .ok_or_else(|| Error::UnsupportedFeature("no drawing part index remains".to_string()))?;
    let drawing_uri = PackURI::new(format!("/xl/drawings/drawing{index}.xml"))?;
    let media_uri = PackURI::new(format!(
        "/xl/media/image{index}.{}",
        plan.format.extension()
    ))?;
    let drawing_xml = crate::package::drawing_write::serialize_drawing(
        std::slice::from_ref(&image),
        &[],
        &[],
        &[],
        &[],
    )?;
    let mut package = workbook.package.clone();
    package.try_add_part(Box::new(BlobPart::new(
        media_uri.clone(),
        plan.format.content_type().to_string(),
        plan.data.to_vec(),
    )))?;
    let mut drawing_part = BlobPart::new(
        drawing_uri.clone(),
        litchi_opc::constants::content_type::OFC_DRAWING.to_string(),
        drawing_xml,
    );
    let strict = is_strict(&package);
    drawing_part.rels_mut().add_relationship(
        if strict {
            litchi_opc::constants::relationship_type::STRICT_IMAGE
        } else {
            litchi_opc::constants::relationship_type::IMAGE
        }
        .to_string(),
        format!("../media/image{index}.{}", plan.format.extension()),
        "rId1".to_string(),
        false,
    );
    package.try_add_part(Box::new(drawing_part))?;
    let drawing_rel_id = package.get_part_mut(&worksheet_uri)?.relate_to(
        &format!("../drawings/drawing{index}.xml"),
        if strict {
            litchi_opc::constants::relationship_type::STRICT_DRAWING
        } else {
            litchi_opc::constants::relationship_type::DRAWING
        },
    );
    let mut drawing_payload = Vec::new();
    Writer::new(&mut drawing_payload).write_wide_string(&drawing_rel_id)?;
    let mut worksheet_output = Vec::new();
    let mut inserted = false;
    for item in Records::new(&worksheet_source) {
        let record = item?;
        if !inserted && matches!(record.kind(), kind::BEGIN_LIST_PARTS | kind::END_SHEET) {
            Writer::new(&mut worksheet_output).write_record(kind::DRAWING, &drawing_payload)?;
            inserted = true;
        }
        copy_record(&worksheet_source, &record, &mut worksheet_output)?;
    }
    if !inserted {
        return Err(Error::InvalidFormat(
            "worksheet has no safe BrtDrawing insertion boundary".to_string(),
        ));
    }
    package
        .get_part_mut(&worksheet_uri)?
        .set_blob(worksheet_output);
    package.unsign();
    *workbook = Workbook::from_opc_package(package)?;
    if workbook.sheet_drawing(sheet).is_none() {
        return Err(Error::InvalidFormat(
            "authored worksheet drawing failed semantic readback".to_string(),
        ));
    }
    Ok(())
}

fn append_image(
    workbook: &mut Workbook,
    sheet: usize,
    worksheet_uri: &PackURI,
    drawing_rel_id: &str,
    plan: &ImagePlan,
    image: &crate::writer::Image,
) -> Result<()> {
    let worksheet_part = workbook.package.get_part(worksheet_uri)?;
    let relationship = worksheet_part.rels().get(drawing_rel_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "BrtDrawing relationship {drawing_rel_id:?} is absent"
        ))
    })?;
    if relationship.is_external()
        || !matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::DRAWING
                | litchi_opc::constants::relationship_type::STRICT_DRAWING
        )
    {
        return Err(Error::InvalidFormat(
            "BrtDrawing relationship is external or has the wrong type".to_string(),
        ));
    }
    let drawing_uri = relationship.target_partname()?;
    let drawing = workbook.sheet_drawing(sheet).ok_or_else(|| {
        Error::InvalidFormat("BrtDrawing has no decoded DrawingML inventory".to_string())
    })?;
    let before_images = drawing.images.len();
    let object_id = crate::package::drawing_write::next_drawing_object_id(
        workbook.package.get_part(&drawing_uri)?.blob(),
    )?;
    let media_index = free_media_index(&workbook.package, plan.format.extension())?;
    let media_uri = PackURI::new(format!(
        "/xl/media/image{media_index}.{}",
        plan.format.extension()
    ))?;
    let strict = is_strict(&workbook.package);
    let mut package = workbook.package.clone();
    package.try_add_part(Box::new(BlobPart::new(
        media_uri,
        plan.format.content_type().to_string(),
        plan.data.to_vec(),
    )))?;
    let drawing_part = package.get_part_mut(&drawing_uri)?;
    let image_rel_id = drawing_part.relate_to(
        &format!("../media/image{media_index}.{}", plan.format.extension()),
        if strict {
            litchi_opc::constants::relationship_type::STRICT_IMAGE
        } else {
            litchi_opc::constants::relationship_type::IMAGE
        },
    );
    let drawing_xml = crate::package::drawing_write::append_image_anchor(
        drawing_part.blob(),
        image,
        object_id,
        &image_rel_id,
    )?;
    drawing_part.set_blob(drawing_xml);
    package.unsign();
    *workbook = Workbook::from_opc_package(package)?;
    let drawing = workbook.sheet_drawing(sheet).ok_or_else(|| {
        Error::InvalidFormat("appended worksheet drawing failed semantic readback".to_string())
    })?;
    let expected_images = before_images
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("drawing image count overflow".to_string()))?;
    if drawing.images.len() != expected_images
        || drawing.images.last().map(|value| value.data.as_ref()) != Some(plan.data.as_ref())
    {
        return Err(Error::InvalidFormat(
            "appended image dependency failed semantic readback".to_string(),
        ));
    }
    Ok(())
}

fn free_media_index(package: &litchi_opc::OpcPackage, extension: &str) -> Result<u32> {
    let search_end =
        u32::try_from(package.iter_parts().count().saturating_add(1)).map_err(|error| {
            Error::InvalidFormat(format!("media part search bound exceeds u32: {error}"))
        })?;
    (1..=search_end)
        .find(|index| {
            PackURI::new(format!("/xl/media/image{index}.{extension}"))
                .is_ok_and(|uri| !package.contains_part(&uri))
        })
        .ok_or_else(|| Error::UnsupportedFeature("no image part index remains".to_string()))
}

fn is_strict(package: &litchi_opc::OpcPackage) -> bool {
    package
        .iter_parts()
        .flat_map(|part| part.rels().iter())
        .any(|relationship| {
            relationship
                .reltype()
                .starts_with("http://purl.oclc.org/ooxml/")
        })
}

fn transfer_cell(
    target: &mut Workbook,
    source_bytes: &[u8],
    source_sheet: usize,
    source_reference: Reference,
    target_sheet: usize,
    target_reference: Reference,
) -> Result<()> {
    let source = validated_workbook(source_bytes)?;
    let source_snapshot = source.cell_values(source_sheet)?;
    let source_cell = source_snapshot
        .cell(source_reference)?
        .cloned()
        .ok_or_else(|| {
            Error::InvalidCellReference(format!(
                "source cell ({}, {}) is absent",
                source_reference.row(),
                source_reference.column()
            ))
        })?;
    if target
        .cell_values(target_sheet)?
        .cell(target_reference)?
        .is_some()
    {
        return Err(Error::UnsupportedFeature(format!(
            "target cell ({}, {}) is occupied",
            target_reference.row(),
            target_reference.column()
        )));
    }
    let mut package = target.package.clone();
    let style =
        super::resources::transfer_style(&source.package, &mut package, source_cell.style())?;
    let value = match source_cell.value() {
        Value::SharedStringIndex(index) => {
            let index = usize::try_from(*index).map_err(|_| {
                Error::InvalidCellReference("shared-string index does not fit usize".to_string())
            })?;
            let string = source.shared_strings().get(index).ok_or_else(|| {
                Error::InvalidCellReference(format!("source shared-string index {index} is absent"))
            })?;
            let string =
                super::resources::transfer_string_fonts(&source.package, &mut package, string)?;
            Value::SharedStringIndex(super::resources::intern_shared_string_for_new_cell(
                &mut package,
                &string,
            )?)
        },
        Value::RichString(string) => Value::RichString(super::resources::transfer_string_fonts(
            &source.package,
            &mut package,
            string,
        )?),
        Value::Blank => Value::Blank,
        Value::RkNumber(value) => Value::RkNumber(*value),
        Value::Error(value) => Value::Error(*value),
        Value::Boolean(value) => Value::Boolean(*value),
        Value::Number(value) => Value::Number(*value),
        Value::InlineString(value) => Value::InlineString(value.clone()),
        Value::FormulaStringCache(value) => Value::FormulaStringCache(value.clone()),
        Value::FormulaNumberCache(value) => Value::FormulaNumberCache(*value),
        Value::FormulaBooleanCache(value) => Value::FormulaBooleanCache(*value),
        Value::FormulaErrorCache(value) => Value::FormulaErrorCache(*value),
    };
    let uri = target.worksheet_uri(target_sheet)?;
    let mut edit = super::workbook::read(&package, &uri)?.edit();
    if let Some(formula) = source_cell.formula() {
        let formula =
            super::formula_transfer::remap(&source, target, source_sheet, target_sheet, formula)?;
        edit.insert_formula(target_reference, style, value, formula)?;
    } else {
        edit.insert(target_reference, style, value)?;
    }
    edit.set_show_phonetic(target_reference, source_cell.show_phonetic())?;
    let commit = edit.commit()?;
    let _published_snapshot = super::workbook::apply(&mut package, &uri, &commit)?;
    *target = Workbook::from_opc_package(package)?;
    Ok(())
}

fn rename_sheet(workbook: &mut Workbook, sheet: usize, name: &str) -> Result<()> {
    validate_sheet_name(workbook, sheet, name)?;
    let uri = workbook.package.main_document_part()?.partname().clone();
    let source = workbook.package.get_part(&uri)?.blob().to_vec();
    let mut output = Vec::new();
    let mut index = 0usize;
    let mut replaced = false;
    for item in Records::new(&source) {
        let record = item?;
        if record.kind() == kind::BUNDLE_SH && index == sheet {
            let payload = renamed_bundle_payload(record.payload(), name)?;
            Writer::new(&mut output).write_record(kind::BUNDLE_SH, &payload)?;
            replaced = true;
        } else {
            copy_record(&source, &record, &mut output)?;
        }
        if record.kind() == kind::BUNDLE_SH {
            index = index.saturating_add(1);
        }
    }
    if !replaced {
        return Err(Error::WorksheetNotFound(format!("sheet index {sheet}")));
    }
    let mut package = workbook.package.clone();
    package.get_part_mut(&uri)?.set_blob(output);
    package.unsign();
    *workbook = Workbook::from_opc_package(package)?;
    Ok(())
}

fn renamed_bundle_payload(source: &[u8], name: &str) -> Result<Vec<u8>> {
    let current_id = litchi_core::binary::read_u32_le_at(source, 4)?;
    let strings_offset = if (1..=0xffff).contains(&current_id) {
        8
    } else {
        12
    };
    let strings = source.get(strings_offset..).ok_or(Error::InvalidLength {
        expected: strings_offset,
        found: source.len(),
    })?;
    let consumed = if litchi_core::binary::read_u32_le_at(strings, 0)? == u32::MAX {
        4
    } else {
        crate::package::records::decode_string(strings)?.1
    };
    let name_offset = strings_offset
        .checked_add(consumed)
        .ok_or(Error::CapacityOverflow {
            resource: "bundle-sheet name offset",
        })?;
    let mut payload = source
        .get(..name_offset)
        .ok_or(Error::InvalidLength {
            expected: name_offset,
            found: source.len(),
        })?
        .to_vec();
    Writer::new(&mut payload).write_wide_string(name)?;
    let _validated_record = crate::package::records::BundleSheetRecord::parse(&payload)?;
    Ok(payload)
}

fn validate_sheet_name(workbook: &Workbook, sheet: usize, name: &str) -> Result<()> {
    if sheet >= workbook.worksheet_count() {
        return Err(Error::WorksheetNotFound(format!("sheet index {sheet}")));
    }
    let units = name.encode_utf16().count();
    if units == 0
        || units > 31
        || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return Err(Error::InvalidFormat(format!("invalid sheet name {name:?}")));
    }
    if workbook
        .worksheet_names()
        .iter()
        .enumerate()
        .any(|(index, existing)| index != sheet && existing.eq_ignore_ascii_case(name))
    {
        return Err(Error::InvalidFormat(format!(
            "duplicate sheet name {name:?}"
        )));
    }
    Ok(())
}

fn validate_all_worksheets(workbook: &Workbook) -> Result<()> {
    for index in 0..workbook.worksheet_count() {
        let uri = workbook.worksheet_uri(index)?;
        if workbook.package.get_part(&uri)?.content_type() == WORKSHEET_CONTENT_TYPE {
            let _cell_snapshot = workbook.cell_values(index)?;
            let _worksheet = workbook.worksheet(index)?;
        }
    }
    Ok(())
}

fn validated_workbook(bytes: &[u8]) -> Result<Workbook> {
    let workbook = workbook_from_bytes(bytes)?;
    validate_all_worksheets(&workbook)?;
    Ok(workbook)
}

fn workbook_from_bytes(bytes: &[u8]) -> Result<Workbook> {
    crate::Package::from_slice(bytes)?.into_workbook()
}

fn workbook_bytes(workbook: &Workbook) -> Result<Vec<u8>> {
    crate::Package::from(workbook.package.clone()).to_bytes()
}

fn encode_operation(operation: &Operation, output: &mut Vec<u8>) -> Result<()> {
    match operation {
        Operation::RenameSheet { sheet, name } => {
            output.push(1);
            push_usize(output, *sheet)?;
            push_usize(output, name.len())?;
            output.extend_from_slice(name.as_bytes());
        },
        Operation::TransferCell {
            source,
            source_sheet,
            source_reference,
            target_sheet,
            target_reference,
        } => {
            output.push(2);
            push_usize(output, source.len())?;
            push_usize(output, *source_sheet)?;
            output.extend_from_slice(&source_reference.row().to_le_bytes());
            output.extend_from_slice(&source_reference.column().to_le_bytes());
            push_usize(output, *target_sheet)?;
            output.extend_from_slice(&target_reference.row().to_le_bytes());
            output.extend_from_slice(&target_reference.column().to_le_bytes());
            output.extend_from_slice(source);
        },
        Operation::AddImage { sheet, image } => {
            output.push(3);
            push_usize(output, *sheet)?;
            output.push(image_format_bits(image.format));
            push_usize(output, image.data.len())?;
            match &image.description {
                Some(description) => push_usize(output, description.len())?,
                None => output.extend_from_slice(&u64::MAX.to_le_bytes()),
            }
            output.extend_from_slice(&image.from_col.to_le_bytes());
            output.extend_from_slice(&image.from_col_offset.to_le_bytes());
            output.extend_from_slice(&image.from_row.to_le_bytes());
            output.extend_from_slice(&image.from_row_offset.to_le_bytes());
            output.extend_from_slice(&image.to_col.to_le_bytes());
            output.extend_from_slice(&image.to_col_offset.to_le_bytes());
            output.extend_from_slice(&image.to_row.to_le_bytes());
            output.extend_from_slice(&image.to_row_offset.to_le_bytes());
            output.extend_from_slice(&image.data);
            if let Some(description) = &image.description {
                output.extend_from_slice(description.as_bytes());
            }
        },
    }
    Ok(())
}

fn decode_operation(data: &[u8], offset: &mut usize) -> Result<Operation> {
    let tag = read_slice(data, offset, 1)?[0];
    match tag {
        1 => {
            let sheet = read_next_usize(data, offset)?;
            let length = read_next_usize(data, offset)?;
            let name = std::str::from_utf8(read_slice(data, offset, length)?)
                .map_err(|error| Error::InvalidFormat(format!("sheet name is not UTF-8: {error}")))?
                .to_string();
            Ok(Operation::RenameSheet { sheet, name })
        },
        2 => {
            let source_len = read_next_usize(data, offset)?;
            let source_sheet = read_next_usize(data, offset)?;
            let source_row = read_next_u32(data, offset)?;
            let source_column = read_next_u32(data, offset)?;
            let target_sheet = read_next_usize(data, offset)?;
            let target_row = read_next_u32(data, offset)?;
            let target_column = read_next_u32(data, offset)?;
            let source = Arc::from(read_slice(data, offset, source_len)?.to_vec());
            Ok(Operation::TransferCell {
                source,
                source_sheet,
                source_reference: Reference::new(source_row, source_column)?,
                target_sheet,
                target_reference: Reference::new(target_row, target_column)?,
            })
        },
        3 => {
            let sheet = read_next_usize(data, offset)?;
            let format = image_format_from_bits(read_slice(data, offset, 1)?[0])?;
            let data_len = read_next_usize(data, offset)?;
            let description_len = read_next_u64(data, offset)?;
            let from_col = read_next_u32(data, offset)?;
            let from_col_offset = read_next_i64(data, offset)?;
            let from_row = read_next_u32(data, offset)?;
            let from_row_offset = read_next_i64(data, offset)?;
            let to_col = read_next_u32(data, offset)?;
            let to_col_offset = read_next_i64(data, offset)?;
            let to_row = read_next_u32(data, offset)?;
            let to_row_offset = read_next_i64(data, offset)?;
            let image_data = Arc::from(read_slice(data, offset, data_len)?.to_vec());
            let description = if description_len == u64::MAX {
                None
            } else {
                let length = usize::try_from(description_len).map_err(|error| {
                    Error::InvalidFormat(format!("image description length overflow: {error}"))
                })?;
                Some(
                    std::str::from_utf8(read_slice(data, offset, length)?)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("image description is not UTF-8: {error}"))
                        })?
                        .to_string(),
                )
            };
            let image = ImagePlan {
                data: image_data,
                format,
                from_col,
                from_col_offset,
                from_row,
                from_row_offset,
                to_col,
                to_col_offset,
                to_row,
                to_row_offset,
                description,
            };
            let _validated_image = image.image()?;
            Ok(Operation::AddImage { sheet, image })
        },
        _ => Err(Error::InvalidFormat(format!(
            "unknown workbook patch operation {tag}"
        ))),
    }
}

const fn image_format_bits(format: crate::package::drawing_image::ImageFormat) -> u8 {
    match format {
        crate::package::drawing_image::ImageFormat::Bmp => 0,
        crate::package::drawing_image::ImageFormat::Gif => 1,
        crate::package::drawing_image::ImageFormat::Jpeg => 2,
        crate::package::drawing_image::ImageFormat::Png => 3,
        crate::package::drawing_image::ImageFormat::Svg => 4,
        crate::package::drawing_image::ImageFormat::Tiff => 5,
        crate::package::drawing_image::ImageFormat::Emf => 6,
        crate::package::drawing_image::ImageFormat::Wmf => 7,
        crate::package::drawing_image::ImageFormat::Wdp => 8,
    }
}

fn image_format_from_bits(value: u8) -> Result<crate::package::drawing_image::ImageFormat> {
    match value {
        0 => Ok(crate::package::drawing_image::ImageFormat::Bmp),
        1 => Ok(crate::package::drawing_image::ImageFormat::Gif),
        2 => Ok(crate::package::drawing_image::ImageFormat::Jpeg),
        3 => Ok(crate::package::drawing_image::ImageFormat::Png),
        4 => Ok(crate::package::drawing_image::ImageFormat::Svg),
        5 => Ok(crate::package::drawing_image::ImageFormat::Tiff),
        6 => Ok(crate::package::drawing_image::ImageFormat::Emf),
        7 => Ok(crate::package::drawing_image::ImageFormat::Wmf),
        8 => Ok(crate::package::drawing_image::ImageFormat::Wdp),
        _ => Err(Error::InvalidFormat(format!(
            "unknown durable image format {value}"
        ))),
    }
}

fn copy_record(source: &[u8], record: &crate::raw::Record<'_>, output: &mut Vec<u8>) -> Result<()> {
    let record_source = source
        .get(record.offset()..)
        .ok_or_else(|| Error::InvalidFormat("record offset is outside workbook.bin".to_string()))?;
    let (_, header_len) = Header::parse(record_source, RawLimits::DEFAULT)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "workbook record range",
        })?;
    output.extend_from_slice(
        source.get(record.offset()..end).ok_or_else(|| {
            Error::InvalidFormat("record range is outside workbook.bin".to_string())
        })?,
    );
    Ok(())
}

fn validate_limits(limits: TransferLimits) -> Result<()> {
    if limits.bytes() == 0
        || limits.changes() == 0
        || limits.history_entries() == 0
        || limits.history_bytes() == 0
    {
        return Err(Error::InvalidFormat(
            "workbook patch transfer/history limits must be nonzero".to_string(),
        ));
    }
    Ok(())
}

fn push_usize(output: &mut Vec<u8>, value: usize) -> Result<()> {
    let value = u64::try_from(value).map_err(|_| Error::CapacityOverflow {
        resource: "workbook patch length",
    })?;
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn read_usize(data: &[u8], offset: usize) -> Result<usize> {
    let bytes = data
        .get(offset..offset.saturating_add(8))
        .ok_or(Error::InvalidLength {
            expected: offset.saturating_add(8),
            found: data.len(),
        })?;
    usize::try_from(u64::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::InvalidFormat("invalid workbook patch length".to_string())
    })?))
    .map_err(|_| Error::CapacityOverflow {
        resource: "workbook patch length",
    })
}

fn read_next_usize(data: &[u8], offset: &mut usize) -> Result<usize> {
    let value = read_usize(data, *offset)?;
    *offset = offset.saturating_add(8);
    Ok(value)
}

fn read_next_u32(data: &[u8], offset: &mut usize) -> Result<u32> {
    let bytes = read_slice(data, offset, 4)?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::InvalidFormat("invalid workbook patch coordinate".to_string())
    })?))
}

fn read_next_u64(data: &[u8], offset: &mut usize) -> Result<u64> {
    let bytes = read_slice(data, offset, 8)?;
    Ok(u64::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::InvalidFormat("invalid workbook patch u64".to_string())
    })?))
}

fn read_next_i64(data: &[u8], offset: &mut usize) -> Result<i64> {
    let bytes = read_slice(data, offset, 8)?;
    Ok(i64::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::InvalidFormat("invalid workbook patch i64".to_string())
    })?))
}

fn read_slice<'a>(data: &'a [u8], offset: &mut usize, len: usize) -> Result<&'a [u8]> {
    let end = offset.checked_add(len).ok_or(Error::CapacityOverflow {
        resource: "workbook patch range",
    })?;
    let value = data.get(*offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    *offset = end;
    Ok(value)
}

fn require(data: &[u8], expected: usize) -> Result<()> {
    if data.len() < expected {
        Err(Error::InvalidLength {
            expected,
            found: data.len(),
        })
    } else {
        Ok(())
    }
}
