//! Durable, source-checked workbook transactions for sheet structure and cells.
//!
//! This family uses module-local transaction names. Worksheet-only cell edits
//! remain in the parent module; the workbook owner supplies this family's
//! entry point through [crate::Package::edit].
//!
//! ```
//! # fn main() -> litchi_xlsb::package::PackageResult<()> {
//! use litchi_core::sheet::traits::WorkbookTrait;
//! use litchi_xlsb::Package;
//! use litchi_xlsb::cell_values::TransferLimits;
//! use litchi_xlsb::cell_values::structure::Patch;
//!
//! let source = Package::new()?;
//! let mut edit = source.edit()?;
//! edit.set_sheet_name(0, "Renamed".to_owned())?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! assert!(!commit.patch().is_noop());
//! assert_eq!(commit.package().workbook()?.worksheet_names(), &["Renamed"]);
//! let (package, patch) = commit.into_parts();
//! let wire = patch.to_bytes(TransferLimits::DEFAULT)?;
//! let decoded = Patch::from_bytes(&wire, TransferLimits::DEFAULT, source.raw_limits_value())?;
//! let applied = source.apply(&decoded)?.into_package();
//! assert_eq!(package.to_bytes()?, applied.to_bytes()?);
//! let restored = package.apply(&patch.inverse())?.into_package();
//! assert_eq!(restored.to_bytes()?, source.to_bytes()?);
//! # Ok(())
//! # }
//! ```

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
    clippy::map_err_ignore,
    clippy::unnecessary_unwrap,
    clippy::wildcard_enum_match_arm,
    reason = "new workbook transaction code uses checked wire conversions and explicit validation"
)]

pub use super::LimitKind;
use super::{CellFormula, Reference, TransferLimits, Value};
use crate::Workbook;
use crate::external_link::ExternalLinkLimits;
use crate::formula::{Parser, Token};
pub use crate::package::error::Error;
use crate::package::error::Result;
use crate::package::formula::excel_name_eq;
use crate::raw::{Encoder, Header, Limits as RawLimits, Records, kind};
use crate::sheet::{Kind, Selector, State};
use crate::workbook::DrawingLoadPolicy;
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{BlobPart, PackURI, Part};
use std::collections::{BTreeMap, HashSet};
use std::mem::size_of;
use std::sync::Arc;

const MAGIC: &[u8; 8] = b"LCXBWRP1";
const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// Finite policies for the primary workbook transaction
/// ([`crate::Package::edit_with_options`]).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    transfer_limits: TransferLimits,
}

impl EditOptions {
    /// The default transfer and patch policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            transfer_limits: TransferLimits::DEFAULT,
        }
    }

    /// Select the transfer, patch and history policy.
    #[must_use]
    pub const fn transfer_limits(mut self, limits: TransferLimits) -> Self {
        self.transfer_limits = limits;
        self
    }

    pub(crate) const fn get_transfer_limits(self) -> TransferLimits {
        self.transfer_limits
    }
}

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
    /// Legacy durable operation: `sheet` is a worksheet-only ordinal.
    ///
    /// Tag 1 is retained so patches written before the complete tab catalog
    /// was exposed remain decodable.  Newly staged edits use
    /// [`Self::SetSheetName`], which identifies the complete catalog position.
    RenameSheet {
        sheet: usize,
        name: String,
    },
    /// Canonical complete-catalog sheet-name operation.
    SetSheetName {
        catalog_position: usize,
        name: String,
        /// Whether replay must update XML dependency carriers owned by the
        /// workbook package.  Tag 6 deliberately retains the historical
        /// metadata-only behavior; newly staged edits use tag 7.
        dependencies: bool,
    },
    /// Replace the complete workbook-tab order by source-catalog positions.
    ///
    /// The vector is a permutation in final-position order.  Keeping source
    /// positions rather than names makes the operation stable when names are
    /// duplicated in a lenient source image and keeps chart/dialog/macro tabs
    /// in the same complete catalog as worksheets.
    ReorderSheets {
        order: Arc<[usize]>,
    },
    /// Author one blank worksheet at a position in the currently staged
    /// complete tab catalog.  The position is replayed sequentially, so a
    /// patch containing several insertions remains deterministic even when
    /// later insertions are staged after earlier ones.
    InsertSheet {
        catalog_position: usize,
        name: String,
        state: State,
    },
    /// Remove one or more complete-catalog tabs by their immutable source
    /// positions.  The positions stay source-relative so a durable patch is
    /// independent of the order in which selectors were staged.
    RemoveSheets {
        positions: Arc<[usize]>,
    },
    SetSheetState {
        catalog_position: usize,
        state: State,
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
    TransferDrawing {
        source: Arc<[u8]>,
        source_sheet: usize,
        source_anchor: usize,
        target_sheet: usize,
    },
    /// Replace every merged range of one worksheet, selected by its
    /// complete-catalog position. Ranges are canonical: sorted and pairwise
    /// disjoint. Durable tag 11.
    SetMergedRanges {
        catalog_position: usize,
        ranges: Arc<[crate::merged_cells::MergedCell]>,
    },
    /// Set (`target` present) or remove a defined name of one scope (a
    /// complete-catalog position, or the workbook). Durable tag 12.
    SetDefinedName {
        name: String,
        scope: Option<usize>,
        target: Option<(usize, crate::formula::Range)>,
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

#[derive(Debug, Clone)]
struct SheetTarget {
    name: String,
    kind: Kind,
    state: State,
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
        let anchor = crate::builder::chart::Anchor::with_offsets(
            self.from_col,
            self.from_col_offset,
            self.from_row,
            self.from_row_offset,
            self.to_col,
            self.to_col_offset,
            self.to_row,
            self.to_row_offset,
        );
        // The bytes were admitted when the source workbook was opened; the
        // copy is re-admitted under the format ceilings.
        let image = crate::writer::Image::new(
            Arc::clone(&self.data),
            self.format,
            anchor,
            &crate::package::drawing_image::Limits::DEFAULT,
        )?;
        match &self.description {
            Some(description) => image.with_description(description.clone()),
            None => Ok(image),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum Selection {
    SheetRename(usize),
    LegacyWorksheetRename(usize),
    SheetOrder,
    SheetInsert(usize),
    SheetRemoval,
    SheetState(usize),
    Cell(usize, Reference),
    Drawing(usize),
    Merges(usize),
    /// A defined name: its case-folded spelling's hash and its scope
    /// (`usize::MAX` for the workbook).
    DefinedName(u64, usize),
}

impl Operation {
    fn selection(&self) -> Selection {
        match self {
            Self::RenameSheet { sheet, .. } => Selection::LegacyWorksheetRename(*sheet),
            Self::SetSheetName {
                catalog_position, ..
            } => Selection::SheetRename(*catalog_position),
            Self::ReorderSheets { .. } => Selection::SheetOrder,
            Self::InsertSheet {
                catalog_position, ..
            } => Selection::SheetInsert(*catalog_position),
            Self::RemoveSheets { .. } => Selection::SheetRemoval,
            Self::SetSheetState {
                catalog_position, ..
            } => Selection::SheetState(*catalog_position),
            Self::TransferCell {
                target_sheet,
                target_reference,
                ..
            } => Selection::Cell(*target_sheet, *target_reference),
            Self::AddImage { sheet, .. } => Selection::Drawing(*sheet),
            Self::TransferDrawing { target_sheet, .. } => Selection::Drawing(*target_sheet),
            Self::SetMergedRanges {
                catalog_position, ..
            } => Selection::Merges(*catalog_position),
            Self::SetDefinedName { name, scope, .. } => {
                Selection::DefinedName(name_key(name), scope.unwrap_or(usize::MAX))
            },
        }
    }

    fn conflicts_with(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::ReorderSheets { .. }, _)
            | (_, Self::ReorderSheets { .. })
            | (Self::InsertSheet { .. }, _)
            | (_, Self::InsertSheet { .. })
            | (Self::RemoveSheets { .. }, _)
            | (_, Self::RemoveSheets { .. }) => true,
            (Self::RenameSheet { .. }, Self::SetSheetState { .. })
            | (Self::SetSheetState { .. }, Self::RenameSheet { .. }) => false,
            (Self::SetSheetName { .. }, Self::SetSheetState { .. })
            | (Self::SetSheetState { .. }, Self::SetSheetName { .. }) => false,
            _ => self.selection() == other.selection(),
        }
    }
}

/// A detached workbook transaction planned against one exact package image.
#[derive(Debug, Clone)]
pub struct Edit {
    base: crate::Package,
    before: Option<Arc<Vec<u8>>>,
    sheet_catalog: Option<Arc<[SheetTarget]>>,
    operations: Vec<Operation>,
    operation_indexes: BTreeMap<Selection, Vec<usize>>,
    limits: TransferLimits,
    external_link_limits: ExternalLinkLimits,
    drawing_load_policy: DrawingLoadPolicy,
}

impl Edit {
    pub(crate) fn new(package: &crate::Package, limits: TransferLimits) -> Result<Self> {
        validate_limits(limits)?;
        Ok(Self {
            base: package.clone(),
            before: None,
            sheet_catalog: None,
            operations: Vec::new(),
            operation_indexes: BTreeMap::new(),
            limits,
            external_link_limits: package.external_link_limits(),
            drawing_load_policy: package.drawing_load_policy(),
        })
    }

    /// Materialize the source catalog and exact package image at the first
    /// operation that needs them. Entry and discard therefore retain only the
    /// immutable package owner and finite edit policy.
    fn ensure_prepared(&mut self) -> Result<()> {
        let sheet_catalog: Option<Arc<[SheetTarget]>> = if self.sheet_catalog.is_none() {
            let workbook = self.base.workbook()?;
            Some(
                workbook
                    .sheets()
                    .map(|tab| SheetTarget {
                        name: tab.name().to_owned(),
                        kind: tab.kind(),
                        state: tab.state(),
                    })
                    .collect::<Vec<_>>()
                    .into(),
            )
        } else {
            None
        };
        let before = if self.before.is_none() {
            Some(Arc::new(self.base.to_bytes()?))
        } else {
            None
        };
        if let Some(sheet_catalog) = sheet_catalog {
            self.sheet_catalog = Some(sheet_catalog);
        }
        if let Some(before) = before {
            self.before = Some(before);
        }
        Ok(())
    }

    fn catalog(&self) -> Result<&[SheetTarget]> {
        self.sheet_catalog.as_deref().ok_or(Error::InvalidFormat(
            "workbook edit catalog is not prepared",
        ))
    }

    fn before_image(&self) -> Result<&Arc<Vec<u8>>> {
        self.before
            .as_ref()
            .ok_or(Error::InvalidFormat("workbook edit source is not prepared"))
    }

    /// Plan a workbook-tab name change without mutating the source workbook.
    ///
    /// The selector resolves against this edit's immutable complete tab
    /// catalog.  A tab renamed earlier in the same edit is therefore still
    /// selected by its source name (or its source catalog position), and its
    /// new name is not itself a selector.  Chart, dialog, macro, and unbound
    /// tabs are addressable because their names are owned by `BrtBundleSh` in
    /// the workbook part just like worksheet names.
    ///
    /// The return value reports whether the selected tab's current staged name
    /// differs from `name`.  A resolved no-op is retained as one operation so
    /// it has the same exact-image and durable-patch behavior as the other
    /// metadata operations.  A missing selector returns `Ok(None)` without
    /// staging.  The final staged names are checked as one workbook-wide
    /// namespace using Excel's Unicode-caseless identity rule.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous selector, an invalid or duplicate
    /// name, or a transaction that exceeds its finite change policy.
    pub fn set_sheet_name<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        name: String,
    ) -> Result<Option<bool>> {
        self.ensure_prepared()?;
        let selector = selector.into();
        let catalog = self.catalog()?;
        let Some(catalog_position) = resolve_catalog_position(catalog, &selector)? else {
            return Ok(None);
        };
        let name_index = self
            .operation_indexes
            .get(&Selection::SheetRename(catalog_position))
            .and_then(|indexes| {
                indexes.iter().copied().find(|index| {
                    matches!(&self.operations[*index], Operation::SetSheetName { .. })
                })
            });
        let current = name_index.map_or_else(
            || catalog[catalog_position].name.as_str(),
            |index| match &self.operations[index] {
                Operation::SetSheetName { name, .. } => name.as_str(),
                Operation::RenameSheet { .. }
                | Operation::SetSheetState { .. }
                | Operation::ReorderSheets { .. }
                | Operation::InsertSheet { .. }
                | Operation::RemoveSheets { .. }
                | Operation::TransferCell { .. }
                | Operation::AddImage { .. }
                | Operation::TransferDrawing { .. }
                | Operation::SetMergedRanges { .. }
                | Operation::SetDefinedName { .. } => catalog[catalog_position].name.as_str(),
            },
        );
        let changed = current != name;
        if changed {
            validate_sheet_name_value(&name)?;
            validate_staged_sheet_name(
                self.catalog()?,
                &self.operations,
                &self.operation_indexes,
                catalog_position,
                &name,
            )?;
        }
        let operation = Operation::SetSheetName {
            catalog_position,
            name,
            dependencies: true,
        };
        if let Some(index) = name_index {
            self.operations[index] = operation;
        } else {
            self.stage(operation)?;
        }
        Ok(Some(changed))
    }

    /// Replace the complete workbook-tab order by semantic selectors.
    ///
    /// Every selector is resolved against this edit's immutable source
    /// catalog, and the slice must name every tab exactly once.  The order
    /// includes chart, dialog, macro, and unbound tabs as well as worksheets.
    /// Repeating the operation replaces the earlier staged permutation, so a
    /// later identity order cancels the change without retaining stale
    /// intermediate positions.
    ///
    /// Reordering is a workbook-structure barrier.  It cannot be composed in
    /// one transaction with metadata, cell, or drawing operations until those
    /// operations have a shared catalog-identity remapping contract.
    ///
    /// ```
    /// use litchi_core::sheet::traits::WorkbookTrait;
    /// use litchi_xlsb::{Builder, sheet::Selector, builder};
    ///
    /// let mut builder = Builder::new();
    /// builder.add_worksheet(builder::Worksheet::new("First"));
    /// builder.add_worksheet(builder::Worksheet::new("Second"));
    /// let source = builder.build()?;
    /// let mut edit = source.edit()?;
    /// assert!(edit.reorder_sheets(&[
    ///     Selector::from("Second"),
    ///     Selector::from("First"),
    /// ])?);
    /// let commit = edit.commit()?;
    /// assert_eq!(
    ///     commit.package().workbook()?.worksheet_names(),
    ///     &["Second", "First"],
    /// );
    /// # Ok::<(), litchi_xlsb::package::PackageError>(())
    /// ```
    pub fn reorder_sheets(&mut self, selectors: &[Selector<'_>]) -> Result<bool> {
        self.ensure_prepared()?;
        if self
            .operations
            .iter()
            .any(|operation| !matches!(operation, Operation::ReorderSheets { .. }))
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reordering cannot be combined with other workbook operations",
            ));
        }
        if selectors.len() != self.catalog()?.len() {
            return Err(Error::InvalidFormat(
                "sheet order must select every catalog tab exactly once",
            ));
        }
        let mut order = Vec::new();
        order
            .try_reserve_exact(selectors.len())
            .map_err(|source| Error::Allocation {
                resource: "XLSB sheet order selectors",
                source,
            })?;
        let mut seen = Vec::new();
        seen.try_reserve_exact(self.catalog()?.len())
            .map_err(|source| Error::Allocation {
                resource: "XLSB sheet order selector set",
                source,
            })?;
        seen.resize(self.catalog()?.len(), false);
        for selector in selectors {
            let catalog_position = resolve_catalog_position(self.catalog()?, selector)?
                .ok_or(Error::WorksheetNotFound(None))?;
            let used = seen.get_mut(catalog_position).ok_or(Error::InvalidFormat(
                "sheet order position is out of bounds",
            ))?;
            if std::mem::replace(used, true) {
                return Err(Error::InvalidFormat("sheet order selectors must be unique"));
            }
            order.push(catalog_position);
        }
        if seen.iter().any(|used| !used) {
            return Err(Error::InvalidFormat(
                "sheet order selectors must cover the complete catalog",
            ));
        }
        let order: Arc<[usize]> = order.into();
        let existing = self
            .operation_indexes
            .get(&Selection::SheetOrder)
            .and_then(|indexes| {
                indexes
                    .iter()
                    .rev()
                    .find_map(|index| match &self.operations[*index] {
                        Operation::ReorderSheets { order } => Some(order),
                        Operation::RenameSheet { .. }
                        | Operation::SetSheetName { .. }
                        | Operation::InsertSheet { .. }
                        | Operation::RemoveSheets { .. }
                        | Operation::SetSheetState { .. }
                        | Operation::TransferCell { .. }
                        | Operation::AddImage { .. }
                        | Operation::TransferDrawing { .. }
                        | Operation::SetMergedRanges { .. }
                        | Operation::SetDefinedName { .. } => None,
                    })
            });
        let identity = order
            .iter()
            .enumerate()
            .all(|(position, source)| position == *source);
        let changed = existing.map_or(!identity, |current| current.as_ref() != order.as_ref());
        let operation = Operation::ReorderSheets { order };
        if let Some(index) =
            self.operation_indexes
                .get(&Selection::SheetOrder)
                .and_then(|indexes| {
                    indexes.iter().rev().copied().find(|index| {
                        matches!(self.operations[*index], Operation::ReorderSheets { .. })
                    })
                })
        {
            self.operations[index] = operation;
        } else {
            self.stage(operation)?;
        }
        Ok(changed)
    }

    /// Author one blank worksheet at a checked complete-catalog position.
    ///
    /// The position is interpreted against the catalog produced by earlier
    /// insertions in this edit, so `Position::new(current_len)` appends and a
    /// repeated insertion at the same position is well-defined.  A newly
    /// authored worksheet has no cells, drawings, or formulas.  Its worksheet
    /// part has the required binary-index relationship, and its explicit
    /// visibility state is carried by the workbook's `BrtBundleSh` record.
    /// Existing workbook operations cannot be mixed with insertion because
    /// their sheet ordinals have no shared remapping contract.
    ///
    /// The durable operation is tag 9.  Its payload is the one-byte tag,
    /// followed by an eight-byte little-endian catalog position, one state
    /// byte, an eight-byte little-endian UTF-8 name length, and the name
    /// bytes.  The operation is source-checked and replayed in order, so a
    /// patch containing several insertions has the same staged-catalog
    /// semantics as this method.
    ///
    /// ```
    /// use litchi_core::sheet::traits::WorkbookTrait;
    /// use litchi_core::Position;
    /// use litchi_xlsb::{Builder, sheet::State, builder};
    ///
    /// let mut builder = Builder::new();
    /// builder.add_worksheet(builder::Worksheet::new("First"));
    /// builder.add_worksheet(builder::Worksheet::new("Second"));
    /// let source = builder.build()?;
    /// let mut edit = source.edit()?;
    /// edit.insert_sheet(Position::new(1), "Inserted".to_owned(), State::Visible)?;
    /// let commit = edit.commit()?;
    /// assert_eq!(
    ///     commit.package().workbook()?.worksheet_names(),
    ///     &["First", "Inserted", "Second"],
    /// );
    /// # Ok::<(), litchi_xlsb::package::PackageError>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range staged position, an invalid or
    /// duplicate name, a structural dependency that cannot be remapped, or
    /// a transaction that exceeds its finite change and patch policy.
    pub fn insert_sheet(
        &mut self,
        position: litchi_core::Position,
        name: String,
        state: State,
    ) -> Result<()> {
        self.ensure_prepared()?;
        if self
            .operations
            .iter()
            .any(|operation| !matches!(operation, Operation::InsertSheet { .. }))
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet insertion cannot be combined with other workbook operations",
            ));
        }
        if self.operations.len() >= self.limits.changes() {
            return Err(Error::Limit {
                kind: LimitKind::Changes.into(),
                observed: self.operations.len().saturating_add(1),
                maximum: self.limits.changes(),
            });
        }
        let catalog_position = position.get();
        let staged_insertions = self
            .operations
            .iter()
            .filter(|operation| matches!(operation, Operation::InsertSheet { .. }))
            .count();
        let staged_len = self.catalog()?.len().checked_add(staged_insertions).ok_or(
            Error::CapacityOverflow {
                resource: "staged XLSB sheet catalog",
            },
        )?;
        if catalog_position > staged_len {
            return Err(Error::InvalidFormat(
                "sheet insertion position exceeds staged catalog length",
            ));
        }
        validate_sheet_name_value(&name)?;
        validate_staged_insert_sheet_name(self.catalog()?, &self.operations, &name)?;
        let operation = Operation::InsertSheet {
            catalog_position,
            name,
            state,
        };
        let mut operation_bytes = 0usize;
        for staged in &self.operations {
            operation_bytes = operation_bytes
                .checked_add(encoded_operation_length(staged)?)
                .ok_or(Error::CapacityOverflow {
                    resource: "durable workbook patch operation bytes",
                })?;
        }
        operation_bytes = operation_bytes
            .checked_add(encoded_operation_length(&operation)?)
            .ok_or(Error::CapacityOverflow {
                resource: "durable workbook patch operation bytes",
            })?;
        let before_len = self.before_image()?.len();
        let minimum_patch_bytes = MAGIC
            .len()
            .checked_add(24)
            .and_then(|size| size.checked_add(before_len))
            .and_then(|size| size.checked_add(operation_bytes))
            .ok_or(Error::CapacityOverflow {
                resource: "durable workbook patch bytes",
            })?;
        if minimum_patch_bytes > self.limits.bytes() {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes.into(),
                observed: minimum_patch_bytes,
                maximum: self.limits.bytes(),
            });
        }
        let candidate_len =
            self.operations
                .len()
                .checked_add(1)
                .ok_or(Error::CapacityOverflow {
                    resource: "staged workbook structural operations",
                })?;
        let mut candidate = Vec::new();
        candidate
            .try_reserve_exact(candidate_len)
            .map_err(|source| Error::Allocation {
                resource: "staged workbook structural operations",
                source,
            })?;
        candidate.extend(self.operations.iter().cloned());
        candidate.push(operation.clone());
        let _validated_candidate = replay_package(&self.base, self.before_image()?, &candidate)?;
        self.stage(operation)
    }

    /// Plan removal of one complete-catalog tab.
    ///
    /// The selector is resolved against the immutable source catalog. Repeating
    /// a selector in one edit is an exact resolved no-op; selecting another tab
    /// extends the one durable removal operation, keeping all positions in
    /// source-catalog coordinates. Worksheet, chart, dialog, macro, and
    /// very-hidden unbound tabs are all addressable.
    ///
    /// Removal is a workbook-topology barrier. The candidate is fully replayed
    /// before this edit adopts it, so relationship closure, catalog-index
    /// dependencies, and the final visible-tab invariant fail atomically.
    ///
    /// ```
    /// use litchi_core::sheet::WorkbookTrait as _;
    /// use litchi_xlsb::{Builder, builder};
    ///
    /// let mut builder = Builder::new();
    /// builder.add_worksheet(builder::Worksheet::new("First"));
    /// builder.add_worksheet(builder::Worksheet::new("Second"));
    /// let source = builder.build()?;
    /// let mut edit = source.edit()?;
    /// assert_eq!(edit.remove_sheet(0_usize)?, Some(true));
    /// let commit = edit.commit()?;
    /// assert_eq!(
    ///     commit.package().workbook()?.worksheet_names(),
    ///     &["Second".to_owned()],
    /// );
    /// # Ok::<(), litchi_xlsb::package::PackageError>(())
    /// ```
    ///
    /// Repeating a resolved selector returns `Some(false)` without changing
    /// the staged operation. A selector that is absent from the immutable
    /// source catalog returns `None`.
    pub fn remove_sheet<'a>(&mut self, selector: impl Into<Selector<'a>>) -> Result<Option<bool>> {
        self.ensure_prepared()?;
        let selector = selector.into();
        let Some(position) = resolve_catalog_position(self.catalog()?, &selector)? else {
            return Ok(None);
        };
        if self
            .operations
            .iter()
            .any(|operation| !matches!(operation, Operation::RemoveSheets { .. }))
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet removal cannot be combined with other workbook operations",
            ));
        }
        let existing_index =
            self.operation_indexes
                .get(&Selection::SheetRemoval)
                .and_then(|indexes| {
                    indexes.iter().rev().copied().find(|index| {
                        matches!(self.operations[*index], Operation::RemoveSheets { .. })
                    })
                });
        let existing_positions = existing_index.and_then(|index| match &self.operations[index] {
            Operation::RemoveSheets { positions } => Some(positions.as_ref()),
            Operation::RenameSheet { .. }
            | Operation::SetSheetName { .. }
            | Operation::ReorderSheets { .. }
            | Operation::InsertSheet { .. }
            | Operation::SetSheetState { .. }
            | Operation::TransferCell { .. }
            | Operation::AddImage { .. }
            | Operation::TransferDrawing { .. }
            | Operation::SetMergedRanges { .. }
            | Operation::SetDefinedName { .. } => None,
        });
        let positions_len = existing_positions.map_or(Ok(1usize), |positions| {
            positions
                .len()
                .checked_add(1)
                .ok_or(Error::CapacityOverflow {
                    resource: "staged sheet-removal positions",
                })
        })?;
        let mut positions = Vec::new();
        positions
            .try_reserve_exact(positions_len)
            .map_err(|source| Error::Allocation {
                resource: "staged sheet-removal positions",
                source,
            })?;
        if let Some(existing_positions) = existing_positions {
            positions.extend_from_slice(existing_positions);
        }
        if positions.binary_search(&position).is_ok() {
            return Ok(Some(false));
        }
        positions.push(position);
        positions.sort_unstable();
        let positions: Arc<[usize]> = positions.into();
        let operation = Operation::RemoveSheets { positions };
        if existing_index.is_none() && self.operations.len() >= self.limits.changes() {
            return Err(Error::Limit {
                kind: LimitKind::Changes.into(),
                observed: self.operations.len().saturating_add(1),
                maximum: self.limits.changes(),
            });
        }
        let mut operation_bytes = 0usize;
        for staged in &self.operations {
            operation_bytes = operation_bytes
                .checked_add(encoded_operation_length(staged)?)
                .ok_or(Error::CapacityOverflow {
                    resource: "durable workbook patch operation bytes",
                })?;
        }
        if let Some(index) = existing_index {
            operation_bytes = operation_bytes
                .checked_sub(encoded_operation_length(&self.operations[index])?)
                .ok_or(Error::CapacityOverflow {
                    resource: "durable workbook patch operation bytes",
                })?;
        }
        operation_bytes = operation_bytes
            .checked_add(encoded_operation_length(&operation)?)
            .ok_or(Error::CapacityOverflow {
                resource: "durable workbook patch operation bytes",
            })?;
        let before_len = self.before_image()?.len();
        let minimum_patch_bytes = MAGIC
            .len()
            .checked_add(24)
            .and_then(|size| size.checked_add(before_len))
            .and_then(|size| size.checked_add(operation_bytes))
            .ok_or(Error::CapacityOverflow {
                resource: "durable workbook patch bytes",
            })?;
        if minimum_patch_bytes > self.limits.bytes() {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes.into(),
                observed: minimum_patch_bytes,
                maximum: self.limits.bytes(),
            });
        }
        let candidate_len = self
            .operations
            .len()
            .checked_add(usize::from(existing_index.is_none()))
            .ok_or(Error::CapacityOverflow {
                resource: "staged workbook structural operations",
            })?;
        let mut candidate = Vec::new();
        candidate
            .try_reserve_exact(candidate_len)
            .map_err(|source| Error::Allocation {
                resource: "staged workbook structural operations",
                source,
            })?;
        candidate.extend(self.operations.iter().cloned());
        if let Some(index) = existing_index {
            candidate[index] = operation.clone();
        } else {
            candidate.push(operation.clone());
        }
        let _validated_candidate = replay_package(&self.base, self.before_image()?, &candidate)?;
        if let Some(index) = existing_index {
            self.operations[index] = operation;
            Ok(Some(true))
        } else {
            self.stage(operation)?;
            Ok(Some(true))
        }
    }

    /// Plan a workbook-tab visibility change using the complete sheet
    /// catalog.  Positions are zero-based [`litchi_core::Position`] values
    /// wrapped by [`Selector`], so chart, dialog, macro, and unbound tabs are
    /// addressable alongside worksheets.
    ///
    /// The return value reports whether the selected tab's current staged
    /// state differs from `state`.  A resolved no-op is still retained as one
    /// operation so its commit has the usual exact-image and durable-patch
    /// behavior.  A missing selector returns `Ok(None)` without staging;
    /// `Ok(Some(false))` is a resolved no-op.
    ///
    /// The final candidate must retain at least one visible tab.  This check
    /// is performed at commit time, allowing a transaction to hide one tab
    /// and reveal another before publication.  An unbound tab can only be
    /// very hidden because it has no physical sheet part.
    ///
    /// ```
    /// use litchi_xlsb::{Builder, sheet::State, builder};
    /// let mut builder = Builder::new();
    /// builder.add_worksheet(builder::Worksheet::new("Visible"));
    /// builder.add_worksheet(builder::Worksheet::new("Archive"));
    /// let source = builder.build()?;
    /// let mut edit = source.edit()?;
    /// assert_eq!(edit.set_sheet_state("Archive", State::Hidden)?, Some(true));
    /// assert_eq!(edit.set_sheet_state("Missing", State::Hidden)?, None);
    /// let commit = edit.commit()?;
    /// assert!(commit.diagnostics().changed());
    /// let restored = commit.package().apply(&commit.patch().inverse())?.into_package();
    /// assert_eq!(restored.to_bytes()?, source.to_bytes()?);
    /// # Ok::<(), litchi_xlsb::package::PackageError>(())
    /// ```
    pub fn set_sheet_state<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        state: State,
    ) -> Result<Option<bool>> {
        self.ensure_prepared()?;
        let selector = selector.into();
        let Some(catalog_position) = resolve_catalog_position(self.catalog()?, &selector)? else {
            return Ok(None);
        };
        let tab = &self.catalog()?[catalog_position];
        if tab.kind == Kind::Unbound && state != State::VeryHidden {
            return Err(Error::InvalidRelationship(
                "unbound XLSB sheet tabs must be very hidden",
            ));
        }
        let state_index = self
            .operation_indexes
            .get(&Selection::SheetState(catalog_position))
            .and_then(|indexes| {
                indexes.iter().copied().find(|index| {
                    matches!(&self.operations[*index], Operation::SetSheetState { .. })
                })
            });
        let current = state_index.map_or(tab.state, |index| match &self.operations[index] {
            Operation::SetSheetState { state, .. } => *state,
            Operation::RenameSheet { .. }
            | Operation::SetSheetName { .. }
            | Operation::ReorderSheets { .. }
            | Operation::InsertSheet { .. }
            | Operation::RemoveSheets { .. }
            | Operation::TransferCell { .. }
            | Operation::AddImage { .. }
            | Operation::TransferDrawing { .. }
            | Operation::SetMergedRanges { .. }
            | Operation::SetDefinedName { .. } => tab.state,
        });
        let changed = current != state;
        if let Some(index) = state_index {
            self.operations[index] = Operation::SetSheetState {
                catalog_position,
                state,
            };
        } else {
            self.stage(Operation::SetSheetState {
                catalog_position,
                state,
            })?;
        }
        Ok(Some(changed))
    }

    /// Merge exactly `range` on the selected worksheet.
    ///
    /// The selector resolves the complete tab catalog; a missing selector is
    /// `Ok(None)`. `Some(false)` reports that the range is already merged.
    /// Cell records are not changed: the covered cells keep their values, as
    /// the binary format stores them, and Excel displays the top-left one.
    /// Repeated merge edits of one worksheet replace each other's staged
    /// range set; publication rewrites only the worksheet's `BrtMergeCell`
    /// block (durable tag 11).
    ///
    /// ```
    /// use litchi_xlsb::merged_cells::MergedCell;
    /// let source = litchi_xlsb::Package::new()?;
    /// let mut edit = source.edit()?;
    /// assert_eq!(edit.merge_cells(0, MergedCell::new(0, 1, 0, 2))?, Some(true));
    /// assert_eq!(edit.merge_cells(0, MergedCell::new(0, 1, 0, 2))?, Some(false));
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(
    ///     package.workbook()?.merged_cell_ranges(0)?,
    ///     vec![MergedCell::new(0, 1, 0, 2)]
    /// );
    /// # Ok::<(), litchi_xlsb::package::PackageError>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a non-worksheet tab, an invalid or single-cell
    /// range, a range overlapping another merge, and the topology and limit
    /// refusals of [`Self::commit`].
    pub fn merge_cells<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        range: crate::merged_cells::MergedCell,
    ) -> Result<Option<bool>> {
        self.ensure_prepared()?;
        self.change_merges(selector.into(), range, true)
    }

    /// Remove exactly the merge `range` from the selected worksheet.
    ///
    /// `Some(false)` reports that no merge covers the range. A range that
    /// partially overlaps a different merge is refused, so a caller cannot
    /// remove a merge it did not name.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::merge_cells`].
    pub fn unmerge_cells<'a>(
        &mut self,
        selector: impl Into<Selector<'a>>,
        range: crate::merged_cells::MergedCell,
    ) -> Result<Option<bool>> {
        self.ensure_prepared()?;
        self.change_merges(selector.into(), range, false)
    }

    fn change_merges(
        &mut self,
        selector: Selector<'_>,
        range: crate::merged_cells::MergedCell,
        merge: bool,
    ) -> Result<Option<bool>> {
        self.ensure_prepared()?;
        let Some(catalog_position) = resolve_catalog_position(self.catalog()?, &selector)? else {
            return Ok(None);
        };
        if self.catalog()?[catalog_position].kind != Kind::Worksheet {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "edit",
                "merged ranges belong to worksheets",
            ));
        }
        range.validate()?;
        let staged = self
            .operation_indexes
            .get(&Selection::Merges(catalog_position))
            .and_then(|indexes| indexes.first().copied());
        let current: Vec<crate::merged_cells::MergedCell> = match staged {
            Some(index) => {
                let Operation::SetMergedRanges { ranges, .. } = &self.operations[index] else {
                    return Err(Error::InvalidFormat(
                        "merge selection names another operation",
                    ));
                };
                ranges.to_vec()
            },
            None => {
                let workbook = self.base.workbook()?;
                let worksheet = workbook
                    .worksheet_for_catalog_position(catalog_position)
                    .ok_or_else(|| {
                        Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedContent,
                            "edit",
                            "the selected tab has no worksheet part",
                        )
                    })?;
                workbook.merged_cell_ranges(worksheet)?
            },
        };
        let exact = current.iter().position(|existing| *existing == range);
        let mut next = current.clone();
        match (merge, exact) {
            (true, Some(_)) | (false, None) => {
                if !merge && current.iter().any(|existing| existing.overlaps(&range)) {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedTarget,
                        "edit",
                        "the range overlaps a different merge; name that merge exactly",
                    ));
                }
                return Ok(Some(false));
            },
            (true, None) => {
                if current.iter().any(|existing| existing.overlaps(&range)) {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::InvalidValue,
                        "edit",
                        "the range overlaps an existing merge",
                    ));
                }
                next.push(range);
            },
            (false, Some(index)) => {
                next.remove(index);
            },
        }
        let ranges: Arc<[crate::merged_cells::MergedCell]> = canonical_merge_ranges(next)?.into();
        let operation = Operation::SetMergedRanges {
            catalog_position,
            ranges,
        };
        match staged {
            Some(index) => self.operations[index] = operation,
            None => self.stage(operation)?,
        }
        Ok(Some(true))
    }

    /// Set a defined name that refers to a range of one worksheet.
    ///
    /// `scope` selects the sheet a local name belongs to; `None` is workbook
    /// scope. An existing name of the same scope (compared as Excel compares
    /// names) keeps its flags and comment and receives the new reference; a
    /// new name is appended after the workbook's existing names, so no name
    /// reference in a formula is renumbered. Returns `Ok(None)` when a
    /// selector does not resolve. Repeated edits of one name replace each
    /// other's staged reference (durable tag 12).
    ///
    /// ```
    /// use litchi_xlsb::formula::Range;
    /// let source = litchi_xlsb::Package::new()?;
    /// let mut edit = source.edit()?;
    /// edit.set_defined_name("Totals", None, 0, Range::new(0, 4, 1, 1)?)?;
    /// let package = edit.commit()?.into_package();
    /// let names = package.workbook()?.defined_name_entries()?;
    /// assert_eq!(names[0].name, "Totals");
    /// assert_eq!(names[0].formula.as_deref(), Some("Sheet1!$B$1:$B$5"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid name, a non-worksheet target, an
    /// extern-sheet table that cannot take another entry, and the topology
    /// and limit refusals of [`Self::commit`].
    pub fn set_defined_name<'a, 'b>(
        &mut self,
        name: &str,
        scope: Option<Selector<'a>>,
        sheet: impl Into<Selector<'b>>,
        range: crate::formula::Range,
    ) -> Result<Option<()>> {
        self.ensure_prepared()?;
        crate::named_ranges::validate_name(name)?;
        let scope = match scope {
            Some(selector) => match resolve_catalog_position(self.catalog()?, &selector)? {
                Some(position) => Some(position),
                None => return Ok(None),
            },
            None => None,
        };
        let Some(sheet) = resolve_catalog_position(self.catalog()?, &sheet.into())? else {
            return Ok(None);
        };
        if self.catalog()?[sheet].kind != Kind::Worksheet {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "a defined name refers to a worksheet range",
            ));
        }
        self.stage_defined_name(Operation::SetDefinedName {
            name: name.to_owned(),
            scope,
            target: Some((sheet, range)),
        })?;
        Ok(Some(()))
    }

    /// Remove a defined name of one scope (`None` is workbook scope).
    ///
    /// Only the workbook's last defined name can be removed: removing an
    /// earlier one would renumber the name references of every formula.
    /// Removing a name that does not exist is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid name, a name that is not the last
    /// one, and the refusals of [`Self::commit`], which include a formula
    /// that still refers to the removed name.
    pub fn remove_defined_name<'a>(
        &mut self,
        name: &str,
        scope: Option<Selector<'a>>,
    ) -> Result<Option<()>> {
        self.ensure_prepared()?;
        crate::named_ranges::validate_name(name)?;
        let scope = match scope {
            Some(selector) => match resolve_catalog_position(self.catalog()?, &selector)? {
                Some(position) => Some(position),
                None => return Ok(None),
            },
            None => None,
        };
        self.stage_defined_name(Operation::SetDefinedName {
            name: name.to_owned(),
            scope,
            target: None,
        })?;
        Ok(Some(()))
    }

    fn stage_defined_name(&mut self, operation: Operation) -> Result<()> {
        self.ensure_prepared()?;
        let selection = operation.selection();
        let staged = self
            .operation_indexes
            .get(&selection)
            .and_then(|indexes| indexes.first().copied());
        let mut candidate = self.operations.clone();
        match staged {
            Some(index) => candidate[index] = operation.clone(),
            None => candidate.push(operation.clone()),
        }
        let _validated_candidate = replay_package(&self.base, self.before_image()?, &candidate)?;
        match staged {
            Some(index) => {
                self.operations[index] = operation;
                Ok(())
            },
            None => self.stage(operation),
        }
    }

    /// Plan dependency-complete transfer of one stored cell from another XLSB.
    ///
    /// Cell XF resources, rich-string fonts, and shared-string entries are
    /// interned into the target during commit. Formula tokens remain inert and
    /// must validate against the target workbook during complete readback.
    /// Donor drawing parts are outside this dependency closure and remain
    /// opaque; the donor's drawing projection is never required for transfer.
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
        self.ensure_prepared()?;
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
        let _validated_candidate = replay_package(&self.base, self.before_image()?, &candidate)?;
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
                "formula authoring requires a Formula*Cache value",
            ));
        }
        let mut donor = self.base.materialize_workbook()?;
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
        self.ensure_prepared()?;
        let operation = Operation::AddImage {
            sheet,
            image: ImagePlan::from_image(&image),
        };
        let mut candidate = self.operations.clone();
        candidate.push(operation.clone());
        let _validated_candidate = replay_package(&self.base, self.before_image()?, &candidate)?;
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
        target_anchor: crate::builder::chart::Anchor,
    ) -> Result<()> {
        let source_catalog_position = source.catalog_position_for_worksheet(source_sheet)?;
        let drawing = source
            .sheet_drawing(source_catalog_position)
            .ok_or_else(|| {
                Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "source sheet has no decoded drawing",
                )
            })?;
        let embedded = drawing.images.get(image_index).ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "source sheet has no embedded image",
            )
        })?;
        let image = crate::writer::Image::new(
            Arc::clone(&embedded.data),
            embedded.format,
            target_anchor,
            &crate::package::drawing_image::Limits::DEFAULT,
        )?;
        let image = match &embedded.description {
            Some(description) => image.with_description(description.clone())?,
            None => image,
        };
        self.insert_image(target_sheet, image)
    }

    /// Transfer one ordinary chart, shape, group, or connector graph.
    ///
    /// Shape/group bytes and chart resource bytes are retained. Drawing object
    /// IDs, colliding object names, relationship IDs, and copied part names are
    /// remapped deterministically against the target. Selecting a connector
    /// also transfers every shape/group anchor referenced by its endpoints.
    ///
    /// # Errors
    ///
    /// Returns a typed [`super::DrawingTransferRefusal`] for pictures (use
    /// [`Self::transfer_image`]), foreign/OLE objects, shape relationships
    /// other than inert external hyperlinks, ambiguous object IDs, unresolved
    /// connector endpoints, and chart resources that escape the bounded
    /// chart-owned package graph.
    pub fn transfer_drawing_object(
        &mut self,
        source: &Workbook,
        source_sheet: usize,
        source_anchor: usize,
        target_sheet: usize,
    ) -> Result<()> {
        self.ensure_prepared()?;
        let source = Arc::from(workbook_bytes(source)?);
        let operation = Operation::TransferDrawing {
            source,
            source_sheet,
            source_anchor,
            target_sheet,
        };
        let mut candidate = self.operations.clone();
        candidate.push(operation.clone());
        let _validated_candidate = replay_package(&self.base, self.before_image()?, &candidate)?;
        self.stage(operation)
    }

    fn insert_authored_string(
        &mut self,
        sheet: usize,
        reference: Reference,
        mut string: crate::package::SharedString,
        style: &AuthoredStyle,
        shared: bool,
    ) -> Result<()> {
        self.ensure_prepared()?;
        let mut donor = self.base.materialize_workbook()?;
        let style_index = author_style(&mut donor, style)?;
        let font_id = donor
            .styles()
            .get_cell_format(
                usize::try_from(style_index.get())
                    .map_err(|_error| Error::InvalidFormat("authored style index overflow"))?,
            )
            .ok_or(Error::InvalidFormat("authored cell XF is absent"))?
            .font_id;
        let font_id = u16::try_from(font_id)
            .map_err(|_error| Error::InvalidFormat("authored font index exceeds u16"))?;
        for run in &mut string.runs {
            run.font_id = font_id;
        }
        if let Some(phonetic) = &mut string.phonetic {
            phonetic.font_id = font_id;
        }
        let value = if shared {
            let mut package = donor.package.clone();
            let index = super::resources::intern_shared_string_for_new_cell(
                &mut package,
                &string,
                donor.raw_limits_value(),
            )?;
            donor = donor.reparse_candidate(package)?;
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
    pub fn commit(self) -> Result<Commit> {
        let mut this = self;
        this.ensure_prepared()?;
        let before = this
            .before
            .take()
            .ok_or(Error::InvalidFormat("workbook edit source is not prepared"))?;
        let (after, package) = replay_package(&this.base, &before, &this.operations)?;
        let patch = Patch {
            before,
            after,
            operations: this.operations,
            external_link_limits: this.external_link_limits,
            drawing_load_policy: this.drawing_load_policy,
            raw_limits: this.base.raw_limits_value(),
        };
        let _bounded_encoding = patch.to_bytes(this.limits)?;
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(package, patch, diagnostics))
    }

    fn stage(&mut self, operation: Operation) -> Result<()> {
        let is_catalog_topology = matches!(
            operation,
            Operation::ReorderSheets { .. }
                | Operation::InsertSheet { .. }
                | Operation::RemoveSheets { .. }
        );
        if self.operations.iter().any(|existing| {
            matches!(
                existing,
                Operation::ReorderSheets { .. }
                    | Operation::InsertSheet { .. }
                    | Operation::RemoveSheets { .. }
            ) != is_catalog_topology
        }) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet catalog topology operations cannot be combined with other workbook operations",
            ));
        }
        if self.operations.len() >= self.limits.changes() {
            return Err(Error::Limit {
                kind: LimitKind::Changes.into(),
                observed: self.operations.len().saturating_add(1),
                maximum: self.limits.changes(),
            });
        }
        let selection = operation.selection();
        let conflicts = self
            .operation_indexes
            .get(&selection)
            .is_some_and(|indexes| {
                indexes.iter().any(|index| {
                    // Multiple insertions are intentionally sequential;
                    // all other topology/selection conflicts retain the
                    // ordinary merge identity rule.
                    !(matches!(self.operations[*index], Operation::InsertSheet { .. })
                        && matches!(operation, Operation::InsertSheet { .. }))
                        && !(matches!(self.operations[*index], Operation::RemoveSheets { .. })
                            && matches!(operation, Operation::RemoveSheets { .. }))
                        && self.operations[*index].conflicts_with(&operation)
                })
            });
        if conflicts {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "workbook transaction selects more than once",
            ));
        }
        let index = self.operations.len();
        self.operations.push(operation);
        self.operation_indexes
            .entry(selection)
            .or_default()
            .push(index);
        Ok(())
    }
}

/// A validated workbook transaction ready for atomic publication.
pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

pub use litchi_core::edit::Diagnostics;

/// An exact before/after workbook patch with semantic replay operations.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    operations: Vec<Operation>,
    external_link_limits: ExternalLinkLimits,
    drawing_load_policy: DrawingLoadPolicy,
    raw_limits: RawLimits,
}

impl Patch {
    fn changed_candidate(&self, package: &crate::Package) -> Result<Workbook> {
        crate::signature_policy::ensure_unsigned(package.opc_package())?;
        let limits = package.opc_package().read_limits();
        // An owned image publishes its exact bytes, so an applied (or
        // inverted) patch reproduces its after image byte for byte.
        // The workbook projection below re-admits every BIFF12 record under
        // the recipient's retained raw policy. A patch's decode profile
        // describes how its wire image was read; it is not a reason to reject
        // an equal-content image whose actual resources fit this owner. Keep
        // the owned OPC path here: it preserves the exact source/provenance
        // publication bytes used by the original package commit path.
        let reader = admit_archive_before_copy(self.after.as_slice(), limits)?;
        let candidate = litchi_opc::OpcPackage::from_shared_vec_with_admitted_reader(
            Arc::clone(&self.after),
            reader,
        )?;
        let workbook = package.project_candidate(candidate)?;
        validate_all_worksheets(&workbook)?;
        Ok(workbook)
    }

    pub(crate) fn apply_to_package(&self, package: &crate::Package) -> Result<crate::Package> {
        if package.to_bytes()?.as_slice() != self.before.as_slice() {
            return Err(Error::PatchConflict { part: "workbook" });
        }
        if self.is_noop() {
            return Ok(package.clone());
        }
        let workbook = self.changed_candidate(package)?;
        Ok(workbook.into_package())
    }

    /// Whether the canonical before and after package images are identical.
    ///
    /// This compares images, not the number of retained forward operations.
    /// It may inspect all image bytes and does not serialize or reparse them.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Canonical exact package image required before application.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before.as_slice()
    }

    /// Canonical validated package image published by application.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after.as_slice()
    }

    /// Number of semantic workbook operations retained for durable replay.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.operations.len()
    }

    /// Apply this patch atomically to a workbook projection after checking
    /// the exact package source (packages replay with
    /// [`crate::Package::apply`]).
    #[cfg(test)]
    pub(crate) fn apply(&self, workbook: &mut Workbook) -> Result<()> {
        let package = crate::Package::from_workbook(workbook);
        if workbook_bytes(workbook)?.as_slice() != self.before.as_slice() {
            return Err(Error::PatchConflict { part: "workbook" });
        }
        if self.is_noop() {
            return Ok(());
        }
        let candidate = self.changed_candidate(&package)?;
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
            external_link_limits: self.external_link_limits,
            drawing_load_policy: self.drawing_load_policy,
            raw_limits: self.raw_limits,
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
    pub fn merge_three_way(&self, other: &Self) -> Result<MergeOutcome> {
        if self.before != other.before {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "workbook patches do not share an exact base",
            ));
        }
        if (self.operations.is_empty() && self.before != self.after)
            || (other.operations.is_empty() && other.before != other.after)
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "exact-image inverse patches do not carry forward semantic merge operations",
            ));
        }
        let mut merged = self.operations.clone();
        let mut conflicts = Vec::new();
        for right in &other.operations {
            if merged.iter().any(|operation| operation == right) {
                continue;
            }
            if merged
                .iter()
                .any(|operation| operation.conflicts_with(right))
            {
                conflicts.push(MergeConflict {
                    selection: right.selection(),
                });
            } else {
                merged.push(right.clone());
            }
        }
        if !conflicts.is_empty() {
            return Ok(MergeOutcome {
                patch: None,
                conflicts,
            });
        }
        let after = Arc::new(replay(
            self.before.as_slice(),
            &merged,
            self.external_link_limits,
            self.drawing_load_policy,
            self.raw_limits,
        )?);
        Ok(MergeOutcome {
            patch: Some(Self {
                before: Arc::clone(&self.before),
                after,
                operations: merged,
                external_link_limits: self.external_link_limits,
                drawing_load_policy: self.drawing_load_policy,
                raw_limits: self.raw_limits,
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
            return Err(Error::Limit {
                kind: LimitKind::Changes.into(),
                observed: self.operations.len(),
                maximum: limits.changes(),
            });
        }
        let mut encoded_length = 0usize;
        for operation in &self.operations {
            encoded_length = encoded_length
                .checked_add(encoded_operation_length(operation)?)
                .ok_or(Error::CapacityOverflow {
                    resource: "durable workbook patch operation bytes",
                })?;
        }
        let total = MAGIC
            .len()
            .checked_add(24)
            .and_then(|size| size.checked_add(self.before.len()))
            .and_then(|size| size.checked_add(self.after.len()))
            .and_then(|size| size.checked_add(encoded_length))
            .ok_or(Error::CapacityOverflow {
                resource: "durable workbook patch bytes",
            })?;
        if total > limits.bytes() {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes.into(),
                observed: total,
                maximum: limits.bytes(),
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
        for operation in &self.operations {
            encode_operation(operation, &mut output, self.raw_limits)?;
        }
        Ok(output)
    }

    /// Decode, replay, and fully validate a durable workbook patch.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed input, exceeded limits, invalid package
    /// images, or forward semantic operations inconsistent with the stored
    /// result. This compatibility entry point validates package images with
    /// eager drawing decoding; use [`Self::from_bytes_without_drawing_parse`]
    /// for an explicit drawing-skipped projection. Exact-image inverse patches
    /// intentionally carry no forward operations and validate both package
    /// endpoints instead. `raw_limits` is required because it bounds every
    /// BIFF12 record and string decoded while validating the images.
    pub fn from_bytes(data: &[u8], limits: TransferLimits, raw_limits: RawLimits) -> Result<Self> {
        Self::from_bytes_with_external_link_limits(
            data,
            limits,
            ExternalLinkLimits::default(),
            raw_limits,
        )
    }

    /// Decode a durable workbook patch with an explicit external-link policy.
    ///
    /// The caller profile governs validation of both package images and
    /// semantic replay; no default-policy parse is performed first. The
    /// required `raw_limits` profile likewise governs every BIFF12 record and
    /// string decoded during that validation. Drawing parts are decoded
    /// eagerly; use
    /// [`Self::from_bytes_without_drawing_parse_with_external_link_limits`]
    /// when the explicit skipped projection is required.
    pub fn from_bytes_with_external_link_limits(
        data: &[u8],
        limits: TransferLimits,
        external_link_limits: ExternalLinkLimits,
        raw_limits: RawLimits,
    ) -> Result<Self> {
        Self::from_bytes_with_external_link_limits_and_drawing_policy(
            data,
            limits,
            external_link_limits,
            DrawingLoadPolicy::Eager,
            raw_limits,
        )
    }

    /// Decode a durable workbook patch while treating SpreadsheetDrawing
    /// parts as opaque package content.
    ///
    /// This is an explicit opt-in counterpart to [`Self::from_bytes`]. The
    /// durable wire format has no drawing-policy field, so callers choose the
    /// projection at decode time; ordinary `from_bytes*` calls remain eager.
    /// Drawing authoring and transfer operations still fail closed when the
    /// decoded patch is replayed under the skipped policy. `raw_limits` is
    /// required for the BIFF12 record and string validation performed here.
    pub fn from_bytes_without_drawing_parse(
        data: &[u8],
        limits: TransferLimits,
        raw_limits: RawLimits,
    ) -> Result<Self> {
        Self::from_bytes_without_drawing_parse_with_external_link_limits(
            data,
            limits,
            ExternalLinkLimits::default(),
            raw_limits,
        )
    }

    /// Decode a durable workbook patch with explicit external-link limits and
    /// the drawing-skipped projection.
    pub fn from_bytes_without_drawing_parse_with_external_link_limits(
        data: &[u8],
        limits: TransferLimits,
        external_link_limits: ExternalLinkLimits,
        raw_limits: RawLimits,
    ) -> Result<Self> {
        Self::from_bytes_with_external_link_limits_and_drawing_policy(
            data,
            limits,
            external_link_limits,
            DrawingLoadPolicy::Skipped,
            raw_limits,
        )
    }

    fn from_bytes_with_external_link_limits_and_drawing_policy(
        data: &[u8],
        limits: TransferLimits,
        external_link_limits: ExternalLinkLimits,
        drawing_load_policy: DrawingLoadPolicy,
        raw_limits: RawLimits,
    ) -> Result<Self> {
        validate_limits(limits)?;
        raw_limits.validate()?;
        if data.len() > limits.bytes() {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes.into(),
                observed: data.len(),
                maximum: limits.bytes(),
            });
        }
        require(data, 32)?;
        if data.get(..8) != Some(MAGIC.as_slice()) {
            return Err(Error::InvalidFormat(
                "unknown durable workbook patch version",
            ));
        }
        let before_len = read_usize(data, 8)?;
        let after_len = read_usize(data, 16)?;
        let operation_count = read_usize(data, 24)?;
        if operation_count > limits.changes() {
            return Err(Error::Limit {
                kind: LimitKind::Changes.into(),
                observed: operation_count,
                maximum: limits.changes(),
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
            operations.push(decode_operation(data, &mut offset, raw_limits)?);
        }
        if offset != data.len() {
            return Err(Error::InvalidLength {
                expected: offset,
                found: data.len(),
            });
        }
        let before = copy_vec_arc(&data[32..before_end], "workbook patch before image")?;
        let after = copy_vec_arc(&data[before_end..after_end], "workbook patch after image")?;
        let validated_before = validated_workbook_with_external_link_limits_and_drawing_policy(
            &before,
            external_link_limits,
            drawing_load_policy,
            raw_limits,
        )?;
        // Tag 1 predates the complete tab catalog and stores a worksheet-only
        // ordinal.  Resolve it against the immutable before image before any
        // replay or merge identity is computed.  This keeps an old patch's
        // physical target stable when chart/dialog/macro tabs precede it and
        // makes old and new encodings select the same merge key.  Re-encoding
        // the decoded patch uses the canonical tag 6 form.
        operations = canonicalize_legacy_operations(&validated_before, operations)?;
        validate_reorder_operations(&operations)?;
        let _validated_after = validated_workbook_with_external_link_limits_and_drawing_policy(
            &after,
            external_link_limits,
            drawing_load_policy,
            raw_limits,
        )?;
        if !operations.is_empty()
            && replay(
                &before,
                &operations,
                external_link_limits,
                drawing_load_policy,
                raw_limits,
            )?
            .as_slice()
                != after.as_slice()
        {
            return Err(Error::InvalidFormat(
                "durable workbook patch operations do not reconstruct its after image",
            ));
        }
        Ok(Self {
            before,
            after,
            operations,
            external_link_limits,
            drawing_load_policy,
            raw_limits,
        })
    }

    /// Return the checked history byte charge of this patch.
    ///
    /// The result includes both canonical package images and the encoded
    /// semantic operations and durable envelope, with checked arithmetic matching [`History`]'s
    /// eviction accounting. It does not serialize or reparse the package.
    ///
    /// ```
    /// let package = litchi_xlsb::Package::new()?;
    /// let commit = package.edit()?.commit()?;
    /// assert!(commit.patch().retained_bytes()? >= commit.patch().before().len());
    /// # Ok::<(), litchi_xlsb::package::PackageError>(())
    /// ```
    pub fn retained_bytes(&self) -> Result<usize> {
        let mut bytes = self
            .before
            .len()
            .checked_add(self.after.len())
            .and_then(|bytes| bytes.checked_add(MAGIC.len() + 24))
            .ok_or(Error::CapacityOverflow {
                resource: "workbook history bytes",
            })?;
        for operation in &self.operations {
            let operation_bytes = match operation {
                Operation::RenameSheet { name, .. } | Operation::SetSheetName { name, .. } => {
                    17usize.checked_add(name.len())
                },
                Operation::ReorderSheets { order } => {
                    9usize.checked_add(order.len().checked_mul(size_of::<u64>()).ok_or(
                        Error::CapacityOverflow {
                            resource: "workbook history reorder operation bytes",
                        },
                    )?)
                },
                Operation::InsertSheet { name, .. } => 18usize.checked_add(name.len()),
                Operation::RemoveSheets { positions } => {
                    9usize.checked_add(positions.len().checked_mul(size_of::<u64>()).ok_or(
                        Error::CapacityOverflow {
                            resource: "workbook history sheet-removal operation bytes",
                        },
                    )?)
                },
                Operation::SetSheetState { .. } => Some(17),
                Operation::TransferCell { source, .. } => 41usize.checked_add(source.len()),
                Operation::AddImage { image, .. } => {
                    74usize.checked_add(image.data.len()).and_then(|size| {
                        size.checked_add(image.description.as_ref().map_or(0, String::len))
                    })
                },
                Operation::TransferDrawing { source, .. } => 33usize.checked_add(source.len()),
                Operation::SetMergedRanges { ranges, .. } => merge_ranges_bytes(ranges.len()),
                Operation::SetDefinedName { name, .. } => defined_name_bytes(name),
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
pub struct MergeConflict {
    selection: Selection,
}

impl MergeConflict {
    /// Sheet index selected by both divergent operations.
    #[must_use]
    pub const fn sheet_index(&self) -> usize {
        match self.selection {
            Selection::SheetRename(sheet)
            | Selection::LegacyWorksheetRename(sheet)
            | Selection::SheetInsert(sheet)
            | Selection::SheetState(sheet)
            | Selection::Cell(sheet, _)
            | Selection::Drawing(sheet)
            | Selection::Merges(sheet) => sheet,
            Selection::DefinedName(_, scope) => scope,
            Selection::SheetOrder | Selection::SheetRemoval => usize::MAX,
        }
    }

    /// Conflicting cell coordinate, or `None` for a sheet-level conflict.
    #[must_use]
    pub const fn cell_reference(&self) -> Option<Reference> {
        match self.selection {
            Selection::SheetRename(_)
            | Selection::LegacyWorksheetRename(_)
            | Selection::SheetInsert(_)
            | Selection::SheetState(_)
            | Selection::SheetOrder
            | Selection::SheetRemoval
            | Selection::Drawing(_)
            | Selection::Merges(_)
            | Selection::DefinedName(..) => None,
            Selection::Cell(_, reference) => Some(reference),
        }
    }
}

/// Atomic result of a workbook-level three-way merge.
#[derive(Debug, Clone)]
pub struct MergeOutcome {
    patch: Option<Patch>,
    conflicts: Vec<MergeConflict>,
}

impl MergeOutcome {
    /// Merged patch, present exactly when there are no conflicts.
    #[must_use]
    pub const fn patch(&self) -> Option<&Patch> {
        self.patch.as_ref()
    }

    /// All divergent selections.
    #[must_use]
    pub fn conflicts(&self) -> &[MergeConflict] {
        &self.conflicts
    }
}

/// Bounded exact-source workbook undo/redo history.
#[derive(Debug, Clone)]
pub struct History {
    entries: Vec<Patch>,
    cursor: usize,
    retained_bytes: usize,
    limits: TransferLimits,
}

impl History {
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
    pub fn push(&mut self, patch: Patch) -> Result<()> {
        let bytes = patch.retained_bytes()?;
        if bytes > self.limits.history_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::HistoryBytes.into(),
                observed: bytes,
                maximum: self.limits.history_bytes(),
            });
        }
        let expected = if self.cursor == 0 {
            self.entries.first().map(|entry| Arc::clone(&entry.before))
        } else {
            self.entries
                .get(self.cursor.saturating_sub(1))
                .map(|entry| Arc::clone(&entry.after))
        };
        if expected.is_some_and(|source| source.as_slice() != patch.before.as_slice()) {
            return Err(Error::PatchConflict {
                part: "workbook history",
            });
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

    /// Apply one exact inverse at the current history tip to `package` and
    /// return the new package; `package` is unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when there is no undo entry, for a package that is
    /// not the entry's after state, and for a failed readback.
    pub fn undo(&mut self, package: &crate::Package) -> Result<Commit> {
        let index = self.cursor.checked_sub(1).ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "workbook history has no undo entry",
            )
        })?;
        let commit = package.apply(&self.entries[index].inverse())?;
        self.cursor = index;
        Ok(commit)
    }

    /// Reapply one entry on the redo side to `package` and return the new
    /// package; `package` is unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when there is no redo entry, for a package that is
    /// not the entry's before state, and for a failed readback.
    pub fn redo(&mut self, package: &crate::Package) -> Result<Commit> {
        let patch = self.entries.get(self.cursor).ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "workbook history has no redo entry",
            )
        })?;
        let commit = package.apply(patch)?;
        self.cursor = self.cursor.saturating_add(1);
        Ok(commit)
    }
}

fn replay(
    before: &[u8],
    operations: &[Operation],
    external_link_limits: ExternalLinkLimits,
    drawing_load_policy: DrawingLoadPolicy,
    raw_limits: RawLimits,
) -> Result<Vec<u8>> {
    let workbook = validated_workbook_with_external_link_limits_and_drawing_policy(
        before,
        external_link_limits,
        drawing_load_policy,
        raw_limits,
    )?;
    replay_workbook(workbook, before, operations).map(|(bytes, _)| bytes)
}

fn replay_package(
    base: &crate::Package,
    before: &Arc<Vec<u8>>,
    operations: &[Operation],
) -> Result<(Arc<Vec<u8>>, crate::Package)> {
    // Durable replay is defined against this exact serialized image, not the
    // in-memory graph's potentially different archive/layout provenance.
    let limits = base.opc_package().read_limits();
    let graph = litchi_opc::OpcPackage::from_bytes(before, limits)?;
    let workbook = base.project_candidate(graph)?;
    validate_all_worksheets(&workbook)?;
    let (bytes, mut workbook) = replay_workbook(workbook, before.as_slice(), operations)?;
    let bytes = if bytes == before.as_slice() {
        return Ok((Arc::clone(before), base.clone()));
    } else {
        // Admit the generated physical image under the same finite policy
        // before publication. This also establishes the precise archive
        // lineage used by later edits/applications. The semantic workbook is
        // already validated by replay; do not parse that projection again.
        Arc::new(bytes)
    };
    let reader = admit_archive_before_copy(bytes.as_slice(), limits)?;
    workbook.package =
        litchi_opc::OpcPackage::from_shared_vec_with_admitted_reader(Arc::clone(&bytes), reader)?;
    Ok((bytes, workbook.into_package()))
}

/// Admit the recipient's physical archive policy before taking ownership of
/// an exact after-image. The returned reader checks input, archive/member
/// counts, names, metadata, and declared ZIP sizes without copying the image;
/// the admitted-reader handoff consumes that same reader while retaining the
/// exact source for publication.
fn admit_archive_before_copy<'data>(
    bytes: &'data [u8],
    limits: litchi_opc::SourceLimits,
) -> Result<litchi_opc::phys_pkg::PhysPkgReader<'data>> {
    let observed = u64::try_from(bytes.len()).map_err(|_error| Error::CapacityOverflow {
        resource: "workbook archive input bytes",
    })?;
    let maximum = limits.max_input_bytes();
    if observed > maximum {
        return Err(Error::Opc(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::InputBytes,
            observed,
            maximum,
        }));
    }
    Ok(litchi_opc::phys_pkg::PhysPkgReader::new(bytes, limits)?)
}

#[cfg(test)]
mod replay_admission_tests {
    use super::admit_archive_before_copy;
    use crate::Package;
    use litchi_opc::{BlobPart, PackURI, SourceLimits};

    #[test]
    fn archive_preflight_checks_member_budget_before_owned_handoff() {
        let bytes = Package::new().unwrap().to_bytes().unwrap();
        let limits = SourceLimits::builder()
            .max_members(1)
            .unwrap()
            .max_parts(1)
            .unwrap()
            .max_relationship_parts(1)
            .unwrap()
            .build()
            .unwrap();

        assert!(
            admit_archive_before_copy(&bytes, limits).is_err(),
            "the borrowed preflight must reject an archive whose member count exceeds the recipient budget"
        );
    }

    #[test]
    fn structural_edit_constructor_and_drop_do_not_capture_package_image() {
        let cold = Package::from(Package::new().unwrap().into_opc());
        assert!(!cold.projection_cached());
        let edit = super::Edit::new(&cold, super::TransferLimits::DEFAULT).unwrap();
        drop(edit);
        assert!(!cold.projection_cached());
    }

    #[test]
    fn structural_preparation_is_failure_atomic_and_retryable() {
        let mut opc = Package::new().unwrap().into_opc();
        opc.add_part(Box::new(BlobPart::new(
            PackURI::new("/invalid.bin").unwrap(),
            "not a content type".to_owned(),
            Vec::new(),
        )));
        let package = Package::from(opc);
        let mut edit = super::Edit::new(&package, super::TransferLimits::DEFAULT).unwrap();

        assert!(edit.ensure_prepared().is_err());
        assert!(edit.sheet_catalog.is_none());
        assert!(edit.before.is_none());
        assert!(edit.ensure_prepared().is_err());
        assert!(edit.sheet_catalog.is_none());
        assert!(edit.before.is_none());
    }
}

fn canonicalize_legacy_operations(
    workbook: &Workbook,
    operations: Vec<Operation>,
) -> Result<Vec<Operation>> {
    let mut canonical = Vec::new();
    canonical
        .try_reserve_exact(operations.len())
        .map_err(|source| Error::Allocation {
            resource: "canonical workbook patch operations",
            source,
        })?;
    for operation in operations {
        let operation = match operation {
            Operation::RenameSheet { sheet, name } => {
                let catalog_position = workbook.catalog_position_for_worksheet(sheet)?;
                Operation::SetSheetName {
                    catalog_position,
                    name,
                    dependencies: false,
                }
            },
            operation @ (Operation::SetSheetName { .. }
            | Operation::ReorderSheets { .. }
            | Operation::InsertSheet { .. }
            | Operation::RemoveSheets { .. }
            | Operation::SetSheetState { .. }
            | Operation::TransferCell { .. }
            | Operation::AddImage { .. }
            | Operation::TransferDrawing { .. }
            | Operation::SetMergedRanges { .. }
            | Operation::SetDefinedName { .. }) => operation,
        };
        canonical.push(operation);
    }
    Ok(canonical)
}

fn validate_reorder_operations(operations: &[Operation]) -> Result<()> {
    let reorder_count = operations
        .iter()
        .filter(|operation| matches!(operation, Operation::ReorderSheets { .. }))
        .count();
    if reorder_count > 1 {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "durable workbook patches may contain at most one sheet-reorder operation",
        ));
    }
    let insert_count = operations
        .iter()
        .filter(|operation| matches!(operation, Operation::InsertSheet { .. }))
        .count();
    let remove_count = operations
        .iter()
        .filter(|operation| matches!(operation, Operation::RemoveSheets { .. }))
        .count();
    let topology_family_count = usize::from(reorder_count > 0)
        .saturating_add(usize::from(insert_count > 0))
        .saturating_add(usize::from(remove_count > 0));
    if topology_family_count > 1 {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet topology operations cannot be combined in one durable patch",
        ));
    }
    if reorder_count == 1
        && operations
            .iter()
            .any(|operation| !matches!(operation, Operation::ReorderSheets { .. }))
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet reordering cannot be combined with other workbook operations",
        ));
    }
    if insert_count > 0
        && operations
            .iter()
            .any(|operation| !matches!(operation, Operation::InsertSheet { .. }))
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet insertion cannot be combined with other workbook operations",
        ));
    }
    if remove_count > 0
        && operations
            .iter()
            .any(|operation| !matches!(operation, Operation::RemoveSheets { .. }))
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet removal cannot be combined with other workbook operations",
        ));
    }
    if remove_count > 1 {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "durable workbook patches may contain at most one sheet-removal operation",
        ));
    }
    Ok(())
}

fn replay_workbook(
    mut workbook: Workbook,
    before: &[u8],
    operations: &[Operation],
) -> Result<(Vec<u8>, Workbook)> {
    if operations.is_empty() {
        return Ok((before.to_vec(), workbook));
    }
    validate_reorder_operations(operations)?;
    let mut changed = false;
    let mut state_changed = false;
    let mut index = 0usize;
    while index < operations.len() {
        if matches!(
            &operations[index],
            Operation::RenameSheet { .. }
                | Operation::SetSheetName { .. }
                | Operation::SetSheetState { .. }
        ) {
            let start = index;
            index += 1;
            while index < operations.len()
                && matches!(
                    &operations[index],
                    Operation::RenameSheet { .. }
                        | Operation::SetSheetName { .. }
                        | Operation::SetSheetState { .. }
                )
            {
                index += 1;
            }
            let (metadata_changed, metadata_state_changed) =
                apply_sheet_metadata(&mut workbook, &operations[start..index])?;
            changed |= metadata_changed;
            state_changed |= metadata_state_changed;
            continue;
        }
        match &operations[index] {
            Operation::TransferCell {
                source,
                source_sheet,
                source_reference,
                target_sheet,
                target_reference,
            } => {
                transfer_cell(
                    &mut workbook,
                    source,
                    *source_sheet,
                    *source_reference,
                    *target_sheet,
                    *target_reference,
                )?;
                changed = true;
            },
            Operation::ReorderSheets { order } => {
                changed |= apply_sheet_order(&mut workbook, order)?;
            },
            Operation::InsertSheet {
                catalog_position,
                name,
                state,
            } => {
                insert_sheet(&mut workbook, *catalog_position, name, *state)?;
                changed = true;
                // Insertion itself does not relocate the active view, but it
                // still must leave the resulting catalog with a valid visible
                // tab before publication.
                state_changed = true;
            },
            Operation::RemoveSheets { positions } => {
                changed |= apply_sheet_removal(&mut workbook, positions)?;
                state_changed = true;
            },
            Operation::AddImage { sheet, image } => {
                add_image(&mut workbook, *sheet, image)?;
                changed = true;
            },
            Operation::SetDefinedName {
                name,
                scope,
                target,
            } => {
                changed |= super::names::apply(&mut workbook, name, *scope, *target)?;
            },
            Operation::SetMergedRanges {
                catalog_position,
                ranges,
            } => {
                let worksheet = workbook
                    .worksheet_for_catalog_position(*catalog_position)
                    .ok_or_else(|| {
                        Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedTarget,
                            "edit",
                            "catalog position is not a worksheet",
                        )
                    })?;
                let before = workbook.merged_cell_ranges(worksheet)?;
                if before[..] != ranges[..] {
                    workbook.set_merged_cell_ranges(worksheet, ranges)?;
                    changed = true;
                }
            },
            Operation::TransferDrawing {
                source,
                source_sheet,
                source_anchor,
                target_sheet,
            } => {
                super::drawing_transfer::transfer(
                    &mut workbook,
                    source,
                    *source_sheet,
                    *source_anchor,
                    *target_sheet,
                )?;
                changed = true;
            },
            Operation::RenameSheet { .. }
            | Operation::SetSheetName { .. }
            | Operation::SetSheetState { .. } => {
                return Err(Error::InvalidFormat(
                    "workbook metadata operation grouping is internally inconsistent",
                ));
            },
        }
        index += 1;
    }
    if state_changed {
        validate_sheet_state_invariants(&workbook)?;
    }
    validate_all_worksheets(&workbook)?;
    let bytes = if changed {
        workbook_bytes(&workbook)?
    } else {
        before.to_vec()
    };
    Ok((bytes, workbook))
}

fn author_style(workbook: &mut Workbook, style: &AuthoredStyle) -> Result<super::StyleIndex> {
    let raw_limits = workbook.raw_limits_value();
    let plan = super::resources::plan_style(style, raw_limits)?;
    crate::signature_policy::ensure_unsigned(&workbook.package)?;
    let mut package = workbook.package.clone();
    let index = super::resources::intern_style_plan(&mut package, &plan, raw_limits)?;
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
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
    let package = workbook.package.clone();
    let leniency = workbook.leniency();
    let worksheet_limits = worksheet_limits_with_raw(workbook);
    let mut edit =
        super::workbook::read_with_limits_and_leniency(&package, &uri, worksheet_limits, leniency)?
            .draft();
    if let Some(formula) = formula {
        edit.insert_formula(reference, style, value, formula)?;
    } else {
        edit.insert(reference, style, value)?;
    }
    let commit = edit.commit()?;
    let external_link_limits = workbook.external_link_limits();
    let applied = super::workbook::apply_retaining_parse(
        &package,
        &uri,
        &commit,
        external_link_limits,
        workbook.drawing_load_policy(),
        leniency,
        workbook.raw_limits_value(),
    )?;
    if let super::workbook::Applied::Published {
        workbook: validated,
        ..
    } = applied
    {
        *workbook = *validated;
    }
    Ok(())
}

fn add_image(workbook: &mut Workbook, sheet: usize, plan: &ImagePlan) -> Result<()> {
    require_eager_drawing_projection(workbook)?;
    crate::signature_policy::ensure_unsigned(&workbook.package)?;
    let worksheet_uri = workbook.worksheet_uri(sheet)?;
    let worksheet_source = workbook.package.get_part(&worksheet_uri)?.blob().to_vec();
    let mut drawing_rel_id = None;
    let raw_limits = workbook.raw_limits_value();
    for item in Records::new(&worksheet_source, raw_limits) {
        let record = item?;
        if record.kind() == kind::DRAWING {
            if drawing_rel_id.is_some() {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "worksheet contains multiple BrtDrawing records",
                ));
            }
            let mut cursor = crate::raw::Cursor::new(record.payload(), "BrtDrawing", raw_limits);
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
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "worksheet has a Drawing relationship without BrtDrawing ownership",
        ));
    }
    let search_end = u32::try_from(workbook.package.iter_parts().count().saturating_add(1))
        .map_err(|_error| Error::InvalidFormat("drawing part search bound exceeds u32"))?;
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
        .ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "no drawing part index remains",
            )
        })?;
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
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
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
    Encoder::new(&mut drawing_payload, raw_limits).write_wide_string(&drawing_rel_id)?;
    let mut worksheet_output = Vec::new();
    let mut inserted = false;
    for item in Records::new(&worksheet_source, raw_limits) {
        let record = item?;
        if !inserted && matches!(record.kind(), kind::BEGIN_LIST_PARTS | kind::END_SHEET) {
            Encoder::new(&mut worksheet_output, raw_limits)
                .write_record(kind::DRAWING, &drawing_payload)?;
            inserted = true;
        }
        copy_record(
            &worksheet_source,
            &record,
            raw_limits,
            &mut worksheet_output,
        )?;
    }
    if !inserted {
        return Err(Error::InvalidFormat(
            "worksheet has no safe BrtDrawing insertion boundary",
        ));
    }
    crate::worksheet_index::maintain(
        &mut package,
        &worksheet_uri,
        &worksheet_source,
        &worksheet_output,
    )?;
    package
        .get_part_mut(&worksheet_uri)?
        .set_blob(worksheet_output);
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
    let catalog_position = workbook.catalog_position_for_worksheet(sheet)?;
    if workbook.sheet_drawing(catalog_position).is_none() {
        return Err(Error::InvalidFormat(
            "authored worksheet drawing failed semantic readback",
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
    let relationship = worksheet_part
        .rels()
        .get(drawing_rel_id)
        .ok_or(Error::InvalidFormat("BrtDrawing relationship is absent"))?;
    if relationship.is_external()
        || !matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::DRAWING
                | litchi_opc::constants::relationship_type::STRICT_DRAWING
        )
    {
        return Err(Error::InvalidFormat(
            "BrtDrawing relationship is external or has the wrong type",
        ));
    }
    let drawing_uri = relationship.target_partname()?;
    let catalog_position = workbook.catalog_position_for_worksheet(sheet)?;
    let drawing = workbook
        .sheet_drawing(catalog_position)
        .ok_or(Error::InvalidFormat(
            "BrtDrawing has no decoded DrawingML inventory",
        ))?;
    let before_images = drawing.images.len();
    let object_id = crate::package::drawing_write::next_drawing_object_id(
        workbook.package.get_part(&drawing_uri)?.blob(),
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
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
        litchi_core::Limits::for_profile(litchi_core::Profile::Server),
    )?;
    drawing_part.set_blob(drawing_xml);
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
    let catalog_position = workbook.catalog_position_for_worksheet(sheet)?;
    let drawing = workbook
        .sheet_drawing(catalog_position)
        .ok_or(Error::InvalidFormat(
            "appended worksheet drawing failed semantic readback",
        ))?;
    let expected_images = before_images
        .checked_add(1)
        .ok_or(Error::InvalidFormat("drawing image count overflow"))?;
    if drawing.images.len() != expected_images
        || drawing.images.last().map(|value| value.data.as_ref()) != Some(plan.data.as_ref())
    {
        return Err(Error::InvalidFormat(
            "appended image dependency failed semantic readback",
        ));
    }
    Ok(())
}

fn free_media_index(package: &litchi_opc::OpcPackage, extension: &str) -> Result<u32> {
    let search_end = u32::try_from(package.iter_parts().count().saturating_add(1))
        .map_err(|_error| Error::InvalidFormat("media part search bound exceeds u32"))?;
    (1..=search_end)
        .find(|index| {
            PackURI::new(format!("/xl/media/image{index}.{extension}"))
                .is_ok_and(|uri| !package.contains_part(&uri))
        })
        .ok_or_else(|| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "no image part index remains",
            )
        })
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

fn require_eager_drawing_projection(workbook: &Workbook) -> Result<()> {
    if workbook.drawing_load_policy() == DrawingLoadPolicy::Skipped {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "drawing authoring requires an eager drawing projection",
        ));
    }
    Ok(())
}

/// Transfer only the selected source cell and its style/string dependencies.
/// The donor's drawing parts are outside that dependency closure and remain
/// opaque regardless of the target workbook's projection policy.
fn transfer_cell(
    target: &mut Workbook,
    source_bytes: &[u8],
    source_sheet: usize,
    source_reference: Reference,
    target_sheet: usize,
    target_reference: Reference,
) -> Result<()> {
    let worksheet_limits = worksheet_limits_with_raw(target);
    let source = validated_workbook_with_external_link_limits_and_drawing_policy(
        source_bytes,
        target.external_link_limits(),
        DrawingLoadPolicy::Skipped,
        target.raw_limits_value(),
    )?;
    let source_snapshot = source.cell_values(source_sheet)?;
    let source_cell = source_snapshot
        .cell(source_reference)?
        .cloned()
        .ok_or(Error::InvalidCellReference("source cell (, ) is absent"))?;
    if target
        .cell_values(target_sheet)?
        .cell(target_reference)?
        .is_some()
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "target cell (, ) is occupied",
        ));
    }
    let mut package = target.package.clone();
    let style = super::resources::transfer_style(
        &source.package,
        &mut package,
        source_cell.style(),
        source.raw_limits_value(),
        target.raw_limits_value(),
    )?;
    let value = match source_cell.value() {
        Value::SharedStringIndex(index) => {
            let index = usize::try_from(*index).map_err(|_error| {
                Error::InvalidCellReference("shared-string index does not fit usize")
            })?;
            let string = source
                .shared_strings()
                .get(index)
                .ok_or(Error::InvalidCellReference(
                    "source shared-string index is absent",
                ))?;
            let string = super::resources::transfer_string_fonts(
                &source.package,
                &mut package,
                string,
                source.raw_limits_value(),
                target.raw_limits_value(),
            )?;
            Value::SharedStringIndex(super::resources::intern_shared_string_for_new_cell(
                &mut package,
                &string,
                target.raw_limits_value(),
            )?)
        },
        Value::RichString(string) => Value::RichString(super::resources::transfer_string_fonts(
            &source.package,
            &mut package,
            string,
            source.raw_limits_value(),
            target.raw_limits_value(),
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
    let leniency = target.leniency();
    let mut edit =
        super::workbook::read_with_limits_and_leniency(&package, &uri, worksheet_limits, leniency)?
            .draft();
    if let Some(formula) = source_cell.formula() {
        let formula =
            super::formula_transfer::remap(&source, target, source_sheet, target_sheet, formula)?;
        edit.insert_formula(target_reference, style, value, formula)?;
    } else {
        edit.insert(target_reference, style, value)?;
    }
    edit.set_show_phonetic(target_reference, source_cell.show_phonetic())?;
    let commit = edit.commit()?;
    let external_link_limits = target.external_link_limits();
    let applied = super::workbook::apply_retaining_parse(
        &package,
        &uri,
        &commit,
        external_link_limits,
        target.drawing_load_policy(),
        leniency,
        target.raw_limits_value(),
    )?;
    if let super::workbook::Applied::Published {
        workbook: validated,
        ..
    } = applied
    {
        *target = *validated;
    }
    Ok(())
}

fn worksheet_limits_with_raw(workbook: &Workbook) -> super::Limits {
    let semantic = workbook.semantic_limits().cell_values_value();
    let raw = workbook.raw_limits_value();
    super::Limits::new(
        semantic.source_bytes(),
        semantic.cells(),
        raw.payload(),
        raw.string_units(),
    )
}

/// Author one blank worksheet and insert its complete-catalog entry.
///
/// The package is edited on a clone and reparsed before the caller adopts it,
/// so a refusal from any dependency leaves the source image untouched. Existing
/// relationship identifiers, tab identifiers, worksheet payloads, and all
/// unrelated package members remain byte-for-byte unchanged.
fn insert_sheet(workbook: &mut Workbook, position: usize, name: &str, state: State) -> Result<()> {
    let catalog_len = workbook.sheets().len();
    if position > catalog_len {
        return Err(Error::InvalidFormat(
            "sheet insertion position exceeds catalog length",
        ));
    }
    validate_sheet_name_value(name)?;
    if workbook.sheets().any(|tab| excel_name_eq(tab.name(), name)) {
        return Err(Error::InvalidFormat("duplicate sheet name"));
    }
    reject_unmodeled_reorder_dependencies(&workbook.package)?;
    reject_shared_workbook_headers(&workbook.package)?;
    crate::signature_policy::ensure_unsigned(&workbook.package)?;

    let workbook_uri = workbook.package.main_document_part()?.partname().clone();
    let workbook_source = workbook.package.get_part(&workbook_uri)?.blob().to_vec();
    let catalog = collect_bundle_records(
        &workbook_source,
        catalog_len,
        workbook.leniency(),
        workbook.raw_limits_value(),
    )?;
    let tab_id = allocate_tab_id(&catalog)?;
    let (worksheet_uri, binary_index_uri) = allocate_worksheet_uris(&workbook.package)?;
    let raw_limits = workbook.raw_limits_value();
    let (worksheet_bytes, binary_index_bytes) = authored_blank_worksheet(name, raw_limits)?;

    let (worksheet_relationship_type, worksheet_content_type) =
        worksheet_relationship_flavor(&workbook.package)?;
    let worksheet_target = worksheet_uri.relative_ref(workbook_uri.base_uri());
    let binary_index_target = binary_index_uri.relative_ref(worksheet_uri.base_uri());

    let mut worksheet_part = BlobPart::new(
        worksheet_uri.clone(),
        worksheet_content_type,
        worksheet_bytes,
    );
    worksheet_part.rels_mut().get_or_add(
        crate::binary_index::BINARY_INDEX_RELATIONSHIP,
        &binary_index_target,
    );
    let binary_index_part = BlobPart::new(
        binary_index_uri,
        crate::binary_index::BINARY_INDEX_CONTENT_TYPE.to_string(),
        binary_index_bytes,
    );

    let mut package = workbook.package.clone();
    package.try_add_part(Box::new(worksheet_part))?;
    package.try_add_part(Box::new(binary_index_part))?;
    let new_rel_id = {
        let workbook_part = package.get_part_mut(&workbook_uri)?;
        workbook_part
            .rels_mut()
            .get_or_add(&worksheet_relationship_type, &worksheet_target)
            .r_id()
            .to_string()
    };
    let new_bundle = bundle_sheet_payload(state, tab_id, &new_rel_id, name, raw_limits)?;
    let workbook_output = rewrite_workbook_insert(
        &workbook_source,
        position,
        catalog_len,
        &workbook.formula_context,
        &new_bundle,
        raw_limits,
    )?;

    let tabs = workbook.sheets().collect::<Vec<_>>();
    let old_titles = tabs
        .iter()
        .map(|tab| tab.name().to_owned())
        .collect::<Vec<_>>();
    let title_refs = old_titles.iter().map(String::as_str).collect::<Vec<_>>();
    let mut replacements = Vec::new();
    replacements
        .try_reserve(package.part_count().saturating_add(2))
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-insertion package replacements",
            source,
        })?;
    replacements.push((workbook_uri, workbook_output));
    for part in package.iter_parts() {
        if part
            .content_type()
            .eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES)
        {
            if let Some(output) = litchi_ooxml_common::xml::sheet_name::insert_title(
                package.get_part(part.partname())?.blob(),
                &title_refs,
                position,
                name,
                excel_name_eq,
            )? {
                replacements.push((part.partname().clone(), output));
            }
        }
    }
    let volatile = crate::volatile_dependencies::reorder_owner(&package, catalog_len)?;
    if let Some(uri) = volatile {
        let source = package.get_part(&uri)?.blob();
        if let Some(output) =
            rewrite_volatile_dependencies_insert(source, position, catalog_len, raw_limits)?
        {
            replacements.push((uri, output));
        }
    }
    package.replace_part_payloads(replacements)?;
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
    Ok(())
}

fn authored_blank_worksheet(
    name: &str,
    raw_limits: crate::raw::Limits,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let mut worksheet = crate::builder::Worksheet::new(name.to_owned());
    // `None` uses the legacy writer default, whose tab-selected flag is true.
    // An inserted tab must not become the active/selected tab implicitly.
    worksheet.set_view(litchi_sheet::view::View {
        tab_selected: false,
        ..Default::default()
    });
    let mut source = Vec::new();
    worksheet.write_with_sparkline_block(
        &mut Encoder::new(&mut source, raw_limits),
        &mut crate::writer::SharedStringsWriter::new(),
        None,
    )?;
    let index = crate::binary_index::encode_for_worksheet(
        &source,
        crate::binary_index::Limits::publication_default(),
    )?;
    Ok((source, index))
}

fn allocate_worksheet_uris(package: &litchi_opc::OpcPackage) -> Result<(PackURI, PackURI)> {
    let attempts = package.part_count().saturating_add(2);
    for ordinal in 1..=attempts {
        let sheet_uri = PackURI::new(format!("/xl/worksheets/sheet{ordinal}.bin"))?;
        let index_uri = PackURI::new(format!("/xl/worksheets/binaryIndex{ordinal}.bin"))?;
        if package.contains_part(&sheet_uri) || package.contains_part(&index_uri) {
            continue;
        }
        if package.validate_new_part_name(&sheet_uri).is_ok()
            && package.validate_new_part_name(&index_uri).is_ok()
        {
            return Ok((sheet_uri, index_uri));
        }
    }
    Err(Error::limit(
        LimitKind::PartNameAttempts,
        attempts,
        attempts,
    ))
}

fn worksheet_relationship_flavor(package: &litchi_opc::OpcPackage) -> Result<(String, String)> {
    let workbook = package.main_document_part()?;
    let mut fallback = litchi_opc::constants::relationship_type::WORKSHEET.to_string();
    for relationship in workbook.rels().iter() {
        let is_worksheet = matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::WORKSHEET
                | litchi_opc::constants::relationship_type::STRICT_WORKSHEET
        );
        if is_worksheet {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            let part = package.get_part(&target)?;
            return Ok((
                relationship.reltype().to_string(),
                part.content_type().to_string(),
            ));
        }
        if relationship.reltype()
            == "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet"
        {
            fallback = litchi_opc::constants::relationship_type::STRICT_WORKSHEET.to_string();
        }
    }
    Ok((fallback, WORKSHEET_CONTENT_TYPE.to_string()))
}

fn collect_bundle_records(
    source: &[u8],
    catalog_len: usize,
    leniency: crate::leniency::Leniency,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<crate::package::records::BundleSheetRecord>> {
    let mut records = Vec::new();
    records
        .try_reserve_exact(catalog_len)
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet insertion catalog records",
            source,
        })?;
    let mut open = false;
    let mut seen = false;
    let mut closed = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_BUNDLE_SHS => {
                if open || seen || !record.payload().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Workbook has duplicate or malformed BrtBundleShs collections",
                    ));
                }
                open = true;
                seen = true;
            },
            kind::BUNDLE_SH if open => {
                let mut tolerance_report = crate::leniency::ToleranceReport::default();
                let ordinal =
                    u32::try_from(records.len()).map_err(|_error| Error::CapacityOverflow {
                        resource: "XLSB sheet insertion catalog ordinal",
                    })?;
                let parsed = crate::package::records::BundleSheetRecord::parse_with(
                    record.payload(),
                    ordinal,
                    crate::leniency::Tolerance::recording_with(&mut tolerance_report, leniency),
                    raw_limits,
                )?;
                records.push(parsed);
            },
            kind::BUNDLE_SH => {
                return Err(Error::InvalidFormat(
                    "BrtBundleSh appears outside its BrtBeginBundleShs collection",
                ));
            },
            kind::END_BUNDLE_SHS => {
                if !open || !record.payload().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Workbook has an unbalanced BrtEndBundleShs collection",
                    ));
                }
                open = false;
                closed = true;
            },
            _ if open => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet insertion cannot detach an unknown record interspersed with BrtBundleSh records",
                ));
            },
            _ => {},
        }
    }
    if open || !closed || !seen || records.len() != catalog_len {
        return Err(Error::InvalidFormat(
            "Workbook BrtBundleSh count differs from catalog size",
        ));
    }
    Ok(records)
}

fn allocate_tab_id(records: &[crate::package::records::BundleSheetRecord]) -> Result<u32> {
    let max_id = usize::from(u16::MAX);
    let mut used = Vec::new();
    used.try_reserve_exact(max_id.saturating_add(1))
        .map_err(|source| Error::Allocation {
            resource: "XLSB stable sheet identifiers",
            source,
        })?;
    used.resize(max_id.saturating_add(1), false);
    for record in records {
        let index = usize::try_from(record.id).map_err(|_error| {
            Error::InvalidFormat("BrtBundleSh stable sheet identifier exceeds usize")
        })?;
        let slot = used.get_mut(index).ok_or(Error::InvalidFormat(
            "BrtBundleSh stable sheet identifier exceeds u16",
        ))?;
        if *slot {
            return Err(Error::InvalidFormat(
                "BrtBundleSh stable sheet identifiers are not unique",
            ));
        }
        *slot = true;
    }
    for candidate in 1..=u32::from(u16::MAX) {
        let index = usize::try_from(candidate)
            .map_err(|_error| Error::InvalidFormat("stable sheet identifier does not fit usize"))?;
        if !used[index] {
            return Ok(candidate);
        }
    }
    Err(Error::limit(
        LimitKind::SheetIdentifiers,
        records.len().saturating_add(1),
        usize::from(u16::MAX),
    ))
}

fn bundle_sheet_payload(
    state: State,
    tab_id: u32,
    rel_id: &str,
    name: &str,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    if !(1..=u32::from(u16::MAX)).contains(&tab_id) || rel_id.is_empty() {
        return Err(Error::InvalidFormat(
            "new BrtBundleSh has an invalid stable identifier or relationship",
        ));
    }
    let mut payload = Vec::new();
    let mut writer = Encoder::new(&mut payload, raw_limits);
    writer.write_u32(u32::from(state_to_wire(state)))?;
    writer.write_u32(tab_id)?;
    writer.write_wide_string(rel_id)?;
    writer.write_wide_string(name)?;
    let _ = crate::package::records::BundleSheetRecord::parse(&payload, raw_limits)?;
    Ok(payload)
}

fn rewrite_workbook_insert(
    source: &[u8],
    position: usize,
    catalog_len: usize,
    context: &crate::package::formula::Context,
    new_bundle_payload: &[u8],
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    if position > catalog_len {
        return Err(Error::InvalidFormat(
            "sheet insertion position exceeds the source catalog",
        ));
    }
    let new_bundle_record = make_record(kind::BUNDLE_SH, new_bundle_payload, raw_limits)?;
    let output_capacity =
        source
            .len()
            .checked_add(new_bundle_record.len())
            .ok_or(Error::CapacityOverflow {
                resource: "XLSB inserted workbook bytes",
            })?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_capacity)
        .map_err(|source| Error::Allocation {
            resource: "XLSB inserted workbook bytes",
            source,
        })?;
    let mut bundle_count = 0usize;
    let mut bundle_open = false;
    let mut saw_bundle_collection = false;
    let mut bundle_collection_closed = false;
    let mut inserted = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        let raw = record_bytes(source, &record, raw_limits)?;
        match record.kind() {
            kind::BEGIN_BUNDLE_SHS => {
                if bundle_open || saw_bundle_collection || !record.payload().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Workbook has duplicate or malformed BrtBundleShs collections",
                    ));
                }
                saw_bundle_collection = true;
                bundle_open = true;
                output.extend_from_slice(raw);
            },
            kind::BUNDLE_SH if bundle_open => {
                if bundle_count == position {
                    output.extend_from_slice(&new_bundle_record);
                    inserted = true;
                }
                output.extend_from_slice(raw);
                bundle_count = bundle_count.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "XLSB worksheet insertion catalog ordinal",
                })?;
            },
            kind::BUNDLE_SH => {
                return Err(Error::InvalidFormat(
                    "BrtBundleSh appears outside its BrtBeginBundleShs collection",
                ));
            },
            kind::END_BUNDLE_SHS => {
                if !bundle_open || !record.payload().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Workbook has an unbalanced BrtEndBundleShs collection",
                    ));
                }
                if bundle_count == position {
                    output.extend_from_slice(&new_bundle_record);
                    inserted = true;
                }
                bundle_open = false;
                bundle_collection_closed = true;
                output.extend_from_slice(raw);
            },
            _ if bundle_open => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet insertion cannot detach an unknown record interspersed with BrtBundleSh records",
                ));
            },
            kind::BEGIN_BOOK
            | kind::END_BOOK
            | kind::FILE_VERSION
            | kind::WORKBOOK_PROP
            | kind::BEGIN_BOOK_VIEWS
            | kind::END_BOOK_VIEWS
            | kind::BOOK_VIEW
            | kind::BEGIN_PIVOT_CACHE_IDS
            | kind::END_PIVOT_CACHE_IDS
            | kind::BEGIN_PIVOT_CACHE_ID
            | kind::END_PIVOT_CACHE_ID
            | kind::BEGIN_EXTERNALS
            | kind::END_EXTERNALS
            | kind::SUP_SELF
            | kind::SUP_SAME
            | kind::SUP_ADDIN
            | kind::SUP_BOOK_SRC
            | kind::EXTERN_SHEET
            | kind::NAME
            | kind::CALC_PROP => {
                let rewritten = rewrite_workbook_record_insert(
                    raw,
                    position,
                    catalog_len,
                    context,
                    raw_limits,
                )?;
                output.extend_from_slice(&rewritten);
            },
            _other => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet insertion cannot preserve unmodeled workbook record at offset without a catalog dependency contract",
                ));
            },
        }
    }
    if bundle_open
        || !bundle_collection_closed
        || !saw_bundle_collection
        || bundle_count != catalog_len
    {
        return Err(Error::InvalidFormat(
            "Workbook BrtBundleSh count differs from catalog size",
        ));
    }
    if !inserted {
        return Err(Error::InvalidFormat(
            "Workbook sheet catalog insertion record was not emitted",
        ));
    }
    Ok(output)
}

fn make_record(
    record_kind: crate::raw::Kind,
    payload: &[u8],
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    Encoder::new(&mut output, raw_limits).write_record(record_kind, payload)?;
    Ok(output)
}

fn rewrite_workbook_record_insert(
    raw: &[u8],
    position: usize,
    catalog_len: usize,
    context: &crate::package::formula::Context,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let (header, header_len) = Header::parse(raw, raw_limits)?;
    let payload = raw.get(header_len..).ok_or(Error::InvalidLength {
        expected: header_len,
        found: raw.len(),
    })?;
    let payload = payload.get(..header.len()).ok_or(Error::InvalidLength {
        expected: header_len.saturating_add(header.len()),
        found: raw.len(),
    })?;
    let mut output = raw.to_vec();
    match header.kind() {
        kind::BOOK_VIEW => {
            if payload.len() != 29 {
                return Err(Error::InvalidLength {
                    expected: 29,
                    found: payload.len(),
                });
            }
            for field_offset in [20usize, 24usize] {
                let old = usize::try_from(u32::from_le_bytes(
                    payload[field_offset..field_offset + 4]
                        .try_into()
                        .map_err(|_error| Error::InvalidLength {
                            expected: field_offset.saturating_add(4),
                            found: payload.len(),
                        })?,
                ))
                .map_err(|_error| Error::InvalidFormula("BrtBookView index overflow"))?;
                if old >= catalog_len {
                    return Err(Error::InvalidFormat(
                        "BrtBookView points outside the sheet catalog",
                    ));
                }
                let new = if old >= position {
                    old.checked_add(1).ok_or(Error::CapacityOverflow {
                        resource: "inserted workbook view index",
                    })?
                } else {
                    old
                };
                output[header_len + field_offset..header_len + field_offset + 4].copy_from_slice(
                    &u32::try_from(new)
                        .map_err(|_error| Error::CapacityOverflow {
                            resource: "inserted workbook view index",
                        })?
                        .to_le_bytes(),
                );
            }
        },
        kind::NAME => {
            if payload.len() < 9 {
                return Err(Error::InvalidLength {
                    expected: 9,
                    found: payload.len(),
                });
            }
            let old = u32::from_le_bytes(payload[5..9].try_into().map_err(|_error| {
                Error::InvalidLength {
                    expected: 9,
                    found: payload.len(),
                }
            })?);
            if old != u32::MAX {
                let old = usize::try_from(old).map_err(|_error| {
                    Error::InvalidFormula("BrtName sheet scope does not fit usize")
                })?;
                if old >= catalog_len {
                    return Err(Error::InvalidFormula(
                        "BrtName sheet scope exceeds the sheet catalog",
                    ));
                }
                let new = if old >= position {
                    old.checked_add(1).ok_or(Error::CapacityOverflow {
                        resource: "inserted BrtName sheet scope",
                    })?
                } else {
                    old
                };
                output[header_len + 5..header_len + 9].copy_from_slice(
                    &u32::try_from(new)
                        .map_err(|_error| Error::CapacityOverflow {
                            resource: "inserted BrtName sheet scope",
                        })?
                        .to_le_bytes(),
                );
            }
        },
        kind::EXTERN_SHEET => {
            let rewritten =
                rewrite_extern_sheet_payload_insert(payload, position, catalog_len, context)?;
            if rewritten != payload {
                output = replace_record_payload(&output, header_len, header.len(), &rewritten)?;
            }
        },
        _ => {},
    }
    Ok(output)
}

fn rewrite_extern_sheet_payload_insert(
    payload: &[u8],
    position: usize,
    catalog_len: usize,
    context: &crate::package::formula::Context,
) -> Result<Vec<u8>> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let count = usize::try_from(u32::from_le_bytes(payload[..4].try_into().map_err(
        |_error| Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        },
    )?))
    .map_err(|_error| Error::InvalidFormula("BrtExternSheet count overflow"))?;
    let expected = 4usize
        .checked_add(count.checked_mul(12).ok_or(Error::CapacityOverflow {
            resource: "BrtExternSheet payload size",
        })?)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtExternSheet payload size",
        })?;
    if payload.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        });
    }
    let mut output = payload.to_vec();
    for entry in 0..count {
        let offset = 4 + entry * 12;
        let external_link =
            u32::from_le_bytes(payload[offset..offset + 4].try_into().map_err(|_error| {
                Error::InvalidLength {
                    expected: offset + 4,
                    found: payload.len(),
                }
            })?);
        let first = i32::from_le_bytes(payload[offset + 4..offset + 8].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: offset + 8,
                found: payload.len(),
            },
        )?);
        let last = i32::from_le_bytes(payload[offset + 8..offset + 12].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: offset + 12,
                found: payload.len(),
            },
        )?);
        let link = context
            .supporting_links
            .get(usize::try_from(external_link).map_err(|_error| {
                Error::InvalidFormula("BrtExternSheet supporting-link index overflow")
            })?)
            .ok_or(Error::InvalidFormula(
                "BrtExternSheet supporting-link index is unresolved",
            ))?;
        if matches!(link, crate::package::formula::SupportingLink::SameSheet)
            && (first, last) != (-2, -2)
        {
            return Err(Error::InvalidFormula(
                "same-sheet BrtExternSheet has invalid endpoints",
            ));
        }
        if !matches!(link, crate::package::formula::SupportingLink::SelfWorkbook) {
            continue;
        }
        if (first, last) == (-2, -2) || (first, last) == (-1, -1) {
            continue;
        }
        if first < 0 || last < first {
            return Err(Error::InvalidFormula(
                "internal BrtExternSheet range has invalid endpoints",
            ));
        }
        let first = usize::try_from(first).map_err(|_error| {
            Error::InvalidFormula("internal BrtExternSheet first endpoint overflow")
        })?;
        let last = usize::try_from(last).map_err(|_error| {
            Error::InvalidFormula("internal BrtExternSheet last endpoint overflow")
        })?;
        if last >= catalog_len {
            return Err(Error::InvalidFormula(
                "internal BrtExternSheet range exceeds the sheet catalog",
            ));
        }
        if first < position && position <= last {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet insertion would broaden an internal BrtExternSheet 3-D range",
            ));
        }
        let first = if first >= position {
            first.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "inserted BrtExternSheet first endpoint",
            })?
        } else {
            first
        };
        let last = if last >= position {
            last.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "inserted BrtExternSheet last endpoint",
            })?
        } else {
            last
        };
        output[offset + 4..offset + 8].copy_from_slice(
            &i32::try_from(first)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "inserted BrtExternSheet first endpoint",
                })?
                .to_le_bytes(),
        );
        output[offset + 8..offset + 12].copy_from_slice(
            &i32::try_from(last)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "inserted BrtExternSheet last endpoint",
                })?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_volatile_dependencies_insert(
    source: &[u8],
    position: usize,
    catalog_len: usize,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Vec<u8>>> {
    let mut output = source.to_vec();
    let mut changed = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        if record.kind() != kind::VOL_REF {
            continue;
        }
        if record.payload().len() != 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: record.payload().len(),
            });
        }
        let old = usize::try_from(u32::from_le_bytes(
            record.payload()[8..12]
                .try_into()
                .map_err(|_error| Error::InvalidLength {
                    expected: 12,
                    found: record.payload().len(),
                })?,
        ))
        .map_err(|_error| Error::InvalidFormula("BrtVolRef sheet index overflow"))?;
        if old >= catalog_len {
            return Err(Error::InvalidFormat(
                "BrtVolRef sheet index exceeds the sheet catalog",
            ));
        }
        let new = if old >= position {
            old.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "inserted BrtVolRef sheet index",
            })?
        } else {
            old
        };
        let (_, header_len) = Header::parse(
            source
                .get(record.offset()..)
                .ok_or(Error::InvalidFormat("BrtVolRef offset is outside its part"))?,
            raw_limits,
        )?;
        output[record.offset() + header_len + 8..record.offset() + header_len + 12]
            .copy_from_slice(
                &u32::try_from(new)
                    .map_err(|_error| Error::CapacityOverflow {
                        resource: "inserted BrtVolRef sheet index",
                    })?
                    .to_le_bytes(),
            );
        changed |= new != old;
    }
    Ok(changed.then_some(output))
}

/// Apply one source-catalog sheet removal. All physical graph work is done on
/// a package clone and adopted only after a strict workbook reparse succeeds.
fn apply_sheet_removal(workbook: &mut Workbook, positions: &[usize]) -> Result<bool> {
    let catalog_len = workbook.sheets().len();
    let (remaining, removed) = validate_sheet_removal_positions(positions, catalog_len)?;
    if removed.iter().all(|value| !*value) {
        return Ok(false);
    }
    reject_unmodeled_reorder_dependencies(&workbook.package)?;
    reject_shared_workbook_headers(&workbook.package)?;
    crate::signature_policy::ensure_unsigned(&workbook.package)?;

    let workbook_uri = workbook.package.main_document_part()?.partname().clone();
    let workbook_source = workbook.package.get_part(&workbook_uri)?.blob().to_vec();
    let raw_limits = workbook.raw_limits_value();
    let bundle_records = collect_bundle_records(
        &workbook_source,
        catalog_len,
        workbook.leniency(),
        raw_limits,
    )?;
    let selected_rel_ids = bundle_records
        .iter()
        .enumerate()
        .filter(|(position, _)| removed[*position])
        .filter_map(|(_, record)| record.rel_id.clone())
        .collect::<Vec<_>>();
    let (roots, modeled_parts) =
        modeled_sheet_removal_parts(&workbook.package, &workbook_uri, &selected_rel_ids)?;
    reject_unmodeled_binary_sheet_name_dependencies(&workbook.package)?;
    reject_surviving_binary_sheet_dependencies(
        &workbook.package,
        &modeled_parts,
        &bundle_records,
        &removed,
        workbook.raw_limits_value(),
    )?;
    reject_sheet_xml_dependencies(&workbook.package, &modeled_parts, &bundle_records, &removed)?;

    for (position, tab) in workbook.sheets().enumerate() {
        if removed[position] {
            continue;
        }
        let Some(ordinal) = tab.worksheet_position() else {
            continue;
        };
        let snapshot = workbook.cell_values(ordinal.get())?;
        for cell in snapshot.cells() {
            if let Some(formula) = cell.formula() {
                let tokens = Parser::with_extra(
                    formula.tokens(),
                    formula.ancillary(),
                    workbook.raw_limits_value(),
                )
                .parse()
                .map_err(map_formula_error)?;
                reject_removed_formula_targets(
                    tokens,
                    &removed,
                    &workbook.formula_context,
                    Some(position),
                )?;
            }
        }
    }

    let workbook_output = rewrite_workbook_remove(
        &workbook_source,
        positions,
        &removed,
        &remaining,
        catalog_len,
        &workbook.formula_context,
        &bundle_records,
        workbook.raw_limits_value(),
    )?;
    let tabs = workbook.sheets().collect::<Vec<_>>();
    let titles = tabs
        .iter()
        .map(|tab| tab.name().to_owned())
        .collect::<Vec<_>>();
    let title_refs = titles.iter().map(String::as_str).collect::<Vec<_>>();
    let volatile = crate::volatile_dependencies::reorder_owner(&workbook.package, catalog_len)?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve(workbook.package.part_count().saturating_add(2))
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-removal package replacements",
            source,
        })?;
    replacements.push((workbook_uri.clone(), workbook_output));
    for part in workbook.package.iter_parts() {
        if part
            .content_type()
            .eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES)
        {
            if let Some(output) = litchi_ooxml_common::xml::sheet_name::remove_titles(
                workbook.package.get_part(part.partname())?.blob(),
                &title_refs,
                positions,
                excel_name_eq,
            )? {
                replacements.push((part.partname().clone(), output));
            }
        }
        if volatile.as_ref().is_some_and(|uri| uri == part.partname()) {
            if let Some(output) = rewrite_volatile_dependencies_remove(
                workbook.package.get_part(part.partname())?.blob(),
                &removed,
                &remaining,
                catalog_len,
                raw_limits,
            )? {
                replacements.push((part.partname().clone(), output));
            }
        }
    }
    let mut package = workbook.package.clone();
    package.replace_part_payloads(replacements)?;
    apply_sheet_removal_view_dependencies(
        &workbook_source,
        &mut package,
        &workbook_uri,
        &bundle_records,
        &removed,
        &tabs,
        raw_limits,
    )?;
    remove_modeled_sheet_parts(
        &mut package,
        &workbook_uri,
        &selected_rel_ids,
        &roots,
        &modeled_parts,
    )?;
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
    Ok(true)
}

fn validate_sheet_removal_positions(
    positions: &[usize],
    catalog_len: usize,
) -> Result<(Vec<usize>, Vec<bool>)> {
    if positions.is_empty() {
        return Err(Error::InvalidFormat(
            "sheet removal must select at least one catalog position",
        ));
    }
    let mut removed = Vec::new();
    removed
        .try_reserve_exact(catalog_len)
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-removal catalog mask",
            source,
        })?;
    removed.resize(catalog_len, false);
    let mut previous = None;
    for &position in positions {
        if position >= catalog_len {
            return Err(Error::InvalidFormat(
                "sheet removal position exceeds catalog length",
            ));
        }
        if previous.is_some_and(|old| position <= old) {
            return Err(Error::InvalidFormat(
                "sheet-removal positions must be sorted and unique",
            ));
        }
        previous = Some(position);
        removed[position] = true;
    }
    let mut remaining = Vec::new();
    remaining
        .try_reserve_exact(catalog_len)
        .map_err(|source| Error::Allocation {
            resource: "XLSB remaining sheet positions",
            source,
        })?;
    remaining.resize(catalog_len, usize::MAX);
    let mut next = 0usize;
    for (old, is_removed) in removed.iter().copied().enumerate() {
        if !is_removed {
            remaining[old] = next;
            next = next.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "XLSB remaining sheet position",
            })?;
        }
    }
    if next == 0 {
        return Err(Error::InvalidFormat(
            "sheet removal must leave at least one catalog tab",
        ));
    }
    Ok((remaining, removed))
}

/// Resolve the physical parts a tab removal is allowed to delete.
///
/// A sheet root is a modeled owner.  The only modeled internal child is its
/// exclusive worksheet binary-index part.  Every other internal outgoing edge
/// would lose content unless a format-specific cascade or detach operation is
/// available, so it is refused before any candidate package is mutated.
fn modeled_sheet_removal_parts(
    package: &litchi_opc::OpcPackage,
    workbook_uri: &PackURI,
    selected_rel_ids: &[String],
) -> Result<(Vec<PackURI>, HashSet<PackURI>)> {
    let workbook = package.get_part(workbook_uri)?;
    let mut roots: Vec<PackURI> = Vec::new();
    roots
        .try_reserve_exact(selected_rel_ids.len())
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-removal relationship roots",
            source,
        })?;
    for rel_id in selected_rel_ids {
        let relationship = workbook.rels().get(rel_id).ok_or({
            Error::InvalidRelationship(
                "selected BrtBundleSh relationship is absent from the workbook",
            )
        })?;
        if relationship.target_mode() != litchi_opc::TargetMode::Internal {
            return Err(Error::InvalidRelationship(
                "selected XLSB sheet relationship must target an internal part",
            ));
        }
        let target = relationship.target_partname()?;
        let canonical = package
            .part_metadata(&target)
            .map(|part| part.partname().clone())
            .ok_or(Error::InvalidRelationship(
                "selected XLSB sheet relationship target is absent",
            ))?;
        if !roots.iter().any(|root| root.is_equivalent_to(&canonical)) {
            roots.push(canonical);
        }
    }
    let mut modeled = HashSet::new();
    modeled
        .try_reserve(roots.len())
        .map_err(|source| Error::Allocation {
            resource: "XLSB modeled sheet-removal parts",
            source,
        })?;
    for root in &roots {
        if !modeled.insert(root.clone()) {
            return Err(Error::InvalidRelationship(
                "multiple selected workbook tabs resolve to one physical sheet part",
            ));
        }
    }

    let mut index_owners: HashSet<PackURI> = HashSet::new();
    for root in &roots {
        let part = package.get_part(root)?;
        let mut index_relationship = None;
        for relationship in part.rels().iter() {
            if relationship.reltype() == crate::binary_index::BINARY_INDEX_RELATIONSHIP {
                if index_relationship.is_some() {
                    return Err(Error::InvalidRelationship(
                        "worksheet has multiple binary-index relationships",
                    ));
                }
                if relationship.target_mode() != litchi_opc::TargetMode::Internal {
                    return Err(Error::InvalidRelationship(
                        "worksheet binary-index relationship cannot be external",
                    ));
                }
                index_relationship = Some(relationship);
            } else if relationship.target_mode() == litchi_opc::TargetMode::Internal {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal requires a disposition for outgoing internal relationship",
                ));
            }
        }
        let Some(relationship) = index_relationship else {
            continue;
        };
        if part.content_type() != WORKSHEET_CONTENT_TYPE {
            return Err(Error::InvalidRelationship(
                "binary-index relationship has a non-worksheet owner",
            ));
        }
        let target = canonical_sheet_removal_target(package, relationship)?;
        if target.is_equivalent_to(root) {
            return Err(Error::InvalidRelationship(
                "worksheet binary-index relationship targets its worksheet",
            ));
        }
        let index = package.get_part(&target)?;
        if index.content_type() != crate::binary_index::BINARY_INDEX_CONTENT_TYPE {
            return Err(Error::InvalidContentType {
                expected: crate::binary_index::BINARY_INDEX_CONTENT_TYPE,
            });
        }
        let canonical = index.partname().clone();
        if index_owners
            .iter()
            .any(|owner| owner.is_equivalent_to(&canonical))
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet removal requires a disposition for a binary-index part shared by selected tabs",
            ));
        }
        index_owners.insert(canonical.clone());
        modeled.insert(canonical);
    }

    // A binary index is an exclusive leaf.  External links may be discarded
    // with it, but an internal child would require another modeled cascade.
    for index in &index_owners {
        let part = package.get_part(index)?;
        for relationship in part.rels().iter() {
            if relationship.target_mode() == litchi_opc::TargetMode::Internal {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal requires a disposition for outgoing internal relationship from binary-index part",
                ));
            }
        }
    }

    // Validate every incoming edge before publication.  Workbook catalog
    // edges to selected roots and the matching root-to-index edge are the
    // only relationships that prove ownership; package-level or surviving
    // part references would otherwise become dangling.
    for relationship in package.rels().iter() {
        if relationship.target_mode() != litchi_opc::TargetMode::Internal {
            continue;
        }
        let target = canonical_sheet_removal_target(package, relationship)?;
        if modeled.iter().any(|part| part.is_equivalent_to(&target)) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet removal requires a disposition for an incoming package relationship to a selected part",
            ));
        }
    }
    for source in package.iter_parts() {
        let source_is_workbook = source.partname().is_equivalent_to(workbook_uri);
        let source_is_root = roots
            .iter()
            .any(|root| root.is_equivalent_to(source.partname()));
        for relationship in source.rels().iter() {
            if relationship.target_mode() != litchi_opc::TargetMode::Internal {
                continue;
            }
            let target = canonical_sheet_removal_target(package, relationship)?;
            let Some(selected_target) = modeled.iter().find(|part| part.is_equivalent_to(&target))
            else {
                continue;
            };
            let target_is_root = roots
                .iter()
                .any(|root| root.is_equivalent_to(selected_target));
            let target_is_index = index_owners
                .iter()
                .any(|index| index.is_equivalent_to(selected_target));
            let allowed = if target_is_root && source_is_workbook {
                selected_rel_ids
                    .iter()
                    .any(|selected_id| selected_id == relationship.r_id())
            } else if target_is_index && source_is_root {
                relationship.reltype() == crate::binary_index::BINARY_INDEX_RELATIONSHIP
            } else {
                false
            };
            if !allowed {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal requires a disposition for incoming relationship from to a selected part",
                ));
            }
        }
    }

    Ok((roots, modeled))
}

fn reject_unmodeled_binary_sheet_name_dependencies(package: &litchi_opc::OpcPackage) -> Result<()> {
    const XLSB_PIVOT_CACHE_DEFINITION_CONTENT_TYPE: &str =
        "application/vnd.ms-excel.pivotCacheDefinition";
    if package.iter_parts().any(|part| {
        part.content_type()
            .eq_ignore_ascii_case(XLSB_PIVOT_CACHE_DEFINITION_CONTENT_TYPE)
            || part
                .content_type()
                .eq_ignore_ascii_case("application/vnd.ms-excel.table")
    }) {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet removal requires a disposition for XLSB pivot-cache or table formula dependencies",
        ));
    }
    Ok(())
}

fn reject_surviving_binary_sheet_dependencies(
    package: &litchi_opc::OpcPackage,
    removed_parts: &HashSet<PackURI>,
    records: &[crate::package::records::BundleSheetRecord],
    removed: &[bool],
    raw_limits: crate::raw::Limits,
) -> Result<()> {
    for part in package.iter_parts() {
        if !part
            .content_type()
            .eq_ignore_ascii_case(WORKSHEET_CONTENT_TYPE)
            || removed_parts
                .iter()
                .any(|removed| removed.is_equivalent_to(part.partname()))
        {
            continue;
        }
        let source = package.get_part(part.partname())?.blob();
        for item in Records::new(source, raw_limits) {
            let record = item?;
            match record.kind() {
                kind::FRT_BEGIN | kind::AC_BEGIN | kind::ARR_FMLA | kind::SHR_FMLA => {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit",
                        "sheet removal requires a disposition for a surviving worksheet extension or formula-group record",
                    ));
                },
                kind::H_LINK => {
                    let hyperlink =
                        crate::hyperlinks::Hyperlink::parse(record.payload(), raw_limits)?;
                    if !hyperlink.r_id.is_empty() {
                        let Some(relationship) = part.rels().get(&hyperlink.r_id) else {
                            return Err(Error::refused(
                                litchi_core::edit::RefusalKind::UnsupportedContent,
                                "edit",
                                "sheet removal requires a disposition for an unresolved surviving BrtHLink relationship",
                            ));
                        };
                        if relationship.target_mode() != litchi_opc::TargetMode::External
                            || !matches!(
                                relationship.reltype(),
                                litchi_opc::constants::relationship_type::HYPERLINK
                                    | litchi_opc::constants::relationship_type::STRICT_HYPERLINK
                            )
                        {
                            return Err(Error::refused(
                                litchi_core::edit::RefusalKind::UnsupportedContent,
                                "edit",
                                "sheet removal requires a disposition for a non-external surviving BrtHLink relationship",
                            ));
                        }
                        continue;
                    }
                    let Some(location) = hyperlink.location.as_deref() else {
                        continue;
                    };
                    if internal_hyperlink_targets_removed_sheet(location, records, removed) {
                        return Err(Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedContent,
                            "edit",
                            "sheet removal requires a disposition for a surviving BrtHLink internal location",
                        ));
                    }
                },
                _ => {},
            }
        }
    }
    Ok(())
}

fn internal_hyperlink_targets_removed_sheet(
    location: &str,
    records: &[crate::package::records::BundleSheetRecord],
    removed: &[bool],
) -> bool {
    let mut location = location.trim();
    if let Some(stripped) = location.strip_prefix('#') {
        location = stripped;
    }
    let (sheet, quoted) = if location.starts_with('\'') {
        let bytes = location.as_bytes();
        let mut index = 1usize;
        let mut close = None;
        while index < bytes.len() {
            if bytes[index] != b'\'' {
                index = index.saturating_add(1);
                continue;
            }
            if bytes.get(index.saturating_add(1)) == Some(&b'\'') {
                index = index.saturating_add(2);
                continue;
            }
            close = Some(index);
            break;
        }
        let Some(close) = close else {
            return true;
        };
        if bytes.get(close.saturating_add(1)) != Some(&b'!') {
            return true;
        }
        (&location[1..close], true)
    } else {
        let Some(separator) = location.find('!') else {
            return true;
        };
        (&location[..separator], false)
    };
    let mut unquoted = sheet.trim();
    if unquoted.starts_with('[') {
        return false;
    }
    if quoted && unquoted.contains("''") {
        let mut decoded = String::with_capacity(unquoted.len());
        let mut chars = unquoted.chars();
        while let Some(character) = chars.next() {
            if character == '\'' && chars.as_str().starts_with('\'') {
                decoded.push('\'');
                chars.next();
            } else {
                decoded.push(character);
            }
        }
        unquoted = decoded.as_str();
        return records.iter().enumerate().any(|(position, record)| {
            removed.get(position).copied().unwrap_or(false) && excel_name_eq(unquoted, &record.name)
        });
    }
    records.iter().enumerate().any(|(position, record)| {
        removed.get(position).copied().unwrap_or(false) && excel_name_eq(unquoted, &record.name)
    })
}

fn canonical_sheet_removal_target(
    package: &litchi_opc::OpcPackage,
    relationship: &litchi_opc::Relationship,
) -> Result<PackURI> {
    let target = relationship.target_partname()?;
    package
        .part_metadata(&target)
        .map(|part| part.partname().clone())
        .ok_or(Error::InvalidRelationship(
            "sheet-removal relationship target is absent",
        ))
}

fn remove_modeled_sheet_parts(
    package: &mut litchi_opc::OpcPackage,
    workbook_uri: &PackURI,
    selected_rel_ids: &[String],
    roots: &[PackURI],
    modeled: &HashSet<PackURI>,
) -> Result<()> {
    let mut to_remove = Vec::new();
    to_remove
        .try_reserve_exact(modeled.len())
        .map_err(|source| Error::Allocation {
            resource: "XLSB modeled sheet-removal deletion set",
            source,
        })?;
    for part in package.iter_parts() {
        if modeled
            .iter()
            .any(|removed| removed.is_equivalent_to(part.partname()))
        {
            to_remove.push(part.partname().clone());
        }
    }
    for part in package.iter_parts() {
        if to_remove
            .iter()
            .any(|partname| partname.is_equivalent_to(part.partname()))
        {
            continue;
        }
        for relationship in part.rels().iter() {
            if relationship.target_mode() != litchi_opc::TargetMode::Internal {
                continue;
            }
            let target = relationship.target_partname()?;
            if !roots.iter().any(|root| target.is_equivalent_to(root)) {
                continue;
            }
            if part.partname().is_equivalent_to(workbook_uri)
                && selected_rel_ids
                    .iter()
                    .any(|selected_id| selected_id == relationship.r_id())
            {
                continue;
            }
            return Err(Error::InvalidRelationship(
                "a surviving XLSB package part references a removed sheet part",
            ));
        }
    }
    {
        let workbook_part = package.get_part_mut(workbook_uri)?;
        for rel_id in selected_rel_ids {
            if workbook_part.rels_mut().remove(rel_id).is_none() {
                return Err(Error::InvalidRelationship(
                    "selected workbook relationship disappeared before removal",
                ));
            }
        }
    }
    for partname in to_remove {
        package.remove_part(&partname);
    }
    Ok(())
}

fn reject_sheet_xml_dependencies(
    package: &litchi_opc::OpcPackage,
    removed_parts: &HashSet<PackURI>,
    records: &[crate::package::records::BundleSheetRecord],
    removed: &[bool],
) -> Result<()> {
    let mut replacement_name = "__litchi_removed_sheet__".to_string();
    let mut suffix = 0usize;
    while records
        .iter()
        .any(|record| excel_name_eq(&record.name, &replacement_name))
    {
        suffix = suffix.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "sheet-removal XML guard name",
        })?;
        replacement_name = format!("__litchi_removed_sheet_{suffix}__");
    }
    let mut renames = Vec::new();
    for (position, record) in records.iter().enumerate() {
        if removed[position] {
            renames.push(litchi_ooxml_common::xml::sheet_name::Rename {
                before: record.name.as_str(),
                after: replacement_name.as_str(),
            });
        }
    }
    if renames.is_empty() {
        return Ok(());
    }
    for part in package.iter_parts() {
        if removed_parts
            .iter()
            .any(|removed| removed.is_equivalent_to(part.partname()))
            || part
                .content_type()
                .eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES)
            || part
                .content_type()
                .eq_ignore_ascii_case(litchi_opc::constants::content_type::SML_EXTERNAL_LINK)
            || part
                .content_type()
                .eq_ignore_ascii_case("application/vnd.ms-excel.externalLink")
        {
            continue;
        }
        let content_type = part.content_type().to_ascii_lowercase();
        let part_name = part.partname().as_str().to_ascii_lowercase();
        if !(content_type.contains("xml") || part_name.ends_with(".vml")) {
            continue;
        }
        if litchi_ooxml_common::xml::sheet_name::rewrite(
            package.get_part(part.partname())?.blob(),
            &renames,
            &[],
            excel_name_eq,
        )?
        .is_some()
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet removal is unsupported while a surviving XML/VML part references a removed sheet",
            ));
        }
    }
    Ok(())
}

fn rewrite_workbook_remove(
    source: &[u8],
    positions: &[usize],
    removed: &[bool],
    remaining: &[usize],
    catalog_len: usize,
    context: &crate::package::formula::Context,
    bundle_records: &[crate::package::records::BundleSheetRecord],
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut bundle_count = 0usize;
    let mut bundle_open = false;
    let mut saw_collection = false;
    let mut closed = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_BUNDLE_SHS => {
                if bundle_open || saw_collection || !record.payload().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Workbook has duplicate or malformed BrtBundleShs collections",
                    ));
                }
                bundle_open = true;
                saw_collection = true;
            },
            kind::BUNDLE_SH if bundle_open => {
                bundle_count = bundle_count.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "XLSB removed catalog records",
                })?;
            },
            kind::BUNDLE_SH => {
                return Err(Error::InvalidFormat(
                    "BrtBundleSh appears outside its BrtBeginBundleShs collection",
                ));
            },
            kind::END_BUNDLE_SHS => {
                if !bundle_open || !record.payload().is_empty() {
                    return Err(Error::InvalidFormat(
                        "Workbook has an unbalanced BrtEndBundleShs collection",
                    ));
                }
                bundle_open = false;
                closed = true;
            },
            _ if bundle_open => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal cannot detach an unknown record interspersed with BrtBundleSh records",
                ));
            },
            kind::BEGIN_BOOK
            | kind::END_BOOK
            | kind::FILE_VERSION
            | kind::WORKBOOK_PROP
            | kind::BEGIN_BOOK_VIEWS
            | kind::END_BOOK_VIEWS
            | kind::BOOK_VIEW
            | kind::BEGIN_PIVOT_CACHE_IDS
            | kind::END_PIVOT_CACHE_IDS
            | kind::BEGIN_PIVOT_CACHE_ID
            | kind::END_PIVOT_CACHE_ID
            | kind::BEGIN_EXTERNALS
            | kind::END_EXTERNALS
            | kind::SUP_SELF
            | kind::SUP_SAME
            | kind::SUP_ADDIN
            | kind::SUP_BOOK_SRC
            | kind::EXTERN_SHEET
            | kind::NAME
            | kind::CALC_PROP => {},
            _other => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal cannot preserve unmodeled workbook record at offset without a catalog dependency contract",
                ));
            },
        }
    }
    if bundle_open || !closed || !saw_collection || bundle_count != catalog_len {
        return Err(Error::InvalidFormat(
            "Workbook BrtBundleSh count differs from catalog size",
        ));
    }
    if bundle_records.len() != catalog_len
        || positions.iter().any(|position| *position >= catalog_len)
    {
        return Err(Error::InvalidFormat(
            "sheet-removal catalog metadata is inconsistent",
        ));
    }
    removal_book_view_transitions(source, removed, bundle_records, raw_limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "XLSB removed workbook bytes",
            source,
        })?;
    let mut bundle_index = 0usize;
    let mut open = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        let raw = record_bytes(source, &record, raw_limits)?;
        match record.kind() {
            kind::BEGIN_BUNDLE_SHS => {
                open = true;
                output.extend_from_slice(raw);
            },
            kind::BUNDLE_SH if open => {
                let position = bundle_index;
                bundle_index = bundle_index.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "XLSB removal catalog ordinal",
                })?;
                if !removed[position] {
                    output.extend_from_slice(raw);
                }
            },
            kind::END_BUNDLE_SHS => {
                open = false;
                output.extend_from_slice(raw);
            },
            _ => {
                let rewritten = rewrite_workbook_record_remove(
                    raw,
                    removed,
                    remaining,
                    catalog_len,
                    context,
                    bundle_records,
                    raw_limits,
                )?;
                output.extend_from_slice(&rewritten);
            },
        }
    }
    Ok(output)
}

fn rewrite_workbook_record_remove(
    raw: &[u8],
    removed: &[bool],
    remaining: &[usize],
    catalog_len: usize,
    context: &crate::package::formula::Context,
    bundle_records: &[crate::package::records::BundleSheetRecord],
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let (header, header_len) = Header::parse(raw, raw_limits)?;
    let payload = raw.get(header_len..).ok_or(Error::InvalidLength {
        expected: header_len,
        found: raw.len(),
    })?;
    let payload = payload.get(..header.len()).ok_or(Error::InvalidLength {
        expected: header_len.saturating_add(header.len()),
        found: raw.len(),
    })?;
    let mut output = raw.to_vec();
    match header.kind() {
        kind::BOOK_VIEW => {
            if payload.len() != 29 {
                return Err(Error::InvalidLength {
                    expected: 29,
                    found: payload.len(),
                });
            }
            for field_offset in [20usize, 24usize] {
                let old = usize::try_from(u32::from_le_bytes(
                    payload[field_offset..field_offset + 4]
                        .try_into()
                        .map_err(|_error| Error::InvalidLength {
                            expected: field_offset.saturating_add(4),
                            found: payload.len(),
                        })?,
                ))
                .map_err(|_error| Error::InvalidFormula("BrtBookView index overflow"))?;
                if old >= catalog_len || old >= removed.len() {
                    return Err(Error::InvalidFormat(
                        "BrtBookView points outside the sheet catalog",
                    ));
                }
                let new = if removed[old] {
                    let fallback = fallback_visible_position(old, removed, bundle_records)?;
                    remaining[fallback]
                } else {
                    *remaining
                        .get(old)
                        .ok_or(Error::InvalidFormat("BrtBookView index is unmapped"))?
                };
                output[header_len + field_offset..header_len + field_offset + 4].copy_from_slice(
                    &u32::try_from(new)
                        .map_err(|_error| Error::CapacityOverflow {
                            resource: "removed workbook view index",
                        })?
                        .to_le_bytes(),
                );
            }
        },
        kind::NAME => {
            reject_name_formula_dependencies(payload, removed, context, raw_limits)?;
            if payload.len() < 9 {
                return Err(Error::InvalidLength {
                    expected: 9,
                    found: payload.len(),
                });
            }
            let old = u32::from_le_bytes(payload[5..9].try_into().map_err(|_error| {
                Error::InvalidLength {
                    expected: 9,
                    found: payload.len(),
                }
            })?);
            if old != u32::MAX {
                let old = usize::try_from(old).map_err(|_error| {
                    Error::InvalidFormula("BrtName sheet scope does not fit usize")
                })?;
                if old >= catalog_len {
                    return Err(Error::InvalidFormula(
                        "BrtName sheet scope exceeds the sheet catalog",
                    ));
                }
                if removed[old] {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit",
                        "sheet removal cannot safely delete a sheet-scoped BrtName",
                    ));
                }
                output[header_len + 5..header_len + 9].copy_from_slice(
                    &u32::try_from(remaining[old])
                        .map_err(|_error| Error::CapacityOverflow {
                            resource: "removed BrtName sheet scope",
                        })?
                        .to_le_bytes(),
                );
            }
        },
        kind::EXTERN_SHEET => {
            let rewritten = rewrite_extern_sheet_payload_remove(
                payload,
                removed,
                remaining,
                catalog_len,
                context,
            )?;
            if rewritten != payload {
                output = replace_record_payload(&output, header_len, header.len(), &rewritten)?;
            }
        },
        _ => {},
    }
    Ok(output)
}

fn reject_name_formula_dependencies(
    payload: &[u8],
    removed: &[bool],
    context: &crate::package::formula::Context,
    raw_limits: crate::raw::Limits,
) -> Result<()> {
    let definition =
        crate::named_ranges::parse(payload, raw_limits).map_err(map_named_range_error)?;
    let Some(formula) = definition.formula.as_deref() else {
        return Ok(());
    };
    let scope = definition
        .sheet_id
        .map(|sheet| {
            usize::try_from(sheet).map_err(|_error| {
                Error::InvalidFormula("defined-name sheet scope does not fit usize")
            })
        })
        .transpose()?;
    let tokens = Parser::new(formula, raw_limits)
        .parse()
        .map_err(map_formula_error)?;
    reject_removed_formula_targets(tokens, removed, context, scope)
}

fn reject_removed_formula_targets(
    tokens: Vec<Token>,
    removed: &[bool],
    context: &crate::package::formula::Context,
    scope: Option<usize>,
) -> Result<()> {
    for token in tokens {
        let xti = match token {
            Token::CellRef3d { sheet_index, .. }
            | Token::AreaRef3d { sheet_index, .. }
            | Token::ExternalName { sheet_index, .. } => Some(sheet_index),
            Token::TableReference(reference) if !reference.invalid => Some(reference.sheet_index),
            Token::Unknown(_kind) => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal requires a disposition for unknown formula token",
                ));
            },
            Token::Number(_)
            | Token::String(_)
            | Token::Bool(_)
            | Token::Error(_)
            | Token::Int(_)
            | Token::MissingArg
            | Token::Parenthesis
            | Token::Attribute(_)
            | Token::Array { .. }
            | Token::Memory { .. }
            | Token::CellRef { .. }
            | Token::AreaRef { .. }
            | Token::ReferenceError { .. }
            | Token::BinaryOp(_)
            | Token::UnaryOp(_)
            | Token::Function { .. }
            | Token::Name(_)
            | Token::TableReference(_)
            | Token::PivotName(_) => None,
        };
        let Some(xti) = xti else {
            continue;
        };
        if name_formula_xti_removed(context, xti, removed, scope)? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet removal requires a disposition for a formula referencing a removed tab",
            ));
        }
    }
    Ok(())
}

fn name_formula_xti_removed(
    context: &crate::package::formula::Context,
    index: u16,
    removed: &[bool],
    scope: Option<usize>,
) -> Result<bool> {
    let xti = context
        .external_sheets
        .get(usize::from(index))
        .ok_or(Error::InvalidFormula("defined-name XTI index is absent"))?;
    let link =
        context
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_error| {
                Error::InvalidFormula("defined-name external-link index overflow")
            })?)
            .ok_or(Error::InvalidFormula(
                "defined-name XTI index has no supporting link",
            ))?;
    match link {
        crate::package::formula::SupportingLink::SelfWorkbook => {
            if (xti.first_sheet, xti.last_sheet) == (-1, -1) {
                return Ok(false);
            }
            if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
                return Err(Error::InvalidFormula(
                    "defined-name XTI index has invalid internal sheet range",
                ));
            }
            let first = usize::try_from(xti.first_sheet)
                .map_err(|_error| Error::InvalidFormula("defined-name XTI first sheet overflow"))?;
            let last = usize::try_from(xti.last_sheet)
                .map_err(|_error| Error::InvalidFormula("defined-name XTI last sheet overflow"))?;
            if last >= removed.len() {
                return Err(Error::InvalidFormula(
                    "defined-name XTI last sheet exceeds the worksheet catalog",
                ));
            }
            Ok((first..=last).all(|sheet| removed[sheet]))
        },
        crate::package::formula::SupportingLink::SameSheet => {
            if (xti.first_sheet, xti.last_sheet) != (-2, -2) {
                return Err(Error::InvalidFormula(
                    "defined-name same-sheet XTI index has invalid endpoints",
                ));
            }
            let Some(scope) = scope else {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal requires a disposition for a workbook-scoped same-sheet defined-name formula",
                ));
            };
            Ok(removed.get(scope).copied().unwrap_or(false))
        },
        crate::package::formula::SupportingLink::ExternalWorkbook(_)
        | crate::package::formula::SupportingLink::AddIn => Ok(false),
    }
}

fn map_named_range_error(error: crate::named_ranges::Error) -> Error {
    match error {
        crate::named_ranges::Error::Wire(error) => Error::Wire(error),
        crate::named_ranges::Error::Formula(error) => map_formula_error(error),
        crate::named_ranges::Error::InvalidLength { expected, found } => {
            Error::InvalidLength { expected, found }
        },
        crate::named_ranges::Error::InvalidFormula(message) => Error::InvalidFormula(message),
    }
}

fn map_formula_error(error: crate::formula::Error) -> Error {
    match error {
        crate::formula::Error::InvalidFormula(message) => Error::InvalidFormula(message),
        crate::formula::Error::InvalidCellReference(message) => {
            Error::InvalidCellReference(message)
        },
        crate::formula::Error::InvalidLength { expected, found } => {
            Error::InvalidLength { expected, found }
        },
        crate::formula::Error::UnsupportedFeature(_) => Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "the structural edit cannot rewrite this formula feature",
        ),
        crate::formula::Error::UnresolvedDependency(dependency) => {
            Error::UnresolvedDependency(dependency)
        },
        crate::formula::Error::Encoding(message) => Error::Encoding(message),
        crate::formula::Error::Wire(error) => Error::Wire(error),
    }
}

fn fallback_visible_position(
    old: usize,
    removed: &[bool],
    records: &[crate::package::records::BundleSheetRecord],
) -> Result<usize> {
    for position in old.saturating_add(1)..removed.len() {
        if !removed[position] && records[position].state == 0 {
            return Ok(position);
        }
    }
    for position in (0..old).rev() {
        if !removed[position] && records[position].state == 0 {
            return Ok(position);
        }
    }
    Err(Error::InvalidFormat(
        "sheet removal leaves no visible tab for BrtBookView",
    ))
}

fn removal_book_view_transitions(
    source: &[u8],
    removed: &[bool],
    records: &[crate::package::records::BundleSheetRecord],
    raw_limits: crate::raw::Limits,
) -> Result<Vec<(usize, usize)>> {
    let mut transitions = Vec::new();
    let mut book_view_count = 0usize;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        if record.kind() != kind::BOOK_VIEW {
            continue;
        }
        book_view_count = book_view_count.saturating_add(1);
        let payload = record.payload();
        if payload.len() != 29 {
            return Err(Error::InvalidLength {
                expected: 29,
                found: payload.len(),
            });
        }
        let old = usize::try_from(u32::from_le_bytes(payload[24..28].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: 28,
                found: payload.len(),
            },
        )?))
        .map_err(|_error| Error::InvalidFormula("BrtBookView index overflow"))?;
        if old >= removed.len() {
            return Err(Error::InvalidFormat(
                "BrtBookView points outside the sheet catalog",
            ));
        }
        if removed[old] {
            let next = fallback_visible_position(old, removed, records)?;
            transitions.push((old, next));
        }
    }
    if book_view_count > 1 && !transitions.is_empty() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "multiple BrtBookView active tabs cannot be relocated during sheet removal",
        ));
    }
    Ok(transitions)
}

fn apply_sheet_removal_view_dependencies(
    source: &[u8],
    package: &mut litchi_opc::OpcPackage,
    workbook_uri: &PackURI,
    records: &[crate::package::records::BundleSheetRecord],
    removed: &[bool],
    tabs: &[crate::sheet::Tab<'_>],
    raw_limits: crate::raw::Limits,
) -> Result<()> {
    let transitions = removal_book_view_transitions(source, removed, records, raw_limits)?;
    let target = if let Some((_, target)) = transitions.first().copied() {
        target
    } else {
        let mut book_view_count = 0usize;
        for item in Records::new(source, raw_limits) {
            if item?.kind() == kind::BOOK_VIEW {
                book_view_count = book_view_count.saturating_add(1);
            }
        }
        if book_view_count != 0 {
            return Ok(());
        }
        let default_source = 0usize;
        if removed.get(default_source).copied().unwrap_or(false) {
            let fallback = fallback_visible_position(default_source, removed, records)?;
            let first_survivor =
                removed
                    .iter()
                    .position(|selected| !*selected)
                    .ok_or(Error::InvalidFormat(
                        "sheet removal leaves no surviving catalog tab",
                    ))?;
            if fallback != first_survivor {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal cannot preserve a hidden default tab without an authored BrtBookView",
                ));
            }
            fallback
        } else {
            if records
                .get(default_source)
                .is_none_or(|record| record.state != 0)
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet removal cannot preserve a hidden default tab without an authored BrtBookView",
                ));
            }
            default_source
        }
    };
    let tab = tabs.get(target).ok_or(Error::InvalidFormat(
        "BrtBookView fallback exceeds the sheet catalog",
    ))?;
    if !matches!(tab.kind(), Kind::Worksheet | Kind::Chart) {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet removal cannot update selected flags for the fallback tab kind",
        ));
    }
    let Some(rel_id) = records[target].rel_id.as_deref() else {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "edit",
            "sheet removal cannot update selected flags for an unbound fallback tab",
        ));
    };
    let target_uri = {
        let workbook_part = package.get_part(workbook_uri)?;
        let relationship = workbook_part
            .rels()
            .get(rel_id)
            .ok_or(Error::InvalidRelationship("sheet relationship is missing"))?;
        if relationship.target_mode() != litchi_opc::TargetMode::Internal {
            return Err(Error::InvalidRelationship(
                "sheet relationship for a fallback tab is external",
            ));
        }
        let target = relationship.target_partname()?;
        package
            .part_metadata(&target)
            .map(|part| part.partname().clone())
            .ok_or(Error::InvalidRelationship(
                "fallback sheet relationship target is absent",
            ))?
    };
    let source_blob = package.get_part(&target_uri)?.blob();
    if let Some(output) = rewrite_sheet_view_flags(source_blob, tab.kind(), true, raw_limits)? {
        package.replace_part_payloads(vec![(target_uri, output)])?;
    }
    Ok(())
}

fn rewrite_extern_sheet_payload_remove(
    payload: &[u8],
    removed: &[bool],
    remaining: &[usize],
    catalog_len: usize,
    context: &crate::package::formula::Context,
) -> Result<Vec<u8>> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let count = usize::try_from(u32::from_le_bytes(payload[..4].try_into().map_err(
        |_error| Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        },
    )?))
    .map_err(|_error| Error::InvalidFormula("BrtExternSheet count overflow"))?;
    let expected = 4usize
        .checked_add(count.checked_mul(12).ok_or(Error::CapacityOverflow {
            resource: "BrtExternSheet payload size",
        })?)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtExternSheet payload size",
        })?;
    if payload.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        });
    }
    let mut output = payload.to_vec();
    for entry in 0..count {
        let offset = 4 + entry * 12;
        let external_link =
            u32::from_le_bytes(payload[offset..offset + 4].try_into().map_err(|_error| {
                Error::InvalidLength {
                    expected: offset + 4,
                    found: payload.len(),
                }
            })?);
        let first = i32::from_le_bytes(payload[offset + 4..offset + 8].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: offset + 8,
                found: payload.len(),
            },
        )?);
        let last = i32::from_le_bytes(payload[offset + 8..offset + 12].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: offset + 12,
                found: payload.len(),
            },
        )?);
        let link = context
            .supporting_links
            .get(usize::try_from(external_link).map_err(|_error| {
                Error::InvalidFormula("BrtExternSheet supporting-link index overflow")
            })?)
            .ok_or(Error::InvalidFormula(
                "BrtExternSheet supporting-link index is unresolved",
            ))?;
        if matches!(link, crate::package::formula::SupportingLink::SameSheet)
            && (first, last) != (-2, -2)
        {
            return Err(Error::InvalidFormula(
                "same-sheet BrtExternSheet has invalid endpoints",
            ));
        }
        if !matches!(link, crate::package::formula::SupportingLink::SelfWorkbook) {
            continue;
        }
        if (first, last) == (-2, -2) || (first, last) == (-1, -1) {
            continue;
        }
        if first < 0 || last < first {
            return Err(Error::InvalidFormula(
                "internal BrtExternSheet range has invalid endpoints",
            ));
        }
        let first = usize::try_from(first).map_err(|_error| {
            Error::InvalidFormula("internal BrtExternSheet first endpoint overflow")
        })?;
        let last = usize::try_from(last).map_err(|_error| {
            Error::InvalidFormula("internal BrtExternSheet last endpoint overflow")
        })?;
        if last >= catalog_len || last >= removed.len() {
            return Err(Error::InvalidFormula(
                "internal BrtExternSheet range exceeds the sheet catalog",
            ));
        }
        let first_remaining = (first..=last).find(|position| !removed[*position]);
        let Some(first_remaining) = first_remaining else {
            output[offset + 4..offset + 8].copy_from_slice(&(-1_i32).to_le_bytes());
            output[offset + 8..offset + 12].copy_from_slice(&(-1_i32).to_le_bytes());
            continue;
        };
        let last_remaining = (first..=last)
            .rev()
            .find(|position| !removed[*position])
            .ok_or(Error::InvalidFormula("internal XTI range has no endpoint"))?;
        let new_first = remaining[first_remaining];
        let new_last = remaining[last_remaining];
        output[offset + 4..offset + 8].copy_from_slice(
            &i32::try_from(new_first)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "removed BrtExternSheet first endpoint",
                })?
                .to_le_bytes(),
        );
        output[offset + 8..offset + 12].copy_from_slice(
            &i32::try_from(new_last)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "removed BrtExternSheet last endpoint",
                })?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_volatile_dependencies_remove(
    source: &[u8],
    removed: &[bool],
    remaining: &[usize],
    catalog_len: usize,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Vec<u8>>> {
    let mut output = source.to_vec();
    let mut changed = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        if record.kind() != kind::VOL_REF {
            continue;
        }
        if record.payload().len() != 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: record.payload().len(),
            });
        }
        let old = usize::try_from(u32::from_le_bytes(
            record.payload()[8..12]
                .try_into()
                .map_err(|_error| Error::InvalidLength {
                    expected: 12,
                    found: record.payload().len(),
                })?,
        ))
        .map_err(|_error| Error::InvalidFormula("BrtVolRef sheet index overflow"))?;
        if old >= catalog_len || old >= removed.len() {
            return Err(Error::InvalidFormat(
                "BrtVolRef sheet index exceeds the sheet catalog",
            ));
        }
        if removed[old] {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet removal cannot safely delete a volatile dependency reference",
            ));
        }
        let new = remaining[old];
        let (_, header_len) = Header::parse(
            source
                .get(record.offset()..)
                .ok_or(Error::InvalidFormat("BrtVolRef offset is outside its part"))?,
            raw_limits,
        )?;
        output[record.offset() + header_len + 8..record.offset() + header_len + 12]
            .copy_from_slice(
                &u32::try_from(new)
                    .map_err(|_error| Error::CapacityOverflow {
                        resource: "removed BrtVolRef sheet index",
                    })?
                    .to_le_bytes(),
            );
        changed |= new != old;
    }
    Ok(changed.then_some(output))
}

/// Apply one complete-catalog permutation to the workbook and every modeled
/// sheet-index dependency.  The package is edited through a replacement
/// batch, so a refusal from any dependency leaves the source image untouched.
fn apply_sheet_order(workbook: &mut Workbook, order: &[usize]) -> Result<bool> {
    let catalog_len = workbook.sheets().len();
    let inverse = validate_sheet_order(order, catalog_len)?;
    if order
        .iter()
        .enumerate()
        .all(|(position, source)| position == *source)
    {
        return Ok(false);
    }
    reject_unmodeled_reorder_dependencies(&workbook.package)?;
    crate::signature_policy::ensure_unsigned(&workbook.package)?;

    let workbook_uri = workbook.package.main_document_part()?.partname().clone();
    let volatile = crate::volatile_dependencies::reorder_owner(&workbook.package, catalog_len)?;
    let workbook_source = workbook.package.get_part(&workbook_uri)?.blob().to_vec();
    let raw_limits = workbook.raw_limits_value();
    let workbook_output = rewrite_workbook_order(
        &workbook_source,
        order,
        &inverse,
        catalog_len,
        &workbook.formula_context,
        raw_limits,
    )?;

    let tabs = workbook.sheets().collect::<Vec<_>>();
    let titles = tabs
        .iter()
        .map(|tab| tab.name().to_owned())
        .collect::<Vec<_>>();
    let title_refs = titles.iter().map(String::as_str).collect::<Vec<_>>();
    let mut replacements = Vec::new();
    replacements
        .try_reserve(workbook.package.part_count().saturating_add(2))
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-order package replacements",
            source,
        })?;
    replacements.push((workbook_uri, workbook_output));

    for part in workbook.package.iter_parts() {
        if part
            .content_type()
            .eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES)
        {
            if let Some(output) = litchi_ooxml_common::xml::sheet_name::reorder_titles(
                workbook.package.get_part(part.partname())?.blob(),
                &title_refs,
                order,
                excel_name_eq,
            )? {
                replacements.push((part.partname().clone(), output));
            }
        }
        if volatile.as_ref().is_some_and(|uri| uri == part.partname()) {
            if let Some(output) = rewrite_volatile_dependencies(
                workbook.package.get_part(part.partname())?.blob(),
                &inverse,
                catalog_len,
                raw_limits,
            )? {
                replacements.push((part.partname().clone(), output));
            }
        }
        if part
            .content_type()
            .eq_ignore_ascii_case(crate::shared_workbook::package::HEADERS_CONTENT_TYPE)
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reordering is unsupported with shared-workbook revision headers because their historical display order is immutable",
            ));
        }
    }

    let mut package = workbook.package.clone();
    package.replace_part_payloads(replacements)?;
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
    Ok(true)
}

fn validate_sheet_order(order: &[usize], catalog_len: usize) -> Result<Vec<usize>> {
    if order.len() != catalog_len {
        return Err(Error::InvalidFormat(
            "sheet order has entries but the catalog",
        ));
    }
    let mut inverse = Vec::new();
    inverse
        .try_reserve_exact(catalog_len)
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-order inverse permutation",
            source,
        })?;
    inverse.resize(catalog_len, usize::MAX);
    for (new_position, &old_position) in order.iter().enumerate() {
        let slot = inverse.get_mut(old_position).ok_or(Error::InvalidFormat(
            "sheet order contains an out-of-range position",
        ))?;
        if *slot != usize::MAX {
            return Err(Error::InvalidFormat(
                "sheet order contains a duplicate position",
            ));
        }
        *slot = new_position;
    }
    if inverse.iter().any(|position| *position == usize::MAX) {
        return Err(Error::InvalidFormat("sheet order omits a catalog position"));
    }
    Ok(inverse)
}

fn reject_unmodeled_reorder_dependencies(package: &litchi_opc::OpcPackage) -> Result<()> {
    for part in package.iter_parts() {
        let content_type = part.content_type().to_ascii_lowercase();
        let part_name = part.partname().as_str().to_ascii_lowercase();
        if content_type.eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_VBA_PROJECT)
            || content_type.eq_ignore_ascii_case(
                litchi_opc::constants::content_type::OFC_VBA_PROJECT_SIGNATURE,
            )
            || content_type.eq_ignore_ascii_case(
                litchi_opc::constants::content_type::OFC_VBA_PROJECT_SIGNATURE_AGILE,
            )
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reordering is unsupported while a VBA project dependency is present",
            ));
        }
        if content_type.contains("calcchain") || part_name.contains("calcchain") {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reordering is unsupported while a calculation-chain dependency is present",
            ));
        }
        if content_type.eq_ignore_ascii_case(crate::slicer::package::CACHE_CONTENT_TYPE)
            || content_type.eq_ignore_ascii_case(crate::slicer::package::VIEWS_CONTENT_TYPE)
            || content_type == "application/vnd.ms-excel.timelinecache+xml"
            || content_type == "application/vnd.ms-excel.timeline+xml"
            || content_type.eq_ignore_ascii_case(crate::shared_workbook::package::LOG_CONTENT_TYPE)
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reordering is unsupported with unmodeled slicer, timeline or revision-log dependencies",
            ));
        }
        if content_type.eq_ignore_ascii_case(crate::data_model::DATA_MODEL_CONTENT_TYPE)
            || content_type.eq_ignore_ascii_case(
                crate::package::connections::package::CONNECTIONS_CONTENT_TYPE,
            )
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reordering is unsupported with unmodeled data-model or connections dependencies",
            ));
        }
    }
    Ok(())
}

fn reject_shared_workbook_headers(package: &litchi_opc::OpcPackage) -> Result<()> {
    if package.iter_parts().any(|part| {
        part.content_type()
            .eq_ignore_ascii_case(crate::shared_workbook::package::HEADERS_CONTENT_TYPE)
    }) {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "sheet insertion is unsupported with shared-workbook revision headers because their historical display order is immutable",
        ));
    }
    Ok(())
}

fn rewrite_workbook_order(
    source: &[u8],
    order: &[usize],
    inverse: &[usize],
    catalog_len: usize,
    context: &crate::package::formula::Context,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let mut bundle_records = Vec::new();
    bundle_records
        .try_reserve_exact(catalog_len)
        .map_err(|source| Error::Allocation {
            resource: "XLSB reordered catalog records",
            source,
        })?;
    let mut saw_bundle_collection = false;
    let mut bundle_open = false;
    let mut bundle_collection_closed = false;
    let mut records = Vec::new();
    for item in Records::new(source, raw_limits) {
        let record = item?;
        let raw = record_bytes(source, &record, raw_limits)?;
        match record.kind() {
            kind::BEGIN_BUNDLE_SHS => {
                if bundle_open || saw_bundle_collection {
                    return Err(Error::InvalidFormat(
                        "Workbook has duplicate or nested BrtBundleShs collections",
                    ));
                }
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                saw_bundle_collection = true;
                bundle_open = true;
            },
            kind::BUNDLE_SH if bundle_open => {
                bundle_records.push(raw);
            },
            kind::BUNDLE_SH => {
                return Err(Error::InvalidFormat(
                    "BrtBundleSh appears outside its BrtBeginBundleShs collection",
                ));
            },
            kind::END_BUNDLE_SHS => {
                if !bundle_open {
                    return Err(Error::InvalidFormat(
                        "BrtEndBundleShs has no matching collection",
                    ));
                }
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                bundle_open = false;
                bundle_collection_closed = true;
            },
            _ if bundle_open => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet reordering cannot detach an unknown record interspersed with BrtBundleSh records",
                ));
            },
            kind::BEGIN_BOOK
            | kind::END_BOOK
            | kind::FILE_VERSION
            | kind::WORKBOOK_PROP
            | kind::BEGIN_BOOK_VIEWS
            | kind::END_BOOK_VIEWS
            | kind::BOOK_VIEW
            | kind::BEGIN_PIVOT_CACHE_IDS
            | kind::END_PIVOT_CACHE_IDS
            | kind::BEGIN_PIVOT_CACHE_ID
            | kind::END_PIVOT_CACHE_ID
            | kind::BEGIN_EXTERNALS
            | kind::END_EXTERNALS
            | kind::SUP_SELF
            | kind::SUP_SAME
            | kind::SUP_ADDIN
            | kind::SUP_BOOK_SRC
            | kind::EXTERN_SHEET
            | kind::NAME
            | kind::CALC_PROP => {},
            _other => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "sheet reordering cannot preserve unmodeled workbook record at offset without a catalog dependency contract",
                ));
            },
        }
        records.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "XLSB workbook record spans",
            source,
        })?;
        records.push((record.kind(), raw));
    }
    if bundle_open || !bundle_collection_closed || !saw_bundle_collection {
        return Err(Error::InvalidFormat(
            "Workbook sheet catalog has no complete BrtBundleSh collection",
        ));
    }
    if bundle_records.len() != catalog_len {
        return Err(Error::InvalidFormat(
            "Workbook BrtBundleSh count differs from catalog size",
        ));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "reordered workbook.bin",
            source,
        })?;
    let mut inserted_bundles = false;
    for (record_kind, raw) in records {
        if record_kind == kind::BUNDLE_SH {
            if !inserted_bundles {
                for &old_position in order {
                    output.extend_from_slice(bundle_records.get(old_position).ok_or(
                        Error::InvalidFormat("sheet order references an absent BrtBundleSh record"),
                    )?);
                }
                inserted_bundles = true;
            }
            continue;
        }
        let raw = rewrite_workbook_record(raw, inverse, catalog_len, context, raw_limits)?;
        output.extend_from_slice(&raw);
    }
    if !inserted_bundles {
        return Err(Error::InvalidFormat(
            "Workbook sheet catalog records were not emitted",
        ));
    }
    Ok(output)
}

fn rewrite_workbook_record(
    raw: &[u8],
    inverse: &[usize],
    catalog_len: usize,
    context: &crate::package::formula::Context,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let (header, header_len) = Header::parse(raw, raw_limits)?;
    let payload = raw.get(header_len..).ok_or(Error::InvalidLength {
        expected: header_len,
        found: raw.len(),
    })?;
    let payload = payload.get(..header.len()).ok_or(Error::InvalidLength {
        expected: header_len.saturating_add(header.len()),
        found: raw.len(),
    })?;
    let mut output = raw.to_vec();
    match header.kind() {
        kind::BOOK_VIEW => {
            if payload.len() != 29 {
                return Err(Error::InvalidLength {
                    expected: 29,
                    found: payload.len(),
                });
            }
            for field_offset in [20usize, 24usize] {
                let old = usize::try_from(u32::from_le_bytes(
                    payload[field_offset..field_offset + 4]
                        .try_into()
                        .map_err(|_error| Error::InvalidLength {
                            expected: field_offset.saturating_add(4),
                            found: payload.len(),
                        })?,
                ))
                .map_err(|_error| Error::InvalidFormula("BrtBookView index overflow"))?;
                let new = *inverse.get(old).ok_or(Error::InvalidFormat(
                    "BrtBookView points outside the sheet catalog",
                ))?;
                let output_len = output.len();
                let target = output
                    .get_mut(header_len + field_offset..header_len + field_offset + 4)
                    .ok_or(Error::InvalidLength {
                        expected: header_len + field_offset + 4,
                        found: output_len,
                    })?;
                target.copy_from_slice(
                    &u32::try_from(new)
                        .map_err(|_error| Error::CapacityOverflow {
                            resource: "reordered workbook view index",
                        })?
                        .to_le_bytes(),
                );
            }
        },
        kind::NAME => {
            if payload.len() < 9 {
                return Err(Error::InvalidLength {
                    expected: 9,
                    found: payload.len(),
                });
            }
            let old = u32::from_le_bytes(payload[5..9].try_into().map_err(|_error| {
                Error::InvalidLength {
                    expected: 9,
                    found: payload.len(),
                }
            })?);
            if old != u32::MAX {
                let old = usize::try_from(old).map_err(|_error| {
                    Error::InvalidFormula("BrtName sheet scope does not fit usize")
                })?;
                let new = *inverse.get(old).ok_or(Error::InvalidFormula(
                    "BrtName sheet scope exceeds the sheet catalog",
                ))?;
                output[header_len + 5..header_len + 9].copy_from_slice(
                    &u32::try_from(new)
                        .map_err(|_error| Error::CapacityOverflow {
                            resource: "reordered BrtName sheet scope",
                        })?
                        .to_le_bytes(),
                );
            }
        },
        kind::EXTERN_SHEET => {
            let rewritten = rewrite_extern_sheet_payload(payload, inverse, catalog_len, context)?;
            if rewritten != payload {
                output = replace_record_payload(&output, header_len, header.len(), &rewritten)?;
            }
        },
        _ => {},
    }
    Ok(output)
}

fn rewrite_extern_sheet_payload(
    payload: &[u8],
    inverse: &[usize],
    catalog_len: usize,
    context: &crate::package::formula::Context,
) -> Result<Vec<u8>> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let count = usize::try_from(u32::from_le_bytes(payload[..4].try_into().map_err(
        |_error| Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        },
    )?))
    .map_err(|_error| Error::InvalidFormula("BrtExternSheet count overflow"))?;
    let expected = 4usize
        .checked_add(count.checked_mul(12).ok_or(Error::CapacityOverflow {
            resource: "BrtExternSheet payload size",
        })?)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtExternSheet payload size",
        })?;
    if payload.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        });
    }
    let mut output = payload.to_vec();
    for entry in 0..count {
        let offset = 4 + entry * 12;
        let external_link =
            u32::from_le_bytes(payload[offset..offset + 4].try_into().map_err(|_error| {
                Error::InvalidLength {
                    expected: offset + 4,
                    found: payload.len(),
                }
            })?);
        let first = i32::from_le_bytes(payload[offset + 4..offset + 8].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: offset + 8,
                found: payload.len(),
            },
        )?);
        let last = i32::from_le_bytes(payload[offset + 8..offset + 12].try_into().map_err(
            |_error| Error::InvalidLength {
                expected: offset + 12,
                found: payload.len(),
            },
        )?);
        let link = context
            .supporting_links
            .get(usize::try_from(external_link).map_err(|_error| {
                Error::InvalidFormula("BrtExternSheet supporting-link index overflow")
            })?)
            .ok_or(Error::InvalidFormula(
                "BrtExternSheet supporting-link index is unresolved",
            ))?;
        match link {
            crate::package::formula::SupportingLink::SelfWorkbook => {},
            crate::package::formula::SupportingLink::SameSheet => {
                if (first, last) != (-2, -2) {
                    return Err(Error::InvalidFormula(
                        "same-sheet BrtExternSheet has invalid endpoints",
                    ));
                }
                continue;
            },
            crate::package::formula::SupportingLink::ExternalWorkbook(_)
            | crate::package::formula::SupportingLink::AddIn => continue,
        }
        if (first, last) == (-2, -2) || (first, last) == (-1, -1) {
            continue;
        }
        if first < 0 || last < first {
            return Err(Error::InvalidFormula(
                "internal BrtExternSheet range has invalid endpoints",
            ));
        }
        let first = usize::try_from(first).map_err(|_error| {
            Error::InvalidFormula("internal BrtExternSheet first endpoint overflow")
        })?;
        let last = usize::try_from(last).map_err(|_error| {
            Error::InvalidFormula("internal BrtExternSheet last endpoint overflow")
        })?;
        if last >= catalog_len {
            return Err(Error::InvalidFormula(
                "internal BrtExternSheet range exceeds the sheet catalog",
            ));
        }
        let span = last
            .checked_sub(first)
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::CapacityOverflow {
                resource: "internal BrtExternSheet range",
            })?;
        let mut minimum = usize::MAX;
        let mut maximum = 0usize;
        for old_position in first..=last {
            let new_position = *inverse.get(old_position).ok_or(Error::InvalidFormula(
                "internal BrtExternSheet endpoint is unmapped",
            ))?;
            minimum = minimum.min(new_position);
            maximum = maximum.max(new_position);
        }
        if maximum.saturating_sub(minimum).saturating_add(1) != span {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "sheet reorder would make an internal BrtExternSheet 3-D range non-contiguous",
            ));
        }
        output[offset + 4..offset + 8].copy_from_slice(
            &i32::try_from(minimum)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "reordered BrtExternSheet first endpoint",
                })?
                .to_le_bytes(),
        );
        output[offset + 8..offset + 12].copy_from_slice(
            &i32::try_from(maximum)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "reordered BrtExternSheet last endpoint",
                })?
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn rewrite_volatile_dependencies(
    source: &[u8],
    inverse: &[usize],
    catalog_len: usize,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Vec<u8>>> {
    let mut output = source.to_vec();
    let mut changed = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        if record.kind() != kind::VOL_REF {
            continue;
        }
        if record.payload().len() != 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: record.payload().len(),
            });
        }
        let old = usize::try_from(u32::from_le_bytes(
            record.payload()[8..12]
                .try_into()
                .map_err(|_error| Error::InvalidLength {
                    expected: 12,
                    found: record.payload().len(),
                })?,
        ))
        .map_err(|_error| Error::InvalidFormula("BrtVolRef sheet index overflow"))?;
        if old >= catalog_len {
            return Err(Error::InvalidFormat(
                "BrtVolRef sheet index exceeds the sheet catalog",
            ));
        }
        let new = *inverse
            .get(old)
            .ok_or(Error::InvalidFormat("BrtVolRef sheet index is unmapped"))?;
        let (_, header_len) = Header::parse(
            source
                .get(record.offset()..)
                .ok_or(Error::InvalidFormat("BrtVolRef offset is outside its part"))?,
            raw_limits,
        )?;
        output[record.offset() + header_len + 8..record.offset() + header_len + 12]
            .copy_from_slice(
                &u32::try_from(new)
                    .map_err(|_error| Error::CapacityOverflow {
                        resource: "reordered BrtVolRef sheet index",
                    })?
                    .to_le_bytes(),
            );
        changed |= new != old;
    }
    Ok(changed.then_some(output))
}

fn replace_record_payload(
    raw: &[u8],
    header_len: usize,
    payload_len: usize,
    payload: &[u8],
) -> Result<Vec<u8>> {
    if payload.len() != payload_len {
        return Err(Error::InvalidLength {
            expected: payload_len,
            found: payload.len(),
        });
    }
    let end = header_len
        .checked_add(payload_len)
        .ok_or(Error::CapacityOverflow {
            resource: "BIFF12 record payload range",
        })?;
    if raw.len() < end {
        return Err(Error::InvalidLength {
            expected: end,
            found: raw.len(),
        });
    }
    let mut output = raw.to_vec();
    output[header_len..end].copy_from_slice(payload);
    Ok(output)
}

fn record_bytes<'a>(
    source: &'a [u8],
    record: &crate::raw::Record<'_>,
    raw_limits: crate::raw::Limits,
) -> Result<&'a [u8]> {
    let record_source = source
        .get(record.offset()..)
        .ok_or(Error::InvalidFormat("record offset is outside its part"))?;
    let (_, header_len) = Header::parse(record_source, raw_limits)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "workbook record range",
        })?;
    source
        .get(record.offset()..end)
        .ok_or(Error::InvalidFormat("record range is outside its part"))
}

fn apply_sheet_metadata(workbook: &mut Workbook, operations: &[Operation]) -> Result<(bool, bool)> {
    let mut names = BTreeMap::<usize, &str>::new();
    let mut dependency_positions = std::collections::BTreeSet::<usize>::new();
    let mut states = BTreeMap::<usize, State>::new();
    let mut changed = false;
    let mut state_changed = false;
    let tabs = workbook.sheets().collect::<Vec<_>>();
    for operation in operations {
        match operation {
            Operation::RenameSheet { sheet, name } => {
                let (catalog_position, tab) = tabs
                    .iter()
                    .enumerate()
                    .find(|(_, tab)| {
                        tab.worksheet_position()
                            .is_some_and(|position| position.get() == *sheet)
                    })
                    .ok_or(Error::WorksheetNotFound(Some(*sheet)))?;
                if tab.name() != name {
                    validate_sheet_name_value(name)?;
                }
                if tab.name() != name {
                    names.insert(catalog_position, name.as_str());
                    changed = true;
                }
            },
            Operation::SetSheetName {
                catalog_position,
                name,
                dependencies,
            } => {
                let tab = tabs
                    .get(*catalog_position)
                    .ok_or(Error::WorksheetNotFound(Some(*catalog_position)))?;
                // A no-op must remain an exact source no-op even when a
                // leniently opened package retained a malformed authored
                // name.  Changed names are validated below before any
                // package payload is touched.
                if tab.name() != name {
                    validate_sheet_name_value(name)?;
                    names.insert(*catalog_position, name.as_str());
                    if *dependencies {
                        dependency_positions.insert(*catalog_position);
                    }
                    changed = true;
                }
            },
            Operation::SetSheetState {
                catalog_position,
                state,
            } => {
                let tab = tabs
                    .get(*catalog_position)
                    .ok_or(Error::WorksheetNotFound(Some(*catalog_position)))?;
                if tab.kind() == Kind::Unbound && *state != State::VeryHidden {
                    return Err(Error::InvalidRelationship(
                        "unbound XLSB sheet tabs must be very hidden",
                    ));
                }
                if tab.state() != *state {
                    states.insert(*catalog_position, *state);
                    changed = true;
                    state_changed = true;
                }
            },
            Operation::TransferCell { .. }
            | Operation::ReorderSheets { .. }
            | Operation::InsertSheet { .. }
            | Operation::RemoveSheets { .. }
            | Operation::AddImage { .. }
            | Operation::TransferDrawing { .. }
            | Operation::SetMergedRanges { .. }
            | Operation::SetDefinedName { .. } => {
                return Err(Error::InvalidFormat(
                    "workbook metadata operation grouping is internally inconsistent",
                ));
            },
        }
    }
    validate_final_sheet_name_changes(&tabs, &names)?;
    if !changed {
        return Ok((false, false));
    }
    let kinds = tabs.iter().map(|tab| tab.kind()).collect::<Vec<_>>();
    let old_sheet_titles = if dependency_positions.is_empty() {
        Vec::new()
    } else {
        tabs.iter()
            .map(|tab| tab.name().to_owned())
            .collect::<Vec<_>>()
    };
    let dependency_renames = dependency_positions
        .iter()
        .filter_map(|catalog_position| {
            names.get(catalog_position).map(|name| {
                (
                    old_sheet_titles[*catalog_position].clone(),
                    (*name).to_owned(),
                )
            })
        })
        .collect::<Vec<_>>();
    let mut final_states = tabs.iter().map(|tab| tab.state()).collect::<Vec<_>>();
    for (catalog_position, state) in &states {
        let final_state = final_states
            .get_mut(*catalog_position)
            .ok_or(Error::WorksheetNotFound(Some(*catalog_position)))?;
        *final_state = *state;
    }
    drop(tabs);
    crate::signature_policy::ensure_unsigned(&workbook.package)?;
    let uri = workbook.package.main_document_part()?.partname().clone();
    let source = workbook.package.get_part(&uri)?.blob().to_vec();
    let raw_limits = workbook.raw_limits_value();
    if names.is_empty() {
        let mut output = source.clone();
        let mut index = 0usize;
        let mut replaced = false;
        for item in Records::new(&source, raw_limits) {
            let record = item?;
            if record.kind() == kind::BUNDLE_SH {
                if let Some(state) = states.get(&index).copied() {
                    let header_source = source.get(record.offset()..).ok_or(
                        Error::InvalidFormat("record offset is outside workbook.bin"),
                    )?;
                    let (_, header_len) = Header::parse(header_source, raw_limits)?;
                    let payload_start =
                        record
                            .offset()
                            .checked_add(header_len)
                            .ok_or(Error::CapacityOverflow {
                                resource: "bundle-sheet state offset",
                            })?;
                    let payload_end = payload_start.saturating_add(4);
                    let output_len = output.len();
                    let state_bytes =
                        output
                            .get_mut(payload_start..payload_end)
                            .ok_or(Error::InvalidLength {
                                expected: payload_end,
                                found: output_len,
                            })?;
                    state_bytes.copy_from_slice(&u32::from(state_to_wire(state)).to_le_bytes());
                    replaced = true;
                }
                index = index.saturating_add(1);
            }
        }
        if !replaced {
            return Err(Error::InvalidFormat("the sheet state record was not"));
        }
        let mut package = workbook.package.clone();
        package.get_part_mut(&uri)?.set_blob(output);
        apply_sheet_view_dependencies(workbook, &mut package, &kinds, &final_states, &states)?;
        let validated = workbook.reparse_candidate(package)?;
        *workbook = validated;
        return Ok((true, state_changed));
    }
    let mut output = Vec::with_capacity(source.len());
    let mut index = 0usize;
    let mut replaced = false;
    for item in Records::new(&source, raw_limits) {
        let record = item?;
        if record.kind() == kind::BUNDLE_SH {
            let name = names.get(&index).copied();
            let state = states.get(&index).copied();
            if name.is_some() || state.is_some() {
                let mut payload = record.payload().to_vec();
                if let Some(name) = name {
                    payload = renamed_bundle_payload(&payload, name, raw_limits)?;
                }
                if let Some(state) = state {
                    let payload_len = payload.len();
                    let state_bytes = payload.get_mut(..4).ok_or(Error::InvalidLength {
                        expected: 4,
                        found: payload_len,
                    })?;
                    state_bytes.copy_from_slice(&u32::from(state_to_wire(state)).to_le_bytes());
                }
                Encoder::new(&mut output, workbook.raw_limits_value())
                    .write_record(kind::BUNDLE_SH, &payload)?;
                replaced = true;
            } else {
                copy_record(&source, &record, raw_limits, &mut output)?;
            }
            index = index.saturating_add(1);
        } else {
            copy_record(&source, &record, raw_limits, &mut output)?;
        }
    }
    if !replaced {
        return Err(Error::InvalidFormat("the sheet metadata record was not"));
    }
    // Parse every dependent XML part and stage all successful replacements
    // before changing the candidate package.  A dependency refusal therefore
    // cannot leave a partially edited workbook image behind.
    let mut package = workbook.package.clone();
    let dependency_replacements = if dependency_renames.is_empty() {
        Vec::new()
    } else {
        collect_sheet_name_dependency_replacements(
            &package,
            &dependency_renames,
            &old_sheet_titles,
        )?
    };
    let mut replacements = dependency_replacements;
    replacements.push((uri, output));
    package.replace_part_payloads(replacements)?;
    apply_sheet_view_dependencies(workbook, &mut package, &kinds, &final_states, &states)?;
    let validated = workbook.reparse_candidate(package)?;
    *workbook = validated;
    Ok((true, state_changed))
}

/// Stage rewrites for every package part whose declared content type is XML.
///
/// XLSB workbooks commonly carry chart, custom-XML, and extended-property
/// members beside their binary workbook part.  The content type is the only
/// authoritative discriminator here: a URI ending in `.xml` is not enough,
/// and an XML part is not required to use that suffix.  The helper performs a
/// namespace-aware lexical rewrite and returns typed refusals for formula-like
/// carriers it cannot prove safe to edit.
fn collect_sheet_name_dependency_replacements(
    package: &litchi_opc::OpcPackage,
    dependency_renames: &[(String, String)],
    old_sheet_titles: &[String],
) -> Result<Vec<(PackURI, Vec<u8>)>> {
    if dependency_renames.is_empty() {
        return Ok(Vec::new());
    }
    let renames = dependency_renames
        .iter()
        .map(|(before, after)| litchi_ooxml_common::xml::sheet_name::Rename { before, after })
        .collect::<Vec<_>>();
    let sheet_titles = old_sheet_titles
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    let mut xml_parts = Vec::new();
    xml_parts
        .try_reserve(package.part_count())
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-name XML part inventory",
            source,
        })?;
    for part in package.iter_parts() {
        if (is_xml_content_type(part.content_type())
            || part
                .content_type()
                .eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_VML_DRAWING))
            && !part
                .content_type()
                .eq_ignore_ascii_case(litchi_opc::constants::content_type::SML_EXTERNAL_LINK)
            && !part.partname().as_str().starts_with("/xl/externalLinks/")
        {
            xml_parts.push(part.partname().clone());
        }
    }
    xml_parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut replacements = Vec::new();
    replacements
        .try_reserve(xml_parts.len().saturating_add(1))
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet-name XML replacements",
            source,
        })?;
    for partname in xml_parts {
        let part = package.get_part(&partname)?;
        let Some(payload) = litchi_ooxml_common::xml::sheet_name::rewrite(
            part.blob(),
            &renames,
            if part
                .content_type()
                .eq_ignore_ascii_case(litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES)
            {
                &sheet_titles
            } else {
                &[]
            },
            excel_name_eq,
        )?
        else {
            continue;
        };
        replacements.push((partname, payload));
    }
    Ok(replacements)
}

fn is_xml_content_type(content_type: &str) -> bool {
    let bytes = content_type.as_bytes();
    content_type.eq_ignore_ascii_case("application/xml")
        || content_type.eq_ignore_ascii_case("text/xml")
        || (bytes.len() > 4 && bytes[bytes.len() - 4..].eq_ignore_ascii_case(b"+xml"))
}

fn apply_sheet_view_dependencies(
    workbook: &Workbook,
    package: &mut litchi_opc::OpcPackage,
    kinds: &[Kind],
    final_states: &[State],
    changed_states: &BTreeMap<usize, State>,
) -> Result<()> {
    if changed_states.is_empty() {
        return Ok(());
    }
    if kinds.len() != final_states.len() {
        return Err(Error::InvalidFormat(
            "sheet catalog and visibility state counts differ",
        ));
    }
    let workbook_uri = workbook.package.main_document_part()?.partname().clone();
    let workbook_source = package.get_part(&workbook_uri)?.blob().to_vec();
    let workbook_part = workbook.package.get_part(&workbook_uri)?;
    let mut workbook_output = workbook_source.clone();
    let raw_limits = workbook.raw_limits_value();
    let transitions = rewrite_book_views(&mut workbook_output, final_states, raw_limits)?;
    let mut relocated_target = None;
    for (_, new_active) in &transitions {
        if let Some(previous) = relocated_target
            && previous != *new_active
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "multiple BrtBookView active tabs cannot be mapped to one selected view",
            ));
        }
        relocated_target = Some(*new_active);
    }
    let old_active = workbook.active_catalog_position().unwrap_or(0);
    if old_active >= final_states.len() {
        return Err(Error::InvalidFormat(
            "active workbook view points outside the sheet catalog",
        ));
    }
    if transitions.is_empty() && final_states[old_active].is_hidden() {
        return Err(Error::InvalidFormat(
            "active workbook view has no relocatable BrtBookView",
        ));
    }

    let mut actions = BTreeMap::<usize, bool>::new();
    for catalog_position in changed_states.keys().copied() {
        if final_states[catalog_position] != State::Visible {
            actions.insert(catalog_position, false);
        }
    }
    for (old_active, new_active) in &transitions {
        if old_active != new_active {
            actions.insert(*old_active, false);
            actions.insert(*new_active, true);
        }
    }

    for catalog_position in actions.keys() {
        match kinds[*catalog_position] {
            Kind::Worksheet | Kind::Chart => {},
            Kind::Dialog | Kind::Macro | Kind::IntlMacro | Kind::Unbound => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "cannot update the selected view for sheet tab",
                ));
            },
        }
    }

    let mut sheet_parts = vec![None; kinds.len()];
    let mut catalog_position = 0usize;
    let raw_limits = workbook.raw_limits_value();
    for item in Records::new(&workbook_source, raw_limits) {
        let record = item?;
        if record.kind() != kind::BUNDLE_SH {
            continue;
        }
        let bundle =
            crate::package::records::BundleSheetRecord::parse(record.payload(), raw_limits)?;
        if let Some(rel_id) = bundle.rel_id.as_deref() {
            let relationship = workbook_part
                .rels()
                .get(rel_id)
                .ok_or(Error::InvalidRelationship("sheet relationship is missing"))?;
            if relationship.is_external() {
                return Err(Error::InvalidRelationship("sheet relationship is external"));
            }
            sheet_parts[catalog_position] = Some(relationship.target_partname()?);
        }
        catalog_position = catalog_position.saturating_add(1);
    }
    if catalog_position != kinds.len() {
        return Err(Error::InvalidFormat(
            "workbook sheet records do not match the sheet catalog",
        ));
    }

    let replacement_capacity = actions
        .len()
        .checked_add(if workbook_output != workbook_source {
            1
        } else {
            0
        })
        .ok_or(Error::CapacityOverflow {
            resource: "XLSB sheet view payload replacements",
        })?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(replacement_capacity)
        .map_err(|source| Error::Allocation {
            resource: "XLSB sheet view payload replacements",
            source,
        })?;
    if workbook_output != workbook_source {
        replacements.push((workbook_uri, workbook_output));
    }
    for (catalog_position, selected) in actions {
        let Some(partname) = sheet_parts[catalog_position].as_ref() else {
            continue;
        };
        let part_kind = kinds[catalog_position];
        if !matches!(part_kind, Kind::Worksheet | Kind::Chart) {
            continue;
        }
        let source = package.get_part(partname)?.blob().to_vec();
        let Some(output) = rewrite_sheet_view_flags(&source, part_kind, selected, raw_limits)?
        else {
            continue;
        };
        replacements.push((partname.clone(), output));
    }
    package.replace_part_payloads(replacements)?;
    Ok(())
}

fn rewrite_book_views(
    output: &mut [u8],
    final_states: &[State],
    raw_limits: crate::raw::Limits,
) -> Result<Vec<(usize, usize)>> {
    if final_states.is_empty() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "workbook must retain at least one visible sheet tab",
        ));
    }
    let source = output.to_vec();
    let mut transitions = Vec::new();
    let mut saw_book_view = false;
    for item in Records::new(&source, raw_limits) {
        let record = item?;
        if record.kind() != kind::BOOK_VIEW {
            continue;
        }
        saw_book_view = true;
        let payload = record.payload();
        let current_bytes = payload.get(24..28).ok_or(Error::InvalidLength {
            expected: 29,
            found: payload.len(),
        })?;
        let current = usize::try_from(u32::from_le_bytes([
            current_bytes[0],
            current_bytes[1],
            current_bytes[2],
            current_bytes[3],
        ]))
        .map_err(|_error| Error::InvalidFormula("BrtBookView itabCur does not fit usize"))?;
        if current >= final_states.len() {
            return Err(Error::InvalidFormat(
                "BrtBookView itabCur points outside the sheet catalog",
            ));
        }
        let next = if final_states[current] == State::Visible {
            current
        } else {
            next_visible_sheet(current, final_states).ok_or_else(|| {
                Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit",
                    "workbook must retain at least one visible sheet tab",
                )
            })?
        };
        if current == next {
            continue;
        }
        let header_source = source.get(record.offset()..).ok_or(Error::InvalidFormat(
            "book-view record offset is outside workbook.bin",
        ))?;
        let (_, header_len) = Header::parse(header_source, raw_limits)?;
        let payload_start = record
            .offset()
            .checked_add(header_len)
            .and_then(|offset| offset.checked_add(24))
            .ok_or(Error::CapacityOverflow {
                resource: "BrtBookView itabCur offset",
            })?;
        let payload_end = payload_start
            .checked_add(4)
            .ok_or(Error::CapacityOverflow {
                resource: "BrtBookView itabCur length",
            })?;
        let output_len = output.len();
        let bytes = output
            .get_mut(payload_start..payload_end)
            .ok_or(Error::InvalidLength {
                expected: payload_end,
                found: output_len,
            })?;
        bytes.copy_from_slice(
            &u32::try_from(next)
                .map_err(|_error| Error::CapacityOverflow {
                    resource: "active workbook sheet index",
                })?
                .to_le_bytes(),
        );
        transitions.push((current, next));
    }
    if !saw_book_view {
        return Ok(Vec::new());
    }
    if !transitions.is_empty() {
        let book_view_count =
            Records::new(&source, raw_limits).try_fold(0usize, |count, item| {
                let record = item?;
                Ok::<_, Error>(count.saturating_add(if record.kind() == kind::BOOK_VIEW {
                    1
                } else {
                    0
                }))
            })?;
        if book_view_count > 1 {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "active-tab relocation across multiple BrtBookView records is unsupported",
            ));
        }
    }
    Ok(transitions)
}

fn next_visible_sheet(current: usize, states: &[State]) -> Option<usize> {
    (1..=states.len())
        .map(|distance| current.saturating_add(distance) % states.len())
        .find(|position| states[*position] == State::Visible)
}

fn rewrite_sheet_view_flags(
    source: &[u8],
    kind: Kind,
    selected: bool,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Vec<u8>>> {
    let (record_kind, flag) = match kind {
        Kind::Worksheet => (kind::BEGIN_WS_VIEW, 0x0040u16),
        Kind::Chart => (kind::BEGIN_CS_VIEW, 0x0001u16),
        Kind::Dialog | Kind::Macro | Kind::IntlMacro | Kind::Unbound => return Ok(None),
    };
    if selected {
        let mut view_count = 0usize;
        for item in Records::new(source, raw_limits) {
            if item?.kind() == record_kind {
                view_count = view_count.saturating_add(1);
                if view_count > 1 {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit",
                        "cannot map multiple view records to one active workbook view",
                    ));
                }
            }
        }
        if view_count == 0 {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit",
                "cannot select active sheet without a view record",
            ));
        }
    }
    let mut output = source.to_vec();
    let mut changed = false;
    for item in Records::new(source, raw_limits) {
        let record = item?;
        if record.kind() != record_kind {
            continue;
        }
        let payload = record.payload();
        let flags = payload.get(..2).ok_or(Error::InvalidLength {
            expected: 2,
            found: payload.len(),
        })?;
        let old = u16::from_le_bytes([flags[0], flags[1]]);
        let updated = if selected { old | flag } else { old & !flag };
        if old == updated {
            continue;
        }
        let header_source = source.get(record.offset()..).ok_or(Error::InvalidFormat(
            "sheet-view record offset is outside its part",
        ))?;
        let (_, header_len) = Header::parse(header_source, raw_limits)?;
        let payload_start =
            record
                .offset()
                .checked_add(header_len)
                .ok_or(Error::CapacityOverflow {
                    resource: "sheet-view flags offset",
                })?;
        let payload_end = payload_start
            .checked_add(2)
            .ok_or(Error::CapacityOverflow {
                resource: "sheet-view flags length",
            })?;
        let output_len = output.len();
        let bytes = output
            .get_mut(payload_start..payload_end)
            .ok_or(Error::InvalidLength {
                expected: payload_end,
                found: output_len,
            })?;
        bytes.copy_from_slice(&updated.to_le_bytes());
        changed = true;
    }
    Ok(changed.then_some(output))
}

fn validate_sheet_state_invariants(workbook: &Workbook) -> Result<()> {
    let visible = workbook
        .sheets()
        .filter(|tab| tab.state() == State::Visible)
        .count();
    if visible == 0 {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit",
            "workbook must retain at least one visible sheet tab",
        ));
    }
    for tab in workbook.sheets() {
        if tab.kind() == Kind::Unbound && tab.state() != State::VeryHidden {
            return Err(Error::InvalidRelationship(
                "unbound XLSB sheet tabs must be very hidden",
            ));
        }
    }
    Ok(())
}

fn renamed_bundle_payload(
    source: &[u8],
    name: &str,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
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
        crate::package::records::decode_string(strings, raw_limits)?.1
    };
    let name_offset = strings_offset
        .checked_add(consumed)
        .ok_or(Error::CapacityOverflow {
            resource: "bundle-sheet name offset",
        })?;
    let name_source = source.get(name_offset..).ok_or(Error::InvalidLength {
        expected: name_offset,
        found: source.len(),
    })?;
    let old_units = usize::try_from(litchi_core::binary::read_u32_le_at(name_source, 0)?).map_err(
        |_overflow| Error::CapacityOverflow {
            resource: "bundle-sheet name units",
        },
    )?;
    let old_name_bytes = old_units
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(4))
        .ok_or(Error::CapacityOverflow {
            resource: "bundle-sheet name length",
        })?;
    let suffix_offset = name_offset
        .checked_add(old_name_bytes)
        .ok_or(Error::CapacityOverflow {
            resource: "bundle-sheet suffix offset",
        })?;
    let suffix = source.get(suffix_offset..).ok_or(Error::InvalidLength {
        expected: suffix_offset,
        found: source.len(),
    })?;
    let mut payload = source
        .get(..name_offset)
        .ok_or(Error::InvalidLength {
            expected: name_offset,
            found: source.len(),
        })?
        .to_vec();
    Encoder::new(&mut payload, raw_limits).write_wide_string(name)?;
    payload.extend_from_slice(suffix);
    let mut tolerance_report = crate::leniency::ToleranceReport::default();
    let _validated_record = crate::package::records::BundleSheetRecord::parse_with(
        &payload,
        0,
        crate::leniency::Tolerance::recording(&mut tolerance_report),
        raw_limits,
    )?;
    Ok(payload)
}

fn validate_sheet_name_value(name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0
        || units > 31
        || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return Err(Error::InvalidFormat("invalid sheet name"));
    }
    Ok(())
}

fn validate_staged_sheet_name(
    catalog: &[SheetTarget],
    operations: &[Operation],
    operation_indexes: &BTreeMap<Selection, Vec<usize>>,
    catalog_position: usize,
    name: &str,
) -> Result<()> {
    if catalog_position >= catalog.len() {
        return Err(Error::WorksheetNotFound(Some(catalog_position)));
    }
    // Every earlier staged name has already passed this check.  Resolve only
    // the effective names of the other catalog entries for the new request;
    // this keeps repeated staging linear in the catalog size instead of
    // rescanning every earlier operation for every earlier operation.
    for (right_position, right_tab) in catalog.iter().enumerate() {
        if right_position == catalog_position {
            continue;
        }
        let right_name = operation_indexes
            .get(&Selection::SheetRename(right_position))
            .and_then(|indexes| {
                indexes
                    .iter()
                    .rev()
                    .find_map(|index| match operations.get(*index) {
                        Some(Operation::SetSheetName { name, .. }) => Some(name.as_str()),
                        _ => None,
                    })
            })
            .unwrap_or(right_tab.name.as_str());
        if excel_name_eq(name, right_name) {
            return Err(Error::InvalidFormat("duplicate sheet name"));
        }
    }
    Ok(())
}

fn validate_staged_insert_sheet_name(
    catalog: &[SheetTarget],
    operations: &[Operation],
    name: &str,
) -> Result<()> {
    if catalog.iter().any(|tab| excel_name_eq(&tab.name, name)) {
        return Err(Error::InvalidFormat("duplicate sheet name"));
    }
    if operations.iter().any(|operation| {
        matches!(operation, Operation::InsertSheet { name: staged, .. }
            if excel_name_eq(staged, name))
    }) {
        return Err(Error::InvalidFormat("duplicate sheet name"));
    }
    Ok(())
}

fn validate_final_sheet_name_changes(
    tabs: &[crate::sheet::Tab<'_>],
    changes: &BTreeMap<usize, &str>,
) -> Result<()> {
    for (&left_position, left_name) in changes {
        if left_position >= tabs.len() {
            return Err(Error::WorksheetNotFound(Some(left_position)));
        }
        for (right_position, right_tab) in tabs.iter().enumerate() {
            if right_position != left_position
                && !changes.contains_key(&right_position)
                && excel_name_eq(left_name, right_tab.name())
            {
                return Err(Error::InvalidFormat("duplicate sheet name"));
            }
        }
        if changes.iter().any(|(&right_position, right_name)| {
            right_position != left_position && excel_name_eq(left_name, right_name)
        }) {
            return Err(Error::InvalidFormat("duplicate sheet name"));
        }
    }
    Ok(())
}

fn resolve_catalog_position(
    catalog: &[SheetTarget],
    selector: &Selector<'_>,
) -> Result<Option<usize>> {
    match selector {
        Selector::Name(name) => {
            let mut matches = catalog
                .iter()
                .enumerate()
                .filter(|(_, tab)| excel_name_eq(&tab.name, name.as_ref()));
            let first = matches.next().map(|(position, _)| position);
            if matches.next().is_some() {
                return Err(Error::InvalidFormat(
                    "sheet selector name is ambiguous in the complete tab catalog",
                ));
            }
            Ok(first)
        },
        Selector::Position(position) => {
            Ok((position.get() < catalog.len()).then_some(position.get()))
        },
    }
}

pub(crate) fn validate_all_worksheets(workbook: &Workbook) -> Result<()> {
    for index in 0..workbook.worksheet_count() {
        let uri = workbook.worksheet_uri(index)?;
        if workbook.package.get_part(&uri)?.content_type() == WORKSHEET_CONTENT_TYPE {
            let _cell_snapshot = workbook.cell_values(index)?;
            let _worksheet = workbook.worksheet(index)?;
        }
    }
    Ok(())
}

pub(super) fn validated_workbook_with_external_link_limits(
    bytes: &[u8],
    external_link_limits: ExternalLinkLimits,
    raw_limits: RawLimits,
) -> Result<Workbook> {
    validated_workbook_with_external_link_limits_and_drawing_policy(
        bytes,
        external_link_limits,
        DrawingLoadPolicy::Eager,
        raw_limits,
    )
}

pub(super) fn validated_workbook_with_external_link_limits_and_drawing_policy(
    bytes: &[u8],
    external_link_limits: ExternalLinkLimits,
    drawing_load_policy: DrawingLoadPolicy,
    raw_limits: RawLimits,
) -> Result<Workbook> {
    let workbook =
        workbook_from_bytes(bytes, external_link_limits, drawing_load_policy, raw_limits)?;
    validate_all_worksheets(&workbook)?;
    Ok(workbook)
}

fn workbook_from_bytes(
    bytes: &[u8],
    external_link_limits: ExternalLinkLimits,
    drawing_load_policy: DrawingLoadPolicy,
    raw_limits: RawLimits,
) -> Result<Workbook> {
    let package = litchi_opc::OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())?;
    Workbook::reparse_candidate_with_policy(
        package,
        external_link_limits,
        drawing_load_policy,
        crate::Leniency::Strict,
        raw_limits,
    )
}

fn workbook_bytes(workbook: &Workbook) -> Result<Vec<u8>> {
    crate::Package::from(workbook.package.clone()).to_bytes()
}

fn copy_arc(bytes: &[u8], resource: &'static str) -> Result<Arc<[u8]>> {
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copied.extend_from_slice(bytes);
    Ok(Arc::from(copied))
}

fn copy_vec_arc(bytes: &[u8], resource: &'static str) -> Result<Arc<Vec<u8>>> {
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copied.extend_from_slice(bytes);
    Ok(Arc::new(copied))
}

fn encoded_operation_length(operation: &Operation) -> Result<usize> {
    let length = match operation {
        Operation::RenameSheet { name, .. } | Operation::SetSheetName { name, .. } => {
            17usize.checked_add(name.len())
        },
        Operation::ReorderSheets { order } => {
            9usize.checked_add(order.len().checked_mul(size_of::<u64>()).ok_or(
                Error::CapacityOverflow {
                    resource: "durable workbook patch sheet order bytes",
                },
            )?)
        },
        Operation::InsertSheet { name, .. } => 18usize.checked_add(name.len()),
        Operation::RemoveSheets { positions } => {
            9usize.checked_add(positions.len().checked_mul(size_of::<u64>()).ok_or(
                Error::CapacityOverflow {
                    resource: "durable workbook patch sheet-removal bytes",
                },
            )?)
        },
        Operation::SetSheetState { .. } => Some(10),
        Operation::TransferCell { source, .. } => 41usize.checked_add(source.len()),
        Operation::AddImage { image, .. } => 74usize
            .checked_add(image.data.len())
            .and_then(|size| size.checked_add(image.description.as_ref().map_or(0, String::len))),
        Operation::TransferDrawing { source, .. } => 33usize.checked_add(source.len()),
        Operation::SetMergedRanges { ranges, .. } => merge_ranges_bytes(ranges.len()),
        Operation::SetDefinedName { name, .. } => defined_name_bytes(name),
    };
    length.ok_or(Error::CapacityOverflow {
        resource: "durable workbook patch operation bytes",
    })
}

fn encode_operation(
    operation: &Operation,
    output: &mut Vec<u8>,
    raw_limits: RawLimits,
) -> Result<()> {
    match operation {
        Operation::RenameSheet { sheet, name } => {
            // Tag 1 is retained for wire compatibility with patches
            // written before complete-catalog selectors.  Current edits do
            // not construct this variant; decoded instances are canonicalized
            // to tag 6 before they can be replayed or re-encoded.  Tag 6 is
            // metadata-only; tag 7 is the dependency-aware form emitted by
            // current edits.
            output.push(1);
            push_usize(output, *sheet)?;
            push_usize(output, name.len())?;
            output.extend_from_slice(name.as_bytes());
        },
        Operation::SetSheetName {
            catalog_position,
            name,
            dependencies,
        } => {
            output.push(if *dependencies { 7 } else { 6 });
            push_usize(output, *catalog_position)?;
            push_usize(output, name.len())?;
            output.extend_from_slice(name.as_bytes());
        },
        Operation::ReorderSheets { order } => {
            output.push(8);
            push_usize(output, order.len())?;
            for position in order.iter().copied() {
                push_usize(output, position)?;
            }
        },
        Operation::InsertSheet {
            catalog_position,
            name,
            state,
        } => {
            output.push(9);
            push_usize(output, *catalog_position)?;
            output.push(state_to_wire(*state));
            push_usize(output, name.len())?;
            output.extend_from_slice(name.as_bytes());
        },
        Operation::RemoveSheets { positions } => {
            output.push(10);
            push_usize(output, positions.len())?;
            for position in positions.iter().copied() {
                push_usize(output, position)?;
            }
        },
        Operation::SetSheetState {
            catalog_position,
            state,
        } => {
            output.push(5);
            push_usize(output, *catalog_position)?;
            output.push(state_to_wire(*state));
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
        Operation::TransferDrawing {
            source,
            source_sheet,
            source_anchor,
            target_sheet,
        } => {
            output.push(4);
            push_usize(output, source.len())?;
            push_usize(output, *source_sheet)?;
            push_usize(output, *source_anchor)?;
            push_usize(output, *target_sheet)?;
            output.extend_from_slice(source);
        },
        Operation::SetDefinedName {
            name,
            scope,
            target,
        } => {
            output.push(12);
            push_usize(output, name.len())?;
            output.extend_from_slice(name.as_bytes());
            output.extend_from_slice(&scope.map_or(u64::MAX, |scope| scope as u64).to_le_bytes());
            match target {
                Some((sheet, range)) => {
                    output.push(1);
                    push_usize(output, *sheet)?;
                    for value in [
                        range.row_first,
                        range.row_last,
                        range.col_first,
                        range.col_last,
                    ] {
                        output.extend_from_slice(&value.to_le_bytes());
                    }
                },
                None => output.push(0),
            }
        },
        Operation::SetMergedRanges {
            catalog_position,
            ranges,
        } => {
            output.push(11);
            push_usize(output, *catalog_position)?;
            push_usize(output, ranges.len())?;
            for range in ranges.iter() {
                output.extend_from_slice(&range.serialize(raw_limits)?);
            }
        },
    }
    Ok(())
}

fn decode_operation(data: &[u8], offset: &mut usize, raw_limits: RawLimits) -> Result<Operation> {
    let tag = read_slice(data, offset, 1)?[0];
    match tag {
        1 => {
            let sheet = read_next_usize(data, offset)?;
            let length = read_next_usize(data, offset)?;
            let name = std::str::from_utf8(read_slice(data, offset, length)?)
                .map_err(|_error| Error::InvalidFormat("sheet name is not UTF-8"))?
                .to_string();
            Ok(Operation::RenameSheet { sheet, name })
        },
        6 => {
            let catalog_position = read_next_usize(data, offset)?;
            let length = read_next_usize(data, offset)?;
            let name = std::str::from_utf8(read_slice(data, offset, length)?)
                .map_err(|_error| Error::InvalidFormat("sheet name is not UTF-8"))?
                .to_string();
            Ok(Operation::SetSheetName {
                catalog_position,
                name,
                dependencies: false,
            })
        },
        7 => {
            let catalog_position = read_next_usize(data, offset)?;
            let length = read_next_usize(data, offset)?;
            let name = std::str::from_utf8(read_slice(data, offset, length)?)
                .map_err(|_error| Error::InvalidFormat("sheet name is not UTF-8"))?
                .to_string();
            Ok(Operation::SetSheetName {
                catalog_position,
                name,
                dependencies: true,
            })
        },
        8 => {
            let count = read_next_usize(data, offset)?;
            let available = data
                .len()
                .checked_sub(*offset)
                .ok_or(Error::InvalidLength {
                    expected: *offset,
                    found: data.len(),
                })?;
            let required = count
                .checked_mul(size_of::<u64>())
                .ok_or(Error::CapacityOverflow {
                    resource: "workbook patch sheet order",
                })?;
            if required > available {
                return Err(Error::InvalidLength {
                    expected: offset.saturating_add(required),
                    found: data.len(),
                });
            }
            let mut order = Vec::new();
            order
                .try_reserve_exact(count)
                .map_err(|source| Error::Allocation {
                    resource: "workbook patch sheet order",
                    source,
                })?;
            for _ in 0..count {
                order.push(read_next_usize(data, offset)?);
            }
            Ok(Operation::ReorderSheets {
                order: order.into(),
            })
        },
        9 => {
            let catalog_position = read_next_usize(data, offset)?;
            let state = state_from_wire(read_slice(data, offset, 1)?[0])?;
            let length = read_next_usize(data, offset)?;
            let name = std::str::from_utf8(read_slice(data, offset, length)?)
                .map_err(|_error| Error::InvalidFormat("sheet name is not UTF-8"))?
                .to_string();
            Ok(Operation::InsertSheet {
                catalog_position,
                name,
                state,
            })
        },
        10 => {
            let count = read_next_usize(data, offset)?;
            let available = data
                .len()
                .checked_sub(*offset)
                .ok_or(Error::InvalidLength {
                    expected: *offset,
                    found: data.len(),
                })?;
            let required = count
                .checked_mul(size_of::<u64>())
                .ok_or(Error::CapacityOverflow {
                    resource: "workbook patch sheet removal",
                })?;
            if required > available {
                return Err(Error::InvalidLength {
                    expected: offset.saturating_add(required),
                    found: data.len(),
                });
            }
            let mut positions = Vec::new();
            positions
                .try_reserve_exact(count)
                .map_err(|source| Error::Allocation {
                    resource: "workbook patch sheet removal",
                    source,
                })?;
            for _ in 0..count {
                positions.push(read_next_usize(data, offset)?);
            }
            Ok(Operation::RemoveSheets {
                positions: positions.into(),
            })
        },
        5 => {
            let catalog_position = read_next_usize(data, offset)?;
            let state = state_from_wire(read_slice(data, offset, 1)?[0])?;
            Ok(Operation::SetSheetState {
                catalog_position,
                state,
            })
        },
        2 => {
            let source_len = read_next_usize(data, offset)?;
            let source_sheet = read_next_usize(data, offset)?;
            let source_row = read_next_u32(data, offset)?;
            let source_column = read_next_u32(data, offset)?;
            let target_sheet = read_next_usize(data, offset)?;
            let target_row = read_next_u32(data, offset)?;
            let target_column = read_next_u32(data, offset)?;
            let source = copy_arc(
                read_slice(data, offset, source_len)?,
                "workbook patch transfer-cell source",
            )?;
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
                let length = usize::try_from(description_len)
                    .map_err(|_error| Error::InvalidFormat("image description length overflow"))?;
                Some(
                    std::str::from_utf8(read_slice(data, offset, length)?)
                        .map_err(|_error| Error::InvalidFormat("image description is not UTF-8"))?
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
        4 => {
            let source_len = read_next_usize(data, offset)?;
            let source_sheet = read_next_usize(data, offset)?;
            let source_anchor = read_next_usize(data, offset)?;
            let target_sheet = read_next_usize(data, offset)?;
            let source = copy_arc(
                read_slice(data, offset, source_len)?,
                "workbook patch transfer-drawing source",
            )?;
            Ok(Operation::TransferDrawing {
                source,
                source_sheet,
                source_anchor,
                target_sheet,
            })
        },
        12 => {
            let length = read_next_usize(data, offset)?;
            let name = std::str::from_utf8(read_slice(data, offset, length)?)
                .map_err(|_error| Error::InvalidFormat("defined name is not UTF-8"))?
                .to_string();
            let scope = match read_next_u64(data, offset)? {
                u64::MAX => None,
                value => Some(
                    usize::try_from(value)
                        .map_err(|_error| Error::InvalidFormat("defined-name scope overflow"))?,
                ),
            };
            let target = match read_slice(data, offset, 1)?[0] {
                0 => None,
                1 => {
                    let sheet = read_next_usize(data, offset)?;
                    let row_first = read_next_u32(data, offset)?;
                    let row_last = read_next_u32(data, offset)?;
                    let col_first = read_next_u32(data, offset)?;
                    let col_last = read_next_u32(data, offset)?;
                    Some((
                        sheet,
                        crate::formula::Range::new(row_first, row_last, col_first, col_last)?,
                    ))
                },
                _flag => {
                    return Err(Error::InvalidFormat("defined-name target flag is invalid"));
                },
            };
            crate::named_ranges::validate_name(&name)?;
            Ok(Operation::SetDefinedName {
                name,
                scope,
                target,
            })
        },
        11 => {
            let catalog_position = read_next_usize(data, offset)?;
            let count = read_next_usize(data, offset)?;
            let available = data
                .len()
                .checked_sub(*offset)
                .ok_or(Error::InvalidLength {
                    expected: *offset,
                    found: data.len(),
                })?;
            let required =
                count
                    .checked_mul(crate::merged_cells::LEN)
                    .ok_or(Error::CapacityOverflow {
                        resource: "workbook patch merged ranges",
                    })?;
            if required > available {
                return Err(Error::InvalidLength {
                    expected: offset.saturating_add(required),
                    found: data.len(),
                });
            }
            let mut ranges = Vec::new();
            ranges
                .try_reserve_exact(count)
                .map_err(|source| Error::Allocation {
                    resource: "workbook patch merged ranges",
                    source,
                })?;
            for _ in 0..count {
                ranges.push(crate::merged_cells::MergedCell::parse(
                    read_slice(data, offset, crate::merged_cells::LEN)?,
                    raw_limits,
                )?);
            }
            Ok(Operation::SetMergedRanges {
                catalog_position,
                ranges: canonical_merge_ranges(ranges)?.into(),
            })
        },
        _ => Err(Error::InvalidFormat("unknown workbook patch operation")),
    }
}

/// Encoded length of a tag-12 defined-name operation.
fn defined_name_bytes(name: &str) -> Option<usize> {
    name.len().checked_add(1 + 8 + 8 + 1 + 8 + 16)
}

/// A stable hash of a defined name's case-folded spelling, for selection.
fn name_key(name: &str) -> u64 {
    name.to_uppercase()
        .bytes()
        .fold(0xcbf2_9ce4_8422_2325_u64, |hash, byte| {
            (hash ^ u64::from(byte)).wrapping_mul(0x0100_0000_01b3)
        })
}

/// Encoded length of a tag-11 merged-range replacement.
fn merge_ranges_bytes(count: usize) -> Option<usize> {
    count
        .checked_mul(crate::merged_cells::LEN)
        .and_then(|bytes| bytes.checked_add(17))
}

/// Sort `ranges` and refuse an invalid or overlapping set.
fn canonical_merge_ranges(
    mut ranges: Vec<crate::merged_cells::MergedCell>,
) -> Result<Vec<crate::merged_cells::MergedCell>> {
    if ranges.len() > crate::merged_cells::MAX_MERGED_CELL_RANGES {
        return Err(Error::InvalidLength {
            expected: crate::merged_cells::MAX_MERGED_CELL_RANGES,
            found: ranges.len(),
        });
    }
    for range in &ranges {
        range.validate()?;
        if range.row_first == range.row_last && range.col_first == range.col_last {
            return Err(Error::InvalidFormat(
                "a merged range must span more than one cell",
            ));
        }
    }
    ranges.sort_unstable_by_key(|range| {
        (
            range.row_first,
            range.col_first,
            range.row_last,
            range.col_last,
        )
    });
    for (index, range) in ranges.iter().enumerate() {
        if ranges[index + 1..]
            .iter()
            .take_while(|next| next.row_first <= range.row_last)
            .any(|next| next.overlaps(range))
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::InvalidValue,
                "edit",
                "merged ranges must not overlap",
            ));
        }
    }
    Ok(ranges)
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
        _ => Err(Error::InvalidFormat("unknown durable image format")),
    }
}

const fn state_to_wire(state: State) -> u8 {
    match state {
        State::Visible => 0,
        State::Hidden => 1,
        State::VeryHidden => 2,
    }
}

fn state_from_wire(value: u8) -> Result<State> {
    match value {
        0 => Ok(State::Visible),
        1 => Ok(State::Hidden),
        2 => Ok(State::VeryHidden),
        _ => Err(Error::InvalidFormat("unknown durable sheet state")),
    }
}

fn copy_record(
    source: &[u8],
    record: &crate::raw::Record<'_>,
    raw_limits: crate::raw::Limits,
    output: &mut Vec<u8>,
) -> Result<()> {
    let record_source = source.get(record.offset()..).ok_or(Error::InvalidFormat(
        "record offset is outside workbook.bin",
    ))?;
    let (_, header_len) = Header::parse(record_source, raw_limits)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "workbook record range",
        })?;
    output.extend_from_slice(
        source
            .get(record.offset()..end)
            .ok_or(Error::InvalidFormat("record range is outside workbook.bin"))?,
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
            "workbook patch transfer/history limits must be nonzero",
        ));
    }
    Ok(())
}

fn push_usize(output: &mut Vec<u8>, value: usize) -> Result<()> {
    let value = u64::try_from(value)
        .map_err(|_error| Error::InvalidFormat("workbook patch length exceeds u64"))?;
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
    usize::try_from(u64::from_le_bytes(bytes.try_into().map_err(|_error| {
        Error::InvalidFormat("invalid workbook patch length")
    })?))
    .map_err(|_error| Error::InvalidFormat("workbook patch length exceeds usize"))
}

fn read_next_usize(data: &[u8], offset: &mut usize) -> Result<usize> {
    let value = read_usize(data, *offset)?;
    *offset = offset.saturating_add(8);
    Ok(value)
}

fn read_next_u32(data: &[u8], offset: &mut usize) -> Result<u32> {
    let bytes = read_slice(data, offset, 4)?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_error| {
        Error::InvalidFormat("invalid workbook patch coordinate")
    })?))
}

fn read_next_u64(data: &[u8], offset: &mut usize) -> Result<u64> {
    let bytes = read_slice(data, offset, 8)?;
    Ok(u64::from_le_bytes(bytes.try_into().map_err(|_error| {
        Error::InvalidFormat("invalid workbook patch u64")
    })?))
}

fn read_next_i64(data: &[u8], offset: &mut usize) -> Result<i64> {
    let bytes = read_slice(data, offset, 8)?;
    Ok(i64::from_le_bytes(bytes.try_into().map_err(|_error| {
        Error::InvalidFormat("invalid workbook patch i64")
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
