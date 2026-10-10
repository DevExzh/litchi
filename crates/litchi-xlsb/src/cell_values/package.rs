//! Package-owned transactions for worksheet cell values.
//!
//! [`crate::cell_values::worksheet`] is the bounded worksheet-stream codec.
//! This module owns the package boundary around that codec: selectors resolve
//! against the complete workbook tab catalog, commits validate the complete
//! workbook, and patches carry exact whole-package images.
//!
//! A durable package-patch wire envelope is intentionally deferred; the
//! worksheet-only `LCXBCLP1` envelope is never used to represent a package
//! patch.

use super::worksheet::{
    CellError, CellFormula, Change, Limits, Patch as WorksheetPatch, Reference, Snapshot,
    StyleIndex, TransferLimits, Value,
};
use crate::cell_values::workbook::{self, Applied};
use crate::package::Package;
use crate::package::error::{Error, Result};
use crate::package::shared_strings::{PhoneticRun, SharedString, SharedStringRun};
use crate::sheet::Selector;
use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use std::io::{self, Write};
use std::mem::size_of;
use std::sync::Arc;

/// Publication diagnostics for a package-owned cell-values commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

/// A format-neutral scalar accepted by the package-owned XLSB cell writer.
///
/// The value is deliberately narrower than [`Value`].  Formula caches,
/// errors, rich-string runs, and shared-string indexes are preserved by the
/// existing family-specific methods; a scalar assignment changes only the
/// displayed direct value and retains the target cell's style and phonetic
/// flag.
#[derive(Debug, Clone, PartialEq)]
pub enum Scalar {
    /// A normalized IEEE-754 number or positive zero.
    Number(f64),
    /// A Boolean value.
    Boolean(bool),
    /// Plain text, bounded by the selected worksheet string policy.
    Text(String),
    /// A direct blank cell value.
    Blank,
}

impl Scalar {
    fn validate(&self, limits: Limits) -> Result<()> {
        match self {
            Self::Number(value) if !(value.is_normal() || value.to_bits() == 0) => {
                Err(Error::UnsupportedFeature(
                    "scalar number edits require a normalized IEEE-754 value or positive zero"
                        .to_string(),
                ))
            },
            Self::Text(value) => {
                let units = value.encode_utf16().count();
                let maximum = limits
                    .raw()
                    .string_units()
                    .min(super::worksheet::MAX_CELL_STRING_UNITS);
                if units > maximum {
                    return Err(Error::InvalidLength {
                        expected: maximum,
                        found: units,
                    });
                }
                Ok(())
            },
            Self::Number(_) | Self::Boolean(_) | Self::Blank => Ok(()),
        }
    }
}

impl Diagnostics {
    const fn new(changed: bool) -> Self {
        Self { changed }
    }

    /// Whether the commit changes the package image.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A detached package-owned edit of one worksheet's cell values.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Package,
    worksheet: PackURI,
    raw: super::worksheet::Edit,
    transfer_limits: TransferLimits,
    shared_strings: Option<Vec<SharedString>>,
    calculation_chain_present: bool,
}

impl Edit {
    fn new(
        base: Package,
        worksheet: PackURI,
        snapshot: Snapshot,
        transfer_limits: TransferLimits,
        calculation_chain_present: bool,
    ) -> Self {
        Self {
            base,
            worksheet,
            raw: snapshot.edit(),
            transfer_limits,
            shared_strings: None,
            calculation_chain_present,
        }
    }

    fn shared_string(&mut self, index: u32) -> Result<SharedString> {
        if self.shared_strings.is_none() {
            let workbook = self.base.workbook()?;
            self.shared_strings = Some(workbook.shared_strings().to_vec());
        }
        let index = usize::try_from(index).map_err(|_error| {
            Error::InvalidCellReference(format!(
                "shared-string index {index} does not fit this platform"
            ))
        })?;
        self.shared_strings
            .as_ref()
            .and_then(|strings| strings.get(index))
            .cloned()
            .ok_or_else(|| {
                Error::InvalidCellReference(format!(
                    "shared-string index {index} is not present in the workbook"
                ))
            })
    }

    /// Assign a direct scalar value, allowing a deliberate storage-family
    /// change while retaining the existing cell header.
    ///
    /// Returns whether this assignment changes the staged cell.
    ///
    /// Missing cells are inserted with the default style.  Assigning
    /// [`Scalar::Blank`] to a missing cell is an exact no-op.  Array, shared,
    /// and data-table formula groups are refused as a unit; an ordinary
    /// formula may be replaced when the workbook has no opaque calculation
    /// chain that this leaf owner could update safely.
    ///
    /// # Examples
    ///
    /// ```
    /// use std::io::Cursor;
    ///
    /// use litchi_xlsb::Package;
    /// use litchi_xlsb::cell_values::{Reference, Scalar};
    /// use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
    ///
    /// let mut sheet = MutableWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, 1.0_f64);
    /// let mut writer = WorkbookWriter::new();
    /// writer.add_worksheet(sheet);
    /// let mut bytes = Cursor::new(Vec::new());
    /// writer.save(&mut bytes)?;
    /// let package = Package::from_bytes(bytes.into_inner())?;
    /// let mut edit = package.edit_cell_values(0)?.expect("worksheet");
    /// edit.set_scalar(Reference::new(0, 0)?, Scalar::Number(2.0))?;
    /// let _commit = edit.commit()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_scalar(&mut self, reference: Reference, scalar: Scalar) -> Result<bool> {
        scalar.validate(self.raw.limits())?;
        self.raw.scalar_assignment_guard(reference)?;
        if self.raw.is_removed(reference) {
            return Err(Error::UnsupportedFeature(
                "scalar edits cannot target a cell already staged for removal".to_string(),
            ));
        }
        let current = self.raw.current_cell(reference)?.cloned();
        let Some(current) = current else {
            if matches!(scalar, Scalar::Blank) {
                if self.raw.source_has_cell(reference)? {
                    return Err(Error::UnsupportedFeature(
                        "scalar edits cannot target an unsupported cell-record family".to_string(),
                    ));
                }
                return Ok(false);
            }
            let value = scalar_value(scalar);
            self.raw.insert(reference, StyleIndex::new(0)?, value)?;
            return Ok(true);
        };

        if (current.formula().is_some() || current.value().is_formula_cache())
            && self.calculation_chain_present
        {
            return Err(Error::UnsupportedFeature(
                "scalar edits cannot replace a formula while an opaque calculation chain is present"
                    .to_string(),
            ));
        }

        let replacement = match scalar {
            Scalar::Number(value) => match current.value() {
                Value::Number(old) | Value::RkNumber(old) if old.to_bits() == value.to_bits() => {
                    return Ok(false);
                },
                Value::Number(_) | Value::FormulaNumberCache(_) => Value::Number(value),
                Value::RkNumber(_) => Value::Number(value),
                Value::Blank
                | Value::Error(_)
                | Value::Boolean(_)
                | Value::InlineString(_)
                | Value::SharedStringIndex(_)
                | Value::RichString(_)
                | Value::FormulaStringCache(_)
                | Value::FormulaBooleanCache(_)
                | Value::FormulaErrorCache(_) => Value::Number(value),
            },
            Scalar::Boolean(value) => match current.value() {
                Value::Boolean(old) if *old == value => return Ok(false),
                Value::Boolean(_) | Value::FormulaBooleanCache(_) => Value::Boolean(value),
                Value::Blank
                | Value::RkNumber(_)
                | Value::Error(_)
                | Value::Number(_)
                | Value::InlineString(_)
                | Value::SharedStringIndex(_)
                | Value::RichString(_)
                | Value::FormulaStringCache(_)
                | Value::FormulaNumberCache(_)
                | Value::FormulaErrorCache(_) => Value::Boolean(value),
            },
            Scalar::Text(value) => match current.value() {
                Value::InlineString(old) if old == &value => return Ok(false),
                Value::SharedStringIndex(index) => {
                    if self.shared_string(*index)?.text == value {
                        return Ok(false);
                    }
                    Value::InlineString(value)
                },
                Value::RichString(old) if old.text == value => return Ok(false),
                Value::InlineString(_) | Value::RichString(_) | Value::FormulaStringCache(_) => {
                    Value::InlineString(value)
                },
                Value::Blank
                | Value::RkNumber(_)
                | Value::Error(_)
                | Value::Boolean(_)
                | Value::Number(_)
                | Value::FormulaNumberCache(_)
                | Value::FormulaBooleanCache(_)
                | Value::FormulaErrorCache(_) => Value::InlineString(value),
            },
            Scalar::Blank => match current.value() {
                Value::Blank => return Ok(false),
                Value::RkNumber(_)
                | Value::Error(_)
                | Value::Boolean(_)
                | Value::Number(_)
                | Value::InlineString(_)
                | Value::SharedStringIndex(_)
                | Value::RichString(_)
                | Value::FormulaStringCache(_)
                | Value::FormulaNumberCache(_)
                | Value::FormulaBooleanCache(_)
                | Value::FormulaErrorCache(_) => Value::Blank,
            },
        };
        self.raw.replace_scalar_value(reference, replacement)?;
        Ok(true)
    }

    /// Replace a value while retaining the existing BIFF12 storage family.
    pub fn set_value(&mut self, reference: Reference, value: Value) -> Result<()> {
        self.raw.set_value(reference, value)
    }

    /// Set an existing direct numeric value, retaining direct or RK storage.
    pub fn set_number(&mut self, reference: Reference, value: f64) -> Result<()> {
        self.raw.set_number(reference, value)
    }

    /// Set an existing Boolean cell.
    pub fn set_boolean(&mut self, reference: Reference, value: bool) -> Result<()> {
        self.raw.set_boolean(reference, value)
    }

    /// Set an existing cell error.
    pub fn set_error(&mut self, reference: Reference, value: CellError) -> Result<()> {
        self.raw.set_error(reference, value)
    }

    /// Set an existing inline string.
    pub fn set_inline_string(&mut self, reference: Reference, value: String) -> Result<()> {
        self.raw.set_inline_string(reference, value)
    }

    /// Set an existing shared-string index.
    pub fn set_shared_string_index(&mut self, reference: Reference, index: u32) -> Result<()> {
        self.raw.set_shared_string_index(reference, index)
    }

    /// Set an existing rich string.
    pub fn set_rich_string(&mut self, reference: Reference, value: SharedString) -> Result<()> {
        self.raw.set_rich_string(reference, value)
    }

    /// Set the cached numeric result of an existing formula cell.
    pub fn set_formula_number_cache(&mut self, reference: Reference, value: f64) -> Result<()> {
        self.raw.set_formula_number_cache(reference, value)
    }

    /// Set the cached Boolean result of an existing formula cell.
    pub fn set_formula_boolean_cache(&mut self, reference: Reference, value: bool) -> Result<()> {
        self.raw.set_formula_boolean_cache(reference, value)
    }

    /// Set the cached error result of an existing formula cell.
    pub fn set_formula_error_cache(
        &mut self,
        reference: Reference,
        value: CellError,
    ) -> Result<()> {
        self.raw.set_formula_error_cache(reference, value)
    }

    /// Set the cached string result of an existing formula cell.
    pub fn set_formula_string_cache(&mut self, reference: Reference, value: String) -> Result<()> {
        self.raw.set_formula_string_cache(reference, value)
    }

    /// Change the cell-XF style index of an existing cell.
    pub fn set_style(&mut self, reference: Reference, style: StyleIndex) -> Result<()> {
        self.raw.set_style(reference, style)
    }

    /// Change the phonetic-display flag of an existing cell.
    pub fn set_show_phonetic(&mut self, reference: Reference, show: bool) -> Result<()> {
        self.raw.set_show_phonetic(reference, show)
    }

    /// Replace the inert formula payload of an existing formula cell.
    pub fn set_formula(&mut self, reference: Reference, formula: CellFormula) -> Result<()> {
        self.raw.set_formula(reference, formula)
    }

    /// Insert one direct scalar cell.
    pub fn insert(&mut self, reference: Reference, style: StyleIndex, value: Value) -> Result<()> {
        self.raw.insert(reference, style, value)
    }

    /// Insert one formula cell with a typed cached-result family.
    pub fn insert_formula(
        &mut self,
        reference: Reference,
        style: StyleIndex,
        cache: Value,
        formula: CellFormula,
    ) -> Result<()> {
        self.raw.insert_formula(reference, style, cache, formula)
    }

    /// Remove one supported stored cell.
    pub fn remove(&mut self, reference: Reference) -> Result<()> {
        self.raw.remove(reference)
    }

    /// Validate the edit and publish an immutable package commit.
    pub fn commit(self) -> Result<Commit> {
        validate_limits(self.transfer_limits)?;
        let worksheet_limits = self.raw.limits();
        let raw_commit = self
            .raw
            .commit_with_change_limit(self.transfer_limits.changes())?;
        validate_transfer_limits(raw_commit.patch(), self.transfer_limits)?;
        let retained = size_of::<Patch>()
            .saturating_add(self.worksheet.as_str().len())
            .saturating_add(worksheet_retained_bytes(raw_commit.patch()));
        let image_budget = remaining_image_budget(retained, self.transfer_limits)?;
        let before: Arc<[u8]> = Arc::from(serialize_package(&self.base, image_budget)?);

        let applied = workbook::apply_retaining_parse(
            self.base.opc_package(),
            &self.worksheet,
            &raw_commit,
            self.base.external_link_limits(),
            self.base.drawing_load_policy(),
            self.base.leniency(),
        )?;

        let (package, after) = match applied {
            Applied::Unchanged(_) => (self.base.clone(), Arc::clone(&before)),
            Applied::Published { workbook, .. } => {
                let workbook = if shared_string_delta(raw_commit.patch().changes()) != 0 {
                    super::scalar_sst::reconcile_with_limits(
                        *workbook,
                        worksheet_limits,
                        shared_string_delta(raw_commit.patch().changes()),
                    )?
                } else {
                    *workbook
                };
                let remaining = remaining_image_budget(
                    retained.saturating_add(before.len()),
                    self.transfer_limits,
                )?;
                let (package, after) = publish_validated_workbook(&self.base, workbook, remaining)?;
                (package, Arc::from(after))
            },
        };
        let (_, worksheet_patch) = raw_commit.into_parts();

        let patch = Patch {
            before,
            after,
            worksheet: self.worksheet,
            worksheet_patch,
            transfer_limits: self.transfer_limits,
        };
        if patch.retained_bytes() > self.transfer_limits.bytes() {
            return Err(Error::LimitExceeded {
                resource: "cell-values retained patch bytes",
                actual: patch.retained_bytes(),
                maximum: self.transfer_limits.bytes(),
            });
        }
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

/// The immutable package result, exact patch, and diagnostics of one edit.
#[derive(Debug, Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Validated package that the commit would publish.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Exact whole-package patch produced by the commit.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume the commit and return its validated package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return the package and exact patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// An exact before/after package patch for one worksheet cell-values owner.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    worksheet: PackURI,
    worksheet_patch: WorksheetPatch,
    transfer_limits: TransferLimits,
}

impl Patch {
    /// Whether the exact package images are identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Exact whole-package image required before applying this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact whole-package image produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Cell-level semantic changes retained by this package patch.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        self.worksheet_patch.changes()
    }

    /// Approximate the bytes retained by this immutable package patch.
    ///
    /// The package images and the worksheet owner's before/after images are
    /// both retained.  Semantic cell deltas are retained as owned values as
    /// well, so their string, rich-string, phonetic, and formula payloads are
    /// included in the estimate.  Capacities are used for the public string
    /// and vector fields where they are available; the result is saturating
    /// because this accessor is used to enforce a finite history budget.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        size_of::<Self>()
            .saturating_add(self.before.len())
            .saturating_add(self.after.len())
            .saturating_add(self.worksheet.as_str().len())
            .saturating_add(worksheet_retained_bytes(&self.worksheet_patch))
    }

    /// Apply this patch to an exact package source.
    pub fn apply(&self, package: &Package) -> Result<Package> {
        validate_limits(self.transfer_limits)?;
        let source = serialize_package(package, self.transfer_limits.bytes())?;
        if source.as_slice() != self.before.as_ref() {
            return Err(Error::UnsupportedFeature(
                "cell-values package patch source is stale".to_string(),
            ));
        }
        if self.is_noop() {
            return Ok(package.clone());
        }
        crate::signature_policy::ensure_unsigned(package.opc_package())?;
        let candidate =
            OpcPackage::from_bytes_with_limits(&self.after, package.opc_package().read_limits())?;
        let workbook = package.project_candidate(candidate)?;
        super::structure::validate_all_worksheets(&workbook)?;
        Ok(Package::from_validated_workbook(workbook))
    }

    /// Return the exact inverse package patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            worksheet: self.worksheet.clone(),
            worksheet_patch: self.worksheet_patch.inverse(),
            transfer_limits: self.transfer_limits,
        }
    }
}

fn worksheet_retained_bytes(patch: &WorksheetPatch) -> usize {
    let mut bytes = patch
        .before()
        .len()
        .saturating_add(patch.after().len())
        .saturating_add(patch.changes().len().saturating_mul(size_of::<Change>()));
    for change in patch.changes() {
        for cell in [change.before(), change.after()].into_iter().flatten() {
            bytes = bytes.saturating_add(retained_stored_cell_bytes(cell));
        }
    }
    bytes
}

fn retained_stored_cell_bytes(cell: &super::worksheet::StoredCell) -> usize {
    size_of::<super::worksheet::StoredCell>()
        .saturating_add(retained_value_bytes(cell.value()))
        .saturating_add(cell.formula().map_or(0, |formula| {
            size_of::<CellFormula>()
                .saturating_add(formula.tokens().len())
                .saturating_add(formula.ancillary().len())
        }))
}

fn retained_value_bytes(value: &Value) -> usize {
    match value {
        Value::InlineString(value) | Value::FormulaStringCache(value) => value.capacity(),
        Value::RichString(value) => retained_shared_string_bytes(value),
        Value::Blank
        | Value::RkNumber(_)
        | Value::Error(_)
        | Value::Boolean(_)
        | Value::Number(_)
        | Value::SharedStringIndex(_)
        | Value::FormulaNumberCache(_)
        | Value::FormulaBooleanCache(_)
        | Value::FormulaErrorCache(_) => 0,
    }
}

fn retained_shared_string_bytes(value: &SharedString) -> usize {
    let mut bytes = value.text.capacity().saturating_add(
        value
            .runs
            .capacity()
            .saturating_mul(size_of::<SharedStringRun>()),
    );
    if let Some(phonetic) = value.phonetic.as_ref() {
        bytes = bytes
            .saturating_add(phonetic.text.capacity())
            .saturating_add(
                phonetic
                    .runs
                    .capacity()
                    .saturating_mul(size_of::<PhoneticRun>()),
            );
    }
    bytes
}

impl Package {
    /// Start a detached cell-values edit selected from the complete tab
    /// catalog using the default worksheet and transfer policies.
    pub fn edit_cell_values<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<Edit>> {
        self.edit_cell_values_with_limits(selector, Limits::DEFAULT, TransferLimits::DEFAULT)
    }

    /// Start a detached cell-values edit with explicit finite policies.
    ///
    /// The transfer byte budget bounds the retained package patch, including
    /// package and worksheet images plus owned semantic cell payloads.
    /// Worksheet parsing uses the separate worksheet policy.
    pub fn edit_cell_values_with_limits<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
        worksheet_limits: Limits,
        transfer_limits: TransferLimits,
    ) -> Result<Option<Edit>> {
        validate_limits(transfer_limits)?;
        let Some(worksheet) = resolve_worksheet(self, selector)? else {
            return Ok(None);
        };
        let snapshot = workbook::read_with_limits_and_leniency(
            self.opc_package(),
            &worksheet,
            worksheet_limits,
            self.leniency(),
        )?;
        let calculation_chain_present = has_calculation_chain(self.opc_package())?;
        Ok(Some(Edit::new(
            self.clone(),
            worksheet,
            snapshot,
            transfer_limits,
            calculation_chain_present,
        )))
    }

    /// Apply an exact-source cell-values patch and return the new package.
    pub fn apply_cell_values(&self, patch: &Patch) -> Result<Self> {
        patch.apply(self)
    }
}

fn resolve_worksheet<'a>(
    package: &Package,
    selector: impl Into<Selector<'a>>,
) -> Result<Option<PackURI>> {
    let workbook = package.workbook()?;
    let selector = selector.into();
    if let Selector::Name(name) = &selector {
        let matches = workbook
            .sheets()
            .filter(|tab| crate::package::formula::excel_name_eq(tab.name(), name.as_ref()))
            .count();
        if matches > 1 {
            return Err(Error::InvalidFormat(
                "cell-values selector name is ambiguous in the complete tab catalog".to_string(),
            ));
        }
    }
    let Some(tab) = workbook.sheet(selector)? else {
        return Ok(None);
    };
    let Some(position) = tab.worksheet_position() else {
        return Err(Error::UnsupportedFeature(
            "cell-values editing requires a worksheet tab with a grid".to_string(),
        ));
    };
    workbook.worksheet_uri(position.get()).map(Some)
}

fn has_calculation_chain(package: &OpcPackage) -> Result<bool> {
    let workbook = package.main_document_part()?;
    Ok(workbook.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::CALC_CHAIN
                | litchi_opc::constants::relationship_type::STRICT_CALC_CHAIN
        )
    }))
}

fn scalar_value(scalar: Scalar) -> Value {
    match scalar {
        Scalar::Number(value) => Value::Number(value),
        Scalar::Boolean(value) => Value::Boolean(value),
        Scalar::Text(value) => Value::InlineString(value),
        Scalar::Blank => Value::Blank,
    }
}

fn shared_string_delta(changes: &[Change]) -> i64 {
    changes.iter().fold(0_i64, |delta, change| {
        let before = change
            .before()
            .is_some_and(|cell| matches!(cell.value(), Value::SharedStringIndex(_)));
        let after = change
            .after()
            .is_some_and(|cell| matches!(cell.value(), Value::SharedStringIndex(_)));
        delta + if after { 1 } else { 0 } - if before { 1 } else { 0 }
    })
}

fn publish_validated_workbook(
    base: &Package,
    mut workbook: crate::Workbook,
    image_budget: usize,
) -> Result<(Package, Vec<u8>)> {
    super::structure::validate_all_worksheets(&workbook)?;
    // Serialize the already validated candidate directly.  The exact bytes
    // are admitted below and installed into the same projection before the
    // package retains it, so a consuming projection never points at the
    // pre-admission OPC graph.
    let after = serialize_opc_package(&workbook.package, image_budget)?;
    let admitted = OpcPackage::from_bytes_with_limits(&after, base.opc_package().read_limits())?;
    workbook.package = admitted;
    Ok((
        Package::from_validated_workbook_with_projection(workbook),
        after,
    ))
}

fn remaining_image_budget(before_len: usize, limits: TransferLimits) -> Result<usize> {
    limits
        .bytes()
        .checked_sub(before_len)
        .ok_or(Error::LimitExceeded {
            resource: "cell-values package patch images",
            actual: before_len.saturating_add(1),
            maximum: limits.bytes(),
        })
}

fn serialize_package(package: &Package, maximum: usize) -> Result<Vec<u8>> {
    serialize_opc_package(package.opc_package(), maximum)
}

fn serialize_opc_package(package: &OpcPackage, maximum: usize) -> Result<Vec<u8>> {
    let mut writer = BoundedWriter::new(maximum);
    let result = PackageWriter::write_to_stream(&mut writer, package);
    if writer.exceeded {
        return Err(Error::LimitExceeded {
            resource: "cell-values package image",
            actual: writer.written.max(maximum.saturating_add(1)),
            maximum,
        });
    }
    result?;
    Ok(writer.bytes)
}

struct BoundedWriter {
    bytes: Vec<u8>,
    maximum: usize,
    written: usize,
    exceeded: bool,
}

impl BoundedWriter {
    fn new(maximum: usize) -> Self {
        Self {
            bytes: Vec::new(),
            maximum,
            written: 0,
            exceeded: false,
        }
    }
}

impl Write for BoundedWriter {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let Some(end) = self.written.checked_add(buffer.len()) else {
            self.exceeded = true;
            return Err(io::Error::other("bounded package image length overflow"));
        };
        if end > self.maximum {
            self.exceeded = true;
            self.written = end;
            return Err(io::Error::other("bounded package image limit exceeded"));
        }
        if end > self.bytes.capacity() {
            let capacity = end
                .max(self.bytes.capacity().saturating_mul(2))
                .min(self.maximum);
            self.bytes
                .try_reserve_exact(capacity - self.bytes.len())
                .map_err(io::Error::other)?;
        }
        self.bytes.extend_from_slice(buffer);
        self.written = end;
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn validate_transfer_limits(patch: &WorksheetPatch, limits: TransferLimits) -> Result<()> {
    validate_limits(limits)?;
    if patch.changes().len() > limits.changes() {
        return Err(Error::LimitExceeded {
            resource: "cell-values patch changes",
            actual: patch.changes().len(),
            maximum: limits.changes(),
        });
    }
    let bytes = patch
        .before()
        .len()
        .checked_add(patch.after().len())
        .ok_or(Error::CapacityOverflow {
            resource: "cell-values patch images",
        })?;
    if bytes > limits.bytes() {
        return Err(Error::LimitExceeded {
            resource: "cell-values patch images",
            actual: bytes,
            maximum: limits.bytes(),
        });
    }
    Ok(())
}

fn validate_limits(limits: TransferLimits) -> Result<()> {
    if limits.bytes() == 0
        || limits.changes() == 0
        || limits.history_entries() == 0
        || limits.history_bytes() == 0
    {
        return Err(Error::InvalidFormat(
            "cell-values transfer/history limits must all be nonzero".to_string(),
        ));
    }
    Ok(())
}
