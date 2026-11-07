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
    CellError, CellFormula, Change, Limits, Patch as WorksheetPatch, Reference, StyleIndex,
    TransferLimits, Value,
};
use crate::cell_values::workbook::{self, Applied};
use crate::package::Package;
use crate::package::error::{Error, Result};
use crate::package::shared_strings::{PhoneticRun, SharedString, SharedStringRun};
use crate::sheet::Selector;
use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use std::borrow::Cow;
use std::io::{self, Write};
use std::mem::size_of;
use std::sync::Arc;

/// Finite policies for one package-owned cell-values edit
/// ([`Package::edit_cell_values_with_options`]).
///
/// The worksheet limits bound parsing of the selected worksheet; the transfer
/// limits bound the retained package patch.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Limits,
    transfer_limits: TransferLimits,
}

impl EditOptions {
    /// The default worksheet and transfer policies.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            limits: Limits::DEFAULT,
            transfer_limits: TransferLimits::DEFAULT,
        }
    }

    /// Select the worksheet parsing policy.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// Select the patch transfer and history policy.
    #[must_use]
    pub const fn transfer_limits(mut self, limits: TransferLimits) -> Self {
        self.transfer_limits = limits;
        self
    }
}

pub use litchi_core::edit::Diagnostics;

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
                Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit_cell_values",
                    "scalar number edits require a normalized IEEE-754 value or positive zero",
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

/// A detached package-owned edit of one worksheet's cell values.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Package,
    selector: Selector<'static>,
    worksheet: Option<PackURI>,
    resolved: bool,
    worksheet_limits: Limits,
    raw: Option<super::worksheet::Draft>,
    transfer_limits: TransferLimits,
    shared_strings: Option<Vec<SharedString>>,
    calculation_chain_present: Option<bool>,
    /// The styles part with the formats staged by [`Self::set_format`].
    styles: Option<StylesDraft>,
}

/// The working styles stream of an edit and the XF each staged format
/// produced from each source XF.
#[derive(Debug, Clone)]
struct StylesDraft {
    uri: PackURI,
    bytes: Vec<u8>,
    changed: bool,
    restyled: Vec<(super::Format, std::collections::BTreeMap<u32, u32>)>,
}

impl StylesDraft {
    fn load(base: &Package) -> Result<Self> {
        let workbook = base.opc_package().main_document_part()?;
        let mut targets = workbook.rels().iter().filter(|relationship| {
            !relationship.is_external()
                && matches!(
                    relationship.reltype(),
                    litchi_opc::constants::relationship_type::STYLES
                        | litchi_opc::constants::relationship_type::STRICT_STYLES
                )
        });
        let Some(relationship) = targets.next() else {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "cell formats need a workbook styles part",
            ));
        };
        if targets.next().is_some() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "the workbook has more than one styles relationship",
            ));
        }
        let uri = relationship.target_partname()?;
        let bytes = base.opc_package().get_part(&uri)?.blob().to_vec();
        Ok(Self {
            uri,
            bytes,
            changed: false,
            restyled: Vec::new(),
        })
    }

    fn restyle(
        &mut self,
        old: u32,
        format: &super::Format,
        limits: crate::raw::Limits,
    ) -> Result<u32> {
        let slot = match self
            .restyled
            .iter()
            .position(|(staged, _)| staged == format)
        {
            Some(slot) => slot,
            None => {
                self.restyled
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "XLSB staged cell formats",
                        source,
                    })?;
                self.restyled
                    .push((format.clone(), std::collections::BTreeMap::new()));
                self.restyled.len() - 1
            },
        };
        if let Some(index) = self.restyled[slot].1.get(&old) {
            return Ok(*index);
        }
        let mut bytes = self.bytes.clone();
        let index = super::format::restyle(&mut bytes, old, format, limits)?;
        if bytes != self.bytes {
            self.bytes = bytes;
            self.changed = true;
        }
        self.restyled[slot].1.insert(old, index);
        Ok(index)
    }
}

impl Edit {
    fn new(
        base: Package,
        selector: Selector<'static>,
        worksheet_limits: Limits,
        transfer_limits: TransferLimits,
    ) -> Self {
        Self {
            base,
            selector,
            worksheet: None,
            resolved: false,
            worksheet_limits,
            raw: None,
            transfer_limits,
            shared_strings: None,
            calculation_chain_present: None,
            styles: None,
        }
    }

    /// Resolve the selected tab against the workbook catalog.
    ///
    /// Construction of an edit is deliberately cold, including when the
    /// package came from [`litchi_opc::OpcPackage`] without a parsed workbook
    /// projection.  This first fallible resolution performs the catalog scan
    /// and preserves the selector's missing-tab capability as `Ok(None)`.
    /// A chart, dialog, macro, or otherwise non-grid tab is reported as the
    /// same typed refusal as an operation that reaches it through `commit`.
    pub fn resolve(mut self) -> Result<Option<Self>> {
        if self.resolve_selector()?.is_some() {
            Ok(Some(self))
        } else {
            Ok(None)
        }
    }

    fn resolve_selector(&mut self) -> Result<Option<PackURI>> {
        if !self.resolved {
            self.worksheet = resolve_worksheet(&self.base, &self.selector)?;
            self.resolved = true;
        }
        Ok(self.worksheet.clone())
    }

    fn worksheet_not_found(&self) -> Error {
        match &self.selector {
            Selector::Name(_) => Error::WorksheetNotFound(None),
            Selector::Position(position) => Error::WorksheetNotFound(Some(position.get())),
        }
    }

    fn ensure_prepared(&mut self) -> Result<()> {
        // A prepared draft implies a resolved worksheet; return before
        // `resolve_selector` clones its part name, since every staged
        // operation passes through here (record 0377).
        if self.raw.is_some() {
            return Ok(());
        }
        let worksheet = self
            .resolve_selector()?
            .ok_or_else(|| self.worksheet_not_found())?;
        let snapshot = workbook::read_with_limits_and_leniency(
            self.base.opc_package(),
            &worksheet,
            self.worksheet_limits,
            self.base.leniency(),
        )?;
        self.calculation_chain_present = Some(has_calculation_chain(self.base.opc_package())?);
        self.raw = Some(snapshot.draft());
        Ok(())
    }

    fn raw_mut(&mut self) -> Result<&mut super::worksheet::Draft> {
        self.ensure_prepared()?;
        self.raw
            .as_mut()
            .ok_or(Error::InvalidFormat("cell-values edit is not prepared"))
    }

    fn chain_present(&self) -> Result<bool> {
        self.calculation_chain_present
            .ok_or(Error::InvalidFormat("cell-values edit is not prepared"))
    }

    fn shared_string(&mut self, index: u32) -> Result<SharedString> {
        if self.shared_strings.is_none() {
            let workbook = self.base.workbook()?;
            self.shared_strings = Some(workbook.shared_strings().to_vec());
        }
        let index = usize::try_from(index).map_err(|_error| {
            Error::InvalidCellReference("shared-string index does not fit this platform")
        })?;
        self.shared_strings
            .as_ref()
            .and_then(|strings| strings.get(index))
            .cloned()
            .ok_or(Error::InvalidCellReference(
                "shared-string index is not present in the workbook",
            ))
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
    /// use litchi_xlsb::Package;
    /// use litchi_xlsb::cell_values::{Reference, Scalar};
    /// use litchi_xlsb::builder::{self, Builder};
    ///
    /// let mut sheet = builder::Worksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, 1.0_f64);
    /// let mut writer = Builder::new();
    /// writer.add_worksheet(sheet);
    /// let package = writer.build()?;
    /// let mut edit = package.edit_cell_values(0)?.resolve()?.expect("worksheet");
    /// edit.set_scalar(Reference::new(0, 0)?, Scalar::Number(2.0))?;
    /// let _commit = edit.commit()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn set_scalar(&mut self, reference: Reference, scalar: Scalar) -> Result<bool> {
        scalar.validate(self.raw_mut()?.limits())?;
        self.raw_mut()?.scalar_assignment_guard(reference)?;
        if self.raw_mut()?.is_removed(reference) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "edit_cell_values",
                "scalar edits cannot target a cell already staged for removal",
            ));
        }
        let current = self.raw_mut()?.current_cell(reference)?.cloned();
        let Some(current) = current else {
            if matches!(scalar, Scalar::Blank) {
                if self.raw_mut()?.source_has_cell(reference)? {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedTarget,
                        "edit_cell_values",
                        "scalar edits cannot target an unsupported cell-record family",
                    ));
                }
                return Ok(false);
            }
            let value = scalar_value(scalar);
            self.raw_mut()?
                .insert(reference, StyleIndex::new(0)?, value)?;
            return Ok(true);
        };

        if (current.formula().is_some() || current.value().is_formula_cache())
            && self.chain_present()?
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "scalar edits cannot replace a formula while an opaque calculation chain is present",
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
        self.raw_mut()?
            .replace_scalar_value(reference, replacement)?;
        Ok(true)
    }

    /// Apply a partial format change to one cell.
    ///
    /// The cell's current XF and its font, fill and border records are
    /// copied, the attributes `format` names are patched, and the results are
    /// interned in the styles part; every other attribute keeps its bytes.
    /// A coordinate without a stored cell becomes a blank cell carrying the
    /// new format. Returns whether the staged cell changed: applying a format
    /// the cell already has is an exact no-op.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_xlsb::Package;
    /// use litchi_xlsb::cell_values::{Format, Reference};
    ///
    /// let package = Package::new()?;
    /// let mut edit = package.edit_cell_values(0)?.resolve()?.expect("worksheet");
    /// let mut format = Format::default();
    /// format.bold = Some(true);
    /// assert!(edit.set_format(Reference::new(0, 0)?, &format)?);
    /// assert!(edit.commit()?.diagnostics().changed());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a value outside its BIFF12 field, a coordinate in
    /// a formula group or with an unsupported cell-record family, a workbook
    /// without exactly one styles part, and malformed style records.
    pub fn set_format(&mut self, reference: Reference, format: &super::Format) -> Result<bool> {
        format.validate()?;
        if format.is_empty() {
            return Ok(false);
        }
        if self.raw_mut()?.is_removed(reference) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "edit_cell_values",
                "format edits cannot target a cell already staged for removal",
            ));
        }
        let current = self
            .raw_mut()?
            .current_cell(reference)?
            .map(|cell| cell.style());
        if current.is_none() && self.raw_mut()?.source_has_cell(reference)? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "edit_cell_values",
                "format edits cannot target an unsupported cell-record family",
            ));
        }
        let old = current.map_or(0, StyleIndex::get);
        if self.styles.is_none() {
            self.styles = Some(StylesDraft::load(&self.base)?);
        }
        let styles = self
            .styles
            .as_mut()
            .ok_or(Error::InvalidFormat("the styles draft was not loaded"))?;
        let new = styles.restyle(old, format, self.base.raw_limits_value())?;
        if new == old {
            return Ok(false);
        }
        let style = StyleIndex::new(new)?;
        match current {
            Some(_) => self.raw_mut()?.set_style(reference, style)?,
            None => self.raw_mut()?.insert(reference, style, Value::Blank)?,
        }
        Ok(true)
    }

    /// Set a cell's formula from Excel formula text (a leading `=` is
    /// optional), keeping the cell's style.
    ///
    /// The formula is stored with `fAlwaysCalc`, so Excel recalculates it
    /// when the workbook opens; no value is computed here. An ordinary
    /// formula cell keeps its cached result; a value cell keeps its value as
    /// the cached result, and a blank or new cell caches an empty string.
    /// Returns whether the staged cell changed.
    ///
    /// Formulas that name another sheet or a defined name need the workbook's
    /// extern-sheet and name tables and are refused
    /// ([`Error::UnsupportedFeature`]); references within the sheet,
    /// operators, literals and built-in functions are compiled.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid or unsupported formula text, a
    /// coordinate in a formula group or with an unsupported cell-record
    /// family, and a new formula while the workbook keeps an opaque
    /// calculation chain this owner cannot extend.
    pub fn set_formula_text(&mut self, reference: Reference, formula: &str) -> Result<bool> {
        let parsed = crate::package::formula::text::Compiler::compile(formula).map_err(|error| {
            if let Error::UnresolvedDependency(_message) = error {
                Error::refused(litchi_core::edit::RefusalKind::UnsupportedContent, "edit_cell_values", "the formula needs workbook context; only formulas within the sheet are compiled")
            } else {
                error
            }
        })?;
        let formula = CellFormula::new(
            0x0002,
            parsed.rgce,
            parsed.rgcb,
            self.base.raw_limits_value(),
        )?;
        self.raw_mut()?.scalar_assignment_guard(reference)?;
        if self.raw_mut()?.is_removed(reference) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "edit_cell_values",
                "formula edits cannot target a cell already staged for removal",
            ));
        }
        let current = self.raw_mut()?.current_cell(reference)?.cloned();
        let Some(current) = current else {
            if self.raw_mut()?.source_has_cell(reference)? {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    "edit_cell_values",
                    "formula edits cannot target an unsupported cell-record family",
                ));
            }
            if self.chain_present()? {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit_cell_values",
                    "a new formula cell cannot be added while an opaque calculation chain is present",
                ));
            }
            self.raw_mut()?.insert_formula(
                reference,
                StyleIndex::new(0)?,
                Value::FormulaStringCache(String::new()),
                formula,
            )?;
            return Ok(true);
        };
        if current.formula().is_some() {
            if current.formula() == Some(&formula) {
                return Ok(false);
            }
            self.raw_mut()?.set_formula(reference, formula)?;
            return Ok(true);
        }
        if self.chain_present()? {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "a value cell cannot become a formula while an opaque calculation chain is present",
            ));
        }
        let cache = match current.value() {
            Value::Number(value) | Value::RkNumber(value) => Value::FormulaNumberCache(*value),
            Value::Boolean(value) => Value::FormulaBooleanCache(*value),
            Value::Error(value) => Value::FormulaErrorCache(*value),
            Value::InlineString(value) => Value::FormulaStringCache(value.clone()),
            Value::SharedStringIndex(index) => {
                Value::FormulaStringCache(self.shared_string(*index)?.text)
            },
            Value::RichString(value) => Value::FormulaStringCache(value.text.clone()),
            Value::Blank => Value::FormulaStringCache(String::new()),
            Value::FormulaStringCache(_)
            | Value::FormulaNumberCache(_)
            | Value::FormulaBooleanCache(_)
            | Value::FormulaErrorCache(_) => {
                return Err(Error::InvalidFormat(
                    "a formula cache without a formula payload",
                ));
            },
        };
        self.raw_mut()?
            .replace_with_formula(reference, cache, formula)?;
        Ok(true)
    }

    /// Replace a value while retaining the existing BIFF12 storage family.
    pub fn set_value(&mut self, reference: Reference, value: Value) -> Result<()> {
        self.raw_mut()?.set_value(reference, value)
    }

    /// Set an existing direct numeric value, retaining direct or RK storage.
    pub fn set_number(&mut self, reference: Reference, value: f64) -> Result<()> {
        self.raw_mut()?.set_number(reference, value)
    }

    /// Set an existing Boolean cell.
    pub fn set_boolean(&mut self, reference: Reference, value: bool) -> Result<()> {
        self.raw_mut()?.set_boolean(reference, value)
    }

    /// Set an existing cell error.
    pub fn set_error(&mut self, reference: Reference, value: CellError) -> Result<()> {
        self.raw_mut()?.set_error(reference, value)
    }

    /// Set an existing inline string.
    pub fn set_inline_string(&mut self, reference: Reference, value: String) -> Result<()> {
        self.raw_mut()?.set_inline_string(reference, value)
    }

    /// Set an existing shared-string index.
    pub fn set_shared_string_index(&mut self, reference: Reference, index: u32) -> Result<()> {
        self.raw_mut()?.set_shared_string_index(reference, index)
    }

    /// Set an existing rich string.
    pub fn set_rich_string(&mut self, reference: Reference, value: SharedString) -> Result<()> {
        self.raw_mut()?.set_rich_string(reference, value)
    }

    /// Set the cached numeric result of an existing formula cell.
    pub fn set_formula_number_cache(&mut self, reference: Reference, value: f64) -> Result<()> {
        self.raw_mut()?.set_formula_number_cache(reference, value)
    }

    /// Set the cached Boolean result of an existing formula cell.
    pub fn set_formula_boolean_cache(&mut self, reference: Reference, value: bool) -> Result<()> {
        self.raw_mut()?.set_formula_boolean_cache(reference, value)
    }

    /// Set the cached error result of an existing formula cell.
    pub fn set_formula_error_cache(
        &mut self,
        reference: Reference,
        value: CellError,
    ) -> Result<()> {
        self.raw_mut()?.set_formula_error_cache(reference, value)
    }

    /// Set the cached string result of an existing formula cell.
    pub fn set_formula_string_cache(&mut self, reference: Reference, value: String) -> Result<()> {
        self.raw_mut()?.set_formula_string_cache(reference, value)
    }

    /// Change the cell-XF style index of an existing cell.
    pub fn set_style(&mut self, reference: Reference, style: StyleIndex) -> Result<()> {
        self.raw_mut()?.set_style(reference, style)
    }

    /// Change the phonetic-display flag of an existing cell.
    pub fn set_show_phonetic(&mut self, reference: Reference, show: bool) -> Result<()> {
        self.raw_mut()?.set_show_phonetic(reference, show)
    }

    /// Replace the inert formula payload of an existing formula cell.
    pub fn set_formula(&mut self, reference: Reference, formula: CellFormula) -> Result<()> {
        self.raw_mut()?.set_formula(reference, formula)
    }

    /// Insert one direct scalar cell.
    pub fn insert(&mut self, reference: Reference, style: StyleIndex, value: Value) -> Result<()> {
        self.raw_mut()?.insert(reference, style, value)
    }

    /// Insert one formula cell with a typed cached-result family.
    pub fn insert_formula(
        &mut self,
        reference: Reference,
        style: StyleIndex,
        cache: Value,
        formula: CellFormula,
    ) -> Result<()> {
        self.raw_mut()?
            .insert_formula(reference, style, cache, formula)
    }

    /// Remove one supported stored cell without shifting its neighbours.
    ///
    /// Returns `false` when the coordinate is already absent, including a
    /// repeated removal in this edit. The cell header, formula payload and
    /// cell/value metadata references are removed together; shared-string
    /// occurrence counts are reconciled at commit. Formula-group members and
    /// ordinary formulas with an opaque calculation chain are refused before staging.
    ///
    /// # Errors
    ///
    /// Returns an error for an ambiguous or unsupported stored record,
    /// formula dependencies that cannot be preserved, or exceeded limits.
    pub fn remove(&mut self, reference: Reference) -> Result<bool> {
        self.raw_mut()?.scalar_assignment_guard(reference)?;
        if self.raw_mut()?.is_removed(reference) {
            return Ok(false);
        }
        let Some(current) = self.raw_mut()?.current_cell(reference)? else {
            if self.raw_mut()?.source_has_cell(reference)? {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    "edit_cell_values",
                    "cell removal cannot target an unsupported cell-record family",
                ));
            }
            return Ok(false);
        };
        if (current.formula().is_some() || current.value().is_formula_cache())
            && self.chain_present()?
            && !self.raw_mut()?.is_inserted(reference)
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "cell removal cannot remove a formula while an opaque calculation chain is present",
            ));
        }
        self.raw_mut()?.remove(reference)?;
        Ok(true)
    }

    /// Validate the edit and publish an immutable package commit.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        validate_limits(self.transfer_limits)?;
        let raw = self
            .raw
            .take()
            .ok_or(Error::InvalidFormat("cell-values edit is not prepared"))?;
        let worksheet_limits = raw.limits();
        let raw_commit = raw.commit_with_change_limit(self.transfer_limits.changes())?;
        validate_transfer_limits(raw_commit.patch(), self.transfer_limits)?;
        let worksheet = self
            .worksheet
            .clone()
            .ok_or_else(|| self.worksheet_not_found())?;
        let retained = size_of::<Patch>()
            .saturating_add(worksheet.as_str().len())
            .saturating_add(worksheet_retained_bytes(raw_commit.patch()));
        let image_budget = remaining_image_budget(retained, self.transfer_limits)?;
        let before: Arc<[u8]> = Arc::from(serialize_package(&self.base, image_budget)?);

        // Staged formats intern their records into the styles part; the
        // worksheet commit is published over that candidate package.
        let styled = match &self.styles {
            Some(styles) if styles.changed => {
                let mut package = self.base.opc_package().clone();
                crate::signature_policy::ensure_unsigned(&package)?;
                package
                    .get_part_mut(&styles.uri)?
                    .set_blob(styles.bytes.clone());
                Some(package)
            },
            _ => None,
        };
        let applied = workbook::apply_retaining_parse(
            styled.as_ref().unwrap_or_else(|| self.base.opc_package()),
            &worksheet,
            &raw_commit,
            self.base.external_link_limits(),
            self.base.drawing_load_policy(),
            self.base.leniency(),
            self.base.raw_limits_value(),
        )?;

        let (package, after) = match applied {
            Applied::Unchanged => (self.base.clone(), Arc::clone(&before)),
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
            worksheet,
            worksheet_patch,
            transfer_limits: self.transfer_limits,
        };
        if patch.retained_bytes() > self.transfer_limits.bytes() {
            return Err(Error::Limit {
                kind: crate::cell_values::LimitKind::PatchBytes.into(),
                observed: patch.retained_bytes(),
                maximum: self.transfer_limits.bytes(),
            });
        }
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(package, patch, diagnostics))
    }
}

/// The immutable package result, exact patch, and diagnostics of one edit.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

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
    pub(crate) fn apply(&self, package: &Package) -> Result<Package> {
        validate_limits(self.transfer_limits)?;
        let source = serialize_package(package, self.transfer_limits.bytes())?;
        if source.as_slice() != self.before.as_ref() {
            return Err(Error::PatchConflict {
                part: "cell values",
            });
        }
        if self.is_noop() {
            return Ok(package.clone());
        }
        crate::signature_policy::ensure_unsigned(package.opc_package())?;
        // The patch owns an exact archive image. Retain that archive at
        // ingress so subsequent save/reopen and inverse use the same bytes.
        let candidate =
            OpcPackage::from_vec(self.after.to_vec(), package.opc_package().read_limits())?;
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
    /// Start a detached cell-values edit carrying a selector and the default
    /// worksheet and transfer policies.
    ///
    /// The selector is owned but resolved lazily.  Call [`Edit::resolve`] when
    /// the caller needs to distinguish a missing tab before staging the first
    /// cell operation; ordinary operations and `commit` resolve it on demand.
    pub fn edit_cell_values<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.edit_cell_values_with_options(selector, EditOptions::default())
    }

    /// Start a detached cell-values edit with explicit finite policies.
    ///
    /// The transfer byte budget bounds the retained package patch, including
    /// package and worksheet images plus owned semantic cell payloads.
    /// Worksheet parsing uses the separate worksheet policy.
    pub fn edit_cell_values_with_options<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
        options: EditOptions,
    ) -> Result<Edit> {
        let worksheet_limits = options.limits;
        let transfer_limits = options.transfer_limits;
        validate_limits(transfer_limits)?;
        Ok(Edit::new(
            self.clone(),
            own_selector(selector.into()),
            worksheet_limits,
            transfer_limits,
        ))
    }

    /// Apply an exact-source cell-values patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch source, a changed patch of a
    /// signed workbook, and the workbook validation errors of the result.
    pub fn apply_cell_values(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new(
            patch.apply(self)?,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

fn resolve_worksheet<'a>(package: &Package, selector: &Selector<'a>) -> Result<Option<PackURI>> {
    let workbook = package.workbook()?;
    if let Selector::Name(name) = selector {
        let matches = workbook
            .sheets()
            .filter(|tab| crate::package::formula::excel_name_eq(tab.name(), name.as_ref()))
            .count();
        if matches > 1 {
            return Err(Error::InvalidFormat(
                "cell-values selector name is ambiguous in the complete tab catalog",
            ));
        }
    }
    let Some(tab) = workbook.sheet(selector.clone())? else {
        return Ok(None);
    };
    let Some(position) = tab.worksheet_position() else {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_cell_values",
            "cell-values editing requires a worksheet tab with a grid",
        ));
    };
    workbook.worksheet_uri(position.get()).map(Some)
}

fn own_selector<'a>(selector: Selector<'a>) -> Selector<'static> {
    match selector {
        Selector::Name(name) => Selector::Name(Cow::Owned(name.into_owned())),
        Selector::Position(position) => Selector::Position(position),
    }
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
    let admitted = OpcPackage::from_vec(after.clone(), base.opc_package().read_limits())?;
    workbook.package = admitted;
    Ok((Package::from_validated_workbook(workbook), after))
}

fn remaining_image_budget(before_len: usize, limits: TransferLimits) -> Result<usize> {
    limits.bytes().checked_sub(before_len).ok_or(Error::Limit {
        kind: crate::cell_values::LimitKind::PatchBytes.into(),
        observed: before_len.saturating_add(1),
        maximum: limits.bytes(),
    })
}

fn serialize_package(package: &Package, maximum: usize) -> Result<Vec<u8>> {
    serialize_opc_package(package.opc_package(), maximum)
}

fn serialize_opc_package(package: &OpcPackage, maximum: usize) -> Result<Vec<u8>> {
    let mut writer = BoundedWriter::new(maximum);
    let result = PackageWriter::write_to(&mut writer, package);
    if writer.exceeded {
        return Err(Error::Limit {
            kind: super::LimitKind::PatchBytes.into(),
            observed: writer.written.max(maximum.saturating_add(1)),
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
        return Err(Error::Limit {
            kind: crate::cell_values::LimitKind::Changes.into(),
            observed: patch.changes().len(),
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
        return Err(Error::Limit {
            kind: crate::cell_values::LimitKind::PatchBytes.into(),
            observed: bytes,
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
            "cell-values transfer/history limits must all be nonzero",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selector_resolution_is_deferred_for_cold_packages() {
        let cold = Package::from(Package::new().expect("package").into_opc());
        assert!(!cold.projection_cached());

        let edit = cold.edit_cell_values("missing").expect("edit entry");
        assert!(!cold.projection_cached());
        assert!(edit.resolve().expect("selector resolution").is_none());
        // The edit owns a cloned immutable package, so its first resolution
        // does not mutate the original owner's private projection cache.
        assert!(!cold.projection_cached());
    }
}
