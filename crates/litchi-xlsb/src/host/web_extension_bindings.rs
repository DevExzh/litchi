#![allow(
    clippy::cast_possible_truncation,
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Worksheet web-extension binding records (MS-XLSB 2.4.868).

use litchi_ooxml_common::web;
use std::collections::HashSet;
use std::sync::Arc;

use crate::package::error::{Error, Result};
use crate::package::formula::{ParsedFormula, Parser, Token};
use crate::package::frt::{parse_formula_header, serialize_formula_header};
use crate::raw::{Records, Writer, kind};

const MAX_BINDINGS: usize = 65_536;
const MAX_APP_REF_CODE_UNITS: usize = 32_767;
const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u32 = 16_384;

/// The reference range encoded by a `BrtWebExtension` FRT formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    /// Index into the workbook's `ExternSheet` (`Xti`) collection.
    pub external_sheet_index: u16,
    pub first_row: u32,
    pub last_row: u32,
    pub first_column: u32,
    pub last_column: u32,
}

/// One binary worksheet-side Office Add-in binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    pub application_reference: String,
    pub range: Range,
    /// Exact `FRTFormula`, retained for lossless authoring.
    pub formula: ParsedFormula,
}

impl Binding {
    /// Construct and validate a binding from its native formula.
    ///
    /// `valid_external_sheet` must verify that the referenced XTI resolves to
    /// one internal worksheet (`firstSheet >= 0` and `firstSheet == lastSheet`).
    pub fn new(
        application_reference: impl Into<String>,
        formula: ParsedFormula,
        valid_external_sheet: impl FnOnce(u16) -> bool,
    ) -> Result<Self> {
        let application_reference = application_reference.into();
        validate_app_ref(&application_reference)?;
        let range = range_from_formula(&formula)?;
        if !valid_external_sheet(range.external_sheet_index) {
            return Err(invalid(
                "BrtWebExtension",
                "formula does not reference one internal worksheet",
            ));
        }
        Ok(Self {
            application_reference,
            range,
            formula,
        })
    }

    /// Parse one `BrtWebExtension` payload.
    pub fn parse_payload(
        data: &[u8],
        valid_external_sheet: impl FnOnce(u16) -> bool,
    ) -> Result<Self> {
        let (mut formulas, consumed) = parse_formula_header(data, "BrtWebExtension", 1)?;
        if formulas.len() != 1 {
            return Err(invalid(
                "BrtWebExtension",
                "FRTHeader must contain exactly one formula",
            ));
        }
        let formula = formulas.pop().ok_or_else(|| {
            invalid(
                "BrtWebExtension",
                "FRTHeader must contain exactly one formula",
            )
        })?;
        let string_data = data.get(consumed..).ok_or(Error::InvalidLength {
            expected: consumed,
            found: data.len(),
        })?;
        let application_reference = parse_wide_string_exact(string_data)?;
        Self::new(application_reference, formula, valid_external_sheet)
    }

    /// Serialize one `BrtWebExtension` payload.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        validate_app_ref(&self.application_reference)?;
        if range_from_formula(&self.formula)? != self.range {
            return Err(invalid(
                "BrtWebExtension",
                "cached range disagrees with its binary formula",
            ));
        }
        let mut output = serialize_formula_header(std::slice::from_ref(&self.formula), 1)?;
        write_wide_string(&mut output, &self.application_reference)?;
        Ok(output)
    }
}

/// Parse a complete `WEBEXTENSIONS` record collection.
pub fn parse_xlsb_web_extension_bindings(
    records: &[u8],
    mut valid_external_sheet: impl FnMut(u16) -> bool,
) -> Result<Vec<Binding>> {
    let mut iterator = Records::new(records);
    let begin = iterator
        .next()
        .ok_or_else(|| Error::UnexpectedEndOfStream("WEBEXTENSIONS".to_string()))??;
    if begin.kind() != kind::BEGIN_WEB_EXTENSIONS || !begin.payload().is_empty() {
        return Err(invalid(
            "WEBEXTENSIONS",
            "collection must start with empty BrtBeginWebExtensions",
        ));
    }
    let mut bindings = Vec::new();
    let mut app_refs = HashSet::new();
    loop {
        let record = iterator
            .next()
            .ok_or_else(|| Error::UnexpectedEndOfStream("WEBEXTENSIONS".to_string()))??;
        match record.kind() {
            kind::WEB_EXTENSION => {
                if bindings.len() == MAX_BINDINGS {
                    return Err(invalid("WEBEXTENSIONS", "binding count exceeds 65,536"));
                }
                let binding =
                    Binding::parse_payload(record.payload(), |index| valid_external_sheet(index))?;
                if !app_refs.insert(binding.application_reference.clone()) {
                    return Err(invalid("WEBEXTENSIONS", "duplicate binding appRef"));
                }
                bindings.push(binding);
            },
            kind::END_WEB_EXTENSIONS => {
                if !record.payload().is_empty() {
                    return Err(invalid("BrtEndWebExtensions", "end record must be empty"));
                }
                if bindings.is_empty() {
                    return Err(invalid(
                        "WEBEXTENSIONS",
                        "collection requires at least one binding",
                    ));
                }
                if iterator.next().is_some() {
                    return Err(invalid(
                        "WEBEXTENSIONS",
                        "records follow BrtEndWebExtensions",
                    ));
                }
                return Ok(bindings);
            },
            other => {
                return Err(invalid(
                    "WEBEXTENSIONS",
                    format!("unexpected record 0x{other:04X}"),
                ));
            },
        }
    }
}

/// Serialize a complete `WEBEXTENSIONS` record collection.
pub fn write_xlsb_web_extension_bindings(bindings: &[Binding]) -> Result<Vec<u8>> {
    validate_bindings(bindings, None)?;
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(kind::BEGIN_WEB_EXTENSIONS, &[])?;
    for binding in bindings {
        writer.write_record(kind::WEB_EXTENSION, &binding.to_payload()?)?;
    }
    writer.write_record(kind::END_WEB_EXTENSIONS, &[])?;
    Ok(output)
}

/// An immutable, source-bound worksheet `WEBEXTENSIONS` collection.
///
/// The snapshot retains the exact BIFF12 record stream.  Its typed bindings
/// are used for semantic edits, while the source bytes remain available for
/// exact no-op publication and stale-source checks.
#[derive(Debug, Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    bindings: Vec<Binding>,
    valid_external_sheets: ValidExternalSheets,
}

impl Snapshot {
    /// Parse and validate one complete `WEBEXTENSIONS` collection.
    ///
    /// The resolver is called only for XTI values actually present in the
    /// source collection.  When an edit needs to introduce a previously
    /// unseen XTI, use [`Self::read_with_external_sheet_indices`] to provide
    /// the complete worksheet relationship set.
    pub fn read(
        records: impl AsRef<[u8]>,
        mut valid_external_sheet: impl FnMut(u16) -> bool,
    ) -> Result<Self> {
        let records = records.as_ref();
        let mut valid_external_sheets = ValidExternalSheets::default();
        let bindings = parse_xlsb_web_extension_bindings(records, |index| {
            let valid = valid_external_sheet(index);
            if valid {
                valid_external_sheets.insert(index);
            }
            valid
        })?;
        validate_bindings(&bindings, Some(&valid_external_sheets))?;
        Ok(Self {
            source: Arc::from(records.to_vec().into_boxed_slice()),
            bindings,
            valid_external_sheets,
        })
    }

    /// Parse using an explicit set of valid worksheet XTI relationships.
    ///
    /// This form avoids probing a relationship resolver and permits later
    /// transactions to introduce any XTI in the supplied set.
    pub fn read_with_external_sheet_indices(
        records: impl AsRef<[u8]>,
        valid_external_sheets: impl IntoIterator<Item = u16>,
    ) -> Result<Self> {
        let valid_external_sheets = ValidExternalSheets::from_indices(valid_external_sheets);
        Self::from_source(
            Arc::from(records.as_ref().to_vec().into_boxed_slice()),
            valid_external_sheets,
        )
    }

    /// Alias for [`Self::read`].
    pub fn parse(
        records: impl AsRef<[u8]>,
        valid_external_sheet: impl FnMut(u16) -> bool,
    ) -> Result<Self> {
        Self::read(records, valid_external_sheet)
    }

    /// Parse and validate a collection against both worksheet XTI and package
    /// `MS-OWEXML` appRef relationships.
    pub fn read_with_package_bindings<'a>(
        records: impl AsRef<[u8]>,
        valid_external_sheet: impl FnMut(u16) -> bool,
        package_bindings: impl IntoIterator<Item = &'a web::Binding>,
    ) -> Result<Self> {
        let snapshot = Self::read(records, valid_external_sheet)?;
        snapshot.validate_apprefs(package_bindings)?;
        Ok(snapshot)
    }

    /// Borrow typed bindings in worksheet record order.
    #[must_use]
    pub fn bindings(&self) -> &[Binding] {
        &self.bindings
    }

    /// Borrow one typed binding by source order.
    #[must_use]
    pub fn binding(&self, index: usize) -> Option<&Binding> {
        self.bindings.get(index)
    }

    /// Return the exact source record stream.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Alias for [`Self::source_bytes`].
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.source_bytes()
    }

    /// Whether this snapshot is bound to an exact source stream.
    #[must_use]
    pub const fn is_source_bound(&self) -> bool {
        true
    }

    /// Validate worksheet bindings against package-level `MS-OWEXML` appRefs.
    pub fn validate_apprefs<'a>(
        &self,
        package_bindings: impl IntoIterator<Item = &'a web::Binding>,
    ) -> Result<()> {
        validate_xlsb_web_extension_apprefs(&self.bindings, package_bindings)
    }

    /// Start a detached, failure-atomic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            bindings: self.bindings.clone(),
            valid_external_sheets: self.valid_external_sheets.clone(),
        }
    }

    fn from_source(source: Arc<[u8]>, valid_external_sheets: ValidExternalSheets) -> Result<Self> {
        let bindings = parse_xlsb_web_extension_bindings(&source, |index| {
            valid_external_sheets.contains(index)
        })?;
        validate_bindings(&bindings, Some(&valid_external_sheets))?;
        Ok(Self {
            source,
            bindings,
            valid_external_sheets,
        })
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// A detached, typed edit over one worksheet `WEBEXTENSIONS` snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    bindings: Vec<Binding>,
    valid_external_sheets: ValidExternalSheets,
}

impl Transaction {
    /// Construct an edit from a validated source snapshot.
    #[must_use]
    pub fn new(source: Snapshot) -> Self {
        source.edit()
    }

    /// Borrow the immutable source snapshot used for conflict checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Borrow the currently staged typed bindings.
    #[must_use]
    pub fn bindings(&self) -> &[Binding] {
        &self.bindings
    }

    /// Borrow one currently staged binding by source order.
    #[must_use]
    pub fn binding(&self, index: usize) -> Option<&Binding> {
        self.bindings.get(index)
    }

    /// Whether staged serialization differs from the source semantics.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.bindings != self.source.bindings
    }

    /// Validate staged bindings against package-level `MS-OWEXML` appRefs.
    pub fn validate_apprefs<'a>(
        &self,
        package_bindings: impl IntoIterator<Item = &'a web::Binding>,
    ) -> Result<()> {
        validate_xlsb_web_extension_apprefs(&self.bindings, package_bindings)
    }

    /// Replace the complete typed collection while preserving source order.
    pub fn replace(&mut self, bindings: Vec<Binding>) -> Result<bool> {
        validate_bindings(&bindings, Some(&self.valid_external_sheets))?;
        if self.bindings == bindings {
            return Ok(false);
        }
        self.bindings = bindings;
        Ok(true)
    }

    /// Insert one binding at a source-relative index.
    pub fn insert(&mut self, index: usize, binding: Binding) -> Result<bool> {
        if index > self.bindings.len() {
            return Err(invalid(
                "WEBEXTENSIONS transaction",
                "binding insertion index is out of bounds",
            ));
        }
        if !self
            .valid_external_sheets
            .contains(binding.range.external_sheet_index)
        {
            return Err(invalid(
                "BrtWebExtension XTI",
                "binding does not reference a validated worksheet relationship",
            ));
        }
        let mut candidate = self.bindings.clone();
        candidate.insert(index, binding);
        validate_bindings(&candidate, Some(&self.valid_external_sheets))?;
        self.bindings = candidate;
        Ok(true)
    }

    /// Append one binding after the current collection.
    pub fn append(&mut self, binding: Binding) -> Result<bool> {
        self.insert(self.bindings.len(), binding)
    }

    /// Replace one binding in place.
    pub fn set(&mut self, index: usize, binding: Binding) -> Result<bool> {
        let mut candidate = self.bindings.clone();
        let current = candidate.get(index).ok_or_else(|| {
            invalid(
                "WEBEXTENSIONS transaction",
                "binding replacement index is out of bounds",
            )
        })?;
        if current == &binding {
            return Ok(false);
        }
        candidate[index] = binding;
        validate_bindings(&candidate, Some(&self.valid_external_sheets))?;
        self.bindings = candidate;
        Ok(true)
    }

    /// Alias for [`Self::set`].
    pub fn replace_at(&mut self, index: usize, binding: Binding) -> Result<bool> {
        self.set(index, binding)
    }

    /// Remove one binding by source-relative index.
    pub fn remove(&mut self, index: usize) -> Result<Option<Binding>> {
        if index >= self.bindings.len() {
            return Ok(None);
        }
        if self.bindings.len() == 1 {
            return Err(invalid(
                "WEBEXTENSIONS transaction",
                "a collection must retain at least one binding",
            ));
        }
        let mut candidate = self.bindings.clone();
        let removed = candidate.remove(index);
        validate_bindings(&candidate, Some(&self.valid_external_sheets))?;
        self.bindings = candidate;
        Ok(Some(removed))
    }

    /// Edit one binding through a failure-atomic closure.
    pub fn edit(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut Binding) -> Result<()>,
    ) -> Result<bool> {
        let mut candidate = self.bindings.clone();
        let binding = candidate.get_mut(index).ok_or_else(|| {
            invalid(
                "WEBEXTENSIONS transaction",
                "binding edit index is out of bounds",
            )
        })?;
        edit(binding)?;
        validate_bindings(&candidate, Some(&self.valid_external_sheets))?;
        if candidate == self.bindings {
            return Ok(false);
        }
        self.bindings = candidate;
        Ok(true)
    }

    /// Replace one binding's appRef while retaining its formula and range.
    pub fn set_application_reference(
        &mut self,
        index: usize,
        application_reference: impl Into<String>,
    ) -> Result<bool> {
        let application_reference = application_reference.into();
        self.edit(index, |binding| {
            binding.application_reference = application_reference;
            Ok(())
        })
    }

    /// Materialize the currently staged collection as a validated snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = write_xlsb_web_extension_bindings(&self.bindings)?;
        if bytes.as_slice() == self.source.source_bytes() {
            return Ok(self.source.clone());
        }
        Snapshot::from_source(
            Arc::from(bytes.into_boxed_slice()),
            self.valid_external_sheets.clone(),
        )
    }

    /// Discard staged edits and return the source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validate and publish the staged collection as a reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = self.snapshot()?;
        let patch = Patch {
            before: Arc::clone(&self.source.source),
            after: Arc::clone(&snapshot.source),
            before_external_sheets: self.source.valid_external_sheets,
            after_external_sheets: snapshot.valid_external_sheets.clone(),
        };
        let changed = !patch.is_empty();
        Ok(Commit {
            snapshot,
            patch,
            changed,
        })
    }
}

/// A successful source-checked worksheet collection publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether publication changed any record byte.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the publication into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible patch guarded by exact source bytes and XTI relationships.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    before_external_sheets: ValidExternalSheets,
    after_external_sheets: ValidExternalSheets,
}

impl Patch {
    /// Whether this patch preserves the exact source record stream.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Alias for [`Self::is_empty`].
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.is_empty()
    }

    /// Borrow the exact source record stream required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Borrow the exact record stream produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Apply only to the exact source snapshot used to create this patch.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.source.as_ref() != self.before.as_ref()
            || source.valid_external_sheets != self.before_external_sheets
        {
            return Err(invalid(
                "WEBEXTENSIONS patch",
                "source snapshot does not match the patch base",
            ));
        }
        if self.is_empty() {
            return Ok(source.clone());
        }
        Snapshot::from_source(Arc::clone(&self.after), self.after_external_sheets.clone())
    }

    /// Apply to exact source bytes and return the resulting record stream.
    pub fn apply_bytes(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before.as_ref() {
            return Err(invalid(
                "WEBEXTENSIONS patch",
                "source bytes do not match the patch base",
            ));
        }
        if !self.is_empty() {
            Snapshot::from_source(Arc::clone(&self.after), self.after_external_sheets.clone())?;
        }
        Ok(self.after.to_vec())
    }

    /// Alias for [`Self::apply`].
    pub fn commit(&self, source: &Snapshot) -> Result<Snapshot> {
        self.apply(source)
    }

    /// Apply the exact inverse to the committed target snapshot.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        self.inverse().apply(target)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            before_external_sheets: self.after_external_sheets.clone(),
            after_external_sheets: self.before_external_sheets.clone(),
        }
    }
}

/// Read one source-bound worksheet `WEBEXTENSIONS` collection.
pub fn read(
    records: impl AsRef<[u8]>,
    valid_external_sheet: impl FnMut(u16) -> bool,
) -> Result<Snapshot> {
    Snapshot::read(records, valid_external_sheet)
}

/// Apply a source-checked patch to exact worksheet collection bytes.
pub fn apply(source: &[u8], patch: &Patch) -> Result<Vec<u8>> {
    patch.apply_bytes(source)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ValidExternalSheets(Box<[u64; 1024]>);

impl Default for ValidExternalSheets {
    fn default() -> Self {
        Self(Box::new([0; 1024]))
    }
}

impl ValidExternalSheets {
    fn from_indices(indices: impl IntoIterator<Item = u16>) -> Self {
        let mut result = Self::default();
        for index in indices {
            result.insert(index);
        }
        result
    }

    fn insert(&mut self, index: u16) {
        let word = usize::from(index) / 64;
        let bit = usize::from(index) % 64;
        self.0[word] |= 1u64 << bit;
    }

    fn contains(&self, index: u16) -> bool {
        let word = usize::from(index) / 64;
        let bit = usize::from(index) % 64;
        self.0[word] & (1u64 << bit) != 0
    }
}

fn validate_bindings(
    bindings: &[Binding],
    valid_external_sheets: Option<&ValidExternalSheets>,
) -> Result<()> {
    if bindings.is_empty() || bindings.len() > MAX_BINDINGS {
        return Err(invalid(
            "WEBEXTENSIONS",
            "binding count must be in 1..=65,536",
        ));
    }
    let mut app_refs = HashSet::new();
    app_refs.try_reserve(bindings.len()).map_err(|_| {
        invalid(
            "WEBEXTENSIONS",
            "unable to reserve binding validation memory",
        )
    })?;
    for binding in bindings {
        validate_binding(binding, valid_external_sheets)?;
        if !app_refs.insert(&binding.application_reference) {
            return Err(invalid("WEBEXTENSIONS", "duplicate binding appRef"));
        }
    }
    Ok(())
}

fn validate_binding(
    binding: &Binding,
    valid_external_sheets: Option<&ValidExternalSheets>,
) -> Result<()> {
    validate_app_ref(&binding.application_reference)?;
    let range = range_from_formula(&binding.formula)?;
    if range != binding.range {
        return Err(invalid(
            "BrtWebExtension",
            "cached range disagrees with its binary formula",
        ));
    }
    if valid_external_sheets.is_some_and(|values| !values.contains(range.external_sheet_index)) {
        return Err(invalid(
            "BrtWebExtension XTI",
            "binding does not reference a validated worksheet relationship",
        ));
    }
    Ok(())
}

/// Require every binary worksheet `appRef` to resolve to one package binding.
pub fn validate_xlsb_web_extension_apprefs<'a>(
    worksheet_bindings: &[Binding],
    package_bindings: impl IntoIterator<Item = &'a web::Binding>,
) -> Result<()> {
    PackageAppRefs::new(package_bindings)?.validate(worksheet_bindings)
}

pub(crate) struct PackageAppRefs<'a> {
    values: HashSet<&'a str>,
}

impl<'a> PackageAppRefs<'a> {
    pub(crate) fn new(
        package_bindings: impl IntoIterator<Item = &'a web::Binding>,
    ) -> Result<Self> {
        let mut values = HashSet::new();
        let mut count = 0usize;
        for binding in package_bindings {
            count = count
                .checked_add(1)
                .ok_or_else(|| invalid("MS-OWEXML bindings", "package binding count overflow"))?;
            if count > MAX_BINDINGS {
                return Err(invalid(
                    "MS-OWEXML bindings",
                    "package binding count exceeds 65,536",
                ));
            }
            if values.len() == values.capacity() {
                values.try_reserve(1).map_err(|_| {
                    invalid(
                        "MS-OWEXML bindings",
                        "unable to reserve binding validation memory",
                    )
                })?;
            }
            if !values.insert(binding.app_ref()) {
                return Err(invalid(
                    "MS-OWEXML bindings",
                    "duplicate package binding appref",
                ));
            }
        }
        Ok(Self { values })
    }

    pub(crate) fn validate(&self, worksheet_bindings: &[Binding]) -> Result<()> {
        if worksheet_bindings.len() > MAX_BINDINGS {
            return Err(invalid(
                "WEBEXTENSIONS",
                "worksheet binding count exceeds 65,536",
            ));
        }
        for binding in worksheet_bindings {
            if !self.values.contains(binding.application_reference.as_str()) {
                return Err(invalid(
                    "BrtWebExtension.appRef",
                    format!(
                        "'{}' has no matching MS-OWEXML binding",
                        binding.application_reference
                    ),
                ));
            }
        }
        Ok(())
    }
}

fn range_from_formula(formula: &ParsedFormula) -> Result<Range> {
    if formula
        .rgce
        .first()
        .is_none_or(|token| token & 0x60 != 0x20)
    {
        return Err(invalid(
            "BrtWebExtension",
            "binding formula root must use the REFERENCE operand class",
        ));
    }
    let tokens = Parser::with_extra(&formula.rgce, &formula.rgcb).parse()?;
    if tokens.len() != 1 {
        return Err(invalid(
            "BrtWebExtension",
            "binding formula must be one reference expression",
        ));
    }
    let token = tokens.into_iter().next().ok_or_else(|| {
        invalid(
            "BrtWebExtension",
            "binding formula must be one reference expression",
        )
    })?;
    let range = match token {
        Token::CellRef3d {
            sheet_index,
            row,
            col,
            ..
        } => Range {
            external_sheet_index: sheet_index,
            first_row: row,
            last_row: row,
            first_column: col,
            last_column: col,
        },
        Token::AreaRef3d {
            sheet_index,
            row_first,
            row_last,
            col_first,
            col_last,
            ..
        } => Range {
            external_sheet_index: sheet_index,
            first_row: row_first,
            last_row: row_last,
            first_column: col_first,
            last_column: col_last,
        },
        Token::CellRef { .. } | Token::AreaRef { .. } | Token::ReferenceError { .. } => {
            return Err(invalid(
                "BrtWebExtension",
                "local and invalid reference tokens are forbidden",
            ));
        },
        _ => {
            return Err(invalid(
                "BrtWebExtension",
                "binding formula root is not a 3D reference",
            ));
        },
    };
    validate_range(&range)?;
    Ok(range)
}

fn validate_range(range: &Range) -> Result<()> {
    if range.first_row > range.last_row
        || range.last_row >= MAX_ROWS
        || range.first_column > range.last_column
        || range.last_column >= MAX_COLUMNS
    {
        return Err(invalid(
            "BrtWebExtension range",
            "range is outside the XLSB worksheet grid",
        ));
    }
    Ok(())
}

fn validate_app_ref(value: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > MAX_APP_REF_CODE_UNITS {
        return Err(invalid(
            "BrtWebExtension.appRef",
            "length must be in 1..=32,767 UTF-16 code units",
        ));
    }
    if value.chars().any(char::is_control) {
        return Err(invalid(
            "BrtWebExtension.appRef",
            "control characters are forbidden",
        ));
    }
    Ok(())
}

fn parse_wide_string_exact(data: &[u8]) -> Result<String> {
    let length_bytes = data.get(..4).ok_or(Error::InvalidLength {
        expected: 4,
        found: data.len(),
    })?;
    let length_bytes = <[u8; 4]>::try_from(length_bytes).map_err(|_| Error::InvalidLength {
        expected: 4,
        found: data.len(),
    })?;
    let count = u32::from_le_bytes(length_bytes) as usize;
    if count == 0 || count > MAX_APP_REF_CODE_UNITS {
        return Err(invalid("BrtWebExtension.appRef", "invalid string length"));
    }
    let expected = 4usize
        .checked_add(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid("BrtWebExtension.appRef", "length overflow"))?,
        )
        .ok_or_else(|| invalid("BrtWebExtension.appRef", "length overflow"))?;
    if data.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    let encoded = data.get(4..).ok_or(Error::InvalidLength {
        expected,
        found: data.len(),
    })?;
    let units = encoded
        .chunks_exact(2)
        .map(|bytes| {
            <[u8; 2]>::try_from(bytes)
                .map(u16::from_le_bytes)
                .map_err(|_| invalid("BrtWebExtension.appRef", "invalid UTF-16 unit"))
        })
        .collect::<Result<Vec<_>>>()?;
    String::from_utf16(&units)
        .map_err(|_| invalid("BrtWebExtension.appRef", "invalid UTF-16 string"))
}

fn write_wide_string(output: &mut Vec<u8>, value: &str) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.is_empty() || units.len() > MAX_APP_REF_CODE_UNITS {
        return Err(invalid("BrtWebExtension.appRef", "invalid string length"));
    }
    output.extend_from_slice(&(units.len() as u32).to_le_bytes());
    output.reserve(units.len() * 2);
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn invalid(typ: impl Into<String>, val: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.into(),
        val: val.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::formula::text::Compiler as TextCompiler;

    struct LyingHint<'a> {
        value: Option<&'a web::Binding>,
    }

    impl<'a> Iterator for LyingHint<'a> {
        type Item = &'a web::Binding;

        fn next(&mut self) -> Option<Self::Item> {
            self.value.take()
        }

        fn size_hint(&self) -> (usize, Option<usize>) {
            (usize::MAX, None)
        }
    }

    fn binding() -> Binding {
        // Public context-free compilation intentionally rejects 3D formulas;
        // construct the canonical PtgArea3d token directly.
        let binary = ParsedFormula {
            rgce: vec![0x3B, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 1, 0],
            rgcb: Vec::new(),
        };
        Binding::new("sales-table", binary, |index| index == 0).unwrap()
    }

    #[test]
    fn payload_and_collection_roundtrip() {
        let binding = binding();
        let payload = binding.to_payload().unwrap();
        assert_eq!(
            Binding::parse_payload(&payload, |index| index == 0).unwrap(),
            binding
        );
        let collection = write_xlsb_web_extension_bindings(std::slice::from_ref(&binding)).unwrap();
        assert_eq!(
            parse_xlsb_web_extension_bindings(&collection, |index| index == 0).unwrap(),
            [binding]
        );
    }

    #[test]
    fn rejects_invalid_xti_local_refs_and_trailing_payload() {
        let binding = binding();
        let payload = binding.to_payload().unwrap();
        assert!(Binding::parse_payload(&payload, |_| false).is_err());
        let local = TextCompiler::compile("$A$1:$B$4").unwrap();
        assert!(Binding::new("local", local, |_| true).is_err());
        let mut trailing = payload;
        trailing.push(0);
        assert!(Binding::parse_payload(&trailing, |_| true).is_err());
    }

    #[test]
    fn validates_package_appref_links() {
        let worksheet = [binding()];
        let package = [web::Binding::new("id", "table", "sales-table").unwrap()];
        validate_xlsb_web_extension_apprefs(&worksheet, &package).unwrap();
        validate_xlsb_web_extension_apprefs(
            &worksheet,
            LyingHint {
                value: Some(&package[0]),
            },
        )
        .unwrap();
        assert!(validate_xlsb_web_extension_apprefs(&worksheet, &[]).is_err());
    }

    #[test]
    fn source_checked_transaction_preserves_noop_and_reverses_edits() {
        let source = write_xlsb_web_extension_bindings(&[binding()]).unwrap();
        let snapshot = Snapshot::read(&source, |index| index == 0).unwrap();

        let noop = snapshot.edit().commit().unwrap();
        assert!(!noop.changed());
        assert!(noop.patch().is_empty());
        assert_eq!(noop.patch().before(), source.as_slice());
        assert_eq!(noop.patch().after(), source.as_slice());
        assert_eq!(noop.patch().apply(&snapshot).unwrap(), snapshot);

        let mut transaction = snapshot.edit();
        assert!(
            transaction
                .set_application_reference(0, "expenses-table")
                .unwrap()
        );
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_ne!(commit.patch().before(), commit.patch().after());

        let applied = commit.patch().apply(&snapshot).unwrap();
        assert_eq!(
            applied.bindings()[0].application_reference,
            "expenses-table"
        );
        let reverted = commit.patch().inverse().apply(&applied).unwrap();
        assert_eq!(reverted, snapshot);
        assert!(commit.patch().apply(commit.snapshot()).is_err());
        assert_eq!(
            commit.patch().apply_bytes(&source).unwrap(),
            commit.patch().after()
        );
    }

    #[test]
    fn transaction_rejects_invalid_bounds_relationships_and_malformed_sources() {
        assert!(
            validate_range(&Range {
                external_sheet_index: 0,
                first_row: MAX_ROWS,
                last_row: MAX_ROWS,
                first_column: 0,
                last_column: 0,
            })
            .is_err()
        );
        assert!(
            validate_range(&Range {
                external_sheet_index: 0,
                first_row: 2,
                last_row: 1,
                first_column: 0,
                last_column: 0,
            })
            .is_err()
        );

        let source = write_xlsb_web_extension_bindings(&[binding()]).unwrap();
        assert!(Snapshot::read(&source, |_| false).is_err());
        assert!(Snapshot::read(&source[..source.len() - 1], |_| true).is_err());
        assert!(Snapshot::read_with_external_sheet_indices(&source, []).is_err());

        let mut transaction = Snapshot::read(&source, |index| index == 0).unwrap().edit();
        let mut invalid = binding();
        invalid.application_reference.clear();
        assert!(transaction.set(0, invalid).is_err());
        assert_eq!(
            transaction.bindings()[0].application_reference,
            "sales-table"
        );
        assert!(transaction.remove(0).is_err());
    }
}
