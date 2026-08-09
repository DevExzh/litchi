//! Lossless edits of existing BIFF8 `Number` cell values.
//!
//! This deliberately narrow owner covers only existing `Number` records
//! (`[MS-XLS]` 2.4.180). It does not create cells or rewrite `RK`, `MulRk`,
//! formula, string, Boolean, error, blank, or shared-string records. Within
//! the Workbook stream, only the selected eight-byte `Xnum` field can change.
//! The complete CFB package is reopened before publication, and every other
//! captured stream retains its exact payload.

use crate::records::{BoundSheetRecord, Encoding, SheetType};
use crate::{Error, Result, Workbook};
use litchi_biff::Records;
use litchi_core::binary;
use litchi_ole_common::object::{Editor as PackageEditor, Limits, Targets};
use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const CODE_PAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const FILE_PASS: u16 = 0x002f;
const NUMBER: u16 = 0x0203;
const BIFF8: u16 = 0x0600;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const NUMBER_PAYLOAD_BYTES: usize = 14;
const NUMBER_VALUE_OFFSET: usize = 6;
const NUMBER_VALUE_BYTES: usize = 8;

/// A checked zero-based BIFF8 cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Reference {
    row: u16,
    column: u8,
}

impl Reference {
    /// Constructs a reference inside the 65,536 by 256 BIFF8 grid.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCellReference`] when either coordinate is
    /// outside the BIFF8 grid.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        let checked_row = u16::try_from(row).map_err(|_error| {
            Error::InvalidCellReference(format!(
                "cell row {row} is outside the BIFF8 worksheet grid"
            ))
        })?;
        let checked_column = u8::try_from(column).map_err(|_error| {
            Error::InvalidCellReference(format!(
                "cell column {column} is outside the BIFF8 worksheet grid"
            ))
        })?;
        Ok(Self {
            row: checked_row,
            column: checked_column,
        })
    }

    /// Returns the zero-based row.
    #[must_use]
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Returns the zero-based column.
    #[must_use]
    pub const fn column(self) -> u8 {
        self.column
    }
}

/// A semantic worksheet selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// Select a worksheet by case-insensitive developer-visible tab name.
    Name(&'a str),
    /// Select a worksheet by checked zero-based workbook tab position.
    Position(usize),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(value)
    }
}

/// An existing BIFF8 `Number` record in Workbook-stream order.
#[derive(Debug, Clone, Copy)]
pub struct Number {
    reference: Reference,
    value: f64,
}

impl Number {
    /// Returns the cell location.
    #[must_use]
    pub const fn reference(self) -> Reference {
        self.reference
    }

    /// Returns the exact stored IEEE-754 value.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.value
    }
}

impl PartialEq for Number {
    fn eq(&self, other: &Self) -> bool {
        self.reference == other.reference && self.value.to_bits() == other.value.to_bits()
    }
}

impl Eq for Number {}

#[derive(Debug, Clone)]
struct Entry {
    number: Number,
    value_offset: usize,
}

#[derive(Debug, Clone)]
struct SheetData {
    name: String,
    workbook_index: usize,
    entries: Vec<Entry>,
}

struct Inner {
    bytes: Arc<[u8]>,
    workbook_path: Vec<String>,
    workbook_stream: Arc<[u8]>,
    sheets: Vec<SheetData>,
}

/// Immutable, cheaply cloned snapshot of an editable XLS package.
#[derive(Clone)]
pub struct Snapshot {
    inner: Arc<Inner>,
}

impl Snapshot {
    /// Opens an unencrypted, unsigned XLS package and captures exact source bytes.
    ///
    /// The ordinary [`Workbook`] reader validates the complete candidate. The
    /// package transaction layer additionally refuses signed, encrypted, and
    /// DRM CFB containers before exposing an edit.
    ///
    /// # Errors
    ///
    /// Returns a typed CFB, BIFF, encryption, allocation, or workbook
    /// validation error before a snapshot is published.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = PackageEditor::open(bytes, Targets::default(), Limits::default())?;
        let workbook_path = [vec!["Workbook".to_string()], vec!["Book".to_string()]]
            .into_iter()
            .find(|path| package.stream(path).is_some())
            .ok_or_else(|| {
                Error::InvalidData("XLS package has no Workbook or Book stream".into())
            })?;
        let workbook_stream = package
            .stream_shared(&workbook_path)
            .ok_or_else(|| Error::InvalidData("selected XLS Workbook stream disappeared".into()))?;
        let source = package.finish()?;

        let sheets = parse_workbook_stream(&workbook_stream)?;
        // A full semantic open catches cross-stream and workbook-global
        // dependencies before the narrower source-offset inventory is kept.
        // The legacy reader intentionally skips some malformed optional sheet
        // projections, so this edit owner additionally requires every sheet
        // it can mutate to have survived that complete semantic open.
        {
            let workbook = Workbook::new(Cursor::new(source.as_slice()))?;
            for sheet in &sheets {
                if workbook
                    .sheet(sheet.workbook_index)
                    .and_then(crate::SheetMetadata::parsed_worksheet_index)
                    .is_none()
                {
                    return Err(Error::UnsafeEdit(format!(
                        "worksheet at tab position {} was not published by the complete XLS reader",
                        sheet.workbook_index
                    )));
                }
            }
        }
        Ok(Self {
            inner: Arc::new(Inner {
                bytes: Arc::from(source),
                workbook_path,
                workbook_stream,
                sheets,
            }),
        })
    }

    /// Returns exact source CFB bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
    }

    /// Returns the exact source Workbook stream.
    #[must_use]
    pub fn workbook_stream(&self) -> &[u8] {
        &self.inner.workbook_stream
    }

    /// Returns the number of worksheet tabs exposed by this bounded owner.
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Iterates worksheet tabs in workbook order.
    #[must_use]
    pub fn worksheets(&self) -> impl ExactSizeIterator<Item = Worksheet<'_>> {
        (0..self.inner.sheets.len()).map(|index| Worksheet {
            snapshot: self,
            index,
        })
    }

    /// Resolves a worksheet without exposing a physical stream identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if a malformed source produced an ambiguous name.
    pub fn worksheet<'a>(&'a self, selector: Selector<'_>) -> Result<Option<Worksheet<'a>>> {
        Ok(self.resolve_sheet(selector)?.map(|index| Worksheet {
            snapshot: self,
            index,
        }))
    }

    /// Starts a detached edit against this immutable artifact.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            changes: Vec::new(),
        }
    }

    fn resolve_sheet(&self, selector: Selector<'_>) -> Result<Option<usize>> {
        match selector {
            Selector::Position(position) => Ok(self
                .inner
                .sheets
                .iter()
                .position(|sheet| sheet.workbook_index == position)),
            Selector::Name(name) => {
                let mut matches = self
                    .inner
                    .sheets
                    .iter()
                    .enumerate()
                    .filter(|(_, sheet)| caseless_eq(&sheet.name, name));
                let found = matches.next().map(|(index, _)| index);
                if matches.next().is_some() {
                    return Err(Error::UnsafeEdit(format!(
                        "worksheet name {name:?} is ambiguous"
                    )));
                }
                Ok(found)
            },
        }
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("artifact_bytes", &self.inner.bytes.len())
            .field("workbook_bytes", &self.inner.workbook_stream.len())
            .field("worksheet_count", &self.inner.sheets.len())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes() == other.bytes()
    }
}

impl Eq for Snapshot {}

/// Borrowed view of one worksheet's editable `Number` records.
#[derive(Debug, Clone, Copy)]
pub struct Worksheet<'a> {
    snapshot: &'a Snapshot,
    index: usize,
}

impl<'a> Worksheet<'a> {
    fn data(self) -> &'a SheetData {
        &self.snapshot.inner.sheets[self.index]
    }

    /// Returns the zero-based workbook tab position.
    #[must_use]
    pub fn position(self) -> usize {
        self.data().workbook_index
    }

    /// Returns the developer-visible worksheet name.
    #[must_use]
    pub fn name(self) -> &'a str {
        &self.data().name
    }

    /// Iterates editable numeric values in source order.
    #[must_use]
    pub fn numbers(self) -> impl ExactSizeIterator<Item = Number> + 'a {
        self.data().entries.iter().map(|entry| entry.number)
    }

    /// Looks up an editable numeric value.
    ///
    /// # Errors
    ///
    /// Returns an exact-source ambiguity error for duplicate `Number` records.
    pub fn number(self, reference: Reference) -> Result<Option<Number>> {
        unique_entry(&self.data().entries, reference).map(|entry| entry.map(|item| item.number))
    }
}

#[derive(Debug, Clone)]
struct Change {
    sheet: usize,
    entry: usize,
    value: f64,
}

/// Detached, failure-atomic edits of existing `Number` fields.
#[derive(Clone)]
pub struct Edit {
    source: Snapshot,
    changes: Vec<Change>,
}

impl Edit {
    /// Sets one existing BIFF8 `Number` value.
    ///
    /// The replacement must be a normal value or positive zero, as required
    /// by `Xnum`. An absent cell, or a cell represented by `RK`, `MulRk`, `Formula`,
    /// string, Boolean, error, or blank records, is a typed refusal rather than
    /// a request to change record families.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid Xnum, a missing or ambiguous worksheet,
    /// a duplicate cell, a non-`Number` cell family, or allocation failure.
    pub fn set_number(
        &mut self,
        selector: Selector<'_>,
        reference: Reference,
        value: f64,
    ) -> Result<()> {
        if !valid_xnum(value) {
            return Err(Error::UnsupportedFeature(
                "BIFF8 Number edits require a normal IEEE-754 value or positive zero".into(),
            ));
        }
        let sheet_index = self
            .source
            .resolve_sheet(selector)?
            .ok_or_else(|| Error::WorksheetNotFound("cell-value edit worksheet selector".into()))?;
        let entries = &self.source.inner.sheets[sheet_index].entries;
        let entry = unique_entry_index(entries, reference)?.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "cell ({}, {}) is absent or is not encoded as a BIFF8 Number record",
                reference.row(),
                reference.column()
            ))
        })?;
        if let Some(change) = self
            .changes
            .iter_mut()
            .find(|change| change.sheet == sheet_index && change.entry == entry)
        {
            change.value = value;
        } else {
            self.changes
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("staging Number cell changes"))?;
            self.changes.push(Change {
                sheet: sheet_index,
                entry,
                value,
            });
        }
        Ok(())
    }

    /// Publishes a fully reopened snapshot and reversible exact-source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if stream patching, CFB reconstruction, complete XLS
    /// reopen, or typed semantic readback fails.
    pub fn commit(self) -> Result<Commit> {
        let changed_fields = self
            .changes
            .iter()
            .filter(|change| {
                self.source.inner.sheets[change.sheet].entries[change.entry]
                    .number
                    .value
                    .to_bits()
                    != change.value.to_bits()
            })
            .count();
        if changed_fields == 0 {
            let patch = Patch::new(
                Arc::clone(&self.source.inner.bytes),
                Arc::clone(&self.source.inner.bytes),
            );
            return Ok(Commit {
                snapshot: self.source,
                patch,
                diagnostics: Diagnostics::default(),
            });
        }

        let mut workbook_bytes = self.source.inner.workbook_stream.to_vec();
        for change in &self.changes {
            let entry = &self.source.inner.sheets[change.sheet].entries[change.entry];
            if entry.number.value.to_bits() == change.value.to_bits() {
                continue;
            }
            let end = entry
                .value_offset
                .checked_add(NUMBER_VALUE_BYTES)
                .ok_or_else(|| Error::InvalidData("Number replacement range overflow".into()))?;
            let destination = workbook_bytes
                .get_mut(entry.value_offset..end)
                .ok_or_else(|| {
                    Error::InvalidData("Number replacement is outside the Workbook stream".into())
                })?;
            destination.copy_from_slice(&change.value.to_le_bytes());
        }

        let workbook: Arc<[u8]> = Arc::from(workbook_bytes);
        parse_workbook_stream(&workbook)?;
        let mut package = PackageEditor::open(
            self.source.inner.bytes.to_vec(),
            Targets::default(),
            Limits::default(),
        )?;
        package.put_stream_shared(&self.source.inner.workbook_path, workbook)?;
        let candidate = package.finish()?;
        let snapshot = Snapshot::from_bytes(candidate)?;
        verify_readback(&snapshot, &self.changes)?;
        let patch = Patch::new(
            Arc::clone(&self.source.inner.bytes),
            Arc::clone(&snapshot.inner.bytes),
        );
        Ok(Commit {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                changed_number_fields: changed_fields,
                touched_streams: 1,
            },
        })
    }
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("source", &self.source)
            .field("staged_changes", &self.changes.len())
            .finish()
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed_number_fields: usize,
    touched_streams: usize,
}

impl Diagnostics {
    /// Number of eight-byte `Xnum` fields changed.
    #[must_use]
    pub const fn changed_number_fields(self) -> usize {
        self.changed_number_fields
    }

    /// Number of CFB streams in the mutation closure.
    #[must_use]
    pub const fn touched_streams(self) -> usize {
        self.touched_streams
    }
}

/// Successful immutable cell-value publication.
#[derive(Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("snapshot", &self.snapshot)
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

impl Commit {
    /// Returns the reopened target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Returns content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Splits this commit into its snapshot, patch, and diagnostics.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

/// Reversible, exact-source-checked replacement of one XLS artifact.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    fn new(before: Arc<[u8]>, after: Arc<[u8]>) -> Self {
        Self { before, after }
    }

    /// Returns whether the patch is byte-exactly empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Returns the exact required source artifact.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact produced artifact.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Applies this patch only to its exact immutable source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsafeEdit`] for a stale source, or a typed package
    /// validation error if the retained target can no longer be reopened.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "XLS cell-value patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(source.clone());
        }
        Snapshot::from_bytes(self.after.to_vec())
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .field("is_empty", &self.is_empty())
            .finish()
    }
}

fn parse_workbook_stream(source: &Arc<[u8]>) -> Result<Vec<SheetData>> {
    let mut records = Records::new(source);
    let first = records.next().ok_or(Error::Eof("Workbook globals BOF"))??;
    require_bof(first.payload(), WORKBOOK_GLOBALS)?;

    let mut encoding = Encoding::from_codepage(1252)?;
    let mut bound_payloads = Vec::new();
    let mut globals_eof = false;
    for record_result in records.by_ref() {
        let record = record_result?;
        match record.kind().get() {
            CODE_PAGE if record.payload().len() == 2 => {
                encoding = Encoding::from_codepage(binary::read_u16_le_at(record.payload(), 0)?)?;
            },
            BOUND_SHEET => {
                let mut payload = Vec::new();
                payload
                    .try_reserve_exact(record.payload().len())
                    .map_err(|_error| Error::Allocation("retaining BoundSheet8 payload"))?;
                payload.extend_from_slice(record.payload());
                bound_payloads
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("indexing BoundSheet8 records"))?;
                bound_payloads.push(payload);
            },
            FILE_PASS => return Err(Error::PasswordRequired),
            EOF => {
                globals_eof = true;
                break;
            },
            _ => {},
        }
    }
    if !globals_eof {
        return Err(Error::Eof("Workbook globals EOF"));
    }

    let mut bound_sheets = Vec::new();
    bound_sheets
        .try_reserve_exact(bound_payloads.len())
        .map_err(|_error| Error::Allocation("decoding BoundSheet8 records"))?;
    for payload in &bound_payloads {
        bound_sheets.push(BoundSheetRecord::parse(payload, &encoding)?);
    }
    let mut positions = std::collections::HashSet::new();
    let mut sheets = Vec::new();
    for (workbook_index, bound) in bound_sheets.iter().enumerate() {
        if positions.contains(&bound.position) {
            return Err(Error::InvalidRecord {
                record_type: BOUND_SHEET,
                message: "multiple BoundSheet8 records point to one substream".into(),
            });
        }
        positions
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("validating BoundSheet8 positions"))?;
        positions.insert(bound.position);
        if bound.sheet_type != SheetType::WorkSheet {
            continue;
        }
        let start = usize::try_from(bound.position).map_err(|_error| {
            Error::InvalidData("BoundSheet8 position does not fit usize".into())
        })?;
        let data = source.get(start..).ok_or_else(|| Error::InvalidRecord {
            record_type: BOUND_SHEET,
            message: "BoundSheet8 points outside the Workbook stream".into(),
        })?;
        let entries = parse_worksheet(data, start)?;
        sheets
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing editable worksheets"))?;
        sheets.push(SheetData {
            name: bound.name.clone(),
            workbook_index,
            entries,
        });
    }
    Ok(sheets)
}

fn parse_worksheet(data: &[u8], base_offset: usize) -> Result<Vec<Entry>> {
    let mut records = Records::new(data);
    let first = records.next().ok_or(Error::Eof("worksheet BOF"))??;
    require_bof(first.payload(), WORKSHEET)?;
    let mut entries = Vec::new();
    let mut found_eof = false;
    for record_result in records {
        let record = record_result?;
        if record.kind().get() == EOF {
            found_eof = true;
            break;
        }
        if record.kind().get() != NUMBER {
            continue;
        }
        if record.payload().len() != NUMBER_PAYLOAD_BYTES {
            return Err(Error::InvalidLength {
                expected: NUMBER_PAYLOAD_BYTES,
                found: record.payload().len(),
            });
        }
        let row = binary::read_u16_le_at(record.payload(), 0)?;
        let encoded_column = binary::read_u16_le_at(record.payload(), 2)?;
        let column = u8::try_from(encoded_column).map_err(|_error| Error::InvalidRecord {
            record_type: NUMBER,
            message: "Number column is outside the BIFF8 worksheet grid".into(),
        })?;
        let value = binary::read_f64_le_at(record.payload(), NUMBER_VALUE_OFFSET)?;
        if !valid_xnum(value) {
            return Err(Error::InvalidRecord {
                record_type: NUMBER,
                message: "Number contains an Xnum forbidden by MS-XLS 2.5.342".into(),
            });
        }
        let value_offset = base_offset
            .checked_add(record.offset())
            .and_then(|offset| offset.checked_add(4 + NUMBER_VALUE_OFFSET))
            .ok_or_else(|| Error::InvalidData("Number value offset overflow".into()))?;
        entries
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("indexing editable Number records"))?;
        entries.push(Entry {
            number: Number {
                reference: Reference { row, column },
                value,
            },
            value_offset,
        });
    }
    if !found_eof {
        return Err(Error::Eof("worksheet EOF"));
    }
    Ok(entries)
}

fn valid_xnum(value: f64) -> bool {
    value.is_normal() || (value == 0.0 && value.is_sign_positive())
}

fn caseless_eq(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

fn require_bof(payload: &[u8], expected_substream: u16) -> Result<()> {
    if payload.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let version = binary::read_u16_le_at(payload, 0)?;
    let substream = binary::read_u16_le_at(payload, 2)?;
    if version != BIFF8 || substream != expected_substream {
        return Err(Error::InvalidRecord {
            record_type: BOF,
            message: format!(
                "expected BIFF8 substream 0x{expected_substream:04X}, found version 0x{version:04X} and substream 0x{substream:04X}"
            ),
        });
    }
    Ok(())
}

fn unique_entry(entries: &[Entry], reference: Reference) -> Result<Option<&Entry>> {
    Ok(unique_entry_index(entries, reference)?.map(|index| &entries[index]))
}

fn unique_entry_index(entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
    let mut found = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.number.reference != reference {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(Error::UnsafeEdit(format!(
                "cell ({}, {}) has duplicate BIFF8 Number records",
                reference.row(),
                reference.column()
            )));
        }
    }
    Ok(found)
}

fn verify_readback(snapshot: &Snapshot, changes: &[Change]) -> Result<()> {
    for change in changes {
        let source_entry = &snapshot
            .inner
            .sheets
            .get(change.sheet)
            .ok_or_else(|| Error::UnsafeEdit("edited worksheet disappeared on readback".into()))?
            .entries;
        let reference = source_entry
            .get(change.entry)
            .ok_or_else(|| Error::UnsafeEdit("edited Number disappeared on readback".into()))?
            .number
            .reference;
        let value = unique_entry(source_entry, reference)?
            .ok_or_else(|| Error::UnsafeEdit("edited Number was not found on readback".into()))?
            .number
            .value;
        if value.to_bits() != change.value.to_bits() {
            return Err(Error::UnsafeEdit(
                "edited Number value failed semantic readback".into(),
            ));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests;
