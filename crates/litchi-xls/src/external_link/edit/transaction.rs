//! Failure-atomic metadata edits over one external-link snapshot.

use std::sync::Arc;

use crate::{Error, Result};

use super::super::codec::{encode_external_name, encode_supporting_book};
use super::super::model::{Name, NameBody, SupportingBook};
use super::super::package::Package;
use super::{Commit, Patch, Snapshot};

/// A detached transaction over inert supporting-book, external-name, and
/// external-cache metadata.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Vec<u8>,
    package: Package,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.bytes().to_vec(),
            package: source.package().clone(),
            source,
        }
    }

    /// Returns the immutable source snapshot used for publication checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Returns the current typed candidate view.
    #[must_use]
    pub fn links(&self) -> &crate::external_link::Links {
        self.package.links()
    }

    /// Materializes the current candidate as a validated immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if self.candidate.as_slice() == self.source.bytes() {
            Ok(self.source.clone())
        } else {
            Snapshot::parse(&self.candidate)
        }
    }

    /// Whether a staged edit changes any stream byte.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.as_slice() != self.source.bytes()
    }

    /// Replaces one external or DDE/OLE supporting-book target string.
    ///
    /// The target remains inert BIFF metadata; this operation never checks the
    /// filesystem or opens the path.  Sheet names, caches, formula bytes,
    /// unknown records, and any owned Continue records are retained.
    pub fn set_supporting_book_target(
        &mut self,
        book_index: usize,
        target: &str,
    ) -> Result<&mut Self> {
        let book = self
            .package
            .links()
            .supporting_books()
            .get(book_index)
            .cloned()
            .ok_or_else(|| {
                Error::UnsafeEdit(format!(
                    "supporting-book index {book_index} is outside the external-link collection"
                ))
            })?;
        let current = supporting_book_target(&book).ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "supporting-book {book_index} does not own editable target metadata"
            ))
        })?;
        if current == target {
            return Ok(self);
        }
        let record = self.package.supporting_book_record(book_index)?;
        let replacement = encode_supporting_book(&book, Some(target), None)?;
        self.replace_record(record, &replacement)?;
        Ok(self)
    }

    /// Renames one sheet owned by an external workbook supporting book.
    pub fn set_sheet_name(
        &mut self,
        book_index: usize,
        sheet_index: usize,
        name: &str,
    ) -> Result<&mut Self> {
        let book = self
            .package
            .links()
            .supporting_books()
            .get(book_index)
            .cloned()
            .ok_or_else(|| {
                Error::UnsafeEdit(format!(
                    "supporting-book index {book_index} is outside the external-link collection"
                ))
            })?;
        let SupportingBook::ExternalWorkbook(workbook) = &book else {
            return Err(Error::UnsafeEdit(format!(
                "supporting-book {book_index} does not own external sheet names"
            )));
        };
        let sheet = workbook.sheets().get(sheet_index).ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "sheet index {sheet_index} is outside supporting book {book_index}"
            ))
        })?;
        if sheet.name() == name {
            return Ok(self);
        }
        let record = self.package.supporting_book_record(book_index)?;
        let replacement = encode_supporting_book(&book, None, Some((sheet_index, name)))?;
        self.replace_record(record, &replacement)?;
        Ok(self)
    }

    /// Renames one `ExternName` while preserving its formula, matrix, flags,
    /// and all continuation payloads.
    pub fn set_external_name(&mut self, name_index: usize, name: &str) -> Result<&mut Self> {
        let current = self
            .package
            .links()
            .external_names()
            .get(name_index)
            .and_then(name_value)
            .ok_or_else(|| {
                Error::UnsafeEdit(format!(
                    "external-name index {name_index} is outside the external-link collection"
                ))
            })?;
        if current == name {
            return Ok(self);
        }
        let record = self.package.external_name_record(name_index)?;
        let replacement = encode_external_name(record.payload(&self.candidate), name)?;
        self.replace_record(record, &replacement)?;
        Ok(self)
    }

    /// Toggles the valid bit carried by one external sheet's `XCT` cache.
    ///
    /// The declared CRN cardinality and every CRN payload remain unchanged. A
    /// zero-cardinality cache cannot represent the invalid state in BIFF8 and
    /// is therefore rejected instead of being ambiguously rewritten.
    pub fn set_cache_valid(
        &mut self,
        book_index: usize,
        sheet_index: usize,
        valid: bool,
    ) -> Result<&mut Self> {
        let book = self
            .package
            .links()
            .supporting_books()
            .get(book_index)
            .ok_or_else(|| {
                Error::UnsafeEdit(format!(
                    "supporting-book index {book_index} is outside the external-link collection"
                ))
            })?;
        let SupportingBook::ExternalWorkbook(workbook) = book else {
            return Err(Error::UnsafeEdit(format!(
                "supporting-book {book_index} does not own external cache sheets"
            )));
        };
        let sheet = workbook.sheets().get(sheet_index).ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "sheet index {sheet_index} is outside supporting book {book_index}"
            ))
        })?;
        let cache = self.package.cache_record(book_index, sheet_index)?;
        if sheet.cache_valid() == valid {
            return Ok(self);
        }
        let magnitude = cache.declared.unsigned_abs();
        if !valid && magnitude == 0 {
            return Err(Error::UnsafeEdit(
                "a zero-cardinality XCT cannot represent an invalid cache".into(),
            ));
        }
        let declared = if valid {
            i16::try_from(magnitude).map_err(|_| {
                Error::UnsafeEdit("XCT cache cardinality does not fit signed BIFF8 count".into())
            })?
        } else {
            -i16::try_from(magnitude).map_err(|_| {
                Error::UnsafeEdit("XCT cache cardinality does not fit signed BIFF8 count".into())
            })?
        };
        let mut replacement = cache.record.payload(&self.candidate).to_vec();
        if replacement.len() != 4 {
            return Err(Error::UnsafeEdit(
                "selected XCT record is not owned by the parsed cache model".into(),
            ));
        }
        replacement[..2].copy_from_slice(&declared.to_le_bytes());
        self.replace_record(cache.record, &replacement)?;
        Ok(self)
    }

    /// Discards all staged edits and returns the original source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate with a reversible source-checked
    /// patch. Failed validation leaves the transaction candidate untouched.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source;
        if self.candidate.as_slice() == source.bytes() {
            let patch = Patch::new(source.clone(), source.clone());
            return Ok(Commit::new(source, patch));
        }
        let snapshot = Snapshot::parse_shared(Arc::from(self.candidate.into_boxed_slice()))?;
        let patch = Patch::new(source, snapshot.clone());
        Ok(Commit::new(snapshot, patch))
    }

    fn replace_record(
        &mut self,
        record: super::super::validation::RecordSpan,
        payload: &[u8],
    ) -> Result<()> {
        let candidate = self
            .package
            .replace_record(&self.candidate, record, payload)?;
        let package = Package::parse(&candidate)?;
        self.candidate = candidate;
        self.package = package;
        Ok(())
    }
}

fn supporting_book_target(book: &SupportingBook) -> Option<&str> {
    match book {
        SupportingBook::ExternalWorkbook(book) => Some(book.encoded_virtual_path()),
        SupportingBook::DdeOrOle {
            encoded_virtual_path,
        } => Some(encoded_virtual_path),
        _ => None,
    }
}

fn name_value(name: &Name) -> Option<&str> {
    match name.body() {
        NameBody::ExternalDefinedName { name, .. }
        | NameBody::AddInFunction { name, .. }
        | NameBody::DdeOrOle { name, .. }
        | NameBody::DdeStandardDocumentName { name } => Some(name),
    }
}
