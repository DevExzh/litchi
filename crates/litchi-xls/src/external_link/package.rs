//! Workbook-level BIFF8 external-link record collection.
//!
//! This layer owns the ordered SupBook/ExternName/XCT/CRN/ExternSheet
//! sequence. It never resolves or opens an external target.

use std::collections::HashSet;

use super::codec::{
    invalid, parse_cache_row, parse_extern_sheet, parse_external_name, parse_ser_ar_values,
    parse_sup_book, read_u16, validate_reference,
};
use super::model::{Links, Name, NameBody, SheetReference, SupportingBook};
use super::validation::{RecordSpan, replace_record, scan_records};
use super::{
    CONTINUE_RECORD_TYPE, CRN_RECORD_TYPE, EXTERN_NAME_RECORD_TYPE, EXTERN_SHEET_RECORD_TYPE,
    MAX_CACHED_CELLS, MAX_EXTERNAL_NAME_BYTES, MAX_EXTERNAL_NAMES, MAX_EXTERNAL_REFERENCES,
    MAX_SUPPORTING_BOOKS, SUP_BOOK_RECORD_TYPE, XCT_RECORD_TYPE,
};
use crate::{Error, Result};

struct PendingCache {
    book: usize,
    sheet: usize,
    remaining: usize,
}

struct PendingDdeMatrix {
    name_index: usize,
    remaining: usize,
}

#[derive(Default)]
pub(crate) struct ExternalLinkCollector {
    books: Vec<SupportingBook>,
    references: Vec<SheetReference>,
    external_names: Vec<Name>,
    pending: Option<PendingCache>,
    pending_dde_matrix: Option<PendingDdeMatrix>,
    names_allowed: bool,
    extern_sheet_seen: bool,
    closed: bool,
    cached_cells: usize,
    external_name_bytes: usize,
}

/// A validated workbook-global BIFF8 stream with contextual external-link
/// ownership.  Raw records remain in their source order; only a transaction's
/// explicitly selected owner may replace its payload.
#[derive(Clone)]
pub(crate) struct Package {
    records: Vec<RecordSpan>,
    supporting_books: Vec<usize>,
    external_names: Vec<(usize, usize)>,
    caches: Vec<CacheRecord>,
    links: Links,
}

#[derive(Clone, Copy)]
pub(crate) struct CacheRecord {
    pub(crate) record: RecordSpan,
    pub(crate) book_index: usize,
    pub(crate) sheet_index: usize,
    pub(crate) declared: i16,
}

impl Package {
    pub(crate) fn parse(bytes: &[u8]) -> Result<Self> {
        let records = scan_records(bytes)?;
        let internal_sheet_count = records
            .iter()
            .filter(|record| record.record_type == 0x0085)
            .count();
        let mut collector = ExternalLinkCollector::new();
        for record in &records {
            collector.feed_record(record.record_type, record.payload(bytes))?;
        }
        let links = collector.finish(internal_sheet_count)?;

        let mut supporting_books = Vec::new();
        let mut external_names = Vec::new();
        let mut caches = Vec::new();
        let mut current_book = None;
        let mut names_allowed = false;
        let mut collection_started = false;
        let mut collection_closed = false;
        for (record_index, record) in records.iter().copied().enumerate() {
            match record.record_type {
                SUP_BOOK_RECORD_TYPE => {
                    collection_started = true;
                    if collection_closed {
                        return Err(Error::InvalidRecord {
                            record_type: record.record_type,
                            message: "SupBook appears outside the external-link collection".into(),
                        });
                    }
                    current_book = Some(supporting_books.len());
                    supporting_books.push(record_index);
                    names_allowed = true;
                },
                EXTERN_NAME_RECORD_TYPE => {
                    let book_index = current_book.ok_or_else(|| Error::InvalidRecord {
                        record_type: record.record_type,
                        message: "ExternName has no owning SupBook".into(),
                    })?;
                    if !names_allowed {
                        return Err(Error::InvalidRecord {
                            record_type: record.record_type,
                            message: "ExternName is outside its owning SupBook name collection"
                                .into(),
                        });
                    }
                    external_names.push((record_index, book_index));
                },
                XCT_RECORD_TYPE => {
                    let book_index = current_book.ok_or_else(|| Error::InvalidRecord {
                        record_type: record.record_type,
                        message: "XCT has no owning SupBook".into(),
                    })?;
                    let payload = record.payload(bytes);
                    if payload.len() != 4 {
                        return Err(Error::InvalidLength {
                            expected: 4,
                            found: payload.len(),
                        });
                    }
                    caches.push(CacheRecord {
                        record,
                        book_index,
                        sheet_index: usize::from(u16::from_le_bytes([payload[2], payload[3]])),
                        declared: i16::from_le_bytes([payload[0], payload[1]]),
                    });
                    names_allowed = false;
                },
                CRN_RECORD_TYPE | EXTERN_SHEET_RECORD_TYPE => names_allowed = false,
                CONTINUE_RECORD_TYPE => {},
                _ if collection_started => collection_closed = true,
                _ => {},
            }
        }

        if supporting_books.len() != links.supporting_books().len()
            || external_names.len() != links.external_names().len()
        {
            return Err(Error::InvalidRecord {
                record_type: SUP_BOOK_RECORD_TYPE,
                message: "external-link owner map does not match the semantic collection".into(),
            });
        }
        for (name_index, (_, book_index)) in external_names.iter().enumerate() {
            if usize::from(links.external_names()[name_index].supporting_book_index())
                != *book_index
            {
                return Err(Error::InvalidRecord {
                    record_type: EXTERN_NAME_RECORD_TYPE,
                    message: "ExternName owner does not match its SupBook record".into(),
                });
            }
        }
        for cache in &caches {
            let Some(SupportingBook::ExternalWorkbook(book)) =
                links.supporting_books().get(cache.book_index)
            else {
                return Err(Error::InvalidRecord {
                    record_type: XCT_RECORD_TYPE,
                    message: "XCT owner map does not match its supporting-book sheet".into(),
                });
            };
            if cache.sheet_index >= book.sheets.len() {
                return Err(Error::InvalidRecord {
                    record_type: XCT_RECORD_TYPE,
                    message: "XCT sheet index exceeds its supporting-book sheet count".into(),
                });
            }
        }
        Ok(Self {
            records,
            supporting_books,
            external_names,
            caches,
            links,
        })
    }

    pub(crate) fn links(&self) -> &Links {
        &self.links
    }

    pub(crate) fn supporting_book_record(&self, index: usize) -> Result<RecordSpan> {
        self.supporting_books
            .get(index)
            .and_then(|record| self.records.get(*record).copied())
            .ok_or_else(|| {
                Error::UnsafeEdit(format!(
                    "supporting-book index {index} is outside the external-link collection"
                ))
            })
    }

    pub(crate) fn external_name_record(&self, index: usize) -> Result<RecordSpan> {
        self.external_names
            .get(index)
            .and_then(|(record, _)| self.records.get(*record).copied())
            .ok_or_else(|| {
                Error::UnsafeEdit(format!(
                    "external-name index {index} is outside the external-link collection"
                ))
            })
    }

    pub(crate) fn cache_record(
        &self,
        book_index: usize,
        sheet_index: usize,
    ) -> Result<CacheRecord> {
        let mut matches = self
            .caches
            .iter()
            .copied()
            .filter(|cache| cache.book_index == book_index && cache.sheet_index == sheet_index);
        let cache = matches.next().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "external cache for supporting book {book_index}, sheet {sheet_index} is not declared"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::UnsafeEdit(format!(
                "external cache for supporting book {book_index}, sheet {sheet_index} is ambiguous"
            )));
        }
        Ok(cache)
    }

    pub(crate) fn replace_record(
        &self,
        bytes: &[u8],
        record: RecordSpan,
        payload: &[u8],
    ) -> Result<Vec<u8>> {
        if !self.records.iter().any(|candidate| {
            candidate.record_start == record.record_start
                && candidate.record_type == record.record_type
                && candidate.payload_end == record.payload_end
        }) {
            return Err(Error::UnsafeEdit(
                "external-link edit does not own the selected BIFF record".into(),
            ));
        }
        replace_record(bytes, record, payload)
    }
}

impl ExternalLinkCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        if self
            .pending
            .as_ref()
            .is_some_and(|pending| pending.remaining > 0)
            && record_type != CRN_RECORD_TYPE
        {
            return invalid(
                XCT_RECORD_TYPE,
                "XCT must be followed immediately by its declared CRN records",
            );
        }
        if self.pending_dde_matrix.is_some() && record_type != CONTINUE_RECORD_TYPE {
            return invalid(
                EXTERN_NAME_RECORD_TYPE,
                "DDE/OLE MOper matrix ended before all declared values were read",
            );
        }

        if record_type == CONTINUE_RECORD_TYPE {
            if let Some(pending) = &mut self.pending_dde_matrix {
                if data.is_empty() || data.len() > 8224 {
                    return invalid(
                        CONTINUE_RECORD_TYPE,
                        "ExternName Continue payload must be 1..=8224 bytes",
                    );
                }
                self.external_name_bytes = self
                    .external_name_bytes
                    .checked_add(data.len())
                    .ok_or_else(|| Error::InvalidRecord {
                        record_type: CONTINUE_RECORD_TYPE,
                        message: "ExternName continuation size overflows".to_string(),
                    })?;
                if self.external_name_bytes > MAX_EXTERNAL_NAME_BYTES {
                    return invalid(
                        CONTINUE_RECORD_TYPE,
                        "ExternName matrix data exceeds resource bound",
                    );
                }
                let NameBody::DdeOrOle {
                    matrix: Some(matrix),
                    ..
                } = &mut self.external_names[pending.name_index].body
                else {
                    unreachable!()
                };
                let mut values =
                    parse_ser_ar_values(data, pending.remaining, CONTINUE_RECORD_TYPE)?;
                if values.is_empty() || values.len() > pending.remaining {
                    return invalid(
                        CONTINUE_RECORD_TYPE,
                        "DDE/OLE Continue contains an invalid number of complete values",
                    );
                }
                pending.remaining -= values.len();
                matrix.values.append(&mut values);
                if pending.remaining == 0 {
                    self.pending_dde_matrix = None;
                }
                return Ok(());
            }
            if !self.books.is_empty() && !self.closed {
                return invalid(
                    CONTINUE_RECORD_TYPE,
                    "Continue is not associated with a DDE/OLE ExternName",
                );
            }
            return Ok(());
        }

        let target = matches!(
            record_type,
            SUP_BOOK_RECORD_TYPE
                | EXTERN_NAME_RECORD_TYPE
                | XCT_RECORD_TYPE
                | CRN_RECORD_TYPE
                | EXTERN_SHEET_RECORD_TYPE
        );
        if !target {
            if !self.books.is_empty() {
                self.closed = true;
            }
            return Ok(());
        }
        if self.closed {
            return invalid(
                record_type,
                "external-link record is outside its contiguous SUPBOOK collection",
            );
        }

        match record_type {
            SUP_BOOK_RECORD_TYPE => {
                if self.extern_sheet_seen {
                    return invalid(record_type, "SupBook must precede ExternSheet");
                }
                if self.books.len() >= MAX_SUPPORTING_BOOKS {
                    return invalid(record_type, "supporting-book count exceeds resource bound");
                }
                self.books.push(parse_sup_book(data)?);
                self.names_allowed = true;
            },
            EXTERN_NAME_RECORD_TYPE => {
                if !self.names_allowed {
                    return invalid(
                        record_type,
                        "ExternName must directly follow its active SupBook name collection",
                    );
                }
                if self.external_names.len() >= MAX_EXTERNAL_NAMES {
                    return invalid(record_type, "external name count exceeds resource bound");
                }
                let book_index =
                    self.books
                        .len()
                        .checked_sub(1)
                        .ok_or_else(|| Error::InvalidRecord {
                            record_type,
                            message: "ExternName appears without a preceding SupBook".to_string(),
                        })?;
                let name = parse_external_name(data, book_index, &self.books[book_index])?;
                self.external_name_bytes = self
                    .external_name_bytes
                    .checked_add(data.len())
                    .ok_or_else(|| Error::InvalidRecord {
                        record_type,
                        message: "external name size overflows".to_string(),
                    })?;
                if self.external_name_bytes > MAX_EXTERNAL_NAME_BYTES {
                    return invalid(record_type, "external name data exceeds resource bound");
                }
                self.external_names.push(name);
                if let NameBody::DdeOrOle {
                    matrix: Some(matrix),
                    ..
                } = &self.external_names.last().unwrap().body
                {
                    let expected =
                        (usize::from(matrix.last_column) + 1) * (usize::from(matrix.last_row) + 1);
                    if matrix.values.len() < expected {
                        self.pending_dde_matrix = Some(PendingDdeMatrix {
                            name_index: self.external_names.len() - 1,
                            remaining: expected - matrix.values.len(),
                        });
                    }
                }
            },
            XCT_RECORD_TYPE => {
                self.names_allowed = false;
                self.parse_xct(data)?;
            },
            CRN_RECORD_TYPE => self.parse_crn(data)?,
            EXTERN_SHEET_RECORD_TYPE => {
                self.names_allowed = false;
                self.extern_sheet_seen = true;
                let mut references = parse_extern_sheet(data)?;
                if self.references.len() + references.len() > MAX_EXTERNAL_REFERENCES {
                    return invalid(
                        record_type,
                        "external sheet reference count exceeds BIFF8 record bound",
                    );
                }
                self.references.append(&mut references);
            },
            _ => unreachable!(),
        }
        Ok(())
    }

    fn parse_xct(&mut self, data: &[u8]) -> Result<()> {
        if data.len() != 4 {
            return Err(Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let book_index = self
            .books
            .len()
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidRecord {
                record_type: XCT_RECORD_TYPE,
                message: "XCT appears without a preceding SupBook".to_string(),
            })?;
        let declared = i16::from_le_bytes([data[0], data[1]]);
        if declared == i16::MIN {
            return invalid(XCT_RECORD_TYPE, "XCT CRN count absolute value overflows");
        }
        let sheet_index = usize::from(read_u16(data, 2));
        let SupportingBook::ExternalWorkbook(book) = &mut self.books[book_index] else {
            return invalid(
                XCT_RECORD_TYPE,
                "XCT cache requires an external-workbook SupBook",
            );
        };
        let sheet = book
            .sheets
            .get_mut(sheet_index)
            .ok_or_else(|| Error::InvalidRecord {
                record_type: XCT_RECORD_TYPE,
                message: "XCT sheet index exceeds SupBook sheet count".to_string(),
            })?;
        if sheet.cache_declared {
            return invalid(XCT_RECORD_TYPE, "duplicate XCT cache for external sheet");
        }
        sheet.cache_declared = true;
        sheet.cache_valid = declared >= 0;
        let remaining = usize::from(declared.unsigned_abs());
        self.pending = Some(PendingCache {
            book: book_index,
            sheet: sheet_index,
            remaining,
        });
        Ok(())
    }

    fn parse_crn(&mut self, data: &[u8]) -> Result<()> {
        let pending = self.pending.as_mut().ok_or_else(|| Error::InvalidRecord {
            record_type: CRN_RECORD_TYPE,
            message: "CRN appears without XCT".to_string(),
        })?;
        if pending.remaining == 0 {
            return invalid(CRN_RECORD_TYPE, "more CRN records than declared by XCT");
        }
        let row = parse_cache_row(data)?;
        self.cached_cells = self
            .cached_cells
            .checked_add(row.values.len())
            .ok_or_else(|| Error::InvalidRecord {
                record_type: CRN_RECORD_TYPE,
                message: "cached cell count overflows".to_string(),
            })?;
        if self.cached_cells > MAX_CACHED_CELLS {
            return invalid(
                CRN_RECORD_TYPE,
                "external cached cell count exceeds resource bound",
            );
        }
        let SupportingBook::ExternalWorkbook(book) = &mut self.books[pending.book] else {
            unreachable!()
        };
        book.sheets[pending.sheet].cache_rows.push(row);
        pending.remaining -= 1;
        Ok(())
    }

    pub(crate) fn finish(self, internal_sheet_count: usize) -> Result<Links> {
        if self
            .pending
            .as_ref()
            .is_some_and(|pending| pending.remaining > 0)
        {
            return invalid(
                XCT_RECORD_TYPE,
                "workbook ended before all CRN records declared by XCT",
            );
        }
        if self.pending_dde_matrix.is_some() {
            return invalid(
                EXTERN_NAME_RECORD_TYPE,
                "workbook ended before all DDE/OLE matrix values were read",
            );
        }
        for reference in &self.references {
            validate_reference(reference, &self.books, internal_sheet_count)?;
        }
        for book in &self.books {
            if let SupportingBook::ExternalWorkbook(book) = book {
                for sheet in &book.sheets {
                    let mut cells = HashSet::new();
                    for row in &sheet.cache_rows {
                        for offset in 0..row.values.len() {
                            let column = usize::from(row.first_column) + offset;
                            if !cells.insert((row.row, column)) {
                                return invalid(
                                    CRN_RECORD_TYPE,
                                    "external cache contains duplicate cells",
                                );
                            }
                        }
                    }
                }
            }
        }
        Ok(Links {
            supporting_books: self.books,
            external_names: self.external_names,
            sheet_references: self.references,
        })
    }
}
