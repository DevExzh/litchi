//! Bounded OLE2 and `[MS-OGRAPH]` chart payload decoding.

use std::io::{Cursor, Read};

use litchi_cfb::OleFile;
use litchi_cfb::consts::STGTY_STREAM;
use litchi_ograph::chart::Book;
use litchi_ograph::{Limits, Package as GraphPackage};

use super::model::Kind;
use crate::embedded::storage::{Compression, Kind as StorageKind, Storage};
use crate::package::{Error, Result};
use crate::presentation::Presentation;

const MAX_EXCEL_ROOT_ENTRIES: usize = 64;
const WORKBOOK: &str = "Workbook";
const BOOK: &str = "Book";

pub(super) enum Parsed {
    Graph {
        package: Box<GraphPackage>,
        book: Book,
    },
    Excel {
        book: Book,
    },
}

pub(super) fn decode(storage: Storage, limits: Limits) -> Result<Vec<u8>> {
    if storage.kind() != StorageKind::OleObject {
        return corrupted("chart persist ID does not reference an OLE object storage");
    }
    match storage.compression() {
        Compression::Uncompressed => {
            check_limit(
                "chart package bytes",
                storage.stored_payload_len(),
                limits.max_package_bytes,
            )?;
            Ok(storage.into_stored_bytes())
        },
        Compression::Zlib => {
            let uncompressed_len = storage.declared_uncompressed_len().ok_or_else(|| {
                Error::Corrupted("compressed chart storage is missing its size".into())
            })?;
            let declared = usize::try_from(uncompressed_len)
                .map_err(|_| Error::Corrupted("chart storage size exceeds usize".into()))?;
            check_limit("chart package bytes", declared, limits.max_package_bytes)?;
            let capacity = declared
                .checked_add(1)
                .ok_or_else(|| Error::Corrupted("chart storage size overflows usize".into()))?;
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(capacity)
                .map_err(|_| litchi_ograph::Error::Allocation {
                    resource: "PPT chart package bytes",
                })?;
            flate2::read::ZlibDecoder::new(storage.stored_bytes())
                .take(u64::from(uncompressed_len).saturating_add(1))
                .read_to_end(&mut bytes)?;
            if bytes.len() != declared {
                return corrupted("compressed chart storage size mismatch");
            }
            Ok(bytes)
        },
    }
}

pub(super) fn parse(
    presentation: &Presentation,
    persist_id: u32,
    kind: Kind,
    limits: Limits,
) -> Result<Parsed> {
    let storage = presentation.ole_storage(persist_id)?.ok_or_else(|| {
        Error::Corrupted(format!(
            "chart object persist ID {persist_id} has no storage"
        ))
    })?;
    let package_bytes = decode(storage, limits)?;
    match kind {
        Kind::Graph => {
            let package = GraphPackage::with_limits(package_bytes, limits)?;
            let workbook = package.workbook()?.into_bytes();
            let book = Book::with_limits(workbook, limits)?;
            ensure_kind(&book, litchi_ograph::chart::Kind::Graph)?;
            Ok(Parsed::Graph {
                package: Box::new(package),
                book,
            })
        },
        Kind::Excel => {
            let workbook = extract_excel_workbook(package_bytes, limits)?;
            let book = Book::with_limits(workbook, limits)?;
            ensure_kind(&book, litchi_ograph::chart::Kind::Excel)?;
            Ok(Parsed::Excel { book })
        },
    }
}

fn ensure_kind(book: &Book, expected: litchi_ograph::chart::Kind) -> Result<()> {
    for chart in book.charts() {
        if chart?.kind() != expected {
            return corrupted("chart Workbook grammar conflicts with its PPT object kind");
        }
    }
    Ok(())
}

fn extract_excel_workbook(package: Vec<u8>, limits: Limits) -> Result<Vec<u8>> {
    let mut cfb = OleFile::open(Cursor::new(package))?;
    let stream = {
        let entries = cfb.list_directory_entries(&[])?;
        check_limit(
            "chart package root entries",
            entries.len(),
            MAX_EXCEL_ROOT_ENTRIES,
        )?;
        let mut stream = None;
        for entry in entries {
            if entry.size > as_u64(limits.max_stream_bytes) {
                return Err(limit_error(
                    "chart package stream bytes",
                    entry.size,
                    as_u64(limits.max_stream_bytes),
                ));
            }
            let candidate =
                entry.name.eq_ignore_ascii_case(WORKBOOK) || entry.name.eq_ignore_ascii_case(BOOK);
            if !candidate {
                continue;
            }
            if entry.entry_type != STGTY_STREAM {
                return corrupted("Excel chart Workbook entry is not a stream");
            }
            if stream.replace(entry.name.clone()).is_some() {
                return corrupted("Excel chart package has multiple Workbook streams");
            }
            if entry.size > as_u64(limits.max_workbook_bytes) {
                return Err(limit_error(
                    "Workbook bytes",
                    entry.size,
                    as_u64(limits.max_workbook_bytes),
                ));
            }
        }
        stream.ok_or_else(|| Error::Corrupted("chart package has no Workbook stream".into()))?
    };
    let workbook = cfb.open_stream(&[stream.as_str()])?;
    check_limit("Workbook bytes", workbook.len(), limits.max_workbook_bytes)?;
    Ok(workbook)
}

fn check_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(limit_error(resource, as_u64(observed), as_u64(maximum)));
    }
    Ok(())
}

fn limit_error(resource: &'static str, observed: u64, maximum: u64) -> Error {
    litchi_ograph::Error::LimitExceeded {
        resource,
        observed,
        maximum,
    }
    .into()
}

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

pub(super) fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
