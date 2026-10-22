//! Worksheet-stream package facade for the Scenario Manager.
//!
//! The facade splices only the balanced scenario collection. Every record
//! outside that span remains byte-for-byte in its original position. A
//! parsed manager containing opaque records can be read and round-tripped,
//! but semantic replacement/removal is refused because the safe placement of
//! an unknown structural record cannot be inferred.

use super::codec;
use super::model::{MAX_CHANGED_CELLS, MAX_SCENARIOS, MAX_UNKNOWN_RECORDS, Manager};
use crate::package::error::{Error, Result};
use crate::raw::{Record, Records, kind};

const MAX_MANAGER_RECORDS: usize = MAX_SCENARIOS * (MAX_CHANGED_CELLS + 2) + MAX_UNKNOWN_RECORDS;

struct Block<'a> {
    begin: usize,
    header: &'a [u8],
    records: Vec<Record<'a>>,
    end: usize,
}

fn locate<'a>(data: &'a [u8], limits: crate::raw::Limits) -> Result<Option<Block<'a>>> {
    let mut iterator = Records::new(data, limits);
    let mut begin = None;
    let mut header = None;
    let mut inner = Vec::new();
    let mut end = None;

    while let Some(result) = iterator.next() {
        let record = result?;
        if end.is_some() {
            if record.kind() == codec::begin_manager() {
                return Err(Error::Unrecognized {
                    typ: "Scenario Manager collection",
                    val: "multiple BrtBeginScenMan records",
                });
            }
            if record.kind() == codec::end_manager() {
                return Err(Error::Unrecognized {
                    typ: "Scenario Manager collection",
                    val: "multiple BrtEndScenMan records",
                });
            }
            continue;
        }

        match begin {
            None if record.kind() == codec::begin_manager() => {
                begin = Some(record.offset());
                header = Some(record.payload());
            },
            None if record.kind() == codec::end_manager() => {
                return Err(Error::Unrecognized {
                    typ: "Scenario Manager collection",
                    val: "BrtEndScenMan has no matching BrtBeginScenMan",
                });
            },
            None => {},
            Some(_) if record.kind() == codec::begin_manager() => {
                return Err(Error::Unrecognized {
                    typ: "Scenario Manager collection",
                    val: "multiple BrtBeginScenMan records",
                });
            },
            Some(_) if record.kind() == codec::end_manager() => {
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                end = Some(iterator.offset());
            },
            Some(_) => {
                if inner.len() >= MAX_MANAGER_RECORDS {
                    return Err(Error::Unrecognized {
                        typ: "Scenario Manager collection",
                        val: "record count exceeds",
                    });
                }
                inner.push(record);
            },
        }
    }

    let Some(begin) = begin else {
        return Ok(None);
    };
    let Some(end) = end else {
        return Err(Error::UnexpectedEndOfStream("BrtEndScenMan"));
    };
    let Some(header) = header else {
        return Err(Error::Unrecognized {
            typ: "Scenario Manager collection",
            val: "BrtBeginScenMan has no payload",
        });
    };
    Ok(Some(Block {
        begin,
        header,
        records: inner,
        end,
    }))
}

fn end_sheet_offset(data: &[u8], limits: crate::raw::Limits) -> Result<usize> {
    let mut iterator = Records::new(data, limits);
    let mut offset = None;
    while let Some(result) = iterator.next() {
        let record = result?;
        if record.kind() == kind::END_SHEET {
            if offset.replace(record.offset()).is_some() {
                return Err(Error::Unrecognized {
                    typ: "Scenario Manager insertion",
                    val: "worksheet must contain exactly one BrtEndSheet",
                });
            }
        }
    }
    offset.ok_or(Error::Unrecognized {
        typ: "Scenario Manager insertion",
        val: "worksheet must contain exactly one BrtEndSheet",
    })
}

/// Parse the worksheet's Scenario Manager collection, if present.
pub fn parse_worksheet(data: &[u8], limits: crate::raw::Limits) -> Result<Option<Manager>> {
    let Some(block) = locate(data, limits)? else {
        return Ok(None);
    };
    let manager = codec::parse_manager(block.header, &block.records, limits)?;
    Ok(Some(manager))
}

/// Replace or remove the worksheet's Scenario Manager collection.
///
/// Passing `None` removes the collection. Passing `Some` inserts it before
/// `BrtEndSheet` when absent, or replaces the existing balanced collection.
/// Unknown records in an existing manager are retained only for an unchanged
/// round trip; changing that manager is rejected as unsafe.
pub fn replace_worksheet(
    data: &[u8],
    manager: Option<&Manager>,
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    let block = locate(data, limits)?;
    let Some(block) = block else {
        let Some(manager) = manager else {
            return Ok(data.to_vec());
        };
        let encoded = codec::write_manager(manager, limits)?;
        let offset = end_sheet_offset(data, limits)?;
        let capacity = data
            .len()
            .checked_add(encoded.len())
            .ok_or(Error::InvalidLength {
                expected: usize::MAX,
                found: data.len(),
            })?;
        let mut output = Vec::with_capacity(capacity);
        output.extend_from_slice(&data[..offset]);
        output.extend_from_slice(&encoded);
        output.extend_from_slice(&data[offset..]);
        return Ok(output);
    };

    let original = codec::parse_manager(block.header, &block.records, limits)?;
    if manager.is_some_and(|candidate| candidate == &original) {
        return Ok(data.to_vec());
    }
    if original.is_opaque() {
        if manager.is_some_and(|candidate| candidate != &original) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_scenarios",
                "the Scenario Manager source contains opaque records",
            ));
        }
        if manager.is_none() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_scenarios",
                "the Scenario Manager source contains opaque records",
            ));
        }
        return Ok(data.to_vec());
    }

    let start = block.begin;
    let end = block.end;
    let replacement = manager
        .map(|manager| codec::write_manager(manager, limits))
        .transpose()?;
    let replacement_len = replacement.as_ref().map_or(0, Vec::len);
    let removed = end.checked_sub(start).ok_or(Error::InvalidLength {
        expected: 0,
        found: end,
    })?;
    let capacity = data
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or(Error::InvalidLength {
            expected: usize::MAX,
            found: data.len(),
        })?;
    let mut output = Vec::with_capacity(capacity);
    output.extend_from_slice(&data[..start]);
    if let Some(replacement) = replacement {
        output.extend_from_slice(&replacement);
    }
    output.extend_from_slice(&data[end..]);
    Ok(output)
}
