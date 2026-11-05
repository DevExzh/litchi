//! Shared-string closure for package-owned scalar cell edits.
//!
//! A `BrtCellIsst` value is an index into a workbook-global table.  Changing
//! one such cell to a direct scalar therefore has a package-wide consequence:
//! the occurrence count and the table's entry count have to describe every
//! remaining worksheet reference, and removing an unused entry requires
//! remapping all later indexes.  This module keeps that operation separate
//! from the worksheet stream editor so that the latter can continue to copy
//! opaque records byte for byte.

#![allow(
    clippy::cast_lossless,
    reason = "BIFF12 count fields are checked below before their wire conversion"
)]

use super::Limits;
use crate::package::error::{Error, Result};
use crate::raw::{Header, Kind, Limits as RawLimits, Records, kind};
use crate::workbook::Workbook;
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI};

const MAX_SST_COUNT: u32 = 0x7fff_ffff;
const SHARED_STRING_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";

/// Reconcile the workbook-global shared-string table after a scalar edit.
///
/// The caller invokes this only when the semantic edit changed the number of
/// `BrtCellIsst` references.  The input is consumed because the caller has
/// already received a validated candidate workbook; this function clones no
/// caller-owned package and returns a newly validated workbook under the same
/// external-link, drawing, leniency, and source-read policies.
///
/// Uses the worksheet limits selected by the owning edit.
///
/// Keeping this seam explicit lets package-owned edits apply their worksheet
/// record and source ceilings to the second, package-wide scan as well as to
/// the initial worksheet snapshot.
pub(super) fn reconcile_with_limits(
    workbook: Workbook,
    worksheet_limits: Limits,
    reference_delta: i64,
) -> Result<Workbook> {
    let external_link_limits = workbook.external_link_limits();
    let drawing_load_policy = workbook.drawing_load_policy();
    let leniency = workbook.leniency();
    let raw_limits = workbook.raw_limits_value();
    let mut worksheet_uris = Vec::new();
    worksheet_uris
        .try_reserve_exact(workbook.worksheet_count())
        .map_err(|source| Error::Allocation {
            resource: "shared-string worksheet URI catalog",
            source,
        })?;
    for index in 0..workbook.worksheet_count() {
        worksheet_uris.push(workbook.worksheet_uri(index)?);
    }
    let mut package = workbook.into_opc_package();
    validate_worksheet_uris(&package, &worksheet_uris)?;
    let shared_strings_uri = shared_strings_uri(&package)?;
    let shared_source = package.get_part(&shared_strings_uri)?.blob().to_vec();
    let table = SharedStringTable::parse(&shared_source, worksheet_limits.raw())?;

    let mut references = zeroed_vec(table.items.len(), 0_u64, "shared-string reference counts")?;
    for worksheet_uri in &worksheet_uris {
        let worksheet = package.get_part(worksheet_uri)?.blob();
        scan_worksheet(
            worksheet,
            worksheet_limits.source_bytes(),
            worksheet_limits.raw(),
            &mut references,
        )?;
    }

    let actual = references.iter().try_fold(0_u64, |sum, count| {
        sum.checked_add(*count).ok_or(Error::CapacityOverflow {
            resource: "shared-string occurrence count",
        })
    })?;
    if i128::from(table.total) + i128::from(reference_delta) != i128::from(actual) {
        return Err(Error::InvalidFormat(
            "shared-string source count does not match worksheet references",
        ));
    }
    let mapping = table.mapping(&references)?;
    let rewritten_sst = table.rewrite(
        &shared_source,
        worksheet_limits.raw(),
        &references,
        &mapping,
    )?;

    let mut replacement_sizes = Vec::new();
    replacement_sizes
        .try_reserve_exact(worksheet_uris.len().saturating_add(1))
        .map_err(|source| Error::Allocation {
            resource: "shared-string package replacement plan",
            source,
        })?;
    replacement_sizes.push((shared_strings_uri.clone(), rewritten_sst.len()));
    for worksheet_uri in &worksheet_uris {
        let source = package.get_part(worksheet_uri)?.blob();
        if source.len() > worksheet_limits.source_bytes() {
            return Err(Error::Limit {
                kind: crate::cell_values::LimitKind::SourceBytes.into(),
                observed: source.len(),
                maximum: worksheet_limits.source_bytes(),
            });
        }
        replacement_sizes.push((worksheet_uri.clone(), source.len()));
    }
    check_replacement_limits(&package, &replacement_sizes)?;

    package
        .get_part_mut(&shared_strings_uri)?
        .set_blob(rewritten_sst);
    for worksheet_uri in worksheet_uris {
        let source = package.get_part(&worksheet_uri)?.blob();
        let rewritten = rewrite_worksheet(
            source,
            worksheet_limits.source_bytes(),
            worksheet_limits.raw(),
            &mapping,
        )?;
        if rewritten.len() > worksheet_limits.source_bytes() {
            return Err(Error::Limit {
                kind: crate::cell_values::LimitKind::SourceBytes.into(),
                observed: rewritten.len(),
                maximum: worksheet_limits.source_bytes(),
            });
        }
        if source != rewritten.as_slice() {
            package.get_part_mut(&worksheet_uri)?.set_blob(rewritten);
        }
    }

    let candidate = Workbook::reparse_candidate_with_policy(
        package,
        external_link_limits,
        drawing_load_policy,
        leniency,
        raw_limits,
    )?;
    super::structure::validate_all_worksheets(&candidate)?;
    Ok(candidate)
}

fn validate_worksheet_uris(package: &OpcPackage, uris: &[PackURI]) -> Result<()> {
    // The temporary workbook above has already checked the complete sheet
    // graph, but confirm every target is a worksheet part before mutating it.
    for uri in uris {
        let part = package.get_part(uri)?;
        if part.content_type() != "application/vnd.ms-excel.worksheet" {
            return Err(Error::InvalidContentType {
                expected: "application/vnd.ms-excel.worksheet",
            });
        }
    }
    Ok(())
}

fn shared_strings_uri(package: &OpcPackage) -> Result<PackURI> {
    let workbook = package.main_document_part()?;
    let mut found = None;
    for relationship in workbook.rels().iter() {
        if !matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::SHARED_STRINGS
                | litchi_opc::constants::relationship_type::STRICT_SHARED_STRINGS
        ) {
            continue;
        }
        if relationship.is_external() {
            return Err(Error::InvalidRelationship(
                "shared-string relationship is external",
            ));
        }
        if found.is_some() {
            return Err(Error::InvalidRelationship(
                "workbook has multiple shared-string relationships",
            ));
        }
        found = Some(relationship.target_partname()?);
    }
    let uri = found.ok_or(Error::InvalidRelationship(
        "workbook has no shared-string relationship",
    ))?;
    let part = package.get_part(&uri)?;
    if part.content_type() != SHARED_STRING_CONTENT_TYPE {
        return Err(Error::InvalidContentType {
            expected: SHARED_STRING_CONTENT_TYPE,
        });
    }
    Ok(uri)
}

#[derive(Debug)]
struct SharedStringTable {
    total: u32,
    items: Vec<()>,
}

impl SharedStringTable {
    fn parse(source: &[u8], limits: RawLimits) -> Result<Self> {
        let mut total = None;
        let mut unique = None;
        let mut items = Vec::new();
        let mut ended = false;
        let mut in_table = false;
        for item in Records::new(source, limits) {
            let record = item?;
            match record.kind() {
                kind::BEGIN_SST => {
                    if in_table || total.is_some() {
                        return Err(Error::InvalidFormat(
                            "shared-string table has duplicate BrtBeginSst records",
                        ));
                    }
                    if record.payload().len() != 8 {
                        return Err(Error::InvalidLength {
                            expected: 8,
                            found: record.payload().len(),
                        });
                    }
                    let values = record.payload();
                    let occurrence_count =
                        u32::from_le_bytes(values[..4].try_into().map_err(|_error| {
                            Error::InvalidLength {
                                expected: 4,
                                found: values.len(),
                            }
                        })?);
                    let entry_count =
                        u32::from_le_bytes(values[4..8].try_into().map_err(|_error| {
                            Error::InvalidLength {
                                expected: 8,
                                found: values.len(),
                            }
                        })?);
                    if occurrence_count > MAX_SST_COUNT || entry_count > occurrence_count {
                        return Err(Error::InvalidFormat("invalid shared-string counts"));
                    }
                    total = Some(occurrence_count);
                    unique = Some(entry_count);
                    in_table = true;
                },
                kind::SST_ITEM if in_table => {
                    items.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "shared-string table entries",
                        source,
                    })?;
                    items.push(());
                },
                kind::END_SST if in_table => {
                    if !record.payload().is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    ended = true;
                    in_table = false;
                },
                kind::SST_ITEM if !in_table && ended => {
                    return Err(Error::InvalidFormat(
                        "shared-string item appears after BrtEndSst",
                    ));
                },
                kind::SST_ITEM if !in_table => {
                    return Err(Error::InvalidFormat(
                        "shared-string item appears before BrtBeginSst",
                    ));
                },
                _ => {},
            }
        }
        let total = total.ok_or(Error::InvalidFormat(
            "shared-string table has no BrtBeginSst",
        ))?;
        let unique = unique.ok_or(Error::InvalidFormat(
            "shared-string table has no unique count",
        ))?;
        if !ended {
            return Err(Error::InvalidFormat("shared-string table has no BrtEndSst"));
        }
        let found = u32::try_from(items.len()).map_err(|_error| {
            Error::limit(
                crate::cell_values::LimitKind::SharedStrings,
                items.len(),
                usize::try_from(MAX_SST_COUNT).unwrap_or(usize::MAX),
            )
        })?;
        if found != unique {
            return Err(Error::InvalidFormat(
                "shared-string unique count declares entries but stream contains",
            ));
        }
        Ok(Self { total, items })
    }

    fn mapping(&self, references: &[u64]) -> Result<Vec<Option<u32>>> {
        if references.len() != self.items.len() {
            return Err(Error::InvalidFormat(
                "shared-string reference scan does not match the table entry count",
            ));
        }
        let occurrence_count = references.iter().try_fold(0_u64, |sum, count| {
            sum.checked_add(*count).ok_or(Error::CapacityOverflow {
                resource: "shared-string occurrence count",
            })
        })?;
        let occurrence_count = u32::try_from(occurrence_count).map_err(|_error| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "shared-string occurrence count exceeds the wire range",
            )
        })?;
        if occurrence_count > MAX_SST_COUNT {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "shared-string occurrence count exceeds the wire range",
            ));
        }
        let live = references.iter().filter(|count| **count != 0).count();
        let capacity = usize::try_from(occurrence_count)
            .unwrap_or(usize::MAX)
            .min(self.items.len());
        if live > capacity {
            return Err(Error::InvalidFormat(
                "shared-string live entry count exceeds the occurrence count",
            ));
        }

        // Existing unreferenced entries are valid source data. Keep all of
        // them while the BIFF12 invariant still permits it; when total falls
        // below the old unique count, remove only the excess at the highest
        // indexes so the indexes of earlier live entries remain stable.
        let mut keep = zeroed_vec(
            self.items.len(),
            false,
            "shared-string retained entry flags",
        )?;
        for (index, count) in references.iter().enumerate() {
            if *count != 0 {
                keep[index] = true;
            }
        }
        let mut extras = capacity.saturating_sub(live);
        for index in 0..self.items.len() {
            if extras == 0 {
                break;
            }
            if !keep[index] {
                keep[index] = true;
                extras = extras.saturating_sub(1);
            }
        }

        let mut mapping = Vec::new();
        mapping
            .try_reserve_exact(self.items.len())
            .map_err(|source| Error::Allocation {
                resource: "shared-string index remap",
                source,
            })?;
        let mut next = 0_u32;
        for retained in keep {
            if !retained {
                mapping.push(None);
            } else {
                mapping.push(Some(next));
                next = next.checked_add(1).ok_or_else(|| {
                    Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit_cell_values",
                        "shared-string unique count exceeds the wire range",
                    )
                })?;
            }
        }
        if next > MAX_SST_COUNT {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "shared-string unique count exceeds the wire range",
            ));
        }
        Ok(mapping)
    }

    fn rewrite(
        &self,
        source: &[u8],
        limits: RawLimits,
        references: &[u64],
        mapping: &[Option<u32>],
    ) -> Result<Vec<u8>> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(source.len())
            .map_err(|source| Error::Allocation {
                resource: "shared-string table rewrite",
                source,
            })?;
        let mut item_index = 0_usize;
        let mut in_table = false;
        let mut saw_begin = false;
        let mut saw_end = false;
        let kept_unique = mapping.iter().filter(|value| value.is_some()).count();
        let kept_unique = u32::try_from(kept_unique).map_err(|_error| {
            Error::limit(
                crate::cell_values::LimitKind::SharedStrings,
                kept_unique,
                usize::try_from(MAX_SST_COUNT).unwrap_or(usize::MAX),
            )
        })?;
        let occurrence_count = references.iter().try_fold(0_u64, |sum, count| {
            sum.checked_add(*count).ok_or(Error::CapacityOverflow {
                resource: "shared-string occurrence count",
            })
        })?;
        let occurrence_count = u32::try_from(occurrence_count).map_err(|_error| {
            Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_cell_values",
                "shared-string occurrence count exceeds the wire range",
            )
        })?;
        if occurrence_count > MAX_SST_COUNT || kept_unique > occurrence_count {
            return Err(Error::InvalidFormat(
                "invalid reconciled shared-string counts",
            ));
        }

        for item in Records::new(source, limits) {
            let record = item?;
            match record.kind() {
                kind::BEGIN_SST => {
                    if saw_begin || record.payload().len() != 8 {
                        return Err(Error::InvalidFormat(
                            "shared-string table has an invalid BrtBeginSst record",
                        ));
                    }
                    saw_begin = true;
                    in_table = true;
                    let mut payload = record.payload().to_vec();
                    payload[..4].copy_from_slice(&occurrence_count.to_le_bytes());
                    payload[4..8].copy_from_slice(&kept_unique.to_le_bytes());
                    append_record_with_payload(source, &record, &payload, limits, &mut output)?;
                },
                kind::SST_ITEM if in_table => {
                    let keep = mapping.get(item_index).and_then(|value| *value).is_some();
                    if keep {
                        append_record(source, &record, limits, &mut output)?;
                    }
                    item_index = item_index.checked_add(1).ok_or(Error::CapacityOverflow {
                        resource: "shared-string item index",
                    })?;
                },
                kind::END_SST if in_table => {
                    if !record.payload().is_empty() {
                        return Err(Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    in_table = false;
                    saw_end = true;
                    append_record(source, &record, limits, &mut output)?;
                },
                kind::SST_ITEM if saw_end => {
                    return Err(Error::InvalidFormat(
                        "shared-string item appears after BrtEndSst",
                    ));
                },
                _ => append_record(source, &record, limits, &mut output)?,
            }
        }
        if !saw_begin || !saw_end || in_table || item_index != self.items.len() {
            return Err(Error::InvalidFormat(
                "shared-string table structure changed during rewrite",
            ));
        }
        Ok(output)
    }
}

fn scan_worksheet(
    source: &[u8],
    source_limit: usize,
    limits: RawLimits,
    references: &mut [u64],
) -> Result<()> {
    if source.len() > source_limit {
        return Err(Error::limit(
            crate::cell_values::LimitKind::SourceBytes,
            source.len(),
            source_limit,
        ));
    }
    let mut in_sheet_data = false;
    let mut frt_depth = 0_usize;
    for item in Records::new(source, limits) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => {
                if in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "worksheet has duplicate BrtBeginSheetData records",
                    ));
                }
                in_sheet_data = true;
                frt_depth = 0;
            },
            kind::END_SHEET_DATA => {
                if !in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "worksheet has BrtEndSheetData without BrtBeginSheetData",
                    ));
                }
                in_sheet_data = false;
                frt_depth = 0;
            },
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => frt_depth = frt_depth.saturating_sub(1),
            kind::CELL_ISST if in_sheet_data && frt_depth != 0 => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit_cell_values",
                    "shared-string closure cannot interpret cell records in an opaque FRT block",
                ));
            },
            kind::CELL_ISST if in_sheet_data => {
                let index = cell_isst_index(record.payload())?;
                let slot = usize::try_from(index).map_err(|_error| {
                    Error::InvalidCellReference("shared-string index is invalid")
                })?;
                let count = references.get_mut(slot).ok_or(Error::InvalidCellReference(
                    "worksheet references missing shared-string entry",
                ))?;
                *count = count.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "shared-string occurrence count",
                })?;
            },
            candidate
                if in_sheet_data && frt_depth == 0 && is_unsupported_cell_family(candidate) =>
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit_cell_values",
                    "shared-string remapping refuses unsupported worksheet cell family",
                ));
            },
            _ => {},
        }
    }
    if in_sheet_data {
        return Err(Error::InvalidFormat(
            "worksheet stream ended before BrtEndSheetData",
        ));
    }
    Ok(())
}

fn rewrite_worksheet(
    source: &[u8],
    source_limit: usize,
    limits: RawLimits,
    mapping: &[Option<u32>],
) -> Result<Vec<u8>> {
    if source.len() > source_limit {
        return Err(Error::limit(
            crate::cell_values::LimitKind::SourceBytes,
            source.len(),
            source_limit,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "shared-string worksheet rewrite",
            source,
        })?;
    let mut in_sheet_data = false;
    let mut frt_depth = 0_usize;
    for item in Records::new(source, limits) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => {
                if in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "worksheet has duplicate BrtBeginSheetData records",
                    ));
                }
                in_sheet_data = true;
                frt_depth = 0;
                append_record(source, &record, limits, &mut output)?;
            },
            kind::END_SHEET_DATA => {
                if !in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "worksheet has BrtEndSheetData without BrtBeginSheetData",
                    ));
                }
                in_sheet_data = false;
                frt_depth = 0;
                append_record(source, &record, limits, &mut output)?;
            },
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
                append_record(source, &record, limits, &mut output)?;
            },
            kind::FRT_END if in_sheet_data => {
                frt_depth = frt_depth.saturating_sub(1);
                append_record(source, &record, limits, &mut output)?;
            },
            kind::CELL_ISST if in_sheet_data && frt_depth != 0 => {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit_cell_values",
                    "shared-string closure cannot interpret cell records in an opaque FRT block",
                ));
            },
            kind::CELL_ISST if in_sheet_data => {
                let index = cell_isst_index(record.payload())?;
                let old_index = usize::try_from(index).map_err(|_error| {
                    Error::InvalidCellReference("shared-string index is invalid")
                })?;
                let new_index = mapping.get(old_index).and_then(|value| *value).ok_or(
                    Error::InvalidCellReference("worksheet references removed shared-string entry"),
                )?;
                let mut payload = record.payload().to_vec();
                payload[8..12].copy_from_slice(&new_index.to_le_bytes());
                append_record_with_payload(source, &record, &payload, limits, &mut output)?;
            },
            candidate
                if in_sheet_data && frt_depth == 0 && is_unsupported_cell_family(candidate) =>
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "edit_cell_values",
                    "shared-string remapping refuses unsupported worksheet cell family",
                ));
            },
            _ => append_record(source, &record, limits, &mut output)?,
        }
    }
    if in_sheet_data {
        return Err(Error::InvalidFormat(
            "worksheet stream ended before BrtEndSheetData",
        ));
    }
    Ok(output)
}

fn cell_isst_index(payload: &[u8]) -> Result<u32> {
    if payload.len() != 12 {
        return Err(Error::InvalidLength {
            expected: 12,
            found: payload.len(),
        });
    }
    Ok(u32::from_le_bytes(payload[8..12].try_into().map_err(
        |_error| Error::InvalidLength {
            expected: 12,
            found: payload.len(),
        },
    )?))
}

fn is_unsupported_cell_family(kind_value: Kind) -> bool {
    // The BIFF12 direct and formula cell families occupy 0x0001..=0x000B;
    // 0x000C..=0x0012 are reserved for additional cell families in versions
    // of the format this crate does not decode.  Refuse these explicitly when
    // they occur at the sheet-data cell boundary rather than guessing where an
    // SST index might live.  Other records are copied as opaque worksheet data.
    (0x000C..=0x0012).contains(&kind_value.get())
}

fn append_record(
    source: &[u8],
    record: &crate::raw::Record<'_>,
    limits: RawLimits,
    output: &mut Vec<u8>,
) -> Result<()> {
    let range = record_range(source, record, limits)?;
    output.extend_from_slice(range);
    Ok(())
}

fn append_record_with_payload(
    source: &[u8],
    record: &crate::raw::Record<'_>,
    payload: &[u8],
    limits: RawLimits,
    output: &mut Vec<u8>,
) -> Result<()> {
    let record_source = source
        .get(record.offset()..)
        .ok_or(Error::InvalidFormat("record offset is outside source"))?;
    let (header, header_len) = Header::parse(record_source, limits)?;
    if header.len() != payload.len() {
        return Err(Error::InvalidLength {
            expected: header.len(),
            found: payload.len(),
        });
    }
    let header_bytes = record_source
        .get(..header_len)
        .ok_or(Error::InvalidFormat("record header is outside source"))?;
    output.extend_from_slice(header_bytes);
    output.extend_from_slice(payload);
    Ok(())
}

fn record_range<'a>(
    source: &'a [u8],
    record: &crate::raw::Record<'_>,
    limits: RawLimits,
) -> Result<&'a [u8]> {
    let record_source = source
        .get(record.offset()..)
        .ok_or(Error::InvalidFormat("record offset is outside source"))?;
    let (header, header_len) = Header::parse(record_source, limits)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|value| value.checked_add(header.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "shared-string record range",
        })?;
    source
        .get(record.offset()..end)
        .ok_or(Error::InvalidFormat("record range is outside source"))
}

fn check_replacement_limits(package: &OpcPackage, replacements: &[(PackURI, usize)]) -> Result<()> {
    let limits = package.read_limits();
    let maximum_part = limits.max_part_bytes();
    let mut total = 0_u64;
    for part in package.try_iter_parts() {
        let part = part?;
        let size = replacements
            .iter()
            .find(|(target, _)| part.partname().is_equivalent_to(target))
            .map_or(part.blob().len(), |(_, size)| *size);
        let size_u64 = u64::try_from(size).map_err(|_error| Error::CapacityOverflow {
            resource: "shared-string total part bytes",
        })?;
        if size_u64 > maximum_part {
            return Err(Error::limit(
                crate::cell_values::LimitKind::SharedStringBytes,
                size,
                usize::try_from(maximum_part).unwrap_or(usize::MAX),
            ));
        }
        total = total.checked_add(size_u64).ok_or(Error::CapacityOverflow {
            resource: "shared-string total part bytes",
        })?;
    }
    if total > limits.max_total_part_bytes() {
        return Err(Error::limit(
            crate::cell_values::LimitKind::SharedStringBytes,
            usize::try_from(total).unwrap_or(usize::MAX),
            usize::try_from(limits.max_total_part_bytes()).unwrap_or(usize::MAX),
        ));
    }
    Ok(())
}

fn zeroed_vec<T: Clone>(length: usize, value: T, resource: &'static str) -> Result<Vec<T>> {
    let mut result = Vec::new();
    result
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation { resource, source })?;
    result.resize(length, value);
    Ok(result)
}
