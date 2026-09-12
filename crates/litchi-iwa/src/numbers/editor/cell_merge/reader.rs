//! Migration-only handoff from cached archives to semantic Numbers reads.

use std::sync::Arc;

use litchi_iwa_archive::{ComponentCatalog, Limits};
use litchi_numbers::{
    MergeReader, PackageReadOptions, PackageSemanticLimits, SheetSelector, TableMergesError,
    TableMergesLimitKind, TableSelector, table::merge::Region,
};

use super::super::{NumbersEditor, selectors};
use crate::{Error, IWorkPackage, Result};

/// Source-built and explicitly unrooted/legacy tables retain the compatibility
/// reader. Every failure after a rooted canonical selection remains terminal.
pub(crate) fn regions_in_editor(editor: &NumbersEditor, table_id: u64) -> Result<Vec<Region>> {
    if editor.package().exact_source_owner().is_none() {
        return super::regions_in_package(editor.package(), table_id);
    }
    let Some((sheet, table)) = selectors::focused_merge_table_indices(editor, table_id)? else {
        return super::regions_in_package(editor.package(), table_id);
    };
    let reader = cached_reader(editor.package())?;
    reader
        .table_merges(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(Error::from)
}

fn cached_reader(package: &IWorkPackage) -> Result<MergeReader> {
    let source_limits = package.limits();
    let limits = Limits::new(
        source_limits.max_input_bytes(),
        source_limits.max_entries(),
        source_limits.max_entry_bytes(),
        source_limits.max_total_bytes(),
        source_limits.max_iwa_stream_bytes(),
    )
    .and_then(|limits| {
        limits.with_archive_limits(
            source_limits
                .effective_archive_limits()
                .map_err(|error| litchi_iwa_archive::Error::InvalidLimits(error.to_string()))?,
        )
    })
    .map_err(map_catalog_error)?;
    let mut records = Vec::new();
    for name in package.iwa_entry_names() {
        if records.len() >= limits.max_entries() {
            return Err(TableMergesError::LimitExceeded {
                kind: TableMergesLimitKind::Entries,
                observed: records.len().saturating_add(1) as u64,
                maximum: limits.max_entries() as u64,
            }
            .into());
        }
        records.try_reserve(1).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers shared merge components",
                amount: 1,
            })
        })?;
        records.push((name, package.parsed_archive(name)?));
    }
    let components =
        ComponentCatalog::__from_shared_archives(records, limits).map_err(map_catalog_error)?;
    MergeReader::__from_shared_catalog(
        Arc::new(components),
        PackageReadOptions::new(limits, PackageSemanticLimits::default()),
    )
    .map_err(Error::from)
}

fn map_catalog_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Iwa(error) => error.into(),
        litchi_iwa_archive::Error::Allocation { resource, amount } => {
            litchi_iwa_common::Error::Allocation { resource, amount }.into()
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            use litchi_iwa_archive::LimitKind;
            let kind = match kind {
                LimitKind::InputBytes => TableMergesLimitKind::InputBytes,
                LimitKind::OutputBytes => TableMergesLimitKind::OutputBytes,
                LimitKind::Entries => TableMergesLimitKind::Entries,
                LimitKind::MemberNameBytes | LimitKind::MetadataBytes => {
                    TableMergesLimitKind::PackageBytes
                },
                LimitKind::CompressedEntryBytes | LimitKind::EntryBytes => {
                    TableMergesLimitKind::EntryBytes
                },
                LimitKind::TotalBytes => TableMergesLimitKind::TotalEntryBytes,
                LimitKind::IwaStreamBytes => TableMergesLimitKind::PayloadBytes,
                LimitKind::IwaTotalBytes => TableMergesLimitKind::TotalPayloadBytes,
            };
            TableMergesError::LimitExceeded {
                kind,
                observed,
                maximum,
            }
            .into()
        },
        error => error.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersEditor;

    const NATIVE_NUMBERS: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/iwork/numbers/table-merges-native.numbers"
    ));

    fn canonical_model_id(editor: &NumbersEditor) -> Result<u64> {
        for name in editor.package().iwa_entry_names() {
            let archive = editor.package().archive(name)?;
            for object in archive.objects {
                if object.messages.iter().any(|message| message.type_ == 6_001) {
                    return object.archive_info.identifier.ok_or_else(|| {
                        Error::InvalidFormat(
                            "native Numbers model object has no identifier".to_owned(),
                        )
                    });
                }
            }
        }
        Err(Error::InvalidFormat(
            "native Numbers fixture has no canonical table model".to_owned(),
        ))
    }

    #[test]
    fn cached_reader_uses_shared_catalog_for_native_table_merges() -> Result<()> {
        let editor = NumbersEditor::from_bytes(NATIVE_NUMBERS)?;
        let model_id = canonical_model_id(&editor)?;
        let (sheet, table) = selectors::focused_merge_table_indices(&editor, model_id)?
            .ok_or_else(|| Error::InvalidFormat("native table was not rooted".to_owned()))?;

        let package = editor.into_package();
        let record = package.parsed_archive("Index/Document.iwa")?;
        let cache_owners = Arc::strong_count(&record);
        let reader = cached_reader(&package)?;
        assert_eq!(
            Arc::strong_count(&record),
            cache_owners + 1,
            "the shared catalog must retain the cached Archive allocation"
        );
        let regions = reader
            .table_merges(SheetSelector::index(sheet), TableSelector::index(table))
            .map_err(Error::from)?;
        assert_eq!(regions, [Region::new(10, 1, 2, 2)?]);
        drop(reader);
        assert_eq!(
            Arc::strong_count(&record),
            cache_owners,
            "dropping the focused reader must release its shared Archive owner"
        );
        Ok(())
    }

    #[test]
    fn cached_reader_does_not_require_exact_source_bytes() -> Result<()> {
        let editor = NumbersEditor::from_bytes(NATIVE_NUMBERS)?;
        let model_id = canonical_model_id(&editor)?;
        let (sheet, table) = selectors::focused_merge_table_indices(&editor, model_id)?
            .ok_or_else(|| Error::InvalidFormat("native table was not rooted".to_owned()))?;
        let mut package = editor.into_package();

        // Build the shared reader once so all component Arcs are resident in
        // the package cache, then discard only the exact ZIP provenance. A
        // focused handoff that reparsed source bytes would no longer have an
        // ingress allocation to borrow at this point.
        let _ = cached_reader(&package)?;
        package.discard_exact_source_for_compatibility();
        let reader = cached_reader(&package)?;
        let regions = reader
            .table_merges(SheetSelector::index(sheet), TableSelector::index(table))
            .map_err(Error::from)?;
        assert_eq!(regions, [Region::new(10, 1, 2, 2)?]);
        Ok(())
    }
}
