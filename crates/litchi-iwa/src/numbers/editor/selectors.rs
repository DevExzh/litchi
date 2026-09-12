//! Semantic Numbers selector resolution at the native archive boundary.
//!
//! The public editor accepts only archive-free selectors from
//! `litchi_numbers`. Native object identifiers are resolved once here and are
//! kept below the semantic API.

use std::collections::{HashMap, HashSet};

use super::NumbersEditor;
use crate::archive::ArchiveObject;
use crate::{Error, IWorkPackage, Result};
use litchi_numbers::{Dimensions, Document, Sheet, SheetSelector, Table, TableSelector};

/// A private semantic sheet catalog used only at the legacy archive boundary.
///
/// The focused Numbers selector is intentionally archive-free.  The host still
/// has to hand a selected native object to its existing writer, so this adapter
/// keeps the native IDs in a parallel private vector and delegates the actual
/// name/index match to `litchi_numbers::Document::sheet`.
struct SheetSelectorAdapter {
    semantic: Document,
    source_names: Vec<String>,
    native_ids: Vec<u64>,
}

impl SheetSelectorAdapter {
    fn from_editor(editor: &NumbersEditor) -> Result<Self> {
        let source = editor.sheets()?;
        let source_names = source
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let semantic_sheets = selector_safe_names(&source_names, "sheet")
            .into_iter()
            .enumerate()
            .map(|(index, name)| Sheet::new(name, index))
            .collect::<Vec<_>>();
        let semantic = Document::from_sheets(semantic_sheets).map_err(|error| {
            Error::InvalidFormat(format!(
                "Numbers sheet selector catalog is invalid: {error}"
            ))
        })?;
        let native_ids = source
            .into_iter()
            .map(|sheet| sheet.native_id())
            .collect::<Vec<_>>();
        Ok(Self {
            semantic,
            source_names,
            native_ids,
        })
    }

    fn native_id(&self, selector: SheetSelector<'_>) -> Result<u64> {
        ensure_unique_sheet_name(&self.source_names, selector)?;
        let index = self
            .semantic
            .sheet(selector)
            .map_err(|error| {
                Error::InvalidFormat(format!("Numbers sheet selector failed: {error}"))
            })?
            .map(Sheet::index)
            .ok_or_else(|| sheet_selector_error(selector))?;
        self.native_ids.get(index).copied().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet selector catalog lost native entry at index {index}"
            ))
        })
    }
}

/// A private semantic table catalog used only at the legacy archive boundary.
///
/// `TableSelector` is scoped to one focused semantic sheet.  The historical
/// editor API predates that scope and exposes one workbook-wide table catalog,
/// so the adapter presents that catalog as one synthetic semantic sheet.  Its
/// source-name preflight retains the host's cross-sheet ambiguity rule.
struct TableSelectorAdapter {
    semantic: Sheet,
    source_names: Vec<String>,
    native_ids: Vec<u64>,
}

impl TableSelectorAdapter {
    fn from_editor(editor: &NumbersEditor) -> Result<Self> {
        let descriptors = super::table_models(&editor.package)?;
        let source_names = descriptors
            .iter()
            .map(|table| table.model.table_name.clone())
            .collect::<Vec<_>>();
        let semantic_tables = selector_safe_names(&source_names, "table")
            .into_iter()
            .zip(&descriptors)
            .map(|(name, descriptor)| {
                Table::new(
                    name,
                    Dimensions::new(
                        descriptor.model.number_of_rows,
                        descriptor.model.number_of_columns,
                    ),
                )
            })
            .collect::<Vec<_>>();
        let semantic = Sheet::try_from_tables("Numbers table selector catalog", 0, semantic_tables)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers table selector catalog is invalid: {error}"
                ))
            })?;
        let native_ids = descriptors
            .into_iter()
            .map(|table| table.object_id)
            .collect::<Vec<_>>();
        Ok(Self {
            semantic,
            source_names,
            native_ids,
        })
    }

    fn native_id(&self, selector: TableSelector<'_>) -> Result<u64> {
        ensure_unique_table_name(&self.source_names, selector)?;
        let selected = self
            .semantic
            .select(selector)
            .map_err(|error| {
                Error::InvalidFormat(format!("Numbers table selector failed: {error}"))
            })?
            .ok_or_else(|| table_selector_error(selector))?;
        let index = self
            .semantic
            .tables()
            .position(|table| std::ptr::eq(table, selected))
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers table selector catalog lost semantic entry".to_owned(),
                )
            })?;
        self.native_ids.get(index).copied().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table selector catalog lost native entry at index {index}"
            ))
        })
    }
}

#[derive(Debug, Clone, Copy)]
struct FocusedObjectLocation {
    archive_index: usize,
    object_index: usize,
}

/// A borrowed archive-backed object index for one focused selector pass.
///
/// The index retains only archive names and object slots. Archive payloads and
/// decoded objects remain in the package cache and are borrowed through
/// `with_parsed_archive` for the duration of each read.
struct FocusedObjectCatalog<'package> {
    package: &'package IWorkPackage,
    archive_names: Vec<String>,
    objects: HashMap<u64, FocusedObjectLocation>,
}

impl<'package> FocusedObjectCatalog<'package> {
    fn build(package: &'package IWorkPackage, required_ids: &HashSet<u64>) -> Result<Self> {
        let archive_count = package.iwa_entry_names().count();
        let mut archive_names = Vec::new();
        archive_names
            .try_reserve_exact(archive_count)
            .map_err(|_| {
                allocation_error("Numbers focused selector archive names", archive_count)
            })?;
        for name in package.iwa_entry_names() {
            let mut owned = String::new();
            owned.try_reserve_exact(name.len()).map_err(|_| {
                allocation_error("Numbers focused selector archive name", name.len())
            })?;
            owned.push_str(name);
            archive_names.push(owned);
        }

        let mut objects = HashMap::new();
        objects.try_reserve(required_ids.len()).map_err(|_| {
            allocation_error("Numbers focused selector object index", required_ids.len())
        })?;
        for (archive_index, archive_name) in archive_names.iter().enumerate() {
            package.with_parsed_archive(archive_name, |archive| {
                for (object_index, object) in archive.objects.iter().enumerate() {
                    let Some(identifier) = object.archive_info.identifier else {
                        continue;
                    };
                    if !required_ids.contains(&identifier) {
                        continue;
                    }
                    objects.try_reserve(1).map_err(|_| {
                        allocation_error("Numbers focused selector object index", 1)
                    })?;
                    if objects
                        .insert(
                            identifier,
                            FocusedObjectLocation {
                                archive_index,
                                object_index,
                            },
                        )
                        .is_some()
                    {
                        return Err(Error::InvalidFormat(format!(
                            "Numbers focused selector object {identifier} is repeated"
                        )));
                    }
                }
                Ok(())
            })?;
        }

        Ok(Self {
            package,
            archive_names,
            objects,
        })
    }

    fn with_object<T>(
        &self,
        identifier: u64,
        read: impl FnOnce(&ArchiveObject) -> Result<T>,
    ) -> Result<T> {
        let location = self.objects.get(&identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers focused selector object {identifier} is missing"
            ))
        })?;
        let archive_name = self
            .archive_names
            .get(location.archive_index)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers focused selector archive is missing".to_owned())
            })?;
        self.package.with_parsed_archive(archive_name, |archive| {
            let object = archive.objects.get(location.object_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers focused selector object {identifier} is missing"
                ))
            })?;
            if object.archive_info.identifier != Some(identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers focused selector object {identifier} location changed"
                )));
            }
            read(object)
        })
    }
}

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
}

#[derive(Debug, Clone, Copy)]
struct FocusedTableLocation {
    sheet_index: usize,
    table_index: usize,
    table_info_id: u64,
}

/// One validated native-table-to-selector index for a workbook snapshot.
///
/// The public editor remains selector-first. This private catalog exists only
/// to avoid repeating the sheet/drawable ownership traversal for every exact
/// table appearance read.
#[derive(Debug)]
pub(super) struct FocusedTableSelectorIndex {
    locations: HashMap<u64, FocusedTableLocation>,
}

impl FocusedTableSelectorIndex {
    pub(super) fn from_descriptors(
        editor: &NumbersEditor,
        descriptors: &[super::model::TableDescriptor],
    ) -> Result<Self> {
        let mut table_info_models = HashMap::new();
        table_info_models
            .try_reserve(descriptors.len())
            .map_err(|_| {
                allocation_error(
                    "Numbers focused selector table-info models",
                    descriptors.len(),
                )
            })?;
        for descriptor in descriptors {
            if table_info_models
                .insert(descriptor.table_info_id, descriptor.object_id)
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table-info object {} has multiple models",
                    descriptor.table_info_id
                )));
            }
        }
        if descriptors.is_empty() {
            return Ok(Self {
                locations: HashMap::new(),
            });
        }

        // `table_models` has already decoded and validated each rooted
        // TableInfo alias. Reuse that result here; this pass only needs the
        // sheet projection order, so it must not reinterpret arbitrary
        // drawable payloads as table ownership.
        let document = super::numbers_document(editor.package())?;
        let mut sheet_ids = HashSet::new();
        sheet_ids.try_reserve(document.sheets.len()).map_err(|_| {
            allocation_error(
                "Numbers focused selector sheet identifiers",
                document.sheets.len(),
            )
        })?;
        for sheet_reference in &document.sheets {
            sheet_ids.insert(sheet_reference.identifier);
        }
        let catalog = FocusedObjectCatalog::build(editor.package(), &sheet_ids)?;

        let mut table_locations = HashMap::new();
        table_locations
            .try_reserve(descriptors.len())
            .map_err(|_| {
                allocation_error("Numbers focused selector locations", descriptors.len())
            })?;
        let mut seen_projections = HashSet::new();
        seen_projections
            .try_reserve(descriptors.len())
            .map_err(|_| {
                allocation_error("Numbers focused selector projections", descriptors.len())
            })?;

        for (sheet_index, sheet_reference) in document.sheets.iter().enumerate() {
            let sheet_id = sheet_reference.identifier;
            if !catalog.objects.contains_key(&sheet_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} is missing"
                )));
            }
            let (_, sheet) = catalog.with_object(sheet_id, super::decode_sheet)?;
            let mut table_index = 0_usize;
            for drawable in &sheet.drawable_infos {
                let Some(&model_id) = table_info_models.get(&drawable.identifier) else {
                    continue;
                };
                let current_table_index = table_index;
                table_index = table_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers focused selector table index overflows".to_owned(),
                    )
                })?;
                seen_projections
                    .try_reserve(1)
                    .map_err(|_| allocation_error("Numbers focused selector projections", 1))?;
                if !seen_projections.insert((sheet_index, drawable.identifier)) {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers table model {model_id} has an ambiguous focused sheet projection"
                    )));
                }
                table_locations
                    .try_reserve(1)
                    .map_err(|_| allocation_error("Numbers focused selector locations", 1))?;
                if table_locations
                    .insert(
                        model_id,
                        FocusedTableLocation {
                            sheet_index,
                            table_index: current_table_index,
                            table_info_id: drawable.identifier,
                        },
                    )
                    .is_some()
                {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers table model {model_id} has multiple owning sheet drawables"
                    )));
                }
            }
        }

        for descriptor in descriptors {
            let Some(location) = table_locations.get(&descriptor.object_id) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table model {} has no owning sheet drawable",
                    descriptor.object_id
                )));
            };
            debug_assert_eq!(location.table_info_id, descriptor.table_info_id);
        }

        Ok(Self {
            locations: table_locations,
        })
    }

    pub(super) fn selectors(
        &self,
        native_id: u64,
    ) -> Result<(SheetSelector<'static>, TableSelector<'static>)> {
        let location = self.locations.get(&native_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table model {native_id} has no owning sheet drawable"
            ))
        })?;
        Ok((
            SheetSelector::index(location.sheet_index),
            TableSelector::index(location.table_index),
        ))
    }
}

/// Make a source catalog safe to pass through focused immutable constructors.
///
/// Native malformed packages can repeat a visible name.  The requested name
/// is checked against the original names before this list is used, while
/// duplicate entries that are irrelevant to an index lookup receive private
/// labels so they cannot make the focused catalog reject an otherwise valid
/// positional selection.  Matching remains exact and case-sensitive.
fn selector_safe_names(source_names: &[String], kind: &str) -> Vec<String> {
    let mut used = source_names.iter().cloned().collect::<HashSet<_>>();
    let mut seen = HashSet::new();
    source_names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            if seen.insert(name.as_str()) {
                return name.clone();
            }
            let mut replacement = format!("\0litchi-selector-{kind}-{index}");
            while used.contains(&replacement) {
                replacement.push('_');
            }
            used.insert(replacement.clone());
            replacement
        })
        .collect()
}

fn ensure_unique_sheet_name(source_names: &[String], selector: SheetSelector<'_>) -> Result<()> {
    let SheetSelector::Name(name) = selector else {
        return Ok(());
    };
    let matches = source_names
        .iter()
        .filter(|candidate| candidate.as_str() == name)
        .count();
    match matches {
        0 => Err(sheet_selector_error(selector)),
        1 => Ok(()),
        _ => Err(Error::ParseError(format!(
            "Numbers sheet name {name:?} is ambiguous"
        ))),
    }
}

fn ensure_unique_table_name(source_names: &[String], selector: TableSelector<'_>) -> Result<()> {
    let TableSelector::Name(name) = selector else {
        return Ok(());
    };
    let matches = source_names
        .iter()
        .filter(|candidate| candidate.as_str() == name)
        .count();
    match matches {
        0 => Err(table_selector_error(selector)),
        1 => Ok(()),
        _ => Err(Error::ParseError(format!(
            "Numbers table name {name:?} is ambiguous"
        ))),
    }
}

fn sheet_selector_error(selector: SheetSelector<'_>) -> Error {
    match selector {
        SheetSelector::Name(name) => {
            Error::ParseError(format!("Numbers sheet named {name:?} not found"))
        },
        SheetSelector::Index(index) => Error::ParseError(format!(
            "Numbers sheet catalog index {index} is out of bounds"
        )),
    }
}

fn table_selector_error(selector: TableSelector<'_>) -> Error {
    match selector {
        TableSelector::Name(name) => {
            Error::ParseError(format!("Numbers table named {name:?} not found"))
        },
        TableSelector::Index(index) => Error::ParseError(format!(
            "Numbers table catalog index {index} is out of bounds"
        )),
    }
}

/// Resolve a semantic sheet selector to its native object identifier.
pub(super) fn sheet_id(editor: &NumbersEditor, selector: SheetSelector<'_>) -> Result<u64> {
    SheetSelectorAdapter::from_editor(editor)?.native_id(selector)
}

/// Resolve a semantic table selector to its native model object identifier.
pub(super) fn table_id(editor: &NumbersEditor, selector: TableSelector<'_>) -> Result<u64> {
    TableSelectorAdapter::from_editor(editor)?.native_id(selector)
}

/// Return the semantic catalog position of a native table identifier for
/// adapter-internal follow-up operations.
pub(super) fn table_index(editor: &NumbersEditor, native_id: u64) -> Result<usize> {
    super::table_models(&editor.package)?
        .iter()
        .position(|table| table.object_id == native_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table {native_id} not found")))
}

/// Resolve one legacy native table identifier to selector-first sheet and
/// table positions without exposing that identifier to the focused owner.
pub(super) fn focused_table_location(
    editor: &NumbersEditor,
    native_id: u64,
) -> Result<(SheetSelector<'static>, TableSelector<'static>)> {
    let owner = super::find_table_owner(editor.package(), native_id)?;
    let sheets = editor.sheets()?;
    let mut sheet_matches = sheets
        .iter()
        .enumerate()
        .filter(|(_, sheet)| sheet.native_id() == owner.sheet_id);
    let (sheet_index, _) = sheet_matches.next().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model {native_id} belongs to an unreachable sheet"
        ))
    })?;
    if sheet_matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Numbers table model {native_id} belongs to an ambiguous sheet"
        )));
    }

    let table_info_ids = super::table_models(editor.package())?
        .into_iter()
        .map(|table| table.table_info_id)
        .collect::<HashSet<_>>();
    let (_, _, native_sheet) = super::numbers_sheet(editor.package(), owner.sheet_id)?;
    let mut table_matches = native_sheet
        .drawable_infos
        .iter()
        .filter(|drawable| table_info_ids.contains(&drawable.identifier))
        .enumerate()
        .filter(|(_, drawable)| drawable.identifier == owner.table_info_id);
    let (table_index, _) = table_matches.next().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model {native_id} is missing from its focused sheet projection"
        ))
    })?;
    if table_matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Numbers table model {native_id} has an ambiguous focused sheet projection"
        )));
    }

    Ok((
        SheetSelector::index(sheet_index),
        TableSelector::index(table_index),
    ))
}

/// Resolve a legacy table identifier to the selector positions understood by
/// the focused merge reader.
///
/// Merge metadata does not need the complete generated table model (or any
/// cell/tile projection). The host therefore derives the table position from
/// the rooted sheet drawable order and the bounded table-info ownership edge.
/// The returned positions follow the same table-info filtering used by the
/// focused Numbers reader, while all native identifiers remain private to this
/// adapter.
pub(super) fn focused_merge_table_indices(
    editor: &NumbersEditor,
    native_id: u64,
) -> Result<Option<(usize, usize)>> {
    let package = editor.package();
    let locations = super::object_locations(package)?;
    let Some(model_archive) = locations.get(&native_id) else {
        return Ok(None);
    };
    let canonical_model = package.with_parsed_archive(model_archive, |archive| {
        Ok(archive
            .object(native_id)
            .is_some_and(|object| object.messages.iter().any(|message| message.type_ == 6_001)))
    })?;
    // The host also admits historical model type aliases. Only canonical
    // models enter the focused rooted reader; the compatibility projection
    // validates a legacy candidate before returning its geometry.
    if !canonical_model {
        return Ok(None);
    }
    let document = super::numbers_document(package)?;
    let limits = litchi_numbers::PackageSemanticLimits::default();
    let check = |kind, observed: usize, maximum: usize| -> Result<()> {
        if observed > maximum {
            return Err(litchi_numbers::TableMergesError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            }
            .into());
        }
        Ok(())
    };
    check(
        litchi_numbers::TableMergesLimitKind::Sheets,
        document.sheets.len(),
        limits.max_sheets(),
    )?;
    let mut seen_sheets = HashSet::new();
    let mut references = 0usize;
    let mut tables = 0usize;
    let mut selected = None;
    for (sheet_position, sheet_reference) in document.sheets.iter().enumerate() {
        seen_sheets.try_reserve(1).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers merge sheet membership",
                amount: 1,
            })
        })?;
        if !seen_sheets.insert(sheet_reference.identifier) {
            return Err(Error::InvalidFormat(
                "Numbers document repeats a sheet reference".to_owned(),
            ));
        }
        let sheet_archive = locations.get(&sheet_reference.identifier).ok_or_else(|| {
            Error::InvalidFormat("Numbers merge sheet is missing from the object index".to_owned())
        })?;
        let sheet = package.with_parsed_archive(sheet_archive, |archive| {
            let object = archive.object(sheet_reference.identifier).ok_or_else(|| {
                Error::InvalidFormat("Numbers merge sheet is missing from its archive".to_owned())
            })?;
            super::model::decode_sheet(object).map(|(_, sheet)| sheet)
        })?;
        references = references.saturating_add(sheet.drawable_infos.len());
        check(
            litchi_numbers::TableMergesLimitKind::PayloadReferences,
            references,
            limits.max_references(),
        )?;
        let mut table_position = 0usize;
        for drawable in &sheet.drawable_infos {
            let Some(archive_name) = locations.get(&drawable.identifier) else {
                continue;
            };
            let model_id = package.with_parsed_archive(archive_name, |archive| {
                let object = archive.object(drawable.identifier).ok_or_else(|| {
                    Error::InvalidFormat("Numbers drawable missing from indexed archive".to_owned())
                })?;
                let mut model = None;
                for message in &object.messages {
                    if let Some(identifier) =
                        super::model::strict_attached_table_info_model_identifier(object, message)?
                        && model.replace(identifier).is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "Numbers drawable has multiple table-info payloads".to_owned(),
                        ));
                    }
                }
                Ok(model)
            })?;
            if let Some(identifier) = model_id {
                tables = tables.saturating_add(1);
                check(
                    litchi_numbers::TableMergesLimitKind::Tables,
                    tables,
                    limits.max_tables(),
                )?;
                if identifier == native_id
                    && selected.replace((sheet_position, table_position)).is_some()
                {
                    return Err(Error::InvalidFormat(
                        "Numbers table has multiple rooted owning drawables".to_owned(),
                    ));
                }
                table_position = table_position.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Numbers table position overflow".to_owned())
                })?;
            }
        }
    }
    // Detached table-info/model pairs remain explicit compatibility scope.
    Ok(selected)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

    #[test]
    fn selectors_share_the_editor_catalog_and_reject_invalid_entries() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let first_sheet = editor.sheets().unwrap().remove(0);
        let first_table = editor.tables().unwrap().remove(0);
        let duplicate = editor.duplicate_table(TableSelector::index(0)).unwrap();

        assert_eq!(
            sheet_id(&editor, SheetSelector::name("Summary")).unwrap(),
            first_sheet.native_id()
        );
        assert_eq!(
            sheet_id(&editor, SheetSelector::index(0)).unwrap(),
            first_sheet.native_id()
        );
        assert_eq!(
            table_id(&editor, TableSelector::name("Revenue")).unwrap(),
            first_table.native_id()
        );
        assert_eq!(
            table_id(&editor, TableSelector::index(0)).unwrap(),
            first_table.native_id()
        );
        assert_eq!(table_index(&editor, duplicate.native_id()).unwrap(), 1);

        assert!(sheet_id(&editor, SheetSelector::name("Missing")).is_err());
        assert!(sheet_id(&editor, SheetSelector::index(1)).is_err());
        assert!(table_id(&editor, TableSelector::name("Missing")).is_err());
        assert!(table_id(&editor, TableSelector::index(2)).is_err());
    }

    #[test]
    fn table_name_resolution_reports_cross_sheet_ambiguity() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        editor.add_empty_sheet("Archive").unwrap();
        editor
            .add_empty_table(SheetSelector::name("Archive"), "Revenue", 2, 2)
            .unwrap();

        assert!(table_id(&editor, TableSelector::name("Revenue")).is_err());
    }

    #[test]
    fn selectors_keep_exact_case_and_ambiguous_operations_atomic() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();

        assert!(sheet_id(&editor, SheetSelector::name("summary")).is_err());
        assert!(table_id(&editor, TableSelector::name("revenue")).is_err());

        editor.add_empty_sheet("Archive").unwrap();
        editor
            .add_empty_table(SheetSelector::name("Archive"), "Revenue", 2, 2)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .duplicate_table(TableSelector::name("Revenue"))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn focused_table_selector_index_preserves_native_table_order() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("First")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        editor
            .add_empty_table(SheetSelector::index(0), "Second", 3, 1)
            .unwrap();
        let exact = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let descriptors = super::super::table_models(exact.package()).unwrap();
        let index = FocusedTableSelectorIndex::from_descriptors(&exact, &descriptors).unwrap();

        assert_eq!(
            descriptors
                .iter()
                .map(|descriptor| descriptor.model.table_name.as_str())
                .collect::<Vec<_>>(),
            ["First", "Second"]
        );
        assert_eq!(
            index.selectors(descriptors[0].object_id).unwrap(),
            (SheetSelector::index(0), TableSelector::index(0))
        );
        assert_eq!(
            index.selectors(descriptors[1].object_id).unwrap(),
            (SheetSelector::index(0), TableSelector::index(1))
        );
    }

    #[test]
    fn focused_table_selector_index_rejects_orphaned_descriptors() {
        let editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("First")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let exact = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let mut descriptors = super::super::table_models(exact.package()).unwrap();
        let descriptor = descriptors.pop().unwrap();
        descriptors.push(super::super::model::TableDescriptor {
            table_info_id: u64::MAX,
            ..descriptor
        });

        assert!(FocusedTableSelectorIndex::from_descriptors(&exact, &descriptors).is_err());
    }

    #[test]
    fn focused_table_selector_index_rejects_cross_sheet_ownership() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("First")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        editor.add_empty_sheet("Archive").unwrap();
        editor
            .add_empty_table(SheetSelector::name("Archive"), "Second", 2, 2)
            .unwrap();
        let exact = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let mut descriptors = super::super::table_models(exact.package()).unwrap();
        assert_eq!(descriptors.len(), 2);
        let first_model = descriptors[0].object_id;
        descriptors[1].object_id = first_model;

        let error = FocusedTableSelectorIndex::from_descriptors(&exact, &descriptors)
            .expect_err("one model cannot own drawables on two sheets");
        assert!(
            error
                .to_string()
                .contains("multiple owning sheet drawables")
        );
    }
}
