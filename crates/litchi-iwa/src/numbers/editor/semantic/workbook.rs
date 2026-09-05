//! Workbook and table lifecycle semantics.

#![allow(unused_imports)]

use super::super::selectors;
use super::*;
use litchi_numbers::{Package, SheetSelector, TableSelector};

const TABLE_APPEARANCE_STYLE_MESSAGE_TYPES: &[u32] = &[6_003];
const TABLE_APPEARANCE_PRESET_MESSAGE_TYPES: &[u32] = &[6_008];
const TABLE_APPEARANCE_NETWORK_MESSAGE_TYPES: &[u32] = &[6_247];

#[derive(Debug, Clone, Copy)]
struct SourceBuiltAppearanceObjectLocation {
    archive_index: usize,
    object_index: usize,
}

/// Payload-free source-built appearance object index.
///
/// The package keeps parsed archives in its bounded cache. This catalog owns
/// only archive names and object slots, so appearance reads can borrow one
/// selected message at a time without cloning every archive in the package.
struct SourceBuiltAppearanceCatalog<'package> {
    package: &'package IWorkPackage,
    archive_names: Vec<String>,
    objects: HashMap<u64, SourceBuiltAppearanceObjectLocation>,
}

impl<'package> SourceBuiltAppearanceCatalog<'package> {
    fn build(package: &'package IWorkPackage) -> Result<Self> {
        let archive_count = package.iwa_entry_names().count();
        let mut archive_names = Vec::new();
        archive_names
            .try_reserve_exact(archive_count)
            .map_err(|_| {
                Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers table appearance archive names",
                    amount: archive_count,
                })
            })?;
        for name in package.iwa_entry_names() {
            let mut owned = String::new();
            owned.try_reserve_exact(name.len()).map_err(|_| {
                Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers table appearance archive name",
                    amount: name.len(),
                })
            })?;
            owned.push_str(name);
            archive_names.push(owned);
        }

        let mut objects = HashMap::new();
        objects.try_reserve(archive_names.len()).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table appearance object index",
                amount: archive_names.len(),
            })
        })?;
        for (archive_index, archive_name) in archive_names.iter().enumerate() {
            package.with_parsed_archive(archive_name, |archive| {
                objects.try_reserve(archive.objects.len()).map_err(|_| {
                    Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                        resource: "Numbers table appearance object index",
                        amount: archive.objects.len(),
                    })
                })?;
                for (object_index, object) in archive.objects.iter().enumerate() {
                    let identifier = object.archive_info.identifier.ok_or_else(|| {
                        Error::InvalidFormat(format!("Object in {archive_name} has no identifier"))
                    })?;
                    if objects
                        .insert(
                            identifier,
                            SourceBuiltAppearanceObjectLocation {
                                archive_index,
                                object_index,
                            },
                        )
                        .is_some()
                    {
                        return Err(Error::InvalidFormat(format!(
                            "Numbers appearance object {identifier} is repeated"
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

    fn with_payload(
        &self,
        identifier: u64,
        message_types: &[u32],
        context: &str,
        read: &mut dyn FnMut(&[u8]),
    ) -> Result<()> {
        let location = self.objects.get(&identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers {context} object {identifier} is missing"))
        })?;
        let archive_name = self
            .archive_names
            .get(location.archive_index)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table appearance archive is missing".to_owned())
            })?;
        self.package.with_parsed_archive(archive_name, |archive| {
            let object = archive.objects.get(location.object_index).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers {context} object {identifier} is missing"))
            })?;
            if object.archive_info.identifier != Some(identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {context} object {identifier} location changed"
                )));
            }
            let mut messages = object
                .messages
                .iter()
                .filter(|message| message_types.contains(&message.type_));
            let message = messages.next().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers {context} object {identifier} has no supported payload"
                ))
            })?;
            if messages.next().is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {context} object {identifier} has multiple supported payloads"
                )));
            }
            read(message.data.as_slice());
            Ok(())
        })
    }
}

impl NumbersEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(IWorkPackage::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        numbers_document(&package)?;
        Ok(Self { package })
    }

    /// Reopen a staged package while retaining legacy source-built
    /// provenance across the internal byte-validation boundary.
    ///
    /// A public `from_bytes` call always denotes an exact source. This helper
    /// is reserved for transactions that started from an `IWorkPackage`
    /// without an exact source and must not turn that compatibility graph into
    /// an exact focused-owner candidate merely because validation serialized
    /// it once.
    pub(crate) fn from_validation_bytes(bytes: &[u8], source_built: bool) -> Result<Self> {
        let mut editor = Self::from_bytes(bytes)?;
        if source_built {
            editor.package.discard_exact_source_for_compatibility();
        }
        Ok(editor)
    }

    pub fn sheets(&self) -> Result<Vec<NumbersSheetInfo>> {
        let document = numbers_document(&self.package)?;
        let locations = object_locations(&self.package)?;
        document
            .sheets
            .into_iter()
            .enumerate()
            .map(|(index, reference)| {
                let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers sheet object {} is missing",
                        reference.identifier
                    ))
                })?;
                let archive = self.package.archive(archive_name)?;
                let object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers sheet object {} is missing",
                        reference.identifier
                    ))
                })?;
                let (_, sheet) = decode_sheet(object)?;
                Ok(NumbersSheetInfo {
                    object_id: reference.identifier,
                    index,
                    name: sheet.name,
                })
            })
            .collect()
    }

    pub fn tables(&self) -> Result<Vec<NumbersTableInfo>> {
        let source_built = !self.package.source_is_exact();
        let descriptors = table_models(&self.package)?;
        // Source-built and older compatibility packages may not carry the
        // exact-source metadata required by focused transactions. Their
        // read-only appearance projection still belongs to the focused
        // Numbers owner through its bounded compatibility seam. Once an exact
        // source is admitted to the focused owner, every ingress/read failure
        // is terminal.
        let source_built_catalog = if source_built {
            Some(SourceBuiltAppearanceCatalog::build(&self.package)?)
        } else {
            None
        };
        let focused_package = if source_built {
            None
        } else {
            let source = self.package.exact_source_owner().ok_or_else(|| {
                Error::InvalidFormat(
                    "focused Numbers table-appearance source is not exact".to_owned(),
                )
            })?;
            Some(
                Package::__from_shared_source_with_options(
                    source,
                    litchi_numbers::PackageReadOptions::default(),
                )
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "focused Numbers table-appearance source failed: {error}"
                    ))
                })?,
            )
        };
        let focused_locations = if source_built {
            None
        } else {
            Some(selectors::FocusedTableSelectorIndex::from_descriptors(
                self,
                &descriptors,
            )?)
        };
        let read_focused_appearance = litchi_numbers::Package::table_appearance;
        let mut tables = descriptors
            .into_iter()
            .map(|descriptor| {
                let appearance = if let Some(package) = focused_package.as_ref() {
                    let locations = focused_locations.as_ref().ok_or_else(|| {
                        Error::InvalidFormat("focused Numbers selector index is missing".to_owned())
                    })?;
                    let (sheet_selector, table_selector) =
                        locations.selectors(descriptor.object_id)?;
                    read_focused_appearance(package, sheet_selector, table_selector).map_err(
                        |error| {
                            Error::InvalidFormat(format!(
                                "focused Numbers table-appearance read failed: {error}"
                            ))
                        },
                    )?
                } else {
                    let catalog = source_built_catalog.as_ref().ok_or_else(|| {
                        Error::InvalidFormat(
                            "source-built Numbers appearance catalog is missing".to_owned(),
                        )
                    })?;
                    source_built_table_appearance(catalog, &descriptor)?
                };
                Ok(NumbersTableInfo {
                    object_id: descriptor.object_id,
                    index: 0,
                    name: descriptor.model.table_name,
                    rows: descriptor.model.number_of_rows as usize,
                    columns: descriptor.model.number_of_columns as usize,
                    appearance,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        for (index, table) in tables.iter_mut().enumerate() {
            table.index = index;
        }
        Ok(tables)
    }

    /// Set a cell to a formula expression.
    ///
    /// The expression is compiled to Numbers' native postfix AST and interned
    /// in the table's formula list. Local and cross-table cells, rectangles,
    /// and whole-row/column references are mirrored into CalculationEngine
    /// dependency records in lockstep with the formula table. Unsupported volatile, lazy,
    /// remote-data, and spill expressions fail before the package is changed.
    pub fn set_formula(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        expression: FormulaExpression,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        table_formula::set_attached_table_formula(
            &mut staged,
            table_id,
            row,
            column,
            expression,
            None,
        )?;
        self.package = staged;
        Ok(())
    }

    /// Set a formula together with the value displayed before the next recalculation.
    pub fn set_formula_with_cached_value(
        &mut self,
        table_id: u64,
        row: usize,
        column: usize,
        expression: FormulaExpression,
        cached_value: FormulaCachedValue,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        table_formula::set_attached_table_formula(
            &mut staged,
            table_id,
            row,
            column,
            expression,
            Some(cached_value),
        )?;
        self.package = staged;
        Ok(())
    }

    /// Resize a table while preserving existing cells and stable row/column UIDs.
    ///
    /// Growth creates blank trailing rows or columns. Shrinkage is accepted only
    /// when the removed trailing region contains no stored cells; this prevents
    /// silently orphaning strings, formulas, rich text, comments, or styles.
    pub fn resize_table(
        &mut self,
        selector: TableSelector<'_>,
        rows: usize,
        columns: usize,
    ) -> Result<()> {
        let table_id = selectors::table_id(self, selector)?;
        let source_built = !self.package.source_is_exact();
        let mut staged = self.package.clone();
        resize_attached_table_in_package(&mut staged, table_id, rows, columns)?;

        let verified = NumbersEditor::from_validation_bytes(&staged.to_bytes()?, source_built)?;
        let resized = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| Error::InvalidFormat("Numbers resized table disappeared".to_owned()))?;
        if (resized.rows, resized.columns) != (rows, columns) {
            return Err(Error::InvalidFormat(
                "Numbers table resize failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Unlink and remove a table model from its owning sheet.
    ///
    /// Private storage, formula dependency owners, UUID registrations, and
    /// now-empty component members are removed. Shared storage and styles are
    /// retained. Deletion is rejected while another table has a formula edge
    /// targeting this table.
    pub fn remove_table(&mut self, selector: TableSelector<'_>) -> Result<NumbersTableInfo> {
        let table_id = selectors::table_id(self, selector)?;
        let table = self
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::ParseError(format!("Numbers table object {table_id} not found"))
            })?;
        let owner = find_table_owner(&self.package, table_id)?;
        let locations = object_locations(&self.package)?;
        let descriptors = table_models(&self.package)?;
        let descriptor = descriptors
            .iter()
            .find(|descriptor| descriptor.object_id == table_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table model {table_id} is missing"))
            })?;
        let owned_graph = table_owned_graph(&self.package, &locations, &descriptor.model)?;
        let mut shared_owned_ids = HashSet::new();
        for other in descriptors
            .iter()
            .filter(|candidate| candidate.object_id != table_id)
        {
            shared_owned_ids
                .extend(table_owned_graph(&self.package, &locations, &other.model)?.into_keys());
        }
        let private_owned_ids = owned_graph
            .into_keys()
            .filter(|identifier| !shared_owned_ids.contains(identifier))
            .collect::<Vec<_>>();
        let sheet_archive = locations.get(&owner.sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
        })?;
        let mut staged = self.package.clone();
        let mut removed_identifiers = remove_table_formula_graph(&mut staged, owner.table_info_id)?;
        staged.update_archive(sheet_archive, |archive| {
            let object = archive.object_mut(owner.sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let previous = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let current = previous
                .iter()
                .copied()
                .filter(|identifier| *identifier != owner.table_info_id)
                .collect::<Vec<_>>();
            if current.len() + 1 != previous.len() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {} does not reference table info {} exactly once",
                    owner.sheet_id, owner.table_info_id
                )));
            }
            replace_sheet_drawable_references(object, message_index, &previous, &current)?;
            object.archive_info.message_infos[message_index]
                .object_references
                .retain(|&identifier| identifier != owner.table_info_id);
            for field in &mut object.archive_info.message_infos[message_index].field_infos {
                field
                    .object_references
                    .retain(|&identifier| identifier != owner.table_info_id);
            }
            Ok(())
        })?;
        let info_archive = locations.get(&owner.table_info_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table info {} is missing",
                owner.table_info_id
            ))
        })?;
        let model_archive = locations.get(&table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model {table_id} is missing"))
        })?;
        let private_owned_locations = private_owned_ids
            .iter()
            .map(|identifier| {
                locations
                    .get(identifier)
                    .map(|entry| (entry.as_str(), *identifier))
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers table storage object {identifier} is missing"
                        ))
                    })
            })
            .collect::<Result<Vec<_>>>()?;
        let mut affected_components = HashMap::<String, u64>::new();
        for (entry, identifier) in std::iter::once((info_archive.as_str(), owner.table_info_id))
            .chain(std::iter::once((model_archive.as_str(), table_id)))
            .chain(private_owned_locations)
        {
            let Some(component) = component_identifier_for_entry(&staged, entry)? else {
                continue;
            };
            affected_components.insert(entry.to_owned(), component);
            remove_component_external_references_to_object(&mut staged, component, identifier)?;
            if component_uuid_identifiers(&staged, component)?
                .is_some_and(|identifiers| identifiers.contains(&identifier))
            {
                remove_component_object_uuids(&mut staged, component, &[identifier])?;
            }
        }
        let dedicated_component = format!("Index/Tables/Table-{}.iwa", owner.table_info_id);
        if info_archive == model_archive && info_archive == &dedicated_component {
            staged.remove_entry(info_archive).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table component {info_archive} is missing"))
            })?;
        } else {
            remove_object_or_empty_entry(&mut staged, &locations, owner.table_info_id)?;
            remove_object_or_empty_entry(&mut staged, &locations, table_id)?;
        }
        for identifier in &private_owned_ids {
            remove_object_or_empty_entry(&mut staged, &locations, *identifier)?;
        }
        for (entry, component) in affected_components {
            if !staged.contains_entry(&entry) {
                remove_component_registration(&mut staged, component)?;
            }
        }
        removed_identifiers.extend([owner.table_info_id, table_id]);
        removed_identifiers.extend(private_owned_ids);
        release_package_identifier_suffix(&mut staged, &removed_identifiers)?;

        let source_built = !self.package.source_is_exact();
        let verified = NumbersEditor::from_validation_bytes(&staged.to_bytes()?, source_built)?;
        if table_models(verified.package())?
            .iter()
            .any(|candidate| candidate.object_id == table_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers table deletion failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(table)
    }

    /// Append an empty sheet to the workbook and return its allocated object ID.
    pub fn add_empty_sheet(&mut self, name: &str) -> Result<NumbersSheetInfo> {
        validate_name(name, "sheet")?;
        let locations = object_locations(&self.package)?;
        let identifier = locations
            .keys()
            .copied()
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let mut staged = self.package.clone();
        staged.update_archive("Index/Document.iwa", |archive| {
            archive.insert_object(crate::archive::ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 2,
                    data: tn::SheetArchive {
                        name: name.to_owned(),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })?;
        update_numbers_document(&mut staged, |document| {
            document.sheets.push(crate::protobuf::tsp::Reference {
                identifier,
                ..Default::default()
            });
            Ok(())
        })?;
        let source_built = !self.package.source_is_exact();
        let verified = NumbersEditor::from_validation_bytes(&staged.to_bytes()?, source_built)?;
        let created = verified
            .sheets()?
            .into_iter()
            .find(|sheet| sheet.object_id == identifier)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers sheet creation failed validation".to_owned())
            })?;
        self.package = staged;
        Ok(created)
    }

    /// Add an independent empty native table to an existing sheet.
    ///
    /// An attached table supplies structural templates when one exists. If the
    /// workbook is table-less, the first table is built from its native theme
    /// preset instead. Cell stores, data lists, row/column UIDs, headers, stroke
    /// state, and the CalculationEngine owner are allocated independently;
    /// workbook styles are shared intentionally.
    #[allow(deprecated)]
    pub fn add_empty_table(
        &mut self,
        selector: SheetSelector<'_>,
        name: &str,
        rows: usize,
        columns: usize,
    ) -> Result<NumbersTableInfo> {
        let sheet_id = selectors::sheet_id(self, selector)?;
        let sheets = self.sheets()?;
        if !sheets.iter().any(|sheet| sheet.object_id == sheet_id) {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not in the workbook"
            )));
        }

        let descriptors = table_models(&self.package)?;
        let mut staged = self.package.clone();
        let graph = if let Some(template) = descriptors.first() {
            let template_owner = find_table_owner(&self.package, template.object_id)?;
            table_create::create_empty_table_graph(
                &mut staged,
                template_owner.table_info_id,
                template.object_id,
                template_owner.sheet_id,
                sheet_id,
                name,
                rows,
                columns,
                (template_owner.sheet_id == sheet_id).then_some(EMPTY_TABLE_POSITION_OFFSET),
            )?
        } else {
            table_bootstrap::bootstrap_empty_table_graph(
                &mut staged,
                sheet_id,
                name,
                rows,
                columns,
            )?
        };
        let new_info_id = graph.info_object_id;
        let new_model_id = graph.model_object_id;
        let locations = object_locations(&staged)?;
        let sheet_archive_name = locations
            .get(&sheet_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
        staged.update_archive(sheet_archive_name, |archive| {
            let object = archive.object_mut(sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing"))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let existing_drawables = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let existing_drawable_set = existing_drawables.iter().copied().collect::<HashSet<_>>();
            let mut current_drawables = existing_drawables.clone();
            current_drawables.push(new_info_id);
            replace_sheet_drawable_references(
                object,
                message_index,
                &existing_drawables,
                &current_drawables,
            )?;
            let references =
                &mut object.archive_info.message_infos[message_index].object_references;
            if !references.contains(&new_info_id) {
                references.push(new_info_id);
            }
            for field in &mut object.archive_info.message_infos[message_index].field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|identifier| existing_drawable_set.contains(identifier))
                    && !field.object_references.contains(&new_info_id)
                {
                    field.object_references.push(new_info_id);
                }
            }
            Ok(())
        })?;

        let source_built = !self.package.source_is_exact();
        let verified = NumbersEditor::from_validation_bytes(&staged.to_bytes()?, source_built)?;
        let created = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == new_model_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table creation failed validation".to_owned())
            })?;
        if (created.rows, created.columns, created.name.as_str()) != (rows, columns, name) {
            return Err(Error::InvalidFormat(
                "Numbers table creation produced unexpected properties".to_owned(),
            ));
        }
        self.package = staged;
        Ok(created)
    }

    /// Duplicate a populated table on its owning sheet.
    ///
    /// Cell tiles, headers, data lists, UID maps, stroke state, formulas, and
    /// CalculationEngine dependency owners are cloned independently. Workbook
    /// styles and referenced rich-text/comment payloads retain their native
    /// copy-on-write sharing.
    #[allow(deprecated)]
    pub fn duplicate_table(&mut self, selector: TableSelector<'_>) -> Result<NumbersTableInfo> {
        let source_built = !self.package.source_is_exact();
        let table_id = selectors::table_id(self, selector)?;
        let descriptors = table_models(&self.package)?;
        let source = descriptors
            .iter()
            .find(|descriptor| descriptor.object_id == table_id)
            .ok_or_else(|| Error::ParseError(format!("Numbers table {table_id} not found")))?;
        let owner = find_table_owner(&self.package, table_id)?;
        let existing_names = descriptors
            .iter()
            .filter_map(|descriptor| {
                find_table_owner(&self.package, descriptor.object_id)
                    .ok()
                    .filter(|candidate| candidate.sheet_id == owner.sheet_id)
                    .map(|_| descriptor.model.table_name.as_str())
            })
            .collect::<HashSet<_>>();
        let name = duplicate_table_name(&source.model.table_name, &existing_names)?;
        let source_package = &self.package;
        let locations = object_locations(source_package)?;
        let sheet_archive_name = locations.get(&owner.sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
        })?;
        let mut staged = source_package.clone();
        let cloned = duplicate_attached_table_graph_in_package(
            source_package,
            &mut staged,
            owner.table_info_id,
            table_id,
            &name,
            TABLE_DUPLICATE_OFFSET,
        )?;
        staged.update_archive(sheet_archive_name, |archive| {
            let object = archive.object_mut(owner.sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let previous = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let mut current = previous.clone();
            current.push(cloned.info_object_id);
            replace_sheet_drawable_references(object, message_index, &previous, &current)?;
            let info = &mut object.archive_info.message_infos[message_index];
            info.object_references.push(cloned.info_object_id);
            for field in &mut info.field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|identifier| previous.contains(identifier))
                {
                    field.object_references.push(cloned.info_object_id);
                }
            }
            Ok(())
        })?;

        let verified = NumbersEditor::from_validation_bytes(&staged.to_bytes()?, source_built)?;
        let created = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == cloned.model_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table duplication failed validation".to_owned())
            })?;
        if (created.name.as_str(), created.rows, created.columns)
            != (
                name.as_str(),
                source.model.number_of_rows as usize,
                source.model.number_of_columns as usize,
            )
        {
            return Err(Error::InvalidFormat(
                "Numbers table duplicate has unexpected properties".to_owned(),
            ));
        }
        self.package = staged;
        Ok(created)
    }

    pub fn package(&self) -> &IWorkPackage {
        &self.package
    }

    pub(super) fn sheet_owned_drawable_ids(&self, sheet_id: u64) -> Result<HashSet<u64>> {
        if !self
            .sheets()?
            .iter()
            .any(|sheet| sheet.object_id == sheet_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not reachable"
            )));
        }
        Ok(numbers_sheet_drawable_owners(&self.package)?
            .into_iter()
            .filter_map(|(drawable_id, owner_id)| (owner_id == sheet_id).then_some(drawable_id))
            .collect())
    }

    pub(super) fn require_sheet_drawable(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<()> {
        if !self
            .sheet_owned_drawable_ids(sheet_id)?
            .contains(&drawable_object_id)
        {
            return Err(Error::ParseError(format!(
                "drawable object {drawable_object_id} is not owned by Numbers sheet {sheet_id}"
            )));
        }
        if !self
            .sheet_drawables(sheet_id)?
            .iter()
            .any(|drawable| drawable.id.get() == drawable_object_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers sheet drawable {drawable_object_id} has no supported direct drawable payload"
            )));
        }
        Ok(())
    }
}

fn source_built_table_appearance(
    catalog: &SourceBuiltAppearanceCatalog<'_>,
    descriptor: &TableDescriptor,
) -> Result<litchi_numbers::Appearance> {
    let style_identifier = descriptor.model.table_style.identifier;
    let style_preset_identifier = descriptor
        .model
        .table_style_preset
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0);
    litchi_numbers::Package::__table_appearance_from_source_built(
        style_identifier,
        style_preset_identifier,
        |identifier, kind, consume| {
            let (message_types, context) = match kind {
                litchi_numbers::SourceBuiltAppearancePayload::TableStyle => {
                    (TABLE_APPEARANCE_STYLE_MESSAGE_TYPES, "table style")
                },
                litchi_numbers::SourceBuiltAppearancePayload::TableStylePreset => {
                    (TABLE_APPEARANCE_PRESET_MESSAGE_TYPES, "table style preset")
                },
                litchi_numbers::SourceBuiltAppearancePayload::TableStyleNetwork => (
                    TABLE_APPEARANCE_NETWORK_MESSAGE_TYPES,
                    "table style network",
                ),
            };
            catalog.with_payload(identifier, message_types, context, consume)
        },
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Numbers source-built table-appearance read failed: {error}"
        ))
    })
}
