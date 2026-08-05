//! Sheet, table, cell, formula, and comment model operations.

use super::*;
use crate::detect::detect_application_from_document;
use crate::package_metadata::PACKAGE_METADATA_ENTRY;
use crate::registry::Application;

const DEFAULT_TILE_SIZE_ROWS: u32 = 256;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
pub(super) fn numbers_document(package: &IWorkPackage) -> Result<tn::DocumentArchive> {
    package.with_parsed_archive("Index/Document.iwa", |archive| {
        let object = archive
            .object(1)
            .ok_or_else(|| Error::InvalidFormat("Numbers root object 1 is missing".to_owned()))?;
        object
            .messages
            .iter()
            .find(|message| {
                detect_application_from_document(&message.data) == Some(Application::Numbers)
            })
            .and_then(|message| tn::DocumentArchive::decode(message.data.as_slice()).ok())
            .ok_or_else(|| {
                Error::InvalidFormat("package does not contain a Numbers root document".to_owned())
            })
    })
}

const NUMBERS_CATALOG_DEFAULT_MAX_ARCHIVES: usize = 1_024;
const NUMBERS_CATALOG_DEFAULT_MAX_ARCHIVE_READS: usize = 2_048;
const NUMBERS_CATALOG_DEFAULT_MAX_OBJECTS: usize = 1_000_000;
const NUMBERS_CATALOG_DEFAULT_MAX_SHEETS: usize = 65_536;
const NUMBERS_CATALOG_DEFAULT_MAX_DRAWABLES: usize = 1_000_000;
const NUMBERS_CATALOG_DEFAULT_MAX_REFERENCE_EDGES: usize = 2_000_000;
const NUMBERS_CATALOG_DEFAULT_MAX_SEMANTIC_DECODES: usize = 2_000_000;

/// Bounded work limits for one operation-scoped Numbers object catalog.
///
/// These limits deliberately sit above normal iWork document sizes while
/// keeping a malformed package from forcing an unbounded index or semantic
/// validation pass. They are operation limits, not package-ingress limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct NumbersObjectCatalogLimits {
    pub(super) max_archives: usize,
    pub(super) max_archive_reads: usize,
    pub(super) max_objects: usize,
    pub(super) max_sheets: usize,
    pub(super) max_drawables: usize,
    pub(super) max_reference_edges: usize,
    pub(super) max_semantic_decodes: usize,
}

impl Default for NumbersObjectCatalogLimits {
    fn default() -> Self {
        Self {
            max_archives: NUMBERS_CATALOG_DEFAULT_MAX_ARCHIVES,
            max_archive_reads: NUMBERS_CATALOG_DEFAULT_MAX_ARCHIVE_READS,
            max_objects: NUMBERS_CATALOG_DEFAULT_MAX_OBJECTS,
            max_sheets: NUMBERS_CATALOG_DEFAULT_MAX_SHEETS,
            max_drawables: NUMBERS_CATALOG_DEFAULT_MAX_DRAWABLES,
            max_reference_edges: NUMBERS_CATALOG_DEFAULT_MAX_REFERENCE_EDGES,
            max_semantic_decodes: NUMBERS_CATALOG_DEFAULT_MAX_SEMANTIC_DECODES,
        }
    }
}

impl NumbersObjectCatalogLimits {
    fn validate(self) -> Result<()> {
        if self.max_archives == 0
            || self.max_archive_reads == 0
            || self.max_objects == 0
            || self.max_sheets == 0
            || self.max_drawables == 0
            || self.max_reference_edges == 0
            || self.max_semantic_decodes == 0
        {
            return Err(Error::InvalidFormat(
                "Numbers object catalog limits must be non-zero".to_owned(),
            ));
        }
        Ok(())
    }
}

/// Bounded measurements for one Numbers object-catalog operation.
///
/// `archive_reads` counts logical catalog archive visits. It is intentionally
/// not called decompressions: the package cache may satisfy a visit without
/// parsing, and the package-level parser has no public observer hook.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct NumbersObjectCatalogStats {
    pub(super) archive_reads: usize,
    pub(super) archives_scanned: usize,
    pub(super) objects_indexed: usize,
    pub(super) sheet_objects_scanned: usize,
    pub(super) drawable_objects_scanned: usize,
    pub(super) reference_edges: usize,
    pub(super) semantic_decodes: usize,
    pub(super) peak_live_archives: usize,
    /// The catalog retains slots and summaries only; it never retains payload
    /// bytes or an expanded `Archive`.
    pub(super) retained_payload_bytes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct NumbersObjectSlot {
    archive_index: u32,
    object_index: u32,
}

#[derive(Debug, Clone, Copy)]
struct NumbersObjectRecord {
    slot: NumbersObjectSlot,
    shape_message_count: u32,
    caption_message_count: u32,
    storage_message_count: u32,
}

#[derive(Debug, Clone, Copy)]
struct NumbersShapeSummary {
    is_text_box: Option<bool>,
    parent_id: Option<u64>,
    owned_storage_id: Option<u64>,
    deprecated_storage_id: Option<u64>,
    title_id: Option<u64>,
    caption_id: Option<u64>,
}

#[derive(Debug, Clone, Copy)]
struct NumbersDrawableSummary {
    owner_sheet_id: u64,
    slot: NumbersObjectSlot,
    shape: NumbersShapeSummary,
}

#[derive(Debug, Clone)]
struct NumbersSheetSummary {
    slot: NumbersObjectSlot,
    drawable_ids: Box<[u64]>,
}

/// Compact, operation-scoped index for reachable Numbers object graphs.
///
/// The catalog owns no parsed archive, protobuf payload, or raw message bytes.
/// It retains only archive names, object slots, and the small semantic facts
/// needed to validate ordinary text-box ownership. A mutation revision binds
/// the catalog to the immutable package generation that was scanned; callers
/// must discard it after copy-on-write mutation.
#[derive(Debug)]
pub(super) struct NumbersObjectCatalog {
    package_revision: u64,
    limits: NumbersObjectCatalogLimits,
    archive_names: Vec<Box<str>>,
    objects: HashMap<u64, NumbersObjectRecord>,
    sheets: HashMap<u64, NumbersSheetSummary>,
    drawable_owners: HashMap<u64, u64>,
    drawable_owner_counts: HashMap<u64, u32>,
    duplicate_drawables: HashSet<u64>,
    drawables: HashMap<u64, NumbersDrawableSummary>,
    storage_owner_counts: HashMap<u64, u32>,
    uuid_identifiers: Option<HashSet<u64>>,
    stats: NumbersObjectCatalogStats,
}

impl NumbersObjectCatalog {
    pub(super) fn build(package: &IWorkPackage) -> Result<Self> {
        Self::build_with_limits(package, NumbersObjectCatalogLimits::default())
    }

    #[allow(deprecated)]
    pub(super) fn build_with_limits(
        package: &IWorkPackage,
        limits: NumbersObjectCatalogLimits,
    ) -> Result<Self> {
        limits.validate()?;
        let document = numbers_document(package)?;
        if document.sheets.len() > limits.max_sheets {
            return Err(Error::InvalidFormat(format!(
                "Numbers object catalog sheet count {} exceeds the {} sheet budget at Numbers document",
                document.sheets.len(),
                limits.max_sheets
            )));
        }

        let mut requested_sheet_ids = Vec::with_capacity(document.sheets.len());
        let mut requested_sheets = HashSet::with_capacity(document.sheets.len());
        for reference in document.sheets {
            if !requested_sheets.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers document repeats sheet object {}",
                    reference.identifier
                )));
            }
            requested_sheet_ids.push(reference.identifier);
        }

        let mut catalog = Self {
            package_revision: package.mutation_revision(),
            limits,
            archive_names: Vec::new(),
            objects: HashMap::new(),
            sheets: HashMap::with_capacity(requested_sheet_ids.len()),
            drawable_owners: HashMap::new(),
            drawable_owner_counts: HashMap::new(),
            duplicate_drawables: HashSet::new(),
            drawables: HashMap::new(),
            storage_owner_counts: HashMap::new(),
            uuid_identifiers: None,
            stats: NumbersObjectCatalogStats::default(),
        };

        for archive_name in package.iwa_entry_names() {
            if catalog.archive_names.len() >= limits.max_archives {
                return Err(Error::InvalidFormat(format!(
                    "Numbers object catalog archive count {} exceeds the {} archive budget at {archive_name}",
                    catalog.archive_names.len().saturating_add(1),
                    limits.max_archives
                )));
            }
            let archive_index = u32::try_from(catalog.archive_names.len()).map_err(|_| {
                Error::InvalidFormat("Numbers object catalog archive index exceeds u32".to_owned())
            })?;
            catalog
                .archive_names
                .push(archive_name.to_owned().into_boxed_str());
            catalog.with_archive(package, archive_name, |catalog, archive| {
                catalog.stats.archives_scanned = catalog
                    .stats
                    .archives_scanned
                    .checked_add(1)
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "Numbers object catalog archive counter overflow".to_owned(),
                        )
                    })?;
                for (object_index, object) in archive.objects.iter().enumerate() {
                    if catalog.objects.len() >= limits.max_objects {
                        return Err(Error::InvalidFormat(format!(
                            "Numbers object catalog object count {} exceeds the {} object budget at {archive_name}",
                            catalog.objects.len().saturating_add(1),
                            limits.max_objects
                        )));
                    }
                    let identifier = object.archive_info.identifier.ok_or_else(|| {
                        Error::Archive(format!(
                            "Object in {archive_name} has no identifier"
                        ))
                    })?;
                    if let Some(previous) = catalog.objects.get(&identifier) {
                        let previous_archive = catalog.archive_name(previous.slot.archive_index)?;
                        return Err(Error::Archive(format!(
                            "Object {identifier} appears in both Numbers archives {previous_archive} and {archive_name}"
                        )));
                    }
                    let object_index = u32::try_from(object_index).map_err(|_| {
                        Error::InvalidFormat(
                            "Numbers object catalog object index exceeds u32".to_owned(),
                        )
                    })?;
                    let record = NumbersObjectRecord {
                        slot: NumbersObjectSlot {
                            archive_index,
                            object_index,
                        },
                        shape_message_count: u32::try_from(
                            object
                                .messages
                                .iter()
                                .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
                                .count(),
                        )
                        .map_err(|_| {
                            Error::InvalidFormat(
                                "Numbers shape message count exceeds u32".to_owned(),
                            )
                        })?,
                        caption_message_count: u32::try_from(
                            object
                                .messages
                                .iter()
                                .filter(|message| {
                                    message.type_ == STANDIN_CAPTION_MESSAGE_TYPE
                                })
                                .count(),
                        )
                        .map_err(|_| {
                            Error::InvalidFormat(
                                "Numbers caption message count exceeds u32".to_owned(),
                            )
                        })?,
                        storage_message_count: u32::try_from(
                            object
                                .messages
                                .iter()
                                .filter(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
                                .count(),
                        )
                        .map_err(|_| {
                            Error::InvalidFormat(
                                "Numbers storage message count exceeds u32".to_owned(),
                            )
                        })?,
                    };
                    catalog.objects.insert(identifier, record);
                    catalog.stats.objects_indexed = catalog
                        .stats
                        .objects_indexed
                        .checked_add(1)
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "Numbers object catalog object counter overflow".to_owned(),
                            )
                        })?;

                    if requested_sheets.contains(&identifier) {
                        catalog.record_semantic_decode(format!("sheet object {identifier}"))?;
                        let (_, sheet) = decode_sheet(object)?;
                        catalog.stats.sheet_objects_scanned = catalog
                            .stats
                            .sheet_objects_scanned
                            .checked_add(1)
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "Numbers object catalog sheet counter overflow".to_owned(),
                                )
                            })?;
                        let drawable_ids = sheet
                            .drawable_infos
                            .into_iter()
                            .map(|reference| reference.identifier)
                            .collect::<Vec<_>>();
                        for drawable_id in &drawable_ids {
                            catalog.stats.reference_edges = catalog
                                .stats
                                .reference_edges
                                .checked_add(1)
                                .ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "Numbers object catalog reference counter overflow"
                                            .to_owned(),
                                    )
                                })?;
                            if catalog.stats.reference_edges > limits.max_reference_edges {
                                return Err(Error::InvalidFormat(format!(
                                    "Numbers object catalog reference count {} exceeds the {} edge budget at sheet {identifier} drawable {drawable_id}",
                                    catalog.stats.reference_edges,
                                    limits.max_reference_edges
                                )));
                            }
                            let count = catalog
                                .drawable_owner_counts
                                .entry(*drawable_id)
                                .or_insert(0);
                            *count = count.checked_add(1).ok_or_else(|| {
                                Error::InvalidFormat(
                                    "Numbers drawable ownership counter overflow".to_owned(),
                                )
                            })?;
                            if *count > 1 {
                                catalog.duplicate_drawables.insert(*drawable_id);
                            } else {
                                catalog
                                    .drawable_owners
                                    .insert(*drawable_id, identifier);
                            }
                        }
                        if catalog.drawable_owner_counts.len() > limits.max_drawables {
                            return Err(Error::InvalidFormat(format!(
                                "Numbers object catalog drawable count {} exceeds the {} drawable budget at sheet {identifier}",
                                catalog.drawable_owner_counts.len(),
                                limits.max_drawables
                            )));
                        }
                        catalog.sheets.insert(
                            identifier,
                            NumbersSheetSummary {
                                slot: record.slot,
                                drawable_ids: drawable_ids.into_boxed_slice(),
                            },
                        );
                    }
                }
                Ok(())
            })?;
        }

        for sheet_id in requested_sheet_ids {
            if !catalog.sheets.contains_key(&sheet_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet object {sheet_id} is missing"
                )));
            }
        }

        let mut drawable_ids_by_archive = BTreeMap::<u32, Vec<u64>>::new();
        for drawable_id in catalog.drawable_owners.keys().copied() {
            let record = catalog.objects.get(&drawable_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers drawable {drawable_id} is missing from the object catalog"
                ))
            })?;
            drawable_ids_by_archive
                .entry(record.slot.archive_index)
                .or_default()
                .push(drawable_id);
        }
        for drawable_ids in drawable_ids_by_archive.values_mut() {
            drawable_ids.sort_unstable();
        }

        for (archive_index, drawable_ids) in drawable_ids_by_archive {
            let archive_name = catalog.archive_name(archive_index)?.to_owned();
            catalog.with_archive(package, &archive_name, |catalog, archive| {
                for drawable_id in drawable_ids {
                    let record = *catalog.objects.get(&drawable_id).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers drawable {drawable_id} is missing from the object catalog"
                        ))
                    })?;
                    let object = archive
                        .objects
                        .get(usize::try_from(record.slot.object_index).map_err(|_| {
                            Error::InvalidFormat(
                                "Numbers drawable object index does not fit usize".to_owned(),
                            )
                        })?)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "Numbers drawable {drawable_id} object slot is out of bounds"
                            ))
                        })?;
                    if object.archive_info.identifier != Some(drawable_id) {
                        return Err(Error::InvalidFormat(format!(
                            "Numbers drawable {drawable_id} object slot does not match its identifier"
                        )));
                    }
                    catalog.stats.drawable_objects_scanned = catalog
                        .stats
                        .drawable_objects_scanned
                        .checked_add(1)
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "Numbers object catalog drawable counter overflow".to_owned(),
                            )
                        })?;
                    let owner_sheet_id = *catalog.drawable_owners.get(&drawable_id).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers drawable {drawable_id} has no owning sheet"
                        ))
                    })?;
                    if record.shape_message_count == 0 {
                        continue;
                    }
                    let drawable_owner_count = *catalog
                        .drawable_owner_counts
                        .get(&drawable_id)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "Numbers drawable {drawable_id} has no owner count"
                            ))
                        })?;
                    let mut decoded_shape = None;
                    for message in object
                        .messages
                        .iter()
                        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
                    {
                        catalog.record_semantic_decode(format!("drawable object {drawable_id}"))?;
                        let shape = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
                        if record.shape_message_count == 1 {
                            decoded_shape = Some(NumbersShapeSummary {
                                is_text_box: shape.is_text_box,
                                parent_id: shape
                                    .super_
                                    .super_
                                    .parent
                                    .as_ref()
                                    .map(|reference| reference.identifier),
                                owned_storage_id: shape
                                    .owned_storage
                                    .as_ref()
                                    .map(|reference| reference.identifier),
                                deprecated_storage_id: shape
                                    .deprecated_storage
                                    .as_ref()
                                    .map(|reference| reference.identifier),
                                title_id: shape
                                    .super_
                                    .super_
                                    .title
                                    .as_ref()
                                    .map(|reference| reference.identifier),
                                caption_id: shape
                                    .super_
                                    .super_
                                    .caption
                                    .as_ref()
                                    .map(|reference| reference.identifier),
                            });
                        }
                        if let Some(storage_id) = shape
                            .owned_storage
                            .as_ref()
                            .map(|reference| reference.identifier)
                        {
                            let owner_count = catalog
                                .storage_owner_counts
                                .entry(storage_id)
                                .or_insert(0);
                            *owner_count = owner_count
                                .checked_add(drawable_owner_count)
                                .ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "Numbers storage ownership counter overflow".to_owned(),
                                    )
                                })?;
                        }
                    }
                    if let Some(shape) = decoded_shape {
                        catalog.drawables.insert(
                            drawable_id,
                            NumbersDrawableSummary {
                                owner_sheet_id,
                                slot: record.slot,
                                shape,
                            },
                        );
                    }
                }
                Ok(())
            })?;
        }

        if catalog
            .drawables
            .values()
            .any(|drawable| drawable.shape.is_text_box == Some(true))
        {
            if package.contains_entry(PACKAGE_METADATA_ENTRY) {
                catalog.record_semantic_decode(PACKAGE_METADATA_ENTRY)?;
            }
            catalog.uuid_identifiers =
                component_uuid_identifiers(package, DOCUMENT_COMPONENT_IDENTIFIER)?;
        }
        Ok(catalog)
    }

    fn with_archive<T, F>(
        &mut self,
        package: &IWorkPackage,
        archive_name: &str,
        read: F,
    ) -> Result<T>
    where
        F: FnOnce(&mut Self, &Archive) -> Result<T>,
    {
        if self.stats.archive_reads >= self.limits.max_archive_reads {
            return Err(Error::InvalidFormat(format!(
                "Numbers object catalog archive reads {} exceed the {} operation budget at {archive_name}",
                self.stats.archive_reads.saturating_add(1),
                self.limits.max_archive_reads
            )));
        }
        self.stats.archive_reads += 1;
        self.stats.peak_live_archives = self.stats.peak_live_archives.max(1);
        package.with_parsed_archive(archive_name, |archive| read(self, archive))
    }

    fn record_semantic_decode(&mut self, path: impl std::fmt::Display) -> Result<()> {
        if self.stats.semantic_decodes >= self.limits.max_semantic_decodes {
            return Err(Error::InvalidFormat(format!(
                "Numbers object catalog semantic decodes {} exceed the {} operation budget at {path}",
                self.stats.semantic_decodes.saturating_add(1),
                self.limits.max_semantic_decodes
            )));
        }
        self.stats.semantic_decodes += 1;
        Ok(())
    }

    fn archive_name(&self, archive_index: u32) -> Result<&str> {
        self.archive_names
            .get(usize::try_from(archive_index).map_err(|_| {
                Error::InvalidFormat("Numbers archive index does not fit usize".to_owned())
            })?)
            .map(Box::<str>::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers object catalog archive slot {archive_index} is missing"
                ))
            })
    }

    fn ensure_current(&self, package: &IWorkPackage) -> Result<()> {
        if package.mutation_revision() != self.package_revision {
            return Err(Error::InvalidFormat(
                "Numbers object catalog is stale after package mutation".to_owned(),
            ));
        }
        Ok(())
    }

    pub(super) fn sheet_drawable_ids<'a>(
        &'a self,
        package: &IWorkPackage,
        sheet_id: u64,
    ) -> Result<&'a [u64]> {
        self.ensure_current(package)?;
        self.sheets
            .get(&sheet_id)
            .map(|sheet| sheet.drawable_ids.as_ref())
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers sheet object {sheet_id} is not reachable exactly once"
                ))
            })
    }

    #[cfg(test)]
    pub(super) fn stats(&self) -> NumbersObjectCatalogStats {
        self.stats
    }

    #[cfg(test)]
    pub(super) fn text_box_graph(
        &self,
        package: &IWorkPackage,
        sheet_id: u64,
        drawable_id: u64,
    ) -> Result<NumbersTextBoxGraph> {
        self.ensure_current(package)?;
        self.text_box_graph_current(sheet_id, drawable_id)
    }

    pub(super) fn text_box_graph_if_supported(
        &self,
        package: &IWorkPackage,
        sheet_id: u64,
        drawable_id: u64,
    ) -> Result<Option<NumbersTextBoxGraph>> {
        self.ensure_current(package)?;
        let record = self.objects.get(&drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} is missing from the object catalog"
            ))
        })?;
        if record.shape_message_count == 0 {
            return Ok(None);
        }
        if record.shape_message_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} must have exactly one shape payload"
            )));
        }
        let Some(drawable) = self.drawables.get(&drawable_id) else {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} shape payload is unavailable"
            )));
        };
        if drawable.shape.is_text_box != Some(true) {
            return Ok(None);
        }
        self.text_box_graph_current(sheet_id, drawable_id).map(Some)
    }

    pub(super) fn text_storage_info(
        &mut self,
        package: &IWorkPackage,
        storage_id: u64,
    ) -> Result<TextStorageInfo> {
        self.ensure_current(package)?;
        let record = *self.objects.get(&storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text storage {storage_id} is missing from the object catalog"
            ))
        })?;
        let archive_name = self.archive_name(record.slot.archive_index)?.to_owned();
        self.with_archive(package, &archive_name, |catalog, archive| {
            let object = archive
                .objects
                .get(usize::try_from(record.slot.object_index).map_err(|_| {
                    Error::InvalidFormat(
                        "Numbers text storage object index does not fit usize".to_owned(),
                    )
                })?)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers text storage {storage_id} object slot is out of bounds"
                    ))
                })?;
            if object.archive_info.identifier != Some(storage_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers text storage {storage_id} object slot does not match its identifier"
                )));
            }
            let mut found = None;
            for (message_index, message) in object.messages.iter().enumerate() {
                if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
                    continue;
                }
                catalog.record_semantic_decode(format!(
                    "storage object {storage_id} message {message_index}"
                ))?;
                let storage = match tswp::StorageArchive::decode(message.data.as_slice()) {
                    Ok(storage) => storage,
                    Err(_error)
                        if message.type_ == 2_022
                            && tswp::ParagraphStyleArchive::decode(message.data.as_slice())
                                .is_ok() =>
                    {
                        continue;
                    },
                    Err(error) => {
                        return Err(Error::InvalidFormat(format!(
                            "iWork text storage {storage_id} has a malformed writable payload in {archive_name} message {message_index}: {error}"
                        )));
                    },
                };
                if found.is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} must have exactly one writable payload"
                    )));
                }
                found = Some((message.type_, storage));
            }
            let (message_type, storage) = found.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers text storage {storage_id} has no writable payload"
                ))
            })?;
            Ok(TextStorageInfo {
                object_id: storage_id,
                message_type,
                kind: storage.kind,
                text: storage.text.concat(),
            })
        })
    }

    fn text_box_graph_current(
        &self,
        sheet_id: u64,
        drawable_id: u64,
    ) -> Result<NumbersTextBoxGraph> {
        let sheet = self.sheets.get(&sheet_id).ok_or_else(|| {
            Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not reachable exactly once"
            ))
        })?;
        let record = self.objects.get(&drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} is missing from the object catalog"
            ))
        })?;
        if record.shape_message_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} must have exactly one shape payload"
            )));
        }
        if sheet
            .drawable_ids
            .iter()
            .filter(|identifier| **identifier == drawable_id)
            .count()
            != 1
        {
            return Err(Error::ParseError(format!(
                "Numbers sheet {sheet_id} does not own drawable {drawable_id} exactly once"
            )));
        }
        if self.duplicate_drawables.contains(&drawable_id)
            || self.drawable_owner_counts.get(&drawable_id).copied() != Some(1)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} is owned more than once"
            )));
        }
        let drawable = self.drawables.get(&drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} is missing from the object catalog"
            ))
        })?;
        let shape = drawable.shape;
        if shape.is_text_box != Some(true) {
            return Err(Error::ParseError(format!(
                "Numbers drawable {drawable_id} is not an ordinary text box"
            )));
        }
        if drawable.owner_sheet_id != sheet_id {
            return Err(Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has a different owning sheet"
            )));
        }
        if drawable.slot.archive_index != sheet.slot.archive_index {
            let archive_name = self.archive_name(sheet.slot.archive_index)?;
            return Err(Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} is outside sheet component {archive_name}"
            )));
        }
        if shape.parent_id != Some(sheet_id) {
            return Err(Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} does not name sheet {sheet_id} as its parent"
            )));
        }
        let storage_id = shape.owned_storage_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has no owned storage"
            ))
        })?;
        if shape.deprecated_storage_id != Some(storage_id) {
            return Err(Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has inconsistent storage ownership"
            )));
        }
        let title_id = shape.title_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has no title stand-in"
            ))
        })?;
        let caption_id = shape.caption_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has no caption stand-in"
            ))
        })?;
        let object_ids = vec![drawable_id, caption_id, title_id, storage_id];
        if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
            return Err(Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has aliased private objects"
            )));
        }

        for (identifier, kind) in [
            (caption_id, "caption stand-in"),
            (title_id, "title stand-in"),
            (storage_id, "storage"),
        ] {
            let record = self.objects.get(&identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers text-box {kind} {identifier} is missing"))
            })?;
            if record.slot.archive_index != sheet.slot.archive_index {
                let archive_name = self.archive_name(sheet.slot.archive_index)?;
                return Err(Error::InvalidFormat(format!(
                    "Numbers text-box {kind} {identifier} is outside {archive_name}"
                )));
            }
            let count = if identifier == storage_id {
                record.storage_message_count
            } else {
                record.caption_message_count
            };
            if count != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Numbers text-box {kind} {identifier} must have exactly one expected payload"
                )));
            }
        }

        if self.storage_owner_counts.get(&storage_id).copied() != Some(1) {
            return Err(Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has a storage {storage_id} with multiple owners"
            )));
        }

        let uuid_object_ids = match &self.uuid_identifiers {
            Some(mapped) => {
                let mapped_count = object_ids
                    .iter()
                    .filter(|identifier| mapped.contains(identifier))
                    .count();
                if mapped_count != object_ids.len() {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers text box {drawable_id} is missing document-component UUID mappings"
                    )));
                }
                object_ids.clone()
            },
            None => Vec::new(),
        };

        Ok(NumbersTextBoxGraph {
            sheet_id,
            archive_name: self.archive_name(sheet.slot.archive_index)?.to_owned(),
            drawable_id,
            storage_id,
            object_ids,
            uuid_object_ids,
        })
    }
}

pub(super) fn numbers_sheet_drawable_owners(package: &IWorkPackage) -> Result<HashMap<u64, u64>> {
    let document = numbers_document(package)?;
    let locations = object_locations(package)?;
    let mut owners = HashMap::new();
    for sheet_reference in document.sheets {
        let sheet_id = sheet_reference.identifier;
        let archive_name = locations.get(&sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet object {sheet_id} is missing"))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet object {sheet_id} is missing"))
        })?;
        let (_, sheet) = decode_sheet(object)?;
        for drawable in sheet.drawable_infos {
            if let Some(previous_owner) = owners.insert(drawable.identifier, sheet_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers drawable {} is owned more than once by sheets {previous_owner} and {sheet_id}",
                    drawable.identifier
                )));
            }
        }
    }
    Ok(owners)
}

pub(super) fn update_numbers_document<F>(package: &mut IWorkPackage, update: F) -> Result<()>
where
    F: FnOnce(&mut tn::DocumentArchive) -> Result<()>,
{
    package.update_archive("Index/Document.iwa", |archive| {
        let object = archive
            .object_mut(1)
            .ok_or_else(|| Error::InvalidFormat("Numbers root object 1 is missing".to_owned()))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                detect_application_from_document(&message.data) == Some(Application::Numbers)
                    && tn::DocumentArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers root document payload is missing".to_owned())
            })?;
        let original = object.messages[message_index].data.clone();
        let mut document = tn::DocumentArchive::decode(original.as_slice())?;
        let previous_sheet_order = document
            .sheets
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let previous_sheets = previous_sheet_order.iter().copied().collect::<HashSet<_>>();
        update(&mut document)?;
        let current_sheet_order = document
            .sheets
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let current_sheets = current_sheet_order.iter().copied().collect::<HashSet<_>>();
        let message_type = object.messages[message_index].type_;
        let data =
            rewrite_reference_list(&original, 1, &previous_sheet_order, &current_sheet_order)?;
        let verified = tn::DocumentArchive::decode(data.as_slice())?;
        if verified
            .sheets
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>()
            != current_sheet_order
        {
            return Err(Error::InvalidFormat(
                "Numbers sheet-order wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let references = &mut object.archive_info.message_infos[message_index].object_references;
        let describes_sheets = references
            .iter()
            .any(|identifier| previous_sheets.contains(identifier));
        references.retain(|identifier| {
            !previous_sheets.contains(identifier) || current_sheets.contains(identifier)
        });
        if describes_sheets {
            for &identifier in &current_sheet_order {
                if !references.contains(&identifier) {
                    references.push(identifier);
                }
            }
        }
        for field in &mut object.archive_info.message_infos[message_index].field_infos {
            let described_sheets = field
                .object_references
                .iter()
                .any(|identifier| previous_sheets.contains(identifier));
            field.object_references.retain(|identifier| {
                !previous_sheets.contains(identifier) || current_sheets.contains(identifier)
            });
            if described_sheets {
                for &identifier in &current_sheet_order {
                    if !field.object_references.contains(&identifier) {
                        field.object_references.push(identifier);
                    }
                }
            }
        }
        Ok(())
    })
}

pub(super) fn decode_sheet(
    object: &crate::archive::ArchiveObject,
) -> Result<(usize, tn::SheetArchive)> {
    object
        .messages
        .iter()
        .enumerate()
        .find_map(|(index, message)| {
            if message.type_ == 3 {
                tn::FormBasedSheetArchive::decode(message.data.as_slice())
                    .ok()
                    .map(|form| (index, form.super_))
            } else {
                tn::SheetArchive::decode(message.data.as_slice())
                    .ok()
                    .map(|sheet| (index, sheet))
            }
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {:?} has no Numbers sheet payload",
                object.archive_info.identifier
            ))
        })
}

pub(super) fn numbers_sheet(
    package: &IWorkPackage,
    sheet_id: u64,
) -> Result<(String, usize, tn::SheetArchive)> {
    let document = numbers_document(package)?;
    if document
        .sheets
        .iter()
        .filter(|reference| reference.identifier == sheet_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet object {sheet_id} is not reachable exactly once"
        )));
    }
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&sheet_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?
        .to_owned();
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(sheet_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
    let (message_index, sheet) = decode_sheet(object)?;
    Ok((archive_name, message_index, sheet))
}

#[allow(deprecated)]
pub(super) fn numbers_text_box_graph(
    package: &IWorkPackage,
    sheet_id: u64,
    drawable_id: u64,
) -> Result<NumbersTextBoxGraph> {
    let (archive_name, _, sheet) = numbers_sheet(package, sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own drawable {drawable_id} exactly once"
        )));
    }

    let locations = object_locations(package)?;
    if locations.get(&drawable_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers text box {drawable_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = package.archive(&archive_name)?;
    let object = archive.object(drawable_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers text box {drawable_id} is missing"))
    })?;
    let shape_messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if shape_messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Numbers text box {drawable_id} must have exactly one shape payload"
        )));
    }
    let shape = tswp::ShapeInfoArchive::decode(shape_messages[0].data.as_slice())?;
    if shape.is_text_box != Some(true) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_id} is not an ordinary text box"
        )));
    }
    if shape
        .super_
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(sheet_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers text box {drawable_id} does not name sheet {sheet_id} as its parent"
        )));
    }
    let storage_id = shape
        .owned_storage
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has no owned storage"
            ))
        })?;
    if shape
        .deprecated_storage
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(storage_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers text box {drawable_id} has inconsistent storage ownership"
        )));
    }
    let title_id = shape
        .super_
        .super_
        .title
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has no title stand-in"
            ))
        })?;
    let caption_id = shape
        .super_
        .super_
        .caption
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers text box {drawable_id} has no caption stand-in"
            ))
        })?;
    let object_ids = vec![drawable_id, caption_id, title_id, storage_id];
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers text box {drawable_id} has aliased private objects"
        )));
    }

    for (identifier, message_types, label) in [
        (
            caption_id,
            &[STANDIN_CAPTION_MESSAGE_TYPE][..],
            "caption stand-in",
        ),
        (
            title_id,
            &[STANDIN_CAPTION_MESSAGE_TYPE][..],
            "title stand-in",
        ),
        (storage_id, STORAGE_MESSAGE_TYPES, "storage"),
    ] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers text-box {label} {identifier} is outside {archive_name}"
            )));
        }
        let private_object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers text-box {label} {identifier} is missing"))
        })?;
        let matches = private_object
            .messages
            .iter()
            .filter(|message| message_types.contains(&message.type_))
            .count();
        if matches != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers text-box {label} {identifier} must have exactly one expected payload"
            )));
        }
    }

    let document = numbers_document(package)?;
    let mut storage_owners = 0usize;
    let mut drawable_owners = 0usize;
    for reference in document.sheets {
        let candidate_archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", reference.identifier))
        })?;
        let candidate_archive = package.archive(candidate_archive_name)?;
        let candidate_object = candidate_archive
            .object(reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {} is missing", reference.identifier))
            })?;
        let (_, candidate_sheet) = decode_sheet(candidate_object)?;
        for drawable in candidate_sheet.drawable_infos {
            if drawable.identifier == drawable_id {
                drawable_owners += 1;
            }
            let Some(candidate_archive_name) = locations.get(&drawable.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers drawable {} is missing",
                    drawable.identifier
                )));
            };
            let candidate_archive = package.archive(candidate_archive_name)?;
            let Some(candidate_object) = candidate_archive.object(drawable.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers drawable {} is missing",
                    drawable.identifier
                )));
            };
            for message in candidate_object
                .messages
                .iter()
                .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            {
                let candidate = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
                if candidate
                    .owned_storage
                    .as_ref()
                    .map(|reference| reference.identifier)
                    == Some(storage_id)
                {
                    storage_owners += 1;
                }
            }
        }
    }
    if drawable_owners != 1 || storage_owners != 1 {
        return Err(Error::InvalidFormat(format!(
            "Numbers text box {drawable_id} has {drawable_owners} sheet owners and storage {storage_id} has {storage_owners} drawable owners"
        )));
    }

    let uuid_object_ids = match component_uuid_identifiers(package, DOCUMENT_COMPONENT_IDENTIFIER)?
    {
        Some(mapped) => {
            let mapped_count = object_ids
                .iter()
                .filter(|identifier| mapped.contains(identifier))
                .count();
            if mapped_count != object_ids.len() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers text box {drawable_id} is missing document-component UUID mappings"
                )));
            }
            object_ids.clone()
        },
        None => Vec::new(),
    };

    Ok(NumbersTextBoxGraph {
        sheet_id,
        archive_name,
        drawable_id,
        storage_id,
        object_ids,
        uuid_object_ids,
    })
}

pub(super) fn remap_numbers_reference_paths(
    data: &[u8],
    paths: &[&[u32]],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    paths.iter().try_fold(data.to_vec(), |data, path| {
        transform_length_delimited_fields_at_path(&data, path, |reference| {
            let decoded = crate::protobuf::tsp::Reference::decode(reference)?;
            let Some(identifier) = remap.get(&decoded.identifier).copied() else {
                return Ok(reference.to_vec());
            };
            let data = patch_varint_field(reference, 1, true, Some(identifier))?;
            if crate::protobuf::tsp::Reference::decode(data.as_slice())?.identifier != identifier {
                return Err(Error::InvalidFormat(
                    "Numbers reference wire remap failed validation".to_owned(),
                ));
            }
            Ok(data)
        })
    })
}

#[allow(deprecated)]
pub(super) fn remap_numbers_shape_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 1, 2],
        &[1, 1, 6],
        &[1, 1, 9],
        &[1, 1, 10],
        &[1, 1, 11],
        &[1, 2],
        &[2],
        &[3],
        &[4],
    ];
    let mut expected = tswp::ShapeInfoArchive::decode(data)?;
    remap_numbers_shape(&mut expected, remap);
    let data = remap_numbers_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tswp::ShapeInfoArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers ShapeInfoArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_numbers_drawable_archive(drawable: &mut tsd::DrawableArchive, remap: &HashMap<u64, u64>) {
    remap_numbers_reference(&mut drawable.parent, remap);
    remap_numbers_reference(&mut drawable.comment, remap);
    for reference in &mut drawable.pencil_annotations {
        if let Some(identifier) = remap.get(&reference.identifier) {
            reference.identifier = *identifier;
        }
    }
    remap_numbers_reference(&mut drawable.title, remap);
    remap_numbers_reference(&mut drawable.caption, remap);
}

fn remap_numbers_chart_wire(
    data: &[u8],
    recorded_references: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    let mut expected = crate::charts::IWorkChartArchive::decode(data)?;
    expected.remap_references(remap, recorded_references)?;
    let data = expected.encode()?;
    if crate::charts::IWorkChartArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_numbers_chart_mediator_wire(
    data: &[u8],
    recorded_references: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    let mut expected = tn::ChartMediatorArchive::decode(data)?;
    let known_reference = expected
        .super_
        .info
        .as_ref()
        .map(|reference| reference.identifier);
    if let Some(identifier) = recorded_references
        .iter()
        .copied()
        .find(|identifier| remap.contains_key(identifier) && Some(*identifier) != known_reference)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart mediator has an unrecognized private reference {identifier}"
        )));
    }
    remap_numbers_reference(&mut expected.super_.info, remap);
    let data = remap_numbers_reference_paths(data, &[&[1, 1]], remap)?;
    if tn::ChartMediatorArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart mediator wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_numbers_image_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[3],
        &[5],
        &[6],
        &[8],
    ];
    let mut expected = tsd::ImageArchive::decode(data)?;
    remap_numbers_drawable_archive(&mut expected.super_, remap);
    remap_numbers_reference(&mut expected.database_data, remap);
    remap_numbers_reference(&mut expected.style, remap);
    remap_numbers_reference(&mut expected.mask, remap);
    remap_numbers_reference(&mut expected.database_thumbnail_data, remap);
    remap_numbers_reference(&mut expected.database_original_data, remap);
    let data = remap_numbers_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::ImageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers ImageArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_numbers_movie_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[10],
        &[11],
        &[19],
    ];
    let mut expected = tsd::MovieArchive::decode(data)?;
    remap_numbers_drawable_archive(&mut expected.super_, remap);
    remap_numbers_reference(&mut expected.database_movie_data, remap);
    remap_numbers_reference(&mut expected.database_poster_image_data, remap);
    remap_numbers_reference(&mut expected.database_audio_only_image_data, remap);
    remap_numbers_reference(&mut expected.style, remap);
    let data = remap_numbers_reference_paths(data, REFERENCE_PATHS, remap)?;
    if tsd::MovieArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers MovieArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_numbers_storage_wire(
    data: &[u8],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    const OBJECT_TABLE_FIELDS: &[u32] = &[5, 7, 8, 9, 11, 12, 15, 16, 17, 18, 21, 22, 23, 27, 28];
    let mut expected = tswp::StorageArchive::decode(data)?;
    remap_numbers_storage(&mut expected, remap);
    let mut data = remap_numbers_reference_paths(data, &[&[2]], remap)?;
    for field in OBJECT_TABLE_FIELDS {
        data = remap_numbers_reference_paths(&data, &[&[*field, 1, 2]], remap)?;
    }
    for field in [25, 26] {
        data = remap_numbers_reference_paths(&data, &[&[field, 1, 2]], remap)?;
    }
    if tswp::StorageArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers StorageArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

#[allow(deprecated)]
fn remap_numbers_caption_info_wire(data: &[u8], remap: &HashMap<u64, u64>) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 1, 1, 2],
        &[1, 1, 1, 6],
        &[1, 1, 1, 9],
        &[1, 1, 1, 10],
        &[1, 1, 1, 11],
        &[1, 1, 2],
        &[1, 2],
        &[1, 3],
        &[1, 4],
        &[2],
    ];
    let mut expected = crate::protobuf::tsa::CaptionInfoArchive::decode(data)?;
    remap_numbers_shape(&mut expected.super_, remap);
    remap_numbers_reference(&mut expected.placement, remap);
    let data = remap_numbers_reference_paths(data, REFERENCE_PATHS, remap)?;
    if crate::protobuf::tsa::CaptionInfoArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers CaptionInfoArchive wire remap failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn clone_numbers_drawable_graph_object(
    source: &ArchiveObject,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let old_identifier = source.archive_info.identifier.ok_or_else(|| {
        Error::InvalidFormat("Numbers drawable object has no identifier".to_owned())
    })?;
    let new_identifier = *remap.get(&old_identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "No clone identifier allocated for Numbers object {old_identifier}"
        ))
    })?;
    let mut messages = Vec::with_capacity(source.messages.len());
    for (message, info) in source
        .messages
        .iter()
        .zip(&source.archive_info.message_infos)
    {
        let data = match message.type_ {
            crate::charts::source::CHART_MESSAGE_TYPE => {
                remap_numbers_chart_wire(&message.data, &info.object_references, remap)?
            },
            crate::charts::source::CHART_PRESET_MESSAGE_TYPE => {
                crate::charts::source::remap_chart_preset_wire(
                    &message.data,
                    &info.object_references,
                    remap,
                )?
            },
            crate::charts::source::CHART_MEDIATOR_MESSAGE_TYPE => {
                remap_numbers_chart_mediator_wire(&message.data, &info.object_references, remap)?
            },
            SHAPE_INFO_MESSAGE_TYPE => remap_numbers_shape_wire(&message.data, remap)?,
            IMAGE_MESSAGE_TYPE => remap_numbers_image_wire(&message.data, remap)?,
            MOVIE_MESSAGE_TYPE => remap_numbers_movie_wire(&message.data, remap)?,
            2_001 | 2_022 => remap_numbers_storage_wire(&message.data, remap)?,
            CAPTION_INFO_MESSAGE_TYPE => remap_numbers_caption_info_wire(&message.data, remap)?,
            STANDIN_CAPTION_MESSAGE_TYPE => message.data.clone(),
            _ => {
                if info
                    .object_references
                    .iter()
                    .any(|identifier| remap.contains_key(identifier))
                {
                    return Err(Error::InvalidFormat(format!(
                        "Cannot safely clone Numbers message type {} with private drawable-graph references",
                        message.type_
                    )));
                }
                message.data.clone()
            },
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    clone_numbers_object_metadata(source, new_identifier, messages, remap)
}

pub(super) fn clone_numbers_object_metadata(
    source: &ArchiveObject,
    new_identifier: u64,
    messages: Vec<RawMessage>,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let mut cloned = ArchiveObject::new(new_identifier, messages)?;
    cloned.archive_info.should_merge = source.archive_info.should_merge;
    for ((target, source), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&source.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_| {
            Error::Archive("IWA message payload exceeds the u32 format limit".to_owned())
        })?;
        *target = source.clone();
        target.length = length;
        target.object_references = source
            .object_references
            .iter()
            .map(|identifier| remap.get(identifier).copied().unwrap_or(*identifier))
            .collect();
        for field in &mut target.field_infos {
            for identifier in &mut field.object_references {
                if let Some(replacement) = remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
        }
    }
    Ok(cloned)
}

pub(super) fn remap_numbers_reference(
    reference: &mut Option<crate::protobuf::tsp::Reference>,
    remap: &HashMap<u64, u64>,
) {
    if let Some(reference) = reference
        && let Some(identifier) = remap.get(&reference.identifier)
    {
        reference.identifier = *identifier;
    }
}

pub(super) fn remap_numbers_required_reference(
    reference: &mut crate::protobuf::tsp::Reference,
    remap: &HashMap<u64, u64>,
) {
    if let Some(identifier) = remap.get(&reference.identifier) {
        reference.identifier = *identifier;
    }
}

#[allow(deprecated)]
pub(super) fn remap_numbers_shape(shape: &mut tswp::ShapeInfoArchive, remap: &HashMap<u64, u64>) {
    let drawable = &mut shape.super_.super_;
    remap_numbers_reference(&mut drawable.parent, remap);
    remap_numbers_reference(&mut drawable.comment, remap);
    for reference in &mut drawable.pencil_annotations {
        if let Some(identifier) = remap.get(&reference.identifier) {
            reference.identifier = *identifier;
        }
    }
    remap_numbers_reference(&mut drawable.title, remap);
    remap_numbers_reference(&mut drawable.caption, remap);
    remap_numbers_reference(&mut shape.super_.style, remap);
    remap_numbers_reference(&mut shape.deprecated_storage, remap);
    remap_numbers_reference(&mut shape.text_flow, remap);
    remap_numbers_reference(&mut shape.owned_storage, remap);
}

pub(super) fn remap_numbers_storage(storage: &mut tswp::StorageArchive, remap: &HashMap<u64, u64>) {
    remap_numbers_reference(&mut storage.style_sheet, remap);
    for table in [
        &mut storage.table_para_style,
        &mut storage.table_list_style,
        &mut storage.table_char_style,
        &mut storage.table_attachment,
        &mut storage.table_smartfield,
        &mut storage.table_layout_style,
        &mut storage.table_bookmark,
        &mut storage.table_footnote,
        &mut storage.table_section,
        &mut storage.table_rubyfield,
        &mut storage.table_insertion,
        &mut storage.table_deletion,
        &mut storage.table_highlight,
        &mut storage.table_tatechuyoko,
        &mut storage.table_drop_cap_style,
    ]
    .into_iter()
    .flatten()
    {
        for entry in &mut table.entries {
            remap_numbers_reference(&mut entry.object, remap);
        }
    }
    for table in [
        &mut storage.table_overlapping_highlight,
        &mut storage.table_pencil_annotation,
    ]
    .into_iter()
    .flatten()
    {
        for entry in &mut table.entries {
            remap_numbers_required_reference(&mut entry.field, remap);
        }
    }
}

pub(super) fn offset_numbers_drawable_clone(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    offset: f32,
) -> Result<()> {
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "Numbers drawable duplicate offset must be finite".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers drawable {drawable_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {drawable_id} must have exactly one shape payload"
            )));
        }
        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let shape = tswp::ShapeInfoArchive::decode(original)?;
        let position = shape
            .super_
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.position.as_ref())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers drawable {drawable_id} has no positioned geometry"
                ))
            })?;
        let x = position.x + offset;
        let y = position.y + offset;
        if !x.is_finite() || !y.is_finite() {
            return Err(Error::ParseError(
                "Numbers drawable duplicate position overflow".to_owned(),
            ));
        }
        let data = patch_nested_fixed32_field(original, &[1, 1, 1, 1, 1], true, Some(x.to_bits()))?;
        let data = patch_nested_fixed32_field(&data, &[1, 1, 1, 1, 2], true, Some(y.to_bits()))?;
        let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
        let verified_position = verified
            .super_
            .super_
            .geometry
            .and_then(|geometry| geometry.position)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers drawable offset removed its position".to_owned())
            })?;
        if verified_position.x != x || verified_position.y != y {
            return Err(Error::InvalidFormat(
                "Numbers drawable offset failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn patch_numbers_sheet_drawable_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    sheet_id: u64,
    remove: Option<u64>,
    add: Option<u64>,
) -> Result<()> {
    if remove.is_some() == add.is_some() {
        return Err(Error::InvalidFormat(
            "Numbers drawable ownership patch must add or remove exactly one object".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive
            .object_mut(sheet_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
        let (message_index, sheet) = decode_sheet(object)?;
        let previous = sheet
            .drawable_infos
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let mut current = previous.clone();
        if let Some(identifier) = remove {
            current.retain(|candidate| *candidate != identifier);
            if current.len() + 1 != previous.len() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} does not own text box {identifier} exactly once"
                )));
            }
        } else if let Some(identifier) = add {
            if current.contains(&identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} already owns drawable {identifier}"
                )));
            }
            current.push(identifier);
        }
        replace_sheet_drawable_references(object, message_index, &previous, &current)?;

        let info = &mut object.archive_info.message_infos[message_index];
        if let Some(identifier) = remove {
            info.object_references
                .retain(|candidate| *candidate != identifier);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|candidate| *candidate != identifier);
            }
        } else if let Some(identifier) = add {
            if !info.object_references.contains(&identifier) {
                info.object_references.push(identifier);
            }
            let existing = previous.iter().copied().collect::<HashSet<_>>();
            for field in &mut info.field_infos {
                if field
                    .object_references
                    .iter()
                    .any(|candidate| existing.contains(candidate))
                    && !field.object_references.contains(&identifier)
                {
                    field.object_references.push(identifier);
                }
            }
        }
        Ok(())
    })
}

pub(super) fn replace_sheet_drawable_references(
    object: &mut ArchiveObject,
    message_index: usize,
    previous: &[u64],
    current: &[u64],
) -> Result<()> {
    let message_type = object.messages[message_index].type_;
    let original = object.messages[message_index].data.as_slice();
    let data = if message_type == 3 {
        transform_length_delimited_field(original, 1, |sheet| {
            rewrite_reference_list(sheet, 2, previous, current)
        })?
    } else {
        rewrite_reference_list(original, 2, previous, current)?
    };
    let verified = if message_type == 3 {
        tn::FormBasedSheetArchive::decode(data.as_slice())?.super_
    } else {
        tn::SheetArchive::decode(data.as_slice())?
    };
    if verified
        .drawable_infos
        .iter()
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>()
        != current
    {
        return Err(Error::InvalidFormat(
            "Numbers sheet drawable-list wire patch failed validation".to_owned(),
        ));
    }
    object.replace_message(
        message_index,
        RawMessage {
            type_: message_type,
            data,
        },
    )?;
    Ok(())
}

pub(super) fn rewrite_reference_list(
    data: &[u8],
    field_number: u32,
    previous: &[u64],
    current: &[u64],
) -> Result<Vec<u8>> {
    let payloads = repeated_length_delimited_payloads(data, field_number)?;
    if payloads.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "protobuf reference field {field_number} has {} raw entries but {} decoded entries",
            payloads.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (&expected, payload) in previous.iter().zip(payloads) {
        let decoded = crate::protobuf::tsp::Reference::decode(payload)?;
        if decoded.identifier != expected {
            return Err(Error::InvalidFormat(format!(
                "protobuf reference field {field_number} changed during mutation"
            )));
        }
        if existing.insert(expected, payload.to_vec()).is_some() {
            return Err(Error::InvalidFormat(format!(
                "protobuf reference field {field_number} contains duplicate object {expected}"
            )));
        }
    }

    let mut seen = HashSet::with_capacity(current.len());
    let mut replacements = Vec::with_capacity(current.len());
    for &identifier in current {
        if !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "protobuf reference field {field_number} would contain duplicate object {identifier}"
            )));
        }
        replacements.push(existing.remove(&identifier).unwrap_or_else(|| {
            crate::protobuf::tsp::Reference {
                identifier,
                ..Default::default()
            }
            .encode_to_vec()
        }));
    }
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

pub(super) fn decode_table_info(object: &ArchiveObject) -> Result<(usize, tst::TableInfoArchive)> {
    object
        .messages
        .iter()
        .enumerate()
        .find_map(|(index, message)| {
            tst::TableInfoArchive::decode(message.data.as_slice())
                .ok()
                .filter(|info| info.table_model.identifier != 0)
                .map(|info| (index, info))
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {:?} has no Numbers table info payload",
                object.archive_info.identifier
            ))
        })
}

pub(super) fn find_table_model_message(object: &ArchiveObject) -> Result<usize> {
    object
        .messages
        .iter()
        .position(|message| {
            (message.type_ == 6000 || message.type_ == 6001)
                && TableModelArchive::decode(message.data.as_slice()).is_ok()
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {:?} has no Numbers table model payload",
                object.archive_info.identifier
            ))
        })
}

pub(super) fn take_identifier(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    Ok(identifier)
}

pub(super) fn allocate_table_uuid(seed: u64, existing: &HashSet<&str>) -> String {
    let mut suffix = seed & 0x0000_ffff_ffff_ffff;
    loop {
        let candidate = format!("00000000-0000-4000-8000-{suffix:012X}");
        if !existing.contains(candidate.as_str()) {
            return candidate;
        }
        suffix = suffix.wrapping_add(1) & 0x0000_ffff_ffff_ffff;
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum TableOwnedKind {
    Tile,
    Header,
    Data,
    UidMap,
    StrokeSidecar,
}

pub(super) fn table_owned_objects(
    model: &TableModelArchive,
) -> Result<BTreeMap<u64, TableOwnedKind>> {
    let mut objects = BTreeMap::new();
    let mut insert = |identifier: u64, kind: TableOwnedKind| -> Result<()> {
        if identifier == 0 {
            return Ok(());
        }
        if let Some(previous) = objects.insert(identifier, kind)
            && previous != kind
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers table object {identifier} has conflicting storage roles"
            )));
        }
        Ok(())
    };
    let store = &model.base_data_store;
    for reference in &store.row_headers.buckets {
        insert(reference.identifier, TableOwnedKind::Header)?;
    }
    insert(store.column_headers.identifier, TableOwnedKind::Header)?;
    for tile in &store.tiles.tiles {
        insert(tile.tile.identifier, TableOwnedKind::Tile)?;
    }
    for reference in [
        Some(&store.string_table),
        Some(&store.style_table),
        Some(&store.formula_table),
        Some(&store.format_table_pre_bnc),
        store.formula_error_table.as_ref(),
        store.multiple_choice_list_format_table.as_ref(),
        store.merge_region_map.as_ref(),
        store.deprecated_custom_format_table.as_ref(),
        store.rich_text_table.as_ref(),
        store.conditionalstyletable.as_ref(),
        store.comment_storage_table.as_ref(),
        store.import_warning_set_table.as_ref(),
        store.control_cell_spec_table.as_ref(),
        store.format_table.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        insert(reference.identifier, TableOwnedKind::Data)?;
    }
    if let Some(reference) = &model.base_column_row_uids {
        insert(reference.identifier, TableOwnedKind::UidMap)?;
    }
    if let Some(reference) = &model.stroke_sidecar {
        insert(reference.identifier, TableOwnedKind::StrokeSidecar)?;
    }
    Ok(objects)
}

pub(super) fn remap_table_reference(
    reference: &mut crate::protobuf::tsp::Reference,
    remap: &HashMap<u64, u64>,
) {
    if let Some(identifier) = remap.get(&reference.identifier) {
        reference.identifier = *identifier;
    }
}

pub(super) fn remap_optional_table_reference(
    reference: &mut Option<crate::protobuf::tsp::Reference>,
    remap: &HashMap<u64, u64>,
) {
    if let Some(reference) = reference {
        remap_table_reference(reference, remap);
    }
}

#[allow(deprecated)]
pub(super) fn prepare_empty_table_model(
    model: &mut TableModelArchive,
    remap: &HashMap<u64, u64>,
    table_uuid: &str,
    name: &str,
    rows: u32,
    columns: u32,
) -> Result<()> {
    model.table_id = table_uuid.to_owned();
    model.from_table_id = None;
    model.was_cut = Some(false);
    model.table_name = name.to_owned();
    model.number_of_rows = rows;
    model.number_of_columns = columns;
    model.number_of_header_rows = model.number_of_header_rows.map(|value| value.min(rows));
    model.number_of_header_columns = model
        .number_of_header_columns
        .map(|value| value.min(columns));
    model.number_of_footer_rows = model.number_of_footer_rows.map(|value| value.min(rows));
    model.number_of_hidden_rows = Some(0);
    model.number_of_hidden_columns = Some(0);
    model.number_of_user_hidden_rows = Some(0);
    model.number_of_user_hidden_columns = Some(0);
    model.number_of_filtered_rows = Some(0);
    model.provider = None;
    model.hidden_state_formula_owner_for_columns = None;
    model.hidden_state_formula_owner_for_rows = None;
    model.row_filter_set_pre_pivot = None;
    model.conditional_style_formula_owner_id = None;
    model.sort_order = None;
    model.sort_rule_reference_tracker = None;
    model.merge_owner = None;
    model.text_import_record = None;
    model.hidden_states_owner = None;
    model.category_owner_deprecated = None;
    model.pencil_annotation_owner = None;
    model.from_group_by_uid = None;
    model.haunted_owner = None;
    model.pivot_owner = None;
    model.category_owner = None;
    model.pivot_value_types_by_col.clear();
    model.pivot_date_grouping_columns.clear();
    model.pivot_date_grouping_types.clear();
    model.spill_owner = None;

    let store = &mut model.base_data_store;
    for reference in &mut store.row_headers.buckets {
        remap_table_reference(reference, remap);
    }
    remap_table_reference(&mut store.column_headers, remap);
    for tile in &mut store.tiles.tiles {
        remap_table_reference(&mut tile.tile, remap);
    }
    remap_table_reference(&mut store.string_table, remap);
    remap_table_reference(&mut store.style_table, remap);
    remap_table_reference(&mut store.formula_table, remap);
    remap_table_reference(&mut store.format_table_pre_bnc, remap);
    remap_optional_table_reference(&mut store.formula_error_table, remap);
    remap_optional_table_reference(&mut store.multiple_choice_list_format_table, remap);
    remap_optional_table_reference(&mut store.merge_region_map, remap);
    remap_optional_table_reference(&mut store.deprecated_custom_format_table, remap);
    remap_optional_table_reference(&mut store.rich_text_table, remap);
    remap_optional_table_reference(&mut store.conditionalstyletable, remap);
    remap_optional_table_reference(&mut store.comment_storage_table, remap);
    remap_optional_table_reference(&mut store.import_warning_set_table, remap);
    remap_optional_table_reference(&mut store.control_cell_spec_table, remap);
    remap_optional_table_reference(&mut store.format_table, remap);
    store.next_row_strip_id = 1;
    store.next_column_strip_id = 0;
    store.row_tile_tree.nodes.clear();
    store.column_tile_tree.nodes.clear();
    remap_optional_table_reference(&mut model.base_column_row_uids, remap);
    remap_optional_table_reference(&mut model.stroke_sidecar, remap);

    let total_uids = usize::try_from(rows)
        .ok()
        .and_then(|rows| {
            usize::try_from(columns)
                .ok()
                .and_then(|columns| rows.checked_add(columns))
        })
        .ok_or_else(|| Error::ParseError("Numbers table dimensions overflow usize".to_owned()))?;
    if total_uids > MAX_TABLE_UIDS {
        return Err(Error::ParseError(format!(
            "Numbers table dimensions require {total_uids} UIDs, exceeding the safety limit {MAX_TABLE_UIDS}"
        )));
    }
    Ok(())
}

pub(super) fn clone_single_payload_object(
    source: &ArchiveObject,
    new_identifier: u64,
    message_index: usize,
    data: Vec<u8>,
    object_references: Vec<u64>,
    remap: &HashMap<u64, u64>,
    clear_data_references: bool,
) -> Result<ArchiveObject> {
    if source.messages.len() != 1 || message_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "Cannot safely clone multi-payload Numbers object {:?}",
            source.archive_info.identifier
        )));
    }
    let mut cloned = ArchiveObject::new(
        new_identifier,
        vec![RawMessage {
            type_: source.messages[0].type_,
            data,
        }],
    )?;
    cloned.archive_info.should_merge = source.archive_info.should_merge;
    let length = u32::try_from(cloned.messages[0].data.len())
        .map_err(|_| Error::Archive("IWA message payload exceeds u32".to_owned()))?;
    cloned.archive_info.message_infos[0] = source.archive_info.message_infos[0].clone();
    cloned.archive_info.message_infos[0].length = length;
    let info = &mut cloned.archive_info.message_infos[0];
    info.object_references = object_references;
    let final_references = info
        .object_references
        .iter()
        .copied()
        .collect::<HashSet<_>>();
    for field in &mut info.field_infos {
        for identifier in &mut field.object_references {
            if let Some(replacement) = remap.get(identifier) {
                *identifier = *replacement;
            }
        }
        field
            .object_references
            .retain(|identifier| final_references.contains(identifier));
        if clear_data_references {
            field.data_references.clear();
        }
    }
    if clear_data_references {
        info.data_references.clear();
    }
    Ok(cloned)
}

pub(super) fn clone_empty_table_storage(
    source: &ArchiveObject,
    new_identifier: u64,
    kind: TableOwnedKind,
    rows: u32,
    columns: u32,
    seed: u64,
) -> Result<ArchiveObject> {
    if source.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Cannot safely clone multi-payload Numbers storage object {:?}",
            source.archive_info.identifier
        )));
    }
    let message = &source.messages[0];
    let data = match kind {
        TableOwnedKind::Tile => {
            let mut tile = Tile::decode(message.data.as_slice())?;
            tile.max_column = 0;
            tile.max_row = 0;
            tile.num_cells = 0;
            tile.numrows = 0;
            tile.row_infos.clear();
            tile.encode_to_vec()
        },
        TableOwnedKind::Header => {
            let mut headers = tst::HeaderStorageBucket::decode(message.data.as_slice())?;
            headers.headers.clear();
            headers.encode_to_vec()
        },
        TableOwnedKind::Data => {
            if let Ok(mut list) = TableDataList::decode(message.data.as_slice()) {
                list.next_list_id = 1;
                list.entries.clear();
                list.segments.clear();
                list.encode_to_vec()
            } else if let Ok(mut merges) =
                tst::MergeRegionMapArchive::decode(message.data.as_slice())
            {
                merges.cell_range.clear();
                merges.encode_to_vec()
            } else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table data object {:?} has an unsupported payload",
                    source.archive_info.identifier
                )));
            }
        },
        TableOwnedKind::UidMap => empty_uid_map(rows, columns, seed)?.encode_to_vec(),
        TableOwnedKind::StrokeSidecar => tst::StrokeSidecarArchive {
            column_count: Some(columns),
            row_count: Some(rows),
            ..Default::default()
        }
        .encode_to_vec(),
    };
    clone_single_payload_object(
        source,
        new_identifier,
        0,
        data,
        Vec::new(),
        &HashMap::new(),
        true,
    )
}

pub(super) fn empty_uid_map(
    rows: u32,
    columns: u32,
    seed: u64,
) -> Result<tst::ColumnRowUidMapArchive> {
    let rows_usize = rows as usize;
    let columns_usize = columns as usize;
    let mut sorted_column_uids = Vec::new();
    let mut column_index_for_uid = Vec::new();
    let mut column_uid_for_index = Vec::new();
    let mut sorted_row_uids = Vec::new();
    let mut row_index_for_uid = Vec::new();
    let mut row_uid_for_index = Vec::new();
    for values in [
        (&mut sorted_column_uids, columns_usize),
        (&mut sorted_row_uids, rows_usize),
    ] {
        values.0.try_reserve_exact(values.1).map_err(|error| {
            Error::ParseError(format!("Unable to allocate Numbers UID map: {error}"))
        })?;
    }
    for values in [
        (&mut column_index_for_uid, columns_usize),
        (&mut column_uid_for_index, columns_usize),
        (&mut row_index_for_uid, rows_usize),
        (&mut row_uid_for_index, rows_usize),
    ] {
        values.0.try_reserve_exact(values.1).map_err(|error| {
            Error::ParseError(format!("Unable to allocate Numbers UID index: {error}"))
        })?;
    }
    for index in 0..columns {
        sorted_column_uids.push(crate::protobuf::tsp::Uuid {
            lower: seed.wrapping_mul(0x9e37_79b9).wrapping_add(index as u64),
            upper: seed.rotate_left(29) ^ 0x434f_4c55_4d4e_0000,
        });
        column_index_for_uid.push(index);
        column_uid_for_index.push(index);
    }
    for index in 0..rows {
        sorted_row_uids.push(crate::protobuf::tsp::Uuid {
            lower: seed.wrapping_mul(0x517c_c1b7).wrapping_add(index as u64),
            upper: seed.rotate_left(31) ^ 0x524f_5700_0000_0000,
        });
        row_index_for_uid.push(index);
        row_uid_for_index.push(index);
    }
    Ok(tst::ColumnRowUidMapArchive {
        sorted_column_uids,
        column_index_for_uid,
        column_uid_for_index,
        sorted_row_uids,
        row_index_for_uid,
        row_uid_for_index,
    })
}

pub(super) fn table_model_references(model: &TableModelArchive) -> Vec<u64> {
    let mut references = Vec::new();
    let mut push = |identifier: u64| {
        if identifier != 0 {
            references.push(identifier);
        }
    };
    for reference in [
        &model.table_style,
        &model.body_text_style,
        &model.header_row_text_style,
        &model.header_column_text_style,
        &model.footer_row_text_style,
        &model.body_cell_style,
        &model.header_row_style,
        &model.header_column_style,
        &model.footer_row_style,
    ] {
        push(reference.identifier);
    }
    for reference in [
        model.table_name_style.as_ref(),
        model.table_name_shape_style.as_ref(),
        model.table_style_preset.as_ref(),
        model.provider.as_ref(),
        model.hidden_state_formula_owner_for_columns.as_ref(),
        model.hidden_state_formula_owner_for_rows.as_ref(),
        model.row_filter_set_pre_pivot.as_ref(),
        model.base_column_row_uids.as_ref(),
        model.stroke_sidecar.as_ref(),
        model.category_level_1_style.as_ref(),
        model.category_level_2_style.as_ref(),
        model.category_level_3_style.as_ref(),
        model.category_level_4_style.as_ref(),
        model.category_level_5_style.as_ref(),
        model.category_level_1_text_style.as_ref(),
        model.category_level_2_text_style.as_ref(),
        model.category_level_3_text_style.as_ref(),
        model.category_level_4_text_style.as_ref(),
        model.category_level_5_text_style.as_ref(),
        model.label_level_1_style.as_ref(),
        model.label_level_2_style.as_ref(),
        model.label_level_3_style.as_ref(),
        model.label_level_4_style.as_ref(),
        model.label_level_5_style.as_ref(),
        model.label_level_1_text_style.as_ref(),
        model.label_level_2_text_style.as_ref(),
        model.label_level_3_text_style.as_ref(),
        model.label_level_4_text_style.as_ref(),
        model.label_level_5_text_style.as_ref(),
        model.pivot_owner.as_ref(),
        model.category_owner.as_ref(),
        model.pivot_body_summary_row_style.as_ref(),
        model.pivot_body_summary_column_style.as_ref(),
        model.pivot_header_column_summary_style.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        push(reference.identifier);
    }
    let store = &model.base_data_store;
    for reference in &store.row_headers.buckets {
        push(reference.identifier);
    }
    push(store.column_headers.identifier);
    for tile in &store.tiles.tiles {
        push(tile.tile.identifier);
    }
    for reference in [
        Some(&store.string_table),
        Some(&store.style_table),
        Some(&store.formula_table),
        Some(&store.format_table_pre_bnc),
        store.formula_error_table.as_ref(),
        store.multiple_choice_list_format_table.as_ref(),
        store.merge_region_map.as_ref(),
        store.deprecated_custom_format_table.as_ref(),
        store.rich_text_table.as_ref(),
        store.conditionalstyletable.as_ref(),
        store.comment_storage_table.as_ref(),
        store.import_warning_set_table.as_ref(),
        store.control_cell_spec_table.as_ref(),
        store.format_table.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        push(reference.identifier);
    }
    references.sort_unstable();
    references.dedup();
    references
}

pub(super) fn validate_name(name: &str, kind: &str) -> Result<()> {
    if name.is_empty() || name.contains('\0') {
        return Err(Error::ParseError(format!(
            "iWork {kind} names must be non-empty and contain no NUL"
        )));
    }
    Ok(())
}

pub(super) fn validate_table_dimensions(rows: usize, columns: usize) -> Result<(u32, u32)> {
    if rows == 0 || columns == 0 {
        return Err(Error::ParseError(
            "iWork tables must contain at least one row and one column".to_owned(),
        ));
    }
    let total_uids = rows
        .checked_add(columns)
        .ok_or_else(|| Error::ParseError("iWork table dimensions overflow usize".to_owned()))?;
    if total_uids > MAX_TABLE_UIDS {
        return Err(Error::ParseError(format!(
            "iWork table dimensions require {total_uids} UIDs, exceeding the safety limit {MAX_TABLE_UIDS}"
        )));
    }
    let rows = u32::try_from(rows)
        .map_err(|_| Error::ParseError("iWork row count exceeds u32".to_owned()))?;
    let columns = u32::try_from(columns)
        .map_err(|_| Error::ParseError("iWork column count exceeds u32".to_owned()))?;
    Ok((rows, columns))
}

pub(super) struct TableOwner {
    pub(super) sheet_id: u64,
    pub(super) table_info_id: u64,
}

pub(super) fn find_table_owner(package: &IWorkPackage, table_id: u64) -> Result<TableOwner> {
    let document = numbers_document(package)?;
    let locations = object_locations(package)?;
    let mut owner = None;
    for sheet_reference in document.sheets {
        let archive_name = locations.get(&sheet_reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {} is missing",
                sheet_reference.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let sheet_object = archive.object(sheet_reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {} is missing",
                sheet_reference.identifier
            ))
        })?;
        let (_, sheet) = decode_sheet(sheet_object)?;
        for drawable in sheet.drawable_infos {
            let Some(drawable_archive) = locations.get(&drawable.identifier) else {
                continue;
            };
            let object_archive = package.archive(drawable_archive)?;
            let Some(object) = object_archive.object(drawable.identifier) else {
                continue;
            };
            let matches = object.messages.iter().any(|message| {
                tst::TableInfoArchive::decode(message.data.as_slice())
                    .is_ok_and(|info| info.table_model.identifier == table_id)
            });
            if matches {
                if owner.is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers table model {table_id} has multiple owning sheet drawables"
                    )));
                }
                owner = Some(TableOwner {
                    sheet_id: sheet_reference.identifier,
                    table_info_id: drawable.identifier,
                });
            }
        }
    }
    owner.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model {table_id} has no owning sheet drawable"
        ))
    })
}

pub(super) fn remove_object_or_empty_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
) -> Result<()> {
    let archive_name = locations
        .get(&identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))?
        .to_owned();
    let mut archive = package.archive(&archive_name)?;
    archive
        .remove_object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Object {identifier} is missing")))?;
    if archive.objects.is_empty() {
        package.remove_entry(&archive_name).ok_or_else(|| {
            Error::InvalidFormat(format!("Package entry {archive_name} is missing"))
        })?;
    } else {
        package.replace_archive(&archive_name, &archive)?;
    }
    Ok(())
}

pub(super) fn validate_and_trim_tiles(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    rows: usize,
    columns: usize,
) -> Result<()> {
    let tile_size = model
        .base_data_store
        .tiles
        .tile_size
        .unwrap_or(DEFAULT_TILE_SIZE_ROWS);
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    for tile_reference in &model.base_data_store.tiles.tiles {
        let tile_id = tile_reference.tile.identifier;
        let archive_name = locations.get(&tile_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
        })?;
        package.update_archive(archive_name, |archive| {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| Tile::decode(message.data.as_slice()).is_ok())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!("Object {tile_id} has no tile payload"))
                })?;
            let mut tile = Tile::decode(object.messages[message_index].data.as_slice())?;
            let base_row = usize::try_from(tile_reference.tileid)
                .ok()
                .and_then(|tile| tile.checked_mul(tile_size as usize))
                .ok_or_else(|| Error::ParseError("Numbers tile row overflow".to_owned()))?;
            let mut remove_rows = Vec::new();
            for (position, row_info) in tile.row_infos.iter().enumerate() {
                let global_row = base_row
                    .checked_add(row_info.tile_row_index as usize)
                    .ok_or_else(|| Error::ParseError("Numbers row index overflow".to_owned()))?;
                let cells = split_row(row_info)?;
                if global_row >= rows {
                    if cells.iter().any(Option::is_some) {
                        return Err(Error::ParseError(format!(
                            "Cannot shrink Numbers table: row {global_row} contains stored cells"
                        )));
                    }
                    remove_rows.push(position);
                } else if cells.iter().skip(columns).any(Option::is_some) {
                    return Err(Error::ParseError(format!(
                        "Cannot shrink Numbers table: row {global_row} has stored cells beyond column {}",
                        columns.saturating_sub(1)
                    )));
                }
            }
            if !remove_rows.is_empty() {
                let previous = tile.clone();
                for position in remove_rows.into_iter().rev() {
                    tile.row_infos.remove(position);
                }
                tile.numrows = tile
                    .row_infos
                    .iter()
                    .map(|row| row.tile_row_index + 1)
                    .max()
                    .unwrap_or(0);
                let message_type = object.messages[message_index].type_;
                let data = rewrite_tile_wire(
                    object.messages[message_index].data.as_slice(),
                    &previous,
                    &tile,
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
            }
            Ok(())
        })?;
    }
    Ok(())
}

pub(super) fn resize_header_buckets(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    rows: u32,
    columns: u32,
) -> Result<()> {
    let row_buckets = model
        .base_data_store
        .row_headers
        .buckets
        .iter()
        .map(|reference| (reference.identifier, rows));
    let column_bucket = std::iter::once((model.base_data_store.column_headers.identifier, columns));
    for (identifier, limit) in row_buckets.chain(column_bucket) {
        if identifier == 0 {
            continue;
        }
        let archive_name = locations.get(&identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers header bucket object {identifier} is missing"
            ))
        })?;
        package.update_archive(archive_name, |archive| {
            let object = archive.object_mut(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers header bucket object {identifier} is missing"
                ))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    tst::HeaderStorageBucket::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {identifier} has no header bucket payload"
                    ))
                })?;
            let mut bucket =
                tst::HeaderStorageBucket::decode(object.messages[message_index].data.as_slice())?;
            let previous = bucket.clone();
            bucket.headers.retain(|header| header.index < limit);
            let message_type = object.messages[message_index].type_;
            let data = rewrite_header_bucket_wire(
                object.messages[message_index].data.as_slice(),
                &previous,
                &bucket,
            )?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
    }
    Ok(())
}

pub(super) fn resize_uid_map(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    old_rows: usize,
    rows: usize,
    old_columns: usize,
    columns: usize,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                tst::ColumnRowUidMapArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Object {identifier} has no UID map payload"))
            })?;
        let mut map =
            tst::ColumnRowUidMapArchive::decode(object.messages[message_index].data.as_slice())?;
        let previous = map.clone();
        resize_uid_axis(
            &mut map.sorted_row_uids,
            &mut map.row_index_for_uid,
            &mut map.row_uid_for_index,
            old_rows,
            rows,
            "row",
        )?;
        resize_uid_axis(
            &mut map.sorted_column_uids,
            &mut map.column_index_for_uid,
            &mut map.column_uid_for_index,
            old_columns,
            columns,
            "column",
        )?;
        let message_type = object.messages[message_index].type_;
        let data = rewrite_uid_map_wire(
            object.messages[message_index].data.as_slice(),
            &previous,
            &map,
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn resize_uid_axis(
    sorted: &mut Vec<crate::protobuf::tsp::Uuid>,
    index_for_uid: &mut Vec<u32>,
    uid_for_index: &mut Vec<u32>,
    old_len: usize,
    new_len: usize,
    axis: &str,
) -> Result<()> {
    if sorted.len() != old_len || index_for_uid.len() != old_len || uid_for_index.len() != old_len {
        return Err(Error::InvalidFormat(format!(
            "Numbers {axis} UID map lengths do not match table dimensions"
        )));
    }
    if new_len < old_len {
        for logical_index in (new_len..old_len).rev() {
            let sorted_index = uid_for_index.pop().ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers {axis} UID map is truncated"))
            })? as usize;
            if sorted_index >= sorted.len()
                || index_for_uid.get(sorted_index).copied() != Some(logical_index as u32)
            {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {axis} UID map is inconsistent"
                )));
            }
            sorted.remove(sorted_index);
            index_for_uid.remove(sorted_index);
            for value in uid_for_index.iter_mut() {
                if *value > sorted_index as u32 {
                    *value -= 1;
                }
            }
        }
    } else {
        for logical_index in old_len..new_len {
            let lower = sorted
                .iter()
                .map(|uuid| uuid.lower)
                .max()
                .unwrap_or(0)
                .checked_add(1)
                .ok_or_else(|| Error::ParseError(format!("Numbers {axis} UUID overflow")))?;
            let upper = sorted
                .iter()
                .map(|uuid| uuid.upper)
                .max()
                .unwrap_or(0)
                .checked_add(1)
                .ok_or_else(|| Error::ParseError(format!("Numbers {axis} UUID overflow")))?;
            sorted.push(crate::protobuf::tsp::Uuid { lower, upper });
            index_for_uid.push(
                u32::try_from(logical_index)
                    .map_err(|_| Error::ParseError(format!("Numbers {axis} index exceeds u32")))?,
            );
            uid_for_index.push(
                u32::try_from(sorted.len() - 1).map_err(|_| {
                    Error::ParseError(format!("Numbers {axis} UID index exceeds u32"))
                })?,
            );
        }
    }
    Ok(())
}

pub(super) fn resize_stroke_sidecar(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    rows: u32,
    columns: u32,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke sidecar {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers stroke sidecar {identifier} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| tst::StrokeSidecarArchive::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Object {identifier} has no stroke sidecar payload"))
            })?;
        let mut sidecar =
            tst::StrokeSidecarArchive::decode(object.messages[message_index].data.as_slice())?;
        let previous = sidecar.clone();
        sidecar.row_count = Some(rows);
        sidecar.column_count = Some(columns);
        let message_type = object.messages[message_index].type_;
        let data = rewrite_stroke_sidecar_wire(
            object.messages[message_index].data.as_slice(),
            &previous,
            &sidecar,
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

#[derive(Debug, Clone)]
pub(super) struct TableDescriptor {
    pub(super) object_id: u64,
    pub(super) table_info_id: u64,
    pub(super) model: TableModelArchive,
}

pub(super) fn formula_external_tables(
    package: &IWorkPackage,
    descriptors: &[TableDescriptor],
) -> Result<HashMap<u64, ExternalFormulaTable>> {
    let Some(component) = package.calculation_engine_entry_name()? else {
        return Ok(HashMap::new());
    };
    let archive = package.archive(component)?;
    let owners = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == 4008)
        .filter_map(|message| {
            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).ok()
        })
        .filter_map(|owner| {
            let identifier = owner.formula_owner.as_ref()?.identifier;
            Some((identifier, owner))
        })
        .collect::<HashMap<_, _>>();
    let mut tables = HashMap::new();
    for descriptor in descriptors {
        let Some(owner) = owners.get(&descriptor.table_info_id) else {
            continue;
        };
        tables.insert(
            descriptor.object_id,
            ExternalFormulaTable {
                rows: descriptor.model.number_of_rows,
                columns: descriptor.model.number_of_columns,
                owner_uid: owner.formula_owner_uid,
                internal_owner_id: owner.internal_formula_owner_id,
            },
        );
    }
    Ok(tables)
}

pub(super) fn formula_pivot_categories(
    package: &IWorkPackage,
) -> Result<HashMap<PivotFormulaKey, ExternalPivotCategory>> {
    let Some(component) = package.calculation_engine_entry_name()? else {
        return Ok(HashMap::new());
    };
    let archive = package.archive(component)?;
    let mut owner_ids = HashMap::<u32, Vec<u32>>::new();
    let mut groups = Vec::new();
    let mut group_nodes = HashMap::new();
    let mut aggregators = HashMap::new();
    for object in &archive.objects {
        let identifier = object.archive_info.identifier;
        for message in &object.messages {
            match message.type_ {
                4008 => {
                    if let Ok(owner) =
                        tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                        && let Some(kind) = owner.owner_kind
                    {
                        owner_ids
                            .entry(kind)
                            .or_default()
                            .push(owner.internal_formula_owner_id);
                    }
                },
                6373 => {
                    if let Ok(group) = tst::GroupByArchive::decode(message.data.as_slice()) {
                        groups.push(group);
                    }
                },
                6382 => {
                    if let Some(identifier) = identifier
                        && let Ok(aggregator) = tst::group_by_archive::AggregatorArchive::decode(
                            message.data.as_slice(),
                        )
                    {
                        aggregators.insert(identifier, aggregator);
                    }
                },
                6383 => {
                    if let Some(identifier) = identifier
                        && let Ok(node) =
                            tst::group_by_archive::GroupNodeArchive::decode(message.data.as_slice())
                    {
                        group_nodes.insert(identifier, node);
                    }
                },
                _ => {},
            }
        }
    }

    let mut result = HashMap::new();
    let mut ambiguous = HashSet::new();
    for group in groups.into_iter().filter(|group| group.is_enabled) {
        let Some(owner_kind) = group.owner_index else {
            continue;
        };
        let Ok(owner_kind) = u32::try_from(owner_kind) else {
            continue;
        };
        let Some(internal_owner_id) = owner_ids
            .get(&owner_kind)
            .filter(|owners| owners.len() == 1)
            .and_then(|owners| owners.first())
            .copied()
        else {
            continue;
        };
        let Some(grouping_columns) = group
            .grouping_columns_formula
            .as_ref()
            .and_then(expanded_formula_coordinate)
        else {
            continue;
        };
        let aggregate_columns = if group.aggregator.is_empty() {
            group
                .aggregator_ref
                .iter()
                .filter_map(|reference| aggregators.get(&reference.identifier).cloned())
                .collect::<Vec<_>>()
        } else {
            group.aggregator.clone()
        };
        if aggregate_columns.is_empty() {
            continue;
        }
        let root = group.group_node_root.clone().or_else(|| {
            group
                .group_node_root_ref
                .as_ref()
                .and_then(|reference| group_nodes.get(&reference.identifier).cloned())
        });
        let Some(root) = root else {
            continue;
        };
        let group_uid = FormulaUuid::new(group.group_by_uid.lower, group.group_by_uid.upper);
        let mut stack = vec![(root, 0u32)];
        let mut visited_references = HashSet::new();
        if let Some(reference) = &group.group_node_root_ref {
            visited_references.insert(reference.identifier);
        }
        let mut visited_nodes = 0usize;
        while let Some((node, depth)) = stack.pop() {
            visited_nodes += 1;
            if visited_nodes > MAX_TABLE_UIDS {
                break;
            }
            let Ok(group_level) = i32::try_from(depth) else {
                continue;
            };
            for (index, aggregate_column) in aggregate_columns.iter().enumerate() {
                let Some(aggregate) = node
                    .agg_formula_coords
                    .get(index)
                    .and_then(expanded_formula_coordinate)
                else {
                    continue;
                };
                let Some(aggregate_type) = group
                    .column_agg_type
                    .iter()
                    .find(|value| value.column_uid == aggregate_column.column_uid)
                    .map(|value| value.agg_type)
                else {
                    continue;
                };
                let key = PivotFormulaKey::new(
                    group_uid,
                    FormulaUuid::new(
                        aggregate_column.column_uid.lower,
                        aggregate_column.column_uid.upper,
                    ),
                    FormulaUuid::new(node.group_uid.lower, node.group_uid.upper),
                );
                let value = ExternalPivotCategory {
                    internal_owner_id,
                    grouping_columns,
                    aggregate,
                    aggregate_type,
                    group_level,
                    label: pivot_group_label(&node).or_else(|| {
                        (group_level == 0 && node.group_uid.lower == 1 && node.group_uid.upper == 0)
                            .then(|| "Grand Total".to_owned())
                    }),
                };
                if !ambiguous.contains(&key)
                    && let Some(previous) = result.insert(key, value.clone())
                    && previous != value
                {
                    result.remove(&key);
                    ambiguous.insert(key);
                }
            }
            let child_depth = depth.saturating_add(1);
            stack.extend(
                node.child
                    .into_iter()
                    .rev()
                    .map(|child| (child, child_depth)),
            );
            for reference in node.child_ref.into_iter().rev() {
                if visited_references.insert(reference.identifier)
                    && let Some(child) = group_nodes.get(&reference.identifier)
                {
                    stack.push((child.clone(), child_depth));
                }
            }
        }
    }
    Ok(result)
}

pub(super) fn expanded_formula_coordinate(
    coordinate: &tsce::CellCoordinateArchive,
) -> Option<(u32, u32)> {
    Some((coordinate.row?, coordinate.column?))
}

pub(super) fn pivot_group_label(node: &tst::group_by_archive::GroupNodeArchive) -> Option<String> {
    let value = node.group_cell_value.as_ref()?;
    if let Some(string) = &value.string_value {
        return Some(string.value.clone());
    }
    if let Some(number) = &value.number_value
        && let Some(number) = number.value
    {
        return Some(number.to_string());
    }
    if let Some(boolean) = &value.boolean_value {
        return Some(if boolean.value { "TRUE" } else { "FALSE" }.to_owned());
    }
    value.date_value.as_ref().map(|date| date.value.to_string())
}

#[derive(Debug, Clone)]
pub(super) enum EncodedValue {
    Clear,
    ClearValuePreservingMetadata,
    /// A previously validated BNC-v5 payload relocated without touching its
    /// referenced string, rich-text, formula, or comment table entries.
    Raw(Vec<u8>),
    Number(f64),
    Boolean(bool),
    Date(f64),
    Duration(f64),
    String(u32),
    RichText(u32),
    Formula(u32),
    FormulaCachedNumber(f64),
    FormulaCachedBoolean(bool),
    Comment(Option<u32>),
    ConditionalStyle {
        identifier: Option<u32>,
        applied_rule: Option<u32>,
    },
}

pub(super) struct CellLocation {
    pub(super) descriptor: TableDescriptor,
    pub(super) object_locations: HashMap<u64, String>,
    pub(super) tile_archive: String,
    pub(super) tile_id: u64,
    pub(super) tile_row: u32,
}

pub(super) fn locate_cell(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<CellLocation> {
    let descriptor = table_models(package)?
        .into_iter()
        .find(|table| table.object_id == table_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table object {table_id} not found")))?;
    locate_cell_in_descriptor(package, descriptor, row, column)
}

pub(super) fn locate_attached_cell(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<CellLocation> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    locate_cell_in_descriptor(package, descriptor, row, column)
}

fn locate_cell_in_descriptor(
    package: &IWorkPackage,
    descriptor: TableDescriptor,
    row: usize,
    column: usize,
) -> Result<CellLocation> {
    let locations = object_locations(package)?;
    let (tile_archive, tile_id, tile_row) =
        cell_tile_location(&descriptor, &locations, row, column)?;
    let tile_archive = tile_archive.to_owned();
    Ok(CellLocation {
        descriptor,
        object_locations: locations,
        tile_archive,
        tile_id,
        tile_row,
    })
}

fn cell_tile_location<'a>(
    descriptor: &TableDescriptor,
    locations: &'a HashMap<u64, String>,
    row: usize,
    column: usize,
) -> Result<(&'a str, u64, u32)> {
    if row >= descriptor.model.number_of_rows as usize
        || column >= descriptor.model.number_of_columns as usize
    {
        return Err(Error::ParseError(format!(
            "Cell ({row}, {column}) is outside Numbers table {:?} dimensions {}x{}",
            descriptor.model.table_name,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns
        )));
    }

    let tile_size = descriptor
        .model
        .base_data_store
        .tiles
        .tile_size
        .unwrap_or(DEFAULT_TILE_SIZE_ROWS);
    if tile_size == 0 {
        return Err(Error::ParseError(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    let row_u32 =
        u32::try_from(row).map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?;
    let tile_key = row_u32 / tile_size;
    let tile_row = row_u32 % tile_size;
    let tile_id = descriptor
        .model
        .base_data_store
        .tiles
        .tiles
        .iter()
        .find(|tile| tile.tileid == tile_key)
        .map(|tile| tile.tile.identifier)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Numbers table {:?} has no tile for row {row}",
                descriptor.model.table_name
            ))
        })?;
    let tile_archive = locations
        .get(&tile_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers tile object {tile_id} is missing")))?;
    Ok((tile_archive, tile_id, tile_row))
}

pub(super) fn attached_table_descriptor(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<TableDescriptor> {
    let locations = object_locations(package)?;
    let model_archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::ParseError(format!("iWork table model object {table_id} not found"))
    })?;
    let model_archive = package.archive(model_archive_name)?;
    let model_object = model_archive.object(table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
    })?;
    let models = decode_attached_table_models(model_object.messages.iter(), table_id)?;
    let [model] = models.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork table model {table_id} must contain exactly one table-model payload"
        )));
    };

    let mut table_info_id = None;
    let archive_names = locations.values().collect::<HashSet<_>>();
    for archive_name in archive_names {
        let archive = package.archive(archive_name)?;
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if identifier == table_id {
                continue;
            }
            let owns_model = object.messages.iter().any(|message| {
                tst::TableInfoArchive::decode(message.data.as_slice())
                    .is_ok_and(|info| info.table_model.identifier == table_id)
            });
            if owns_model && table_info_id.replace(identifier).is_some() {
                return Err(Error::InvalidFormat(format!(
                    "iWork table model {table_id} has multiple table-info owners"
                )));
            }
        }
    }
    Ok(TableDescriptor {
        object_id: table_id,
        table_info_id: table_info_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table model {table_id} has no table-info owner"
            ))
        })?,
        model: model.clone(),
    })
}

pub(super) fn attached_table_descriptors(package: &IWorkPackage) -> Result<Vec<TableDescriptor>> {
    let locations = object_locations(package)?;
    let archive_names = locations.values().collect::<HashSet<_>>();
    let mut descriptors = BTreeMap::<u64, TableDescriptor>::new();

    for archive_name in archive_names {
        let archive = package.archive(archive_name)?;
        for object in &archive.objects {
            let Some(table_info_id) = object.archive_info.identifier else {
                continue;
            };
            let mut owned_models = HashSet::new();
            for message in &object.messages {
                let Ok(table_info) = tst::TableInfoArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                let model_id = table_info.table_model.identifier;
                let Some(model_archive_name) = locations.get(&model_id) else {
                    continue;
                };
                let model_archive = package.archive(model_archive_name)?;
                let Some(model_object) = model_archive.object(model_id) else {
                    continue;
                };
                let models = decode_attached_table_models(model_object.messages.iter(), model_id)?;
                let [model] = models.as_slice() else {
                    continue;
                };
                if !owned_models.insert(model_id) {
                    continue;
                }
                let descriptor = TableDescriptor {
                    object_id: model_id,
                    table_info_id,
                    model: model.clone(),
                };
                if descriptors.insert(model_id, descriptor).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "iWork table model {model_id} has multiple table-info owners"
                    )));
                }
            }
        }
    }

    Ok(descriptors.into_values().collect())
}

fn decode_attached_table_models<'a>(
    messages: impl Iterator<Item = &'a RawMessage>,
    table_id: u64,
) -> Result<Vec<TableModelArchive>> {
    messages
        .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
        .map(|message| {
            TableModelArchive::decode(message.data.as_slice()).map_err(|error| {
                Error::InvalidFormat(format!(
                    "iWork table model {table_id} contains malformed table-model payload: {error}"
                ))
            })
        })
        .collect()
}

pub(super) fn rename_attached_table_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    name: &str,
) -> Result<()> {
    validate_name(name, "table")?;
    attached_table_descriptor(package, table_id)?;
    let locations = object_locations(package)?;
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                (message.type_ == 6000 || message.type_ == 6001)
                    && TableModelArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {table_id} has no iWork table-model payload"
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let data = patch_length_delimited_field(
            object.messages[message_index].data.as_slice(),
            8,
            true,
            Some(name.as_bytes()),
        )?;
        let verified = TableModelArchive::decode(data.as_slice())?;
        if verified.table_name != name {
            return Err(Error::InvalidFormat(
                "iWork table-name wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn resize_attached_table_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    rows: usize,
    columns: usize,
) -> Result<()> {
    let (rows_u32, columns_u32) = validate_table_dimensions(rows, columns)?;
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_rows = descriptor.model.number_of_rows as usize;
    let old_columns = descriptor.model.number_of_columns as usize;
    if (rows, columns) == (old_rows, old_columns) {
        return Ok(());
    }

    let locations = object_locations(package)?;
    validate_and_trim_tiles(package, &locations, &descriptor.model, rows, columns)?;
    resize_header_buckets(
        package,
        &locations,
        &descriptor.model,
        rows_u32,
        columns_u32,
    )?;
    if let Some(reference) = &descriptor.model.base_column_row_uids {
        resize_uid_map(
            package,
            &locations,
            reference.identifier,
            old_rows,
            rows,
            old_columns,
            columns,
        )?;
    }
    if let Some(reference) = &descriptor.model.stroke_sidecar {
        resize_stroke_sidecar(
            package,
            &locations,
            reference.identifier,
            rows_u32,
            columns_u32,
        )?;
    }
    let table_archive = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
    })?;
    package.update_archive(table_archive, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                (message.type_ == 6000 || message.type_ == 6001)
                    && TableModelArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {table_id} has no iWork table-model payload"
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let mut data = patch_varint_field(original, 6, true, Some(u64::from(rows_u32)))?;
        data = patch_varint_field(&data, 7, true, Some(u64::from(columns_u32)))?;
        let verified = TableModelArchive::decode(data.as_slice())?;
        if (verified.number_of_rows, verified.number_of_columns) != (rows_u32, columns_u32) {
            return Err(Error::InvalidFormat(
                "iWork table-dimension wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn set_cell_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<()> {
    table_sparse_storage::ensure_cell_storage(package, table_id, row, column)?;
    let location = locate_cell(package, table_id, row, column)?;
    set_cell_at_location(package, location, row, column, value)
}

pub(super) fn set_attached_cell_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<()> {
    table_sparse_storage::ensure_attached_cell_storage(package, table_id, row, column)?;
    let location = locate_attached_cell(package, table_id, row, column)?;
    set_cell_at_location(package, location, row, column, value)
}

/// Return whether an attached table cell stores a native formula reference.
pub(super) fn attached_cell_is_formula(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    let Some(data) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(false);
    };
    Ok(matches!(
        BncCell::parse(&data)?.stored_value(),
        StoredValue::Formula(_)
    ))
}

/// Move one attached table cell's exact BNC payload without changing any
/// referenced table-data-list entry counts.
///
/// This is used when a native merged-cell anchor survives a deletion by moving
/// into the following row or column. Formula dependency coordinates must be
/// updated before invoking this operation.
pub(super) fn relocate_attached_cell_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    source_row: usize,
    source_column: usize,
    destination_row: usize,
    destination_column: usize,
) -> Result<bool> {
    let source = locate_attached_cell(package, table_id, source_row, source_column)?;
    let Some(source_data) = read_tile_cell(
        package,
        &source.tile_archive,
        source.tile_id,
        source.tile_row,
        source_column,
    )?
    else {
        return Ok(false);
    };

    let destination = locate_attached_cell(package, table_id, destination_row, destination_column)?;
    if read_tile_cell(
        package,
        &destination.tile_archive,
        destination.tile_id,
        destination.tile_row,
        destination_column,
    )?
    .is_some()
    {
        if attached_cell_comment_in_package(package, table_id, destination_row, destination_column)?
            .is_some()
        {
            clear_attached_cell_comment_in_package(
                package,
                table_id,
                destination_row,
                destination_column,
            )?;
        }
        set_attached_cell_in_package(
            package,
            table_id,
            destination_row,
            destination_column,
            CellValue::Empty,
        )?;
    }

    let destination_count = update_tile(
        package,
        &destination.tile_archive,
        destination.tile_id,
        destination.tile_row,
        destination_column,
        destination.descriptor.model.number_of_columns as usize,
        EncodedValue::Raw(source_data),
    )?;
    update_row_header(
        package,
        &destination.object_locations,
        &destination.descriptor.model,
        destination_row,
        destination_count,
    )?;
    let source_count = update_tile(
        package,
        &source.tile_archive,
        source.tile_id,
        source.tile_row,
        source_column,
        source.descriptor.model.number_of_columns as usize,
        EncodedValue::Clear,
    )?;
    update_row_header(
        package,
        &source.object_locations,
        &source.descriptor.model,
        source_row,
        source_count,
    )?;
    Ok(true)
}

pub(super) fn set_cells_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    updates: Vec<TableCellUpdate>,
) -> Result<()> {
    validate_batch_values(&updates)?;
    let coordinates = updates
        .iter()
        .map(|update| (update.row, update.column))
        .collect::<Vec<_>>();
    table_sparse_storage::ensure_cells_storage(package, table_id, &coordinates)?;
    let descriptor = table_models(package)?
        .into_iter()
        .find(|table| table.object_id == table_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table object {table_id} not found")))?;
    set_cells_for_descriptor(package, descriptor, updates)
}

pub(super) fn set_attached_cells_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    updates: Vec<TableCellUpdate>,
) -> Result<()> {
    validate_batch_values(&updates)?;
    let coordinates = updates
        .iter()
        .map(|update| (update.row, update.column))
        .collect::<Vec<_>>();
    table_sparse_storage::ensure_attached_cells_storage(package, table_id, &coordinates)?;
    let descriptor = attached_table_descriptor(package, table_id)?;
    set_cells_for_descriptor(package, descriptor, updates)
}

fn validate_batch_values(updates: &[TableCellUpdate]) -> Result<()> {
    if updates
        .iter()
        .any(|update| matches!(&update.value, CellValue::Formula(_) | CellValue::Error(_)))
    {
        return Err(Error::ParseError(
            "Formula and error cell writes require referenced-table construction".to_owned(),
        ));
    }
    Ok(())
}

fn set_cells_for_descriptor(
    package: &mut IWorkPackage,
    descriptor: TableDescriptor,
    updates: Vec<TableCellUpdate>,
) -> Result<()> {
    let locations = object_locations(package)?;
    let mut resolved = Vec::with_capacity(updates.len());
    for update in updates {
        let (tile_archive, tile_id, tile_row) =
            cell_tile_location(&descriptor, &locations, update.row, update.column)?;
        resolved.push((update, tile_archive, tile_id, tile_row));
    }
    for (update, tile_archive, tile_id, tile_row) in resolved {
        let context = CellWriteContext {
            descriptor: &descriptor,
            object_locations: &locations,
            tile_archive,
            tile_id,
            tile_row,
        };
        set_cell_with_context(package, &context, update.row, update.column, update.value)?;
    }
    Ok(())
}

fn set_cell_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<()> {
    let context = CellWriteContext {
        descriptor: &location.descriptor,
        object_locations: &location.object_locations,
        tile_archive: &location.tile_archive,
        tile_id: location.tile_id,
        tile_row: location.tile_row,
    };
    set_cell_with_context(package, &context, row, column, value)
}

struct CellWriteContext<'a> {
    descriptor: &'a TableDescriptor,
    object_locations: &'a HashMap<u64, String>,
    tile_archive: &'a str,
    tile_id: u64,
    tile_row: u32,
}

fn set_cell_with_context(
    package: &mut IWorkPackage,
    context: &CellWriteContext<'_>,
    row: usize,
    column: usize,
    value: CellValue,
) -> Result<()> {
    let CellWriteContext {
        descriptor,
        object_locations,
        tile_archive,
        tile_id,
        tile_row,
    } = context;
    if matches!(value, CellValue::Formula(_) | CellValue::Error(_)) {
        return Err(Error::ParseError(
            "Formula and error cell writes require referenced-table construction".to_string(),
        ));
    }
    let old_cell = read_tile_cell(package, tile_archive, *tile_id, *tile_row, column)?;
    let old_bnc = old_cell.as_deref().map(BncCell::parse).transpose()?;
    let old_formula_error = old_bnc.as_ref().and_then(BncCell::formula_error_identifier);
    let old_comment = old_bnc.as_ref().and_then(BncCell::comment_identifier);
    let old_value = old_bnc
        .as_ref()
        .map_or(StoredValue::Empty, |cell| cell.stored_value());

    if matches!(old_value, StoredValue::Unsupported(_)) {
        return Err(Error::ParseError(format!(
            "Replacing {old_value:?} cells is not yet safe because its referenced table must be updated"
        )));
    }

    if let Some(identifier) = old_formula_error {
        decrement_formula_error_table(package, object_locations, &descriptor.model, identifier)?;
    }

    if let (StoredValue::RichText(identifier), CellValue::Text(replacement)) = (old_value, &value) {
        let replacement_identifier = set_rich_text(
            package,
            object_locations,
            &descriptor.model,
            identifier,
            row,
            column,
            replacement,
        )?;
        if replacement_identifier == identifier && old_formula_error.is_none() {
            return Ok(());
        }
        let cell_count = update_tile(
            package,
            tile_archive,
            *tile_id,
            *tile_row,
            column,
            descriptor.model.number_of_columns as usize,
            EncodedValue::RichText(replacement_identifier),
        )?;
        return update_row_header(
            package,
            object_locations,
            &descriptor.model,
            row,
            cell_count,
        );
    }

    if let StoredValue::RichText(identifier) = old_value {
        release_rich_text(package, object_locations, &descriptor.model, identifier)?;
    }

    let old_string = match old_value {
        StoredValue::Text(identifier) => Some(identifier),
        _ => None,
    };
    let old_formula = match old_value {
        StoredValue::Formula(identifier) => Some(identifier),
        _ => None,
    };
    let encoded_value = match value {
        CellValue::Empty => {
            if let Some(identifier) = old_string {
                update_string_table(
                    package,
                    object_locations,
                    descriptor.model.base_data_store.string_table.identifier,
                    Some(identifier),
                    None,
                )?;
            }
            if old_comment.is_some() {
                EncodedValue::ClearValuePreservingMetadata
            } else {
                EncodedValue::Clear
            }
        },
        CellValue::Text(text) => {
            let identifier = update_string_table(
                package,
                object_locations,
                descriptor.model.base_data_store.string_table.identifier,
                old_string,
                Some(&text),
            )?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers string-table insertion returned no identifier".to_owned(),
                )
            })?;
            EncodedValue::String(identifier)
        },
        CellValue::Number(value) => {
            decrement_old_string(package, object_locations, &descriptor.model, old_string)?;
            EncodedValue::Number(value)
        },
        CellValue::Boolean(value) => {
            decrement_old_string(package, object_locations, &descriptor.model, old_string)?;
            EncodedValue::Boolean(value)
        },
        CellValue::Date(value) => {
            decrement_old_string(package, object_locations, &descriptor.model, old_string)?;
            EncodedValue::Date(value)
        },
        CellValue::Duration(value) => {
            decrement_old_string(package, object_locations, &descriptor.model, old_string)?;
            EncodedValue::Duration(value)
        },
        CellValue::Formula(_) | CellValue::Error(_) => unreachable!("validated above"),
    };

    if let Some(identifier) = old_formula {
        decrement_formula_table(
            package,
            object_locations,
            descriptor.model.base_data_store.formula_table.identifier,
            identifier,
        )?;
        update_formula_dependencies(
            package,
            descriptor.table_info_id,
            row,
            column,
            false,
            &[],
            &[],
        )?;
    }

    let cell_count = update_tile(
        package,
        tile_archive,
        *tile_id,
        *tile_row,
        column,
        descriptor.model.number_of_columns as usize,
        encoded_value,
    )?;
    update_row_header(
        package,
        object_locations,
        &descriptor.model,
        row,
        cell_count,
    )
}

#[derive(Debug, Clone)]
pub(super) struct CommentEntryLocation {
    table_id: u64,
    storage_id: u64,
    storage_archive: String,
    refcount: u32,
    owner: TableDataListEntryOwner,
}

pub(super) fn comment_entry_location(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    identifier: u32,
) -> Result<CommentEntryLocation> {
    let table_id = model
        .base_data_store
        .comment_storage_table
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers cell references comment entry {identifier}, but table {:?} has no comment list",
                model.table_name
            ))
        })?;
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::CommentStorage,
    )?;
    let located = resolved
        .entries
        .iter()
        .find(|candidate| candidate.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers comment table has no entry {identifier}"))
        })?;
    if located.entry.refcount == 0 {
        return Err(Error::InvalidFormat(format!(
            "Numbers comment entry {identifier} has a zero reference count"
        )));
    }
    let storage_id = located
        .entry
        .comment_storage
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers comment entry {identifier} has no storage reference"
            ))
        })?;
    let storage_archive = locations.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers comment storage object {storage_id} is missing"
        ))
    })?;
    let storage_component = package.archive(storage_archive)?;
    let object = storage_component.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers comment storage object {storage_id} is missing"
        ))
    })?;
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == 3056)
        .collect::<Vec<_>>();
    if payloads.len() != 1
        || tsd::CommentStorageArchive::decode(payloads[0].data.as_slice()).is_err()
    {
        return Err(Error::InvalidFormat(format!(
            "Object {storage_id} must contain exactly one TSD comment-storage payload"
        )));
    }
    Ok(CommentEntryLocation {
        table_id,
        storage_id,
        storage_archive: storage_archive.clone(),
        refcount: located.entry.refcount,
        owner: located.owner.clone(),
    })
}

pub(super) fn read_comment_storage(
    package: &IWorkPackage,
    entry: &CommentEntryLocation,
) -> Result<NumbersCellComment> {
    read_comment_storage_object(package, &entry.storage_archive, entry.storage_id)
}

pub(super) fn read_comment_storage_object(
    package: &IWorkPackage,
    archive_name: &str,
    storage_id: u64,
) -> Result<NumbersCellComment> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers comment storage object {storage_id} is missing"
        ))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == 3056)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Object {storage_id} must contain exactly one TSD comment-storage payload"
        )));
    }
    let comment = tsd::CommentStorageArchive::decode(messages[0].data.as_slice())?;
    Ok(NumbersCellComment {
        text: comment.text.unwrap_or_default(),
        creation_date_seconds: comment.creation_date.map(|date| date.seconds),
        author_object_id: comment.author.map(|author| author.identifier),
        reply_object_ids: comment
            .replies
            .into_iter()
            .map(|reply| reply.identifier)
            .collect(),
        storage_uuid: comment.storage_uuid.map(|uuid| NumbersCommentUuid {
            lower: uuid.lower,
            upper: uuid.upper,
        }),
    })
}

pub(super) fn read_comment_storage_by_id(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    storage_id: u64,
) -> Result<NumbersCellComment> {
    let archive_name = locations.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers comment storage object {storage_id} is missing"
        ))
    })?;
    read_comment_storage_object(package, archive_name, storage_id)
}

pub(super) fn cell_comment_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<NumbersCellCommentInfo>> {
    let location = locate_cell(package, table_id, row, column)?;
    cell_comment_at_location(package, location, table_id, row, column)
}

pub(super) fn attached_cell_comment_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<NumbersCellCommentInfo>> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    cell_comment_at_location(package, location, table_id, row, column)
}

fn cell_comment_at_location(
    package: &IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<NumbersCellCommentInfo>> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(None);
    };
    let Some(identifier) = BncCell::parse(&cell)?.comment_identifier() else {
        return Ok(None);
    };
    let entry = comment_entry_location(
        package,
        &location.object_locations,
        &location.descriptor.model,
        identifier,
    )?;
    let comment = read_comment_storage(package, &entry)?;
    Ok(Some(NumbersCellCommentInfo {
        table_id,
        row,
        column,
        list_identifier: identifier,
        storage_object_id: entry.storage_id,
        comment,
    }))
}

pub(super) fn cell_comment_replies_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<NumbersCellCommentReplyInfo>> {
    let location = locate_cell(package, table_id, row, column)?;
    cell_comment_replies_at_location(package, location, table_id, row, column)
}

pub(super) fn attached_cell_comment_replies_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<NumbersCellCommentReplyInfo>> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    cell_comment_replies_at_location(package, location, table_id, row, column)
}

fn cell_comment_replies_at_location(
    package: &IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<NumbersCellCommentReplyInfo>> {
    let root =
        cell_comment_at_location(package, location, table_id, row, column)?.ok_or_else(|| {
            Error::ParseError(format!(
                "iWork table cell ({row}, {column}) in table {table_id} has no comment"
            ))
        })?;
    let locations = object_locations(package)?;
    validate_cell_comment_reply_graph(package, &locations, root.storage_object_id, &root.comment)?;
    root.comment
        .reply_object_ids
        .into_iter()
        .map(|storage_object_id| {
            Ok(NumbersCellCommentReplyInfo {
                table_id,
                row,
                column,
                root_storage_object_id: root.storage_object_id,
                storage_object_id,
                comment: read_comment_storage_by_id(package, &locations, storage_object_id)?,
            })
        })
        .collect()
}

pub(super) fn validate_cell_comment_reply_graph(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    root_storage_id: u64,
    root: &NumbersCellComment,
) -> Result<()> {
    let mut seen = HashSet::new();
    for reply_id in &root.reply_object_ids {
        if *reply_id == root_storage_id || !seen.insert(*reply_id) {
            return Err(Error::InvalidFormat(format!(
                "Numbers comment storage {root_storage_id} contains a duplicate or cyclic reply reference to {reply_id}"
            )));
        }
        read_comment_storage_by_id(package, locations, *reply_id)?;
    }
    Ok(())
}

pub(super) fn validate_cell_comment_reply_reference(
    root_storage_id: u64,
    root: &NumbersCellComment,
    reply_storage_id: u64,
) -> Result<()> {
    if root_storage_id == reply_storage_id {
        return Err(Error::InvalidFormat(format!(
            "Numbers comment storage {root_storage_id} references itself as a reply"
        )));
    }
    match root
        .reply_object_ids
        .iter()
        .filter(|identifier| **identifier == reply_storage_id)
        .count()
    {
        1 => Ok(()),
        0 => Err(Error::ParseError(format!(
            "Numbers comment storage {reply_storage_id} is not a direct reply to {root_storage_id}"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Numbers comment storage {root_storage_id} duplicates reply {reply_storage_id}"
        ))),
    }
}

pub(super) fn ensure_comment_storage_metadata(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    author_id: Option<u64>,
) -> Result<bool> {
    let storage_uuid = fresh_comment_storage_uuid(package)?;
    let mut changed = false;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers comment storage object {storage_id} is missing"
            ))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == 3056)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Object {storage_id} must contain exactly one TSD comment-storage payload"
            )));
        }
        let index = indexes[0];
        let original = object.messages[index].data.as_slice();
        let before = tsd::CommentStorageArchive::decode(original)?;
        let mut data = original.to_vec();
        if before.creation_date.is_none() {
            data = patch_length_delimited_field(
                &data,
                2,
                false,
                Some(&current_apple_reference_date()?.encode_to_vec()),
            )?;
        }
        if before.author.is_none()
            && let Some(author_id) = author_id
        {
            data = patch_length_delimited_field(
                &data,
                3,
                false,
                Some(
                    &tsp::Reference {
                        identifier: author_id,
                        ..Default::default()
                    }
                    .encode_to_vec(),
                ),
            )?;
        }
        if before.storage_uuid.is_none() {
            data =
                patch_length_delimited_field(&data, 5, false, Some(&storage_uuid.encode_to_vec()))?;
        }
        if data == original {
            return Ok(());
        }
        let verified = tsd::CommentStorageArchive::decode(data.as_slice())?;
        if verified.creation_date.is_none()
            || (author_id.is_some() && verified.author.is_none())
            || verified.storage_uuid.is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers comment storage {storage_id} metadata patch failed validation"
            )));
        }
        object.replace_message(index, RawMessage { type_: 3056, data })?;
        if before.author.is_none()
            && let Some(author_id) = author_id
            && !object.archive_info.message_infos[index]
                .object_references
                .contains(&author_id)
        {
            object.archive_info.message_infos[index]
                .object_references
                .push(author_id);
        }
        changed = true;
        Ok(())
    })?;
    Ok(changed)
}

pub(super) fn update_comment_storage_text(
    package: &mut IWorkPackage,
    entry: &CommentEntryLocation,
    text: String,
) -> Result<()> {
    package.update_archive(&entry.storage_archive, |archive| {
        let object = archive.object_mut(entry.storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers comment storage object {} is missing",
                entry.storage_id
            ))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 3056)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no TSD comment-storage payload",
                    entry.storage_id
                ))
            })?;
        let comment =
            tsd::CommentStorageArchive::decode(object.messages[message_index].data.as_slice())?;
        let data = patch_length_delimited_field(
            object.messages[message_index].data.as_slice(),
            1,
            comment.text.is_some(),
            Some(text.as_bytes()),
        )?;
        let verified = tsd::CommentStorageArchive::decode(data.as_slice())?;
        if verified.text.as_deref() != Some(text.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers comment storage object {} text patch failed validation",
                entry.storage_id
            )));
        }
        object.replace_message(message_index, RawMessage { type_: 3056, data })?;
        Ok(())
    })
}

pub(super) fn ensure_comment_table(
    package: &mut IWorkPackage,
    location: &CellLocation,
    table_id: u64,
) -> Result<(u64, String)> {
    if let Some(reference) = &location
        .descriptor
        .model
        .base_data_store
        .comment_storage_table
    {
        let archive = location
            .object_locations
            .get(&reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers comment table object {} is missing",
                    reference.identifier
                ))
            })?;
        return Ok((reference.identifier, archive.clone()));
    }

    let model_archive = location
        .object_locations
        .get(&location.descriptor.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table object {} is missing",
                location.descriptor.object_id
            ))
        })?
        .clone();
    package.update_archive(&model_archive, |archive| {
        archive.insert_object(ArchiveObject::new(
            table_id,
            vec![RawMessage {
                type_: 6005,
                data: TableDataList {
                    list_type: tst::table_data_list::ListType::CommentStorage as i32,
                    next_list_id: 1,
                    entries: Vec::new(),
                    segments: Vec::new(),
                    // Native Numbers and Keynote create cell-comment tables in
                    // the original table-data-list mode. Keep the generated
                    // graph wire-compatible with their inline thread storage.
                    is_new_for_bnc: None,
                }
                .encode_to_vec(),
            }],
        )?)?;
        let object = archive
            .object_mut(location.descriptor.object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table object {} is missing",
                    location.descriptor.object_id
                ))
            })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6000 || message.type_ == 6001)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers table-model payload",
                    location.descriptor.object_id
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let mut model = TableModelArchive::decode(object.messages[message_index].data.as_slice())?;
        let previous = model.clone();
        model.base_data_store.comment_storage_table = Some(crate::protobuf::tsp::Reference {
            identifier: table_id,
            ..Default::default()
        });
        let data = rewrite_table_model_comment_table_wire(
            object.messages[message_index].data.as_slice(),
            &previous,
            &model,
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let references = &mut object.archive_info.message_infos[message_index].object_references;
        if !references.contains(&table_id) {
            references.push(table_id);
        }
        Ok(())
    })?;
    Ok((table_id, model_archive))
}

pub(super) fn append_comment_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    storage_id: u64,
) -> Result<u32> {
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::CommentStorage,
    )?;
    // Native Numbers reserves comment identifier 1. Empty comment lists carry
    // next_list_id=1, but the first attached cell uses key 2 and advances the
    // counter to 3. Other TableDataList kinds are allowed to use key 1.
    let key = next_table_data_list_key(&resolved.list, &resolved.entries)?.max(2);
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers comment table object {table_id} is missing"
            ))
        })?;
        let message_index =
            table_data_list_message_index(object, tst::table_data_list::ListType::CommentStorage)
                .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {table_id} has no Numbers comment TableDataList payload"
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let previous = TableDataList::decode(original)?;
        let mut list = previous.clone();
        list.next_list_id = key
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers comment identifier overflow".to_owned()))?;
        list.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            comment_storage: Some(crate::protobuf::tsp::Reference {
                identifier: storage_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        let data = rewrite_table_data_list_wire(original, &previous, &list)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let references = &mut object.archive_info.message_infos[message_index].object_references;
        if !references.contains(&storage_id) {
            references.push(storage_id);
        }
        Ok(())
    })?;
    Ok(key)
}

pub(super) fn update_comment_cell(
    package: &mut IWorkPackage,
    location: &CellLocation,
    row: usize,
    column: usize,
    identifier: Option<u32>,
) -> Result<()> {
    let cell_count = update_tile(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
        location.descriptor.model.number_of_columns as usize,
        EncodedValue::Comment(identifier),
    )?;
    update_row_header(
        package,
        &location.object_locations,
        &location.descriptor.model,
        row,
        cell_count,
    )
}

pub(super) fn add_cell_comment_author_reference(
    package: &mut IWorkPackage,
    comment_archive: &str,
    author_archive: Option<&str>,
    author_id: Option<u64>,
) -> Result<()> {
    let (Some(author_archive), Some(author_id)) = (author_archive, author_id) else {
        return Ok(());
    };
    let source_component = component_identifier_for_entry(package, comment_archive)?;
    let author_component = component_identifier_for_entry(package, author_archive)?;
    if let (Some(source_component), Some(author_component)) = (source_component, author_component)
        && source_component != author_component
    {
        add_component_external_reference(package, source_component, author_component, author_id)?;
    }
    Ok(())
}

pub(super) fn add_cell_comment_storage_reference(
    package: &mut IWorkPackage,
    list_archive: &str,
    storage_archive: &str,
    storage_id: u64,
) -> Result<()> {
    let source_component = component_identifier_for_entry(package, list_archive)?;
    let storage_component = component_identifier_for_entry(package, storage_archive)?;
    if let (Some(source_component), Some(storage_component)) = (source_component, storage_component)
        && source_component != storage_component
    {
        add_component_external_reference(package, source_component, storage_component, storage_id)?;
    }
    Ok(())
}

pub(super) fn replace_cell_comment_root(
    package: &mut IWorkPackage,
    location: &CellLocation,
    row: usize,
    column: usize,
    old_identifier: u32,
    new_entry: &CommentEntryLocation,
) -> Result<bool> {
    let locations = object_locations(package)?;
    let resolved = resolve_table_data_list(
        package,
        &locations,
        new_entry.table_id,
        tst::table_data_list::ListType::CommentStorage,
    )?;
    let old = resolved
        .entries
        .iter()
        .find(|candidate| candidate.entry.key == old_identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers comment table has no entry {old_identifier}"
            ))
        })?;
    if old.owner != new_entry.owner || old.entry.refcount != new_entry.refcount {
        return Err(Error::InvalidFormat(format!(
            "Numbers comment entry {old_identifier} changed during copy-on-write"
        )));
    }
    let removed = decrement_table_data_list_entry(
        package,
        &locations,
        &resolved,
        old,
        tst::table_data_list::ListType::CommentStorage,
    )?;
    let locations = object_locations(package)?;
    let new_identifier = append_comment_entry(
        package,
        &locations,
        new_entry.table_id,
        new_entry.storage_id,
    )?;
    add_cell_comment_storage_reference(
        package,
        &resolved.table_archive,
        &new_entry.storage_archive,
        new_entry.storage_id,
    )?;
    update_comment_cell(package, location, row, column, Some(new_identifier))?;
    Ok(removed)
}

pub(super) fn set_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<()> {
    table_sparse_storage::ensure_cell_storage(package, table_id, row, column)?;
    let location = locate_cell(package, table_id, row, column)?;
    set_cell_comment_at_location(package, location, row, column, text)
}

pub(super) fn set_attached_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<()> {
    table_sparse_storage::ensure_attached_cell_storage(package, table_id, row, column)?;
    let location = locate_attached_cell(package, table_id, row, column)?;
    set_cell_comment_at_location(package, location, row, column, text)
}

fn set_cell_comment_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    row: usize,
    column: usize,
    text: String,
) -> Result<()> {
    let old_identifier = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .as_deref()
    .map(BncCell::parse)
    .transpose()?
    .and_then(|cell| cell.comment_identifier());

    if let Some(identifier) = old_identifier {
        let entry = comment_entry_location(
            package,
            &location.object_locations,
            &location.descriptor.model,
            identifier,
        )?;
        let old_comment = read_comment_storage(package, &entry)?;
        validate_cell_comment_reply_graph(
            package,
            &location.object_locations,
            entry.storage_id,
            &old_comment,
        )?;
        if old_comment.text == text
            && old_comment.creation_date_seconds.is_some()
            && old_comment.author_object_id.is_some()
            && old_comment.storage_uuid.is_some()
        {
            return Ok(());
        }
        let (author_id, author_component_entry, created_author) =
            if old_comment.author_object_id.is_none() {
                preferred_or_ensure_table_annotation_author(package)?
            } else {
                (old_comment.author_object_id, None, false)
            };
        if entry.refcount == 1 {
            let text_changed = old_comment.text != text;
            if text_changed {
                update_comment_storage_text(package, &entry, text)?;
            }
            let repaired = ensure_comment_storage_metadata(
                package,
                &entry.storage_archive,
                entry.storage_id,
                author_id,
            )?;
            if !text_changed && !repaired && !created_author {
                return Ok(());
            }
            add_cell_comment_author_reference(
                package,
                &entry.storage_archive,
                author_component_entry.as_deref(),
                author_id,
            )?;
            if created_author && let Some(author_id) = author_id {
                set_package_last_object_identifier(package, author_id)?;
            }
            let mut modified_entries = vec![entry.storage_archive];
            if created_author && let Some(author_entry) = author_component_entry {
                modified_entries.push(author_entry);
            }
            return advance_save_tokens_for_entries(package, &modified_entries);
        }

        let new_storage_id = next_object_identifier(package)?;
        clone_comment_storage_exact(
            package,
            &location.object_locations,
            entry.storage_id,
            new_storage_id,
        )?;
        let new_entry = CommentEntryLocation {
            storage_id: new_storage_id,
            ..entry
        };
        update_comment_storage_text(package, &new_entry, text)?;
        ensure_comment_storage_metadata(
            package,
            &new_entry.storage_archive,
            new_storage_id,
            author_id,
        )?;
        replace_cell_comment_root(package, &location, row, column, identifier, &new_entry)?;
        set_package_last_object_identifier(package, new_storage_id)?;
        add_cell_comment_author_reference(
            package,
            &new_entry.storage_archive,
            author_component_entry.as_deref(),
            author_id,
        )?;
        let mut modified_entries = vec![
            new_entry.storage_archive,
            location.tile_archive,
            location
                .object_locations
                .get(&new_entry.table_id)
                .cloned()
                .unwrap_or_default(),
        ];
        modified_entries.retain(|entry| !entry.is_empty());
        if created_author && let Some(author_entry) = author_component_entry {
            modified_entries.push(author_entry);
        }
        return advance_save_tokens_for_entries(package, &modified_entries);
    }

    let (author_id, author_component_entry, created_author) =
        preferred_or_ensure_table_annotation_author(package)?;
    let comment_table_id = match &location
        .descriptor
        .model
        .base_data_store
        .comment_storage_table
    {
        Some(reference) => reference.identifier,
        None => next_object_identifier(package)?,
    };
    let (_, comment_table_archive) = ensure_comment_table(package, &location, comment_table_id)?;
    let storage_id = next_object_identifier(package)?;
    insert_comment_storage(
        package,
        &comment_table_archive,
        storage_id,
        text,
        author_id,
        fresh_comment_storage_uuid(package)?,
    )?;
    let locations = object_locations(package)?;
    let identifier = append_comment_entry(package, &locations, comment_table_id, storage_id)?;
    update_comment_cell(package, &location, row, column, Some(identifier))?;
    set_package_last_object_identifier(package, storage_id)?;
    add_cell_comment_author_reference(
        package,
        &comment_table_archive,
        author_component_entry.as_deref(),
        author_id,
    )?;
    let mut modified_entries = vec![comment_table_archive, location.tile_archive];
    if created_author && let Some(author_entry) = author_component_entry {
        modified_entries.push(author_entry);
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn required_cell_comment_root_at_location(
    package: &IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<(CellLocation, u32, CommentEntryLocation, NumbersCellComment)> {
    let identifier = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .as_deref()
    .map(BncCell::parse)
    .transpose()?
    .and_then(|cell| cell.comment_identifier())
    .ok_or_else(|| {
        Error::ParseError(format!(
            "iWork table cell ({row}, {column}) in table {table_id} has no comment"
        ))
    })?;
    let entry = comment_entry_location(
        package,
        &location.object_locations,
        &location.descriptor.model,
        identifier,
    )?;
    let root = read_comment_storage(package, &entry)?;
    validate_cell_comment_reply_graph(
        package,
        &location.object_locations,
        entry.storage_id,
        &root,
    )?;
    Ok((location, identifier, entry, root))
}

pub(super) fn add_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<u64> {
    let location = locate_cell(package, table_id, row, column)?;
    add_cell_comment_reply_at_location(package, location, table_id, row, column, text)
}

pub(super) fn add_attached_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<u64> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    add_cell_comment_reply_at_location(package, location, table_id, row, column, text)
}

fn add_cell_comment_reply_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
    text: String,
) -> Result<u64> {
    let (location, identifier, entry, _) =
        required_cell_comment_root_at_location(package, location, table_id, row, column)?;
    let original_locations = location.object_locations.clone();
    let (author_id, author_component_entry, created_author) =
        preferred_or_ensure_table_annotation_author(package)?;
    let new_root_id = next_object_identifier(package)?;
    let root_archive =
        clone_comment_storage_exact(package, &original_locations, entry.storage_id, new_root_id)?;
    let reply_id = next_object_identifier(package)?;
    insert_comment_storage(
        package,
        &root_archive,
        reply_id,
        text,
        author_id,
        fresh_comment_storage_uuid(package)?,
    )?;
    update_comment_reply_reference(package, new_root_id, None, Some(reply_id))?;
    let new_entry = CommentEntryLocation {
        storage_id: new_root_id,
        storage_archive: root_archive.clone(),
        ..entry.clone()
    };
    let removed_root =
        replace_cell_comment_root(package, &location, row, column, identifier, &new_entry)?;
    set_package_last_object_identifier(package, reply_id)?;
    add_cell_comment_author_reference(
        package,
        &root_archive,
        author_component_entry.as_deref(),
        author_id,
    )?;
    let mut modified_entries = cell_comment_modified_entries(&location, &entry, &root_archive);
    if created_author && let Some(author_entry) = author_component_entry {
        modified_entries.push(author_entry);
    }
    if removed_root {
        let mut removed =
            remove_unreferenced_comment_graph(package, &original_locations, entry.storage_id)?;
        cleanup_removed_cell_comment_graph(
            package,
            &original_locations,
            &mut removed,
            &mut modified_entries,
        )?;
        release_package_identifier_suffix(package, &removed.object_ids)?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)?;
    Ok(reply_id)
}

pub(super) fn set_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
    text: String,
) -> Result<u64> {
    let location = locate_cell(package, table_id, row, column)?;
    set_cell_comment_reply_at_location(
        package,
        location,
        table_id,
        row,
        column,
        reply_storage_object_id,
        text,
    )
}

pub(super) fn set_attached_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
    text: String,
) -> Result<u64> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    set_cell_comment_reply_at_location(
        package,
        location,
        table_id,
        row,
        column,
        reply_storage_object_id,
        text,
    )
}

#[allow(clippy::too_many_arguments)]
fn set_cell_comment_reply_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
    text: String,
) -> Result<u64> {
    let (location, identifier, entry, root) =
        required_cell_comment_root_at_location(package, location, table_id, row, column)?;
    validate_cell_comment_reply_reference(entry.storage_id, &root, reply_storage_object_id)?;
    let reply =
        read_comment_storage_by_id(package, &location.object_locations, reply_storage_object_id)?;
    if reply.text == text {
        return Ok(reply_storage_object_id);
    }
    let original_locations = location.object_locations.clone();
    let new_root_id = next_object_identifier(package)?;
    let root_archive =
        clone_comment_storage_exact(package, &original_locations, entry.storage_id, new_root_id)?;
    let new_reply_id = next_object_identifier(package)?;
    let reply_archive = clone_comment_storage_exact(
        package,
        &original_locations,
        reply_storage_object_id,
        new_reply_id,
    )?;
    let updated_locations = object_locations(package)?;
    let new_reply_entry = CommentEntryLocation {
        table_id: entry.table_id,
        storage_id: new_reply_id,
        storage_archive: reply_archive.clone(),
        refcount: 1,
        owner: entry.owner.clone(),
    };
    update_comment_storage_text(package, &new_reply_entry, text)?;
    update_comment_reply_reference(
        package,
        new_root_id,
        Some(reply_storage_object_id),
        Some(new_reply_id),
    )?;
    debug_assert!(updated_locations.contains_key(&new_root_id));
    let new_entry = CommentEntryLocation {
        storage_id: new_root_id,
        storage_archive: root_archive.clone(),
        ..entry.clone()
    };
    let removed_root =
        replace_cell_comment_root(package, &location, row, column, identifier, &new_entry)?;
    set_package_last_object_identifier(package, new_reply_id)?;
    let mut modified_entries = cell_comment_modified_entries(&location, &entry, &root_archive);
    if !modified_entries.contains(&reply_archive) {
        modified_entries.push(reply_archive);
    }
    if removed_root {
        let mut removed =
            remove_unreferenced_comment_graph(package, &original_locations, entry.storage_id)?;
        cleanup_removed_cell_comment_graph(
            package,
            &original_locations,
            &mut removed,
            &mut modified_entries,
        )?;
        release_package_identifier_suffix(package, &removed.object_ids)?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)?;
    Ok(new_reply_id)
}

pub(super) fn remove_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
) -> Result<()> {
    let location = locate_cell(package, table_id, row, column)?;
    remove_cell_comment_reply_at_location(
        package,
        location,
        table_id,
        row,
        column,
        reply_storage_object_id,
    )
}

pub(super) fn remove_attached_cell_comment_reply_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
) -> Result<()> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    remove_cell_comment_reply_at_location(
        package,
        location,
        table_id,
        row,
        column,
        reply_storage_object_id,
    )
}

fn remove_cell_comment_reply_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
    reply_storage_object_id: u64,
) -> Result<()> {
    let (location, identifier, entry, root) =
        required_cell_comment_root_at_location(package, location, table_id, row, column)?;
    validate_cell_comment_reply_reference(entry.storage_id, &root, reply_storage_object_id)?;
    read_comment_storage_by_id(package, &location.object_locations, reply_storage_object_id)?;
    let original_locations = location.object_locations.clone();
    let new_root_id = next_object_identifier(package)?;
    let root_archive =
        clone_comment_storage_exact(package, &original_locations, entry.storage_id, new_root_id)?;
    update_comment_reply_reference(package, new_root_id, Some(reply_storage_object_id), None)?;
    let new_entry = CommentEntryLocation {
        storage_id: new_root_id,
        storage_archive: root_archive.clone(),
        ..entry.clone()
    };
    let removed_root =
        replace_cell_comment_root(package, &location, row, column, identifier, &new_entry)?;
    set_package_last_object_identifier(package, new_root_id)?;
    let mut modified_entries = cell_comment_modified_entries(&location, &entry, &root_archive);
    if removed_root {
        let mut removed =
            remove_unreferenced_comment_graph(package, &original_locations, entry.storage_id)?;
        cleanup_removed_cell_comment_graph(
            package,
            &original_locations,
            &mut removed,
            &mut modified_entries,
        )?;
        release_package_identifier_suffix(package, &removed.object_ids)?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

pub(super) fn cell_comment_modified_entries(
    location: &CellLocation,
    entry: &CommentEntryLocation,
    root_archive: &str,
) -> Vec<String> {
    let mut entries = vec![location.tile_archive.clone(), root_archive.to_owned()];
    if let Some(table_archive) = location.object_locations.get(&entry.table_id)
        && !entries.contains(table_archive)
    {
        entries.push(table_archive.clone());
    }
    entries
}

pub(super) fn clear_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_cell(package, table_id, row, column)?;
    clear_cell_comment_at_location(package, location, row, column)
}

pub(super) fn clear_attached_cell_comment_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    clear_cell_comment_at_location(package, location, row, column)
}

fn clear_cell_comment_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    row: usize,
    column: usize,
) -> Result<()> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(());
    };
    let Some(identifier) = BncCell::parse(&cell)?.comment_identifier() else {
        return Ok(());
    };
    let entry = comment_entry_location(
        package,
        &location.object_locations,
        &location.descriptor.model,
        identifier,
    )?;
    let original_locations = location.object_locations.clone();
    let resolved = resolve_table_data_list(
        package,
        &original_locations,
        entry.table_id,
        tst::table_data_list::ListType::CommentStorage,
    )?;
    let located = resolved
        .entries
        .iter()
        .find(|candidate| candidate.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers comment table has no entry {identifier}"))
        })?;
    let removed = decrement_table_data_list_entry(
        package,
        &original_locations,
        &resolved,
        located,
        tst::table_data_list::ListType::CommentStorage,
    )?;
    update_comment_cell(package, &location, row, column, None)?;
    let mut modified_entries =
        cell_comment_modified_entries(&location, &entry, &entry.storage_archive);
    if removed {
        let mut removed_graph =
            remove_unreferenced_comment_graph(package, &original_locations, entry.storage_id)?;
        cleanup_removed_cell_comment_graph(
            package,
            &original_locations,
            &mut removed_graph,
            &mut modified_entries,
        )?;
        release_package_identifier_suffix(package, &removed_graph.object_ids)?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

#[derive(Debug, Default)]
pub(super) struct RemovedCellCommentGraph {
    object_ids: Vec<u64>,
    author_ids: HashSet<u64>,
}

pub(super) fn remove_unreferenced_comment_graph(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    root: u64,
) -> Result<RemovedCellCommentGraph> {
    let mut pending = vec![root];
    let mut visited = HashSet::new();
    let mut removed = RemovedCellCommentGraph::default();
    while let Some(identifier) = pending.pop() {
        if !visited.insert(identifier) || package_references_object(package, locations, identifier)?
        {
            continue;
        }
        let Some(archive_name) = locations.get(&identifier) else {
            continue;
        };
        if !package.contains_entry(archive_name) {
            continue;
        }
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let mut replies = Vec::new();
        for message in object
            .messages
            .iter()
            .filter(|message| message.type_ == 3056)
        {
            let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
            if let Some(author) = comment.author {
                removed.author_ids.insert(author.identifier);
            }
            replies.extend(comment.replies.into_iter().map(|reply| reply.identifier));
        }
        if let Some(component_identifier) = component_identifier_for_entry(package, archive_name)? {
            remove_component_external_references_to_object(
                package,
                component_identifier,
                identifier,
            )?;
        }
        remove_object_or_empty_entry(package, locations, identifier)?;
        removed.object_ids.push(identifier);
        pending.extend(replies);
    }
    Ok(removed)
}

pub(super) fn cleanup_removed_cell_comment_graph(
    package: &mut IWorkPackage,
    original_locations: &HashMap<u64, String>,
    removed: &mut RemovedCellCommentGraph,
    modified_entries: &mut Vec<String>,
) -> Result<()> {
    for identifier in &removed.object_ids {
        if let Some(entry) = original_locations.get(identifier)
            && !modified_entries.contains(entry)
        {
            modified_entries.push(entry.clone());
        }
    }
    for author_id in std::mem::take(&mut removed.author_ids) {
        if remove_generated_annotation_author_if_unused(package, author_id)? {
            if let Some(entry) = original_locations.get(&author_id)
                && !modified_entries.contains(entry)
            {
                modified_entries.push(entry.clone());
            }
            removed.object_ids.push(author_id);
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn malformed_attached_table_model_payload_is_reported() {
        let messages = [RawMessage {
            type_: TABLE_MODEL_MESSAGE_TYPES[0],
            data: vec![0x80],
        }];

        let error = decode_attached_table_models(messages.iter(), 41)
            .expect_err("malformed table model payload");
        assert!(matches!(
            error,
            Error::InvalidFormat(message)
                if message.contains("iWork table model 41")
                    && message.contains("malformed table-model payload")
        ));
    }

    #[test]
    fn unrelated_messages_do_not_masquerade_as_table_models() {
        let messages = [RawMessage {
            type_: 9_999,
            data: vec![0x80],
        }];

        assert!(
            decode_attached_table_models(messages.iter(), 41)
                .expect("unrelated messages are ignored")
                .is_empty()
        );
    }
}
