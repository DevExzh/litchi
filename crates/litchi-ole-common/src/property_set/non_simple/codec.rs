//! Bounded CFB capture, closure validation, and candidate rendering.

use super::model::{ElementKind, EntryRecord, Limits, State, same_path};
use crate::object::directory::{self, Catalog, EntryKind, Sid};
use crate::property_set::model::{
    CodePage, Guid, IndirectPropertyName, VT_STORAGE, VT_STORED_OBJECT, VT_STREAM,
    VT_STREAMED_OBJECT, VT_VERSIONED_STREAM, Value, invalid,
};
use litchi_cfb::{
    DirectoryEntry, DirectoryNameKey, OleError, OleWriter, SequentialOleWriter,
    SequentialWriteError, SequentialWriterLimits, SequentialWriterOptions, SharedOleFile,
    directory_names_equal, validate_directory_name,
};
use litchi_core::{ReadAt, SourceVersion};
use std::collections::{HashMap, HashSet};
use std::io::{self, Seek, SeekFrom, Write};
use std::sync::Arc;

/// A stable positional adapter over the exact source allocation retained by a
/// non-simple snapshot.  `SharedOleFile::open_with_limits` accepts this
/// adapter, allowing owner-specific CFB limits without copying an `Arc<[u8]>`
/// into an unrelated `Vec<u8>`.
#[derive(Debug)]
struct ArcSource {
    bytes: Arc<[u8]>,
    version: SourceVersion,
}

impl ReadAt for ArcSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_error| io::Error::other("source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

pub(crate) fn capture(source: Arc<[u8]>, limits: Limits) -> Result<Arc<State>, OleError> {
    limits.validate()?;
    let cfb_limits = limits.cfb_limits()?;
    let source_adapter: Arc<dyn ReadAt> = Arc::new(ArcSource {
        bytes: Arc::clone(&source),
        version: SourceVersion::new(0x4f4c_4550_5300_0001, 0),
    });
    let cfb = Arc::new(SharedOleFile::open_with_limits(source_adapter, cfb_limits)?);
    let mut source_entries = Vec::new();
    source_entries
        .try_reserve(limits.max_directory_entries.min(1024))
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set directory entries",
            source,
        })?;
    for entry in cfb.directory_entries() {
        if source_entries.len() >= limits.max_directory_entries {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set directory entries",
                observed: source_entries.len() as u64 + 1,
                maximum: limits.max_directory_entries as u64,
            });
        }
        source_entries
            .try_reserve(1)
            .map_err(|source| OleError::Allocation {
                resource: "non-simple Property Set directory entries",
                source,
            })?;
        source_entries.push(entry);
    }
    preflight_source_entries(&source_entries, limits)?;
    let root_name = try_clone_string(
        &source_entries
            .first()
            .ok_or_else(|| invalid("non-simple Property Set CFB has no root entry"))?
            .name,
        "non-simple Property Set root name",
    )?;

    let mut owned_entries = Vec::new();
    owned_entries
        .try_reserve_exact(source_entries.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set directory entries",
            source,
        })?;
    for entry in source_entries {
        owned_entries.push(shallow_entry(entry)?);
    }
    let directory_limits = directory::Limits {
        max_entries: limits.max_directory_entries,
        max_name_bytes: limits.max_name_bytes,
        max_total_bytes: limits.max_total_name_bytes,
        max_raw_children: limits.max_directory_entries,
        max_raw_depth: limits.max_storage_depth,
    };
    let catalog = Arc::new(Catalog::from_entries(owned_entries, directory_limits)?);
    let records: Arc<[EntryRecord]> = Arc::from(collect_records(&catalog, limits)?);
    let contents_record = find_record(&records, &["CONTENTS"])
        .ok_or_else(|| invalid("non-simple Property Set CONTENTS stream is missing"))?;
    if contents_record.path.len() != 1 || contents_record.kind != Some(ElementKind::Stream) {
        return Err(invalid(
            "non-simple Property Set CONTENTS must be a root stream",
        ));
    }
    if contents_record.size > limits.max_contents_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set CONTENTS bytes",
            observed: contents_record.size,
            maximum: limits.max_contents_bytes,
        });
    }
    let root = records
        .iter()
        .find(|record| record.is_root())
        .cloned()
        .ok_or_else(|| invalid("non-simple Property Set CFB has no root entry"))?;
    let root_metadata = root
        .metadata
        .ok_or_else(|| invalid("non-simple Property Set root metadata is unavailable"))?;
    if root_metadata.kind() != EntryKind::Root {
        return Err(invalid(
            "non-simple Property Set root entry has the wrong kind",
        ));
    }

    Ok(Arc::new(State {
        source,
        cfb,
        records,
        root_metadata,
        root_name,
        limits,
    }))
}

fn shallow_entry(entry: &DirectoryEntry) -> Result<DirectoryEntry, OleError> {
    let name = try_clone_string(&entry.name, "non-simple Property Set directory name")?;
    let clsid = try_clone_string(&entry.clsid, "non-simple Property Set directory CLSID")?;
    Ok(DirectoryEntry {
        sid: entry.sid,
        name,
        entry_type: entry.entry_type,
        sid_left: entry.sid_left,
        sid_right: entry.sid_right,
        sid_child: entry.sid_child,
        clsid,
        state_bits: entry.state_bits,
        creation_time: entry.creation_time,
        modified_time: entry.modified_time,
        start_sector: entry.start_sector,
        size: entry.size,
        is_minifat: entry.is_minifat,
        children: Vec::new(),
    })
}

fn try_clone_string(value: &str, resource: &'static str) -> Result<String, OleError> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| OleError::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

fn preflight_source_entries(entries: &[&DirectoryEntry], limits: Limits) -> Result<(), OleError> {
    if entries.is_empty() {
        return Err(invalid(
            "non-simple Property Set CFB has no directory entries",
        ));
    }
    if entries.len() > limits.max_directory_entries {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set directory entries",
            observed: entries.len() as u64,
            maximum: limits.max_directory_entries as u64,
        });
    }
    let mut stream_count = 0usize;
    let mut total_stream_bytes = 0u64;
    let mut total_name_bytes = 0usize;
    for entry in entries {
        if entry.name.len() > limits.max_name_bytes || entry.clsid.len() > limits.max_name_bytes {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set directory name bytes",
                observed: entry.name.len().max(entry.clsid.len()) as u64,
                maximum: limits.max_name_bytes as u64,
            });
        }
        total_name_bytes = total_name_bytes
            .checked_add(entry.name.len())
            .and_then(|value| value.checked_add(entry.clsid.len()))
            .ok_or_else(|| invalid("non-simple Property Set directory metadata overflows"))?;
        if total_name_bytes > limits.max_total_name_bytes {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set directory metadata bytes",
                observed: total_name_bytes as u64,
                maximum: limits.max_total_name_bytes as u64,
            });
        }
        if entry.entry_type == EntryKind::Stream.raw() {
            stream_count = stream_count
                .checked_add(1)
                .ok_or_else(|| invalid("non-simple Property Set stream count overflows"))?;
            if stream_count > limits.max_streams {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set streams",
                    observed: stream_count as u64,
                    maximum: limits.max_streams as u64,
                });
            }
            if entry.size > limits.max_stream_bytes {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set stream bytes",
                    observed: entry.size,
                    maximum: limits.max_stream_bytes,
                });
            }
            total_stream_bytes = total_stream_bytes
                .checked_add(entry.size)
                .ok_or_else(|| invalid("non-simple Property Set stream bytes overflow"))?;
            if total_stream_bytes > limits.max_total_stream_bytes {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set aggregate stream bytes",
                    observed: total_stream_bytes,
                    maximum: limits.max_total_stream_bytes,
                });
            }
        }
    }
    Ok(())
}

fn collect_records(catalog: &Catalog, limits: Limits) -> Result<Vec<EntryRecord>, OleError> {
    let raw = catalog.raw_entries();
    let root = raw
        .first()
        .ok_or_else(|| invalid("non-simple Property Set CFB has no root entry"))?;
    let mut by_sid = HashMap::<u32, &DirectoryEntry>::new();
    by_sid
        .try_reserve(raw.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set directory SID index",
            source,
        })?;
    for entry in raw {
        if by_sid.insert(entry.sid, entry).is_some() {
            return Err(invalid(
                "non-simple Property Set directory has duplicate SIDs",
            ));
        }
    }

    let mut records = Vec::new();
    records
        .try_reserve_exact(raw.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set element records",
            source,
        })?;
    let root_metadata = catalog
        .metadata(Sid::new(root.sid)?)
        .copied()
        .ok_or_else(|| invalid("non-simple Property Set root metadata is unavailable"))?;
    records.push(EntryRecord {
        path: Vec::new(),
        kind: None,
        metadata: Some(root_metadata),
        raw_kind: root.entry_type,
        size: root.size,
    });

    #[derive(Debug)]
    struct Work {
        sid: u32,
        parent: Vec<String>,
    }
    let mut path_budget = PathBudget::new(limits.max_total_path_bytes);
    let mut stack = Vec::new();
    if root.sid_child != litchi_cfb::consts::NOSTREAM {
        stack
            .try_reserve(1)
            .map_err(|source| OleError::Allocation {
                resource: "non-simple Property Set directory traversal stack",
                source,
            })?;
        stack.push(Work {
            sid: root.sid_child,
            parent: Vec::new(),
        });
    }
    let mut seen = HashSet::new();
    seen.try_reserve(raw.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set directory traversal",
            source,
        })?;
    while let Some(Work { sid, parent }) = stack.pop() {
        if sid == litchi_cfb::consts::NOSTREAM {
            continue;
        }
        if !seen.insert(sid) {
            return Err(invalid(
                "non-simple Property Set directory graph repeats a SID",
            ));
        }
        let entry = by_sid
            .get(&sid)
            .ok_or_else(|| invalid("non-simple Property Set directory link is dangling"))?;
        let mut path = clone_path_bounded(
            &parent,
            "non-simple Property Set element path",
            &mut path_budget,
        )?;
        append_path_component(&mut path, &entry.name, &mut path_budget)?;
        if path.len() > limits.max_storage_depth.saturating_add(1) {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set storage depth",
                observed: path.len() as u64,
                maximum: limits.max_storage_depth.saturating_add(1) as u64,
            });
        }
        let metadata = catalog.metadata(Sid::new(entry.sid)?).copied();
        let kind = match entry.entry_type {
            value if value == EntryKind::Stream.raw() => Some(ElementKind::Stream),
            value if value == EntryKind::Storage.raw() => Some(ElementKind::Storage),
            _ => None,
        };
        records.push(EntryRecord {
            path: clone_path_bounded(
                &path,
                "non-simple Property Set element path",
                &mut path_budget,
            )?,
            kind,
            metadata,
            raw_kind: entry.entry_type,
            size: entry.size,
        });

        // The red-black sibling tree is traversed with an explicit stack so
        // producer-controlled links cannot consume the Rust call stack.
        if entry.sid_right != litchi_cfb::consts::NOSTREAM {
            stack
                .try_reserve(1)
                .map_err(|source| OleError::Allocation {
                    resource: "non-simple Property Set directory traversal stack",
                    source,
                })?;
            let sibling_parent = clone_path_bounded(
                &parent,
                "non-simple Property Set directory traversal",
                &mut path_budget,
            )?;
            stack.push(Work {
                sid: entry.sid_right,
                parent: sibling_parent,
            });
        }
        if (entry.entry_type == EntryKind::Storage.raw()
            || entry.entry_type == EntryKind::Root.raw())
            && entry.sid_child != litchi_cfb::consts::NOSTREAM
        {
            stack
                .try_reserve(1)
                .map_err(|source| OleError::Allocation {
                    resource: "non-simple Property Set directory traversal stack",
                    source,
                })?;
            stack.push(Work {
                sid: entry.sid_child,
                parent: path,
            });
        }
        if entry.sid_left != litchi_cfb::consts::NOSTREAM {
            stack
                .try_reserve(1)
                .map_err(|source| OleError::Allocation {
                    resource: "non-simple Property Set directory traversal stack",
                    source,
                })?;
            stack.push(Work {
                sid: entry.sid_left,
                parent,
            });
        }
    }
    if seen.len() + 1 != raw.len() {
        return Err(invalid(
            "non-simple Property Set directory contains unreachable entries",
        ));
    }
    Ok(records)
}

#[derive(Debug)]
struct PathBudget {
    used: usize,
    maximum: usize,
}

impl PathBudget {
    const fn new(maximum: usize) -> Self {
        Self { used: 0, maximum }
    }

    fn charge(&mut self, bytes: usize) -> Result<(), OleError> {
        let observed = self
            .used
            .checked_add(bytes)
            .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))?;
        if observed > self.maximum {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set copied path bytes",
                observed: observed as u64,
                maximum: self.maximum as u64,
            });
        }
        self.used = observed;
        Ok(())
    }
}

fn clone_path_bounded(
    path: &[String],
    resource: &'static str,
    budget: &mut PathBudget,
) -> Result<Vec<String>, OleError> {
    let bytes = path.iter().try_fold(0usize, |total, component| {
        total
            .checked_add(component.len())
            .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))
    })?;
    budget.charge(bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(path.len())
        .map_err(|source| OleError::Allocation { resource, source })?;
    for component in path {
        let mut clone = String::new();
        clone
            .try_reserve_exact(component.len())
            .map_err(|source| OleError::Allocation { resource, source })?;
        clone.push_str(component);
        output.push(clone);
    }
    Ok(output)
}

fn append_path_component(
    path: &mut Vec<String>,
    component: &str,
    budget: &mut PathBudget,
) -> Result<(), OleError> {
    budget.charge(component.len())?;
    path.try_reserve(1).map_err(|source| OleError::Allocation {
        resource: "non-simple Property Set element path",
        source,
    })?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(component.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set element name",
            source,
        })?;
    owned.push_str(component);
    path.push(owned);
    Ok(())
}

pub(crate) fn find_record<'a>(
    records: &'a [EntryRecord],
    path: &[&str],
) -> Option<&'a EntryRecord> {
    records.iter().find(|record| {
        record.path.len() == path.len()
            && record
                .path
                .iter()
                .zip(path)
                .all(|(left, right)| directory_names_equal(left, right))
    })
}

pub(crate) fn read_stream(state: &State, record: &EntryRecord) -> Result<Arc<[u8]>, OleError> {
    if record.kind != Some(ElementKind::Stream) {
        return Err(invalid("non-simple element is not a stream"));
    }
    if record.size > state.limits.max_stream_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set stream bytes",
            observed: record.size,
            maximum: state.limits.max_stream_bytes,
        });
    }
    let refs = path_refs(
        &record.path,
        "non-simple Property Set stream path references",
    )?;
    let bytes = state.cfb.open_stream(&refs)?;
    if bytes.len() as u64 != record.size {
        return Err(invalid("non-simple Property Set stream size changed"));
    }
    Ok(bytes.into())
}

pub(crate) fn read_contents(state: &State) -> Result<crate::property_set::Stream, OleError> {
    let bytes = read_contents_bytes(state)?;
    preflight_contents(&bytes, state.limits)?;
    let stream = crate::property_set::codec::parse_non_simple_stream(&bytes)?;
    validate_closure(state, &stream, &state.records)?;
    Ok(stream)
}

fn preflight_contents(data: &[u8], limits: Limits) -> Result<(), OleError> {
    if data.len() < 48 {
        return Err(invalid("Property Set CONTENTS stream is too short"));
    }
    let section_count = usize::try_from(u32::from_le_bytes(
        data[24..28]
            .try_into()
            .map_err(|_| invalid("Property Set section count is truncated"))?,
    ))
    .map_err(|_| invalid("Property Set section count is too large"))?;
    if !(1..=2).contains(&section_count) {
        return Err(invalid("Property Set section count is outside 1..=2"));
    }
    let descriptor_end = 28usize
        .checked_add(
            section_count
                .checked_mul(20)
                .ok_or_else(|| invalid("Property Set descriptor table overflows"))?,
        )
        .ok_or_else(|| invalid("Property Set descriptor table overflows"))?;
    if descriptor_end > data.len() {
        return Err(invalid("Property Set descriptor table is truncated"));
    }
    let mut properties = 0usize;
    for index in 0..section_count {
        let descriptor = 28 + index * 20;
        let offset = usize::try_from(u32::from_le_bytes(
            data[descriptor + 16..descriptor + 20]
                .try_into()
                .map_err(|_| invalid("Property Set section offset is truncated"))?,
        ))
        .map_err(|_| invalid("Property Set section offset is too large"))?;
        let count_end = offset
            .checked_add(8)
            .ok_or_else(|| invalid("Property Set section header overflows"))?;
        if count_end > data.len() {
            return Err(invalid("Property Set section header is truncated"));
        }
        let count = usize::try_from(u32::from_le_bytes(
            data[offset + 4..offset + 8]
                .try_into()
                .map_err(|_| invalid("Property Set property count is truncated"))?,
        ))
        .map_err(|_| invalid("Property Set property count is too large"))?;
        properties = properties
            .checked_add(count)
            .ok_or_else(|| invalid("Property Set property count overflows"))?;
        if properties > limits.max_properties {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set properties",
                observed: properties as u64,
                maximum: limits.max_properties as u64,
            });
        }
    }
    Ok(())
}

pub(crate) fn read_contents_bytes(state: &State) -> Result<Arc<[u8]>, OleError> {
    let record = find_record(&state.records, &["CONTENTS"]).ok_or(OleError::StreamNotFound)?;
    if record.path.len() != 1 || record.kind != Some(ElementKind::Stream) {
        return Err(invalid(
            "non-simple Property Set CONTENTS must be a root stream",
        ));
    }
    if record.size > state.limits.max_contents_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set CONTENTS bytes",
            observed: record.size,
            maximum: state.limits.max_contents_bytes,
        });
    }
    read_stream(state, record)
}

pub(crate) fn validate_closure(
    state: &State,
    contents: &crate::property_set::Stream,
    records: &[EntryRecord],
) -> Result<(), OleError> {
    let root = records
        .iter()
        .find(|record| record.is_root())
        .and_then(|record| record.metadata)
        .ok_or_else(|| invalid("non-simple Property Set root metadata is unavailable"))?;
    if root.kind() != EntryKind::Root {
        return Err(invalid(
            "non-simple Property Set root entry has the wrong kind",
        ));
    }
    let root_class_id = root.class_id().unwrap_or(Guid::from_bytes([0; 16]));
    if root_class_id != contents.class_identifier {
        return Err(invalid(
            "non-simple Property Set root CLSID does not match CONTENTS CLSID",
        ));
    }
    let contents_record = find_record(records, &["CONTENTS"])
        .ok_or_else(|| invalid("non-simple Property Set CONTENTS stream is missing"))?;
    if contents_record.path.len() != 1 || contents_record.kind != Some(ElementKind::Stream) {
        return Err(invalid(
            "non-simple Property Set CONTENTS must be a root stream",
        ));
    }

    let mut direct_count = 0usize;
    let mut direct_key_bytes = 0usize;
    for record in records {
        if record.path.len() != 1 {
            continue;
        }
        let name = record
            .path
            .first()
            .ok_or_else(|| invalid("non-simple Property Set path index is corrupt"))?;
        direct_count = direct_count
            .checked_add(1)
            .ok_or_else(|| invalid("non-simple Property Set direct-name count overflows"))?;
        direct_key_bytes = direct_key_bytes
            .checked_add(direct_name_key_bytes(name)?)
            .ok_or_else(|| invalid("non-simple Property Set physical-name key overflows"))?;
    }
    if direct_key_bytes > state.limits.max_total_path_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set copied path bytes",
            observed: direct_key_bytes as u64,
            maximum: state.limits.max_total_path_bytes as u64,
        });
    }
    let mut direct_index = HashMap::<DirectoryNameKey, usize>::new();
    direct_index
        .try_reserve(direct_count)
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set indirect path index",
            source,
        })?;
    let mut index_budget = PathBudget::new(state.limits.max_total_path_bytes);
    for (index, record) in records.iter().enumerate() {
        let Some(name) = record.path.first().filter(|_| record.path.len() == 1) else {
            continue;
        };
        let key = direct_name_key(name, &mut index_budget)?;
        if let Some(previous) = direct_index.insert(key, index) {
            let previous_name = records[previous]
                .path
                .first()
                .ok_or_else(|| invalid("non-simple Property Set path index is corrupt"))?;
            if directory_names_equal(previous_name, name) {
                return Err(invalid(
                    "non-simple Property Set has duplicate case-insensitive root elements",
                ));
            }
            return Err(invalid(
                "non-simple Property Set physical-name index normalization mismatch",
            ));
        }
    }

    let mut property_count = 0usize;
    let mut indirect_count = 0usize;
    for section in &contents.sections {
        if section.property_ids().count() > state.limits.max_properties {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set properties",
                observed: section.property_ids().count() as u64,
                maximum: state.limits.max_properties as u64,
            });
        }
        property_count = property_count
            .checked_add(section.property_ids().count())
            .ok_or_else(|| invalid("non-simple Property Set property count overflows"))?;
        if property_count > state.limits.max_properties {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set properties",
                observed: property_count as u64,
                maximum: state.limits.max_properties as u64,
            });
        }
        for property_id in section.property_ids() {
            let Some(value) = section.property(property_id) else {
                continue;
            };
            let requirement = indirect_requirement(
                value,
                property_id,
                section.codepage.unwrap_or(CodePage::WINDOWS_1252).id(),
                state.limits.max_contents_bytes,
            )?;
            if let Some((name, kind)) = requirement {
                name.validate_for_property(property_id)?;
                validate_indirect_name(
                    state,
                    records,
                    &direct_index,
                    &mut index_budget,
                    &name,
                    kind,
                    &mut indirect_count,
                )?;
            }
        }
    }
    Ok(())
}

fn indirect_requirement(
    value: &Value,
    property_identifier: u32,
    codepage: u16,
    max_contents_bytes: u64,
) -> Result<Option<(IndirectPropertyName, ElementKind)>, OleError> {
    match value {
        Value::Stream(name) => Ok(Some((name.clone(), ElementKind::Stream))),
        Value::Storage(name) => Ok(Some((name.clone(), ElementKind::Storage))),
        Value::StreamedObject(name) => Ok(Some((name.clone(), ElementKind::Stream))),
        Value::StoredObject(name) => Ok(Some((name.clone(), ElementKind::Storage))),
        Value::VersionedStream(versioned) => {
            let name = IndirectPropertyName::from_wire(
                versioned.stream_name().to_owned(),
                Some(property_identifier),
            )?;
            Ok(Some((name, ElementKind::Stream)))
        },
        Value::Unknown { variant_type, data }
            if matches!(
                *variant_type,
                VT_STREAM
                    | VT_STORAGE
                    | VT_STREAMED_OBJECT
                    | VT_STORED_OBJECT
                    | VT_VERSIONED_STREAM
            ) =>
        {
            let observed = u64::try_from(data.len())
                .ok()
                .and_then(|length| length.checked_add(4))
                .unwrap_or(u64::MAX);
            if observed > max_contents_bytes {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set CONTENTS bytes",
                    observed,
                    maximum: max_contents_bytes,
                });
            }
            let decoded = crate::property_set::codec::decode_non_simple_unknown_property(
                *variant_type,
                data,
                codepage,
                property_identifier,
            )?;
            indirect_requirement(&decoded, property_identifier, codepage, max_contents_bytes)
        },
        _ => Ok(None),
    }
}

fn validate_indirect_name(
    state: &State,
    records: &[EntryRecord],
    direct_index: &HashMap<DirectoryNameKey, usize>,
    index_budget: &mut PathBudget,
    name: &IndirectPropertyName,
    kind: ElementKind,
    indirect_count: &mut usize,
) -> Result<(), OleError> {
    *indirect_count = indirect_count
        .checked_add(1)
        .ok_or_else(|| invalid("non-simple Property Set indirect count overflows"))?;
    if *indirect_count > state.limits.max_indirect_properties {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set indirect properties",
            observed: *indirect_count as u64,
            maximum: state.limits.max_indirect_properties as u64,
        });
    }
    let key = direct_name_key(name.as_str(), index_budget)?;
    let record = direct_index
        .get(&key)
        .and_then(|index| records.get(*index))
        .ok_or_else(|| {
            invalid(format!(
                "non-simple Property Set indirect element {} is missing",
                name.as_str()
            ))
        })?;
    if record.path.len() != 1 || record.kind != Some(kind) {
        return Err(invalid(format!(
            "non-simple Property Set indirect element {} has the wrong kind",
            name.as_str()
        )));
    }
    if !directory_names_equal(&record.path[0], name.as_str()) {
        return Err(invalid(format!(
            "non-simple Property Set indirect element {} has an incompatible physical name",
            name.as_str()
        )));
    }
    Ok(())
}

fn direct_name_key(name: &str, budget: &mut PathBudget) -> Result<DirectoryNameKey, OleError> {
    let key_bytes = direct_name_key_bytes(name)?;
    budget.charge(key_bytes)?;
    DirectoryNameKey::new(name)
}

fn direct_name_key_bytes(name: &str) -> Result<usize, OleError> {
    validate_directory_name(name)?;
    let utf16_len = name.encode_utf16().count();
    utf16_len
        .checked_add(1)
        .and_then(|units| units.checked_mul(2))
        .ok_or_else(|| invalid("non-simple Property Set physical-name key overflows"))
}

#[derive(Debug, Clone)]
pub(crate) struct PayloadEdit {
    pub(crate) path: Vec<String>,
    pub(crate) bytes: Option<Arc<[u8]>>,
}

pub(crate) fn render(
    state: &State,
    records: &[EntryRecord],
    contents: &crate::property_set::Stream,
    contents_bytes: &[u8],
    edits: &[PayloadEdit],
) -> Result<Arc<[u8]>, OleError> {
    if !directory_names_equal(&state.root_name, "Root Entry") {
        return Err(invalid(
            "changed non-simple Property Set CFB has a nonstandard root name",
        ));
    }
    if contents_bytes.len() as u64 > state.limits.max_contents_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set CONTENTS bytes",
            observed: contents_bytes.len() as u64,
            maximum: state.limits.max_contents_bytes,
        });
    }
    if records.iter().any(|record| {
        record.raw_kind != EntryKind::Root.raw()
            && record.raw_kind != EntryKind::Storage.raw()
            && record.raw_kind != EntryKind::Stream.raw()
    }) {
        return Err(invalid(
            "changed non-simple Property Set CFB contains an unsupported directory kind",
        ));
    }
    if records.len() > state.limits.max_directory_entries {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set directory entries",
            observed: records.len() as u64,
            maximum: state.limits.max_directory_entries as u64,
        });
    }
    let mut streams = 0usize;
    let mut total_bytes = 0u64;
    for record in records {
        if record.kind == Some(ElementKind::Stream) {
            streams = streams
                .checked_add(1)
                .ok_or_else(|| invalid("non-simple Property Set stream count overflows"))?;
            let size = payload_size(state, record, edits, contents_bytes)?;
            if size > state.limits.max_stream_bytes {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set stream bytes",
                    observed: size,
                    maximum: state.limits.max_stream_bytes,
                });
            }
            total_bytes = total_bytes
                .checked_add(size)
                .ok_or_else(|| invalid("non-simple Property Set output size overflows"))?;
        }
    }
    if streams > state.limits.max_streams {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set streams",
            observed: streams as u64,
            maximum: state.limits.max_streams as u64,
        });
    }
    if total_bytes > state.limits.max_total_stream_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set aggregate stream bytes",
            observed: total_bytes,
            maximum: state.limits.max_total_stream_bytes,
        });
    }
    // Ask the forward-only CFB planner for the exact physical layout before
    // retaining any edited/source stream payload.  The ordinary writer below
    // is still used for source metadata replay; this bounded zero-reader pass
    // only proves the output-size admission and never copies a payload.
    preflight_render_output(state, records, contents, edits, contents_bytes)?;
    let root = records
        .iter()
        .find(|record| record.is_root())
        .and_then(|record| record.metadata)
        .ok_or_else(|| invalid("non-simple Property Set root metadata is unavailable"))?;
    let mut writer = OleWriter::new();
    writer.set_root_clsid(*contents.class_identifier.as_bytes());
    writer.set_root_state_bits(root.state_bits());
    writer.set_root_modified_time(root.modified_time());
    writer.set_root_creation_time_from_source(root.creation_time());

    let mut storage_indices = Vec::new();
    storage_indices
        .try_reserve(records.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set storage render plan",
            source,
        })?;
    for (index, record) in records.iter().enumerate() {
        if record.kind == Some(ElementKind::Storage) {
            storage_indices.push(index);
        }
    }
    storage_indices.sort_by(|left, right| {
        records[*left]
            .path
            .len()
            .cmp(&records[*right].path.len())
            .then_with(|| records[*left].path.cmp(&records[*right].path))
    });
    for index in storage_indices {
        let record = &records[index];
        let refs = path_refs(
            &record.path,
            "non-simple Property Set render path references",
        )?;
        writer.create_storage(&refs)?;
        if let Some(metadata) = record.metadata
            && let Some(class_id) = metadata.class_id()
        {
            writer.set_storage_clsid(&refs, *class_id.as_bytes())?;
        }
        if let Some(metadata) = record.metadata
            && (metadata.state_bits() != 0
                || metadata.creation_time() != 0
                || metadata.modified_time() != 0)
        {
            writer.set_storage_metadata(
                &refs,
                metadata.state_bits(),
                metadata.creation_time(),
                metadata.modified_time(),
            )?;
        }
    }
    for record in records {
        if record.kind != Some(ElementKind::Stream) {
            continue;
        }
        let refs = path_refs(
            &record.path,
            "non-simple Property Set render path references",
        )?;
        let data = payload(state, record, edits, contents_bytes)?;
        if let Some(metadata) = record.metadata
            && (metadata.state_bits() != 0
                || metadata.creation_time() != 0
                || metadata.modified_time() != 0)
        {
            writer.create_stream_with_metadata(
                &refs,
                &data,
                metadata.state_bits(),
                metadata.creation_time(),
                metadata.modified_time(),
            )?;
        } else {
            writer.create_stream_owned(&refs, data)?;
        }
    }
    let mut output = BoundedOutput::new(state.limits.max_output_bytes);
    let result = writer.write_to(&mut output);
    if output.exceeded {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set output bytes",
            observed: state.limits.max_output_bytes.saturating_add(1),
            maximum: state.limits.max_output_bytes,
        });
    }
    result?;
    Ok(output.into_inner().into())
}

fn preflight_render_output(
    state: &State,
    records: &[EntryRecord],
    contents: &crate::property_set::Stream,
    edits: &[PayloadEdit],
    contents_bytes: &[u8],
) -> Result<(), OleError> {
    let max_path_components = records
        .iter()
        .map(|record| record.path.len())
        .max()
        .unwrap_or(1)
        .max(1);
    let planner_limits = SequentialWriterLimits::new(
        u64::try_from(state.limits.max_streams).unwrap_or(u64::MAX),
        u64::try_from(state.limits.max_directory_entries).unwrap_or(u64::MAX),
        u64::try_from(max_path_components).unwrap_or(u64::MAX),
        u64::MAX,
        state.limits.max_stream_bytes,
        u64::MAX,
        state.limits.max_output_bytes,
    );
    let options = SequentialWriterOptions::default()
        .with_limits(planner_limits)
        .with_publication_buffer_bytes(512);
    let mut planner = SequentialOleWriter::with_options(options).map_err(map_sequential_error)?;
    planner.set_root_clsid(*contents.class_identifier.as_bytes());
    let root = records
        .iter()
        .find(|record| record.is_root())
        .and_then(|record| record.metadata)
        .ok_or_else(|| invalid("non-simple Property Set root metadata is unavailable"))?;
    planner.set_root_state_bits(root.state_bits());
    planner.set_root_modified_time(root.modified_time());

    for record in records {
        if record.kind != Some(ElementKind::Storage) {
            continue;
        }
        let refs = path_refs(
            &record.path,
            "non-simple Property Set output preflight path references",
        )?;
        planner
            .create_storage(&refs)
            .map_err(map_sequential_error)?;
        if let Some(metadata) = record.metadata {
            if let Some(class_id) = metadata.class_id() {
                planner
                    .set_storage_clsid(&refs, *class_id.as_bytes())
                    .map_err(map_sequential_error)?;
            }
            planner
                .set_storage_metadata(
                    &refs,
                    metadata.state_bits(),
                    metadata.creation_time(),
                    metadata.modified_time(),
                )
                .map_err(map_sequential_error)?;
        }
    }

    for record in records {
        if record.kind != Some(ElementKind::Stream) {
            continue;
        }
        let refs = path_refs(
            &record.path,
            "non-simple Property Set output preflight path references",
        )?;
        let size = payload_size(state, record, edits, contents_bytes)?;
        planner
            .add_stream(&refs, size, ZeroReader::new(size))
            .map_err(map_sequential_error)?;
        if let Some(metadata) = record.metadata {
            planner
                .set_stream_metadata(
                    &refs,
                    metadata.state_bits(),
                    metadata.creation_time(),
                    metadata.modified_time(),
                )
                .map_err(map_sequential_error)?;
        }
    }

    let mut probe = ProbeOutput::default();
    match planner.write_to(&mut probe) {
        // The layout planner runs to completion before the first sink call.
        // Accept the first physical segment, then stop at the next write so
        // no payload-sized output or source allocation is retained here.
        Err(SequentialWriteError::WriteZero { .. }) if probe.writes >= 2 => Ok(()),
        Ok(_) => Ok(()),
        Err(error) => Err(map_sequential_error(error)),
    }
}

fn map_sequential_error(error: SequentialWriteError) -> OleError {
    match error {
        SequentialWriteError::LimitExceeded {
            resource: "output bytes",
            observed,
            limit,
        } => OleError::LimitExceeded {
            resource: "non-simple Property Set output bytes",
            observed,
            maximum: limit,
        },
        SequentialWriteError::Planning(error) => error,
        other => OleError::InvalidFormat(format!(
            "non-simple Property Set output preflight failed: {other}"
        )),
    }
}

#[derive(Debug)]
struct ZeroReader {
    remaining: u64,
}

impl ZeroReader {
    fn new(remaining: u64) -> Self {
        Self { remaining }
    }
}

impl io::Read for ZeroReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.remaining == 0 || output.is_empty() {
            return Ok(0);
        }
        let count = usize::try_from(self.remaining)
            .unwrap_or(usize::MAX)
            .min(output.len());
        output[..count].fill(0);
        self.remaining -= u64::try_from(count).unwrap_or(0);
        Ok(count)
    }
}

#[derive(Debug, Default)]
struct ProbeOutput {
    writes: usize,
}

impl Write for ProbeOutput {
    fn write(&mut self, data: &[u8]) -> io::Result<usize> {
        self.writes = self.writes.saturating_add(1);
        if self.writes == 1 {
            Ok(data.len())
        } else {
            Ok(0)
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct BoundedOutput {
    bytes: Vec<u8>,
    position: u64,
    maximum: u64,
    exceeded: bool,
}

impl BoundedOutput {
    fn new(maximum: u64) -> Self {
        Self {
            bytes: Vec::new(),
            position: 0,
            maximum,
            exceeded: false,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }

    fn limit_error(&mut self) -> io::Error {
        self.exceeded = true;
        io::Error::new(
            io::ErrorKind::WriteZero,
            "non-simple Property Set output exceeds its configured limit",
        )
    }
}

impl Write for BoundedOutput {
    fn write(&mut self, data: &[u8]) -> io::Result<usize> {
        let count = u64::try_from(data.len()).map_err(|_| self.limit_error())?;
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| self.limit_error())?;
        if end > self.maximum {
            return Err(self.limit_error());
        }
        let start = usize::try_from(self.position).map_err(|_| self.limit_error())?;
        let end = usize::try_from(end).map_err(|_| self.limit_error())?;
        if end > self.bytes.len() {
            self.bytes
                .try_reserve_exact(end - self.bytes.len())
                .map_err(|error| io::Error::other(error.to_string()))?;
            self.bytes.resize(end, 0);
        }
        self.bytes[start..end].copy_from_slice(data);
        self.position = end as u64;
        Ok(data.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for BoundedOutput {
    fn seek(&mut self, from: SeekFrom) -> io::Result<u64> {
        let current = self.position;
        let end = self.bytes.len() as u64;
        let next = match from {
            SeekFrom::Start(value) => value,
            SeekFrom::Current(value) if value >= 0 => current
                .checked_add(value as u64)
                .ok_or_else(|| self.limit_error())?,
            SeekFrom::Current(value) => {
                current.checked_sub(value.unsigned_abs()).ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "bounded output seek underflow")
                })?
            },
            SeekFrom::End(value) if value >= 0 => end
                .checked_add(value as u64)
                .ok_or_else(|| self.limit_error())?,
            SeekFrom::End(value) => end.checked_sub(value.unsigned_abs()).ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "bounded output seek underflow")
            })?,
        };
        if next > self.maximum {
            return Err(self.limit_error());
        }
        self.position = next;
        Ok(next)
    }
}

fn payload_size(
    _state: &State,
    record: &EntryRecord,
    edits: &[PayloadEdit],
    contents_bytes: &[u8],
) -> Result<u64, OleError> {
    if record.path.len() == 1 && directory_names_equal(&record.path[0], "CONTENTS") {
        return Ok(contents_bytes.len() as u64);
    }
    if let Some(edit) = edits
        .iter()
        .find(|edit| same_path(&edit.path, &record.path))
    {
        return Ok(edit.bytes.as_ref().map_or(0, |bytes| bytes.len() as u64));
    }
    Ok(record.size)
}

fn payload(
    state: &State,
    record: &EntryRecord,
    edits: &[PayloadEdit],
    contents_bytes: &[u8],
) -> Result<Vec<u8>, OleError> {
    if record.path.len() == 1 && directory_names_equal(&record.path[0], "CONTENTS") {
        return Ok(contents_bytes.to_vec());
    }
    if let Some(edit) = edits
        .iter()
        .find(|edit| same_path(&edit.path, &record.path))
    {
        let Some(bytes) = edit.bytes.as_ref() else {
            return Err(invalid("render plan removed a required stream"));
        };
        return Ok(bytes.as_ref().to_vec());
    }
    Ok(read_stream(state, record)?.as_ref().to_vec())
}

fn path_refs<'a>(path: &'a [String], resource: &'static str) -> Result<Vec<&'a str>, OleError> {
    let mut refs = Vec::new();
    refs.try_reserve_exact(path.len())
        .map_err(|source| OleError::Allocation { resource, source })?;
    refs.extend(path.iter().map(String::as_str));
    Ok(refs)
}
