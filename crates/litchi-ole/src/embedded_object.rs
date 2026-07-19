//! Inert, bounded discovery and mutation of legacy Office embedded objects.
//!
//! No API in this module activates an object, executes macros, COM, DDE, or
//! native content, fetches a link, or deserializes an embedded payload.

use crate::{OleError, OleFile, OleWriter};
use std::collections::HashSet;
use std::io::{Cursor, Read, Seek};

const STORAGE: u8 = 1;
const STREAM: u8 = 2;
const COMPOBJ: &str = "\u{1}CompObj";
const OLE10_NATIVE: &str = "\u{1}Ole10Native";
const OBJINFO: &str = "\u{3}ObjInfo";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyOfficeObjectFormat {
    Doc,
    Xls,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyOfficeObjectKind {
    Embedded,
    Linked,
    ActiveXControl,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyOfficePreviewKind {
    OlePresentation,
    PrintWmf,
    PrintEmf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyOfficePreview {
    pub kind: LegacyOfficePreviewKind,
    pub stream_name: String,
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompObjMetadata {
    pub clsid: String,
    pub user_type: Option<String>,
    pub clipboard_format: Option<String>,
    pub prog_id: Option<String>,
    pub unicode_user_type: Option<String>,
    pub unicode_clipboard_format: Option<String>,
    pub unicode_prog_id: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocObjectDescriptor {
    pub default_handler: bool,
    pub linked: bool,
    pub display_as_icon: bool,
    pub ole1: bool,
    pub manual_update: bool,
    pub recompose_on_resize: bool,
    pub activex_control: bool,
    pub stream_control: bool,
    pub view_object: bool,
    pub enhanced_metafile: bool,
    pub queried_enhanced_metafile: bool,
    pub stored_as_enhanced_metafile: bool,
    pub clipboard_format: u16,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleNativePackage {
    pub flags: u16,
    pub label: String,
    pub file_name: String,
    pub command: String,
    /// Opaque bytes only; callers must not execute or deserialize them.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyOfficeEmbeddedObject {
    pub id: String,
    pub storage_path: Vec<String>,
    pub internal_storage_reference: Option<u32>,
    pub kind: LegacyOfficeObjectKind,
    pub clsid: String,
    pub prog_id: Option<String>,
    pub display_name: Option<String>,
    pub clipboard_format: Option<String>,
    /// Stored link text. This is never resolved or fetched.
    pub link_metadata: Option<String>,
    pub comp_obj: Option<CompObjMetadata>,
    pub doc_descriptor: Option<DocObjectDescriptor>,
    pub native_package: Option<OleNativePackage>,
    pub previews: Vec<LegacyOfficePreview>,
    /// Standalone CFB containing the selected storage's children.
    pub compound_file: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyOfficeObjectLimits {
    pub max_objects: usize,
    pub max_storage_depth: usize,
    pub max_streams_per_object: usize,
    pub max_stream_size: u64,
    pub max_object_size: u64,
    pub max_total_size: u64,
    pub max_metadata_string_bytes: usize,
}

impl Default for LegacyOfficeObjectLimits {
    fn default() -> Self {
        Self {
            max_objects: 1_024,
            max_storage_depth: 32,
            max_streams_per_object: 4_096,
            max_stream_size: 128 * 1024 * 1024,
            max_object_size: 256 * 1024 * 1024,
            max_total_size: 512 * 1024 * 1024,
            max_metadata_string_bytes: 64 * 1024,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyOfficeObjectCollection {
    objects: Vec<LegacyOfficeEmbeddedObject>,
}

impl LegacyOfficeObjectCollection {
    pub fn new(objects: Vec<LegacyOfficeEmbeddedObject>) -> Result<Self, OleError> {
        Self::validate_objects(&objects)?;
        Ok(Self { objects })
    }

    pub fn as_slice(&self) -> &[LegacyOfficeEmbeddedObject] {
        &self.objects
    }

    pub fn into_vec(self) -> Vec<LegacyOfficeEmbeddedObject> {
        self.objects
    }

    pub fn find(&self, id: &str) -> Option<&LegacyOfficeEmbeddedObject> {
        self.objects.iter().find(|object| object.id == id)
    }

    pub fn add(&mut self, object: LegacyOfficeEmbeddedObject) -> Result<(), OleError> {
        let mut candidate = self.objects.clone();
        candidate.push(object);
        Self::validate_objects(&candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn update<F>(&mut self, id: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut LegacyOfficeEmbeddedObject) -> Result<(), OleError>,
    {
        let mut candidate = self.objects.clone();
        let object = candidate.iter_mut().find(|value| value.id == id).ok_or_else(|| {
            OleError::InvalidFormat(format!("embedded object {id:?} not found"))
        })?;
        edit(object)?;
        Self::validate_objects(&candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn replace(
        &mut self,
        id: &str,
        replacement: LegacyOfficeEmbeddedObject,
    ) -> Result<LegacyOfficeEmbeddedObject, OleError> {
        let mut candidate = self.objects.clone();
        let index = candidate.iter().position(|value| value.id == id).ok_or_else(|| {
            OleError::InvalidFormat(format!("embedded object {id:?} not found"))
        })?;
        let previous = std::mem::replace(&mut candidate[index], replacement);
        Self::validate_objects(&candidate)?;
        self.objects = candidate;
        Ok(previous)
    }

    pub fn remove(&mut self, id: &str) -> Result<LegacyOfficeEmbeddedObject, OleError> {
        let index = self.objects.iter().position(|value| value.id == id).ok_or_else(|| {
            OleError::InvalidFormat(format!("embedded object {id:?} not found"))
        })?;
        Ok(self.objects.remove(index))
    }

    pub fn reorder(&mut self, ids: &[String]) -> Result<(), OleError> {
        if ids.len() != self.objects.len() {
            return Err(OleError::InvalidFormat(
                "reorder must contain every embedded object exactly once".into(),
            ));
        }
        let mut remaining = self.objects.clone();
        let mut candidate = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining.iter().position(|value| value.id == *id).ok_or_else(|| {
                OleError::InvalidFormat(format!("unknown or repeated object id {id:?}"))
            })?;
            candidate.push(remaining.remove(index));
        }
        Self::validate_objects(&candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn validate(&self) -> Result<(), OleError> {
        Self::validate_objects(&self.objects)
    }

    fn validate_objects(objects: &[LegacyOfficeEmbeddedObject]) -> Result<(), OleError> {
        let mut ids = HashSet::new();
        let mut paths = HashSet::new();
        for object in objects {
            if object.id.is_empty() || object.storage_path.is_empty() {
                return Err(OleError::InvalidFormat(
                    "embedded object id and storage path must not be empty".into(),
                ));
            }
            if !ids.insert(object.id.clone()) || !paths.insert(object.storage_path.clone()) {
                return Err(OleError::InvalidFormat(format!(
                    "duplicate embedded object id or path {:?}", object.id
                )));
            }
        }
        Ok(())
    }
}

pub fn discover_legacy_office_objects<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: LegacyOfficeObjectFormat,
    limits: LegacyOfficeObjectLimits,
) -> Result<LegacyOfficeObjectCollection, OleError> {
    validate_limits(limits)?;
    reject_protected_package(ole)?;
    let paths = find_object_storages(ole, format, limits.max_objects)?;
    let mut total = 0u64;
    let mut objects = Vec::with_capacity(paths.len());
    for path in paths {
        let object = read_object(ole, format, path, limits)?;
        total = total.checked_add(object.compound_file.len() as u64).ok_or_else(|| {
            OleError::InvalidFormat("embedded object total size overflow".into())
        })?;
        if total > limits.max_total_size {
            return Err(OleError::InvalidFormat(format!(
                "embedded object total size {total} exceeds limit {}", limits.max_total_size
            )));
        }
        objects.push(object);
    }
    LegacyOfficeObjectCollection::new(objects)
}

#[derive(Debug, Clone)]
struct CapturedStream {
    path: Vec<String>,
    data: Vec<u8>,
}

/// Targeted atomic CFB rewrite of an already-referenced object storage.
///
/// Add/remove are intentionally collection-only: package-level add/remove is
/// unsafe without rewriting DOC field records or XLS Obj/FtPictFmla records.
#[derive(Debug, Clone)]
pub struct LegacyOfficeObjectEditor {
    format: LegacyOfficeObjectFormat,
    limits: LegacyOfficeObjectLimits,
    sector_size: usize,
    root_clsid: Option<[u8; 16]>,
    storages: Vec<Vec<String>>,
    streams: Vec<CapturedStream>,
    objects: LegacyOfficeObjectCollection,
    changed: bool,
}

impl LegacyOfficeObjectEditor {
    pub fn open(
        bytes: &[u8],
        format: LegacyOfficeObjectFormat,
        limits: LegacyOfficeObjectLimits,
    ) -> Result<Self, OleError> {
        validate_limits(limits)?;
        let mut ole = OleFile::open(Cursor::new(bytes.to_vec()))?;
        reject_protected_package(&mut ole)?;
        let sector_size = ole.sector_size();
        let root_clsid = ole.root_entry().and_then(|entry| parse_clsid_string(&entry.clsid));
        let mut storages = Vec::new();
        let mut streams = Vec::new();
        capture_container(&mut ole, Vec::new(), &mut storages, &mut streams, limits)?;
        let objects = discover_legacy_office_objects(&mut ole, format, limits)?;
        Ok(Self {
            format,
            limits,
            sector_size,
            root_clsid,
            storages,
            streams,
            objects,
            changed: false,
        })
    }

    pub fn objects(&self) -> &LegacyOfficeObjectCollection {
        &self.objects
    }

    pub fn is_changed(&self) -> bool {
        self.changed
    }

    pub fn package_stream(&self, path: &[String]) -> Option<&[u8]> {
        self.streams
            .iter()
            .find(|stream| stream.path == path)
            .map(|stream| stream.data.as_slice())
    }

    pub fn update<F>(&mut self, id: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut Vec<u8>) -> Result<(), OleError>,
    {
        let mut bytes = self.objects.find(id).ok_or_else(|| {
            OleError::InvalidFormat(format!("embedded object {id:?} not found"))
        })?.compound_file.clone();
        edit(&mut bytes)?;
        self.replace(id, bytes)
    }

    pub fn replace(&mut self, id: &str, compound_file: Vec<u8>) -> Result<(), OleError> {
        if compound_file.len() as u64 > self.limits.max_object_size {
            return Err(OleError::InvalidFormat("replacement object exceeds size limit".into()));
        }
        let object_path = self.objects.find(id).ok_or_else(|| {
            OleError::InvalidFormat(format!("embedded object {id:?} not found"))
        })?.storage_path.clone();
        let mut replacement = OleFile::open(Cursor::new(compound_file))?;
        reject_protected_package(&mut replacement)?;
        let mut new_storages = Vec::new();
        let mut new_streams = Vec::new();
        capture_container(
            &mut replacement,
            Vec::new(),
            &mut new_storages,
            &mut new_streams,
            self.limits,
        )?;

        let mut candidate = self.clone();
        candidate.storages.retain(|path| {
            path == &object_path || !(path.len() > object_path.len() && path.starts_with(&object_path))
        });
        candidate.streams.retain(|stream| !stream.path.starts_with(&object_path));
        for relative in new_storages {
            let path = join(&object_path, &relative);
            if path != object_path && !candidate.storages.contains(&path) {
                candidate.storages.push(path);
            }
        }
        for stream in new_streams {
            candidate.streams.push(CapturedStream {
                path: join(&object_path, &stream.path),
                data: stream.data,
            });
        }
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover_legacy_office_objects(
            &mut check,
            candidate.format,
            candidate.limits,
        )?;
        if candidate.objects.find(id).is_none() {
            return Err(OleError::InvalidFormat(
                "replacement invalidated the object storage reference".into(),
            ));
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Replaces an existing package stream while preserving all other streams.
    /// Format-specific editors use this only after validating and rebuilding
    /// references and offsets in the stream.
    pub fn replace_package_stream(
        &mut self,
        path: &[String],
        data: Vec<u8>,
    ) -> Result<(), OleError> {
        if data.len() as u64 > self.limits.max_stream_size {
            return Err(OleError::InvalidFormat("replacement stream exceeds size limit".into()));
        }
        let mut candidate = self.clone();
        let stream = candidate.streams.iter_mut().find(|stream| stream.path == path)
            .ok_or(OleError::StreamNotFound)?;
        stream.data = data;
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover_legacy_office_objects(&mut check, candidate.format, candidate.limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    pub fn add_package_stream(&mut self, path: Vec<String>, data: Vec<u8>) -> Result<(), OleError> {
        if path.is_empty() || data.len() as u64 > self.limits.max_stream_size {
            return Err(OleError::InvalidFormat("new package stream path or size is invalid".into()));
        }
        if self.streams.iter().any(|stream| stream.path == path) {
            return Err(OleError::InvalidFormat(format!("package stream {path:?} already exists")));
        }
        if path.len() > 1 && !self.storages.iter().any(|storage| storage == &path[..path.len() - 1]) {
            return Err(OleError::InvalidFormat("new package stream parent storage is missing".into()));
        }
        let mut candidate = self.clone();
        candidate.streams.push(CapturedStream { path, data });
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover_legacy_office_objects(&mut check, candidate.format, candidate.limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Adds a storage whose format-specific reference has already been staged.
    pub fn add_referenced_storage(
        &mut self,
        id: &str,
        compound_file: Vec<u8>,
    ) -> Result<(), OleError> {
        let path = match self.format {
            LegacyOfficeObjectFormat::Doc if is_doc_name(id) => {
                vec!["ObjectPool".to_string(), id.to_string()]
            }
            LegacyOfficeObjectFormat::Xls if is_xls_name(id) => vec![id.to_string()],
            _ => return Err(OleError::InvalidFormat("invalid format-specific object storage name".into())),
        };
        if self.storages.iter().any(|value| value == &path) {
            return Err(OleError::InvalidFormat(format!("object storage {id:?} already exists")));
        }
        let mut nested = OleFile::open(Cursor::new(compound_file))?;
        reject_protected_package(&nested)?;
        let mut storages = Vec::new();
        let mut streams = Vec::new();
        capture_container(&mut nested, Vec::new(), &mut storages, &mut streams, self.limits)?;
        let mut candidate = self.clone();
        candidate.storages.push(path.clone());
        for storage in storages { candidate.storages.push(join(&path, &storage)); }
        for stream in streams {
            candidate.streams.push(CapturedStream { path: join(&path, &stream.path), data: stream.data });
        }
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover_legacy_office_objects(&mut check, candidate.format, candidate.limits)?;
        if candidate.objects.find(id).is_none() {
            return Err(OleError::InvalidFormat("new object storage was not discoverable".into()));
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Removes a storage after its format-specific references have been removed.
    pub fn remove_referenced_storage(&mut self, id: &str) -> Result<Vec<u8>, OleError> {
        let object = self.objects.find(id).ok_or_else(|| {
            OleError::InvalidFormat(format!("embedded object {id:?} not found"))
        })?;
        let removed = object.compound_file.clone();
        let path = object.storage_path.clone();
        let mut candidate = self.clone();
        candidate.storages.retain(|value| !value.starts_with(&path));
        candidate.streams.retain(|value| !value.path.starts_with(&path));
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover_legacy_office_objects(&mut check, candidate.format, candidate.limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(removed)
    }

    pub fn finish(&self) -> Result<Vec<u8>, OleError> {
        self.render()
    }

    fn render(&self) -> Result<Vec<u8>, OleError> {
        let mut writer = OleWriter::with_sector_size(self.sector_size);
        if let Some(clsid) = self.root_clsid {
            writer.set_root_clsid(clsid);
        }
        let mut storages = self.storages.clone();
        storages.sort_by_key(Vec::len);
        for path in &storages {
            let refs = path_refs(path);
            writer.create_storage(&refs)?;
        }
        for stream in &self.streams {
            let refs = path_refs(&stream.path);
            writer.create_stream(&refs, &stream.data)?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }
}

fn validate_limits(limits: LegacyOfficeObjectLimits) -> Result<(), OleError> {
    if limits.max_objects == 0
        || limits.max_storage_depth == 0
        || limits.max_streams_per_object == 0
        || limits.max_stream_size == 0
        || limits.max_object_size == 0
        || limits.max_total_size == 0
        || limits.max_metadata_string_bytes == 0
    {
        return Err(OleError::InvalidFormat("all object limits must be non-zero".into()));
    }
    Ok(())
}

fn find_object_storages<R: Read + Seek>(
    ole: &OleFile<R>,
    format: LegacyOfficeObjectFormat,
    max: usize,
) -> Result<Vec<Vec<String>>, OleError> {
    let mut paths = Vec::new();
    match format {
        LegacyOfficeObjectFormat::Doc if ole.exists(&["ObjectPool"]) => {
            for entry in ole.list_directory_entries(&["ObjectPool"])? {
                if entry.entry_type == STORAGE && is_doc_name(&entry.name) {
                    paths.push(vec!["ObjectPool".into(), entry.name.clone()]);
                }
            }
        }
        LegacyOfficeObjectFormat::Doc => {}
        LegacyOfficeObjectFormat::Xls => {
            for entry in ole.list_directory_entries(&[])? {
                if entry.entry_type == STORAGE && is_xls_name(&entry.name) {
                    paths.push(vec![entry.name.clone()]);
                }
            }
        }
    }
    if paths.len() > max {
        return Err(OleError::InvalidFormat(format!(
            "embedded object count {} exceeds limit {max}", paths.len()
        )));
    }
    paths.sort();
    Ok(paths)
}

fn read_object<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: LegacyOfficeObjectFormat,
    path: Vec<String>,
    limits: LegacyOfficeObjectLimits,
) -> Result<LegacyOfficeEmbeddedObject, OleError> {
    let parent = &path[..path.len() - 1];
    let parent_refs = path_refs(parent);
    let id = path.last().cloned().unwrap_or_default();
    let storage_clsid = ole.list_directory_entries(&parent_refs)?.into_iter()
        .find(|entry| entry.name == id && entry.entry_type == STORAGE)
        .map(|entry| entry.clsid.clone())
        .ok_or_else(|| OleError::InvalidFormat(format!("object storage {path:?} not found")))?;
    let mut storages = Vec::new();
    let mut streams = Vec::new();
    capture_subtree(ole, &path, Vec::new(), &mut storages, &mut streams, limits)?;

    let comp_obj = stream_named(&streams, COMPOBJ)
        .map(|bytes| parse_comp_obj(bytes, limits.max_metadata_string_bytes)).transpose()?;
    let doc_descriptor = if format == LegacyOfficeObjectFormat::Doc {
        stream_named(&streams, OBJINFO).map(parse_doc_descriptor).transpose()?
    } else { None };
    let native_package = stream_named(&streams, OLE10_NATIVE)
        .map(|bytes| parse_native_package(bytes, limits)).transpose()?;
    let mut previews = Vec::new();
    for stream in &streams {
        let Some(name) = stream.path.last() else { continue };
        let kind = if name.starts_with("\u{2}OlePres") {
            Some(LegacyOfficePreviewKind::OlePresentation)
        } else if name == "\u{3}PRINT" {
            Some(LegacyOfficePreviewKind::PrintWmf)
        } else if name == "\u{3}EPRINT" {
            Some(LegacyOfficePreviewKind::PrintEmf)
        } else { None };
        if let Some(kind) = kind {
            previews.push(LegacyOfficePreview { kind, stream_name: name.clone(), data: stream.data.clone() });
        }
    }

    let mut writer = OleWriter::new();
    if let Some(clsid) = parse_clsid_string(&storage_clsid) { writer.set_root_clsid(clsid); }
    storages.sort_by_key(Vec::len);
    for storage in &storages { writer.create_storage(&path_refs(storage))?; }
    for stream in &streams { writer.create_stream(&path_refs(&stream.path), &stream.data)?; }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let compound_file = output.into_inner();
    if compound_file.len() as u64 > limits.max_object_size {
        return Err(OleError::InvalidFormat(format!("object {id:?} exceeds size limit")));
    }

    let internal_storage_reference = match format {
        LegacyOfficeObjectFormat::Doc => id.strip_prefix('_').and_then(|v| v.parse().ok()),
        LegacyOfficeObjectFormat::Xls => id.get(3..).and_then(|v| u32::from_str_radix(v, 16).ok()),
    };
    let kind = if format == LegacyOfficeObjectFormat::Xls && id.starts_with("LNK") {
        LegacyOfficeObjectKind::Linked
    } else if doc_descriptor.map(|v| v.activex_control).unwrap_or(false) {
        LegacyOfficeObjectKind::ActiveXControl
    } else if doc_descriptor.map(|v| v.linked).unwrap_or(false) {
        LegacyOfficeObjectKind::Linked
    } else { LegacyOfficeObjectKind::Embedded };
    let clsid = comp_obj.as_ref().map(|v| v.clsid.clone())
        .filter(|v| v != "00000000-0000-0000-0000-000000000000")
        .unwrap_or(storage_clsid);
    let prog_id = comp_obj.as_ref().and_then(|v| v.unicode_prog_id.clone().or_else(|| v.prog_id.clone()));
    let display_name = comp_obj.as_ref().and_then(|v| v.unicode_user_type.clone().or_else(|| v.user_type.clone()));
    let clipboard_format = comp_obj.as_ref().and_then(|v| v.unicode_clipboard_format.clone().or_else(|| v.clipboard_format.clone()));
    let link_metadata = if kind == LegacyOfficeObjectKind::Linked {
        native_package.as_ref().map(|v| v.command.clone()).filter(|v| !v.is_empty()).or_else(|| display_name.clone())
    } else { None };
    Ok(LegacyOfficeEmbeddedObject {
        id, storage_path: path, internal_storage_reference, kind, clsid, prog_id,
        display_name, clipboard_format, link_metadata, comp_obj, doc_descriptor,
        native_package, previews, compound_file,
    })
}

fn capture_container<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: Vec<String>,
    storages: &mut Vec<Vec<String>>,
    streams: &mut Vec<CapturedStream>,
    limits: LegacyOfficeObjectLimits,
) -> Result<(), OleError> {
    if path.len() > limits.max_storage_depth {
        return Err(OleError::InvalidFormat("CFB storage nesting limit exceeded".into()));
    }
    let entries = ole.list_directory_entries(&path_refs(&path))?.into_iter().cloned().collect::<Vec<_>>();
    for entry in entries {
        let mut child = path.clone();
        child.push(entry.name);
        if entry.entry_type == STORAGE {
            storages.push(child.clone());
            capture_container(ole, child, storages, streams, limits)?;
        } else if entry.entry_type == STREAM {
            if entry.size > limits.max_stream_size {
                return Err(OleError::InvalidFormat(format!("stream {child:?} exceeds size limit")));
            }
            let aggregate_limit = limits.max_streams_per_object.checked_mul(limits.max_objects)
                .unwrap_or(usize::MAX);
            if streams.len() >= aggregate_limit {
                return Err(OleError::InvalidFormat("CFB stream count limit exceeded".into()));
            }
            let data = ole.open_stream(&path_refs(&child))?;
            streams.push(CapturedStream { path: child, data });
        }
    }
    Ok(())
}

fn capture_subtree<R: Read + Seek>(
    ole: &mut OleFile<R>,
    absolute: &[String],
    relative: Vec<String>,
    storages: &mut Vec<Vec<String>>,
    streams: &mut Vec<CapturedStream>,
    limits: LegacyOfficeObjectLimits,
) -> Result<(), OleError> {
    if relative.len() > limits.max_storage_depth {
        return Err(OleError::InvalidFormat("object storage nesting limit exceeded".into()));
    }
    let current = join(absolute, &relative);
    let entries = ole.list_directory_entries(&path_refs(&current))?.into_iter().cloned().collect::<Vec<_>>();
    for entry in entries {
        let mut child = relative.clone();
        child.push(entry.name);
        if entry.entry_type == STORAGE {
            storages.push(child.clone());
            capture_subtree(ole, absolute, child, storages, streams, limits)?;
        } else if entry.entry_type == STREAM {
            if streams.len() >= limits.max_streams_per_object || entry.size > limits.max_stream_size {
                return Err(OleError::InvalidFormat("object stream resource limit exceeded".into()));
            }
            let full = join(absolute, &child);
            let data = ole.open_stream(&path_refs(&full))?;
            streams.push(CapturedStream { path: child, data });
        }
    }
    Ok(())
}

fn reject_protected_package<R: Read + Seek>(ole: &OleFile<R>) -> Result<(), OleError> {
    for path in ole.list_streams() {
        if path.iter().any(|name| {
            matches!(name.to_ascii_lowercase().as_str(),
                "_xmlsignatures" | "_signatures" | "\u{5}digitalsignature" |
                "\u{6}dataspaces" | "encryptioninfo" | "encryptedpackage" | "\u{9}drmcontent")
        }) {
            return Err(OleError::InvalidFormat(
                "signed, encrypted, or DRM packages are not eligible for object editing".into(),
            ));
        }
    }
    Ok(())
}

fn parse_comp_obj(data: &[u8], max: usize) -> Result<CompObjMetadata, OleError> {
    if data.len() < 28 { return Err(OleError::InvalidFormat("CompObj stream is truncated".into())); }
    let clsid = format_clsid(&data[12..28]);
    let mut offset = 28;
    let user_type = read_ansi(data, &mut offset, max)?;
    let clipboard_format = read_ansi(data, &mut offset, max)?;
    let prog_id = read_ansi(data, &mut offset, max)?;
    let (mut unicode_user_type, mut unicode_clipboard_format, mut unicode_prog_id) = (None, None, None);
    if data.len().saturating_sub(offset) >= 4 && read_u32(data, offset)? == 0x71B2_39F4 {
        offset += 4;
        unicode_user_type = read_utf16(data, &mut offset, max)?;
        unicode_clipboard_format = read_utf16(data, &mut offset, max)?;
        unicode_prog_id = read_utf16(data, &mut offset, max)?;
    }
    Ok(CompObjMetadata { clsid, user_type, clipboard_format, prog_id,
        unicode_user_type, unicode_clipboard_format, unicode_prog_id })
}

fn parse_doc_descriptor(data: &[u8]) -> Result<DocObjectDescriptor, OleError> {
    if data.len() != 4 && data.len() != 6 {
        return Err(OleError::InvalidFormat("DOC ObjInfo ODT must be 4 or 6 bytes".into()));
    }
    let first = u16::from_le_bytes([data[0], data[1]]);
    if first & ((1 << 10) | (1 << 11)) != 0 {
        return Err(OleError::InvalidFormat("DOC ObjInfo reserved bits are set".into()));
    }
    let activex_control = first & (1 << 12) != 0;
    let stream_control = first & (1 << 13) != 0;
    if stream_control && !activex_control {
        return Err(OleError::InvalidFormat("DOC stream control requires ActiveX".into()));
    }
    let second = if data.len() == 6 { u16::from_le_bytes([data[4], data[5]]) } else { 0 };
    if second & 2 != 0 { return Err(OleError::InvalidFormat("DOC ObjInfo reserved bit is set".into())); }
    Ok(DocObjectDescriptor {
        default_handler: first & (1 << 1) != 0, linked: first & (1 << 4) != 0,
        display_as_icon: first & (1 << 6) != 0, ole1: first & (1 << 7) != 0,
        manual_update: first & (1 << 8) != 0, recompose_on_resize: first & (1 << 9) != 0,
        activex_control, stream_control, view_object: first & (1 << 15) != 0,
        enhanced_metafile: second & 1 != 0, queried_enhanced_metafile: second & 4 != 0,
        stored_as_enhanced_metafile: second & 8 != 0,
        clipboard_format: u16::from_le_bytes([data[2], data[3]]),
    })
}

fn parse_native_package(data: &[u8], limits: LegacyOfficeObjectLimits) -> Result<OleNativePackage, OleError> {
    if data.len() < 6 { return Err(OleError::InvalidFormat("Ole10Native stream is truncated".into())); }
    let declared = read_u32(data, 0)? as usize;
    let end = 4usize.checked_add(declared).ok_or_else(|| OleError::InvalidFormat("Ole10Native size overflow".into()))?;
    if end > data.len() { return Err(OleError::InvalidFormat("Ole10Native size exceeds stream".into())); }
    let bytes = &data[..end];
    let flags = u16::from_le_bytes([bytes[4], bytes[5]]);
    let mut offset = 6;
    let label = read_c_string(bytes, &mut offset, limits.max_metadata_string_bytes)?;
    let file_name = read_c_string(bytes, &mut offset, limits.max_metadata_string_bytes)?;
    if bytes.len().saturating_sub(offset) < 4 { return Err(OleError::InvalidFormat("Ole10Native flags truncated".into())); }
    offset += 4;
    let command = read_c_string(bytes, &mut offset, limits.max_metadata_string_bytes)?;
    let native_size = read_u32(bytes, offset)? as usize;
    offset += 4;
    let native_end = offset.checked_add(native_size).ok_or_else(|| OleError::InvalidFormat("Ole10Native payload overflow".into()))?;
    if native_end > bytes.len() || native_size as u64 > limits.max_object_size {
        return Err(OleError::InvalidFormat("Ole10Native payload exceeds limits".into()));
    }
    Ok(OleNativePackage { flags, label, file_name, command, data: bytes[offset..native_end].to_vec() })
}

fn read_ansi(data: &[u8], offset: &mut usize, max: usize) -> Result<Option<String>, OleError> {
    let length = read_u32(data, *offset)? as usize;
    *offset += 4;
    if length == 0 || length == u32::MAX as usize { return Ok(None); }
    if length > max || data.len().saturating_sub(*offset) < length {
        return Err(OleError::InvalidFormat("CompObj ANSI string exceeds bounds".into()));
    }
    let bytes = &data[*offset..*offset + length];
    *offset += length;
    Ok(Some(String::from_utf8_lossy(bytes.strip_suffix(&[0]).unwrap_or(bytes)).into_owned()))
}

fn read_utf16(data: &[u8], offset: &mut usize, max: usize) -> Result<Option<String>, OleError> {
    let units = read_u32(data, *offset)? as usize;
    *offset += 4;
    if units == 0 || units == u32::MAX as usize { return Ok(None); }
    let byte_len = units.checked_mul(2).ok_or_else(|| OleError::InvalidFormat("CompObj Unicode size overflow".into()))?;
    if byte_len > max || data.len().saturating_sub(*offset) < byte_len {
        return Err(OleError::InvalidFormat("CompObj Unicode string exceeds bounds".into()));
    }
    let mut units = data[*offset..*offset + byte_len].chunks_exact(2)
        .map(|v| u16::from_le_bytes([v[0], v[1]])).collect::<Vec<_>>();
    *offset += byte_len;
    if units.last() == Some(&0) { units.pop(); }
    Ok(Some(String::from_utf16_lossy(&units)))
}

fn read_c_string(data: &[u8], offset: &mut usize, max: usize) -> Result<String, OleError> {
    let remaining = data.get(*offset..).ok_or_else(|| OleError::InvalidFormat("invalid string offset".into()))?;
    let end = remaining.iter().position(|byte| *byte == 0)
        .ok_or_else(|| OleError::InvalidFormat("unterminated native-package string".into()))?;
    if end > max { return Err(OleError::InvalidFormat("native-package string exceeds limit".into())); }
    *offset += end + 1;
    Ok(String::from_utf8_lossy(&remaining[..end]).into_owned())
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, OleError> {
    let value = data.get(offset..offset + 4)
        .ok_or_else(|| OleError::InvalidFormat("object metadata is truncated".into()))?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn stream_named<'a>(streams: &'a [CapturedStream], name: &str) -> Option<&'a [u8]> {
    streams.iter().find(|stream| stream.path.len() == 1 && stream.path[0] == name)
        .map(|stream| stream.data.as_slice())
}

fn is_doc_name(name: &str) -> bool {
    name.strip_prefix('_').map(|v| !v.is_empty() && v.bytes().all(|b| b.is_ascii_digit())).unwrap_or(false)
}

fn is_xls_name(name: &str) -> bool {
    (name.starts_with("MBD") || name.starts_with("LNK")) && name.len() == 11
        && name.as_bytes()[3..].iter().all(|b| b.is_ascii_hexdigit())
}

fn path_refs(path: &[String]) -> Vec<&str> { path.iter().map(String::as_str).collect() }
fn join(left: &[String], right: &[String]) -> Vec<String> { left.iter().chain(right).cloned().collect() }

fn format_clsid(bytes: &[u8]) -> String {
    format!("{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
        bytes[3], bytes[2], bytes[1], bytes[0], bytes[5], bytes[4], bytes[7], bytes[6],
        bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15])
}

fn parse_clsid_string(value: &str) -> Option<[u8; 16]> {
    let value = value.trim_matches(|c| c == '{' || c == '}').replace('-', "");
    if value.len() != 32 { return None; }
    let mut canonical = [0u8; 16];
    for (index, byte) in canonical.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16).ok()?;
    }
    Some([canonical[3], canonical[2], canonical[1], canonical[0], canonical[5], canonical[4],
        canonical[7], canonical[6], canonical[8], canonical[9], canonical[10], canonical[11],
        canonical[12], canonical[13], canonical[14], canonical[15]])
}
