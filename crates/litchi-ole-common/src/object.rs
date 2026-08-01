//! Inert, bounded discovery and mutation of legacy Office embedded objects.
//!
//! No API in this module activates an object, executes macros, COM, DDE, or
//! native content, fetches a link, or deserializes an embedded payload.

use litchi_cfb::{OleError, OleFile, OleWriter};
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek};
use std::sync::Arc;

const STORAGE: u8 = 1;
const STREAM: u8 = 2;
const COMPOBJ: &str = "\u{1}CompObj";
const OLE10_NATIVE: &str = "\u{1}Ole10Native";
const OBJINFO: &str = "\u{3}ObjInfo";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Format {
    Doc,
    Xls,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Embedded,
    Linked,
    ActiveXControl,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreviewKind {
    OlePresentation,
    PrintWmf,
    PrintEmf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Preview {
    pub kind: PreviewKind,
    pub stream: String,
    pub data: Arc<[u8]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompObj {
    pub clsid: String,
    pub user_type: Option<String>,
    pub clipboard_format: Option<String>,
    pub prog_id: Option<String>,
    pub unicode_user_type: Option<String>,
    pub unicode_clipboard_format: Option<String>,
    pub unicode_prog_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Native {
    pub flags: u16,
    pub label: String,
    pub file_name: String,
    pub command: String,
    /// Opaque bytes only; callers must not execute or deserialize them.
    pub data: Arc<[u8]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Object {
    pub id: String,
    pub path: Vec<String>,
    pub storage_ref: Option<u32>,
    pub kind: Kind,
    pub clsid: String,
    pub prog_id: Option<String>,
    pub display_name: Option<String>,
    pub clipboard_format: Option<String>,
    /// Stored link text. This is never resolved or fetched.
    pub link: Option<String>,
    pub metadata: Option<CompObj>,
    /// Format-owned metadata bytes, such as DOC's `\u{3}ObjInfo` stream.
    /// The host crate is responsible for interpreting them.
    pub host: Option<Arc<[u8]>>,
    pub native: Option<Native>,
    pub previews: Vec<Preview>,
    /// Standalone CFB containing the selected storage's children.
    pub compound: Arc<[u8]>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_objects: usize,
    pub max_storage_depth: usize,
    pub max_streams_per_object: usize,
    pub max_stream_size: u64,
    pub max_object_size: u64,
    pub max_total_size: u64,
    pub max_metadata_string_bytes: usize,
}

impl Default for Limits {
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

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Objects {
    objects: Vec<Object>,
}

impl Objects {
    pub fn new(objects: Vec<Object>) -> Result<Self, OleError> {
        Self::validate_objects(&objects)?;
        Ok(Self { objects })
    }

    pub fn as_slice(&self) -> &[Object] {
        &self.objects
    }

    pub fn len(&self) -> usize {
        self.objects.len()
    }

    pub fn is_empty(&self) -> bool {
        self.objects.is_empty()
    }

    pub fn iter(&self) -> std::slice::Iter<'_, Object> {
        self.objects.iter()
    }

    pub fn into_vec(self) -> Vec<Object> {
        self.objects
    }

    /// Looks up an object by its semantic storage identifier.
    pub fn get(&self, id: &str) -> Option<&Object> {
        self.objects.iter().find(|object| object.id == id)
    }

    /// Looks up an object by raw discovery order without panicking.
    pub fn at(&self, index: usize) -> Option<&Object> {
        self.objects.get(index)
    }

    pub fn add(&mut self, object: Object) -> Result<(), OleError> {
        Self::validate_one(&self.objects, &object, None)?;
        self.objects.push(object);
        Ok(())
    }

    pub fn update<F>(&mut self, id: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut Object) -> Result<(), OleError>,
    {
        let index = self
            .objects
            .iter()
            .position(|value| value.id == id)
            .ok_or_else(|| OleError::InvalidFormat(format!("embedded object {id:?} not found")))?;
        let mut candidate = self.objects[index].clone();
        edit(&mut candidate)?;
        Self::validate_one(&self.objects, &candidate, Some(index))?;
        self.objects[index] = candidate;
        Ok(())
    }

    pub fn replace(&mut self, id: &str, replacement: Object) -> Result<Object, OleError> {
        let index = self
            .objects
            .iter()
            .position(|value| value.id == id)
            .ok_or_else(|| OleError::InvalidFormat(format!("embedded object {id:?} not found")))?;
        Self::validate_one(&self.objects, &replacement, Some(index))?;
        Ok(std::mem::replace(&mut self.objects[index], replacement))
    }

    pub fn remove(&mut self, id: &str) -> Result<Object, OleError> {
        let index = self
            .objects
            .iter()
            .position(|value| value.id == id)
            .ok_or_else(|| OleError::InvalidFormat(format!("embedded object {id:?} not found")))?;
        Ok(self.objects.remove(index))
    }

    pub fn reorder(&mut self, ids: &[String]) -> Result<(), OleError> {
        if ids.len() != self.objects.len() {
            return Err(OleError::InvalidFormat(
                "reorder must contain every embedded object exactly once".into(),
            ));
        }
        let known = self
            .objects
            .iter()
            .map(|object| object.id.as_str())
            .collect::<HashSet<_>>();
        let requested = ids.iter().map(String::as_str).collect::<HashSet<_>>();
        if requested.len() != ids.len() || requested != known {
            return Err(OleError::InvalidFormat(
                "reorder contains an unknown or repeated object id".into(),
            ));
        }
        let mut remaining = std::mem::take(&mut self.objects)
            .into_iter()
            .map(|object| (object.id.clone(), object))
            .collect::<HashMap<_, _>>();
        let mut candidate = Vec::with_capacity(ids.len());
        for id in ids {
            if let Some(object) = remaining.remove(id) {
                candidate.push(object);
            }
        }
        self.objects = candidate;
        Ok(())
    }

    pub fn validate(&self) -> Result<(), OleError> {
        Self::validate_objects(&self.objects)
    }

    fn validate_objects(objects: &[Object]) -> Result<(), OleError> {
        let mut ids = HashSet::new();
        let mut paths = HashSet::new();
        for object in objects {
            if object.id.is_empty() || object.path.is_empty() {
                return Err(OleError::InvalidFormat(
                    "embedded object id and storage path must not be empty".into(),
                ));
            }
            if !ids.insert(object.id.as_str()) || !paths.insert(object.path.as_slice()) {
                return Err(OleError::InvalidFormat(format!(
                    "duplicate embedded object id or path {:?}",
                    object.id
                )));
            }
        }
        Ok(())
    }

    fn validate_one(
        objects: &[Object],
        candidate: &Object,
        skip: Option<usize>,
    ) -> Result<(), OleError> {
        if candidate.id.is_empty() || candidate.path.is_empty() {
            return Err(OleError::InvalidFormat(
                "embedded object id and storage path must not be empty".into(),
            ));
        }
        if objects.iter().enumerate().any(|(index, object)| {
            Some(index) != skip && (object.id == candidate.id || object.path == candidate.path)
        }) {
            return Err(OleError::InvalidFormat(format!(
                "duplicate embedded object id or path {:?}",
                candidate.id
            )));
        }
        Ok(())
    }
}

impl AsRef<[Object]> for Objects {
    fn as_ref(&self) -> &[Object] {
        self.as_slice()
    }
}

impl<'a> IntoIterator for &'a Objects {
    type Item = &'a Object;
    type IntoIter = std::slice::Iter<'a, Object>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

pub fn discover<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: Format,
    limits: Limits,
) -> Result<Objects, OleError> {
    validate_limits(limits)?;
    reject_protected_package(ole)?;
    let paths = find_object_storages(ole, format, limits.max_objects)?;
    let mut total = 0u64;
    let mut objects = Vec::with_capacity(paths.len());
    for path in paths {
        let object = read_object(ole, format, path, limits)?;
        total = total
            .checked_add(object.compound.len() as u64)
            .ok_or_else(|| OleError::InvalidFormat("embedded object total size overflow".into()))?;
        if total > limits.max_total_size {
            return Err(OleError::InvalidFormat(format!(
                "embedded object total size {total} exceeds limit {}",
                limits.max_total_size
            )));
        }
        objects.push(object);
    }
    Objects::new(objects)
}

#[derive(Debug, Clone)]
struct CapturedStream {
    path: Vec<String>,
    data: Arc<[u8]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CapturedStorage {
    path: Vec<String>,
    clsid: Option<[u8; 16]>,
}

/// Targeted atomic CFB rewrite of an already-referenced object storage.
///
/// Add/remove are intentionally collection-only: package-level add/remove is
/// unsafe without rewriting DOC field records or XLS Obj/FtPictFmla records.
#[derive(Debug, Clone)]
pub struct Editor {
    format: Format,
    limits: Limits,
    original: Arc<Vec<u8>>,
    sector_size: usize,
    root_clsid: Option<[u8; 16]>,
    storages: Vec<CapturedStorage>,
    streams: Vec<CapturedStream>,
    objects: Objects,
    changed: bool,
}

impl Editor {
    pub fn open(bytes: Vec<u8>, format: Format, limits: Limits) -> Result<Self, OleError> {
        validate_limits(limits)?;
        let original = Arc::new(bytes);
        let mut ole = OleFile::open(Cursor::new(original.as_slice()))?;
        reject_protected_package(&ole)?;
        let sector_size = ole.sector_size();
        let root_clsid = ole
            .root_entry()
            .and_then(|entry| parse_clsid_string(&entry.clsid));
        let mut storages = Vec::new();
        let mut streams = Vec::new();
        capture_container(&mut ole, Vec::new(), &mut storages, &mut streams, limits)?;
        let objects = discover(&mut ole, format, limits)?;
        Ok(Self {
            format,
            limits,
            original,
            sector_size,
            root_clsid,
            storages,
            streams,
            objects,
            changed: false,
        })
    }

    pub fn objects(&self) -> &Objects {
        &self.objects
    }

    pub fn is_changed(&self) -> bool {
        self.changed
    }

    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.streams
            .iter()
            .find(|stream| stream.path == path)
            .map(|stream| stream.data.as_ref())
    }

    pub fn update<F>(&mut self, id: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut Vec<u8>) -> Result<(), OleError>,
    {
        let mut bytes = self
            .objects
            .get(id)
            .ok_or_else(|| OleError::InvalidFormat(format!("embedded object {id:?} not found")))?
            .compound
            .to_vec();
        edit(&mut bytes)?;
        self.replace(id, bytes)
    }

    pub fn replace(&mut self, id: &str, compound_file: Vec<u8>) -> Result<(), OleError> {
        if compound_file.len() as u64 > self.limits.max_object_size {
            return Err(OleError::InvalidFormat(
                "replacement object exceeds size limit".into(),
            ));
        }
        let object_path = self
            .objects
            .get(id)
            .ok_or_else(|| OleError::InvalidFormat(format!("embedded object {id:?} not found")))?
            .path
            .clone();
        let mut replacement = OleFile::open(Cursor::new(compound_file))?;
        reject_protected_package(&replacement)?;
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
        let replacement_clsid = replacement
            .root_entry()
            .and_then(|entry| parse_clsid_string(&entry.clsid));
        candidate.storages.retain(|storage| {
            storage.path == object_path
                || !(storage.path.len() > object_path.len()
                    && storage.path.starts_with(&object_path))
        });
        if let Some(storage) = candidate
            .storages
            .iter_mut()
            .find(|storage| storage.path == object_path)
        {
            storage.clsid = replacement_clsid;
        }
        candidate
            .streams
            .retain(|stream| !stream.path.starts_with(&object_path));
        for relative in new_storages {
            let path = join(&object_path, &relative.path);
            if path != object_path
                && !candidate
                    .storages
                    .iter()
                    .any(|storage| storage.path == path)
            {
                candidate.storages.push(CapturedStorage {
                    path,
                    clsid: relative.clsid,
                });
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
        candidate.objects = discover(&mut check, candidate.format, candidate.limits)?;
        if candidate.objects.get(id).is_none() {
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
    pub fn put_stream(&mut self, path: &[String], data: Vec<u8>) -> Result<(), OleError> {
        if data.len() as u64 > self.limits.max_stream_size {
            return Err(OleError::InvalidFormat(
                "replacement stream exceeds size limit".into(),
            ));
        }
        let mut candidate = self.clone();
        let stream = candidate
            .streams
            .iter_mut()
            .find(|stream| stream.path == path)
            .ok_or(OleError::StreamNotFound)?;
        stream.data = data.into();
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover(&mut check, candidate.format, candidate.limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    pub fn add_stream(&mut self, path: Vec<String>, data: Vec<u8>) -> Result<(), OleError> {
        if path.is_empty() || data.len() as u64 > self.limits.max_stream_size {
            return Err(OleError::InvalidFormat(
                "new package stream path or size is invalid".into(),
            ));
        }
        if self.streams.iter().any(|stream| stream.path == path) {
            return Err(OleError::InvalidFormat(format!(
                "package stream {path:?} already exists"
            )));
        }
        if path.len() > 1
            && !self
                .storages
                .iter()
                .any(|storage| storage.path == path[..path.len() - 1])
        {
            return Err(OleError::InvalidFormat(
                "new package stream parent storage is missing".into(),
            ));
        }
        let mut candidate = self.clone();
        candidate.streams.push(CapturedStream {
            path,
            data: data.into(),
        });
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover(&mut check, candidate.format, candidate.limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Adds a storage whose format-specific reference has already been staged.
    pub fn add_storage(&mut self, id: &str, compound_file: Vec<u8>) -> Result<(), OleError> {
        let path = match self.format {
            Format::Doc if is_doc_name(id) => {
                vec!["ObjectPool".to_string(), id.to_string()]
            },
            Format::Xls if is_xls_name(id) => vec![id.to_string()],
            _ => {
                return Err(OleError::InvalidFormat(
                    "invalid format-specific object storage name".into(),
                ));
            },
        };
        if self.storages.iter().any(|value| value.path == path) {
            return Err(OleError::InvalidFormat(format!(
                "object storage {id:?} already exists"
            )));
        }
        let mut nested = OleFile::open(Cursor::new(compound_file))?;
        reject_protected_package(&nested)?;
        let mut storages = Vec::new();
        let mut streams = Vec::new();
        capture_container(
            &mut nested,
            Vec::new(),
            &mut storages,
            &mut streams,
            self.limits,
        )?;
        let object_clsid = nested
            .root_entry()
            .and_then(|entry| parse_clsid_string(&entry.clsid));
        let mut candidate = self.clone();
        candidate.storages.push(CapturedStorage {
            path: path.clone(),
            clsid: object_clsid,
        });
        for storage in storages {
            candidate.storages.push(CapturedStorage {
                path: join(&path, &storage.path),
                clsid: storage.clsid,
            });
        }
        for stream in streams {
            candidate.streams.push(CapturedStream {
                path: join(&path, &stream.path),
                data: stream.data,
            });
        }
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover(&mut check, candidate.format, candidate.limits)?;
        if candidate.objects.get(id).is_none() {
            return Err(OleError::InvalidFormat(
                "new object storage was not discoverable".into(),
            ));
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Removes a storage after its format-specific references have been removed.
    pub fn remove_storage(&mut self, id: &str) -> Result<Arc<[u8]>, OleError> {
        let object = self
            .objects
            .get(id)
            .ok_or_else(|| OleError::InvalidFormat(format!("embedded object {id:?} not found")))?;
        let removed = Arc::clone(&object.compound);
        let path = object.path.clone();
        let mut candidate = self.clone();
        candidate
            .storages
            .retain(|value| !value.path.starts_with(&path));
        candidate
            .streams
            .retain(|value| !value.path.starts_with(&path));
        let rendered = candidate.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        candidate.objects = discover(&mut check, candidate.format, candidate.limits)?;
        candidate.changed = true;
        *self = candidate;
        Ok(removed)
    }

    pub fn finish(self) -> Result<Vec<u8>, OleError> {
        if self.changed {
            return self.render();
        }
        Ok(match Arc::try_unwrap(self.original) {
            Ok(bytes) => bytes,
            Err(bytes) => bytes.as_ref().clone(),
        })
    }

    fn render(&self) -> Result<Vec<u8>, OleError> {
        let mut writer = OleWriter::with_sector_size(self.sector_size);
        if let Some(clsid) = self.root_clsid {
            writer.set_root_clsid(clsid);
        }
        let mut storages = self.storages.clone();
        storages.sort_by_key(|storage| storage.path.len());
        for storage in &storages {
            let refs = path_refs(&storage.path);
            writer.create_storage(&refs)?;
            if let Some(clsid) = storage.clsid {
                writer.set_storage_clsid(&refs, clsid)?;
            }
        }
        for stream in &self.streams {
            let refs = path_refs(&stream.path);
            writer.create_stream(&refs, stream.data.as_ref())?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }
}

fn validate_limits(limits: Limits) -> Result<(), OleError> {
    if limits.max_objects == 0
        || limits.max_storage_depth == 0
        || limits.max_streams_per_object == 0
        || limits.max_stream_size == 0
        || limits.max_object_size == 0
        || limits.max_total_size == 0
        || limits.max_metadata_string_bytes == 0
    {
        return Err(OleError::InvalidFormat(
            "all object limits must be non-zero".into(),
        ));
    }
    Ok(())
}

fn find_object_storages<R: Read + Seek>(
    ole: &OleFile<R>,
    format: Format,
    max: usize,
) -> Result<Vec<Vec<String>>, OleError> {
    let mut paths = Vec::new();
    match format {
        Format::Doc if ole.exists(&["ObjectPool"]) => {
            for entry in ole.list_directory_entries(&["ObjectPool"])? {
                if entry.entry_type == STORAGE && is_doc_name(&entry.name) {
                    paths.push(vec!["ObjectPool".into(), entry.name.clone()]);
                }
            }
        },
        Format::Doc => {},
        Format::Xls => {
            for entry in ole.list_directory_entries(&[])? {
                if entry.entry_type == STORAGE && is_xls_name(&entry.name) {
                    paths.push(vec![entry.name.clone()]);
                }
            }
        },
    }
    if paths.len() > max {
        return Err(OleError::InvalidFormat(format!(
            "embedded object count {} exceeds limit {max}",
            paths.len()
        )));
    }
    paths.sort();
    Ok(paths)
}

fn read_object<R: Read + Seek>(
    ole: &mut OleFile<R>,
    format: Format,
    path: Vec<String>,
    limits: Limits,
) -> Result<Object, OleError> {
    let (id, parent) = path
        .split_last()
        .ok_or_else(|| OleError::InvalidFormat("object storage path is empty".into()))?;
    let parent_refs = path_refs(parent);
    let id = id.clone();
    let storage_clsid = ole
        .list_directory_entries(&parent_refs)?
        .into_iter()
        .find(|entry| entry.name == id && entry.entry_type == STORAGE)
        .map(|entry| entry.clsid.clone())
        .ok_or_else(|| OleError::InvalidFormat(format!("object storage {path:?} not found")))?;
    let mut storages = Vec::new();
    let mut streams = Vec::new();
    capture_subtree(ole, &path, Vec::new(), &mut storages, &mut streams, limits)?;

    let comp_obj = stream_named(&streams, COMPOBJ)
        .map(|bytes| parse_comp_obj(bytes, limits.max_metadata_string_bytes))
        .transpose()?;
    let (host, doc_flags) = if format == Format::Doc {
        match stream_named(&streams, OBJINFO) {
            Some(bytes) => (
                Some(Arc::<[u8]>::from(bytes)),
                Some(parse_doc_flags(bytes)?),
            ),
            None => (None, None),
        }
    } else {
        (None, None)
    };
    let native = stream_named(&streams, OLE10_NATIVE)
        .map(|bytes| parse_native_package(bytes, limits))
        .transpose()?;
    let mut previews = Vec::new();
    for stream in &streams {
        let Some(name) = stream.path.last() else {
            continue;
        };
        let kind = if name.starts_with("\u{2}OlePres") {
            Some(PreviewKind::OlePresentation)
        } else if name == "\u{3}PRINT" {
            Some(PreviewKind::PrintWmf)
        } else if name == "\u{3}EPRINT" {
            Some(PreviewKind::PrintEmf)
        } else {
            None
        };
        if let Some(kind) = kind {
            previews.push(Preview {
                kind,
                stream: name.clone(),
                data: Arc::clone(&stream.data),
            });
        }
    }

    let mut writer = OleWriter::new();
    if let Some(clsid) = parse_clsid_string(&storage_clsid) {
        writer.set_root_clsid(clsid);
    }
    storages.sort_by_key(|storage| storage.path.len());
    for storage in &storages {
        let refs = path_refs(&storage.path);
        writer.create_storage(&refs)?;
        if let Some(clsid) = storage.clsid {
            writer.set_storage_clsid(&refs, clsid)?;
        }
    }
    for stream in &streams {
        writer.create_stream(&path_refs(&stream.path), stream.data.as_ref())?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let compound: Arc<[u8]> = output.into_inner().into();
    if compound.len() as u64 > limits.max_object_size {
        return Err(OleError::InvalidFormat(format!(
            "object {id:?} exceeds size limit"
        )));
    }

    let storage_ref = match format {
        Format::Doc => id.strip_prefix('_').and_then(|v| v.parse().ok()),
        Format::Xls => id.get(3..).and_then(|v| u32::from_str_radix(v, 16).ok()),
    };
    let kind = if format == Format::Xls && id.starts_with("LNK") {
        Kind::Linked
    } else if doc_flags.is_some_and(|flags| flags.activex) {
        Kind::ActiveXControl
    } else if doc_flags.is_some_and(|flags| flags.linked) {
        Kind::Linked
    } else {
        Kind::Embedded
    };
    let clsid = comp_obj
        .as_ref()
        .map(|v| v.clsid.clone())
        .filter(|v| v != "00000000-0000-0000-0000-000000000000")
        .unwrap_or(storage_clsid);
    let prog_id = comp_obj
        .as_ref()
        .and_then(|v| v.unicode_prog_id.clone().or_else(|| v.prog_id.clone()));
    let display_name = comp_obj
        .as_ref()
        .and_then(|v| v.unicode_user_type.clone().or_else(|| v.user_type.clone()));
    let clipboard_format = comp_obj.as_ref().and_then(|v| {
        v.unicode_clipboard_format
            .clone()
            .or_else(|| v.clipboard_format.clone())
    });
    let link = if kind == Kind::Linked {
        native
            .as_ref()
            .map(|v| v.command.clone())
            .filter(|v| !v.is_empty())
            .or_else(|| display_name.clone())
    } else {
        None
    };
    Ok(Object {
        id,
        path,
        storage_ref,
        kind,
        clsid,
        prog_id,
        display_name,
        clipboard_format,
        link,
        metadata: comp_obj,
        host,
        native,
        previews,
        compound,
    })
}

fn capture_container<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: Vec<String>,
    storages: &mut Vec<CapturedStorage>,
    streams: &mut Vec<CapturedStream>,
    limits: Limits,
) -> Result<(), OleError> {
    if path.len() > limits.max_storage_depth {
        return Err(OleError::InvalidFormat(
            "CFB storage nesting limit exceeded".into(),
        ));
    }
    let entries = ole
        .list_directory_entries(&path_refs(&path))?
        .into_iter()
        .cloned()
        .collect::<Vec<_>>();
    for entry in entries {
        let mut child = path.clone();
        child.push(entry.name);
        if entry.entry_type == STORAGE {
            storages.push(CapturedStorage {
                path: child.clone(),
                clsid: parse_clsid_string(&entry.clsid),
            });
            capture_container(ole, child, storages, streams, limits)?;
        } else if entry.entry_type == STREAM {
            if entry.size > limits.max_stream_size {
                return Err(OleError::InvalidFormat(format!(
                    "stream {child:?} exceeds size limit"
                )));
            }
            let aggregate_limit = limits
                .max_streams_per_object
                .saturating_mul(limits.max_objects);
            if streams.len() >= aggregate_limit {
                return Err(OleError::InvalidFormat(
                    "CFB stream count limit exceeded".into(),
                ));
            }
            let data = ole.open_stream(&path_refs(&child))?;
            streams.push(CapturedStream {
                path: child,
                data: data.into(),
            });
        }
    }
    Ok(())
}

fn capture_subtree<R: Read + Seek>(
    ole: &mut OleFile<R>,
    absolute: &[String],
    relative: Vec<String>,
    storages: &mut Vec<CapturedStorage>,
    streams: &mut Vec<CapturedStream>,
    limits: Limits,
) -> Result<(), OleError> {
    if relative.len() > limits.max_storage_depth {
        return Err(OleError::InvalidFormat(
            "object storage nesting limit exceeded".into(),
        ));
    }
    let current = join(absolute, &relative);
    let entries = ole
        .list_directory_entries(&path_refs(&current))?
        .into_iter()
        .cloned()
        .collect::<Vec<_>>();
    for entry in entries {
        let mut child = relative.clone();
        child.push(entry.name);
        if entry.entry_type == STORAGE {
            storages.push(CapturedStorage {
                path: child.clone(),
                clsid: parse_clsid_string(&entry.clsid),
            });
            capture_subtree(ole, absolute, child, storages, streams, limits)?;
        } else if entry.entry_type == STREAM {
            if streams.len() >= limits.max_streams_per_object || entry.size > limits.max_stream_size
            {
                return Err(OleError::InvalidFormat(
                    "object stream resource limit exceeded".into(),
                ));
            }
            let full = join(absolute, &child);
            let data = ole.open_stream(&path_refs(&full))?;
            streams.push(CapturedStream {
                path: child,
                data: data.into(),
            });
        }
    }
    Ok(())
}

fn reject_protected_package<R: Read + Seek>(ole: &OleFile<R>) -> Result<(), OleError> {
    for path in ole.list_streams() {
        if path.iter().any(|name| {
            matches!(
                name.to_ascii_lowercase().as_str(),
                "_xmlsignatures"
                    | "_signatures"
                    | "\u{5}digitalsignature"
                    | "\u{6}dataspaces"
                    | "encryptioninfo"
                    | "encryptedpackage"
                    | "\u{9}drmcontent"
            )
        }) {
            return Err(OleError::InvalidFormat(
                "signed, encrypted, or DRM packages are not eligible for object editing".into(),
            ));
        }
    }
    Ok(())
}

fn parse_comp_obj(data: &[u8], max: usize) -> Result<CompObj, OleError> {
    if data.len() < 28 {
        return Err(OleError::InvalidFormat(
            "CompObj stream is truncated".into(),
        ));
    }
    let clsid_bytes: &[u8; 16] = data
        .get(12..28)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or_else(|| OleError::InvalidFormat("CompObj CLSID is truncated".into()))?;
    let clsid = format_clsid(clsid_bytes);
    let mut offset = 28;
    let user_type = read_ansi(data, &mut offset, max)?;
    let clipboard_format = read_ansi(data, &mut offset, max)?;
    let prog_id = read_ansi(data, &mut offset, max)?;
    let (mut unicode_user_type, mut unicode_clipboard_format, mut unicode_prog_id) =
        (None, None, None);
    if data.len().saturating_sub(offset) >= 4 && read_u32(data, offset)? == 0x71B2_39F4 {
        offset += 4;
        unicode_user_type = read_utf16(data, &mut offset, max)?;
        unicode_clipboard_format = read_utf16(data, &mut offset, max)?;
        unicode_prog_id = read_utf16(data, &mut offset, max)?;
    }
    Ok(CompObj {
        clsid,
        user_type,
        clipboard_format,
        prog_id,
        unicode_user_type,
        unicode_clipboard_format,
        unicode_prog_id,
    })
}

#[derive(Clone, Copy)]
struct DocFlags {
    linked: bool,
    activex: bool,
}

fn parse_doc_flags(data: &[u8]) -> Result<DocFlags, OleError> {
    if data.len() != 4 && data.len() != 6 {
        return Err(OleError::InvalidFormat(
            "DOC ObjInfo ODT must be 4 or 6 bytes".into(),
        ));
    }
    let first = u16::from_le_bytes(
        data.get(..2)
            .and_then(|bytes| bytes.try_into().ok())
            .ok_or_else(|| OleError::InvalidFormat("DOC ObjInfo ODT is truncated".into()))?,
    );
    if first & ((1 << 10) | (1 << 11)) != 0 {
        return Err(OleError::InvalidFormat(
            "DOC ObjInfo reserved bits are set".into(),
        ));
    }
    let activex = first & (1 << 12) != 0;
    let stream_control = first & (1 << 13) != 0;
    if stream_control && !activex {
        return Err(OleError::InvalidFormat(
            "DOC stream control requires ActiveX".into(),
        ));
    }
    let second = if data.len() == 6 {
        u16::from_le_bytes(
            data.get(4..6)
                .and_then(|bytes| bytes.try_into().ok())
                .ok_or_else(|| {
                    OleError::InvalidFormat("DOC ObjInfo extension is truncated".into())
                })?,
        )
    } else {
        0
    };
    if second & 2 != 0 {
        return Err(OleError::InvalidFormat(
            "DOC ObjInfo reserved bit is set".into(),
        ));
    }
    Ok(DocFlags {
        linked: first & (1 << 4) != 0,
        activex,
    })
}

fn parse_native_package(data: &[u8], limits: Limits) -> Result<Native, OleError> {
    if data.len() < 6 {
        return Err(OleError::InvalidFormat(
            "Ole10Native stream is truncated".into(),
        ));
    }
    let declared = read_u32(data, 0)? as usize;
    let end = 4usize
        .checked_add(declared)
        .ok_or_else(|| OleError::InvalidFormat("Ole10Native size overflow".into()))?;
    if end > data.len() {
        return Err(OleError::InvalidFormat(
            "Ole10Native size exceeds stream".into(),
        ));
    }
    let bytes = &data[..end];
    let flags = u16::from_le_bytes([bytes[4], bytes[5]]);
    let mut offset = 6;
    let label = read_c_string(bytes, &mut offset, limits.max_metadata_string_bytes)?;
    let file_name = read_c_string(bytes, &mut offset, limits.max_metadata_string_bytes)?;
    if bytes.len().saturating_sub(offset) < 4 {
        return Err(OleError::InvalidFormat(
            "Ole10Native flags truncated".into(),
        ));
    }
    offset += 4;
    let command = read_c_string(bytes, &mut offset, limits.max_metadata_string_bytes)?;
    let native_size = read_u32(bytes, offset)? as usize;
    offset += 4;
    let native_end = offset
        .checked_add(native_size)
        .ok_or_else(|| OleError::InvalidFormat("Ole10Native payload overflow".into()))?;
    if native_end > bytes.len() || native_size as u64 > limits.max_object_size {
        return Err(OleError::InvalidFormat(
            "Ole10Native payload exceeds limits".into(),
        ));
    }
    Ok(Native {
        flags,
        label,
        file_name,
        command,
        data: Arc::from(&bytes[offset..native_end]),
    })
}

fn read_ansi(data: &[u8], offset: &mut usize, max: usize) -> Result<Option<String>, OleError> {
    let length = read_u32(data, *offset)? as usize;
    *offset += 4;
    if length == 0 || length == u32::MAX as usize {
        return Ok(None);
    }
    if length > max || data.len().saturating_sub(*offset) < length {
        return Err(OleError::InvalidFormat(
            "CompObj ANSI string exceeds bounds".into(),
        ));
    }
    let bytes = &data[*offset..*offset + length];
    *offset += length;
    Ok(Some(
        String::from_utf8_lossy(bytes.strip_suffix(&[0]).unwrap_or(bytes)).into_owned(),
    ))
}

fn read_utf16(data: &[u8], offset: &mut usize, max: usize) -> Result<Option<String>, OleError> {
    let units = read_u32(data, *offset)? as usize;
    *offset += 4;
    if units == 0 || units == u32::MAX as usize {
        return Ok(None);
    }
    let byte_len = units
        .checked_mul(2)
        .ok_or_else(|| OleError::InvalidFormat("CompObj Unicode size overflow".into()))?;
    if byte_len > max || data.len().saturating_sub(*offset) < byte_len {
        return Err(OleError::InvalidFormat(
            "CompObj Unicode string exceeds bounds".into(),
        ));
    }
    let mut units = data[*offset..*offset + byte_len]
        .chunks_exact(2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .collect::<Vec<_>>();
    *offset += byte_len;
    if units.last() == Some(&0) {
        units.pop();
    }
    Ok(Some(String::from_utf16_lossy(&units)))
}

fn read_c_string(data: &[u8], offset: &mut usize, max: usize) -> Result<String, OleError> {
    let remaining = data
        .get(*offset..)
        .ok_or_else(|| OleError::InvalidFormat("invalid string offset".into()))?;
    let end = remaining
        .iter()
        .position(|byte| *byte == 0)
        .ok_or_else(|| OleError::InvalidFormat("unterminated native-package string".into()))?;
    if end > max {
        return Err(OleError::InvalidFormat(
            "native-package string exceeds limit".into(),
        ));
    }
    *offset += end + 1;
    Ok(String::from_utf8_lossy(&remaining[..end]).into_owned())
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, OleError> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| OleError::InvalidFormat("object metadata offset overflow".into()))?;
    let value = data
        .get(offset..end)
        .ok_or_else(|| OleError::InvalidFormat("object metadata is truncated".into()))?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn stream_named<'a>(streams: &'a [CapturedStream], name: &str) -> Option<&'a [u8]> {
    streams
        .iter()
        .find(|stream| stream.path.len() == 1 && stream.path[0] == name)
        .map(|stream| stream.data.as_ref())
}

fn is_doc_name(name: &str) -> bool {
    name.strip_prefix('_')
        .map(|v| !v.is_empty() && v.bytes().all(|b| b.is_ascii_digit()))
        .unwrap_or(false)
}

fn is_xls_name(name: &str) -> bool {
    (name.starts_with("MBD") || name.starts_with("LNK"))
        && name.len() == 11
        && name.as_bytes()[3..].iter().all(|b| b.is_ascii_hexdigit())
}

fn path_refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}
fn join(left: &[String], right: &[String]) -> Vec<String> {
    left.iter().chain(right).cloned().collect()
}

fn format_clsid(bytes: &[u8; 16]) -> String {
    format!(
        "{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
        bytes[3],
        bytes[2],
        bytes[1],
        bytes[0],
        bytes[5],
        bytes[4],
        bytes[7],
        bytes[6],
        bytes[8],
        bytes[9],
        bytes[10],
        bytes[11],
        bytes[12],
        bytes[13],
        bytes[14],
        bytes[15]
    )
}

fn parse_clsid_string(value: &str) -> Option<[u8; 16]> {
    let value = value
        .trim_matches(|c| c == '{' || c == '}')
        .replace('-', "");
    let bytes = value.as_bytes();
    if bytes.len() != 32 || !bytes.iter().all(u8::is_ascii_hexdigit) {
        return None;
    }
    let mut canonical = [0u8; 16];
    for (index, byte) in canonical.iter_mut().enumerate() {
        let high = hex(bytes[index * 2])?;
        let low = hex(bytes[index * 2 + 1])?;
        *byte = (high << 4) | low;
    }
    Some([
        canonical[3],
        canonical[2],
        canonical[1],
        canonical[0],
        canonical[5],
        canonical[4],
        canonical[7],
        canonical[6],
        canonical[8],
        canonical[9],
        canonical[10],
        canonical[11],
        canonical[12],
        canonical[13],
        canonical[14],
        canonical[15],
    ])
}

fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}
