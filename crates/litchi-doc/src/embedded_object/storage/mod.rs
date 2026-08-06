//! Inert OLE2 ObjectPool discovery, wrapping, and copy-on-write storage.
//!
//! Storage operations copy bounded CFB entries as opaque bytes. They never
//! instantiate, resolve, or execute an embedded object.

use super::Limits;
use super::codec::corrupted;
use crate::package::{Error as PackageError, Result};
use litchi_cfb::consts::STGTY_STORAGE;
use litchi_cfb::{OleError, OleFile, OleWriter};
use litchi_ole_common::object::{Target, Targets};
use std::io::{Cursor, Read, Seek};

pub(in crate::embedded_object) const OBJECT_POOL: &str = "ObjectPool";

pub(in crate::embedded_object) fn discover_targets(
    bytes: &[u8],
    limits: Limits,
) -> Result<(Targets, bool)> {
    let ole = OleFile::open(Cursor::new(bytes)).map_err(PackageError::from)?;
    let entries = match ole.list_directory_entries(&[OBJECT_POOL]) {
        Ok(entries) => entries,
        Err(OleError::StreamNotFound) => return Ok((Targets::default(), false)),
        Err(error) => return Err(PackageError::from(error)),
    };
    let mut targets = Targets::default();
    for entry in entries {
        if entry.entry_type != STGTY_STORAGE || !is_object_storage_name(&entry.name) {
            continue;
        }
        if targets.len() >= limits.max_objects {
            return Err(corrupted("ObjectPool storage count exceeds resource limit"));
        }
        let target = Target::new(
            entry.name.clone(),
            [OBJECT_POOL.to_owned(), entry.name.clone()],
        )
        .map_err(PackageError::from)?;
        targets.push(target).map_err(PackageError::from)?;
    }
    Ok((targets, true))
}

pub(in crate::embedded_object) fn object_target(storage_id: u32) -> Result<Target> {
    if storage_id == 0 || storage_id > i32::MAX as u32 {
        return Err(corrupted("storage ID must be a positive signed integer"));
    }
    let name = format!("_{storage_id}");
    Target::new(name.clone(), [OBJECT_POOL.to_owned(), name]).map_err(PackageError::from)
}

pub(in crate::embedded_object) fn is_object_storage_name(name: &str) -> bool {
    let Some(decimal) = name.strip_prefix('_') else {
        return false;
    };
    let digits = decimal.strip_prefix('-').unwrap_or(decimal);
    !digits.is_empty() && digits.as_bytes().iter().all(u8::is_ascii_digit)
}

fn wrap_object_storage(
    storage_name: &str,
    compound_file: Vec<u8>,
    limits: Limits,
) -> Result<Vec<u8>> {
    if compound_file.len() as u64 > limits.max_object_size {
        return Err(corrupted("embedded object exceeds resource limit"));
    }
    let mut source = OleFile::open(Cursor::new(compound_file)).map_err(PackageError::from)?;
    let mut writer = OleWriter::new();
    writer
        .create_storage(&[storage_name])
        .map_err(PackageError::from)?;
    if let Some(clsid) = source
        .root_entry()
        .and_then(|entry| parse_clsid(&entry.clsid))
    {
        writer
            .set_storage_clsid(&[storage_name], clsid)
            .map_err(PackageError::from)?;
    }
    let mut budget = ObjectCopyBudget::default();
    copy_object_contents(
        &mut source,
        &[],
        &[storage_name.to_owned()],
        &mut writer,
        &mut budget,
        limits,
    )?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).map_err(PackageError::from)?;
    Ok(output.into_inner())
}

pub(in crate::embedded_object) fn add_wrapped_storage(
    storage_name: &str,
    compound_file: Vec<u8>,
    limits: Limits,
) -> Result<Vec<u8>> {
    wrap_object_storage(storage_name, compound_file, limits)
}

#[derive(Default)]
struct ObjectCopyBudget {
    streams: usize,
    bytes: u64,
}

impl ObjectCopyBudget {
    fn charge(&mut self, bytes: u64, limits: Limits) -> Result<()> {
        if self.streams >= limits.max_streams_per_object {
            return Err(corrupted(
                "embedded object stream count exceeds resource limit",
            ));
        }
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or_else(|| corrupted("embedded object size overflow"))?;
        if self.bytes > limits.max_object_size {
            return Err(corrupted("embedded object exceeds resource limit"));
        }
        self.streams += 1;
        Ok(())
    }
}

fn copy_object_contents<R: Read + Seek>(
    source: &mut OleFile<R>,
    source_path: &[String],
    destination_path: &[String],
    writer: &mut OleWriter,
    budget: &mut ObjectCopyBudget,
    limits: Limits,
) -> Result<()> {
    let entries = source
        .list_directory_entries(&path_refs(source_path))
        .map_err(PackageError::from)?
        .into_iter()
        .cloned()
        .collect::<Vec<_>>();
    for entry in entries {
        let mut source_child = source_path.to_vec();
        source_child.push(entry.name.clone());
        let mut destination_child = destination_path.to_vec();
        destination_child.push(entry.name.clone());
        match entry.entry_type {
            STGTY_STORAGE => {
                if source_child.len() > limits.max_storage_depth {
                    return Err(corrupted(
                        "embedded object storage depth exceeds resource limit",
                    ));
                }
                let destination_refs = path_refs(&destination_child);
                writer
                    .create_storage(&destination_refs)
                    .map_err(PackageError::from)?;
                if let Some(clsid) = parse_clsid(&entry.clsid) {
                    writer
                        .set_storage_clsid(&destination_refs, clsid)
                        .map_err(PackageError::from)?;
                }
                copy_object_contents(
                    source,
                    &source_child,
                    &destination_child,
                    writer,
                    budget,
                    limits,
                )?;
            },
            litchi_cfb::consts::STGTY_STREAM => {
                if entry.size > limits.max_stream_size {
                    return Err(corrupted("embedded object stream exceeds resource limit"));
                }
                let data = source
                    .open_stream(&path_refs(&source_child))
                    .map_err(PackageError::from)?;
                if data.len() as u64 != entry.size {
                    return Err(corrupted(
                        "embedded object stream size changed during capture",
                    ));
                }
                budget.charge(entry.size, limits)?;
                writer
                    .create_stream_owned(&path_refs(&destination_child), data)
                    .map_err(PackageError::from)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn path_refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}

fn parse_clsid(input: &str) -> Option<[u8; 16]> {
    let value = input
        .trim_matches(|character| character == '{' || character == '}')
        .replace('-', "");
    if value.len() != 32 || !value.as_bytes().iter().all(u8::is_ascii_hexdigit) {
        return None;
    }
    let mut bytes = [0u8; 16];
    for (index, byte) in bytes.iter_mut().enumerate() {
        *byte = hex(value.as_bytes()[index * 2])? << 4 | hex(value.as_bytes()[index * 2 + 1])?;
    }
    Some([
        bytes[3], bytes[2], bytes[1], bytes[0], bytes[5], bytes[4], bytes[7], bytes[6], bytes[8],
        bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15],
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
