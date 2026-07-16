//! Inert RTF file-table metadata.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::collections::HashSet;

pub(crate) const MAX_FILE_TABLE_ENTRIES: usize = 4_096;
pub(crate) const MAX_FILE_NAME_BYTES: usize = 4_096;
pub(crate) const MAX_FILE_TABLE_TEXT_BYTES: usize = 1_048_576;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct FileSystemValidity {
    pub mac: bool,
    pub dos: bool,
    pub ntfs: bool,
    pub hpfs: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FileLocation {
    #[default]
    Local,
    Network,
    NonFileSystem,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileTableEntry<'a> {
    pub id: u32,
    pub name: Cow<'a, str>,
    pub relative_path_level: Option<u8>,
    pub operating_system: Option<u8>,
    pub valid_on: FileSystemValidity,
    pub location: FileLocation,
}

impl<'a> FileTableEntry<'a> {
    pub fn new(id: u32, name: Cow<'a, str>) -> Self {
        Self {
            id,
            name,
            relative_path_level: None,
            operating_system: None,
            valid_on: FileSystemValidity::default(),
            location: FileLocation::Local,
        }
    }

    pub fn validate(&self) -> RtfResult<()> {
        if self.id > i32::MAX as u32 {
            return Err(RtfError::MalformedDocument(
                "RTF file identifier exceeds the signed parameter range".to_string(),
            ));
        }
        if self.name.is_empty() || self.name.len() > MAX_FILE_NAME_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF file-table name is empty or exceeds the safety limit".to_string(),
            ));
        }
        if self.name.contains('\0') {
            return Err(RtfError::MalformedDocument(
                "RTF file-table name contains a NUL character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> FileTableEntry<'static> {
        FileTableEntry {
            id: self.id,
            name: Cow::Owned(self.name.into_owned()),
            relative_path_level: self.relative_path_level,
            operating_system: self.operating_system,
            valid_on: self.valid_on,
            location: self.location,
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FileTable<'a> {
    entries: Vec<FileTableEntry<'a>>,
}

impl<'a> FileTable<'a> {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn entries(&self) -> &[FileTableEntry<'a>] {
        &self.entries
    }

    pub fn get(&self, id: u32) -> Option<&FileTableEntry<'a>> {
        self.entries.iter().find(|entry| entry.id == id)
    }

    pub fn add(&mut self, entry: FileTableEntry<'a>) -> RtfResult<()> {
        entry.validate()?;
        if self.entries.len() >= MAX_FILE_TABLE_ENTRIES {
            return Err(RtfError::MalformedDocument("RTF file table exceeds the entry limit".to_string()));
        }
        if self.entries.last().is_some_and(|previous| previous.id >= entry.id) {
            return Err(RtfError::MalformedDocument("RTF file-table IDs are duplicated or out of order".to_string()));
        }
        self.entries.push(entry);
        Ok(())
    }

    pub fn validate(&self) -> RtfResult<()> {
        if self.entries.is_empty() || self.entries.len() > MAX_FILE_TABLE_ENTRIES {
            return Err(RtfError::MalformedDocument("RTF file table has an invalid entry count".to_string()));
        }
        let mut ids = HashSet::with_capacity(self.entries.len());
        let mut previous = None;
        let mut text_bytes = 0usize;
        for entry in &self.entries {
            entry.validate()?;
            if !ids.insert(entry.id) || previous.is_some_and(|id| id >= entry.id) {
                return Err(RtfError::MalformedDocument("RTF file-table IDs are duplicated or out of order".to_string()));
            }
            previous = Some(entry.id);
            text_bytes = text_bytes.checked_add(entry.name.len()).ok_or_else(|| {
                RtfError::MalformedDocument("RTF file-table text size overflow".to_string())
            })?;
        }
        if text_bytes > MAX_FILE_TABLE_TEXT_BYTES {
            return Err(RtfError::MalformedDocument("RTF file-table text exceeds the safety limit".to_string()));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> FileTable<'static> {
        FileTable { entries: self.entries.into_iter().map(FileTableEntry::into_owned).collect() }
    }
}
