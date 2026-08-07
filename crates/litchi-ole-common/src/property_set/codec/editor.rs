//! Transactional semantic editor for OLE Property Set streams.

use super::super::binding::Binding;
use super::super::model::*;
use super::package::try_path_refs;
use super::semantic::validate_section;
use super::support::allocation;
use crate::protection::reject_protected_container;
use litchi_cfb::{OleError, OleFile, OleWriter};
use std::collections::HashMap;
use std::io::Cursor;

/// Transactional editor for standard OLE Property Set streams.
pub struct Editor {
    original: Vec<u8>,
    streams: Vec<(Vec<String>, Vec<u8>)>,
    staged: HashMap<Vec<String>, Option<Vec<u8>>>,
}

impl Editor {
    pub fn new(bytes: Vec<u8>) -> Result<Self, OleError> {
        let streams = {
            let mut ole = OleFile::open(Cursor::new(bytes.as_slice()))?;
            reject_protected_container(&ole, "Property Set editing")?;
            let paths = ole.list_streams();
            let mut streams = try_vec_with_capacity(paths.len(), "property-set editor streams")?;
            for path in paths {
                let refs = try_path_refs(&path)?;
                let data = ole.open_stream(&refs)?;
                streams.push((path, data));
            }
            streams
        };
        Ok(Self {
            original: bytes,
            streams,
            staged: HashMap::new(),
        })
    }

    pub fn property_set(&self, kind: Binding) -> Result<Option<Section>, OleError> {
        let Some(stream) = self.load_stream(kind)? else {
            return Ok(None);
        };
        stream
            .section(kind.format_identifier())
            .map(try_clone_property_set)
            .transpose()
    }

    pub fn update<F>(&mut self, kind: Binding, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut Section) -> Result<(), OleError>,
    {
        let mut stream = self.load_stream(kind)?.unwrap_or_else(|| {
            let base = if kind == Binding::UserDefinedProperties {
                Binding::DocumentSummaryInformation.format_identifier()
            } else {
                kind.format_identifier()
            };
            Stream::new(Section::new(base))
        });
        if stream.section(kind.format_identifier()).is_none() {
            stream.add_section(Section::new(kind.format_identifier()))?;
        }
        let version = stream.version;
        let section = stream
            .section_mut(kind.format_identifier())
            .ok_or_else(|| invalid("Property Set section was not available after insertion"))?;
        let mut candidate = try_clone_property_set(section)?;
        edit(&mut candidate)?;
        validate_section(&candidate, version)?;
        *section = candidate;
        let bytes = stream.to_bytes()?;
        self.stage(kind, Some(bytes))?;
        Ok(())
    }

    pub fn replace(
        &mut self,
        kind: Binding,
        section: Section,
    ) -> Result<Option<Section>, OleError> {
        if section.format_identifier != kind.format_identifier() {
            return Err(invalid(
                "Replacement section format ID does not match target",
            ));
        }
        let previous = self.property_set(kind)?;
        self.update(kind, |target| {
            *target = section;
            Ok(())
        })?;
        Ok(previous)
    }

    pub fn remove(&mut self, kind: Binding) -> Result<Option<Section>, OleError> {
        let Some(previous) = self.property_set(kind)? else {
            return Ok(None);
        };
        if kind == Binding::UserDefinedProperties {
            let mut stream = self
                .load_stream(kind)?
                .ok_or_else(|| invalid("Property Set stream disappeared during removal"))?;
            stream.remove_section(USER_DEFINED_PROPERTIES_FMTID);
            self.stage(kind, Some(stream.to_bytes()?))?;
        } else {
            self.stage(kind, None)?;
        }
        Ok(Some(previous))
    }

    pub fn finish(self) -> Result<Vec<u8>, OleError> {
        if self.staged.is_empty() {
            return Ok(self.original);
        }
        let mut writer = OleWriter::new();
        let written_capacity = self
            .streams
            .len()
            .checked_add(self.staged.len())
            .ok_or_else(|| invalid("written Property Set stream count overflow"))?;
        let mut written = try_hash_set_with_capacity(written_capacity, "written streams")?;
        for (path, data) in &self.streams {
            let replacement = self.staged.get(path);
            if replacement.is_some_and(Option::is_none) {
                continue;
            }
            let bytes = replacement.and_then(Option::as_ref).unwrap_or(data);
            let refs = try_path_refs(path)?;
            writer.create_stream(&refs, bytes)?;
            written.insert(path.as_slice());
        }
        for (path, replacement) in &self.staged {
            if !written.contains(path.as_slice())
                && let Some(data) = replacement
            {
                let refs = try_path_refs(path)?;
                writer.create_stream(&refs, data)?;
            }
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    fn stage(&mut self, kind: Binding, replacement: Option<Vec<u8>>) -> Result<(), OleError> {
        let name = kind.name();
        let mut key = try_vec_with_capacity(1, "staged property-set stream path")?;
        key.push(try_clone_string(
            name.as_str(),
            "staged property-set stream path",
        )?);
        if !self.staged.contains_key(&key) {
            self.staged
                .try_reserve(1)
                .map_err(|source| allocation("staged property-set streams", source))?;
        }
        self.staged.insert(key, replacement);
        Ok(())
    }

    fn load_stream(&self, kind: Binding) -> Result<Option<Stream>, OleError> {
        let name = kind.name();
        if let Some((_, staged)) = self.staged.iter().find(|(candidate, _)| {
            candidate.len() == 1
                && candidate
                    .first()
                    .is_some_and(|candidate| candidate == name.as_str())
        }) {
            return staged.as_ref().map(|data| Stream::parse(data)).transpose();
        }
        self.streams
            .iter()
            .find(|(candidate, _)| candidate.len() == 1 && candidate[0] == name.as_str())
            .map(|(_, data)| Stream::parse(data))
            .transpose()
    }
}
