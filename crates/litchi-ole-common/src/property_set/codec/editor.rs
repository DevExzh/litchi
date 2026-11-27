//! Transactional semantic editor for OLE Property Set streams.

use super::super::binding::Binding;
use super::super::model::{
    CodePage, Section, Stream, USER_DEFINED_PROPERTIES_FMTID, invalid, try_clone_property_set,
    try_clone_string, try_hash_set_with_capacity, try_vec_with_capacity,
};
use super::package::try_path_refs;
use super::semantic::validate_section;
use super::support::allocation;
use crate::Error;
use crate::protection::{Limits as ProtectionLimits, reject_protected_container};
use litchi_cfb::{OleError, OleFile, OleFileLimits, OleWriter};
use std::collections::HashMap;
use std::io::Cursor;

/// Transactional editor for standard OLE Property Set streams.
pub struct Edit {
    original: Vec<u8>,
    cfb_limits: OleFileLimits,
    streams: Vec<(Vec<String>, Vec<u8>)>,
    staged: HashMap<Vec<String>, Option<Vec<u8>>>,
}

impl Edit {
    /// Opens a compound file for transactional Property Set editing under an
    /// explicit finite CFB ingress policy.
    ///
    /// # Errors
    ///
    /// Returns an error if `bytes` are not a readable unprotected compound
    /// file, a contained stream cannot be read, or editor storage cannot be
    /// allocated.
    pub fn new(bytes: Vec<u8>, limits: OleFileLimits) -> Result<Self, Error> {
        let streams = {
            let mut ole = OleFile::open(Cursor::new(bytes.as_slice()), limits)?;
            let max_entries = ole.directory_entry_count();
            reject_protected_container(
                &ole,
                "Property Set editing",
                ProtectionLimits { max_entries },
            )?;
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
            cfb_limits: limits,
            streams,
            staged: HashMap::new(),
        })
    }

    /// Loads one property-set section by its standard binding.
    ///
    /// # Errors
    ///
    /// Returns an error if the backing stream cannot be parsed or the selected
    /// section cannot be cloned.
    pub fn property_set(&self, kind: Binding) -> Result<Option<Section>, Error> {
        let Some(stream) = self.load_stream(kind)? else {
            return Ok(None);
        };
        Ok(stream
            .section(kind.format_identifier())
            .map(try_clone_property_set)
            .transpose()?)
    }

    /// Applies an isolated typed edit to one bound property-set section.
    ///
    /// # Errors
    ///
    /// Returns an error if loading, cloning, or validating the stream fails,
    /// `edit` returns an error, or the updated stream cannot be serialized.
    pub fn update<F>(&mut self, kind: Binding, edit: F) -> Result<(), Error>
    where
        F: FnOnce(&mut Section) -> Result<(), Error>,
    {
        let mut stream = self.load_stream_for_rewrite(kind)?.unwrap_or_else(|| {
            if kind == Binding::UserDefinedProperties {
                // The user-defined section lives second in the
                // `DocumentSummaryInformation` stream, whose own section
                // [MS-OSHARED] 2.3.3.2.2 requires to carry a code page: a
                // blank one would make the stream unreadable to the typed
                // `document_summary` owner.
                let mut base =
                    Section::new(Binding::DocumentSummaryInformation.format_identifier());
                base.set_page(CodePage::WINDOWS_1252);
                Stream::new(base)
            } else {
                Stream::new(Section::new(kind.format_identifier()))
            }
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

    /// Replaces one bound property-set section after validating the result.
    ///
    /// # Errors
    ///
    /// Returns an error if the section FMTID does not match `kind` or the
    /// replacement cannot be loaded, validated, or serialized.
    pub fn replace(&mut self, kind: Binding, section: Section) -> Result<Option<Section>, Error> {
        if section.format_identifier != kind.format_identifier() {
            return Err(invalid("Replacement section format ID does not match target").into());
        }
        let previous = self.property_set(kind)?;
        self.update(kind, |target| {
            *target = section;
            Ok(())
        })?;
        Ok(previous)
    }

    /// Removes one bound property-set section or its complete stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the existing stream cannot be read or the retained
    /// document-summary stream cannot be serialized.
    pub fn remove(&mut self, kind: Binding) -> Result<Option<Section>, Error> {
        let Some(previous) = self.property_set(kind)? else {
            return Ok(None);
        };
        if kind == Binding::UserDefinedProperties {
            let mut stream = self
                .load_stream_for_rewrite(kind)?
                .ok_or_else(|| invalid("Property Set stream disappeared during removal"))?;
            stream.remove_section(USER_DEFINED_PROPERTIES_FMTID);
            self.stage(kind, Some(stream.to_bytes()?))?;
        } else {
            self.stage(kind, None)?;
        }
        Ok(Some(previous))
    }

    /// Writes all staged Property Set changes into a replacement compound file.
    ///
    /// # Errors
    ///
    /// Returns an error if composing a stream path, allocating output, or
    /// writing the replacement compound file fails.
    pub fn finish(self) -> Result<Vec<u8>, Error> {
        if self.staged.is_empty() {
            return Ok(self.original);
        }
        let mut writer = OleWriter::new();
        // Adopting the source keeps its root and storage class identifiers
        // (a Word document's `Word.Document.8` CLSID, for one) and lets the
        // writer reuse the source sector layout while the stream set still
        // matches; a source it cannot adopt is serialized from scratch.
        writer.adopt_source_layout(&self.original, self.cfb_limits)?;
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

    /// Loads a stream that is about to be re-serialized.
    ///
    /// A read repair that changed content (a lossy decode, a dropped
    /// property) would be written back as if it were the source, so such a
    /// stream is refused here; content-preserving repairs simply rewrite in
    /// canonical form.
    fn load_stream_for_rewrite(&self, kind: Binding) -> Result<Option<Stream>, OleError> {
        let stream = self.load_stream(kind)?;
        if let Some(stream) = &stream {
            if !stream.tolerance_report().preserves_content() {
                return Err(invalid(
                    "Property Set stream was read with content-changing repairs; refusing to rewrite it",
                ));
            }
        }
        Ok(stream)
    }

    fn load_stream(&self, kind: Binding) -> Result<Option<Stream>, OleError> {
        let name = kind.name();
        if let Some((_, staged)) = self.staged.iter().find(|(staged_path, _)| {
            staged_path.len() == 1
                && staged_path
                    .first()
                    .is_some_and(|component| component == name.as_str())
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
