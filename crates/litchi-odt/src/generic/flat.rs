//! Flat-document snapshot and transactional XML edits.

use super::codec::{classify_mimetype, validate_flat_document};
use super::model::{Family, FlatDocument};
use crate::core::Meta;
use litchi_core::{Error, Metadata, Resource, ResourceLimit, Result};
use std::io::Read;
use std::path::Path;

/// Default maximum encoded size for a generic flat OpenDocument snapshot.
pub(super) const DEFAULT_MAX_FLAT_DOCUMENT_BYTES: u64 = 256 * 1024 * 1024;
/// Hard maximum encoded size accepted by the generic flat-document reader.
pub(super) const HARD_MAX_FLAT_DOCUMENT_BYTES: u64 = 512 * 1024 * 1024;

impl FlatDocument {
    /// Parses the optional flat-document `office:settings` inventory.
    pub fn settings(&self) -> Result<crate::Settings> {
        crate::settings::parse_flat(self.xml())
    }

    /// Open and validate a flat `OpenDocument` XML file.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read and validate a flat `OpenDocument` XML stream.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limit(reader, DEFAULT_MAX_FLAT_DOCUMENT_BYTES)
    }

    /// Read and validate a flat `OpenDocument` XML stream under a finite
    /// encoded-byte limit.
    pub fn from_reader_with_limit(reader: impl Read, maximum: u64) -> Result<Self> {
        validate_flat_limit(maximum)?;
        let bytes = read_flat_input(reader, maximum)?;
        Self::from_bytes_with_limit(bytes, maximum)
    }

    /// Validate flat `OpenDocument` XML from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limit(bytes, DEFAULT_MAX_FLAT_DOCUMENT_BYTES)
    }

    /// Validate flat `OpenDocument` XML from owned bytes under an explicit
    /// finite encoded-byte limit.
    pub fn from_bytes_with_limit(bytes: Vec<u8>, maximum: u64) -> Result<Self> {
        validate_flat_limit(maximum)?;
        let observed = u64::try_from(bytes.len()).map_err(|_| {
            Error::InvalidFormat("flat OpenDocument input exceeds platform limits".to_string())
        })?;
        if observed > maximum {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::InputBytes,
                observed,
                limit: maximum,
                scope: "flat OpenDocument input".into(),
            }));
        }
        let mimetype = crate::detect::flat_mime(&bytes)
            .ok_or_else(|| Error::InvalidFormat("invalid flat OpenDocument root".to_string()))?;
        let (family, template) = classify_mimetype(&mimetype).ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported OpenDocument mimetype '{mimetype}'"))
        })?;
        if template || matches!(family, Family::Master | Family::Web | Family::Database) {
            return Err(Error::InvalidFormat(format!(
                "mimetype '{mimetype}' has no standard flat OpenDocument form"
            )));
        }
        let xml = String::from_utf8(bytes).map_err(|_error| {
            Error::InvalidFormat("invalid UTF-8 in flat OpenDocument".to_string())
        })?;
        validate_flat_document(&xml, family)?;
        Ok(Self {
            xml,
            family,
            mimetype,
        })
    }

    /// Return the document family.
    pub fn family(&self) -> Family {
        self.family
    }

    /// Return the root `office:mimetype` value.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the conventional flat `OpenDocument` extension.
    pub fn extension(&self) -> &'static str {
        match self.family {
            Family::Text => "fodt",
            Family::Spreadsheet => "fods",
            Family::Presentation => "fodp",
            Family::Drawing => "fodg",
            Family::Chart => "fodc",
            Family::Formula => "fodf",
            Family::Image => "fodi",
            Family::Master | Family::Web | Family::Database => {
                unreachable!("master and web flat documents are rejected")
            },
        }
    }

    /// Return the complete flat XML document.
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Extract common document metadata from the combined XML document.
    pub fn metadata(&self) -> Result<Metadata> {
        Meta::from_bytes(self.xml.as_bytes())?.try_extract_metadata()
    }

    /// Extract the complete format-specific metadata model.
    pub fn odf_metadata(&self) -> Result<crate::Metadata> {
        Meta::from_bytes(self.xml.as_bytes())?.odf_metadata()
    }

    /// Discover inline and inert linked images in the flat document.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        crate::media::scan_flat(&self.xml)
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::form::Forms> {
        crate::form::parse_form_parts(&[(self.xml(), crate::form::Part::Flat)])
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        crate::variable_declaration::parse_parts(&[(
            self.xml(),
            crate::variable_declaration::Part::Flat,
        )])
    }

    /// Return the unnamed fallback page layout, when one is declared.
    pub fn default_page_layout(&self) -> Result<Option<crate::page_layout::PageLayout>> {
        crate::page_layout::parse_default_page_layout(self.xml())
    }

    /// Atomically insert or replace one variable declaration container.
    ///
    /// The group must target the flat part. Formulas and cached values remain
    /// inert; this method only updates XML metadata and never evaluates fields.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::variable_declaration::Group,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        if group.part != crate::variable_declaration::Part::Flat {
            return Err(Error::InvalidFormat(
                "FlatDocument requires Part::Flat".to_string(),
            ));
        }
        let current = self.variable_declarations()?;
        let old = current
            .groups
            .iter()
            .find(|candidate| candidate.scope == group.scope && candidate.kind == group.kind)
            .cloned();
        let updated = crate::variable_declaration::set_xml(&self.xml, group)?;
        validate_flat_document(&updated, self.family)?;
        crate::variable_declaration::parse_parts(&[(
            updated.as_str(),
            crate::variable_declaration::Part::Flat,
        )])?;
        self.xml = updated;
        Ok(old)
    }

    /// Atomically remove one variable declaration container.
    ///
    /// Removal fails without mutation if any remaining field references a
    /// declaration owned by the container.
    pub fn remove_variable_declaration_group(
        &mut self,
        scope: &crate::variable_declaration::Scope,
        kind: crate::variable_declaration::Kind,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        let current = self.variable_declarations()?;
        let Some(old) = current
            .groups
            .iter()
            .find(|candidate| candidate.scope == *scope && candidate.kind == kind)
            .cloned()
        else {
            return Ok(None);
        };
        let updated = crate::variable_declaration::remove_xml(&self.xml, scope, kind)?;
        validate_flat_document(&updated, self.family)?;
        crate::variable_declaration::parse_parts(&[(
            updated.as_str(),
            crate::variable_declaration::Part::Flat,
        )])?;
        self.xml = updated;
        Ok(Some(old))
    }

    /// Discover inert inline and linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::Object>> {
        crate::embedded::scan_flat(&self.xml)
    }

    /// Return the exact original bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.xml.as_bytes()
    }

    /// Clone the exact original bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Consume this wrapper and return the exact original bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.xml.into_bytes()
    }

    /// Save the flat document without reconstructing its XML.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

fn validate_flat_limit(maximum: u64) -> Result<()> {
    if maximum == 0 || maximum > HARD_MAX_FLAT_DOCUMENT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "flat OpenDocument input limit must be between 1 and {HARD_MAX_FLAT_DOCUMENT_BYTES}"
        )));
    }
    usize::try_from(maximum).map(|_| ()).map_err(|_| {
        Error::InvalidFormat("flat OpenDocument input limit exceeds platform limits".to_string())
    })
}

fn read_flat_input(reader: impl Read, maximum: u64) -> Result<Vec<u8>> {
    let read_limit = maximum.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("flat OpenDocument input limit overflows".to_string())
    })?;
    let reserve = usize::try_from(read_limit.min(64 * 1024)).map_err(|_| {
        Error::InvalidFormat("flat OpenDocument input exceeds platform limits".to_string())
    })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(reserve)
        .map_err(|source| Error::Allocation {
            resource: "flat OpenDocument input",
            source,
        })?;

    let mut reader = reader.take(read_limit);
    let mut chunk = [0_u8; 8 * 1024];
    loop {
        let read = reader.read(&mut chunk)?;
        if read == 0 {
            break;
        }
        bytes
            .try_reserve(read)
            .map_err(|source| Error::Allocation {
                resource: "flat OpenDocument input",
                source,
            })?;
        bytes.extend_from_slice(&chunk[..read]);
    }
    Ok(bytes)
}
