use litchi_iwa_core::Archive;

use crate::zip::{ZipArchive, parse_iwa_components};
use crate::{Limits, Result};

/// One parsed `.iwa` component in deterministic member-name order.
#[derive(Debug)]
pub struct Component {
    name: Box<str>,
    archive: Archive,
}

impl Component {
    pub(crate) fn new(name: &str, archive: Archive) -> Self {
        Self {
            name: name.into(),
            archive,
        }
    }

    /// Return the normalized ZIP member name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the neutral parsed IWA archive.
    #[must_use]
    pub const fn archive(&self) -> &Archive {
        &self.archive
    }

    /// Consume the component and return its owned name and archive.
    #[must_use]
    pub fn into_parts(self) -> (String, Archive) {
        (self.name.into(), self.archive)
    }
}

/// Deterministic parsed `.iwa` components from one physical iWork ZIP input.
///
/// This catalog owns only ZIP/IWA ingress. Metadata, media, package
/// transactions, object indexing, and application-specific message decoding
/// remain in their respective adapter crates.
#[derive(Debug)]
pub struct ComponentCatalog {
    components: Box<[Component]>,
}

impl ComponentCatalog {
    /// Parse a ZIP bundle from memory using the default physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP, nested `Index.zip`, Snappy stream, or IWA
    /// framing is malformed, encrypted, or exceeds a resource ceiling.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a ZIP bundle from memory under explicit physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the limits are invalid, input exceeds a ceiling,
    /// or the ZIP/IWA stream is malformed or encrypted.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let archive = ZipArchive::new_with_limits(bytes, validated_limits)?;
        let components = parse_iwa_components(&archive, validated_limits)?.into_boxed_slice();
        Ok(Self { components })
    }

    /// Return the number of parsed components.
    #[must_use]
    pub fn len(&self) -> usize {
        self.components.len()
    }

    /// Return whether no IWA components were found.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.components.is_empty()
    }

    /// Iterate over components in deterministic normalized-name order.
    pub fn iter(&self) -> impl Iterator<Item = &Component> {
        self.components.iter()
    }

    /// Find one component by normalized ZIP member name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Component> {
        self.components
            .binary_search_by(|component| component.name().cmp(name))
            .ok()
            .map(|index| &self.components[index])
    }
}

impl IntoIterator for ComponentCatalog {
    type Item = Component;
    type IntoIter = std::vec::IntoIter<Component>;

    fn into_iter(self) -> Self::IntoIter {
        self.components.into_vec().into_iter()
    }
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn iwa_bytes(identifier: u64, message_type: u32) -> Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: message_type,
                    data: vec![1, 2, 3],
                }],
            )?],
        };
        Ok(SnappyStream::compress(&archive.to_bytes()?)?)
    }

    #[test]
    fn parses_direct_iwa_and_skips_operation_storage() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        writer.write_stored("Index/OperationStorage.iwa", b"bvxn opaque data")?;
        let bytes = writer.finish_to_bytes()?;

        let catalog = ComponentCatalog::from_bytes(&bytes)?;
        assert_eq!(catalog.len(), 1);
        assert!(!catalog.is_empty());
        assert_eq!(
            catalog.get("Index/Document.iwa").map(Component::name),
            Some("Index/Document.iwa")
        );
        let component = catalog.iter().next().ok_or_else(|| {
            crate::Error::InvalidBundle("component catalog unexpectedly empty".to_owned())
        })?;
        assert_eq!(component.name(), "Index/Document.iwa");
        assert_eq!(component.archive().objects[0].messages[0].type_, 6000);
        Ok(())
    }

    #[test]
    fn consumes_component_name_and_archive() -> Result<()> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let bytes = writer.finish_to_bytes()?;

        let component = ComponentCatalog::from_bytes(&bytes)?
            .into_iter()
            .next()
            .ok_or_else(|| {
                crate::Error::InvalidBundle("component catalog unexpectedly empty".to_owned())
            })?;
        let (name, archive) = component.into_parts();

        assert_eq!(name, "Index/Document.iwa");
        assert_eq!(archive.objects[0].messages[0].type_, 6000);
        Ok(())
    }

    #[test]
    fn parses_nested_index_and_rejects_encryption() -> Result<()> {
        let mut index = StreamingArchiveWriter::new();
        index.write_stored("Index/Document.iwa", &iwa_bytes(1, 6000)?)?;
        let index_bytes = index.finish_to_bytes()?;

        let mut outer = StreamingArchiveWriter::new();
        outer.write_stored("legacy.pages/Index.zip", &index_bytes)?;
        let outer_bytes = outer.finish_to_bytes()?;
        assert_eq!(ComponentCatalog::from_bytes(&outer_bytes)?.len(), 1);

        let mut encrypted = StreamingArchiveWriter::new();
        encrypted.write_stored(".iwpv2", b"metadata")?;
        encrypted.write_stored("Index/Document.iwa", b"ciphertext")?;
        let encrypted_bytes = encrypted.finish_to_bytes()?;
        let result = ComponentCatalog::from_bytes(&encrypted_bytes);
        assert!(matches!(result, Err(crate::Error::Encrypted)));
        Ok(())
    }

    #[test]
    fn rejects_input_above_profile() -> Result<()> {
        let limits = Limits::new(1, 10, 100, 100, 100)?;
        let result = ComponentCatalog::from_bytes_with_limits(b"not a zip", limits);
        assert!(matches!(result, Err(crate::Error::Limit { .. })));
        Ok(())
    }
}
