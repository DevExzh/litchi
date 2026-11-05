//! Workbook document properties: `\x05SummaryInformation` and
//! `\x05DocumentSummaryInformation`.

use super::super::model::Builder;
use crate::error::{Error, Result};
use litchi_ole_common::property_set::SummaryProperties;
use litchi_ole_common::property_set::summary_properties::Streams;

impl Builder {
    /// Replaces the workbook's metadata with `metadata`.
    ///
    /// The fields are written to the `\x05SummaryInformation` and
    /// `\x05DocumentSummaryInformation` streams by the shared
    /// [`SummaryProperties`] conversion
    /// ([`SummaryProperties::merge_metadata`] lists the mapping). It replaces
    /// every property set earlier, including custom ones, so set custom
    /// properties through [`Self::summary_properties_mut`] afterwards. Empty
    /// metadata writes no property stream, exactly as a builder that never
    /// set one. Nothing reads the clock: a creation or save time is written
    /// only when given.
    ///
    /// ```
    /// let mut builder = litchi_xls::Builder::new();
    /// builder.set_metadata(litchi_core::Metadata {
    ///     title: Some("Budget".to_owned()),
    ///     ..litchi_core::Metadata::default()
    /// })?;
    /// let package = builder.build()?;
    /// assert_eq!(package.metadata()?.title.as_deref(), Some("Budget"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::Metadata`], leaving the properties unchanged, for a
    /// field the property sets cannot store (`identifier`, `hyperlink_base`,
    /// unzoned timestamps) or a refused value.
    pub fn set_metadata(&mut self, metadata: litchi_core::Metadata) -> Result<()> {
        self.properties = SummaryProperties::from_metadata(&metadata)?;
        Ok(())
    }

    /// The workbook properties the builder writes, including typed counts and
    /// custom properties that [`litchi_core::Metadata`] does not carry.
    #[must_use]
    pub const fn summary_properties(&self) -> &SummaryProperties {
        &self.properties
    }

    /// Mutable access to the workbook properties, for custom properties and
    /// typed section edits.
    pub fn summary_properties_mut(&mut self) -> &mut SummaryProperties {
        &mut self.properties
    }

    /// Serializes the property streams, or `None` when no property is set
    /// (the builder then writes no property stream, as it always did).
    ///
    /// `CryptoAPI` RC4 would need the encrypted summary stream
    /// ([MS-OFFCRYPTO] 2.3.5.4), which is not produced, so that combination
    /// is refused instead of storing the properties in plaintext. XOR
    /// obfuscation and Office binary RC4 have no encrypted form for the
    /// property streams; there they are written in plaintext.
    pub(in crate::writer::core) fn property_streams(&self) -> Result<Option<Streams>> {
        if self.properties.is_empty() {
            return Ok(None);
        }
        if self.encryption.as_ref().is_some_and(|encryption| {
            matches!(
                encryption.profile,
                crate::EncryptionProfile::CryptoApiRc4 { .. }
            )
        }) {
            return Err(Error::EncryptedMetadata);
        }
        Ok(self.properties.to_streams()?)
    }
}
