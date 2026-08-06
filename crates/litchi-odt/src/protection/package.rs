//! Package, flat-document, and source-package seams for protection edits.

use litchi_core::{Error, Result};

use crate::constants;
use crate::core::{OwnedPackage, PackageWriter};
use crate::generic::{FlatDocument, Package};

use super::model::Policy;
use super::transaction::Transaction;
use super::{parse_flat, parse_package};

impl Package {
    /// Read the optional document interaction policy from `settings.xml`.
    pub fn protection(&self) -> Result<Option<Policy>> {
        let Some(xml) = self.settings_xml()? else {
            return Ok(None);
        };
        parse_package(xml.as_bytes()).map(Some)
    }

    /// Replace document interaction policy metadata atomically.
    ///
    /// The package is rebuilt only after the settings transaction and all
    /// resulting XML have validated. Unmodeled settings nodes and auxiliary
    /// package entries remain untouched; policy values are inert metadata and
    /// are never enforced by this method.
    pub fn set_protection(&mut self, policy: &Policy) -> Result<Option<Policy>> {
        policy.validate()?;
        let before = self.protection()?;
        if before.as_ref().is_some_and(|value| value == policy)
            || before.is_none() && policy.is_empty()
        {
            return Ok(before);
        }
        let bytes = rewrite_owned_package(self.owned_package(), self.mimetype(), policy)?;
        *self = Self::from_bytes(bytes)?;
        Ok(before)
    }
}

impl FlatDocument {
    /// Read the document interaction policy from the flat XML document.
    pub fn protection(&self) -> Result<Policy> {
        parse_flat(self.xml().as_bytes())
    }

    /// Replace document interaction policy metadata atomically.
    pub fn set_protection(&mut self, policy: &Policy) -> Result<Policy> {
        policy.validate()?;
        let before = self.protection()?;
        if &before == policy {
            return Ok(before);
        }
        let mut transaction = Transaction::flat(self.xml().as_bytes())?;
        transaction.set(policy.clone())?;
        let committed = transaction.commit()?;
        let replacement = Self::from_bytes(committed.into_bytes())?;
        let old = before;
        *self = replacement;
        Ok(old)
    }
}

/// Rewrite one ODT package while changing only the settings part.
pub(crate) fn rewrite_owned_package(
    source: &OwnedPackage,
    mimetype: &str,
    policy: &Policy,
) -> Result<Vec<u8>> {
    let package = source.package()?;
    if package.manifest().has_encrypted_entries() {
        return Err(Error::InvalidFormat(
            "protection settings edits cannot rewrite encrypted ODF entries".to_string(),
        ));
    }

    let settings = if package.has_file(constants::ODF_SETTINGS) {
        let source_xml = package.get_file(constants::ODF_SETTINGS)?;
        let mut transaction = Transaction::package(&source_xml)?;
        transaction.set(policy.clone())?;
        Some(transaction.commit()?.into_bytes())
    } else if policy.is_empty() {
        None
    } else {
        Some(super::codec::empty_package(policy)?)
    };

    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype)?;
    for path in [
        constants::ODF_CONTENT,
        constants::ODF_STYLES,
        constants::ODF_META,
    ] {
        if package.has_file(path) {
            let bytes = package.get_file(path)?;
            writer.add_file(path, &bytes)?;
        }
    }
    if let Some(settings) = settings {
        writer.add_file(constants::ODF_SETTINGS, &settings)?;
    }
    writer.copy_auxiliary_files_from_except(source, &[constants::ODF_SETTINGS.to_string()], &[])?;
    writer.finish_to_bytes()
}
