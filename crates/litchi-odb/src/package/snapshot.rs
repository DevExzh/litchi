//! Immutable database package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::family::Package;
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = litchi_odf_common::constants::ODF_DATABASE;
const BODY_MARKER: &str = "<";

struct State {
    package: Package,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Package::open(path, MIMETYPE, BODY_MARKER, "ODB").and_then(Self::validated)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODB").and_then(Self::validated)
    }

    fn validated(package: Package) -> Result<Self> {
        crate::codec::validate(package.content_xml())?;
        Ok(Self(Arc::new(State { package })))
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    pub(crate) fn protection_status(&self) -> Result<crate::ProtectionStatus> {
        let files = self.files()?;
        let signed = files.iter().any(|path| {
            let lower = path.to_ascii_lowercase();
            lower.starts_with("meta-inf/") && lower.contains("signatures")
        });
        let encrypted = self
            .0
            .package
            .package()
            .package()?
            .manifest()
            .has_encrypted_entries();
        Ok(crate::ProtectionStatus::new(signed, encrypted))
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }

    pub(crate) fn rebuild_with_content(&self, content: &str) -> Result<Self> {
        let files = self.files()?;
        if files.iter().any(|path| {
            matches!(
                path.as_str(),
                "META-INF/documentsignatures.xml" | "META-INF/macrosignatures.xml"
            )
        }) {
            return Err(Error::InvalidFormat(
                "ODB package edits refuse signed packages".to_string(),
            ));
        }
        // Opened producer documents are edited by byte-splicing only the
        // selected XML range.  Formatting whitespace in the unchanged source
        // is therefore lossless input, not generated output that needs the
        // fresh-authoring compactness gate.
        crate::codec::validate(content)?;
        let bytes = super::splice::rebuild_content(self.0.package.package(), content)?;
        Self::from_bytes(bytes)
    }
}
