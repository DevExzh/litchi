//! Immutable master-document package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{OwnedPackage, PackageWriter, family::Package};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-master";
const BODY_MARKER: &str = "<office:text";

/// An immutable, validated package snapshot.
pub(crate) struct Snapshot {
    package: Package,
    semantics: crate::codec::Semantics,
}

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path, MIMETYPE, BODY_MARKER, "ODM")?;
        let semantics = crate::codec::parse(package.content_xml())?;
        Ok(Self { package, semantics })
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODM")?;
        let semantics = crate::codec::parse(package.content_xml())?;
        Ok(Self { package, semantics })
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        let package = Package::from_shared_bytes(bytes, MIMETYPE, BODY_MARKER, "ODM")?;
        let semantics = crate::codec::parse(package.content_xml())?;
        Ok(Self { package, semantics })
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    pub(crate) fn meta_xml(&self) -> Result<Option<String>> {
        let archive = self.package.package();
        archive
            .has_file("meta.xml")?
            .then(|| archive.get_file("meta.xml"))
            .transpose()?
            .map(|bytes| {
                String::from_utf8(bytes).map_err(|error| {
                    Error::InvalidFormat(format!("ODM meta.xml is not UTF-8: {error}"))
                })
            })
            .transpose()
    }

    pub(crate) fn with_meta_xml(&self, meta_xml: &str) -> Result<Self> {
        let archive = self.package.package();
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIMETYPE)?;
        add_preserving_media_type(
            &mut writer,
            archive,
            "content.xml",
            self.content_xml().as_bytes(),
        )?;
        if let Some(styles_xml) = self.styles_xml() {
            add_preserving_media_type(&mut writer, archive, "styles.xml", styles_xml.as_bytes())?;
        }
        add_preserving_media_type(&mut writer, archive, "meta.xml", meta_xml.as_bytes())?;
        writer.copy_auxiliary_files_from(archive)?;
        Self::from_bytes(writer.finish_to_bytes()?)
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    pub(crate) fn shared_bytes(&self) -> Arc<Vec<u8>> {
        self.package.shared_bytes()
    }

    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    pub(crate) fn references(&self) -> &[crate::model::subdocument::Reference] {
        self.semantics.references()
    }
}

fn add_preserving_media_type(
    writer: &mut PackageWriter,
    source: &OwnedPackage,
    path: &str,
    bytes: &[u8],
) -> Result<()> {
    let package = source.package()?;
    if let Some(media_type) = package.manifest().get_media_type(path) {
        writer.add_file_with_media_type(path, bytes, media_type)
    } else {
        writer.add_file(path, bytes)
    }
}
