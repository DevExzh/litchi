//! Immutable master-document package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{OwnedPackage, PackageWriter, family::Package},
};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-master";
// Family structure is validated namespace-aware after package ingress. An
// empty compatibility marker avoids making arbitrary XML prefixes semantic.
const BODY_MARKER: &str = "";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

struct State {
    package: Package,
    semantics: crate::codec::Semantics,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path, MIMETYPE, BODY_MARKER, "ODM")?;
        Self::from_package(package)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODM")?;
        Self::from_package(package)
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        let package = Package::from_shared_bytes(bytes, MIMETYPE, BODY_MARKER, "ODM")?;
        Self::from_package(package)
    }

    fn from_package(package: Package) -> Result<Self> {
        let semantics = crate::codec::parse(package.content_xml())?;
        Ok(Self(Arc::new(State { package, semantics })))
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

    pub(crate) fn meta_xml(&self) -> Result<Option<String>> {
        let archive = self.0.package.package();
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
        compact_xml::validate(meta_xml.as_bytes()).map_err(Error::from)?;
        self.rebuild(self.content_xml(), Some(meta_xml))
    }

    pub(crate) fn with_content_xml(&self, content_xml: &str) -> Result<Self> {
        compact_xml::validate(content_xml.as_bytes()).map_err(Error::from)?;
        self.rebuild(content_xml, None)
    }

    fn rebuild(&self, content_xml: &str, replacement_meta_xml: Option<&str>) -> Result<Self> {
        let archive = self.0.package.package();
        ensure_editable_xml(archive, &self.files()?)?;
        let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
        writer.set_mimetype(MIMETYPE)?;
        add_preserving_media_type(&mut writer, archive, "content.xml", content_xml.as_bytes())?;
        if let Some(styles_xml) = self.styles_xml() {
            add_preserving_media_type(&mut writer, archive, "styles.xml", styles_xml.as_bytes())?;
        }
        if let Some(meta_xml) = replacement_meta_xml {
            add_preserving_media_type(&mut writer, archive, "meta.xml", meta_xml.as_bytes())?;
        } else if archive.has_file("meta.xml")? {
            add_preserving_media_type(
                &mut writer,
                archive,
                "meta.xml",
                &archive.get_file("meta.xml")?,
            )?;
        }
        writer.copy_auxiliary_files_from(archive)?;
        Self::from_bytes(writer.finish_to_bounded_bytes()?)
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    pub(crate) fn shared_bytes(&self) -> Arc<Vec<u8>> {
        self.0.package.shared_bytes()
    }

    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }

    pub(crate) fn references(&self) -> &[crate::model::subdocument::Reference] {
        self.0.semantics.references()
    }

    pub(crate) fn href_span(&self, reference: usize) -> Option<&std::ops::Range<usize>> {
        self.0.semantics.href_span(reference)
    }
}

fn ensure_editable_xml(archive: &OwnedPackage, files: &[String]) -> Result<()> {
    if files.iter().any(|path| {
        matches!(
            path.as_str(),
            "META-INF/documentsignatures.xml" | "META-INF/macrosignatures.xml"
        )
    }) {
        return Err(Error::InvalidFormat(
            "ODM package edits refuse signed packages".to_string(),
        ));
    }
    for path in files {
        if Path::new(path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            compact_xml::validate(&archive.get_file(path)?).map_err(Error::from)?;
        }
    }
    Ok(())
}

fn add_preserving_media_type<W: std::io::Write>(
    writer: &mut PackageWriter<W>,
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
