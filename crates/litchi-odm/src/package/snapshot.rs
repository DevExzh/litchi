//! Immutable master-document package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{OwnedPackage, PackageWriter, family::Package};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-master";
// Family structure is validated namespace-aware after package ingress. An
// empty compatibility marker avoids making arbitrary XML prefixes semantic.
const BODY_MARKER: &str = "";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

struct State {
    package: Package,
    semantics: crate::codec::Semantics,
    styles: Vec<crate::style::Definition>,
    resources: crate::resource::Graph,
}

pub(crate) struct ResourceWrite {
    pub(crate) path: String,
    pub(crate) media_type: String,
    pub(crate) bytes: Vec<u8>,
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

    pub(crate) fn from_bytes_with_password(
        bytes: Vec<u8>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let package =
            Package::from_bytes_with_password(bytes, password, MIMETYPE, BODY_MARKER, "ODM")?;
        Self::from_package(package)
    }

    fn from_package(package: Package) -> Result<Self> {
        let semantics = crate::codec::parse(package.content_xml())?;
        let styles = crate::codec::parse_catalog(package.content_xml(), package.styles_xml())?;
        let resources = build_resource_graph(&package, semantics.references())?;
        Ok(Self(Arc::new(State {
            package,
            semantics,
            styles,
            resources,
        })))
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
        self.rebuild(self.content_xml(), None, Some(meta_xml), &[], &[])
    }

    pub(crate) fn with_content_xml(&self, content_xml: &str) -> Result<Self> {
        self.rebuild(content_xml, None, None, &[], &[])
    }

    pub(crate) fn with_transaction_parts(
        &self,
        content_xml: &str,
        replacement_styles_xml: Option<&str>,
        replacement_meta_xml: Option<&str>,
        removed_resources: &[String],
        resource_writes: &[ResourceWrite],
    ) -> Result<Self> {
        self.rebuild(
            content_xml,
            replacement_styles_xml,
            replacement_meta_xml,
            removed_resources,
            resource_writes,
        )
    }

    fn rebuild(
        &self,
        content_xml: &str,
        replacement_styles_xml: Option<&str>,
        replacement_meta_xml: Option<&str>,
        removed_resources: &[String],
        resource_writes: &[ResourceWrite],
    ) -> Result<Self> {
        let archive = self.0.package.package();
        ensure_editable(archive, &self.files()?)?;
        let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
        writer.set_mimetype(MIMETYPE)?;
        let compact_content = crate::codec::compact_source_xml(content_xml)?;
        add_preserving_media_type(
            &mut writer,
            archive,
            "content.xml",
            compact_content.as_bytes(),
        )?;
        if let Some(styles_xml) = replacement_styles_xml.or_else(|| self.styles_xml()) {
            let compact_styles = crate::codec::compact_source_xml(styles_xml)?;
            add_preserving_media_type(
                &mut writer,
                archive,
                "styles.xml",
                compact_styles.as_bytes(),
            )?;
        }
        if let Some(meta_xml) = replacement_meta_xml {
            let compact_meta = crate::codec::compact_source_xml(meta_xml)?;
            add_preserving_media_type(&mut writer, archive, "meta.xml", compact_meta.as_bytes())?;
        } else if archive.has_file("meta.xml")? {
            let source_meta =
                String::from_utf8(archive.get_file("meta.xml")?).map_err(|error| {
                    Error::InvalidFormat(format!("ODM meta.xml is not UTF-8: {error}"))
                })?;
            let compact_meta = crate::codec::compact_source_xml(&source_meta)?;
            add_preserving_media_type(&mut writer, archive, "meta.xml", compact_meta.as_bytes())?;
        }
        let mut excluded = Vec::new();
        excluded
            .try_reserve(
                removed_resources
                    .len()
                    .saturating_add(resource_writes.len()),
            )
            .map_err(|source| Error::Allocation {
                resource: "ODM resource exclusions",
                source,
            })?;
        excluded.extend_from_slice(removed_resources);
        excluded.extend(resource_writes.iter().map(|write| write.path.clone()));
        writer.copy_auxiliary_files_from_except(archive, &excluded, &[])?;
        for write in resource_writes {
            writer.add_file_with_media_type(&write.path, &write.bytes, &write.media_type)?;
        }
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

    pub(crate) fn section_tree(&self) -> &crate::section::Tree {
        self.0.semantics.tree()
    }

    pub(crate) fn styles(&self) -> &[crate::style::Definition] {
        &self.0.styles
    }

    pub(crate) fn resources(&self) -> &crate::resource::Graph {
        &self.0.resources
    }

    pub(crate) fn resource_bytes(&self, path: &str) -> Result<Vec<u8>> {
        self.0.package.package().get_file(path)
    }

    pub(crate) fn local_section_references(&self) -> &[(String, std::ops::Range<usize>)] {
        self.0.semantics.local_section_references()
    }
}

fn ensure_editable(archive: &OwnedPackage, files: &[String]) -> Result<()> {
    let package = archive.package()?;
    if package.manifest().has_encrypted_entries() {
        return Err(Error::Unsupported(
            "ODM package edits refuse encrypted packages".to_string(),
        ));
    }
    if files.iter().any(|path| {
        matches!(
            path.as_str(),
            "META-INF/documentsignatures.xml" | "META-INF/macrosignatures.xml"
        )
    }) {
        return Err(Error::Unsupported(
            "ODM package edits refuse signed packages".to_string(),
        ));
    }
    Ok(())
}

fn build_resource_graph(
    package: &Package,
    references: &[crate::subdocument::Reference],
) -> Result<crate::resource::Graph> {
    use litchi_core::Position;
    use std::collections::HashMap;

    let archive = package.package().package()?;
    let paths = archive.files()?;
    let mut resources = Vec::new();
    let mut indexes = HashMap::new();
    resources
        .try_reserve(paths.len())
        .map_err(|source| Error::Allocation {
            resource: "ODM package resource graph",
            source,
        })?;
    indexes
        .try_reserve(paths.len())
        .map_err(|source| Error::Allocation {
            resource: "ODM package resource index",
            source,
        })?;
    for path in paths {
        if path.ends_with('/') || matches!(path.as_str(), "mimetype" | "META-INF/manifest.xml") {
            continue;
        }
        let index = resources.len();
        indexes.insert(path.clone(), index);
        resources.push(crate::resource::Resource {
            media_type: archive.manifest().get_media_type(&path).map(str::to_owned),
            path,
            references: Vec::new(),
        });
    }
    let mut missing = Vec::new();
    for (reference_index, reference) in references.iter().enumerate() {
        let crate::subdocument::Target::Package(path) = reference.target() else {
            continue;
        };
        let position = Position::new(reference_index);
        if let Some(resource_index) = indexes.get(path).copied() {
            resources[resource_index]
                .references
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM package resource references",
                    source,
                })?;
            resources[resource_index].references.push(position);
        } else {
            missing.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "ODM missing package resources",
                source,
            })?;
            missing.push(position);
        }
    }
    Ok(crate::resource::Graph { resources, missing })
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
