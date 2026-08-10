//! Immutable image package ownership and bounded content-part rebuilding.

use crate::{
    FlatImage,
    frame::Frame,
    resource::{Edge, Graph, Node, Resource},
};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
    media,
};
use std::{collections::HashMap, path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.image";
// The shared package owner only needs a coarse non-empty XML marker here;
// `FlatImage::from_content_xml` performs the namespace-aware ODI grammar check.
const BODY_MARKER: &str = "<";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

pub(crate) struct ResourceReplacement<'a> {
    pub(crate) path: &'a str,
    pub(crate) media_type: &'a str,
    pub(crate) bytes: Option<&'a [u8]>,
}

#[derive(Clone, Copy)]
pub(crate) enum StylesReplacement<'a> {
    Preserve,
    Remove,
    Set(&'a str),
}

struct State {
    package: Package,
    content: FlatImage,
    resources: Vec<Resource>,
    resource_graph: Graph,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(Package::open(path, MIMETYPE, BODY_MARKER, "ODI")?)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODI")?)
    }

    fn from_package(package: Package) -> Result<Self> {
        let content = FlatImage::from_content_xml(package.content_xml().as_bytes().to_vec())?;
        let resources = scan_resources(&package)?;
        let resource_graph = build_resource_graph(&package, &resources)?;
        Ok(Self(Arc::new(State {
            package,
            content,
            resources,
            resource_graph,
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
                    Error::InvalidFormat(format!("ODI meta.xml is not UTF-8: {error}"))
                })
            })
            .transpose()
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
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

    pub(crate) fn frames(&self) -> &[Frame] {
        self.0.content.frames()
    }

    pub(crate) fn content_snapshot(&self) -> FlatImage {
        self.0.content.clone()
    }

    pub(crate) fn resources(&self) -> &[Resource] {
        &self.0.resources
    }

    pub(crate) fn resource_graph(&self) -> &Graph {
        &self.0.resource_graph
    }

    pub(crate) fn resource_bytes(&self, index: usize) -> Result<Option<Vec<u8>>> {
        let resource =
            self.0.resources.get(index).ok_or_else(|| {
                Error::InvalidFormat("ODI resource selector is out of bounds".into())
            })?;
        if !resource.is_present() {
            return Ok(None);
        }
        self.0.package.package().get_file(resource.path()).map(Some)
    }

    pub(crate) fn resource_file(&self, path: &str) -> Result<Option<Vec<u8>>> {
        if !self.0.package.package().has_file(path)? {
            return Ok(None);
        }
        self.0.package.package().get_file(path).map(Some)
    }

    pub(crate) fn rewrite_capability(&self) -> Result<crate::RewriteCapability> {
        let files = self.files()?;
        let mut blockers = Vec::new();
        if files.iter().any(|path| is_signature_path(path)) {
            blockers.push(crate::RewriteBlocker::Signature);
        }
        let archive = self.0.package.package();
        let inspected = archive.package()?;
        if inspected.manifest().has_encrypted_entries() {
            blockers.push(crate::RewriteBlocker::Encryption);
        }
        let mut noncompact_xml = false;
        let mut unreadable_xml = false;
        for path in files.iter().filter(|path| {
            Path::new(path)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
                && path.as_str() != "META-INF/manifest.xml"
        }) {
            if inspected
                .manifest()
                .get_entry(path)
                .is_some_and(|entry| entry.encryption.is_some())
            {
                continue;
            }
            match inspected.get_file(path) {
                Ok(bytes) => noncompact_xml |= compact_xml::validate(&bytes).is_err(),
                Err(_) => unreadable_xml = true,
            }
        }
        if noncompact_xml {
            blockers.push(crate::RewriteBlocker::NonCompactXml);
        }
        if unreadable_xml {
            blockers.push(crate::RewriteBlocker::UnreadableXml);
        }
        Ok(crate::RewriteCapability::new(blockers))
    }

    pub(crate) fn rebuild(
        &self,
        content: &str,
        replacement_styles_xml: StylesReplacement<'_>,
        replacement_meta_xml: Option<&str>,
        replacements: &[ResourceReplacement<'_>],
    ) -> Result<Self> {
        let files = self.files()?;
        if files.iter().any(|path| is_signature_path(path)) {
            return Err(Error::InvalidFormat(
                "ODI package edits refuse signed packages".to_string(),
            ));
        }
        ensure_compact_rewrite_source(self, &files)?;
        let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
        writer.set_mimetype(MIMETYPE)?;
        writer.add_file("content.xml", content.as_bytes())?;
        match replacement_styles_xml {
            StylesReplacement::Set(styles_xml) => {
                compact_xml::validate(styles_xml.as_bytes()).map_err(Error::from)?;
                writer.add_file("styles.xml", styles_xml.as_bytes())?;
            },
            StylesReplacement::Preserve => {
                if self.0.package.package().has_file("styles.xml")? {
                    writer.add_file(
                        "styles.xml",
                        &self.0.package.package().get_file("styles.xml")?,
                    )?;
                }
            },
            StylesReplacement::Remove => {},
        }
        if self.0.package.package().has_file("settings.xml")? {
            writer.add_file(
                "settings.xml",
                &self.0.package.package().get_file("settings.xml")?,
            )?;
        }
        if let Some(meta_xml) = replacement_meta_xml {
            compact_xml::validate(meta_xml.as_bytes()).map_err(Error::from)?;
            writer.add_file("meta.xml", meta_xml.as_bytes())?;
        } else if self.0.package.package().has_file("meta.xml")? {
            writer.add_file("meta.xml", &self.0.package.package().get_file("meta.xml")?)?;
        }
        let excluded = replacements
            .iter()
            .map(|replacement| replacement.path.to_string())
            .collect::<Vec<_>>();
        writer.copy_auxiliary_files_from_except(self.0.package.package(), &excluded, &[])?;
        for replacement in replacements {
            if let Some(bytes) = replacement.bytes {
                writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
            }
        }
        Self::from_bytes(writer.finish_to_bounded_bytes()?)
    }
}

fn scan_resources(package: &Package) -> Result<Vec<Resource>> {
    let archive = package.package().package()?;
    let images = media::scan_package(package.content_xml(), None, &archive)?;
    let mut resources = Vec::new();
    for (frame, image) in images.into_iter().enumerate() {
        match image.source {
            media::Source::PackagePart {
                href,
                path,
                manifest_media_type,
            } => resources.push(Resource::new(frame, href, path, manifest_media_type, true)),
            media::Source::MissingPackagePart {
                href,
                resolved_path,
            } => resources.push(Resource::new(frame, href, resolved_path, None, false)),
            media::Source::Inline { .. }
            | media::Source::Linked { .. }
            | media::Source::Missing
            | _ => {},
        }
    }
    Ok(resources)
}

fn build_resource_graph(package: &Package, resources: &[Resource]) -> Result<Graph> {
    let archive = package.package().package()?;
    let mut paths = package.files()?;
    for resource in resources {
        if !paths.iter().any(|path| path == resource.path()) {
            paths.push(resource.path().to_owned());
        }
    }
    paths.sort_unstable();
    paths.dedup();
    let referenced = resources
        .iter()
        .map(Resource::path)
        .collect::<std::collections::HashSet<_>>();
    let nodes = paths
        .into_iter()
        .map(|path| {
            let present = archive.has_file(&path);
            Node::new(
                path.clone(),
                archive.manifest().get_media_type(&path).map(str::to_owned),
                present,
                referenced.contains(path.as_str()),
            )
        })
        .collect::<Vec<_>>();
    let positions = nodes
        .iter()
        .enumerate()
        .map(|(index, node)| (node.path().to_owned(), index))
        .collect::<HashMap<_, _>>();
    let edges = resources
        .iter()
        .map(|resource| {
            positions
                .get(resource.path())
                .copied()
                .map(|node| Edge::new(resource.frame(), resource.href().to_owned(), node))
                .ok_or_else(|| Error::InvalidFormat("ODI resource graph target disappeared".into()))
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(Graph::new(nodes, edges))
}

fn is_signature_path(path: &str) -> bool {
    path.strip_prefix("META-INF/")
        .is_some_and(|name| name.contains("signatures"))
}

fn ensure_compact_rewrite_source(source: &Snapshot, files: &[String]) -> Result<()> {
    let archive = source.0.package.package();
    for path in files {
        if Path::new(path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
            && path != "META-INF/manifest.xml"
        {
            compact_xml::validate(&archive.get_file(path)?).map_err(Error::from)?;
        }
    }
    Ok(())
}
