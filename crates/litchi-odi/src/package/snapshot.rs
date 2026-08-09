//! Immutable image package ownership and bounded content-part rebuilding.

use crate::{FlatImage, frame::Frame, resource::Resource};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
    media,
};
use std::{path::Path, sync::Arc};

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

struct State {
    package: Package,
    content: FlatImage,
    resources: Vec<Resource>,
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
        Ok(Self(Arc::new(State {
            package,
            content,
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

    pub(crate) fn rebuild(
        &self,
        content: &str,
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
        for path in ["styles.xml", "meta.xml", "settings.xml"] {
            if self.0.package.package().has_file(path)? {
                writer.add_file(path, &self.0.package.package().get_file(path)?)?;
            }
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
