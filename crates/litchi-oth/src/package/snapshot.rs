//! Immutable web-template package ownership and bounded content replacement.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{PackageWriter, family::Package};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-web";
// The common package shell requires a cheap marker before this family applies
// its namespace-aware structural contract below. A literal prefix is not a
// valid ODF namespace check, so it must remain deliberately non-semantic.
const PRELIMINARY_XML_MARKER: &str = "<";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

struct State {
    bookmarks: Vec<crate::bookmark::Bookmark>,
    forms: Vec<crate::form::Form>,
    headings: Vec<crate::heading::Heading>,
    lists: Vec<crate::list::List>,
    order: Vec<crate::codec::BlockOrder>,
    package: Package,
    paragraphs: Vec<crate::paragraph::Paragraph>,
    replacement_sites: Vec<Option<crate::codec::ReplacementSite>>,
    resources: Vec<crate::resource::Resource>,
    styles: Vec<crate::style::Style>,
    text_close: usize,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
    pub(crate) fn is_same(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(Package::open(
            path,
            MIMETYPE,
            PRELIMINARY_XML_MARKER,
            "OTH",
        )?)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(Package::from_bytes(
            bytes,
            MIMETYPE,
            PRELIMINARY_XML_MARKER,
            "OTH",
        )?)
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        Self::from_package(Package::from_shared_bytes(
            bytes,
            MIMETYPE,
            PRELIMINARY_XML_MARKER,
            "OTH",
        )?)
    }

    fn from_package(package: Package) -> Result<Self> {
        let projection = crate::codec::project(package.content_xml())?;
        let mut paragraphs = Vec::new();
        let mut replacement_sites = Vec::new();
        paragraphs
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph snapshot",
                source,
            })?;
        replacement_sites
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph edit sites",
                source,
            })?;
        for site in projection.paragraphs {
            paragraphs.push(site.value);
            replacement_sites.push(site.replacement);
        }
        let mut styles =
            crate::codec::project_styles(package.content_xml(), crate::style::Origin::Content)?;
        if let Some(styles_xml) = package.styles_xml() {
            styles.extend(crate::codec::project_styles(
                styles_xml,
                crate::style::Origin::Styles,
            )?);
        }
        Ok(Self(Arc::new(State {
            bookmarks: projection.bookmarks,
            forms: projection.forms,
            headings: projection.headings,
            lists: projection.lists,
            order: projection.order,
            package,
            paragraphs,
            replacement_sites,
            resources: projection.resources,
            styles,
            text_close: projection.text_close,
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

    pub(crate) fn paragraphs(&self) -> &[crate::paragraph::Paragraph] {
        &self.0.paragraphs
    }

    pub(crate) fn headings(&self) -> &[crate::heading::Heading] {
        &self.0.headings
    }

    pub(crate) fn bookmarks(&self) -> &[crate::bookmark::Bookmark] {
        &self.0.bookmarks
    }

    pub(crate) fn lists(&self) -> &[crate::list::List] {
        &self.0.lists
    }

    pub(crate) fn resources(&self) -> &[crate::resource::Resource] {
        &self.0.resources
    }

    pub(crate) fn forms(&self) -> &[crate::form::Form] {
        &self.0.forms
    }

    pub(crate) fn styles(&self) -> &[crate::style::Style] {
        &self.0.styles
    }

    pub(crate) fn order(&self) -> &[crate::codec::BlockOrder] {
        &self.0.order
    }

    pub(crate) fn replacement_site(&self, index: usize) -> Option<&crate::codec::ReplacementSite> {
        self.0.replacement_sites.get(index).and_then(Option::as_ref)
    }

    pub(crate) fn text_close(&self) -> usize {
        self.0.text_close
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
                "OTH package edits refuse signed packages".to_string(),
            ));
        }
        let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
        writer.set_mimetype(MIMETYPE)?;
        writer.add_file("content.xml", content.as_bytes())?;
        for path in ["meta.xml", "styles.xml"] {
            if self.0.package.package().has_file(path)? {
                let bytes = self.0.package.package().get_file(path)?;
                let xml = std::str::from_utf8(&bytes).map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH {path} UTF-8: {error}"))
                })?;
                let compact = crate::codec::compact_for_publication(xml)?;
                writer.add_file(path, compact.as_bytes())?;
            }
        }
        let excluded_paths = ["content.xml".to_string()];
        writer.copy_auxiliary_files_from_except(self.0.package.package(), &excluded_paths, &[])?;
        Self::from_bytes(writer.finish_to_bounded_bytes()?)
    }
}
