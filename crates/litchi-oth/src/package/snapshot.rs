//! Immutable web-template package ownership and bounded content replacement.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{
    PackageWriter, XmlSourcePart, XmlSplicePublication, family::Package,
};
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
    forms_site: Option<crate::codec::ReplacementSite>,
    headings: Vec<crate::heading::Heading>,
    heading_content_sites: Vec<crate::codec::ReplacementSite>,
    heading_full_sites: Vec<crate::codec::ReplacementSite>,
    heading_inline_replaceable: Vec<bool>,
    heading_replacement_sites: Vec<Option<crate::codec::ReplacementSite>>,
    lists: Vec<crate::list::List>,
    list_sites: Vec<crate::codec::ReplacementSite>,
    meta_xml: Option<String>,
    order: Vec<crate::codec::BlockOrder>,
    package: Package,
    paragraphs: Vec<crate::paragraph::Paragraph>,
    paragraph_content_sites: Vec<crate::codec::ReplacementSite>,
    paragraph_full_sites: Vec<crate::codec::ReplacementSite>,
    paragraph_inline_replaceable: Vec<bool>,
    replacement_sites: Vec<Option<crate::codec::ReplacementSite>>,
    resources: Vec<crate::resource::Resource>,
    resource_sites: Vec<crate::codec::ReplacementSite>,
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
        let meta_xml = if package.package().has_file("meta.xml")? {
            Some(
                String::from_utf8(package.package().get_file("meta.xml")?).map_err(|error| {
                    Error::InvalidFormat(format!("invalid OTH meta.xml UTF-8: {error}"))
                })?,
            )
        } else {
            None
        };
        let projection = crate::codec::project(package.content_xml())?;
        let resource_sites = crate::codec::resource_sites(package.content_xml())?;
        if resource_sites.len() != projection.resources.len() {
            return Err(Error::InvalidFormat(
                "OTH resource projection/site count differs".to_string(),
            ));
        }
        let mut headings = Vec::new();
        let mut heading_content_sites = Vec::new();
        let mut heading_full_sites = Vec::new();
        let mut heading_inline_replaceable = Vec::new();
        let mut heading_replacement_sites = Vec::new();
        headings
            .try_reserve(projection.headings.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH heading snapshot",
                source,
            })?;
        heading_content_sites
            .try_reserve(projection.headings.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH heading content edit sites",
                source,
            })?;
        heading_full_sites
            .try_reserve(projection.headings.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH heading full source sites",
                source,
            })?;
        heading_inline_replaceable
            .try_reserve(projection.headings.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH heading inline edit flags",
                source,
            })?;
        heading_replacement_sites
            .try_reserve(projection.headings.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH heading edit sites",
                source,
            })?;
        for site in projection.headings {
            heading_content_sites.push(site.content);
            heading_full_sites.push(site.full);
            heading_inline_replaceable.push(site.inline_replaceable);
            headings.push(site.value);
            heading_replacement_sites.push(site.replacement);
        }
        let mut paragraphs = Vec::new();
        let mut paragraph_content_sites = Vec::new();
        let mut paragraph_full_sites = Vec::new();
        let mut paragraph_inline_replaceable = Vec::new();
        let mut replacement_sites = Vec::new();
        paragraphs
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph snapshot",
                source,
            })?;
        paragraph_content_sites
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph content edit sites",
                source,
            })?;
        paragraph_full_sites
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph full source sites",
                source,
            })?;
        paragraph_inline_replaceable
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph inline edit flags",
                source,
            })?;
        replacement_sites
            .try_reserve(projection.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph edit sites",
                source,
            })?;
        for site in projection.paragraphs {
            paragraph_content_sites.push(site.content);
            paragraph_full_sites.push(site.full);
            paragraph_inline_replaceable.push(site.inline_replaceable);
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
            forms_site: projection.forms_site,
            headings,
            heading_content_sites,
            heading_full_sites,
            heading_inline_replaceable,
            heading_replacement_sites,
            lists: projection.lists,
            list_sites: projection.list_sites,
            meta_xml,
            order: projection.order,
            package,
            paragraphs,
            paragraph_content_sites,
            paragraph_full_sites,
            paragraph_inline_replaceable,
            replacement_sites,
            resources: projection.resources,
            resource_sites,
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

    pub(crate) fn meta_xml(&self) -> Option<&str> {
        self.0.meta_xml.as_deref()
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

    pub(crate) fn list_site(&self, index: usize) -> Option<&crate::codec::ReplacementSite> {
        self.0.list_sites.get(index)
    }

    pub(crate) fn resources(&self) -> &[crate::resource::Resource] {
        &self.0.resources
    }

    pub(crate) fn resource_site(&self, index: usize) -> Option<&crate::codec::ReplacementSite> {
        self.0.resource_sites.get(index)
    }

    pub(crate) fn resource_sites(&self) -> &[crate::codec::ReplacementSite] {
        &self.0.resource_sites
    }

    pub(crate) fn member(&self, path: &str) -> Result<Option<Vec<u8>>> {
        self.0
            .package
            .package()
            .has_file(path)?
            .then(|| self.0.package.package().get_file(path))
            .transpose()
    }

    pub(crate) fn member_media_type(&self, path: &str) -> Result<Option<String>> {
        let package = self.0.package.package().package()?;
        Ok(package.manifest().get_media_type(path).map(str::to_owned))
    }

    pub(crate) fn forms(&self) -> &[crate::form::Form] {
        &self.0.forms
    }

    pub(crate) fn forms_site(&self) -> Option<&crate::codec::ReplacementSite> {
        self.0.forms_site.as_ref()
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

    pub(crate) fn heading_replacement_site(
        &self,
        index: usize,
    ) -> Option<&crate::codec::ReplacementSite> {
        self.0
            .heading_replacement_sites
            .get(index)
            .and_then(Option::as_ref)
    }

    pub(crate) fn paragraph_content_site(
        &self,
        index: usize,
    ) -> Option<&crate::codec::ReplacementSite> {
        self.0.paragraph_content_sites.get(index)
    }

    pub(crate) fn heading_content_site(
        &self,
        index: usize,
    ) -> Option<&crate::codec::ReplacementSite> {
        self.0.heading_content_sites.get(index)
    }

    pub(crate) fn paragraph_full_site(
        &self,
        index: usize,
    ) -> Option<&crate::codec::ReplacementSite> {
        self.0.paragraph_full_sites.get(index)
    }

    pub(crate) fn heading_full_site(&self, index: usize) -> Option<&crate::codec::ReplacementSite> {
        self.0.heading_full_sites.get(index)
    }

    pub(crate) fn paragraph_inline_replaceable(&self, index: usize) -> bool {
        self.0
            .paragraph_inline_replaceable
            .get(index)
            .copied()
            .unwrap_or(false)
    }

    pub(crate) fn heading_inline_replaceable(&self, index: usize) -> bool {
        self.0
            .heading_inline_replaceable
            .get(index)
            .copied()
            .unwrap_or(false)
    }

    pub(crate) fn text_close(&self) -> usize {
        self.0.text_close
    }

    pub(crate) fn rebuild_with_parts(
        &self,
        content: &str,
        metadata: &crate::facade::PartChange,
        styles: &crate::facade::PartChange,
        members: &[crate::facade::MemberChange],
    ) -> Result<Self> {
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
        for (path, change) in [("meta.xml", metadata), ("styles.xml", styles)] {
            match change {
                crate::facade::PartChange::Set(xml) => {
                    let compact = crate::codec::compact_for_publication(xml)?;
                    writer.add_file(path, compact.as_bytes())?;
                },
                crate::facade::PartChange::Keep if self.0.package.package().has_file(path)? => {
                    let source = XmlSourcePart::load(self.0.package.package(), path)?;
                    XmlSplicePublication::new(source).publish(&mut writer)?;
                },
                crate::facade::PartChange::Keep | crate::facade::PartChange::Remove => {},
            }
        }
        for change in members {
            if let Some(bytes) = &change.after {
                writer.add_file_with_media_type(&change.path, bytes, &change.media_type)?;
            }
        }
        let mut excluded_paths = vec![
            "content.xml".to_string(),
            "meta.xml".to_string(),
            "styles.xml".to_string(),
        ];
        excluded_paths.extend(members.iter().map(|change| change.path.clone()));
        writer.copy_auxiliary_files_from_except(self.0.package.package(), &excluded_paths, &[])?;
        Self::from_bytes(writer.finish_to_bounded_bytes()?)
    }
}
