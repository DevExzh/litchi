//! Immutable web-template package ownership and bounded content replacement.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
};
use std::{ops::Range, path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-web";
// The common package shell requires a cheap marker before this family applies
// its namespace-aware structural contract below. A literal prefix is not a
// valid ODF namespace check, so it must remain deliberately non-semantic.
const PRELIMINARY_XML_MARKER: &str = "<";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

struct State {
    package: Package,
    paragraphs: Vec<crate::paragraph::Paragraph>,
    replacement_sites: Vec<Option<Range<usize>>>,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
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

    fn from_package(package: Package) -> Result<Self> {
        let sites = crate::codec::paragraphs_with_sites(package.content_xml())?;
        let mut paragraphs = Vec::new();
        let mut replacement_sites = Vec::new();
        paragraphs
            .try_reserve(sites.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph snapshot",
                source,
            })?;
        replacement_sites
            .try_reserve(sites.len())
            .map_err(|source| Error::Allocation {
                resource: "OTH paragraph edit sites",
                source,
            })?;
        for site in sites {
            paragraphs.push(site.value);
            replacement_sites.push(site.replacement);
        }
        Ok(Self(Arc::new(State {
            package,
            paragraphs,
            replacement_sites,
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

    pub(crate) fn replacement_site(&self, index: usize) -> Option<&Range<usize>> {
        self.0.replacement_sites.get(index).and_then(Option::as_ref)
    }

    pub(crate) fn rebuild_with_content(&self, content: &str) -> Result<Self> {
        ensure_compact_rewrite_source(self)?;
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
        for path in ["styles.xml", "meta.xml", "settings.xml"] {
            if self.0.package.package().has_file(path)? {
                writer.add_file(path, &self.0.package.package().get_file(path)?)?;
            }
        }
        writer.copy_auxiliary_files_from(self.0.package.package())?;
        Self::from_bytes(writer.finish_to_bounded_bytes()?)
    }
}

fn ensure_compact_rewrite_source(source: &Snapshot) -> Result<()> {
    let archive = source.0.package.package();
    for path in source.files()? {
        if path.ends_with(".xml") && path != "META-INF/manifest.xml" {
            compact_xml::validate(&archive.get_file(&path)?).map_err(Error::from)?;
        }
    }
    Ok(())
}
