//! Immutable chart package ownership and bounded content-part rebuilding.

use crate::FlatChart;
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.chart";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

struct State {
    package: Package,
    content: FlatChart,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(Package::open(path, MIMETYPE, "", "ODC")?)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // The chart reader performs namespace-aware content validation after
        // the package MIME check; a lexical body marker would reject valid
        // producer documents that use a different namespace prefix.
        Self::from_package(Package::from_bytes(bytes, MIMETYPE, "", "ODC")?)
    }

    fn from_package(package: Package) -> Result<Self> {
        let content = FlatChart::from_content_xml(package.content_xml().as_bytes().to_vec())?;
        Ok(Self(Arc::new(State { package, content })))
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

    pub(crate) fn content_snapshot(&self) -> FlatChart {
        self.0.content.clone()
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
                "ODC package edits refuse signed packages".to_string(),
            ));
        }
        compact_xml::validate(content.as_bytes()).map_err(Error::from)?;
        crate::codec::validate(content)?;
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
        if Path::new(&path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
            && path != "META-INF/manifest.xml"
        {
            compact_xml::validate(&archive.get_file(&path)?).map_err(Error::from)?;
        }
    }
    Ok(())
}
