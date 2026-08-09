//! Immutable database package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, family::Package},
};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = litchi_odf_common::constants::ODF_DATABASE;
const BODY_MARKER: &str = "<";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

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

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
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
                "ODB package edits refuse signed packages".to_string(),
            ));
        }
        compact_xml::validate(content.as_bytes()).map_err(Error::from)?;
        crate::codec::validate(content)?;
        let archive = self.0.package.package();
        let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
        writer.set_mimetype(MIMETYPE)?;
        writer.add_file("content.xml", content.as_bytes())?;
        for path in ["styles.xml", "meta.xml", "settings.xml"] {
            if archive.has_file(path)? {
                writer.add_file(path, &archive.get_file(path)?)?;
            }
        }
        writer.copy_auxiliary_files_from(archive)?;
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
