//! Immutable database package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::family::Package;
use std::{
    io::{Cursor, Write},
    path::Path,
    sync::Arc,
};

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
        // Opened producer documents are edited by byte-splicing only the
        // selected XML range.  Formatting whitespace in the unchanged source
        // is therefore lossless input, not generated output that needs the
        // fresh-authoring compactness gate.
        crate::codec::validate(content)?;
        Self::from_bytes(rebuild_archive(self.as_bytes(), content)?)
    }
}

fn rebuild_archive(source: &[u8], content: &str) -> Result<Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(source))
        .map_err(|error| Error::ZipError(error.to_string()))?;
    let mut output = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut output);
        for index in 0..archive.len() {
            let file = archive
                .by_index_raw(index)
                .map_err(|error| Error::ZipError(error.to_string()))?;
            if file.name() == "content.xml" {
                let mut options =
                    zip::write::SimpleFileOptions::default().compression_method(file.compression());
                if let Some(mode) = file.unix_mode() {
                    options = options.unix_permissions(mode);
                }
                if let Some(modified) = file.last_modified() {
                    options = options.last_modified_time(modified);
                }
                writer
                    .start_file("content.xml", options)
                    .and_then(|()| writer.write_all(content.as_bytes()).map_err(Into::into))
                    .map_err(|error: zip::result::ZipError| Error::ZipError(error.to_string()))?;
            } else {
                writer
                    .raw_copy_file(file)
                    .map_err(|error| Error::ZipError(error.to_string()))?;
            }
        }
        writer
            .finish()
            .map_err(|error| Error::ZipError(error.to_string()))?;
    }
    let bytes = output.into_inner();
    if bytes.len() > MAX_OUTPUT_BYTES {
        return Err(Error::InvalidFormat(
            "ODB rebuilt package exceeds the output limit".to_string(),
        ));
    }
    Ok(bytes)
}
