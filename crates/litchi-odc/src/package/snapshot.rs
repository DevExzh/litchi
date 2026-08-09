//! Immutable chart package ownership and bounded content-part rebuilding.

use crate::FlatChart;
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{
    compact_xml,
    core::{
        AuthoredXmlFragment, OwnedPackage, PackageWriter, XmlSourcePart, XmlSplicePublication,
        family::Package,
    },
};
use std::{fs, io::Write, path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.chart";
struct State {
    package: Package,
    content: FlatChart,
    limits: crate::Limits,
    resources: Vec<crate::Resource>,
    signed: bool,
    encrypted: bool,
}

pub(crate) struct ResourceReplacement<'a> {
    pub(crate) path: &'a str,
    pub(crate) media_type: &'a str,
    pub(crate) bytes: Option<&'a [u8]>,
}

#[derive(Clone, Copy)]
pub(crate) enum StylesReplacement<'a> {
    Unchanged,
    Replace(&'a str),
    Remove,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, crate::Limits::default())
    }

    pub(crate) fn open_with_limits(path: impl AsRef<Path>, limits: crate::Limits) -> Result<Self> {
        Self::from_bytes_with_limits(fs::read(path)?, limits)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, crate::Limits::default())
    }

    pub(crate) fn from_bytes_with_limits(bytes: Vec<u8>, limits: crate::Limits) -> Result<Self> {
        if bytes.len() > limits.max_package_bytes() {
            return Err(Error::InvalidFormat(
                "ODC package exceeds the caller-selected byte limit".into(),
            ));
        }
        // The chart reader performs namespace-aware content validation after
        // the package MIME check; a lexical body marker would reject valid
        // producer documents that use a different namespace prefix.
        Self::from_package(Package::from_bytes(bytes, MIMETYPE, "", "ODC")?, limits)
    }

    fn from_package(package: Package, limits: crate::Limits) -> Result<Self> {
        if let Some(styles) = package.styles_xml() {
            crate::codec::validate_styles(styles, limits)?;
        }
        let content =
            FlatChart::from_content_xml(package.content_xml().as_bytes().to_vec(), limits)?;
        let resources = scan_resources(&package, limits)?;
        let (signed, encrypted) = scan_security(&package)?;
        Ok(Self(Arc::new(State {
            package,
            content,
            limits,
            resources,
            signed,
            encrypted,
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

    pub(crate) fn content_snapshot(&self) -> FlatChart {
        self.0.content.clone()
    }

    pub(crate) fn resources(&self) -> &[crate::Resource] {
        &self.0.resources
    }

    pub(crate) fn is_signed(&self) -> bool {
        self.0.signed
    }

    pub(crate) fn is_encrypted(&self) -> bool {
        self.0.encrypted
    }

    pub(crate) fn resource_bytes(&self, index: usize) -> Result<Vec<u8>> {
        let resource =
            self.0.resources.get(index).ok_or_else(|| {
                Error::InvalidFormat("ODC resource selector is out of bounds".into())
            })?;
        self.0.package.package().get_file(resource.path())
    }

    pub(crate) fn limits(&self) -> crate::Limits {
        self.0.limits
    }

    pub(crate) fn rebuild(
        &self,
        content: &str,
        content_splice: Option<&crate::FlatChartPatch>,
        styles: StylesReplacement<'_>,
        replacements: &[ResourceReplacement<'_>],
    ) -> Result<Self> {
        if self.0.signed {
            return Err(Error::InvalidFormat(
                "ODC package edits refuse signed packages".to_string(),
            ));
        }
        if self.0.encrypted {
            return Err(Error::Unsupported(
                "ODC package edits refuse encrypted package members".to_string(),
            ));
        }
        let compact_limits =
            compact_xml::Limits::new(self.0.limits.max_content_bytes(), self.0.limits.max_depth())
                .map_err(Error::from)?;
        if content_splice.is_none() {
            compact_xml::validate_with_limits(content.as_bytes(), compact_limits)
                .map_err(Error::from)?;
        }
        crate::codec::validate(content)?;
        let mut writer = PackageWriter::new_bounded(self.0.limits.max_package_bytes());
        writer.set_mimetype(MIMETYPE)?;
        if let Some(patch) = content_splice {
            publish_content_splice(self.0.package.package(), patch, content, &mut writer)?;
        } else {
            writer.add_file("content.xml", content.as_bytes())?;
        }
        match styles {
            StylesReplacement::Unchanged => {
                if self.0.package.package().has_file("styles.xml")? {
                    publish_retained_xml(self.0.package.package(), "styles.xml", &mut writer)?;
                }
            },
            StylesReplacement::Replace(xml) => {
                compact_xml::validate_with_limits(xml.as_bytes(), compact_limits)
                    .map_err(Error::from)?;
                crate::codec::validate_styles(xml, self.0.limits)?;
                writer.add_file("styles.xml", xml.as_bytes())?;
            },
            StylesReplacement::Remove => {},
        }
        for path in ["meta.xml", "settings.xml"] {
            if self.0.package.package().has_file(path)? {
                publish_retained_xml(self.0.package.package(), path, &mut writer)?;
            }
        }
        let excluded = replacements
            .iter()
            .map(|replacement| replacement.path.to_string())
            .collect::<Vec<_>>();
        writer.copy_auxiliary_files_from_except(self.0.package.package(), &excluded, &[])?;
        for replacement in replacements {
            if let Some(bytes) = replacement.bytes {
                validate_authored_resource(
                    replacement.path,
                    replacement.media_type,
                    bytes,
                    compact_limits,
                )?;
                writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
            }
        }
        Self::from_bytes_with_limits(writer.finish_to_bounded_bytes()?, self.0.limits)
    }
}

fn publish_content_splice<W: Write>(
    source: &OwnedPackage,
    patch: &crate::FlatChartPatch,
    content: &str,
    writer: &mut PackageWriter<W>,
) -> Result<()> {
    if patch.target_bytes() != content.as_bytes() {
        return Err(Error::InvalidFormat(
            "ODC content splice target does not match the committed chart".into(),
        ));
    }
    let source_part = XmlSourcePart::load(source, "content.xml")?;
    if source_part.bytes() != patch.source_bytes() {
        return Err(Error::InvalidFormat(
            "ODC content splice has different package provenance".into(),
        ));
    }
    let mut publication = XmlSplicePublication::new(source_part.clone());
    for splice in patch.tag_splices()? {
        let proof = source_part.checked_range(splice.range, &splice.expected)?;
        let fragment = if splice.replacement.ends_with(b"/>") {
            AuthoredXmlFragment::markup(splice.replacement)?
        } else {
            AuthoredXmlFragment::start_tag(splice.replacement)?
        };
        publication.replace(proof, fragment)?;
    }
    publication.publish(writer)
}

fn publish_retained_xml<W: Write>(
    source: &OwnedPackage,
    path: &str,
    writer: &mut PackageWriter<W>,
) -> Result<()> {
    XmlSplicePublication::new(XmlSourcePart::load(source, path)?).publish(writer)
}

fn scan_resources(package: &Package, limits: crate::Limits) -> Result<Vec<crate::Resource>> {
    let archive = package.package().package()?;
    let mut resources = Vec::new();
    for path in archive.files()? {
        if path.ends_with('/')
            || matches!(
                path.as_str(),
                "mimetype"
                    | "content.xml"
                    | "styles.xml"
                    | "meta.xml"
                    | "settings.xml"
                    | "META-INF/manifest.xml"
            )
            || path.starts_with("META-INF/")
        {
            continue;
        }
        if resources.len() >= limits.max_resources() {
            return Err(Error::InvalidFormat(
                "ODC resource count exceeds the caller-selected limit".into(),
            ));
        }
        let bytes = archive.get_file(&path)?;
        resources.push(crate::Resource::new(
            path.clone(),
            archive.manifest().get_media_type(&path).map(str::to_owned),
            bytes.len(),
        ));
    }
    Ok(resources)
}

fn scan_security(package: &Package) -> Result<(bool, bool)> {
    let archive = package.package().package()?;
    let signed = archive.files()?.iter().any(|path| {
        path.strip_prefix("META-INF/")
            .is_some_and(|name| name.to_ascii_lowercase().contains("signatures"))
    });
    Ok((signed, archive.manifest().has_encrypted_entries()))
}

pub(crate) fn validate_authored_resource(
    path: &str,
    media_type: &str,
    bytes: &[u8],
    limits: compact_xml::Limits,
) -> Result<()> {
    let is_xml = Path::new(path)
        .extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        || media_type == "application/xml"
        || media_type == "text/xml"
        || media_type.ends_with("+xml");
    if is_xml {
        compact_xml::validate_with_limits(bytes, limits).map_err(Error::from)?;
    }
    Ok(())
}
