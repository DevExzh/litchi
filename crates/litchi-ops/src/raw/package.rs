//! The published package of a handle, as parts: OPC for OOXML, the ZIP
//! members of a packaged ODF document.

use litchi_core::FileFormat;

use super::{Part, PartInfo, no_parts};
use crate::{Code, Error, File, Result};

/// A handle's published package.
#[cfg_attr(
    all(
        any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"),
        any(feature = "odt", feature = "ods", feature = "odp", odf_objects)
    ),
    expect(
        clippy::large_enum_variant,
        reason = "This transient stack owner keeps the full ODF policy without another heap allocation"
    )
)]
pub(super) enum Container {
    #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
    Opc {
        format: FileFormat,
        package: Box<litchi_opc::OpcPackage>,
    },
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    Odf {
        bytes: Vec<u8>,
        options: litchi_odf_common::OpenOptions,
    },
}

impl Container {
    /// Publish `file` and index the package.
    pub(super) fn of(file: &File) -> Result<Self> {
        let Some(format) = file.format() else {
            return Err(no_parts(file));
        };
        if is_opc(format) {
            #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
            {
                let bytes = file.to_bytes()?;
                return Ok(Self::Opc {
                    format,
                    package: Box::new(open_opc(bytes, file.ooxml_read_limits())?),
                });
            }
        }
        if is_odf(format) {
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            {
                let options = file.odf_read_limits();
                let bytes = file.to_bytes()?;
                if !bytes.starts_with(b"PK") {
                    // A flat ODF document is one XML file, not a package.
                    return Err(no_parts(file));
                }
                return Ok(Self::Odf { bytes, options });
            }
        }
        Err(no_parts(file))
    }

    /// The parts, in name order.
    pub(super) fn parts(&self) -> Result<Vec<PartInfo>> {
        match self {
            #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
            Self::Opc { package, .. } => {
                let mut parts: Vec<PartInfo> = package
                    .iter_parts()
                    .map(|part| PartInfo {
                        name: part.partname().as_str().to_owned(),
                        content_type: part.content_type().to_owned(),
                    })
                    .collect();
                parts.sort_unstable_by(|left, right| left.name.cmp(&right.name));
                Ok(parts)
            },
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            Self::Odf { bytes, options } => {
                let archive = litchi_odf_common::package::Archive::new(
                    bytes,
                    archive_limits(options.source_limits_value()),
                )?;
                let manifest =
                    litchi_odf_common::package::read_manifest(&archive, options.xml_limits_value())
                        .ok();
                let mut parts: Vec<PartInfo> = archive
                    .file_names()
                    .filter(|name| !name.ends_with('/'))
                    .map(|name| PartInfo {
                        name: format!("/{name}"),
                        content_type: manifest
                            .as_ref()
                            .and_then(|manifest| manifest.get_media_type(name))
                            .unwrap_or("")
                            .to_owned(),
                    })
                    .collect();
                parts.sort_unstable_by(|left, right| left.name.cmp(&right.name));
                Ok(parts)
            },
        }
    }

    /// Read the part `spec` names.
    pub(super) fn read(&self, spec: &str) -> Result<Part> {
        match self {
            #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
            Self::Opc { format, package } => {
                if let Some(part) = relationships(package, spec) {
                    return Ok(part);
                }
                let name = super::parts::resolve(package, *format, spec)?;
                let part = package.get_part(&name)?;
                Ok(Part {
                    name: part.partname().as_str().to_owned(),
                    content_type: part.content_type().to_owned(),
                    bytes: part.blob().to_vec(),
                })
            },
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            Self::Odf { bytes, options } => {
                let archive = litchi_odf_common::package::Archive::new(
                    bytes,
                    archive_limits(options.source_limits_value()),
                )?;
                let member = odf_member(&archive, spec)?;
                let manifest =
                    litchi_odf_common::package::read_manifest(&archive, options.xml_limits_value())
                        .ok();
                let content_type = manifest
                    .as_ref()
                    .and_then(|manifest| manifest.get_media_type(&member))
                    .unwrap_or(if member == "mimetype" {
                        "text/plain"
                    } else {
                        ""
                    })
                    .to_owned();
                Ok(Part {
                    bytes: archive.read(&member)?,
                    name: format!("/{member}"),
                    content_type,
                })
            },
        }
    }
}

/// Convert the retained ODF source admission profile to the archive codec's
/// format-owned catalog profile without widening any selected dimension.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn archive_limits(
    limits: litchi_odf_common::SourceLimits,
) -> litchi_odf_common::package::ArchiveLimits {
    litchi_odf_common::package::ArchiveLimits::new(
        limits.max_members(),
        limits.max_archive_member_name_bytes(),
        limits.max_archive_metadata_bytes(),
        limits.max_archive_compressed_bytes(),
        limits.max_entry_bytes(),
        limits.max_aggregate_bytes(),
    )
}

const fn is_opc(format: FileFormat) -> bool {
    matches!(
        format,
        FileFormat::Docx | FileFormat::Xlsx | FileFormat::Pptx | FileFormat::Xlsb
    )
}

const fn is_odf(format: FileFormat) -> bool {
    matches!(
        format,
        FileFormat::Odt
            | FileFormat::Ods
            | FileFormat::Odp
            | FileFormat::Odg
            | FileFormat::Odc
            | FileFormat::Odi
            | FileFormat::Odb
            | FileFormat::Odf
            | FileFormat::Oth
            | FileFormat::Odm
    )
}

/// A relationships member (`/_rels/.rels`, `/word/_rels/document.xml.rels`):
/// the package keeps relationships apart from parts, so the member is
/// serialized from its parsed relationships.
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
fn relationships(package: &litchi_opc::OpcPackage, spec: &str) -> Option<Part> {
    const RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
    let lower = spec.to_ascii_lowercase();
    if !lower.ends_with(".rels") {
        return None;
    }
    let rels = if lower == "/_rels/.rels" {
        package.rels()
    } else {
        let (directory, file) = spec.rsplit_once("/_rels/")?;
        let source =
            litchi_opc::PackURI::new(format!("{directory}/{}", file.strip_suffix(".rels")?))
                .ok()?;
        package.part_metadata(&source)?.rels()
    };
    Some(Part {
        name: spec.to_owned(),
        content_type: RELATIONSHIPS.to_owned(),
        bytes: rels.to_xml().into_bytes(),
    })
}

/// Index published OPC bytes without copying them.
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
pub(super) fn open_opc(
    bytes: Vec<u8>,
    limits: litchi_opc::SourceLimits,
) -> Result<litchi_opc::OpcPackage> {
    litchi_opc::OpcPackage::from_shared_vec(std::sync::Arc::new(bytes), limits).map_err(Error::from)
}

/// The ODF member `spec` names: `/content.xml`, or an alias.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn odf_member(archive: &litchi_odf_common::package::Archive<'_>, spec: &str) -> Result<String> {
    let Some(name) = spec.strip_prefix('/') else {
        return Err(Error::new(Code::InvalidPath, "a part name starts with '/'"));
    };
    if archive.contains(name) {
        return Ok(name.to_owned());
    }
    const ALIASES: [(&str, &str); 6] = [
        ("content", "content.xml"),
        ("styles", "styles.xml"),
        ("meta", "meta.xml"),
        ("settings", "settings.xml"),
        ("manifest", "META-INF/manifest.xml"),
        ("mimetype", "mimetype"),
    ];
    ALIASES
        .iter()
        .find(|(alias, _)| name.eq_ignore_ascii_case(alias))
        .map(|(_, member)| *member)
        .filter(|member| archive.contains(member))
        .map(str::to_owned)
        .ok_or_else(part_not_found)
}

/// `path_not_found` for a part name or alias the package does not hold.
pub(super) fn part_not_found() -> Error {
    Error::new(Code::PathNotFound, "the package has no such part").with_suggestion(
        "list the part names with `parts` (CLI: `litchi raw <file>`), or use an alias such as /document, /styles, /slide[1] or /Sheet1",
    )
}
