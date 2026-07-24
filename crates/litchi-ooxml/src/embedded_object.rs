//! Inert discovery of OOXML embedded-object and embedded-package parts.

use crate::error::{OoxmlError, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};
use std::fmt;

/// Maximum number of embedded-part relationships returned from one package.
pub const MAX_EMBEDDED_RELATIONSHIPS: usize = 1024;

const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

/// Normative OOXML embedded-part relationship family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EmbeddedPartKind {
    /// An Embedded Object Part (§15.2.10).
    OleObject,
    /// An Embedded Package Part (§15.2.11).
    Package,
}

/// A borrowed internal embedded payload.
///
/// This is intentionally only an OPC byte view. It never parses OLE/CFB,
/// opens nested packages, sniffs formats, activates objects, or performs I/O.
#[derive(Clone, Copy)]
pub struct EmbeddedPayload<'a> {
    part: &'a dyn Part,
}

impl EmbeddedPayload<'_> {
    /// Absolute OPC part name of the payload.
    #[inline]
    pub fn part_name(&self) -> &PackURI {
        self.part.partname()
    }

    /// Declared OPC content type without format sniffing.
    #[inline]
    pub fn content_type(&self) -> &str {
        self.part.content_type()
    }

    /// Original payload bytes held by the OPC package.
    #[inline]
    pub fn bytes(&self) -> &[u8] {
        self.part.blob()
    }
}

impl fmt::Debug for EmbeddedPayload<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("EmbeddedPayload")
            .field("part_name", self.part_name())
            .field("content_type", &self.content_type())
            .field("byte_len", &self.bytes().len())
            .finish()
    }
}

/// Internal bytes or an inert external relationship target.
#[derive(Debug, Clone, Copy)]
pub enum EmbeddedTarget<'a> {
    Internal(EmbeddedPayload<'a>),
    External { target: &'a str },
}

/// One explicit relationship occurrence referencing an embedded part.
#[derive(Debug, Clone, Copy)]
pub struct EmbeddedPart<'a> {
    source_part_name: &'a PackURI,
    relationship_id: &'a str,
    kind: EmbeddedPartKind,
    target: EmbeddedTarget<'a>,
}

impl<'a> EmbeddedPart<'a> {
    #[inline]
    pub fn source_part_name(&self) -> &'a PackURI {
        self.source_part_name
    }

    #[inline]
    pub fn relationship_id(&self) -> &'a str {
        self.relationship_id
    }

    #[inline]
    pub fn kind(&self) -> EmbeddedPartKind {
        self.kind
    }

    #[inline]
    pub fn target(&self) -> EmbeddedTarget<'a> {
        self.target
    }
}

/// Discover embedded parts solely through normative explicit relationships.
pub(crate) fn discover_embedded_parts(package: &OpcPackage) -> Result<Vec<EmbeddedPart<'_>>> {
    for relationship in package.rels().iter() {
        if embedded_kind(relationship.reltype()).is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "package-level embedded relationship '{}' has no normative source part",
                relationship.r_id()
            )));
        }
    }

    let mut embedded = Vec::new();
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            let Some(kind) = embedded_kind(relationship.reltype()) else {
                continue;
            };
            if embedded.len() >= MAX_EMBEDDED_RELATIONSHIPS {
                return Err(OoxmlError::InvalidFormat(format!(
                    "embedded relationship count exceeds {MAX_EMBEDDED_RELATIONSHIPS}"
                )));
            }
            if !is_allowed_source(kind, source.content_type()) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "{} is not a normative source for {:?} relationship '{}'",
                    source.partname().as_str(),
                    kind,
                    relationship.r_id()
                )));
            }

            let target = if relationship.is_external() {
                EmbeddedTarget::External {
                    target: relationship.target_ref(),
                }
            } else {
                let target_name = relationship.target_partname().map_err(|error| {
                    OoxmlError::InvalidFormat(format!(
                        "invalid embedded target from {} relationship '{}': {error}",
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;
                let part = package.get_part(&target_name).map_err(|error| {
                    OoxmlError::PartNotFound(format!(
                        "embedded target '{}' from {} relationship '{}': {error}",
                        target_name.as_str(),
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;
                validate_payload_relationships(part)?;
                EmbeddedTarget::Internal(EmbeddedPayload { part })
            };

            embedded.push(EmbeddedPart {
                source_part_name: source.partname(),
                relationship_id: relationship.r_id(),
                kind,
                target,
            });
        }
    }

    embedded.sort_unstable_by(|left, right| {
        left.source_part_name
            .as_str()
            .cmp(right.source_part_name.as_str())
            .then_with(|| left.relationship_id.cmp(right.relationship_id))
    });
    Ok(embedded)
}

fn embedded_kind(relationship: &str) -> Option<EmbeddedPartKind> {
    match relationship {
        relationship_type::OLE_OBJECT | relationship_type::STRICT_OLE_OBJECT => {
            Some(EmbeddedPartKind::OleObject)
        },
        relationship_type::PACKAGE | relationship_type::STRICT_PACKAGE => {
            Some(EmbeddedPartKind::Package)
        },
        _ => None,
    }
}

fn is_allowed_source(kind: EmbeddedPartKind, source_content_type: &str) -> bool {
    let common = matches!(
        source_content_type,
        content_type::WML_COMMENTS
            | content_type::WML_ENDNOTES
            | content_type::WML_FOOTER
            | content_type::WML_FOOTNOTES
            | content_type::WML_HEADER
            | content_type::WML_DOCUMENT_MAIN
            | content_type::SML_WORKSHEET
            | content_type::PML_HANDOUT_MASTER
            | content_type::PML_NOTES_SLIDE
            | content_type::PML_NOTES_MASTER
            | content_type::PML_SLIDE
            | content_type::PML_SLIDE_LAYOUT
            | content_type::PML_SLIDE_MASTER
    );
    common || (kind == EmbeddedPartKind::Package && source_content_type == CHART_CONTENT_TYPE)
}

fn validate_payload_relationships(payload: &dyn Part) -> Result<()> {
    for relationship in payload.rels().iter() {
        if !matches!(
            relationship.reltype(),
            relationship_type::HYPERLINK | relationship_type::STRICT_HYPERLINK
        ) {
            return Err(OoxmlError::InvalidFormat(format!(
                "embedded target '{}' has forbidden relationship '{}' of type '{}'",
                payload.partname().as_str(),
                relationship.r_id(),
                relationship.reltype()
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::{BlobPart, Part};

    const POI_DOCX_OLE: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/document/EmbeddedDocument.docx");
    const LO_DOCX_PACKAGE: &[u8] = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/ooxmlimport/data/tdf108545_embeddedDocxIcon.docx"
    );
    const LO_DOCX_BAD_COMPOUND: &[u8] = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/ooxmlimport/data/tdf119039_bad_embedded_compound.docx"
    );
    const POI_DOCX_RECURSIVE: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/integration/test_recursive_embedded.docx");
    const POI_XLSX_TWO: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/spreadsheet/WithEmbeded.xlsx");
    const LO_PPTX_PACKAGE: &[u8] =
        include_bytes!("../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/ole.pptx");
    const POI_PPTX_MIXED: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/slideshow/bug62513.pptx");

    #[test]
    fn inventories_all_seven_real_world_fixtures_without_opening_payloads() {
        assert_fixture(POI_DOCX_OLE, 1, 1, 0);
        assert_fixture(LO_DOCX_PACKAGE, 1, 0, 1);
        assert_fixture(LO_DOCX_BAD_COMPOUND, 1, 1, 0);
        assert_fixture(POI_DOCX_RECURSIVE, 1, 1, 0);
        assert_fixture(POI_XLSX_TWO, 2, 2, 0);
        assert_fixture(LO_PPTX_PACKAGE, 1, 0, 1);
        assert_fixture(POI_PPTX_MIXED, 9, 4, 5);
    }

    #[test]
    fn preserves_malformed_and_recursive_payload_bytes_opaquely() {
        let malformed = OpcPackage::from_bytes(LO_DOCX_BAD_COMPOUND).unwrap();
        let entries = discover_embedded_parts(&malformed).unwrap();
        let EmbeddedTarget::Internal(payload) = entries[0].target() else {
            panic!("expected internal malformed payload")
        };
        assert_eq!(payload.bytes().len(), 2560);

        let recursive = OpcPackage::from_bytes(POI_DOCX_RECURSIVE).unwrap();
        let entries = discover_embedded_parts(&recursive).unwrap();
        let EmbeddedTarget::Internal(payload) = entries[0].target() else {
            panic!("expected internal recursive payload")
        };
        assert_eq!(&payload.bytes()[..8], b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1");
    }

    #[test]
    fn accepts_strict_relationships_and_returns_external_targets_inertly() {
        for (relationship, expected) in [
            (
                relationship_type::STRICT_OLE_OBJECT,
                EmbeddedPartKind::OleObject,
            ),
            (relationship_type::STRICT_PACKAGE, EmbeddedPartKind::Package),
        ] {
            let package =
                synthetic_package(relationship, true, true, content_type::WML_DOCUMENT_MAIN);
            let entries = discover_embedded_parts(&package).unwrap();
            assert_eq!(entries[0].kind(), expected);
            assert!(matches!(
                entries[0].target(),
                EmbeddedTarget::External {
                    target: "https://example.invalid/payload"
                }
            ));
        }
    }

    #[test]
    fn rejects_missing_targets_invalid_sources_and_package_root_relationships() {
        let missing = synthetic_package(
            relationship_type::OLE_OBJECT,
            false,
            false,
            content_type::WML_DOCUMENT_MAIN,
        );
        assert!(discover_embedded_parts(&missing).is_err());

        let invalid_source = synthetic_package(
            relationship_type::OLE_OBJECT,
            true,
            false,
            content_type::WML_STYLES,
        );
        assert!(discover_embedded_parts(&invalid_source).is_err());

        let mut root = OpcPackage::new();
        root.rels_mut().add_relationship(
            relationship_type::PACKAGE.into(),
            "payload.bin".into(),
            "rId1".into(),
            true,
        );
        assert!(discover_embedded_parts(&root).is_err());
    }

    #[test]
    fn ignores_orphans_preserves_duplicate_occurrences_and_sorts_deterministically() {
        let mut orphan = OpcPackage::new();
        orphan.add_part(Box::new(payload_part("/word/embeddings/orphan.bin")));
        assert!(discover_embedded_parts(&orphan).unwrap().is_empty());

        let mut package = OpcPackage::new();
        package.add_part(Box::new(payload_part("/word/embeddings/payload.bin")));
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        source.rels_mut().add_relationship(
            relationship_type::PACKAGE.into(),
            "embeddings/payload.bin".into(),
            "rId9".into(),
            false,
        );
        source.rels_mut().add_relationship(
            relationship_type::OLE_OBJECT.into(),
            "embeddings/payload.bin".into(),
            "rId2".into(),
            false,
        );
        package.add_part(Box::new(source));
        let entries = discover_embedded_parts(&package).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].relationship_id(), "rId2");
        assert_eq!(entries[1].relationship_id(), "rId9");
        let EmbeddedTarget::Internal(first) = entries[0].target() else {
            panic!("expected first internal target")
        };
        let EmbeddedTarget::Internal(second) = entries[1].target() else {
            panic!("expected second internal target")
        };
        assert_eq!(first.part_name(), second.part_name());
    }

    #[test]
    fn rejects_forbidden_payload_relationships_and_relationship_limit() {
        let mut forbidden = OpcPackage::new();
        let mut payload = payload_part("/word/embeddings/payload.bin");
        payload.rels_mut().add_relationship(
            relationship_type::IMAGE.into(),
            "image.png".into(),
            "rId1".into(),
            false,
        );
        forbidden.add_part(Box::new(payload));
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        source.rels_mut().add_relationship(
            relationship_type::OLE_OBJECT.into(),
            "embeddings/payload.bin".into(),
            "rId1".into(),
            false,
        );
        forbidden.add_part(Box::new(source));
        assert!(discover_embedded_parts(&forbidden).is_err());

        let mut limited = OpcPackage::new();
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        for index in 0..=MAX_EMBEDDED_RELATIONSHIPS {
            source.rels_mut().add_relationship(
                relationship_type::OLE_OBJECT.into(),
                format!("https://example.invalid/{index}"),
                format!("rId{index}"),
                true,
            );
        }
        limited.add_part(Box::new(source));
        assert!(discover_embedded_parts(&limited).is_err());
    }

    fn assert_fixture(bytes: &[u8], total: usize, ole: usize, packages: usize) {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let entries = discover_embedded_parts(&package).unwrap();
        assert_eq!(entries.len(), total);
        assert_eq!(
            entries
                .iter()
                .filter(|entry| entry.kind() == EmbeddedPartKind::OleObject)
                .count(),
            ole
        );
        assert_eq!(
            entries
                .iter()
                .filter(|entry| entry.kind() == EmbeddedPartKind::Package)
                .count(),
            packages
        );
        assert!(
            entries
                .iter()
                .all(|entry| matches!(entry.target(), EmbeddedTarget::Internal(_)))
        );
    }

    fn synthetic_package(
        relationship: &str,
        external: bool,
        include_target: bool,
        source_content_type: &str,
    ) -> OpcPackage {
        let mut package = OpcPackage::new();
        let mut source = source_part("/word/document.xml", source_content_type);
        source.rels_mut().add_relationship(
            relationship.into(),
            if external {
                "https://example.invalid/payload"
            } else {
                "embeddings/payload.bin"
            }
            .into(),
            "rId1".into(),
            external,
        );
        package.add_part(Box::new(source));
        if include_target && !external {
            package.add_part(Box::new(payload_part("/word/embeddings/payload.bin")));
        }
        package
    }

    fn source_part(name: &str, content_type: &str) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).unwrap(),
            content_type.into(),
            b"<source/>".to_vec(),
        )
    }

    fn payload_part(name: &str) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).unwrap(),
            "application/octet-stream".into(),
            b"opaque payload".to_vec(),
        )
    }
}
