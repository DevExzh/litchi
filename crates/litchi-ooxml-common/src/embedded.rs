//! Inert inventory of embedded-object and embedded-package relationships.
//!
//! This module only traverses the OPC relationship graph and lends the bytes
//! already owned by the package. It never sniffs, parses, opens, activates,
//! renders, executes, or performs I/O for an embedded payload.

use crate::{Error, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};
use std::collections::HashSet;
use std::fmt;

const DEFAULT_RELATIONSHIPS: usize = 1_024;
const DEFAULT_PAYLOAD_RELATIONSHIPS: usize = 1_024;

// [MS-XLSB] File Structure, sections 2.1.7.36 and 2.1.7.37.
const XLSB_DIALOG_SHEET: &str = "application/vnd.ms-excel.dialogsheet";
const XLSB_EXTERNAL_LINK: &str = "application/vnd.ms-excel.externalLink";
const XLSB_MACRO_SHEET: &str = "application/vnd.ms-excel.macrosheet";
const XLSB_WORKSHEET: &str = "application/vnd.ms-excel.worksheet";

/// Normative OOXML embedded-part relationship family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// An Embedded Object Part (ISO/IEC 29500-1 section 15.2.10).
    Object,
    /// An Embedded Package Part (ISO/IEC 29500-1 section 15.2.11).
    Package,
}

/// Resource budgets for an embedded-part inventory.
///
/// [`Limits::default`] is the safe general-purpose policy. Callers that know
/// their workload may explicitly tighten or loosen either independent budget
/// with [`scan_with`]. Payload relationships are charged once per unique
/// internal target, even when several source relationships reference it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum embedded-object or embedded-package relationship occurrences.
    pub relationships: usize,
    /// Maximum aggregate relationships on unique internal payload parts.
    pub payload_relationships: usize,
}

impl Limits {
    /// Return the safe general-purpose limits.
    #[inline]
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            relationships: DEFAULT_RELATIONSHIPS,
            payload_relationships: DEFAULT_PAYLOAD_RELATIONSHIPS,
        }
    }
}

impl Default for Limits {
    #[inline]
    fn default() -> Self {
        Self::standard()
    }
}

/// A borrowed internal embedded payload.
///
/// The view retains no independent byte allocation and cannot outlive its OPC
/// package.
#[derive(Clone, Copy)]
pub struct Payload<'a> {
    part: &'a dyn Part,
}

impl<'a> Payload<'a> {
    /// Absolute OPC part name of the payload.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a PackURI {
        self.part.partname()
    }

    /// Declared OPC content type, without format sniffing.
    #[inline]
    #[must_use]
    pub fn content_type(&self) -> &'a str {
        self.part.content_type()
    }

    /// Original payload bytes held by the OPC package.
    #[inline]
    #[must_use]
    pub fn bytes(&self) -> &'a [u8] {
        self.part.blob()
    }
}

impl fmt::Debug for Payload<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Payload")
            .field("part", self.part())
            .field("content_type", &self.content_type())
            .field("byte_len", &self.bytes().len())
            .finish()
    }
}

/// Internal payload bytes or an inert external relationship target.
#[derive(Debug, Clone, Copy)]
pub enum Target<'a> {
    /// A package-owned payload exposed as a borrowed byte view.
    Internal(Payload<'a>),
    /// An external target retained verbatim and never contacted.
    External(&'a str),
}

/// One explicit relationship occurrence referencing an embedded part.
#[derive(Debug, Clone, Copy)]
pub struct Entry<'a> {
    source: &'a PackURI,
    id: &'a str,
    kind: Kind,
    target: Target<'a>,
}

impl<'a> Entry<'a> {
    /// Source part that owns the relationship.
    #[inline]
    #[must_use]
    pub fn source(&self) -> &'a PackURI {
        self.source
    }

    /// Relationship identifier within [`Self::source`].
    #[inline]
    #[must_use]
    pub fn id(&self) -> &'a str {
        self.id
    }

    /// Embedded relationship family.
    #[inline]
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Borrowed internal payload or inert external target.
    #[inline]
    #[must_use]
    pub fn target(&self) -> Target<'a> {
        self.target
    }
}

/// Inventory embedded parts with the safe general-purpose resource policy.
#[inline]
pub fn scan(package: &OpcPackage) -> Result<Vec<Entry<'_>>> {
    scan_with(package, &Limits::default())
}

/// Inventory embedded parts with explicit resource budgets.
///
/// Results are ordered by source part name and then relationship identifier,
/// independent of the package's internal hash-map order. Duplicate internal
/// targets remain distinct entries, while their payload relationship graph is
/// validated and charged exactly once under its canonical package part name.
pub fn scan_with<'a>(package: &'a OpcPackage, limits: &Limits) -> Result<Vec<Entry<'a>>> {
    for relationship in package.rels().iter() {
        if kind(relationship.reltype()).is_some() {
            return Err(Error::Relationship(format!(
                "package-level embedded relationship '{}' has no normative source part",
                relationship.r_id()
            )));
        }
    }

    let mut entries = Vec::new();
    let mut relationship_count = 0usize;
    let mut payload_relationship_count = 0usize;
    let mut validated_targets = HashSet::new();

    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            let Some(kind) = kind(relationship.reltype()) else {
                continue;
            };
            charge(
                &mut relationship_count,
                1,
                limits.relationships,
                "embedded relationships",
            )?;
            if !is_allowed_source(kind, source.content_type()) {
                return Err(Error::Relationship(format!(
                    "{} is not a normative source for {kind:?} relationship '{}'",
                    source.partname().as_str(),
                    relationship.r_id()
                )));
            }

            let target = if relationship.is_external() {
                Target::External(relationship.target_ref())
            } else {
                if relationship.target_query().is_some() || relationship.target_fragment().is_some()
                {
                    return Err(Error::Relationship(format!(
                        "internal embedded target from {} relationship '{}' cannot contain a query or fragment",
                        source.partname().as_str(),
                        relationship.r_id()
                    )));
                }
                let target_name = relationship.target_partname().map_err(|error| {
                    Error::Relationship(format!(
                        "invalid embedded target from {} relationship '{}': {error}",
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;
                let part = package.get_part(&target_name).map_err(|error| {
                    Error::Missing(format!(
                        "embedded target '{}' from {} relationship '{}': {error}",
                        target_name.as_str(),
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;

                // `get_part` resolves ASCII-case differences to the stored part.
                // Keying on that stored name therefore memoizes the canonical
                // target even when source relationships use different casing.
                if validated_targets.insert(part.partname()) {
                    validate_payload_relationships(
                        part,
                        &mut payload_relationship_count,
                        limits.payload_relationships,
                    )?;
                }
                Target::Internal(Payload { part })
            };

            entries.push(Entry {
                source: source.partname(),
                id: relationship.r_id(),
                kind,
                target,
            });
        }
    }

    entries.sort_unstable_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(right.id))
    });
    Ok(entries)
}

fn kind(relationship: &str) -> Option<Kind> {
    match relationship {
        relationship_type::OLE_OBJECT | relationship_type::STRICT_OLE_OBJECT => Some(Kind::Object),
        relationship_type::PACKAGE | relationship_type::STRICT_PACKAGE => Some(Kind::Package),
        _ => None,
    }
}

fn is_allowed_source(kind: Kind, source_content_type: &str) -> bool {
    let common = matches!(
        source_content_type,
        content_type::WML_COMMENTS
            | content_type::WML_ENDNOTES
            | content_type::WML_FOOTER
            | content_type::WML_FOOTNOTES
            | content_type::WML_HEADER
            | content_type::WML_DOCUMENT_MAIN
            | content_type::WML_TEMPLATE_MAIN
            | content_type::WML_DOCUMENT_MACRO_MAIN
            | content_type::WML_TEMPLATE_MACRO_MAIN
            | content_type::SML_WORKSHEET
            | content_type::PML_HANDOUT_MASTER
            | content_type::PML_NOTES_SLIDE
            | content_type::PML_NOTES_MASTER
            | content_type::PML_SLIDE
            | content_type::PML_SLIDE_LAYOUT
            | content_type::PML_SLIDE_MASTER
    );
    if common {
        return true;
    }

    matches!(
        (kind, source_content_type),
        (Kind::Object, XLSB_EXTERNAL_LINK)
            | (
                Kind::Object | Kind::Package,
                XLSB_DIALOG_SHEET | XLSB_MACRO_SHEET | XLSB_WORKSHEET
            )
            | (Kind::Package, content_type::DML_CHART)
    )
}

fn validate_payload_relationships(payload: &dyn Part, count: &mut usize, max: usize) -> Result<()> {
    for relationship in payload.rels().iter() {
        charge(count, 1, max, "embedded payload relationships")?;
        if !matches!(
            relationship.reltype(),
            relationship_type::HYPERLINK | relationship_type::STRICT_HYPERLINK
        ) {
            return Err(Error::Relationship(format!(
                "embedded target '{}' has forbidden relationship '{}' of type '{}'",
                payload.partname().as_str(),
                relationship.r_id(),
                relationship.reltype()
            )));
        }
    }
    Ok(())
}

fn charge(total: &mut usize, amount: usize, max: usize, resource: &'static str) -> Result<()> {
    let actual = total.checked_add(amount).ok_or(Error::Limit {
        resource,
        max,
        actual: usize::MAX,
    })?;
    if actual > max {
        return Err(Error::Limit {
            resource,
            max,
            actual,
        });
    }
    *total = actual;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::Relationships;
    use litchi_opc::part::BlobPart;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};

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
        let malformed = OpcPackage::from_bytes(LO_DOCX_BAD_COMPOUND).expect("fixture opens");
        let entries = scan(&malformed).expect("malformed payload remains opaque");
        let Target::Internal(payload) = entries[0].target() else {
            panic!("expected internal malformed payload")
        };
        assert_eq!(payload.bytes().len(), 2_560);

        let recursive = OpcPackage::from_bytes(POI_DOCX_RECURSIVE).expect("fixture opens");
        let entries = scan(&recursive).expect("recursive payload remains opaque");
        let Target::Internal(payload) = entries[0].target() else {
            panic!("expected internal recursive payload")
        };
        assert_eq!(&payload.bytes()[..8], b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1");
    }

    #[test]
    fn preserves_server_specific_object_content_types_without_sniffing() {
        let package = OpcPackage::from_bytes(POI_XLSX_TWO).expect("fixture opens");
        let entries = scan(&package).expect("fixture graph is valid");
        assert!(entries.iter().any(|entry| {
            entry.kind() == Kind::Object
                && matches!(
                    entry.target(),
                    Target::Internal(payload)
                        if payload.content_type() == "application/vnd.ms-excel"
                )
        }));
    }

    #[test]
    fn accepts_strict_relationships_and_returns_external_targets_inertly() {
        for (relationship, expected) in [
            (relationship_type::STRICT_OLE_OBJECT, Kind::Object),
            (relationship_type::STRICT_PACKAGE, Kind::Package),
        ] {
            let package =
                synthetic_package(relationship, true, true, content_type::WML_DOCUMENT_MAIN);
            let entries = scan(&package).expect("strict relationship is valid");
            assert_eq!(entries[0].kind(), expected);
            assert!(matches!(
                entries[0].target(),
                Target::External("https://example.invalid/payload")
            ));
        }
    }

    #[test]
    fn accepts_word_template_and_macro_main_sources() {
        for source in [
            content_type::WML_TEMPLATE_MAIN,
            content_type::WML_DOCUMENT_MACRO_MAIN,
            content_type::WML_TEMPLATE_MACRO_MAIN,
        ] {
            for relationship in [relationship_type::OLE_OBJECT, relationship_type::PACKAGE] {
                let package = synthetic_package(relationship, true, false, source);
                assert_eq!(scan(&package).expect("Word source is valid").len(), 1);
            }
        }
    }

    #[test]
    fn applies_kind_specific_xlsb_source_policy() {
        for source in [XLSB_DIALOG_SHEET, XLSB_MACRO_SHEET, XLSB_WORKSHEET] {
            for relationship in [relationship_type::OLE_OBJECT, relationship_type::PACKAGE] {
                let package = synthetic_package(relationship, true, false, source);
                assert_eq!(scan(&package).expect("XLSB source is valid").len(), 1);
            }
        }

        let object = synthetic_package(
            relationship_type::OLE_OBJECT,
            true,
            false,
            XLSB_EXTERNAL_LINK,
        );
        assert_eq!(
            scan(&object)
                .expect("XLSB external-link object is valid")
                .len(),
            1
        );
        let package =
            synthetic_package(relationship_type::PACKAGE, true, false, XLSB_EXTERNAL_LINK);
        assert!(scan(&package).is_err());
    }

    #[test]
    fn accepts_chart_packages_but_not_chart_objects() {
        let package = synthetic_package(
            relationship_type::PACKAGE,
            true,
            false,
            content_type::DML_CHART,
        );
        assert_eq!(scan(&package).expect("chart package is valid").len(), 1);
        let object = synthetic_package(
            relationship_type::OLE_OBJECT,
            true,
            false,
            content_type::DML_CHART,
        );
        assert!(scan(&object).is_err());
    }

    #[test]
    fn rejects_missing_targets_invalid_sources_and_package_root_relationships() {
        let missing = synthetic_package(
            relationship_type::OLE_OBJECT,
            false,
            false,
            content_type::WML_DOCUMENT_MAIN,
        );
        assert!(scan(&missing).is_err());

        let invalid_source = synthetic_package(
            relationship_type::OLE_OBJECT,
            true,
            false,
            content_type::WML_STYLES,
        );
        assert!(scan(&invalid_source).is_err());

        let mut root = OpcPackage::new();
        root.rels_mut().add_relationship(
            relationship_type::PACKAGE.into(),
            "payload.bin".into(),
            "rId1".into(),
            true,
        );
        assert!(scan(&root).is_err());
    }

    #[test]
    fn rejects_query_and_fragment_components_on_internal_targets() {
        for suffix in ["?version=1", "#payload"] {
            let mut package = OpcPackage::new();
            package.add_part(Box::new(payload_part("/word/embeddings/payload.bin")));
            let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
            source.rels_mut().add_relationship(
                relationship_type::OLE_OBJECT.into(),
                format!("embeddings/payload.bin{suffix}"),
                "rId1".into(),
                false,
            );
            package.add_part(Box::new(source));
            assert!(matches!(scan(&package), Err(Error::Relationship(_))));
        }
    }

    #[test]
    fn ignores_orphans_preserves_duplicate_occurrences_and_sorts_deterministically() {
        let mut orphan = OpcPackage::new();
        orphan.add_part(Box::new(payload_part("/word/embeddings/orphan.bin")));
        assert!(scan(&orphan).expect("orphan is inert").is_empty());

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
        let entries = scan(&package).expect("duplicate occurrences are valid");
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].id(), "rId2");
        assert_eq!(entries[1].id(), "rId9");
        let Target::Internal(first) = entries[0].target() else {
            panic!("expected first internal target")
        };
        let Target::Internal(second) = entries[1].target() else {
            panic!("expected second internal target")
        };
        assert_eq!(first.part(), second.part());
        assert!(std::ptr::eq(first.bytes(), second.bytes()));
    }

    #[test]
    fn rejects_forbidden_payload_relationships_and_default_relationship_limit() {
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
        assert!(scan(&forbidden).is_err());

        let mut limited = OpcPackage::new();
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        for index in 0..=Limits::default().relationships {
            source.rels_mut().add_relationship(
                relationship_type::OLE_OBJECT.into(),
                format!("https://example.invalid/{index}"),
                format!("rId{index}"),
                true,
            );
        }
        limited.add_part(Box::new(source));
        assert!(matches!(scan(&limited), Err(Error::Limit { .. })));
    }

    #[test]
    fn custom_limits_are_checked_and_payload_budget_is_aggregate() {
        let mut entries = OpcPackage::new();
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        for index in 1..=2 {
            source.rels_mut().add_relationship(
                relationship_type::OLE_OBJECT.into(),
                format!("https://example.invalid/{index}"),
                format!("rId{index}"),
                true,
            );
        }
        entries.add_part(Box::new(source));
        let one_entry = Limits {
            relationships: 1,
            ..Limits::default()
        };
        assert!(matches!(
            scan_with(&entries, &one_entry),
            Err(Error::Limit { .. })
        ));

        let mut payloads = OpcPackage::new();
        for index in 1..=2 {
            let name = format!("/word/embeddings/payload{index}.bin");
            let mut payload = payload_part(&name);
            payload.rels_mut().add_relationship(
                relationship_type::HYPERLINK.into(),
                format!("https://example.invalid/{index}"),
                "rIdLink".into(),
                true,
            );
            payloads.add_part(Box::new(payload));
        }
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        for index in 1..=2 {
            source.rels_mut().add_relationship(
                relationship_type::OLE_OBJECT.into(),
                format!("embeddings/payload{index}.bin"),
                format!("rId{index}"),
                false,
            );
        }
        payloads.add_part(Box::new(source));
        let one_payload_relationship = Limits {
            payload_relationships: 1,
            ..Limits::default()
        };
        assert!(matches!(
            scan_with(&payloads, &one_payload_relationship),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn duplicate_internal_targets_are_validated_and_charged_once() {
        let scans = Arc::new(AtomicUsize::new(0));
        let mut payload = CountingPart::new(
            payload_part("/word/embeddings/payload.bin"),
            Arc::clone(&scans),
        );
        payload.rels_mut().add_relationship(
            relationship_type::HYPERLINK.into(),
            "https://example.invalid".into(),
            "rIdLink".into(),
            true,
        );

        let mut package = OpcPackage::new();
        package.add_part(Box::new(payload));
        let mut source = source_part("/word/document.xml", content_type::WML_DOCUMENT_MAIN);
        for id in ["rId1", "rId2"] {
            source.rels_mut().add_relationship(
                relationship_type::OLE_OBJECT.into(),
                "embeddings/payload.bin".into(),
                id.into(),
                false,
            );
        }
        package.add_part(Box::new(source));

        let limits = Limits {
            payload_relationships: 1,
            ..Limits::default()
        };
        assert_eq!(
            scan_with(&package, &limits)
                .expect("duplicate target consumes one payload budget")
                .len(),
            2
        );
        // One call while the payload is visited as a potential source and one
        // while its own relationships are validated. Without target
        // memoization the second source occurrence would make a third call.
        assert_eq!(scans.load(Ordering::Relaxed), 2);
    }

    fn assert_fixture(bytes: &[u8], total: usize, objects: usize, packages: usize) {
        let package = OpcPackage::from_bytes(bytes).expect("fixture opens");
        let entries = scan(&package).expect("fixture graph is valid");
        assert_eq!(entries.len(), total);
        assert_eq!(
            entries
                .iter()
                .filter(|entry| entry.kind() == Kind::Object)
                .count(),
            objects
        );
        assert_eq!(
            entries
                .iter()
                .filter(|entry| entry.kind() == Kind::Package)
                .count(),
            packages
        );
        assert!(
            entries
                .iter()
                .all(|entry| matches!(entry.target(), Target::Internal(_)))
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
            PackURI::new(name).expect("valid test part name"),
            content_type.into(),
            b"<source/>".to_vec(),
        )
    }

    fn payload_part(name: &str) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).expect("valid test part name"),
            "application/octet-stream".into(),
            b"opaque payload".to_vec(),
        )
    }

    #[derive(Clone)]
    struct CountingPart {
        inner: BlobPart,
        relationship_scans: Arc<AtomicUsize>,
    }

    impl CountingPart {
        fn new(inner: BlobPart, relationship_scans: Arc<AtomicUsize>) -> Self {
            Self {
                inner,
                relationship_scans,
            }
        }
    }

    impl Part for CountingPart {
        fn partname(&self) -> &PackURI {
            self.inner.partname()
        }

        fn content_type(&self) -> &str {
            self.inner.content_type()
        }

        fn blob(&self) -> &[u8] {
            self.inner.blob()
        }

        fn blob_arc(&self) -> Arc<Vec<u8>> {
            self.inner.blob_arc()
        }

        fn set_blob(&mut self, blob: Vec<u8>) {
            self.inner.set_blob(blob);
        }

        fn rels(&self) -> &Relationships {
            self.relationship_scans.fetch_add(1, Ordering::Relaxed);
            self.inner.rels()
        }

        fn rels_mut(&mut self) -> &mut Relationships {
            self.inner.rels_mut()
        }
    }
}
