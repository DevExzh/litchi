//! Where the reader draws the line between an OPC part and archive junk.
//!
//! Each case builds a minimal package so the rule under test is the only thing
//! that varies, and asserts both halves of the contract: the package's real
//! content stays reachable, and anything the reader declined to model as a part
//! is still reported.

use litchi_opc::constants::content_type as ct;
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_opc::pkgreader::PackageReader;
use litchi_opc::{NonPartReason, OpcError, OpcPackage};
use soapberry_zip::office::StreamingArchiveWriter;

/// Content types manifest declaring only a default for the `xml` extension.
const XML_ONLY_MANIFEST: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#;
/// Content types manifest declaring defaults for both `rels` and `xml`.
const RELS_AND_XML_MANIFEST: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;

/// Build a ZIP archive from `(member name, contents)` pairs, in order.
fn archive(members: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    for (name, blob) in members {
        writer
            .write_stored(name, blob)
            .unwrap_or_else(|error| panic!("write {name}: {error}"));
    }
    writer.finish_to_bytes().expect("finish archive")
}

/// Parse an archive through the package reader.
fn read(bytes: &[u8]) -> Result<PackageReader, OpcError> {
    PackageReader::from_phys_reader(&PhysPkgReader::new(bytes)?)
}

/// Return the blob of the named part, or fail the test.
fn blob<'reader>(reader: &'reader PackageReader, partname: &str) -> &'reader [u8] {
    reader
        .iter_sparts()
        .find(|spart| spart.partname.as_str() == partname)
        .unwrap_or_else(|| panic!("package is missing {partname}"))
        .blob
        .as_slice()
}

#[test]
fn reports_unreferenced_untyped_member_without_losing_the_real_parts() {
    // An untyped item nothing refers to is archive junk, not a part
    // (ECMA-376 Part 2 §10.1.2.2), so the package still opens — but the reader
    // has to say the item was there.
    let bytes = archive(&[
        ("[Content_Types].xml", XML_ONLY_MANIFEST),
        ("custom/orphan.bin", b"orphan"),
        ("word/document.xml", b"<w:document/>"),
    ]);

    let physical = PhysPkgReader::new(&bytes).unwrap();
    let reader = PackageReader::from_phys_reader(&physical).unwrap();
    assert_eq!(blob(&reader, "/word/document.xml"), b"<w:document/>");
    assert!(
        !reader
            .iter_sparts()
            .any(|spart| spart.partname.as_str() == "/custom/orphan.bin")
    );

    let junk = reader.non_part_members();
    assert_eq!(junk.len(), 1);
    assert_eq!(junk[0].name(), "custom/orphan.bin");
    assert_eq!(junk[0].reason(), NonPartReason::UntypedAndUnreferenced);
    // The bytes are still reachable through the physical archive.
    assert_eq!(physical.archive().read(junk[0].name()).unwrap(), b"orphan");
}

#[test]
fn rejects_referenced_part_without_content_type_mapping() {
    // A part a relationship points at is a part, so the requirement that every
    // part carry a content type applies to it.
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="custom/orphan.bin"/></Relationships>"#;
    let bytes = archive(&[
        ("[Content_Types].xml", RELS_AND_XML_MANIFEST),
        ("_rels/.rels", relationships),
        ("custom/orphan.bin", b"orphan"),
    ]);

    assert!(matches!(
        read(&bytes),
        Err(OpcError::ContentTypeNotFound(_))
    ));
}

#[test]
fn reports_members_whose_names_cannot_denote_a_part() {
    // `[` and `]` are outside RFC 3986 pchar, so `[trash]/0000.dat` — which
    // Excel leaves behind — cannot be a part name (§9.1.1.1).
    let bytes = archive(&[
        ("[Content_Types].xml", XML_ONLY_MANIFEST),
        ("[trash]/0000.dat", b"junk"),
        ("xl/workbook.xml", b"<workbook/>"),
    ]);

    let reader = read(&bytes).expect("junk item must not fail the package");
    assert_eq!(blob(&reader, "/xl/workbook.xml"), b"<workbook/>");

    let junk = reader.non_part_members();
    assert_eq!(junk.len(), 1);
    assert_eq!(junk[0].name(), "[trash]/0000.dat");
    assert_eq!(junk[0].reason(), NonPartReason::UnmappablePartName);
}

#[test]
fn keeps_dangling_relationships_without_inventing_their_target_part() {
    // OPC states no rule that an internal target must resolve at load time, so
    // the package opens and the relationship stays visible on its source.
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="word/document.xml"/><Relationship Id="rId2" Type="urn:missing" Target="word/footer1.xml"/></Relationships>"#;
    let bytes = archive(&[
        ("[Content_Types].xml", RELS_AND_XML_MANIFEST),
        ("_rels/.rels", relationships),
        ("word/document.xml", b"<w:document/>"),
    ]);

    let package =
        OpcPackage::from_bytes(&bytes).expect("dangling target must not fail the package");
    let document = package
        .iter_parts()
        .find(|part| part.partname().as_str() == "/word/document.xml")
        .expect("present target must load");
    assert_eq!(document.blob(), b"<w:document/>");
    assert!(
        !package
            .iter_parts()
            .any(|part| part.partname().as_str() == "/word/footer1.xml"),
        "an absent target must not be fabricated as a part"
    );
    assert_eq!(
        package
            .rels()
            .iter()
            .find(|relationship| relationship.r_id() == "rId2")
            .map(|relationship| relationship.target_ref()),
        Some("word/footer1.xml")
    );
}

#[test]
fn accepts_the_content_types_stream_under_a_case_variant_name() {
    // Item-name comparison is ASCII case-insensitive (§9.1.1.2).
    let bytes = archive(&[
        ("[content_types].xml", XML_ONLY_MANIFEST),
        ("xl/workbook.xml", b"<workbook/>"),
    ]);

    let reader = read(&bytes).expect("case-variant manifest must resolve");
    assert_eq!(blob(&reader, "/xl/workbook.xml"), b"<workbook/>");
    assert!(reader.non_part_members().is_empty());
}

#[test]
fn types_relationship_parts_from_the_reserved_name_but_rejects_contradictions() {
    // §9.2 fixes the Relationships part content type, so an omitted mapping is
    // fine and a conflicting one is not.
    fn package_with_rels_default(rels_content_type: Option<&str>) -> Vec<u8> {
        let default = rels_content_type
            .map(|value| format!(r#"<Default Extension="rels" ContentType="{value}"/>"#))
            .unwrap_or_default();
        let manifest = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{default}<Default Extension="xml" ContentType="application/xml"/></Types>"#
        );
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="word/document.xml"/></Relationships>"#;
        archive(&[
            ("[Content_Types].xml", manifest.as_bytes()),
            ("_rels/.rels", relationships),
            ("word/document.xml", b"<w:document/>"),
        ])
    }

    for manifest in [None, Some(ct::OPC_RELATIONSHIPS)] {
        let bytes = package_with_rels_default(manifest);
        let reader = read(&bytes).expect("relationship part typing must not fail the package");
        assert_eq!(reader.pkg_srels().len(), 1);
        assert_eq!(blob(&reader, "/word/document.xml"), b"<w:document/>");
        assert!(reader.non_part_members().is_empty());
    }

    let bytes = package_with_rels_default(Some("application/xml"));
    assert!(matches!(
        read(&bytes),
        Err(OpcError::InvalidContentType { .. })
    ));
}

/// OPC compares part names case-insensitively, which is why a package holding
/// two names differing only by case is rejected as ambiguous. Because that
/// ambiguity cannot survive loading, a relationship target must still resolve
/// when a writer stored the part under a different case — otherwise the part is
/// unreachable and its content silently reads as absent, which is exactly what
/// happens to the shared strings of `poi/test-data/spreadsheet/49609.xlsx`.
#[test]
fn a_part_resolves_when_its_stored_name_differs_only_by_case() {
    let bytes = archive(&[
        (
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
        ),
        ("xl/sharedstrings.xml", b"<sst/>"),
    ]);
    let package = OpcPackage::from_bytes(&bytes).expect("package loads");

    // The stored spelling resolves, as always.
    let stored = litchi_opc::PackURI::new("/xl/sharedstrings.xml").unwrap();
    assert_eq!(package.get_part(&stored).unwrap().blob(), b"<sst/>");

    // The spelling a relationship would use also resolves.
    let referenced = litchi_opc::PackURI::new("/xl/sharedStrings.xml").unwrap();
    assert_eq!(
        package.get_part(&referenced).unwrap().blob(),
        b"<sst/>",
        "a case-only difference must not make the part unreachable"
    );

    // A genuinely absent part is still absent.
    let missing = litchi_opc::PackURI::new("/xl/styles.xml").unwrap();
    assert!(package.get_part(&missing).is_err());
}
