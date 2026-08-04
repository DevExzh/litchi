//! Focused tests for the layered font owner.

use super::codec::{parse, write};
use super::model::*;
use super::package::{put, read, remove};
use litchi_opc::{OpcPackage, PackURI, XmlPart};
use std::mem::size_of;
use std::sync::Arc;

const KEY: FontKey = FontKey::new([
    0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF,
]);

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let document = PackURI::new("/word/document.xml").expect("test URI");
    package.add_part(Box::new(XmlPart::new(
            document,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .into(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        )));
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "word/document.xml".into(),
        "rId1".into(),
        false,
    );
    package
}

#[test]
fn strict_round_trip_and_safe_selectors() {
    let xml = br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="A&amp;B"><w:altName w:val="Alias"/><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:notTrueType w:val="0"/><w:pitch w:val="variable"/><w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/><w:embedRegular r:id="rId1" w:fontKey="{01014A78-CABC-4EF0-12AC-5CD89AEFDE01}" w:subsetted="1"/></w:font></w:fonts>"#;
    let table = parse(xml).expect("parse");
    let by_name = table.get("a&b").expect("lookup").expect("font");
    let by_index = table.get(0usize).expect("lookup").expect("font");
    assert_eq!(by_name, by_index);
    assert!(table.get(9usize).expect("lookup").is_none());
    assert_eq!(
        by_name.signature().expect("signature").code_pages()[0],
        0x19F
    );

    let strict = write(&table, Conformance::Strict).expect("write");
    assert!(std::str::from_utf8(&strict).expect("UTF-8").contains(WS));
    assert_eq!(parse(&strict).expect("reparse"), table);
}

#[test]
fn mce_and_real_strict_fixture() {
    let xml = br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:font/></mc:Choice><mc:Fallback><w:font w:name="Fallback"><w:family w:val="roman"/></w:font></mc:Fallback></mc:AlternateContent></w:fonts>"#;
    assert_eq!(
        parse(xml)
            .expect("MCE parse")
            .get("Fallback")
            .expect("lookup")
            .expect("font")
            .family(),
        Some(Family::Roman)
    );

    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
        root.join("test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx"),
    )
    .expect("open fixture");
    let uri = PackURI::new("/word/fontTable.xml").expect("test URI");
    let table = parse(&physical.blob_for(&uri).expect("font table")).expect("parse");
    assert!(table.get("Calibri").expect("lookup").is_some());
    assert_eq!(
        table.get(0usize).expect("lookup").expect("font").charset(),
        Some(Charset::Ansi)
    );
}

#[test]
fn malformed_order_and_bounds_are_rejected() {
    for xml in [
        r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font/></w:fonts>"#,
        r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:family w:val="fantasy"/></w:font></w:fonts>"#,
        r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:panose1 w:val="1234"/></w:font></w:fonts>"#,
        r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:pitch w:val="fixed"/><w:family w:val="roman"/></w:font></w:fonts>"#,
        r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="x"><w:embedRegular r:id="rId1" w:fontKey="bad"/></w:font></w:fonts>"#,
        r#"<!DOCTYPE x><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
    ] {
        assert!(parse(xml.as_bytes()).is_err(), "{xml}");
    }
    assert!(parse(&vec![b' '; MAX_XML + 1]).is_err());
}

#[test]
fn poi_resources_share_package_allocations() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let package = OpcPackage::open(root.join("test-data/poi/test-data/document/saut_page.docx"))
        .expect("open fixture");
    let table = read(&package).expect("read").expect("font table");
    assert_eq!(table.len(), 7);
    let embeds = table.iter().flat_map(Font::embeds).collect::<Vec<_>>();
    assert_eq!(embeds.len(), 20);
    let first = embeds
        .iter()
        .find_map(|embed| embed.resource())
        .expect("embedded resource");
    let uri = PackURI::new(&first.part_name).expect("fixture part URI");
    assert!(Arc::ptr_eq(
        &first.data,
        &package.get_part(&uri).expect("font part").blob_arc()
    ));
    assert!(embeds.iter().all(|embed| embed.resource().is_some()));
}

#[test]
fn obfuscation_and_compact_license_are_checked() {
    let original = (0u8..64).collect::<Vec<_>>();
    let mut data = original.clone();
    obfuscate(&mut data, KEY).expect("obfuscate");
    assert_ne!(data, original);
    assert_eq!(&data[32..], &original[32..]);
    deobfuscate(&mut data, KEY).expect("deobfuscate");
    assert_eq!(data, original);
    assert!(obfuscate(&mut [0; 31], KEY).is_err());
    assert!("bad".parse::<FontKey>().is_err());
    assert_eq!(KEY.to_string(), "{00112233-4455-6677-8899-AABBCCDDEEFF}");

    assert!(License::new(0).expect("license").installable());
    let editable = License::new(0x0108).expect("license");
    assert!(editable.editable() && editable.no_subsetting());
    assert_eq!(editable.bits(), 0x0108);
    assert!(License::new(0x0006).is_err());
    assert!(License::new(0x8000).is_err());
    assert_eq!(size_of::<License>(), size_of::<u16>());
}

#[test]
fn move_first_crud_preserves_shared_resources_and_extensions() {
    let mut package = package();
    let shared = Resource::new((0u8..64).collect()).expect("resource");
    let first = Font::new("Alpha")
        .expect("font")
        .with_alt("Alpha Alt")
        .expect("alternate")
        .with_panose([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
        .with_charset(Charset::Ansi)
        .with_family(Family::Swiss)
        .with_pitch(Pitch::Variable)
        .with_signature(Signature::new([1, 2, 3, 4], [5, 6]))
        .with_embed(Embed::new(Style::Regular, KEY, shared.clone()).with_subset(true))
        .expect("face")
        .with_attr(raw::Attr::new("x:flag", "kept").expect("attribute"));
    let mut table = Table::new()
        .with_namespace(raw::Attr::new("xmlns:x", "urn:test-fonts").expect("namespace"))
        .expect("namespace");
    table.add(first).expect("add");
    put(&mut package, table, Conformance::Transitional).expect("put");

    let mut table = read(&package).expect("read").expect("table");
    table
        .add(
            Font::new("Beta")
                .expect("font")
                .with_embed(Embed::new(Style::Regular, KEY, shared))
                .expect("face"),
        )
        .expect("add");
    table.reorder(&["Beta", "Alpha"]).expect("reorder");
    put(&mut package, table, Conformance::Transitional).expect("put");

    let mut table = read(&package).expect("read").expect("table");
    assert_eq!(
        table.get(0usize).expect("lookup").expect("font").name(),
        "Beta"
    );
    assert_eq!(
        table
            .get("alpha")
            .expect("lookup")
            .expect("font")
            .attrs()
            .first()
            .expect("attribute")
            .value(),
        "kept"
    );
    let beta = table.get("Beta").expect("lookup").expect("font");
    let alpha = table.get("Alpha").expect("lookup").expect("font");
    assert!(
        beta.embeds()[0]
            .resource()
            .expect("resource")
            .shares_with(alpha.embeds()[0].resource().expect("resource"))
    );
    let shared_part = beta.embeds()[0]
        .resource()
        .expect("resource")
        .part_name
        .clone();

    assert!(table.remove("Alpha").expect("remove").is_some());
    put(&mut package, table, Conformance::Transitional).expect("put");
    assert!(
        package
            .get_part(&PackURI::new(&shared_part).expect("part URI"))
            .is_ok()
    );
    assert!(
        read(&package)
            .expect("read")
            .expect("table")
            .get("Beta")
            .expect("lookup")
            .is_some()
    );

    put(&mut package, Table::new(), Conformance::Transitional).expect("remove graph");
    assert!(read(&package).expect("read").is_none());
    assert!(
        package
            .get_part(&PackURI::new(&shared_part).expect("part URI"))
            .is_err()
    );
}

#[test]
fn graph_delete_keeps_resources_referenced_outside_the_table() {
    let mut package = package();
    let font = Font::new("Shared")
        .expect("font")
        .with_embed(Embed::new(
            Style::Regular,
            KEY,
            Resource::new(vec![0; 32]).expect("resource"),
        ))
        .expect("face");
    let mut table = Table::new();
    table.add(font).expect("add");
    put(&mut package, table, Conformance::Transitional).expect("put");

    let table = read(&package).expect("read").expect("table");
    let resource_name = table.fonts[0].embedded_fonts[0]
        .resource
        .as_ref()
        .expect("resource")
        .part_name
        .clone();
    let resource = PackURI::new(&resource_name).expect("resource URI");
    let main_name = package
        .main_document_part()
        .expect("main")
        .partname()
        .clone();
    package
        .get_part_mut(&main_name)
        .expect("main")
        .rels_mut()
        .add_relationship(
            "urn:litchi:test:keep-font".into(),
            resource.relative_ref(main_name.base_uri()),
            "rIdKeepFont".into(),
            false,
        );

    assert!(remove(&mut package).expect("remove graph"));
    assert!(read(&package).expect("read").is_none());
    assert!(package.get_part(&resource).is_ok());
    assert!(!remove(&mut package).expect("already absent"));
}

#[test]
fn constructors_prevent_invalid_authoring_state() {
    assert!(Font::new("").is_err());
    assert!(Font::new("12345678901234567890123456789012").is_err());
    assert!(Resource::new(vec![0; 31]).is_err());
    assert!("bad".parse::<FontKey>().is_err());
    assert!(raw::Attr::new("", "value").is_err());

    let mut table = Table::new();
    table.add(Font::new("Alpha").expect("font")).expect("add");
    assert!(table.add(Font::new("alpha").expect("font")).is_err());
    assert!(
        table
            .replace(9usize, Font::new("Beta").expect("font"))
            .expect("replace")
            .is_none()
    );
    assert!(table.remove(9usize).expect("remove").is_none());
}

#[test]
fn unicode_caseless_identity_is_consistent_and_ambiguous_sources_fail() {
    let mut table = Table::new();
    table.add(Font::new("Straße").expect("font")).expect("add");
    assert!(table.add(Font::new("STRASSE").expect("font")).is_err());
    table.add(Font::new("École").expect("font")).expect("add");

    assert!(table.get("strasse").expect("lookup").is_some());
    assert!(table.get("e\u{301}COLE").expect("lookup").is_some());
    table
        .reorder(&["e\u{301}cole", "STRASSE"])
        .expect("reorder");
    assert_eq!(
        table.get(0usize).expect("lookup").expect("font").name(),
        "École"
    );

    let malformed_xml = r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Straße"/><w:font w:name="STRASSE"/></w:fonts>"#;
    let malformed = parse(malformed_xml.as_bytes()).expect("parse malformed producer table");
    assert!(malformed.get("strasse").is_err());
}
