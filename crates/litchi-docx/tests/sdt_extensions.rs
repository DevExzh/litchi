use std::collections::HashSet;

use litchi_docx::content_control::{
    Checksum, ChecksumStatus, ChecksumValue, FormattingAllowed, Inventory, Limits,
    PackageChecksumStatus, PackageLimits, Snapshot,
};
use litchi_docx::custom_xml::NewStore;
use litchi_docx::package::{StoryDialect, StoryKind, StoryLimits};
use litchi_docx::{Error, Package};
use litchi_ooxml_common::custom_xml::Conformance;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::rel::Relationships;
use litchi_opc::{OpcPackage, PackURI};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const HASH: &str = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";
const FORMAT: &str = "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock";
const W15: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const ITEM: &str = "{11111111-1111-4111-8111-111111111111}";
const ITEM_B: &str = "{22222222-2222-4222-8222-222222222222}";
const STRICT_RELS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

fn binding(id: Option<u32>, checksum: Option<&str>) -> String {
    let id = id.map_or_else(String::new, |id| format!(r#"<w:id w:val="{id}"/>"#));
    let checksum = checksum.map_or_else(String::new, |value| {
        format!(r#" h:storeItemChecksum="{value}""#)
    });
    format!(
        r#"<w:sdtPr>{id}<w:dataBinding w:xpath="/x:root/x:value" w:storeItemID="{ITEM}" w:prefixMappings="xmlns:x='urn:test'"{checksum}/></w:sdtPr>"#
    )
}

fn main_xml(controls: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h"><w:body>{controls}</w:body></w:document>"#
    )
    .into_bytes()
}

fn store(xml: &[u8]) -> NewStore {
    store_for(ITEM, xml, Conformance::Transitional)
}

fn store_for(id: &str, xml: &[u8], conformance: Conformance) -> NewStore {
    NewStore {
        xml: xml.to_vec(),
        content_type: "application/xml".to_owned(),
        id: id.to_owned(),
        schemas: vec!["urn:test".to_owned()],
        conformance,
    }
}

fn mark_signed(package: &mut Package) {
    package
        .edit_opc(|opc| {
            opc.try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
            )))?;
            opc.rels_mut().add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rSignature".to_owned(),
                false,
            );
            Ok::<_, Error>(())
        })
        .unwrap();
}

#[derive(Debug, PartialEq, Eq)]
struct OpcPartState {
    name: String,
    content_type: String,
    payload: Vec<u8>,
    relationships: Vec<(String, String, String, bool)>,
}

#[derive(Debug, PartialEq, Eq)]
struct OpcState {
    parts: Vec<OpcPartState>,
    root_relationships: Vec<(String, String, String, bool)>,
}

fn relationship_state(rels: &Relationships) -> Vec<(String, String, String, bool)> {
    let mut state = rels
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            )
        })
        .collect::<Vec<_>>();
    state.sort();
    state
}

fn opc_state(package: &Package) -> OpcState {
    let opc = package.opc_package();
    let mut parts = opc
        .iter_parts()
        .map(|part| OpcPartState {
            name: part.partname().as_str().to_owned(),
            content_type: part.content_type().to_owned(),
            payload: part.blob().to_vec(),
            relationships: relationship_state(part.rels()),
        })
        .collect::<Vec<_>>();
    parts.sort_by(|left, right| left.name.cmp(&right.name));
    OpcState {
        parts,
        root_relationships: relationship_state(opc.rels()),
    }
}

fn bound_package(payload: &[u8]) -> Package {
    let checksum = Checksum::compute(payload, &Limits::default())
        .unwrap()
        .to_base64();
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(payload)).unwrap();
    let document_name = PackURI::new("/word/document.xml").unwrap();
    let header_name = PackURI::new("/word/header1.xml").unwrap();
    let document_xml = main_xml(&format!(
        "{}{}",
        binding(Some(1), Some(&checksum)),
        binding(Some(2), None)
    ));
    let header_xml = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{}</w:hdr>"#,
        binding(Some(3), Some(&checksum))
    )
    .into_bytes();
    package
        .edit_opc(|opc| {
            let document = opc.get_part_mut(&document_name)?;
            document.set_blob(document_xml);
            document.rels_mut().add_relationship(
                format!("{}/header", STRICT_RELS.trim_end_matches('/')).replace(
                    "http://purl.oclc.org/ooxml/officeDocument/relationships/header",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                ),
                "header1.xml".to_owned(),
                "rIdHeader1".to_owned(),
                false,
            );
            opc.add_part(Box::new(BlobPart::new(
                header_name,
                ct::WML_HEADER.to_owned(),
                header_xml,
            )));
            Ok::<_, Error>(())
        })
        .unwrap();
    package
}

#[test]
fn crc_vectors_are_over_exact_bytes_and_use_the_office_little_endian_form() {
    let limits = Limits::default();
    for (source, word, lexical) in [
        (&b""[..], 0x0000_0000, "AAAAAA=="),
        (&b"\x01"[..], 0x0000_00af, "rwAAAA=="),
        (&b"123456789"[..], 0xbd0b_e338, "OOMLvQ=="),
        (&b"<r/>"[..], 0xe19d_c302, "AsOd4Q=="),
        (&b"<r />"[..], 0x9bfc_2d85, "hS38mw=="),
        (&b"<r>\xc3\xa9</r>\n"[..], 0xc14f_1e92, "kh5PwQ=="),
    ] {
        let checksum = Checksum::compute(source, &limits).unwrap();
        assert_eq!(checksum.word_value(), word);
        assert_eq!(checksum.to_base64(), lexical);
        assert_eq!(
            Checksum::parse(lexical).unwrap().as_bytes(),
            checksum.as_bytes()
        );
    }

    let bom_crlf = b"\xef\xbb\xbf<r>\r\n  <v>alpha</v>\r\n</r>";
    assert_eq!(
        Checksum::compute(bom_crlf, &limits).unwrap().to_base64(),
        "b04H7Q=="
    );
    assert_ne!(
        Checksum::compute(b"<r/>", &limits).unwrap(),
        Checksum::compute(b"<r />", &limits).unwrap()
    );
}

#[test]
fn checksum_equality_is_semantic_while_source_lexical_provenance_is_retained() {
    let authored = Checksum::from_word_value(0xbd0b_e338);
    let parsed = Checksum::parse("OOMLvQ==").unwrap();
    assert_eq!(authored.original_lexical(), None);
    assert_eq!(parsed.original_lexical(), Some("OOMLvQ=="));
    assert_eq!(authored.lexical(), "OOMLvQ==");
    assert_eq!(parsed.lexical(), "OOMLvQ==");
    assert_eq!(authored, parsed);

    let mut values = HashSet::new();
    values.insert(authored.clone());
    values.insert(parsed.clone());
    assert_eq!(values.len(), 1);

    let authored_value = ChecksumValue::Valid(authored);
    let parsed_value = ChecksumValue::Valid(parsed);
    assert_eq!(authored_value.lexical(), "OOMLvQ==");
    assert_eq!(parsed_value.lexical(), "OOMLvQ==");
    assert_eq!(authored_value, parsed_value);

    let mut values = HashSet::new();
    values.insert(authored_value);
    values.insert(parsed_value);
    assert_eq!(values.len(), 1);
}

#[test]
fn malformed_base64_is_distinct_from_a_valid_checksum_mismatch() {
    for invalid in ["", "OOMLvQ=", "OOMLvQ===", "OOMLvQ==\n", "OOMLvQ--"] {
        assert!(Checksum::parse(invalid).is_err(), "accepted {invalid:?}");
    }
    assert_eq!(
        Checksum::parse("vQvjOA==").unwrap().word_value(),
        0x38e3_0bbd
    );

    let source = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h"><w:body><w:sdtPr><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}" h:storeItemChecksum="bad"/></w:sdtPr><w:sdtPr><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}" h:storeItemChecksum="OOMLvQ=="/></w:sdtPr></w:body></w:document>"#
    );
    let inventory = Inventory::parse(source.as_bytes()).unwrap();
    let first = inventory.occurrences()[0].control().data_binding().unwrap();
    assert!(
        matches!(first.checksum_status(), ChecksumStatus::Malformed(value) if &*value == "bad")
    );
    let second = inventory.occurrences()[1].control().data_binding().unwrap();
    assert!(matches!(
        second
            .verify_checksum(b"different", &Limits::default())
            .unwrap(),
        ChecksumStatus::Mismatch { .. }
    ));
}

#[test]
fn expanded_names_ignorable_scope_and_mce_select_only_active_metadata() {
    let missing_ignorable = format!(
        r#"<w:sdtPr xmlns:w="{W}" xmlns:h="{HASH}"><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}" h:storeItemChecksum="OOMLvQ=="/></w:sdtPr>"#
    );
    assert!(Inventory::parse(missing_ignorable.as_bytes()).is_err());

    let spoofed = format!(
        r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="urn:not-the-hash-namespace" mc:Ignorable="h"><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}" h:storeItemChecksum="OOMLvQ=="/></w:sdtPr>"#
    );
    let spoofed = Inventory::parse(spoofed.as_bytes()).unwrap();
    assert!(matches!(
        spoofed.occurrences()[0]
            .control()
            .data_binding()
            .unwrap()
            .checksum_status(),
        ChecksumStatus::Absent
    ));

    let alternate = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" xmlns:u="urn:unsupported" mc:Ignorable="h u"><mc:AlternateContent><mc:Choice Requires="u">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent></w:document>"#,
        binding(Some(11), Some("AAAAAA==")),
        binding(Some(22), Some("OOMLvQ=="))
    );
    let selected = Inventory::parse(alternate.as_bytes()).unwrap();
    assert_eq!(selected.occurrences().len(), 1);
    assert_eq!(selected.occurrences()[0].id(), Some(22));
}

#[test]
fn occurrences_preserve_missing_and_duplicate_ids_and_all_st_on_off_forms() {
    let ids = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:sdtPr/><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtPr><w:id w:val="7"/></w:sdtPr></w:body></w:document>"#
    );
    let ids = Inventory::parse(ids.as_bytes()).unwrap();
    assert_eq!(
        ids.occurrences()
            .iter()
            .map(|value| value.id())
            .collect::<Vec<_>>(),
        vec![None, Some(7), Some(7)]
    );
    assert_eq!(
        ids.occurrences()
            .iter()
            .map(|value| value.ordinal())
            .collect::<Vec<_>>(),
        vec![0, 1, 2]
    );

    for (lexical, expected) in [
        ("true", true),
        ("1", true),
        ("on", true),
        ("false", false),
        ("0", false),
        ("off", false),
    ] {
        let source = format!(
            r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{MC}" xmlns:f="{FORMAT}" mc:Ignorable="f"><w:lock w:val="contentLocked" f:formattingAllowed="{lexical}"/></w:sdtPr>"#
        );
        let inventory = Inventory::parse(source.as_bytes()).unwrap();
        assert_eq!(
            inventory.occurrences()[0]
                .control()
                .formatting_allowed()
                .unwrap()
                .as_bool(),
            expected
        );
    }

    let invalid = format!(
        r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{MC}" xmlns:f="{FORMAT}" mc:Ignorable="f"><w:lock w:val="sdtLocked" f:formattingAllowed="true"/></w:sdtPr>"#
    );
    assert!(Inventory::parse(invalid.as_bytes()).is_err());
}

#[test]
fn core_and_word_2012_bindings_remain_distinct_exact_occurrences() {
    let source = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w15="{W15}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="w15 h"><w:body>
        <w:sdtPr><w:id w:val="1"/><w:dataBinding w:xpath="/core" w:storeItemID="{ITEM}" h:storeItemChecksum="AAAAAA=="/></w:sdtPr>
        <w:sdtPr><w:id w:val="2"/><w15:dataBinding w:xpath="/compat" w:storeItemID="{ITEM}" h:storeItemChecksum="rwAAAA=="/></w:sdtPr>
        <w:sdtPr><w:id w:val="3"/><w:dataBinding w:xpath="/preferred" w:storeItemID="{ITEM}" h:storeItemChecksum="AAAAAA=="/><w15:dataBinding w:xpath="/compatibility" w:storeItemID="{ITEM}" h:storeItemChecksum="rwAAAA=="/></w:sdtPr>
        </w:body></w:document>"#
    );
    let inventory = Inventory::parse(source.as_bytes()).unwrap();
    assert_eq!(inventory.occurrences().len(), 3);
    assert_eq!(
        inventory.occurrences()[0]
            .control()
            .data_binding()
            .unwrap()
            .xpath(),
        "/core"
    );
    assert_eq!(
        inventory.occurrences()[1]
            .control()
            .data_binding()
            .unwrap()
            .xpath(),
        "/compat"
    );
    assert_eq!(
        inventory.occurrences()[2]
            .control()
            .data_binding()
            .unwrap()
            .xpath(),
        "/preferred"
    );

    let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
    assert_eq!(snapshot.occurrences()[0].bindings().len(), 1);
    assert_eq!(snapshot.occurrences()[1].bindings().len(), 1);
    assert_eq!(snapshot.occurrences()[2].bindings().len(), 2);
}

#[test]
fn managed_custom_xml_refresh_updates_both_bindings_in_one_sdt() {
    let initial = b"<root xmlns='urn:test'><value>dual-old</value></root>";
    let changed = b"<root xmlns='urn:test'><value>dual-new</value></root>";
    let initial_checksum = Checksum::compute(initial, &Limits::default())
        .unwrap()
        .to_base64();
    let expected = Checksum::compute(changed, &Limits::default())
        .unwrap()
        .to_base64();
    let document_xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w15="{W15}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="w15 h"><w:body><w:sdtPr><w:id w:val="9"/><w:dataBinding w:xpath="/core" w:storeItemID="{ITEM}" h:storeItemChecksum="{initial_checksum}"/><w15:dataBinding w:xpath="/compat" w:storeItemID="{ITEM}" h:storeItemChecksum="{initial_checksum}"/></w:sdtPr></w:body></w:document>"#
    )
    .into_bytes();

    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(initial)).unwrap();
    let main = PackURI::new("/word/document.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&main)?.set_blob(document_xml);
            Ok::<_, Error>(())
        })
        .unwrap();
    package.set_custom_xml(ITEM, changed.to_vec()).unwrap();

    let source =
        std::str::from_utf8(package.opc_package().get_part(&main).unwrap().blob()).unwrap();
    assert_eq!(source.matches(&expected).count(), 2);
    assert!(!source.contains(&initial_checksum));
    let reports = package.verify_content_control_checksums().unwrap();
    assert_eq!(reports.len(), 2);
    assert!(
        reports
            .iter()
            .all(|entry| matches!(entry.status(), PackageChecksumStatus::Matches))
    );
}

fn strict_binding(id: u32, checksum: &str) -> String {
    format!(
        r#"<s:sdtPr><s:id s:val="{id}"/><s:dataBinding s:xpath="/x:root/x:value" s:storeItemID="{ITEM}" s:prefixMappings="xmlns:x='urn:test'" h:storeItemChecksum="{checksum}"/></s:sdtPr>"#
    )
}

fn strict_story_source(kind: StoryKind, id: u32, checksum: &str) -> Vec<u8> {
    let control = strict_binding(id, checksum);
    let root = match kind {
        StoryKind::Main => format!(
            r#"<s:document xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h"><s:body>{control}</s:body></s:document>"#
        ),
        StoryKind::Header => format!(
            r#"<s:hdr xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{control}</s:hdr>"#
        ),
        StoryKind::Footer => format!(
            r#"<s:ftr xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{control}</s:ftr>"#
        ),
        StoryKind::Footnotes => format!(
            r#"<s:footnotes xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{control}</s:footnotes>"#
        ),
        StoryKind::Endnotes => format!(
            r#"<s:endnotes xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{control}</s:endnotes>"#
        ),
        StoryKind::Comments => format!(
            r#"<s:comments xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{control}</s:comments>"#
        ),
        StoryKind::Glossary => format!(
            r#"<s:glossaryDocument xmlns:s="{WS}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h">{control}</s:glossaryDocument>"#
        ),
    };
    root.into_bytes()
}

#[test]
fn changed_payload_refreshes_all_seven_strict_story_kinds() {
    let initial = b"<root xmlns='urn:test'><value>strict-old</value></root>";
    let changed = b"<root xmlns='urn:test'><value>strict-new</value></root>";
    let checksum = Checksum::compute(initial, &Limits::default())
        .unwrap()
        .to_base64();
    let mut package = strict_story_package(ct::WML_DOCUMENT_MAIN);
    package
        .add_custom_xml(store_for(ITEM, initial, Conformance::Strict))
        .unwrap();
    let replacements = package
        .story_inventory()
        .unwrap()
        .stories()
        .iter()
        .enumerate()
        .map(|(index, story)| {
            (
                story.part().clone(),
                strict_story_source(story.kind(), index as u32 + 1, &checksum),
            )
        })
        .collect::<Vec<_>>();
    package
        .edit_opc(|opc| {
            for (part, source) in replacements {
                opc.get_part_mut(&part)?.set_blob(source);
            }
            Ok::<_, Error>(())
        })
        .unwrap();

    package.set_custom_xml(ITEM, changed.to_vec()).unwrap();
    let reports = package.verify_content_control_checksums().unwrap();
    assert_eq!(reports.len(), 7);
    assert!(
        reports
            .iter()
            .all(|entry| matches!(entry.status(), PackageChecksumStatus::Matches))
    );
    assert_eq!(
        package
            .content_control_snapshot()
            .unwrap()
            .stories()
            .iter()
            .map(|story| story.kind())
            .collect::<HashSet<_>>(),
        HashSet::from([
            StoryKind::Main,
            StoryKind::Header,
            StoryKind::Footer,
            StoryKind::Footnotes,
            StoryKind::Endnotes,
            StoryKind::Comments,
            StoryKind::Glossary,
        ])
    );
}

#[test]
fn transitional_docm_main_is_a_reachable_inert_checksum_owner() {
    let initial = b"<root xmlns='urn:test'><value>docm-old</value></root>";
    let changed = b"<root xmlns='urn:test'><value>docm-new</value></root>";
    let checksum = Checksum::compute(initial, &Limits::default())
        .unwrap()
        .to_base64();
    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MACRO_MAIN.to_owned(),
        main_xml(&binding(Some(41), Some(&checksum))),
    )));
    opc.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "word/document.xml".to_owned(),
        "rIdMain".to_owned(),
        false,
    );
    let mut package = Package::from_opc_package(opc).unwrap();
    package.add_custom_xml(store(initial)).unwrap();
    package.set_custom_xml(ITEM, changed.to_vec()).unwrap();
    let reports = package.verify_content_control_checksums().unwrap();
    assert_eq!(reports.len(), 1);
    assert!(matches!(
        reports[0].status(),
        PackageChecksumStatus::Matches
    ));
    assert_eq!(
        package.story_inventory().unwrap().stories()[0].content_type(),
        ct::WML_DOCUMENT_MACRO_MAIN
    );
}

fn add_strict_story(
    opc: &mut OpcPackage,
    owner: &mut BlobPart,
    name: &str,
    target: &str,
    kind: &str,
    content_type: &str,
    root: &str,
) {
    owner.rels_mut().add_relationship(
        format!("{STRICT_RELS}{kind}"),
        target.to_owned(),
        format!("rId{name}"),
        false,
    );
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new(name).unwrap(),
        content_type.to_owned(),
        format!(r#"<s:{root} xmlns:s="{WS}"/>"#).into_bytes(),
    )));
}

fn strict_story_package(main_content_type: &str) -> Package {
    let mut opc = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        main_content_type.to_owned(),
        format!(r#"<s:document xmlns:s="{WS}"><s:body/></s:document>"#).into_bytes(),
    );
    for (name, target, kind, content_type, root) in [
        (
            "/word/header1.xml",
            "header1.xml",
            "header",
            ct::WML_HEADER,
            "hdr",
        ),
        (
            "/word/footer1.xml",
            "footer1.xml",
            "footer",
            ct::WML_FOOTER,
            "ftr",
        ),
        (
            "/word/footnotes.xml",
            "footnotes.xml",
            "footnotes",
            ct::WML_FOOTNOTES,
            "footnotes",
        ),
        (
            "/word/endnotes.xml",
            "endnotes.xml",
            "endnotes",
            ct::WML_ENDNOTES,
            "endnotes",
        ),
        (
            "/word/comments.xml",
            "comments.xml",
            "comments",
            ct::WML_COMMENTS,
            "comments",
        ),
        (
            "/word/glossary/document.xml",
            "glossary/document.xml",
            "glossaryDocument",
            ct::WML_DOCUMENT_GLOSSARY,
            "glossaryDocument",
        ),
    ] {
        add_strict_story(&mut opc, &mut main, name, target, kind, content_type, root);
    }
    opc.add_part(Box::new(main));
    opc.rels_mut().add_relationship(
        rt::STRICT_OFFICE_DOCUMENT.to_owned(),
        "word/document.xml".to_owned(),
        "rIdMain".to_owned(),
        false,
    );
    Package::from_opc_package(opc).unwrap()
}

#[test]
fn strict_dotx_and_dotm_inventory_every_reachable_story_role() {
    for main_type in [ct::WML_TEMPLATE_MAIN, ct::WML_TEMPLATE_MACRO_MAIN] {
        let package = strict_story_package(main_type);
        let inventory = package.story_inventory().unwrap();
        assert_eq!(inventory.dialect(), StoryDialect::Strict);
        assert_eq!(inventory.stories().len(), 7);
        assert_eq!(inventory.stories()[0].kind(), StoryKind::Main);
        let kinds = inventory
            .stories()
            .iter()
            .map(|story| story.kind())
            .collect::<HashSet<_>>();
        assert_eq!(
            kinds,
            HashSet::from([
                StoryKind::Main,
                StoryKind::Header,
                StoryKind::Footer,
                StoryKind::Footnotes,
                StoryKind::Endnotes,
                StoryKind::Comments,
                StoryKind::Glossary,
            ])
        );

        let exact = StoryLimits {
            max_stories: 7,
            ..StoryLimits::default()
        };
        assert_eq!(
            package
                .story_inventory_with_limits(exact)
                .unwrap()
                .stories()
                .len(),
            7
        );
        assert!(
            package
                .story_inventory_with_limits(StoryLimits {
                    max_stories: 6,
                    ..StoryLimits::default()
                })
                .is_err()
        );
    }
}

#[test]
fn detached_edits_are_reversible_stale_safe_and_output_bounded() {
    let source = format!(
        r#"<w:sdtPr xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" xmlns:f="{FORMAT}" mc:Ignorable="h f"><w:lock w:val="contentLocked"/><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}"/></w:sdtPr>"#
    )
    .into_bytes();
    let snapshot = Snapshot::from_xml(source.clone()).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .set_checksum(0, Some(Checksum::from_word_value(0xbd0b_e338)))
        .unwrap()
        .set_formatting_allowed(0, Some(FormattingAllowed::Allowed))
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());

    let patch = commit.patch().clone();
    let stale = Snapshot::from_xml(
        source
            .iter()
            .copied()
            .chain(b" ".iter().copied())
            .collect::<Vec<_>>(),
    )
    .unwrap();
    assert!(patch.apply(&stale).is_err());
    assert!(!patch.is_applied());
    let applied = patch.apply(&snapshot).unwrap();
    assert!(patch.is_applied());
    assert!(patch.apply(&snapshot).is_err());
    let restored = patch.inverse().apply(&applied).unwrap();
    assert_eq!(restored.source(), source);

    let limited_source = format!(
        r#"<w:sdtPr xmlns:w="{W}"><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}"/></w:sdtPr>"#
    )
    .into_bytes();
    let limited = Snapshot::from_xml_with_limits(
        limited_source.clone(),
        Limits {
            max_output_bytes: limited_source.len(),
            ..Limits::default()
        },
    )
    .unwrap();
    let mut transaction = limited.edit();
    transaction
        .set_checksum(0, Some(Checksum::from_word_value(0xbd0b_e338)))
        .unwrap();
    assert!(transaction.commit().is_err());
    assert_eq!(limited.source(), limited_source);
}

#[test]
fn custom_xml_updates_refresh_all_declared_bindings_and_preserve_signed_noops() {
    let initial = b"<root xmlns='urn:test'><value>old</value></root>";
    let changed = b"<root xmlns='urn:test'><value>new</value></root>";
    let mut package = bound_package(initial);
    let reports = package.verify_content_control_checksums().unwrap();
    assert_eq!(
        reports
            .iter()
            .filter(|entry| matches!(entry.status(), PackageChecksumStatus::Matches))
            .count(),
        2
    );
    assert_eq!(
        reports
            .iter()
            .filter(|entry| matches!(entry.status(), PackageChecksumStatus::Absent))
            .count(),
        1
    );

    mark_signed(&mut package);
    let signed_state = opc_state(&package);
    package
        .set_custom_xml(&ITEM.to_ascii_lowercase(), initial.to_vec())
        .unwrap();
    assert!(package.is_signed());
    assert_eq!(opc_state(&package), signed_state);

    assert!(
        package
            .set_custom_xml(&ITEM.to_ascii_lowercase(), changed.to_vec())
            .is_err()
    );
    assert!(package.is_signed());
    assert_eq!(opc_state(&package), signed_state);

    package.unsign();
    assert!(!package.is_signed());
    package
        .set_custom_xml(&ITEM.to_ascii_lowercase(), changed.to_vec())
        .unwrap();

    let reports = package.verify_content_control_checksums().unwrap();
    assert_eq!(
        reports
            .iter()
            .filter(|entry| matches!(entry.status(), PackageChecksumStatus::Matches))
            .count(),
        2
    );
    assert_eq!(
        reports
            .iter()
            .filter(|entry| matches!(entry.status(), PackageChecksumStatus::Absent))
            .count(),
        1
    );
}

#[test]
fn changed_bound_replace_requires_explicit_unsign_and_revalidates_every_checksum() {
    let initial = b"<root xmlns='urn:test'><value>replace-old</value></root>";
    let changed = b"<root xmlns='urn:test'><value>replace-new</value></root>";
    let mut package = bound_package(initial);
    mark_signed(&mut package);
    let replacement = || {
        let mut replacement = store(changed);
        replacement.content_type = "application/vnd.litchi.test+xml".to_owned();
        replacement.schemas.push("urn:test:replacement".to_owned());
        replacement
    };
    let signed_state = opc_state(&package);
    assert!(package.replace_custom_xml(ITEM, replacement()).is_err());
    assert!(package.is_signed());
    assert_eq!(opc_state(&package), signed_state);

    package.unsign();
    assert!(!package.is_signed());
    package.replace_custom_xml(ITEM, replacement()).unwrap();
    assert!(
        package
            .verify_content_control_checksums()
            .unwrap()
            .iter()
            .all(|entry| {
                matches!(
                    entry.status(),
                    PackageChecksumStatus::Matches | PackageChecksumStatus::Absent
                )
            })
    );
}

#[test]
fn signed_remove_preserves_noops_and_requires_explicit_unsign_for_real_publication() {
    let payload = b"<root xmlns='urn:test'><value>bound</value></root>";
    let mut bound = bound_package(payload);
    mark_signed(&mut bound);
    let bound_signed_state = opc_state(&bound);
    assert!(!bound.remove_custom_xml(ITEM_B).unwrap());
    assert!(bound.is_signed());
    assert_eq!(opc_state(&bound), bound_signed_state);
    assert!(bound.remove_custom_xml(ITEM).is_err());
    assert!(bound.is_signed());
    assert!(bound.custom_xml_by_id(ITEM).unwrap().is_some());
    assert_eq!(opc_state(&bound), bound_signed_state);

    let mut unbound = Package::new().unwrap();
    unbound.add_custom_xml(store(payload)).unwrap();
    mark_signed(&mut unbound);
    let unbound_signed_state = opc_state(&unbound);
    assert!(unbound.remove_custom_xml(ITEM).is_err());
    assert!(unbound.is_signed());
    assert_eq!(opc_state(&unbound), unbound_signed_state);

    unbound.unsign();
    assert!(!unbound.is_signed());
    assert!(unbound.remove_custom_xml(ITEM).unwrap());
    assert!(unbound.custom_xml_by_id(ITEM).unwrap().is_none());
}

#[test]
fn pending_managed_document_state_refuses_custom_xml_removal() {
    let payload = b"<root xmlns='urn:test'><value>pending</value></root>";
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(payload)).unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("pending managed edit");
    assert!(package.remove_custom_xml(ITEM).is_err());
    assert!(package.custom_xml_by_id(ITEM).unwrap().is_some());
}

fn two_store_package() -> (Package, usize, usize) {
    let first = b"<a>first-limit-payload</a>";
    let second = b"<b>second-limit-payload</b>";
    let first_crc = Checksum::compute(first, &Limits::default())
        .unwrap()
        .to_base64();
    let second_crc = Checksum::compute(second, &Limits::default())
        .unwrap()
        .to_base64();
    let controls = format!(
        r#"<w:sdtPr><w:id w:val="1"/><w:dataBinding w:xpath="/a" w:storeItemID="{ITEM}" h:storeItemChecksum="{first_crc}"/></w:sdtPr><w:sdtPr><w:id w:val="2"/><w:dataBinding w:xpath="/b" w:storeItemID="{ITEM_B}" h:storeItemChecksum="{second_crc}"/></w:sdtPr>"#
    );
    let source = main_xml(&controls);
    let source_len = source.len();
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(first)).unwrap();
    package
        .add_custom_xml(store_for(ITEM_B, second, Conformance::Transitional))
        .unwrap();
    let main = PackURI::new("/word/document.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&main)?.set_blob(source);
            Ok::<_, Error>(())
        })
        .unwrap();
    (package, source_len, first.len() + second.len())
}

fn minimum_accepted(mut accepts: impl FnMut(usize) -> bool) -> usize {
    let mut upper = 1usize;
    while !accepts(upper) {
        upper = upper.checked_mul(2).expect("test limit search overflow");
    }
    let mut lower = upper / 2 + 1;
    while lower < upper {
        let middle = lower + (upper - lower) / 2;
        if accepts(middle) {
            upper = middle;
        } else {
            lower = middle + 1;
        }
    }
    lower
}

#[test]
fn package_limits_accept_exact_boundaries_and_reject_one_over() {
    let (mut package, source_len, crc_bytes) = two_store_package();
    let mut limits = PackageLimits::default();
    limits.controls.max_input_bytes = source_len;
    limits.max_content_controls = 2;
    assert!(
        package
            .content_control_snapshot_with_limits(limits.clone())
            .is_ok()
    );
    limits.controls.max_input_bytes = source_len - 1;
    assert!(
        package
            .content_control_snapshot_with_limits(limits)
            .is_err()
    );

    let mut limits = PackageLimits {
        max_content_controls: 1,
        ..PackageLimits::default()
    };
    assert!(
        package
            .content_control_snapshot_with_limits(limits.clone())
            .is_err()
    );
    limits.max_content_controls = 2;
    assert!(package.content_control_snapshot_with_limits(limits).is_ok());

    let mut limits = PackageLimits {
        max_crc_parts: 2,
        max_crc_bytes: crc_bytes,
        ..PackageLimits::default()
    };
    assert!(
        package
            .content_control_snapshot_with_limits(limits.clone())
            .unwrap()
            .verify_checksums()
            .is_ok()
    );
    limits.max_crc_parts = 1;
    assert!(
        package
            .content_control_snapshot_with_limits(limits.clone())
            .and_then(|snapshot| snapshot.verify_checksums())
            .is_err()
    );
    limits.max_crc_parts = 2;
    limits.max_crc_bytes = crc_bytes - 1;
    assert!(
        package
            .content_control_snapshot_with_limits(limits)
            .and_then(|snapshot| snapshot.verify_checksums())
            .is_err()
    );

    mark_signed(&mut package);
    let signature_minimum = minimum_accepted(|value| {
        package
            .content_control_snapshot_with_limits(PackageLimits {
                max_signature_bytes: value,
                ..PackageLimits::default()
            })
            .is_ok()
    });
    assert!(signature_minimum > 1);
    assert!(
        package
            .content_control_snapshot_with_limits(PackageLimits {
                max_signature_bytes: signature_minimum,
                ..PackageLimits::default()
            })
            .is_ok()
    );
    assert!(
        package
            .content_control_snapshot_with_limits(PackageLimits {
                max_signature_bytes: signature_minimum - 1,
                ..PackageLimits::default()
            })
            .is_err()
    );

    let custom_graph_minimum = minimum_accepted(|value| {
        package
            .content_control_snapshot_with_limits(PackageLimits {
                max_custom_graph_bytes: value,
                ..PackageLimits::default()
            })
            .is_ok()
    });
    assert!(custom_graph_minimum > 1);
    assert!(
        package
            .content_control_snapshot_with_limits(PackageLimits {
                max_custom_graph_bytes: custom_graph_minimum,
                ..PackageLimits::default()
            })
            .is_ok()
    );
    assert!(
        package
            .content_control_snapshot_with_limits(PackageLimits {
                max_custom_graph_bytes: custom_graph_minimum - 1,
                ..PackageLimits::default()
            })
            .is_err()
    );
}

#[test]
fn package_authored_output_limit_is_exact() {
    let payload = b"<root>output-limit</root>";
    let mut package = Package::new().unwrap();
    package.add_custom_xml(store(payload)).unwrap();
    let main = PackURI::new("/word/document.xml").unwrap();
    let source = main_xml(&binding(Some(1), None));
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&main)?.set_blob(source);
            Ok::<_, Error>(())
        })
        .unwrap();
    let checksum = Checksum::compute(payload, &Limits::default()).unwrap();
    let base = package.content_control_snapshot().unwrap();
    let mut transaction = base.edit().unwrap();
    transaction
        .set_checksum(&main, 0, Some(checksum.clone()))
        .unwrap();
    let expected = transaction.commit().unwrap().stories()[0]
        .snapshot()
        .source()
        .len();

    let exact = package
        .content_control_snapshot_with_limits(PackageLimits {
            max_output_bytes: expected,
            ..PackageLimits::default()
        })
        .unwrap();
    let mut transaction = exact.edit().unwrap();
    transaction
        .set_checksum(&main, 0, Some(checksum.clone()))
        .unwrap();
    assert!(transaction.commit().is_ok());

    let over = package
        .content_control_snapshot_with_limits(PackageLimits {
            max_output_bytes: expected - 1,
            ..PackageLimits::default()
        })
        .unwrap();
    let mut transaction = over.edit().unwrap();
    transaction.set_checksum(&main, 0, Some(checksum)).unwrap();
    assert!(transaction.commit().is_err());
}

#[test]
fn every_story_limit_has_an_exact_and_one_over_fixture() {
    let mut package = strict_story_package(ct::WML_DOCUMENT_MAIN);
    let main = PackURI::new("/word/document.xml").unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&main)?;
            let mut source = b"<!--prolog-->".to_vec();
            source.extend_from_slice(part.blob());
            part.set_blob(source);
            Ok::<_, Error>(())
        })
        .unwrap();
    let inventory = package.story_inventory().unwrap();
    let max_story_bytes = inventory
        .stories()
        .iter()
        .map(|story| story.source().len())
        .max()
        .unwrap();
    let exact = StoryLimits {
        max_package_parts: package.opc_package().part_count(),
        max_stories: inventory.stories().len(),
        max_story_bytes,
        max_total_story_bytes: inventory.total_story_bytes(),
        max_relationships_per_owner: 6,
        max_total_relationships: 6,
        max_topology_bytes: inventory.topology().as_bytes().len(),
        max_xml_prolog_events: 2,
    };
    assert!(package.story_inventory_with_limits(exact).is_ok());

    for limited in [
        StoryLimits {
            max_package_parts: exact.max_package_parts - 1,
            ..exact
        },
        StoryLimits {
            max_stories: exact.max_stories - 1,
            ..exact
        },
        StoryLimits {
            max_story_bytes: exact.max_story_bytes - 1,
            ..exact
        },
        StoryLimits {
            max_total_story_bytes: exact.max_total_story_bytes - 1,
            ..exact
        },
        StoryLimits {
            max_relationships_per_owner: exact.max_relationships_per_owner - 1,
            ..exact
        },
        StoryLimits {
            max_total_relationships: exact.max_total_relationships - 1,
            ..exact
        },
        StoryLimits {
            max_topology_bytes: exact.max_topology_bytes - 1,
            ..exact
        },
        StoryLimits {
            max_xml_prolog_events: exact.max_xml_prolog_events - 1,
            ..exact
        },
    ] {
        assert!(package.story_inventory_with_limits(limited).is_err());
    }
}

#[test]
fn package_patch_rejects_stale_story_bytes_without_partial_publication() {
    let payload = b"<root xmlns='urn:test'><value>stable</value></root>";
    let mut package = bound_package(payload);
    let snapshot = package.content_control_snapshot().unwrap();
    let main = PackURI::new("/word/document.xml").unwrap();
    let mut transaction = snapshot.edit().unwrap();
    transaction
        .set_checksum(
            &main,
            1,
            Some(Checksum::compute(payload, &Limits::default()).unwrap()),
        )
        .unwrap();
    let commit = transaction.commit().unwrap();

    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&main)?;
            let mut changed = part.blob().to_vec();
            let insertion = changed.len() - "</w:document>".len();
            changed.splice(insertion..insertion, b"<!--stale-->".iter().copied());
            part.set_blob(changed);
            Ok::<_, Error>(())
        })
        .unwrap();
    let stale_source = package
        .opc_package()
        .get_part(&main)
        .unwrap()
        .blob()
        .to_vec();
    assert!(package.apply_content_controls(&commit).is_err());
    assert_eq!(
        package.opc_package().get_part(&main).unwrap().blob(),
        stale_source
    );
}
