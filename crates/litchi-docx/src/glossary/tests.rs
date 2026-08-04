use super::codec::*;
use super::graph::*;
use super::model::*;
use super::package::*;
use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use std::sync::Arc;

fn fixture(bytes: &[u8]) -> (Catalog, Conformance) {
    let package = OpcPackage::from_bytes(bytes).expect("package");
    load(&package).expect("load").expect("glossary")
}

fn package_for(conformance: Conformance) -> OpcPackage {
    let mut package = OpcPackage::new();
    let (word, office_document) = match conformance {
        Conformance::Transitional => (W, rt::OFFICE_DOCUMENT),
        Conformance::Strict => (WS, rt::STRICT_OFFICE_DOCUMENT),
    };
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").expect("URI"),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            format!(r#"<w:document xmlns:w="{word}"><w:body/></w:document>"#).into_bytes(),
        )))
        .expect("main part");
    package.relate_to("word/document.xml", office_document);
    package
}

fn package() -> OpcPackage {
    package_for(Conformance::Transitional)
}

fn add_empty_glossary(package: &mut OpcPackage, conformance: Conformance) -> PackURI {
    let root = PackURI::new("/word/glossary/document.xml").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            root.clone(),
            CT.to_owned(),
            write(&Catalog::new(), conformance).unwrap(),
        )))
        .unwrap();
    let main = package.main_document_part().unwrap().partname().clone();
    package
        .get_part_mut(&main)
        .unwrap()
        .rels_mut()
        .get_or_add(conformance.glossary_relationship(), "glossary/document.xml");
    root
}

fn mark_signed(package: &mut OpcPackage) {
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )))
        .unwrap();
    package.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rSignature".to_owned(),
        false,
    );
}

fn entry(name: &str) -> Entry {
    Entry::new(
        name,
        format!(r#"<w:docPartBody xmlns:w="{W}"><w:p/></w:docPartBody>"#).into_bytes(),
    )
    .expect("entry")
}

#[test]
fn poi_placeholders_and_strict_roundtrip() {
    let (catalog, _) = fixture(include_bytes!(
        "../../../../test-data/poi/test-data/document/Bug54849.docx"
    ));
    assert_eq!(catalog.len(), 3);
    let props = catalog.at(0).unwrap().props().unwrap();
    assert_eq!(props.kinds, Kind::SDT_PLACEHOLDER);
    assert_eq!(
        props.category.as_ref().unwrap().gallery().as_str(),
        "placeholder"
    );
    let xml = write(&catalog, Conformance::Strict).unwrap();
    assert_eq!(read(&xml).unwrap().0.len(), 3);
}

#[test]
fn libreoffice_multiple_autotext_and_tables_stay_inert() {
    let (catalog, _) = fixture(include_bytes!(
        "../../../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/autotext-multiple.dotx"
    ));
    assert_eq!(catalog.len(), 3);
    assert_eq!(catalog.at(0).unwrap().name(), Some("Multiple"));
    let body = std::str::from_utf8(catalog.at(2).unwrap().body().unwrap()).unwrap();
    assert!(body.contains("w:tbl"));
    assert!(body.contains("jksdjkfdskjfds"));
}

#[test]
fn libreoffice_empty_glossary_is_valid() {
    let (catalog, conformance) = fixture(include_bytes!(
        "../../../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx"
    ));
    assert!(catalog.is_empty());
    assert!(write(&catalog, conformance).is_ok());
}

#[test]
fn xsd_all_is_order_independent_and_last_duplicate_wins() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:description w:val="first"/><w:name w:val="First"/><w:types><w:type w:val="normal"/></w:types><w:name w:val="Last"/><w:description w:val="last"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (catalog, _) = read(xml.as_bytes()).unwrap();
    let entry = catalog.at(0).unwrap();
    assert_eq!(entry.name(), Some("Last"));
    assert_eq!(entry.props().unwrap().description.as_deref(), Some("last"));
    assert_eq!(entry.props().unwrap().kinds, Kind::NORMAL);
}

#[test]
fn compact_option_flags_collapse_duplicates_and_reject_unknown_or_empty_values() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Flags"/><w:types><w:type w:val="normal"/><w:type w:val="toolbar"/></w:types><w:behaviors><w:behavior w:val="content"/><w:behavior w:val="pg"/></w:behaviors></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (catalog, conformance) = read(xml.as_bytes()).unwrap();
    let props = catalog.at(0).unwrap().props().unwrap();
    assert_eq!(props.kinds, Kind::NORMAL | Kind::TOOLBAR);
    assert_eq!(props.inserts, Insert::CONTENT | Insert::PAGE);
    assert_eq!(
        read(&write(&catalog, conformance).unwrap()).unwrap().0,
        catalog
    );

    let duplicate = xml.replace(
        r#"<w:type w:val="toolbar"/>"#,
        r#"<w:type w:val="normal"/>"#,
    );
    let duplicate_catalog = read(duplicate.as_bytes()).unwrap().0;
    assert_eq!(
        duplicate_catalog.at(0).unwrap().props().unwrap().kinds,
        Kind::NORMAL
    );

    let word_empty_types = xml.replace(
        r#"<w:types><w:type w:val="normal"/><w:type w:val="toolbar"/></w:types>"#,
        "<w:types/>",
    );
    let (word_catalog, word_conformance) = read(word_empty_types.as_bytes()).unwrap();
    assert!(
        word_catalog
            .at(0)
            .unwrap()
            .props()
            .unwrap()
            .kinds
            .is_empty()
    );
    assert!(write(&word_catalog, word_conformance).is_ok());

    let empty_behaviors = xml.replace(
        r#"<w:behavior w:val="content"/><w:behavior w:val="pg"/>"#,
        "",
    );
    assert!(read(empty_behaviors.as_bytes()).is_err());

    let mut props = Props::new(Name::new("Unknown flag").unwrap());
    props.kinds = Kind::from_bits_retain(0x80);
    assert!(entry("Unknown flag").with_props(props).is_err());
}

#[test]
fn xml_forbidden_metadata_is_rejected_without_catalog_mutation() {
    assert!(Name::new("bad\0name").is_err());
    assert!(Name::new("bad\u{fffe}name").is_err());
    assert!(Category::new("bad\0category", Gallery::new("autoTxt").unwrap()).is_err());

    let mut catalog = Catalog::new();
    catalog.add(entry("Keep")).unwrap();
    let before = catalog.clone();
    let mut invalid = Props::new(Name::new("Invalid").unwrap());
    invalid.description = Some("bad\0description".to_owned());
    assert!(entry("Invalid").with_props(invalid).is_err());
    assert_eq!(catalog, before);
}

#[test]
fn invalid_earlier_duplicates_are_ignored_before_last_value_validation() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name/><w:guid w:val="invalid"/><w:types><w:type w:val="invalid"/></w:types><w:name w:val="Last"/><w:guid w:val="{{12345678-1234-4ABC-8DEF-1234567890AB}}"/><w:types><w:type w:val="normal"/></w:types></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (catalog, _) = read(xml.as_bytes()).unwrap();
    let props = catalog.at(0).unwrap().props().unwrap();
    assert_eq!(props.name().unwrap().as_str(), "Last");
    assert_eq!(
        props.id.as_ref().unwrap().as_str(),
        "{12345678-1234-4ABC-8DEF-1234567890AB}"
    );
    assert_eq!(props.kinds, Kind::NORMAL);
}

#[test]
fn mixed_dialects_and_foreign_typed_qnames_are_rejected() {
    let mixed = format!(
        r#"<w:glossaryDocument xmlns:w="{W}" xmlns:s="{WS}"><w:docParts><w:docPart><w:docPartPr><s:name s:val="Mixed"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    assert!(read(mixed.as_bytes()).is_err());

    let spoofed = format!(
        r#"<w:glossaryDocument xmlns:w="{W}" xmlns:u="urn:spoof"><w:docParts><w:docPart><u:docPartPr><w:name w:val="Spoof"/></u:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    assert!(read(spoofed.as_bytes()).is_err());
}

#[test]
fn duplicate_producer_names_are_readable_but_semantically_ambiguous() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Résumé"/></w:docPartPr></w:docPart><w:docPart><w:docPartPr><w:name w:val="RÉSUMÉ"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (catalog, _) = read(xml.as_bytes()).unwrap();
    assert!(catalog.get("résumé").is_err());
    assert_eq!(catalog.at(0).unwrap().name(), Some("Résumé"));
    assert_eq!(catalog.at(1).unwrap().name(), Some("RÉSUMÉ"));
}

#[test]
fn selectors_support_atomic_rename_replacement_and_numeric_ambiguity_repair() {
    let mut catalog = Catalog::new();
    catalog.add(entry("Alpha")).unwrap();
    catalog.add(entry("Beta")).unwrap();
    assert_eq!(catalog.get("ALPHA").unwrap().unwrap().name(), Some("Alpha"));

    let previous = catalog.replace("alpha", entry("Gamma")).unwrap().unwrap();
    assert_eq!(previous.name(), Some("Alpha"));
    assert!(catalog.get("alpha").unwrap().is_none());
    let body_pointer = catalog
        .get("gamma")
        .unwrap()
        .unwrap()
        .body()
        .unwrap()
        .as_ptr();
    assert!(
        catalog
            .rename("GAMMA", Name::new("Epsilon").unwrap())
            .unwrap()
    );
    assert_eq!(
        catalog
            .get("epsilon")
            .unwrap()
            .unwrap()
            .body()
            .unwrap()
            .as_ptr(),
        body_pointer
    );
    let before = catalog.clone();
    assert!(
        catalog
            .rename("epsilon", Name::new("Beta").unwrap())
            .is_err()
    );
    assert_eq!(catalog, before);
    assert!(
        catalog
            .replace("missing", entry("Unused"))
            .unwrap()
            .is_none()
    );
    assert!(catalog.replace_at(usize::MAX, entry("Unused")).is_err());

    let duplicate = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Résumé"/></w:docPartPr></w:docPart><w:docPart><w:docPartPr><w:name w:val="RÉSUMÉ"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (mut ambiguous, conformance) = read(duplicate.as_bytes()).unwrap();
    assert!(ambiguous.get("résumé").is_err());
    assert!(
        ambiguous
            .rename_at(1, Name::new("Repaired").unwrap())
            .unwrap()
    );
    assert_eq!(
        ambiguous.get("résumé").unwrap().unwrap().name(),
        Some("Résumé")
    );
    assert_eq!(
        ambiguous.get("repaired").unwrap().unwrap().name(),
        Some("Repaired")
    );
    let plan = plan_write(&ambiguous, conformance).unwrap();
    let index = conformance.index();
    let mut shell = XmlSize::default();
    write_catalog_open(&mut shell, conformance).unwrap();
    shell.push_str("<w:docParts></w:docParts>").unwrap();
    shell.push_str("</w:glossaryDocument>").unwrap();
    assert_eq!(
        plan.bytes,
        shell.bytes + ambiguous.state.background_bytes[index] + ambiguous.state.entry_bytes[index]
    );
}

#[test]
fn empty_producer_names_and_categories_survive_unrelated_crud() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val=""/><w:category><w:name w:val=""/><w:gallery w:val="autoTxt"/></w:category></w:docPartPr></w:docPart><w:docPart><w:docPartPr><w:name w:val="  "/><w:category><w:name w:val="  "/><w:gallery w:val="autoTxt"/></w:category></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (mut catalog, conformance) = read(xml.as_bytes()).unwrap();
    assert_eq!(catalog.at(0).unwrap().name(), Some(""));
    assert_eq!(catalog.at(1).unwrap().name(), Some("  "));
    catalog.add(entry("Authored")).unwrap();
    let output = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
    assert!(output.contains(r#"<w:name w:val=""/>"#), "{output}");
    assert!(output.contains(r#"<w:name w:val="  "/>"#), "{output}");
    assert!(Name::new("").is_err());
    assert!(Category::new(" ", Gallery::new("autoTxt").unwrap()).is_err());
}

#[test]
fn carriage_returns_remain_distinct_from_line_feeds() {
    let body = format!(
        r#"<w:docPartBody xmlns:w="{W}"><w:p><w:r><w:t>A&#xD;B
C</w:t></w:r></w:p></w:docPartBody>"#
    );
    let mut catalog = Catalog::new();
    catalog
        .add(Entry::new("Line endings", body.into_bytes()).unwrap())
        .unwrap();
    let output = String::from_utf8(write(&catalog, Conformance::Transitional).unwrap()).unwrap();
    assert!(output.contains("A&#xD;B\nC"), "{output}");
    let (round_trip, _) = read(output.as_bytes()).unwrap();
    let body = std::str::from_utf8(round_trip.at(0).unwrap().body().unwrap()).unwrap();
    assert!(body.contains("A&#xD;B\nC"), "{body}");

    let forbidden = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="bad&#x0;name"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    assert!(read(forbidden.as_bytes()).is_err());
}

#[test]
fn opaque_namespace_mapping_does_not_rewrite_literal_values() {
    let literal = W;
    let body = format!(
        r#"<w:docPartBody xmlns:w="{W}" xmlns:u="urn:test"><w:p u:value="prefix-{literal}-suffix"><w:r><w:t>{literal}</w:t></w:r></w:p></w:docPartBody>"#
    );
    let catalog = Catalog {
        background: None,
        background_refs: Arc::from([]),
        background_lineage: None,
        entries: vec![Entry::new("Literal", body.into_bytes()).unwrap()],
        binding: None,
        state: CatalogState::default(),
    };
    let strict = String::from_utf8(write(&catalog, Conformance::Strict).unwrap()).unwrap();
    assert!(strict.contains(&format!(r#"u:value="prefix-{W}-suffix""#)));
    assert!(strict.contains(&format!(">{W}</w:t>")), "{strict}");
    assert!(strict.contains(WS));
    read(strict.as_bytes()).unwrap();
}

#[test]
fn strict_mce_and_malformed_input_are_bounded() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><w:docParts><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><w:docPart><w:docPartPr><w:name w:val="MCE"/><w:behaviors><w:behavior w:val="p"/></w:behaviors></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></mc:Fallback></mc:AlternateContent></w:docParts></w:glossaryDocument>"#
    );
    assert_eq!(read(xml.as_bytes()).unwrap().0.len(), 1);
    for bad in [
        format!(
            r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:behaviors><w:behavior w:val="run"/></w:behaviors></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        ),
        format!(r#"<!DOCTYPE x><w:glossaryDocument xmlns:w="{W}"/>"#),
    ] {
        assert!(read(bad.as_bytes()).is_err(), "accepted {bad}");
    }
}

#[test]
fn producer_optional_metadata_is_readable_and_retained() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:style w:val="ProducerStyle"/><w:guid/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart><w:docPart><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (mut catalog, conformance) = read(xml.as_bytes()).unwrap();
    assert_eq!(catalog.len(), 2);
    assert!(catalog.at(0).unwrap().name().is_none());
    assert!(catalog.at(0).unwrap().props().unwrap().id.is_none());
    assert!(catalog.at(1).unwrap().props().is_none());

    catalog.add(entry("Authored")).unwrap();
    let rewritten = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
    assert!(rewritten.contains("<w:guid/>"), "{rewritten}");
    assert_eq!(read(rewritten.as_bytes()).unwrap().0.len(), 3);
}

#[test]
fn unrelated_crud_retains_ignorable_producer_content() {
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:docParts><w:docPart mc:Ignorable="w15"><w:docPartPr><w:name w:val="Keep"/></w:docPartPr><w:docPartBody><w:p><w15:producerExtension w15:value="preserve-me"/></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (mut catalog, conformance) = read(xml.as_bytes()).unwrap();
    assert!(
        !std::str::from_utf8(catalog.at(0).unwrap().body().unwrap())
            .unwrap()
            .contains("producerExtension")
    );
    catalog.add(entry("Added")).unwrap();

    let rewritten = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
    assert!(rewritten.contains("producerExtension"), "{rewritten}");
    assert!(rewritten.contains("mc:Ignorable=\"w15\""), "{rewritten}");
    assert_eq!(read(rewritten.as_bytes()).unwrap().0.len(), 2);

    let keep = catalog.remove("Keep").unwrap().unwrap();
    let mut props = keep.props().unwrap().clone();
    props.description = Some("changed".to_owned());
    catalog.add(keep.with_props(props).unwrap()).unwrap();
    let targeted = String::from_utf8(write(&catalog, conformance).unwrap()).unwrap();
    assert!(!targeted.contains("producerExtension"), "{targeted}");
    assert!(targeted.contains(r#"w:description w:val="changed""#));
}

#[test]
fn strict_rejects_transitional_relationships_lexicals_and_vml() {
    for xml in [
        format!(
            r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:r="{R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Mixed"/></w:docPartPr><w:docPartBody><w:p r:embed="rId1"/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        ),
        format!(
            r#"<w:glossaryDocument xmlns:w="{WS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Lexical" w:decorated="on"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        ),
        format!(
            r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:v="{VML}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="VML"/></w:docPartPr><w:docPartBody><v:shape/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
        ),
    ] {
        assert!(read(xml.as_bytes()).is_err(), "accepted {xml}");
    }

    let transitional_vml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}" xmlns:v="{VML}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="VML"/></w:docPartPr><w:docPartBody><v:shape/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (catalog, conformance) = read(transitional_vml.as_bytes()).unwrap();
    assert_eq!(conformance, Conformance::Transitional);
    assert!(write(&catalog, Conformance::Transitional).is_ok());
    assert!(write(&catalog, Conformance::Strict).is_err());
}

#[test]
fn inherited_namespace_scope_is_emitted_once_per_opaque_root() {
    let declarations = (0..32)
        .map(|index| format!(r#" xmlns:u{index}="urn:{index}:{}""#, "x".repeat(256)))
        .collect::<String>();
    let children = "<w:r/>".repeat(1_000);
    let xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"{declarations}><w:docParts><w:docPart><w:docPartPr><w:name w:val="Bounded"/></w:docPartPr><w:docPartBody><w:p>{children}</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    let (catalog, _) = read(xml.as_bytes()).unwrap();
    let body = catalog.at(0).unwrap().body().unwrap();
    assert!(
        body.len() < 32 * 1024,
        "opaque body grew to {} bytes",
        body.len()
    );
}

#[test]
fn dense_common_prefix_mce_scope_merges_linearly() {
    let common = "a".repeat(256);
    let attribute = |index: usize, value: &str| {
        let local = format!("{common}{index:04}");
        Attr {
            q: format!("mc:{local}"),
            ns: Arc::from(MC),
            l: local,
            v: value.to_owned(),
        }
    };
    let mut output = (0..2_048)
        .map(|index| attribute(index, "old"))
        .collect::<Vec<_>>();
    let node = Node {
        q: "w:docParts".to_owned(),
        ns: Arc::from(W),
        l: "docParts".to_owned(),
        attrs: (0..2_048)
            .rev()
            .map(|index| attribute(index, "new"))
            .collect(),
        bindings: Arc::new(NamespaceFrame::default()),
        content: Vec::new(),
    };

    merge_mce_attributes(&mut output, &node).unwrap();
    assert_eq!(output.len(), 2_048);
    assert!(output.iter().all(|attribute| attribute.v == "new"));
}

#[test]
fn aggregate_projection_and_dense_content_are_bounded() {
    let declarations = (0..512)
        .map(|index| format!(r#" xmlns:u{index}="urn:{index}:{}""#, "x".repeat(128)))
        .collect::<String>();
    let entries = (0..600)
            .map(|index| {
                format!(
                    r#"<w:docPart><w:docPartPr><w:name w:val="Entry {index}"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart>"#
                )
            })
            .collect::<String>();
    let amplified = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"{declarations}><w:docParts>{entries}</w:docParts></w:glossaryDocument>"#
    );
    assert!(amplified.len() < MAX);
    let root = parse_dom(amplified.as_bytes()).unwrap();
    let error = project(&root).unwrap_err().to_string();
    assert!(error.contains("semantic XML exceeds"), "{error}");

    let dense = "x<!--c-->".repeat((MAX_DOM_CONTENT / 2) + 1);
    let dense = format!(
        r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Dense"/></w:docPartPr><w:docPartBody><w:p>{dense}</w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#
    );
    assert!(dense.len() < MAX);
    assert!(read(dense.as_bytes()).is_err());
}

#[test]
fn exact_noop_preserves_custom_root_bytes_and_signature() {
    let mut package = package();
    let mut catalog = Catalog::new();
    catalog.add(entry("Keep")).unwrap();
    let source = Arc::new(write(&catalog, Conformance::Transitional).unwrap());
    let custom = PackURI::new("/word/glossary/custom.xml").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new_shared(
            custom.clone(),
            CT.to_owned(),
            Arc::clone(&source),
        )))
        .unwrap();
    let main = package.main_document_part().unwrap().partname().clone();
    package
        .get_part_mut(&main)
        .unwrap()
        .rels_mut()
        .get_or_add(REL, "glossary/custom.xml");
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )))
        .unwrap();
    package.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rSignature".to_owned(),
        false,
    );
    assert!(package.is_signed());
    let (loaded, conformance) = load(&package).unwrap().unwrap();
    assert!(!put(&mut package, loaded, conformance).unwrap());
    assert_eq!(package.get_part(&custom).unwrap().blob(), source.as_slice());
    assert!(package.is_signed());
    assert_eq!(
        raw::load(&package).unwrap().unwrap().root_name(),
        custom.as_str()
    );
}

#[test]
fn semantic_create_allocates_around_unrelated_canonical_parts() {
    let mut package = package();
    let occupied = [
        "/word/glossary/document.xml",
        "/word/glossary/styles.xml",
        "/word/glossary/settings.xml",
        "/word/glossary/fontTable.xml",
        "/word/glossary/webSettings.xml",
    ];
    for (index, name) in occupied.iter().enumerate() {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(*name).unwrap(),
                "application/vnd.example.unrelated+xml".to_owned(),
                format!("<unrelated id=\"{index}\"/>").into_bytes(),
            )))
            .unwrap();
    }

    let mut catalog = Catalog::new();
    catalog.add(entry("Allocated")).unwrap();
    assert!(put(&mut package, catalog, Conformance::Transitional).unwrap());

    let graph = raw::load(&package).unwrap().unwrap();
    assert_eq!(graph.root_name(), "/word/glossary/document1.xml");
    let names = graph
        .parts
        .iter()
        .map(|part| part.name.as_str())
        .collect::<HashSet<_>>();
    assert_eq!(
        names,
        HashSet::from([
            "/word/glossary/styles1.xml",
            "/word/glossary/settings1.xml",
            "/word/glossary/fontTable1.xml",
            "/word/glossary/webSettings1.xml",
        ])
    );

    raw::remove(&mut package).unwrap().unwrap();
    for (index, name) in occupied.iter().enumerate() {
        assert_eq!(
            package
                .get_part(&PackURI::new(*name).unwrap())
                .unwrap()
                .blob(),
            format!("<unrelated id=\"{index}\"/>").as_bytes()
        );
    }
}

#[test]
fn repeated_changed_raw_graph_is_a_signature_preserving_noop() {
    let mut package = package();
    let mut initial = Catalog::new();
    initial.add(entry("Initial")).unwrap();
    raw::put(
        &mut package,
        &raw::Graph::new(initial, Conformance::Transitional),
    )
    .unwrap();

    let mut changed = raw::load(&package).unwrap().unwrap();
    changed.catalog.add(entry("Changed")).unwrap();
    assert!(raw::put(&mut package, &changed).unwrap());
    mark_signed(&mut package);
    assert!(package.is_signed());

    assert!(!raw::put(&mut package, &changed).unwrap());
    assert!(package.is_signed());
    assert!(
        raw::load(&package)
            .unwrap()
            .unwrap()
            .catalog
            .get("changed")
            .unwrap()
            .is_some()
    );
}

#[test]
fn signature_namespace_auxiliary_is_rejected_atomically() {
    let mut package = package();
    let mut existing = Catalog::new();
    existing.add(entry("Keep")).unwrap();
    raw::put(
        &mut package,
        &raw::Graph::new(existing, Conformance::Transitional),
    )
    .unwrap();
    mark_signed(&mut package);

    let mut invalid = raw::load(&package).unwrap().unwrap();
    invalid.rels.push(raw::Rel {
        id: "rIdSignatureImage".to_owned(),
        kind: rt::IMAGE.to_owned(),
        target: "/_XMLSIGNATURES/glossary.png".to_owned(),
        external: false,
    });
    assert!(
        raw::Part::new(
            "/_XMLSIGNATURES/glossary.png",
            "image/png",
            vec![0x89, b'P', b'N', b'G'],
        )
        .is_err()
    );
    invalid.parts.push(raw::Part::from_shared(
        "/_XMLSIGNATURES/glossary.png".to_owned(),
        "image/png".to_owned(),
        Arc::new(vec![0x89, b'P', b'N', b'G']),
        Vec::new(),
    ));

    assert!(raw::put(&mut package, &invalid).is_err());
    assert!(package.is_signed());
    assert!(
        raw::load(&package)
            .unwrap()
            .unwrap()
            .catalog
            .get("keep")
            .unwrap()
            .is_some()
    );
}

#[test]
fn reserved_parts_and_invalid_raw_metadata_are_rejected_atomically() {
    assert!(raw::Part::new("/[Content_Types].xml", "application/xml", Vec::new()).is_err());
    assert!(
        raw::Part::new(
            "/word/glossary/media/image.png",
            "image/png;broken",
            Vec::new()
        )
        .is_err()
    );
    for (index, content_type) in [
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
        ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
        ct::OPC_RELATIONSHIPS,
    ]
    .into_iter()
    .enumerate()
    {
        assert!(
            raw::Part::new(
                format!("/word/glossary/reserved-{index}.bin"),
                content_type,
                Vec::new(),
            )
            .is_err()
        );

        let mut candidate = package();
        mark_signed(&mut candidate);
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdReservedType".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: "reserved.bin".to_owned(),
            external: false,
        });
        graph.parts.push(raw::Part::from_shared(
            "/word/glossary/reserved.bin".to_owned(),
            content_type.to_owned(),
            Arc::new(Vec::new()),
            Vec::new(),
        ));
        assert!(raw::put(&mut candidate, &graph).is_err());
        assert!(candidate.is_signed());
        assert!(raw::load(&candidate).unwrap().is_none());
    }

    let mut package = package();
    mark_signed(&mut package);
    let mut reserved = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    reserved.rels.push(raw::Rel {
        id: "rIdImage".to_owned(),
        kind: rt::IMAGE.to_owned(),
        target: "/[Content_Types].xml".to_owned(),
        external: false,
    });
    reserved.parts.push(raw::Part::from_shared(
        "/[Content_Types].xml".to_owned(),
        "image/png".to_owned(),
        Arc::new(vec![0x89, b'P', b'N', b'G']),
        Vec::new(),
    ));
    assert!(raw::put(&mut package, &reserved).is_err());
    assert!(package.is_signed());
    assert!(raw::load(&package).unwrap().is_none());

    let mut invalid_id = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    invalid_id.rels.push(raw::Rel {
        id: "bad id".to_owned(),
        kind: rt::HYPERLINK.to_owned(),
        target: "https://example.invalid".to_owned(),
        external: true,
    });
    assert!(raw::put(&mut package, &invalid_id).is_err());
    assert!(package.is_signed());
}

#[test]
fn graph_wide_relationship_limit_is_failure_atomic() {
    let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    for part_index in 0..2 {
        let name = format!("numbering{part_index}.xml");
        graph.rels.push(raw::Rel {
            id: format!("rIdNumbering{part_index}"),
            kind: rt::NUMBERING.to_owned(),
            target: name.clone(),
            external: false,
        });
        let mut part = raw::Part::new(
            format!("/word/glossary/{name}"),
            ct::WML_NUMBERING,
            format!(r#"<w:numbering xmlns:w="{W}"/>"#).into_bytes(),
        )
        .unwrap();
        for relationship_index in 0..(MAX_VALUES / 2) {
            part.rels.push(raw::Rel {
                id: format!("rIdLink{relationship_index}"),
                kind: rt::HYPERLINK.to_owned(),
                target: format!("https://example.invalid/{part_index}/{relationship_index}"),
                external: true,
            });
        }
        graph.parts.push(part);
    }

    let mut package = package();
    mark_signed(&mut package);
    let error = raw::put(&mut package, &graph).unwrap_err().to_string();
    assert!(error.contains("graph-wide relationship limit"), "{error}");
    assert!(package.is_signed());
    assert!(raw::load(&package).unwrap().is_none());
}

#[test]
fn shared_inbound_owned_parts_block_changed_store_and_remove() {
    let mut package = package();
    let mut catalog = Catalog::new();
    catalog.add(entry("Keep")).unwrap();
    put(&mut package, catalog.clone(), Conformance::Transitional).unwrap();
    let root = PackURI::new("/word/glossary/document.xml").unwrap();
    let mut referrer = BlobPart::new(
        PackURI::new("/word/header1.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml".into(),
        format!(r#"<w:hdr xmlns:w="{W}"/>"#).into_bytes(),
    );
    referrer
        .rels_mut()
        .get_or_add(rt::IMAGE, "glossary/document.xml");
    package.try_add_part(Box::new(referrer)).unwrap();

    let mut changed = catalog;
    changed.add(entry("Changed")).unwrap();
    assert!(put(&mut package, changed, Conformance::Transitional).is_err());
    assert!(remove(&mut package).is_err());
    assert!(package.get_part(&root).is_ok());
}

#[test]
fn auxiliary_relationships_must_match_the_package_conformance_family() {
    for (conformance, wrong_kind) in [
        (Conformance::Transitional, rt::STRICT_IMAGE),
        (Conformance::Strict, rt::IMAGE),
    ] {
        let mut package = package_for(conformance);
        let mut graph = raw::Graph::new(Catalog::new(), conformance);
        graph.rels.push(raw::Rel {
            id: "rIdImage".to_owned(),
            kind: wrong_kind.to_owned(),
            target: "media/image.bin".to_owned(),
            external: false,
        });

        let error = raw::put(&mut package, &graph).unwrap_err().to_string();
        assert!(error.contains("conformance"), "{error}");
        assert!(raw::load(&package).unwrap().is_none());
    }
}

#[test]
fn relationship_targets_use_exact_content_type_profiles() {
    let mut invalid_package = package();
    let mut invalid = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    invalid.rels.push(raw::Rel {
        id: "rIdStyles".to_owned(),
        kind: rt::STYLES.to_owned(),
        target: "styles.png".to_owned(),
        external: false,
    });
    invalid.parts.push(
        raw::Part::new(
            "/word/glossary/styles.png",
            "image/png",
            vec![0x89, b'P', b'N', b'G'],
        )
        .unwrap(),
    );
    assert!(raw::put(&mut invalid_package, &invalid).is_err());
    assert!(raw::load(&invalid_package).unwrap().is_none());

    let mut valid_package = package();
    let mut valid = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    valid.rels.push(raw::Rel {
        id: "rIdCustomizations".to_owned(),
        kind: CUSTOMIZATIONS_REL.to_owned(),
        target: "customizations.xml".to_owned(),
        external: false,
    });
    valid.parts.push(
        raw::Part::new(
            "/word/glossary/customizations.xml",
            CUSTOMIZATIONS_CT,
            Vec::new(),
        )
        .unwrap(),
    );
    assert!(raw::put(&mut valid_package, &valid).unwrap());
    assert_eq!(raw::load(&valid_package).unwrap().unwrap().parts.len(), 1);
}

#[test]
fn standards_relationship_matrix_round_trips_owned_parts_and_keeps_references() {
    fn relationship(id: &str, suffix: &str, target: &str, external: bool) -> raw::Rel {
        raw::Rel {
            id: id.to_owned(),
            kind: format!("{R}/{suffix}"),
            target: target.to_owned(),
            external,
        }
    }

    let body = format!(
        r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p r:dm="rIdDiagram"/></w:docPartBody>"#
    );
    let mut catalog = Catalog::new();
    catalog
        .add(Entry::new("Matrix", body.into_bytes()).unwrap())
        .unwrap();
    let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
    graph.rels = vec![
        relationship("rIdComments", "comments", "comments.xml", false),
        relationship("rIdSettings", "settings", "settings.xml", false),
        relationship("rIdFonts", "fontTable", "fontTable.xml", false),
        relationship("rIdWeb", "webSettings", "webSettings.xml", false),
        relationship("rIdDiagram", "diagramData", "diagram/data.xml", false),
        relationship("rIdCustomXml", "customXml", "customXml/item1.xml", false),
        relationship("rIdControl", "control", "activeX/activeX1.xml", false),
        relationship("rIdGenericControl", "control", "controls/vendor.bin", false),
        relationship("rIdChunk", "aFChunk", "chunks/chunk.html", false),
        relationship(
            "rIdReference",
            "hyperlink",
            "../../shared/reference.xml",
            false,
        ),
        relationship(
            "rIdExternalImage",
            "image",
            "https://example.invalid/image.png",
            true,
        ),
        raw::Rel {
            id: "rIdCustomizations".to_owned(),
            kind: CUSTOMIZATIONS_REL.to_owned(),
            target: "customizations.xml".to_owned(),
            external: false,
        },
    ];

    let mut comments = raw::Part::new(
        "/word/glossary/comments.xml",
        ct::WML_COMMENTS,
        format!(r#"<w:comments xmlns:w="{W}"/>"#).into_bytes(),
    )
    .unwrap();
    comments.rels = vec![
        relationship("rIdChart", "chart", "charts/chart1.xml", false),
        relationship(
            "rIdVideo",
            "video",
            "https://example.invalid/video.mp4",
            true,
        ),
    ];
    let mut chart = raw::Part::new(
        "/word/glossary/charts/chart1.xml",
        ct::DML_CHART,
        b"<c:chartSpace xmlns:c=\"urn:chart\"/>".to_vec(),
    )
    .unwrap();
    chart.rels.push(relationship(
        "rIdPackage",
        "package",
        "https://example.invalid/data.xlsx",
        true,
    ));
    chart.rels.extend([
        raw::Rel {
            id: "rIdChartStyle".to_owned(),
            kind: CHART_STYLE_REL.to_owned(),
            target: "style1.xml".to_owned(),
            external: false,
        },
        raw::Rel {
            id: "rIdChartColors".to_owned(),
            kind: CHART_COLOR_STYLE_REL.to_owned(),
            target: "colors1.xml".to_owned(),
            external: false,
        },
        relationship(
            "rIdChartShapes",
            "chartUserShapes",
            "userShapes1.xml",
            false,
        ),
        raw::Rel {
            id: "rIdChartStyle2012".to_owned(),
            kind: CHART_STYLE_REL_2012.to_owned(),
            target: "style1.xml".to_owned(),
            external: false,
        },
        raw::Rel {
            id: "rIdChartColors2012".to_owned(),
            kind: CHART_COLOR_STYLE_REL_2012.to_owned(),
            target: "colors1.xml".to_owned(),
            external: false,
        },
        relationship(
            "rIdThemeOverride",
            "themeOverride",
            "themeOverride1.xml",
            false,
        ),
    ]);
    let mut chart_shapes = raw::Part::new(
        "/word/glossary/charts/userShapes1.xml",
        ct::DML_CHARTSHAPES,
        Vec::new(),
    )
    .unwrap();
    chart_shapes.rels.push(relationship(
        "rIdLinkedImage",
        "image",
        "https://example.invalid/chart-image.png",
        true,
    ));
    chart_shapes.rels.extend([
        relationship("rIdChartBack", "chart", "chart1.xml", false),
        relationship(
            "rIdInkCustomXml",
            "customXml",
            "../customXml/item1.xml",
            false,
        ),
    ]);
    let mut custom_xml = raw::Part::new(
        "/word/glossary/customXml/item1.xml",
        "application/xml",
        b"<data/>".to_vec(),
    )
    .unwrap();
    custom_xml.rels.push(relationship(
        "rIdCustomXmlProps",
        "customXmlProps",
        "itemProps1.xml",
        false,
    ));
    let mut control = raw::Part::new(
        "/word/glossary/activeX/activeX1.xml",
        ACTIVE_X_DESCRIPTOR_CT,
        b"<ax:ocx xmlns:ax=\"urn:active-x\"/>".to_vec(),
    )
    .unwrap();
    control.rels.push(raw::Rel {
        id: "rIdControlBinary".to_owned(),
        kind: ACTIVE_X_BINARY_REL.to_owned(),
        target: "activeX1.bin".to_owned(),
        external: false,
    });
    let mut settings = raw::Part::new(
        "/word/glossary/settings.xml",
        ct::WML_SETTINGS,
        format!(r#"<w:settings xmlns:w="{W}"/>"#).into_bytes(),
    )
    .unwrap();
    settings.rels.push(relationship(
        "rIdTemplate",
        "attachedTemplate",
        "https://example.invalid/template.dotx",
        true,
    ));
    settings.rels.push(relationship(
        "rIdRecipientData",
        "recipientData",
        "recipientData.xml",
        false,
    ));
    let mut font_table = raw::Part::new(
        "/word/glossary/fontTable.xml",
        ct::WML_FONT_TABLE,
        format!(r#"<w:fonts xmlns:w="{W}"/>"#).into_bytes(),
    )
    .unwrap();
    font_table
        .rels
        .push(relationship("rIdFont", "font", "fonts/font.ttf", false));
    let mut web_settings = raw::Part::new(
        "/word/glossary/webSettings.xml",
        ct::WML_WEB_SETTINGS,
        format!(r#"<w:webSettings xmlns:w="{W}"/>"#).into_bytes(),
    )
    .unwrap();
    web_settings.rels.push(relationship(
        "rIdFrame",
        "frame",
        "https://example.invalid/frame.html",
        true,
    ));
    let mut diagram_data = raw::Part::new(
        "/word/glossary/diagram/data.xml",
        ct::DML_DIAGRAM_DATA,
        b"<d:dataModel xmlns:d=\"urn:diagram\"/>".to_vec(),
    )
    .unwrap();
    diagram_data.rels.push(raw::Rel {
        id: "rIdDrawing".to_owned(),
        kind: DIAGRAM_DRAWING_REL.to_owned(),
        target: "drawing.xml".to_owned(),
        external: false,
    });
    diagram_data.rels.push(relationship(
        "rIdDiagramLink",
        "hyperlink",
        "https://example.invalid/diagram",
        true,
    ));
    let mut customizations = raw::Part::new(
        "/word/glossary/customizations.xml",
        CUSTOMIZATIONS_CT,
        Vec::new(),
    )
    .unwrap();
    customizations.rels.push(raw::Rel {
        id: "rIdToolbars".to_owned(),
        kind: ATTACHED_TOOLBARS_REL.to_owned(),
        target: "attachedToolbars.bin".to_owned(),
        external: false,
    });
    graph.parts = vec![
        comments,
        chart,
        raw::Part::new(
            "/word/glossary/charts/style1.xml",
            CHART_STYLE_CT,
            Vec::new(),
        )
        .unwrap(),
        raw::Part::new(
            "/word/glossary/charts/colors1.xml",
            CHART_COLOR_STYLE_CT,
            Vec::new(),
        )
        .unwrap(),
        raw::Part::new(
            "/word/glossary/charts/themeOverride1.xml",
            ct::OFC_THEME_OVERRIDE,
            Vec::new(),
        )
        .unwrap(),
        chart_shapes,
        custom_xml,
        raw::Part::new(
            "/word/glossary/customXml/itemProps1.xml",
            litchi_ooxml_common::custom_xml::PROPS_CONTENT_TYPE,
            b"<ds:datastoreItem xmlns:ds=\"urn:custom-xml-props\"/>".to_vec(),
        )
        .unwrap(),
        control,
        raw::Part::new(
            "/word/glossary/activeX/activeX1.bin",
            ACTIVE_X_BINARY_CT,
            vec![0, 1, 2],
        )
        .unwrap(),
        raw::Part::new(
            "/word/glossary/controls/vendor.bin",
            "application/vnd.example.control",
            vec![9, 8, 7],
        )
        .unwrap(),
        settings,
        raw::Part::new(
            "/word/glossary/recipientData.xml",
            RECIPIENT_DATA_CT,
            Vec::new(),
        )
        .unwrap(),
        font_table,
        web_settings,
        diagram_data,
        raw::Part::new(
            "/word/glossary/diagram/drawing.xml",
            ct::DML_DIAGRAM_DRAWING,
            Vec::new(),
        )
        .unwrap(),
        raw::Part::new(
            "/word/glossary/chunks/chunk.html",
            "text/html",
            b"<p>chunk</p>".to_vec(),
        )
        .unwrap(),
        raw::Part::new("/word/glossary/fonts/font.ttf", FONT_TTF_CT, vec![0, 1, 2]).unwrap(),
        customizations,
        raw::Part::new(
            "/word/glossary/attachedToolbars.bin",
            ATTACHED_TOOLBARS_CT,
            Vec::new(),
        )
        .unwrap(),
    ];

    let mut package = package();
    let reference = PackURI::new("/shared/reference.xml").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            reference.clone(),
            "application/xml".to_owned(),
            b"<shared/>".to_vec(),
        )))
        .unwrap();
    assert!(raw::put(&mut package, &graph).unwrap());
    let loaded = raw::load(&package).unwrap().unwrap();
    assert!(
        !loaded
            .parts
            .iter()
            .any(|part| part.name == reference.as_str())
    );
    let removed = raw::remove(&mut package).unwrap().unwrap();
    assert_eq!(removed.parts.len(), graph.parts.len());
    assert_eq!(package.get_part(&reference).unwrap().blob(), b"<shared/>");
}

#[test]
fn relationship_roles_reject_invalid_modes_names_and_content_type_spoofing() {
    for (kind, target, external) in [
        (format!("{R}/theme"), "theme.xml", false),
        (format!("{R}/afChunk"), "chunk.html", false),
        (
            format!("{R}/styles"),
            "https://example.invalid/styles",
            true,
        ),
    ] {
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdInvalid".to_owned(),
            kind,
            target: target.to_owned(),
            external,
        });
        if !external {
            graph.parts.push(
                raw::Part::new(
                    format!("/word/glossary/{target}"),
                    if target == "theme.xml" {
                        ct::OFC_THEME
                    } else {
                        "text/html"
                    },
                    Vec::new(),
                )
                .unwrap(),
            );
        }
        assert!(raw::put(&mut package(), &graph).is_err());
    }

    let mut spoofed = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    spoofed.rels.push(raw::Rel {
        id: "rIdPackage".to_owned(),
        kind: format!("{R}/package"),
        target: "embedded.bin".to_owned(),
        external: false,
    });
    let mut embedded =
        raw::Part::new("/word/glossary/embedded.bin", ct::DML_CHART, Vec::new()).unwrap();
    embedded.rels.push(raw::Rel {
        id: "rIdShapes".to_owned(),
        kind: format!("{R}/chartUserShapes"),
        target: "shapes.xml".to_owned(),
        external: false,
    });
    spoofed.parts = vec![
        embedded,
        raw::Part::new("/word/glossary/shapes.xml", ct::DML_CHARTSHAPES, Vec::new()).unwrap(),
    ];
    assert!(raw::put(&mut package(), &spoofed).is_err());

    fn child_graph(
        root_kind: &str,
        parent_content_type: &str,
        child_kind: &str,
        child_content_type: Option<&str>,
        external: bool,
    ) -> raw::Graph {
        let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdParent".to_owned(),
            kind: root_kind.to_owned(),
            target: "parent.xml".to_owned(),
            external: false,
        });
        let mut parent =
            raw::Part::new("/word/glossary/parent.xml", parent_content_type, Vec::new()).unwrap();
        parent.rels.push(raw::Rel {
            id: "rIdChild".to_owned(),
            kind: child_kind.to_owned(),
            target: if external {
                "https://example.invalid/child".to_owned()
            } else {
                "child.bin".to_owned()
            },
            external,
        });
        graph.parts.push(parent);
        if let Some(content_type) = child_content_type {
            graph.parts.push(
                raw::Part::new("/word/glossary/child.bin", content_type, Vec::new()).unwrap(),
            );
        }
        graph
    }

    for invalid in [
        child_graph(
            &format!("{R}/customXml"),
            "application/xml",
            &format!("{R}/customXmlProps"),
            Some("application/xml"),
            false,
        ),
        child_graph(
            &format!("{R}/chart"),
            ct::DML_CHART,
            CHART_STYLE_REL,
            Some("application/xml"),
            false,
        ),
        child_graph(
            &format!("{R}/control"),
            ACTIVE_X_DESCRIPTOR_CT,
            ACTIVE_X_BINARY_REL,
            Some(ACTIVE_X_BINARY_CT),
            true,
        ),
        child_graph(
            &format!("{R}/control"),
            "application/vnd.example.control",
            ACTIVE_X_BINARY_REL,
            Some(ACTIVE_X_BINARY_CT),
            false,
        ),
        child_graph(
            &format!("{R}/customXml"),
            "application/xml",
            &format!("{R}/hyperlink"),
            None,
            true,
        ),
    ] {
        assert!(raw::put(&mut package(), &invalid).is_err());
    }
}

#[test]
fn semantic_catalogs_cannot_rebind_colliding_relationship_ids() {
    fn picture_graph(target: &str, payload: &[u8]) -> raw::Graph {
        let body = format!(
            r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r><w:drawing r:embed="rIdImage"/></w:r></w:p></w:docPartBody>"#
        );
        let mut catalog = Catalog::new();
        catalog
            .add(Entry::new("Picture", body.into_bytes()).unwrap())
            .unwrap();
        let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
        graph.rels.push(raw::Rel {
            id: "rIdImage".to_owned(),
            kind: rt::IMAGE.to_owned(),
            target: format!("media/{target}.png"),
            external: false,
        });
        graph.parts.push(
            raw::Part::new(
                format!("/word/glossary/media/{target}.png"),
                "image/png",
                payload.to_vec(),
            )
            .unwrap(),
        );
        graph
    }

    let source_graph = picture_graph("source", b"source image");
    let destination_graph = picture_graph("destination", b"destination image");
    let mut source = package();
    let mut destination = package();
    raw::put(&mut source, &source_graph).unwrap();
    raw::put(&mut destination, &destination_graph).unwrap();

    let (mut source_catalog, _) = load(&source).unwrap().unwrap();
    let source_entry = source_catalog.remove("picture").unwrap().unwrap();
    let (mut destination_catalog, _) = load(&destination).unwrap().unwrap();
    let before = destination_catalog.clone();
    assert!(
        destination_catalog
            .replace("picture", source_entry)
            .is_err()
    );
    assert_eq!(destination_catalog, before);

    let destination_entry = destination_catalog.remove("picture").unwrap().unwrap();
    let rebound_body = format!(
        r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}"><w:p r:embed="rIdImage"/><w:p/></w:docPartBody>"#
    );
    let destination_entry = destination_entry
        .with_body(rebound_body.into_bytes())
        .unwrap();
    assert!(destination_catalog.add(destination_entry).is_err());
    let background = format!(r#"<w:background xmlns:w="{W}" xmlns:r="{R}" r:id="rIdImage"/>"#);
    assert!(
        destination_catalog
            .set_background(background.into_bytes())
            .is_err()
    );
    assert!(destination_catalog.background().is_none());

    let (foreign, conformance) = load(&source).unwrap().unwrap();
    assert!(put(&mut destination, foreign, conformance).is_err());
    let retained = raw::load(&destination).unwrap().unwrap();
    assert_eq!(retained.parts[0].data(), b"destination image");
}

#[test]
fn inactive_mce_relationship_ids_remain_bound_to_their_physical_graph() {
    let body = format!(
        r#"<w:docPartBody xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="{MC}" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><w:p r:embed="rIdInactiveImage"/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent></w:docPartBody>"#
    );
    let mut catalog = Catalog::new();
    catalog
        .add(Entry::new("Inactive reference", body.into_bytes()).unwrap())
        .unwrap();
    let mut graph = raw::Graph::new(catalog, Conformance::Transitional);
    graph.rels.push(raw::Rel {
        id: "rIdInactiveImage".to_owned(),
        kind: rt::IMAGE.to_owned(),
        target: "media/inactive.png".to_owned(),
        external: false,
    });
    graph.parts.push(
        raw::Part::new(
            "/word/glossary/media/inactive.png",
            "image/png",
            b"inactive image".to_vec(),
        )
        .unwrap(),
    );

    let mut source = package();
    raw::put(&mut source, &graph).unwrap();
    let (mut loaded, conformance) = load(&source).unwrap().unwrap();
    assert!(
        !std::str::from_utf8(loaded.at(0).unwrap().body().unwrap())
            .unwrap()
            .contains("rIdInactiveImage")
    );
    let mut destination = package();
    assert!(put(&mut destination, loaded.clone(), conformance).is_err());
    assert!(raw::load(&destination).unwrap().is_none());

    loaded.add(entry("Unrelated")).unwrap();
    assert!(put(&mut source, loaded, conformance).unwrap());
    let preserved = raw::load(&source).unwrap().unwrap();
    assert!(
        String::from_utf8(write(&preserved.catalog, conformance).unwrap())
            .unwrap()
            .contains("rIdInactiveImage")
    );
    assert!(
        preserved
            .parts
            .iter()
            .any(|part| part.data() == b"inactive image")
    );
}

#[test]
fn complete_graph_follows_valid_parts_outside_the_conventional_directory() {
    let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    graph.rels.push(raw::Rel {
        id: "rIdImage".to_owned(),
        kind: rt::IMAGE.to_owned(),
        target: "../../assets/glossary.png".to_owned(),
        external: false,
    });
    graph.parts.push(
        raw::Part::new(
            "/assets/glossary.png",
            "image/png",
            vec![0x89, b'P', b'N', b'G'],
        )
        .unwrap(),
    );
    let mut package = package();
    assert!(raw::put(&mut package, &graph).unwrap());
    let removed = raw::remove(&mut package).unwrap().unwrap();
    assert_eq!(removed.parts[0].name, "/assets/glossary.png");
    assert!(
        package
            .get_part(&PackURI::new("/assets/glossary.png").unwrap())
            .is_err()
    );
}

#[test]
fn raw_remove_returns_a_graph_that_raw_put_can_restore() {
    let payload = vec![0xA5; MAX + 1];
    let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    graph.rels.push(raw::Rel {
        id: "rIdImage".to_owned(),
        kind: rt::IMAGE.to_owned(),
        target: "media/large.bin".to_owned(),
        external: false,
    });
    graph
        .parts
        .push(raw::Part::new("/word/glossary/media/large.bin", "image/png", payload).unwrap());

    let mut source = package();
    assert!(raw::put(&mut source, &graph).unwrap());
    let removed = raw::remove(&mut source).unwrap().unwrap();
    let mut destination = package();
    assert!(raw::put(&mut destination, &removed).unwrap());
    assert_eq!(
        raw::load(&destination).unwrap().unwrap().parts[0]
            .data()
            .len(),
        MAX + 1
    );
}

#[test]
fn raw_transfer_preserves_the_main_owner_relationship_metadata() {
    let mut source = package();
    let root = PackURI::new("/word/glossary/custom.xml").unwrap();
    source
        .try_add_part(Box::new(BlobPart::new(
            root,
            CT.to_owned(),
            write(&Catalog::new(), Conformance::Transitional).unwrap(),
        )))
        .unwrap();
    let main = source.main_document_part().unwrap().partname().clone();
    source
        .get_part_mut(&main)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            REL.to_owned(),
            "glossary/custom.xml".to_owned(),
            "rCustomGlossary".to_owned(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();

    let graph = raw::remove(&mut source).unwrap().unwrap();
    assert_eq!(graph.owner_relationship_id(), Some("rCustomGlossary"));
    assert_eq!(graph.owner_target(), Some("glossary/custom.xml"));
    let mut destination = package();
    raw::put(&mut destination, &graph).unwrap();
    let restored = raw::load(&destination).unwrap().unwrap();
    assert_eq!(restored.owner_relationship_id(), Some("rCustomGlossary"));
    assert_eq!(restored.owner_target(), Some("glossary/custom.xml"));
}

#[test]
fn raw_transfer_rebases_owner_target_and_allocates_a_free_id() {
    let mut source = package();
    let root = PackURI::new("/word/glossary/custom.xml").unwrap();
    source
        .try_add_part(Box::new(BlobPart::new(
            root,
            CT.to_owned(),
            write(&Catalog::new(), Conformance::Transitional).unwrap(),
        )))
        .unwrap();
    let source_main = source.main_document_part().unwrap().partname().clone();
    source
        .get_part_mut(&source_main)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            REL.to_owned(),
            "glossary/custom.xml".to_owned(),
            "rCustomGlossary".to_owned(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();
    let graph = raw::remove(&mut source).unwrap().unwrap();

    let mut destination = OpcPackage::new();
    let destination_main = PackURI::new("/custom/main.xml").unwrap();
    destination
        .try_add_part(Box::new(BlobPart::new(
            destination_main.clone(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#).into_bytes(),
        )))
        .unwrap();
    destination.relate_to("custom/main.xml", rt::OFFICE_DOCUMENT);
    destination
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/unrelated.bin").unwrap(),
            "application/octet-stream".to_owned(),
            vec![1, 2, 3],
        )))
        .unwrap();
    destination
        .get_part_mut(&destination_main)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:unrelated".to_owned(),
            "unrelated.bin".to_owned(),
            "rCustomGlossary".to_owned(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();

    assert!(raw::put(&mut destination, &graph).unwrap());
    let restored = raw::load(&destination).unwrap().unwrap();
    assert_eq!(restored.root_name(), "/word/glossary/custom.xml");
    assert_eq!(restored.owner_main_part(), Some("/custom/main.xml"));
    assert_eq!(restored.owner_target(), Some("../word/glossary/custom.xml"));
    assert_ne!(restored.owner_relationship_id(), Some("rCustomGlossary"));
    assert_eq!(
        destination
            .get_part(&destination_main)
            .unwrap()
            .rels()
            .get("rCustomGlossary")
            .unwrap()
            .reltype(),
        "urn:unrelated"
    );
    mark_signed(&mut destination);
    assert!(!raw::put(&mut destination, &graph).unwrap());
    assert!(destination.is_signed());
    let mut bytes = Vec::new();
    destination.to_stream(&mut bytes).unwrap();
    let reopened = OpcPackage::from_bytes(&bytes).unwrap();
    assert_eq!(
        raw::load(&reopened).unwrap().unwrap().owner_main_part(),
        Some("/custom/main.xml")
    );
}

#[test]
fn ownership_closure_resolves_case_varied_relationship_targets() {
    let mut package = package();
    let root = add_empty_glossary(&mut package, Conformance::Transitional);
    let auxiliary = PackURI::new("/word/glossary/media/image.bin").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            auxiliary.clone(),
            "image/png".to_owned(),
            vec![1, 2, 3],
        )))
        .unwrap();
    package
        .get_part_mut(&root)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "/WORD/GLOSSARY/MEDIA/IMAGE.BIN".to_owned(),
            "rIdImage".to_owned(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();

    let graph = raw::remove(&mut package).unwrap().unwrap();
    assert!(
        graph
            .parts
            .iter()
            .any(|part| part.name == auxiliary.as_str())
    );
    assert!(package.get_part(&auxiliary).is_err());
    assert!(raw::load(&package).unwrap().is_none());
}

#[test]
fn unowned_parts_in_the_conventional_glossary_directory_are_preserved() {
    for with_root in [false, true] {
        let mut package = package();
        if with_root {
            add_empty_glossary(&mut package, Conformance::Transitional);
        }
        let orphan = PackURI::new("/word/glossary/media/orphan.bin").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                orphan.clone(),
                "application/octet-stream".to_owned(),
                vec![4, 5, 6],
            )))
            .unwrap();

        assert_eq!(raw::load(&package).unwrap().is_some(), with_root);
        assert_eq!(raw::remove(&mut package).unwrap().is_some(), with_root);
        assert_eq!(package.get_part(&orphan).unwrap().blob(), [4, 5, 6]);
    }
}

#[test]
fn semantic_payload_limits_fail_before_xml_parsing() {
    let oversized = vec![b'x'; MAX + 1];
    let entry_error = Entry::new("Oversized", oversized).unwrap_err().to_string();
    assert!(
        entry_error.contains("payload exceeds 32 MiB"),
        "{entry_error}"
    );

    let mut catalog = Catalog::new();
    let background_error = catalog
        .set_background(vec![b'x'; MAX + 1])
        .unwrap_err()
        .to_string();
    assert!(
        background_error.contains("payload exceeds 32 MiB"),
        "{background_error}"
    );
    assert!(catalog.background().is_none());
}

#[test]
fn aggregate_catalog_budget_is_checked_atomically() {
    let mut catalog = Catalog::new();
    for index in 0..5 {
        let name = format!("Large {index}");
        let props = Props {
            description: Some("\"".repeat(MAX_STRING)),
            ..Props::new(Name::new(&name).unwrap())
        };
        catalog
            .add(entry(&name).with_props(props).unwrap())
            .unwrap();
    }
    let before = catalog.len();
    let rejected_name = "Too large";
    let rejected_props = Props {
        description: Some("\"".repeat(MAX_STRING)),
        ..Props::new(Name::new(rejected_name).unwrap())
    };
    let error = catalog
        .add(entry(rejected_name).with_props(rejected_props).unwrap())
        .unwrap_err()
        .to_string();
    assert!(error.contains("exceeds 32 MiB"), "{error}");
    assert_eq!(catalog.len(), before);
    assert!(catalog.get(rejected_name).unwrap().is_none());

    let replacement_props = Props {
        style: Some("\"".repeat(MAX_STRING)),
        description: Some("\"".repeat(MAX_STRING)),
        ..Props::new(Name::new("Large 0").unwrap())
    };
    let replacement_error = catalog
        .put(entry("Large 0").with_props(replacement_props).unwrap())
        .unwrap_err()
        .to_string();
    assert!(
        replacement_error.contains("exceeds 32 MiB"),
        "{replacement_error}"
    );
    assert!(
        catalog
            .get("Large 0")
            .unwrap()
            .unwrap()
            .props()
            .unwrap()
            .style
            .is_none()
    );

    let background = format!(
        r#"<w:background xmlns:w="{W}"><w:color>{}</w:color></w:background>"#,
        ">".repeat(600_000)
    );
    let background_error = catalog
        .set_background(background.into_bytes())
        .unwrap_err()
        .to_string();
    assert!(
        background_error.contains("exceeds 32 MiB"),
        "{background_error}"
    );
    assert!(catalog.background().is_none());
}

#[test]
fn canonical_write_matches_its_checked_plan() {
    let mut catalog = Catalog::new();
    catalog.add(entry("Planned")).unwrap();
    catalog
        .set_background(format!(r#"<w:background xmlns:w="{W}" w:color="A&amp;B"/>"#).into_bytes())
        .unwrap();

    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let plan = plan_write(&catalog, conformance).unwrap();
        let xml = write(&catalog, conformance).unwrap();
        assert_eq!(xml.len(), plan.bytes);
        assert_eq!(read(&xml).unwrap().1, conformance);
    }
}

#[test]
fn aggregate_auxiliary_payload_limit_rejects_before_mutation() {
    let mut package = package();
    let payload = Arc::new(vec![0u8; MAX]);
    let mut graph = raw::Graph::new(Catalog::new(), Conformance::Transitional);
    for index in 0..9 {
        graph.parts.push(raw::Part::from_shared(
            format!("/word/glossary/media/data{index}.bin"),
            "application/octet-stream".to_owned(),
            Arc::clone(&payload),
            Vec::new(),
        ));
    }
    assert!(raw::put(&mut package, &graph).is_err());
    assert!(load(&package).unwrap().is_none());
}
