//! Native Word parents with explicit inert smart-tag metadata.
use litchi_core::{OwnedSource, Position};
use litchi_docx::{Builder, Error, Package, source_backed};
use litchi_opc::{PackURI, SourceLimits, phys_pkg::OwnedPhysPkgReader};
use std::{collections::BTreeMap, path::Path, sync::Arc};
const MAIN: &str = "word/document.xml";
const SETTINGS: &str = "word/settings.xml";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PARENTS: [&str; 3] = [
    "word-text-structure.docx",
    "word-recorded-macro.docm",
    "word-math-navigation.dotx",
];
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = OwnedPhysPkgReader::from_bytes(bytes.to_vec(), SourceLimits::default()).unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = r.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}
fn rewrite(p: &Package, part: &str, xml: Vec<u8>) -> Package {
    let mut e = p.edit_opc();
    e.package_mut()
        .get_part_mut(&PackURI::new(format!("/{part}")).unwrap())
        .unwrap()
        .set_blob(xml);
    e.commit().unwrap().into_package()
}
fn append(p: &Package, fragment: &str) -> Package {
    let xml = String::from_utf8(members(&p.to_bytes().unwrap())[MAIN].clone()).unwrap();
    let at = xml.rfind("<w:sectPr").unwrap();
    let mut out = xml.as_bytes()[..at].to_vec();
    out.extend_from_slice(fragment.as_bytes());
    out.extend_from_slice(&xml.as_bytes()[at..]);
    rewrite(p, MAIN, out)
}
fn only_changed(a: &[u8], b: &[u8], part: &str) {
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name != part {
            assert!(b.get(&name) == Some(&bytes), "changed {name}");
        }
    }
}
fn built_paragraph() -> String {
    let mut builder = Builder::new().unwrap();
    let tag = builder
        .document()
        .unwrap()
        .add_paragraph()
        .add_smart_tag("person");
    tag.set_uri("urn:native:contacts");
    tag.add_attribute("kind", "friend & peer")
        .set_uri("urn:native:meta");
    tag.add_run_with_text("SMART_TAG_076 日本語 & ");
    let nested = tag.add_smart_tag("givenName");
    nested.add_attribute("locale", "ja");
    nested.add_run_with_text("name 🍎");
    tag.add_smart_tag("empty");
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let xml = String::from_utf8(members(&bytes)[MAIN].clone()).unwrap();
    let start = xml.find("<w:p>").unwrap();
    let end = xml.find("</w:p>").unwrap() + "</w:p>".len();
    xml[start..end].to_owned()
}
fn tags(p: &Package) -> Vec<litchi_docx::smart_tag::SmartTag> {
    let doc = p.document().unwrap();
    let paragraphs = doc.paragraphs().unwrap();
    let paragraph = paragraphs
        .iter()
        .find(|p| p.text().unwrap().contains("SMART_TAG_076"))
        .unwrap();
    paragraph.smart_tags().unwrap()
}
fn assert_tags(p: &Package) {
    let tags = tags(p);
    assert_eq!(tags.len(), 3);
    assert_eq!(
        (tags[0].uri.as_deref(), tags[0].element.as_str()),
        (Some("urn:native:contacts"), "person")
    );
    assert_eq!(
        tags[0].attributes,
        [litchi_docx::smart_tag::SmartTagAttribute {
            uri: Some("urn:native:meta".into()),
            name: "kind".into(),
            value: "friend & peer".into()
        }]
    );
    assert_eq!(tags[0].text().unwrap(), "SMART_TAG_076 日本語 & name 🍎");
    assert_eq!(tags[1].element, "givenName");
    assert_eq!(tags[1].attributes[0].name, "locale");
    assert_eq!(tags[1].attributes[0].value, "ja");
    assert_eq!(tags[1].text().unwrap(), "name 🍎");
    assert_eq!(tags[2].element, "empty");
    assert!(tags[2].text().unwrap().is_empty());
}

#[test]
fn native_docx_builder_smart_tags_and_settings_read_exactly_without_resolving_urls() {
    for file in PARENTS {
        let bytes = native(file);
        let base = Package::from_bytes(&bytes).unwrap();
        let p = append(&base, &built_paragraph());
        assert_tags(&p);
        only_changed(&bytes, &p.to_bytes().unwrap(), MAIN);
        let xml = String::from_utf8(members(&p.to_bytes().unwrap())[SETTINGS].clone()).unwrap();
        assert!(!xml.contains("<w:smartTagType") && !xml.contains("<w:doNotEmbedSmartTags"));
        let fragment = r#"<w:smartTagType w:namespaceuri="urn:native:contacts" w:name="person" w:url="https://example.invalid/never-fetched?a=1&amp;b=2"/><w:doNotEmbedSmartTags w:val="0"/>"#;
        let at = xml
            .find("<w:decimalSymbol")
            .or_else(|| xml.find("<w:listSeparator"))
            .unwrap_or_else(|| xml.rfind("</w:settings>").unwrap());
        let mut modified = xml[..at].to_owned();
        modified.push_str(fragment);
        modified.push_str(&xml[at..]);
        let q = rewrite(&p, SETTINGS, modified.into_bytes());
        only_changed(&p.to_bytes().unwrap(), &q.to_bytes().unwrap(), SETTINGS);
        let settings = q.document().unwrap().settings().unwrap().unwrap();
        assert!(!settings.do_not_embed_smart_tags());
        assert_eq!(settings.smart_tag_types().len(), 1);
        let t = &settings.smart_tag_types()[0];
        assert_eq!(
            (t.namespace_uri(), t.name(), t.url()),
            (
                "urn:native:contacts",
                "person",
                "https://example.invalid/never-fetched?a=1&b=2"
            )
        );
        let file_out = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
        q.save(file_out.path()).unwrap();
        let reopened = Package::open(file_out.path()).unwrap();
        assert_tags(&reopened);
        assert_eq!(
            members(&reopened.to_bytes().unwrap()),
            members(&q.to_bytes().unwrap())
        );
        assert_eq!(base.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_docx_smart_tag_namespace_declarations_and_aliases_do_not_change_metadata() {
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let xml=built_paragraph().replace("<w:smartTag ",r#"<w:smartTag xmlns:element="urn:namespace-only" xmlns:uri="urn:another-namespace" "#).replace("<w:attr ",r#"<w:attr xmlns:name="urn:namespace-name" xmlns:val="urn:namespace-value" "#);
        for fragment in [
            xml.clone(),
            xml.replace("w:", "q:")
                .replacen("<q:p>", &format!(r#"<q:p xmlns:q="{W}">"#), 1),
        ] {
            let q = append(&p, &fragment);
            assert_tags(&q);
            let reread = Package::from_bytes(&q.to_bytes().unwrap()).unwrap();
            assert_tags(&reread);
            only_changed(&bytes, &q.to_bytes().unwrap(), MAIN);
        }
        let fragment = built_paragraph()
            .replace("w:smartTag", "q:smartTag")
            .replace("w:attr", "q:attr");
        let xml = String::from_utf8(members(&p.to_bytes().unwrap())[MAIN].clone()).unwrap();
        assert!(!xml.contains("xmlns:q="));
        let inherited = rewrite(
            &p,
            MAIN,
            xml.replacen("<w:document ", &format!(r#"<w:document xmlns:q="{W}" "#), 1)
                .into_bytes(),
        );
        let inherited = append(&inherited, &fragment);
        assert_tags(&inherited);
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_docx_smart_tag_foreign_attributes_do_not_satisfy_required_word_metadata() {
    let bytes = native(PARENTS[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    for (fragment, expected) in [
        (
            format!(
                r#"<w:p xmlns:w="{W}"><w:smartTag xmlns:f="urn:foreign" f:element="person"><w:r><w:t>SMART_TAG_076</w:t></w:r></w:smartTag></w:p>"#
            ),
            "Word smart tag is missing its element attribute",
        ),
        (
            format!(
                r#"<w:p xmlns:w="{W}"><w:smartTag w:element="person"><w:smartTagPr><w:attr xmlns:f="urn:foreign" w:name="kind" f:val="bad"/></w:smartTagPr><w:r><w:t>SMART_TAG_076</w:t></w:r></w:smartTag></w:p>"#
            ),
            "Word smart-tag property is missing its value",
        ),
    ] {
        let q = append(&p, &fragment);
        let doc = q.document().unwrap();
        let paragraphs = doc.paragraphs().unwrap();
        let para = paragraphs
            .iter()
            .find(|p| p.text().unwrap().contains("SMART_TAG_076"))
            .unwrap();
        assert!(matches!(para.smart_tags(),Err(Error::InvalidFormat(m)) if m==expected));
    }
    assert_eq!(p.to_bytes().unwrap(), bytes);
    assert_eq!(native(PARENTS[0]), bytes);
}

#[test]
fn native_docx_story_text_edits_preserve_smart_tag_properties_and_nesting() {
    use source_backed::story_text::Selector;
    for file in PARENTS {
        let bytes = native(file);
        let base = Package::from_bytes(&bytes).unwrap();
        let p = append(&base, &built_paragraph());
        assert_tags(&p);
        let input = p.to_bytes().unwrap();
        let source =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(input.clone())))
                .unwrap();
        let snapshot = source.story_text_snapshot(Selector::Main).unwrap();
        let last = snapshot.paragraph_count().unwrap() - 1;
        assert_eq!(
            snapshot.paragraph_text(last).unwrap().as_deref(),
            Some("SMART_TAG_076 日本語 & name 🍎")
        );
        let mut e = source.edit_story_text(Selector::Main).unwrap();
        e.replace_paragraph_text(Position::new(0), "Reviewed native title")
            .unwrap();
        assert!(matches!(
            e.replace_paragraph_text(Position::new(last), "SMART_TAG_076 edited 日本語 🍎"),
            Err(source_backed::story_text::Error::DocumentEdit(litchi_docx::document::Error::Refused { position, reason: litchi_docx::document::Refusal::ComplexContent })) if position == last
        ));
        let c = e.commit().unwrap();
        let mut output = Vec::new();
        c.package().write_to(&mut output).unwrap();
        let q = Package::from_bytes(&output).unwrap();
        only_changed(&input, &output, MAIN);
        let old = tags(&p);
        let new = tags(&q);
        assert_eq!(new.len(), old.len());
        for (a, b) in old.iter().zip(&new) {
            assert_eq!(
                (&a.uri, &a.element, &a.attributes),
                (&b.uri, &b.element, &b.attributes)
            );
        }
        assert_eq!(new, old);
        assert_eq!(new[0].text().unwrap(), "SMART_TAG_076 日本語 & name 🍎");
        assert_eq!(
            q.document().unwrap().paragraphs().unwrap()[0]
                .text()
                .unwrap(),
            "Reviewed native title"
        );
        let target =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
        let restored = target.apply_story_text(&c.patch().inverse()).unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(members(&restored_bytes), members(&input));
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_template_footnote_and_macro_variant_stories_support_source_text_edits() {
    use source_backed::story_text::Selector;
    for (file, selector, part) in [
        ("word-recorded-macro.docm", Selector::Main, MAIN),
        ("word-recorded-macro.dotm", Selector::Main, MAIN),
        ("word-math-navigation.dotx", Selector::Main, MAIN),
        (
            "word-layout-extensions.dotx",
            Selector::Footnote(1),
            "word/footnotes.xml",
        ),
    ] {
        let bytes = native(file);
        let source =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
                .unwrap();
        let before = source.story_text_snapshot(selector.clone()).unwrap();
        assert!(
            before
                .paragraph_text(0)
                .unwrap()
                .is_some_and(|t| !t.is_empty())
        );
        let mut edit = source.edit_story_text(selector.clone()).unwrap();
        edit.replace_paragraph_text(Position::new(0), "Variant story edited 日本語")
            .unwrap();
        let c = edit.commit().unwrap();
        let mut after = Vec::new();
        c.package().write_to(&mut after).unwrap();
        only_changed(&bytes, &after, part);
        let p = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(after))).unwrap();
        assert_eq!(
            p.story_text_snapshot(selector)
                .unwrap()
                .paragraph_text(0)
                .unwrap()
                .as_deref(),
            Some("Variant story edited 日本語")
        );
        let inverse = p.apply_story_text(&c.patch().inverse()).unwrap();
        let mut restored = Vec::new();
        inverse.package().write_to(&mut restored).unwrap();
        assert_eq!(members(&restored), members(&bytes));
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_header_story_remains_editable_in_an_explicit_macro_template_variant() {
    use source_backed::story_text::Selector;
    let bytes = native("word-tables-layout.docx");
    let base = Package::from_bytes(&bytes).unwrap();
    let mut e = base.edit_opc();
    e.package_mut()
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_content_type(litchi_opc::constants::content_type::WML_TEMPLATE_MACRO_MAIN.into())
        .unwrap();
    let variant = e.commit().unwrap().into_package();
    let input = variant.to_bytes().unwrap();
    only_changed(&bytes, &input, "[Content_Types].xml");
    let source =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(input.clone()))).unwrap();
    assert_eq!(
        source
            .story_text_snapshot(Selector::Header(0))
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("HEADER_DEFAULT_002")
    );
    let mut e = source.edit_story_text(Selector::Header(0)).unwrap();
    e.replace_paragraph_text(Position::new(0), "HEADER_VARIANT_076")
        .unwrap();
    let c = e.commit().unwrap();
    let mut after = Vec::new();
    c.package().write_to(&mut after).unwrap();
    only_changed(&input, &after, "word/header1.xml");
    let reopened = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(after))).unwrap();
    assert_eq!(
        reopened
            .story_text_snapshot(Selector::Header(0))
            .unwrap()
            .paragraph_text(0)
            .unwrap()
            .as_deref(),
        Some("HEADER_VARIANT_076")
    );
    let inverse = reopened.apply_story_text(&c.patch().inverse()).unwrap();
    let mut restored = Vec::new();
    inverse.package().write_to(&mut restored).unwrap();
    assert_eq!(members(&restored), members(&input));
    assert_eq!(base.to_bytes().unwrap(), bytes);
    assert_eq!(native("word-tables-layout.docx"), bytes);
}
