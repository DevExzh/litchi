//! Native Word document-property and web-settings transactions.

use litchi_docx::{Package, web};
use litchi_ooxml_common::custom::Value;
use std::{collections::BTreeMap, path::PathBuf};

fn native(name: &str) -> Package {
    Package::open(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn changed(before: &[u8], after: &[u8]) -> Vec<String> {
    let a = members(before);
    let b = members(after);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}

// These graph patches retain exact part payloads and typed relationship
// identities. Removing/recreating parts can regenerate the manifest and rels
// streams; their complete element/attribute values must remain equivalent.
fn graph_elements(xml: &[u8]) -> Vec<(String, BTreeMap<String, String>)> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut elements = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                let attributes = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
                elements.push((
                    String::from_utf8(e.name().as_ref().to_vec()).unwrap(),
                    attributes,
                ));
            },
            Event::End(_) | Event::Decl(_) => {},
            Event::Text(t) => assert!(t.as_ref().iter().all(u8::is_ascii_whitespace)),
            Event::Eof => break,
            _ => panic!("unexpected opaque graph markup"),
        }
    }
    elements.sort();
    elements
}

fn assert_removal_inverse(before: &[u8], restored: &[u8], relationship_member: &str) {
    let a = members(before);
    let b = members(restored);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name == "[Content_Types].xml" || name == relationship_member {
            assert_eq!(graph_elements(&bytes), graph_elements(&b[&name]), "{name}");
        } else {
            assert!(bytes == b[&name], "changed part {name}");
        }
    }
}

fn fragment<'a>(xml: &'a [u8], tag: &str) -> &'a str {
    let xml = std::str::from_utf8(xml).unwrap();
    let start = xml.find(&format!("<{tag}>")).unwrap();
    let close = format!("</{tag}>");
    let end = start + xml[start..].find(&close).unwrap() + close.len();
    &xml[start..end]
}

#[test]
fn native_review_properties_keep_typed_values_and_exact_noop_members() {
    let package = native("word-review-controls.docx");
    let props = package.metadata().unwrap();
    assert_eq!(props.title.as_deref(), Some("Native Word review fixture"));
    assert_eq!(
        props.subject.as_deref(),
        Some("Tracked revisions, bookmarks, fields, comments")
    );
    assert_eq!(props.author.as_deref(), Some("Litchi Fixture Author"));
    assert_eq!(props.last_modified_by.as_deref(), Some("Ryker Zhu"));
    assert_eq!(
        props.description.as_deref(),
        Some("Synthetic local Microsoft Word fixture created through Computer Use.")
    );
    assert_eq!(props.keywords, ["native", "word", "review", "003"]);
    assert_eq!(props.category.as_deref(), Some("Regression fixture"));
    assert_eq!(props.revision.as_deref(), Some("3"));
    assert_eq!(props.created.unwrap().to_string(), "2026-10-15T05:48:00Z");
    assert_eq!(props.modified.unwrap().to_string(), "2026-10-15T05:53:00Z");
    let app = package.app_props().unwrap().unwrap();
    assert_eq!(app.template.as_deref(), Some("Normal.dotm"));
    assert_eq!(app.application.as_deref(), Some("Microsoft Office Word"));
    assert_eq!(app.app_version.as_deref(), Some("16.0000"));
    assert_eq!(app.manager.as_deref(), Some("Fixture Manager"));
    assert_eq!(app.company.as_deref(), Some("Litchi Native Fixtures"));
    assert_eq!(
        (
            app.pages,
            app.words,
            app.characters,
            app.characters_with_spaces
        ),
        (Some(1), Some(18), Some(239), Some(246))
    );
    assert_eq!(
        (app.lines, app.paragraphs, app.total_time),
        (Some(11), Some(11), Some(5))
    );
    assert_eq!(app.doc_security, Some(0));
    assert_eq!(
        (
            app.scale_crop,
            app.links_up_to_date,
            app.shared_doc,
            app.hyperlinks_changed
        ),
        (Some(false), Some(false), Some(false), Some(false))
    );
    let custom = package.custom_props();
    assert_eq!(custom.len(), 3);
    assert_eq!(
        custom.get("fixtureid"),
        Some(&Value::Text("word-review-controls-003".into()))
    );
    assert_eq!(custom.get("EXPECTEDCOUNT"), Some(&Value::I32(42)));
    assert_eq!(custom.get("IsNative"), Some(&Value::Bool(true)));
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_metadata();
    edit.set(props.clone())
        .set_app_props(Some(app))
        .set_custom_props(custom.clone());
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn native_property_crud_preserves_unmodeled_app_vectors_and_reverses_after_reopen() {
    let package = native("word-review-controls.docx");
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_metadata();
    edit.update(|props| {
        props.title = Some("Revised native review title".into());
        props.keywords.push("edited".into());
    });
    let mut app = edit.app_props().unwrap().unwrap();
    app.company = Some("Edited fixture company".into());
    edit.set_app_props(Some(app));
    let mut custom = edit.custom_props().clone();
    assert_eq!(
        custom.insert("ExpectedCount", Value::I32(43)).unwrap(),
        Some(Value::I32(42))
    );
    assert_eq!(
        custom.remove("FixtureId"),
        Some(Value::Text("word-review-controls-003".into()))
    );
    assert_eq!(
        custom.insert("ReviewedBy", "Native regression").unwrap(),
        None
    );
    edit.set_custom_props(custom);
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    assert_eq!(
        changed(&before, &after),
        [
            "docProps/app.xml",
            "docProps/core.xml",
            "docProps/custom.xml"
        ]
    );
    let a = members(&before);
    let b = members(&after);
    for tag in ["HeadingPairs", "TitlesOfParts"] {
        assert_eq!(
            fragment(&a["docProps/app.xml"], tag),
            fragment(&b["docProps/app.xml"], tag)
        );
    }
    let output = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    updated.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let mut expected_core = package.metadata().unwrap().clone();
    expected_core.title = Some("Revised native review title".into());
    expected_core.keywords.push("edited".into());
    assert_eq!(reopened.metadata(), Some(&expected_core));
    let mut expected_app = package.app_props().unwrap().unwrap();
    expected_app.company = Some("Edited fixture company".into());
    assert_eq!(reopened.app_props().unwrap(), Some(expected_app));
    assert_eq!(
        reopened.metadata().unwrap().title.as_deref(),
        Some("Revised native review title")
    );
    assert_eq!(
        reopened.metadata().unwrap().keywords,
        ["native", "word", "review", "003", "edited"]
    );
    assert_eq!(
        reopened.metadata().unwrap().created,
        package.metadata().unwrap().created
    );
    assert_eq!(
        reopened.app_props().unwrap().unwrap().company.as_deref(),
        Some("Edited fixture company")
    );
    assert_eq!(reopened.custom_props().len(), 3);
    assert_eq!(reopened.custom_props().get("FixtureId"), None);
    assert_eq!(
        reopened.custom_props().get("ExpectedCount"),
        Some(&Value::I32(43))
    );
    assert_eq!(
        reopened.custom_props().get("IsNative"),
        Some(&Value::Bool(true))
    );
    assert_eq!(
        reopened.custom_props().get("ReviewedBy"),
        Some(&Value::Text("Native regression".into()))
    );
    let restored = reopened
        .apply_metadata(&patch.inverse())
        .unwrap()
        .into_package();
    assert!(
        changed(&before, &restored.to_bytes().unwrap()).is_empty(),
        "inverse changed {:?}",
        changed(&before, &restored.to_bytes().unwrap())
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}

#[test]
fn native_property_family_removal_and_inverse_restore_relationships_and_parts() {
    let package = native("word-review-controls.docx");
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_metadata();
    edit.clear().set_app_props(None).remove_custom_props();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    assert_eq!(
        changed(&before, &after),
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "docProps/custom.xml"
        ]
    );
    let reopened = Package::from_bytes(&after).unwrap();
    assert!(reopened.metadata().is_none());
    assert!(reopened.app_props().unwrap().is_none());
    assert!(reopened.custom_props().is_empty());
    for name in [
        "docProps/core.xml",
        "docProps/app.xml",
        "docProps/custom.xml",
    ] {
        assert!(!members(&after).contains_key(name));
    }
    let restored = reopened
        .apply_metadata(&patch.inverse())
        .unwrap()
        .into_package();
    assert_removal_inverse(&before, &restored.to_bytes().unwrap(), "_rels/.rels");
}

#[test]
fn native_word_variants_metadata_edit_keeps_template_and_inert_macro_parts() {
    for (name, template, words) in [
        ("word-text-structure.docx", "Normal.dotm", 59),
        ("word-math-navigation.dotx", "word-math-navigation.dotx", 48),
        ("word-recorded-macro.docm", "Normal.dotm", 12),
        ("word-recorded-macro.dotm", "word-recorded-macro.dotm", 12),
    ] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let app = package.app_props().unwrap().unwrap();
        assert_eq!(app.template.as_deref(), Some(template));
        assert_eq!(app.words, Some(words));
        assert_eq!(
            package.metadata().unwrap().author.as_deref(),
            Some("Ryker Zhu")
        );
        let mut edit = package.edit_metadata();
        edit.update(|props| props.title = Some("Native property variant edit".into()));
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let after = updated.to_bytes().unwrap();
        assert_eq!(changed(&before, &after), ["docProps/core.xml"]);
        let reopened = Package::from_bytes(&after).unwrap();
        assert_eq!(reopened.app_props().unwrap().unwrap(), app);
        assert_eq!(
            reopened.metadata().unwrap().title.as_deref(),
            Some("Native property variant edit")
        );
        let restored = reopened
            .apply_metadata(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(
            changed(&before, &restored.to_bytes().unwrap()).is_empty(),
            "inverse changed {:?}",
            changed(&before, &restored.to_bytes().unwrap())
        );
    }
}

#[test]
fn native_word_web_settings_read_and_screen_edit_preserve_other_source_xml() {
    for name in [
        "word-text-structure.docx",
        "word-math-navigation.dotx",
        "word-recorded-macro.docm",
        "word-recorded-macro.dotm",
    ] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let (settings, conformance) = package.web().unwrap().unwrap();
        assert_eq!(conformance, web::Conformance::Transitional);
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.allow_png(), Some(true));
        assert_eq!(settings.target_screen_size(), None);
        assert_eq!(settings.pixels_per_inch(), None);
        assert_eq!(settings.encoding(), None);
        assert!(settings.frameset().is_none() && settings.divs().is_none());
        let mut noop = package.edit_web();
        assert!(!noop.put(settings, conformance).unwrap());
        assert_eq!(
            noop.commit().unwrap().into_package().to_bytes().unwrap(),
            before
        );
        let mut edit = package.edit_web();
        edit.set_target_screen_size(Some(web::Screen::Pixels1024x768))
            .unwrap();
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let after = updated.to_bytes().unwrap();
        assert_eq!(changed(&before, &after), ["word/webSettings.xml"]);
        let expected = String::from_utf8(members(&before)["word/webSettings.xml"].clone())
            .unwrap()
            .replace(
                "</w:webSettings>",
                "<w:targetScreenSz w:val=\"1024x768\"/></w:webSettings>",
            );
        assert_eq!(members(&after)["word/webSettings.xml"], expected.as_bytes());
        let output = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
        updated.save(output.path()).unwrap();
        let reopened = Package::open(output.path()).unwrap();
        let settings = reopened.web().unwrap().unwrap().0;
        assert_eq!(
            settings.target_screen_size(),
            Some(web::Screen::Pixels1024x768)
        );
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.allow_png(), Some(true));
        let restored = reopened.apply_web(&patch.inverse()).unwrap().into_package();
        assert!(
            changed(&before, &restored.to_bytes().unwrap()).is_empty(),
            "inverse changed {:?}",
            changed(&before, &restored.to_bytes().unwrap())
        );
    }
}

#[test]
fn native_word_web_settings_remove_and_restore_preserves_document_relationships() {
    let package = native("word-review-controls.docx");
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_web();
    assert!(edit.remove().unwrap());
    assert!(!edit.remove().unwrap());
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    assert_eq!(
        changed(&before, &after),
        [
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/webSettings.xml"
        ]
    );
    let reopened = Package::from_bytes(&after).unwrap();
    assert!(reopened.web().unwrap().is_none());
    let restored = reopened.apply_web(&patch.inverse()).unwrap().into_package();
    assert_removal_inverse(
        &before,
        &restored.to_bytes().unwrap(),
        "word/_rels/document.xml.rels",
    );
}
