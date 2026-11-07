//! Native numbering, bibliography and Custom XML owners; no numbering or field evaluation.
use litchi_docx::bibliography::{
    BibliographyPerson as Person, BibliographySourceBuilder as Source,
    BibliographySourceKind as Kind, OOXML_BIBLIOGRAPHY_NAMESPACE as BIB,
};
use litchi_docx::{
    Error, Package, custom_xml,
    numbering::{Format, MultiLevel, Restart, Suffix},
};
use std::{collections::BTreeMap, io::Cursor};
const FILES: [&str; 3] = [
    "word-text-structure.docx",
    "word-math-navigation.docx",
    "word-math-navigation.dotx",
];
const STORE_IDS: [&str; 2] = [
    "{8F890C7E-E3B9-FA4F-9FD3-B127A8AD5D7E}",
    "{72EE7E13-CEB2-0148-AD62-BE2A8CF06DAC}",
];
const ADDED_ID: &str = "{44FE7082-5517-4A11-9258-B07286C2DE88}";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}
fn graph_elements(xml: &[u8]) -> Vec<(String, BTreeMap<String, String>)> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut out = Vec::new();
    loop {
        match r.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                let attrs = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                r.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
                out.push((
                    String::from_utf8(e.name().as_ref().to_vec()).unwrap(),
                    attrs,
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::End(_) => {},
            Event::Text(t) => assert!(t.as_ref().iter().all(u8::is_ascii_whitespace)),
            _ => panic!("unexpected graph markup"),
        }
    }
    out.sort();
    out
}
fn content_types(parts: &BTreeMap<String, Vec<u8>>) -> BTreeMap<String, String> {
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    for (name, attrs) in graph_elements(&parts["[Content_Types].xml"]) {
        match name.as_str() {
            "Default" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    defaults
                        .insert(attrs["Extension"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Override" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    overrides
                        .insert(attrs["PartName"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Types" => assert_eq!(
                attrs,
                BTreeMap::from([(
                    "xmlns".into(),
                    "http://schemas.openxmlformats.org/package/2006/content-types".into()
                )])
            ),
            _ => panic!("unexpected content-type element: {name}"),
        }
    }
    parts
        .keys()
        .filter(|n| *n != "[Content_Types].xml")
        .map(|n| {
            let kind = overrides
                .get(&format!("/{n}"))
                .or_else(|| defaults.get(n.rsplit('.').next().unwrap()))
                .unwrap();
            (n.clone(), kind.clone())
        })
        .collect()
}
fn assert_same_payloads_and_graph(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    // A default and per-part overrides can describe the same OPC typing.
    // Check every member's resolved type, not that serialization choice.
    assert_eq!(content_types(&a), content_types(&b));
    for (n, data) in a {
        if n == "[Content_Types].xml" {
            continue;
        }
        if n.ends_with(".rels") {
            assert_eq!(graph_elements(&data), graph_elements(&b[&n]), "{n}");
        } else {
            assert_eq!(data, b[&n], "{n}");
        }
    }
}

fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn book(title: &str) -> Source {
    Source::new(Kind::Book, "NativeBook")
        .unwrap()
        .title(title)
        .unwrap()
        .year("2026")
        .unwrap()
        .lcid("1033")
        .unwrap()
        .publisher("Fixture Press")
        .unwrap()
        .city("Test City")
        .unwrap()
        .person(Person::new("Rivera").unwrap().with_first("Ana").unwrap())
        .unwrap()
        .custom_field(vec!["BookTitle".into()], "A & B <reference> 日本語".into())
        .unwrap()
}
fn store(xml: &[u8], schemas: &[&str]) -> custom_xml::NewStore {
    custom_xml::NewStore {
        xml: xml.to_vec(),
        content_type: "application/xml".into(),
        id: ADDED_ID.into(),
        schemas: schemas.iter().map(|s| s.to_string()).collect(),
        conformance: custom_xml::Conformance::Transitional,
    }
}
fn root_metadata(xml: &[u8]) -> BTreeMap<String, String> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    loop {
        match r.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                return e
                    .attributes()
                    // The canonical bibliography writer may omit redundant
                    // namespace declarations; preserve all stored metadata.
                    .filter(|a| {
                        let a = a.as_ref().unwrap();
                        a.key.as_ref() != b"xmlns" && !a.key.as_ref().starts_with(b"xmlns:")
                    })
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                r.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
            },
            Event::Eof => panic!("missing root"),
            _ => {},
        }
    }
}

#[test]
fn native_numbering_definitions_resolve_exact_levels_paragraphs_and_restart_state() {
    for (index, name) in FILES.into_iter().enumerate() {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let snap = p.numbering_snapshot().unwrap().unwrap();
        assert_eq!(snap.xml_bytes(), members(&bytes)["word/numbering.xml"]);
        let c = snap.collection();
        assert_eq!(c.abstract_num_count(), 1);
        assert_eq!(c.num_count(), 1);
        assert!(c.picture_bullets().is_empty());
        let def = c.get_abstract_num(0).unwrap();
        assert_eq!(
            def.num_type(),
            Some(if index == 0 {
                MultiLevel::Hybrid
            } else {
                MultiLevel::Multi
            })
        );
        assert_eq!(def.restart_numbering_after_break(), Some(false));
        assert_eq!(snap.restart_numbering_after_break(0).unwrap(), Some(false));
        assert_eq!(def.levels().len(), 9);
        for (i, level) in def.levels().iter().enumerate() {
            assert_eq!(level.level, i as u8);
            assert_eq!(level.start, 1);
            assert_eq!(level.suffix, Suffix::Tab);
            assert_eq!(level.restart, Restart::Default);
            assert!(!level.legal);
            assert_eq!(
                level.format,
                if index == 0 {
                    Format::Bullet
                } else {
                    Format::Decimal
                }
            );
            let text = if index == 0 {
                ["\u{f0b7}", "o", "\u{f0a7}"][i % 3].to_owned()
            } else {
                (1..=i + 1).map(|n| format!("%{n}.")).collect::<String>()
            };
            assert_eq!(level.level_text.as_deref(), Some(text.as_str()));
            assert_eq!(level.picture_bullet_id, None);
        }
        let instance = c.get_num(1).unwrap();
        assert_eq!(instance.abstract_num_id(), 0);
        assert!(instance.overrides().is_empty());
        let d = p.document().unwrap();
        let paras = d.paragraphs().unwrap();
        let numbered = paras
            .iter()
            .filter_map(|q| {
                q.numbering()
                    .unwrap()
                    .map(|n| (q.text().unwrap(), n.num_id, n.level))
            })
            .collect::<Vec<_>>();
        let expected = if index == 0 {
            vec![
                ("List item one", 0),
                ("List item two", 0),
                ("List item three", 0),
            ]
        } else {
            vec![
                ("NUMBERED_FIRST_005", 0),
                ("NUMBERED_NESTED_005", 1),
                ("NUMBERED_SECOND_005", 0),
            ]
        };
        assert_eq!(
            numbered,
            expected
                .into_iter()
                .map(|(s, n)| (s.to_owned(), 1, n))
                .collect::<Vec<_>>()
        );
        let mut e = p.edit_numbering();
        e.set_restart_numbering_after_break(0, Some(false)).unwrap();
        assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_restart_attribute_crud_preserves_list_definitions_and_failed_edits() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let old = members(&bytes);
        let mut e = p.edit_numbering();
        e.set_restart_numbering_after_break(0, Some(true)).unwrap();
        assert!(matches!(
            e.set_restart_numbering_after_break(99, Some(false)),
            Err(Error::InvalidFormat(message))
                if message == "abstract numbering definition 99 does not exist"
        ));
        assert_eq!(e.restart_numbering_after_break(0).unwrap(), Some(true));
        let c = e.commit().unwrap();
        let after = save(c.package());
        let output = after.to_bytes().unwrap();
        assert_eq!(
            after
                .numbering_snapshot()
                .unwrap()
                .unwrap()
                .restart_numbering_after_break(0)
                .unwrap(),
            Some(true)
        );
        assert_eq!(changed(&bytes, &output), ["word/numbering.xml"]);
        let expected = std::str::from_utf8(&old["word/numbering.xml"])
            .unwrap()
            .replace(
                "restartNumberingAfterBreak=\"0\"",
                "restartNumberingAfterBreak=\"true\"",
            );
        assert_eq!(members(&output)["word/numbering.xml"], expected.as_bytes());
        assert!(
            members(
                &p.apply_numbering(c.patch())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap()
            ) == members(&output)
        );
        assert!(matches!(
            after.apply_numbering(c.patch()),
            Err(Error::PatchConflict { .. })
        ));
        assert!(
            members(
                &after
                    .apply_numbering(&c.patch().inverse())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap()
            ) == old
        );
        let mut e = after.edit_numbering();
        e.clear_restart_numbering_after_break(0).unwrap();
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(
            absent
                .numbering_snapshot()
                .unwrap()
                .unwrap()
                .restart_numbering_after_break(0)
                .unwrap(),
            None
        );
        let mut e = absent.edit_numbering();
        e.set_restart_numbering_after_break(0, Some(false)).unwrap();
        let recreated = e.commit().unwrap();
        let off = save(recreated.package());
        assert_eq!(
            off.numbering_snapshot()
                .unwrap()
                .unwrap()
                .restart_numbering_after_break(0)
                .unwrap(),
            Some(false)
        );
        assert_same_payloads_and_graph(
            &absent.to_bytes().unwrap(),
            &off.apply_numbering(&recreated.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let mut reset = p.edit_numbering();
        reset.clear_restart_numbering_after_break(0).unwrap();
        reset.reset();
        assert_eq!(reset.commit().unwrap().package().to_bytes().unwrap(), bytes);
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_bibliography_stores_retain_exact_apa_metadata_and_distinct_item_ids() {
    for (name, id) in FILES[1..].iter().zip(STORE_IDS) {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let parts = members(&bytes);
        let stores = p.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        let s = &stores[0];
        assert_eq!(s.data_store_item_id(), Some(id));
        assert_eq!(s.source_part_name().as_str(), "/word/document.xml");
        assert_eq!(s.data_part_name().as_str(), "/customXml/item1.xml");
        assert_eq!(
            s.properties_part_name().unwrap().as_str(),
            "/customXml/itemProps1.xml"
        );
        assert_eq!(s.content_type(), "application/xml");
        assert_eq!(s.schema_references(), [BIB]);
        assert_eq!(s.selected_style(), Some("/APASixthEditionOfficeOnline.xsl"));
        assert_eq!(s.style_name(), Some("APA"));
        assert_eq!(s.source_count(), 0);
        assert_eq!(root_metadata(&parts["customXml/item1.xml"])["Version"], "6");
        let items = p.custom_xml().unwrap();
        assert_eq!(items.len(), 1);
        let item = &items[0];
        assert_eq!(item.props().unwrap().id, id);
        assert_eq!(item.props().unwrap().schemas, [BIB]);
        assert_eq!(item.xml(), parts["customXml/item1.xml"]);
        assert_eq!(item.props_xml().unwrap(), parts["customXml/itemProps1.xml"]);
        assert_eq!(item.rel_id(), s.relationship_id());
        assert!(!item.source_relationship().is_external());
        assert_eq!(item.relationships().len(), 1);
        assert_eq!(item.relationships()[0].id, "rId1");
        assert_eq!(item.relationships()[0].target, "itemProps1.xml");
        assert_eq!(
            p.edit_bibliography()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
        let mut e = p.edit_custom_xml();
        e.set(&id.to_ascii_lowercase(), item.xml().to_vec())
            .unwrap();
        assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_bibliography_source_crud_preserves_style_owner_and_unrelated_parts() {
    for (name, id) in FILES[1..].iter().zip(STORE_IDS) {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let old = members(&bytes);
        let mut e = p.edit_bibliography();
        assert_eq!(e.add_source(book("Native reference")).unwrap(), id);
        assert!(matches!(
            e.add_source(book("Duplicate")),
            Err(Error::Invalid(message))
                if message == "bibliography tag 'NativeBook' already exists in the source store"
        ));
        let created = e.commit().unwrap();
        let after = save(created.package());
        let output = after.to_bytes().unwrap();
        assert_eq!(changed(&bytes, &output), ["customXml/item1.xml"]);
        let sources = after.bibliography_sources().unwrap();
        assert_eq!(sources.len(), 1);
        let b = &sources[0];
        assert_eq!(b.tag(), Some("NativeBook"));
        assert_eq!(b.source_type(), Some("Book"));
        assert_eq!(b.title(), Some("Native reference"));
        assert_eq!(b.year(), Some("2026"));
        assert_eq!(b.lcid(), Some("1033"));
        assert_eq!(
            b.value(&["Author", "Author", "NameList", "Person", "Last"]),
            Some("Rivera")
        );
        assert_eq!(
            b.value(&["Author", "Author", "NameList", "Person", "First"]),
            Some("Ana")
        );
        assert_eq!(b.value(&["BookTitle"]), Some("A & B <reference> 日本語"));
        assert_eq!(
            root_metadata(&members(&output)["customXml/item1.xml"]),
            root_metadata(&old["customXml/item1.xml"])
        );
        assert!(
            members(
                &p.apply_bibliography(created.patch())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap()
            ) == members(&output)
        );
        assert!(matches!(
            after.apply_bibliography(created.patch()),
            Err(Error::PatchConflict { .. })
        ));
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_bibliography(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let mut e = after.edit_bibliography();
        e.replace_source("NativeBook", book("Revised reference"))
            .unwrap();
        e.add_source(
            Source::new(Kind::Report, "NativeReport")
                .unwrap()
                .title("Second reference")
                .unwrap()
                .corporate_author("Fixture Group")
                .unwrap(),
        )
        .unwrap();
        let updated = e.commit().unwrap();
        let updated = save(updated.package());
        let sources = updated.bibliography_sources().unwrap();
        assert_eq!(sources.len(), 2);
        assert_eq!(sources[0].title(), Some("Revised reference"));
        assert_eq!(
            sources[1].value(&["Author", "Author", "Corporate"]),
            Some("Fixture Group")
        );
        let mut e = updated.edit_bibliography();
        assert!(!e.remove_source("Missing").unwrap());
        assert!(e.remove_source("NativeBook").unwrap());
        assert!(e.remove_source("NativeReport").unwrap());
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(absent.bibliography_source_count().unwrap(), 0);
        assert_eq!(absent.custom_xml().unwrap().len(), 1);
        assert_eq!(
            root_metadata(absent.custom_xml().unwrap()[0].xml()),
            root_metadata(&old["customXml/item1.xml"])
        );
        assert_same_payloads_and_graph(
            &updated.to_bytes().unwrap(),
            &absent
                .apply_bibliography(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_custom_xml_crud_keeps_the_bibliography_owner_and_staged_payloads() {
    const FIRST: &[u8] = br#"<n:check xmlns:n="urn:litchi:native" value="first"/>"#;
    const SECOND: &[u8] = br#"<n:check xmlns:n="urn:litchi:native" value="A &amp; B"/>"#;
    for name in &FILES[1..] {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let original = p.custom_xml().unwrap().remove(0);
        let mut e = p.edit_custom_xml();
        let added = e.add(store(FIRST, &["urn:litchi:native"])).unwrap();
        assert_eq!(added.part().as_str(), "/customXml/item2.xml");
        assert!(matches!(
            e.add(store(SECOND, &[])),
            Err(Error::Refused(reason))
                if reason.kind() == litchi_core::edit::RefusalKind::UnsupportedTarget
                    && reason.reason() == "the Custom XML item ID is already occupied"
        ));
        e.set(ADDED_ID, SECOND.to_vec()).unwrap();
        let created = e.commit().unwrap();
        let after = save(created.package());
        let output = after.to_bytes().unwrap();
        let items = after.custom_xml().unwrap();
        assert_eq!(items.len(), 2);
        assert_eq!(items[0].props(), original.props());
        assert_eq!(items[0].xml(), original.xml());
        let added = items
            .iter()
            .find(|i| i.props().unwrap().id == ADDED_ID)
            .unwrap();
        assert_eq!(added.xml(), SECOND);
        assert_eq!(added.props().unwrap().schemas, ["urn:litchi:native"]);
        assert_eq!(
            changed(&bytes, &output),
            [
                "[Content_Types].xml",
                "customXml/_rels/item1.xml.rels",
                "customXml/_rels/item2.xml.rels",
                "customXml/item2.xml",
                "customXml/itemProps2.xml",
                "word/_rels/document.xml.rels"
            ]
        );
        assert_eq!(
            graph_elements(&members(&bytes)["customXml/_rels/item1.xml.rels"]),
            graph_elements(&members(&output)["customXml/_rels/item1.xml.rels"]),
        );
        assert_same_payloads_and_graph(
            &output,
            &p.apply_custom_xml(created.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert!(matches!(
            after.apply_custom_xml(created.patch()),
            Err(Error::PatchConflict { .. })
        ));
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_custom_xml(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let mut e = after.edit_custom_xml();
        e.replace(ADDED_ID, store(FIRST, &["urn:litchi:updated"]))
            .unwrap();
        let updated = e.commit().unwrap();
        let updated = save(updated.package());
        let items = updated.custom_xml().unwrap();
        let added = items
            .iter()
            .find(|i| i.props().unwrap().id == ADDED_ID)
            .unwrap();
        assert_eq!(added.xml(), FIRST);
        assert_eq!(added.props().unwrap().schemas, ["urn:litchi:updated"]);
        assert_eq!(
            changed(&output, &updated.to_bytes().unwrap()),
            ["customXml/item2.xml", "customXml/itemProps2.xml"]
        );
        let mut e = updated.edit_custom_xml();
        assert!(e.remove(ADDED_ID).unwrap());
        assert!(!e.remove(ADDED_ID).unwrap());
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(absent.custom_xml().unwrap().len(), 1);
        assert_same_payloads_and_graph(&bytes, &absent.to_bytes().unwrap());
        assert_same_payloads_and_graph(
            &updated.to_bytes().unwrap(),
            &absent
                .apply_custom_xml(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_bibliography_store_removal_collects_only_its_unbound_graph() {
    for (name, id) in FILES[1..].iter().zip(STORE_IDS) {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        let mut e = p.edit_custom_xml();
        assert!(e.remove(&id.to_ascii_lowercase()).unwrap());
        let removed = e.commit().unwrap();
        let after = save(removed.package());
        let output = after.to_bytes().unwrap();
        assert!(after.custom_xml().unwrap().is_empty());
        assert!(after.bibliography_source_stores().unwrap().is_empty());
        assert_eq!(
            changed(&bytes, &output),
            [
                "[Content_Types].xml",
                "customXml/_rels/item1.xml.rels",
                "customXml/item1.xml",
                "customXml/itemProps1.xml",
                "word/_rels/document.xml.rels"
            ]
        );
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_custom_xml(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}
