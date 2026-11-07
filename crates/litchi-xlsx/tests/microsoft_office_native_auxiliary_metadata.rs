//! Inert, explicitly authored metadata on unchanged native workbook parents.
use litchi_opc::{PackURI, SourceLimits, phys_pkg::OwnedPhysPkgReader};
use litchi_xlsx::{
    Error, Package,
    custom_data::{self, CustomData, ExtensionList, RemovalDisposition},
    phonetic_properties::{PhoneticAlignment as Alignment, PhoneticType as Type},
    smart_tags::{self, Cell, Collection, Property, Tag},
};
use std::{collections::BTreeMap, path::Path};

const PARENTS: [&str; 3] = [
    "excel-metadata.xlsx",
    "excel-recorded-macro.xlsm",
    "excel-visual-analytics-strict.xlsx",
];
const SHEET: &str = "xl/worksheets/sheet1.xml";
const CONNECTIONS: &str = "xl/connections.xml";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STORAGE: &str = "Native 🍎_x0041_";

fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}

fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = OwnedPhysPkgReader::from_bytes(p.to_bytes().unwrap(), SourceLimits::default()).unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = r.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

type Graph = BTreeMap<String, (String, Vec<(String, String, String, bool)>)>;
fn graph(p: &Package) -> Graph {
    let p = p.clone().into_plain_opc();
    let rels = |rels: &litchi_opc::Relationships| {
        let mut items = rels
            .iter()
            .map(|r| {
                (
                    r.r_id().into(),
                    r.reltype().into(),
                    r.target_ref().into(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        items.sort();
        items
    };
    let mut out = BTreeMap::from([("/".into(), (String::new(), rels(p.rels())))]);
    for part in p.iter_parts() {
        out.insert(
            part.partname().to_string(),
            (part.content_type().into(), rels(part.rels())),
        );
    }
    out
}

fn equivalent(a: &Package, b: &Package) {
    assert_eq!(graph(a), graph(b));
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name != "[Content_Types].xml" && name != "xl/_rels/workbook.xml.rels" {
            assert!(b.get(&name) == Some(&bytes), "changed {name}");
        }
    }
}

fn only_changed(a: &Package, b: &Package, owned: &[&str]) {
    assert_eq!(graph(a), graph(b));
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if !owned.contains(&name.as_str()) {
            assert!(
                b.get(&name) == Some(&bytes),
                "unowned member changed: {name}"
            );
        }
    }
}

fn workbook_without_recalculation_flags(xml: &[u8]) -> Vec<Vec<u8>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut out = Vec::new();
    loop {
        let event = reader.read_event().unwrap();
        match &event {
            Event::Eof => break,
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Empty(element) if element.local_name().as_ref() == b"calcPr" => {
                let mut retained = BTreeMap::new();
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if !matches!(
                        attribute.key.as_ref(),
                        b"calcId"
                            | b"fullCalcOnLoad"
                            | b"forceFullCalc"
                            | b"calcCompleted"
                            | b"calcOnSave"
                    ) {
                        retained.insert(attribute.key.as_ref().to_vec(), attribute.value.to_vec());
                    }
                }
                if !retained.is_empty() {
                    out.push(format!("calcPr:{retained:?}").into_bytes());
                }
            },
            _ => out.push(event.as_ref().to_vec()),
        }
    }
    out
}

fn cell_edit_preserved(a: &Package, b: &Package) {
    let mut expected_graph = graph(a);
    expected_graph.remove("/xl/calcChain.xml");
    expected_graph
        .get_mut("/xl/workbook.xml")
        .unwrap()
        .1
        .retain(|r| !r.1.ends_with("/calcChain"));
    assert_eq!(graph(b), expected_graph);
    let a = members(a);
    let b_members = members(b);
    assert_eq!(
        b_members.keys().collect::<Vec<_>>(),
        a.keys()
            .filter(|n| n.as_str() != "xl/calcChain.xml")
            .collect::<Vec<_>>()
    );
    for (name, bytes) in &a {
        if ![
            SHEET,
            "xl/calcChain.xml",
            "xl/workbook.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]
        .contains(&name.as_str())
        {
            assert!(
                b_members.get(name) == Some(bytes),
                "cell edit changed {name}"
            );
        }
    }
    assert_eq!(
        workbook_without_recalculation_flags(&a["xl/workbook.xml"]),
        workbook_without_recalculation_flags(&b_members["xl/workbook.xml"])
    );
    let metadata = b.workbook().calculation_metadata().unwrap();
    let properties = metadata.properties().unwrap();
    assert_eq!(properties.calculation_id(), 0);
    assert!(properties.full_calculation_on_load());
    assert!(properties.force_full_calculation());
    assert!(!properties.calculation_completed());
    assert!(properties.calculate_on_save());
}

fn native_parts_retained(base: &Package, after: &Package) {
    let a = members(base);
    let b = members(after);
    for (name, bytes) in a {
        if name != "[Content_Types].xml" && name != "xl/_rels/workbook.xml.rels" {
            assert!(
                b.get(&name) == Some(&bytes),
                "native member changed: {name}"
            );
        }
    }
    let mut actual = graph(after);
    for (name, expected) in graph(base) {
        let entry = actual.get_mut(&name).unwrap();
        if name == "/xl/workbook.xml" {
            entry
                .1
                .retain(|r| r.1 != custom_data::PROPERTIES_RELATIONSHIP_TYPE);
        }
        assert_eq!(entry, &expected);
    }
}

fn reopen(p: &Package) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    p.save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    assert!(members(p) == members(&q));
    q
}

fn custom_publish(p: &Package, c: &custom_data::Commit) -> Package {
    let q = reopen(c.package());
    equivalent(&p.apply_custom_data(c.patch()).unwrap().into_package(), &q);
    equivalent(
        &q.apply_custom_data(&c.patch().inverse())
            .unwrap()
            .into_package(),
        p,
    );
    q
}

fn storage() -> CustomData {
    let mut value = CustomData::new(STORAGE, b"\0native opaque payload\xff\r\n".to_vec());
    value.properties.extension_list = Some(ExtensionList {
        xml: format!(r#"<x:extLst xmlns:x="{X14}" xmlns:s="{SML}"><s:ext uri="urn:litchi:native-storage"><v:opaque xmlns:v="urn:inert" value="A &amp; B"><![CDATA[retained <text>]]></v:opaque></s:ext></x:extLst>"#).into_bytes(),
    });
    value
}

#[test]
fn native_xlsx_custom_data_catalog_edits_keep_opaque_payloads_and_native_parts() {
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.custom_data().unwrap().is_empty());
        let value = storage();
        let mut e = p.edit_custom_data().unwrap();
        let index = e.insert(value.clone()).unwrap();
        assert_eq!(index, 0);
        assert!(
            matches!(e.insert(value.clone()), Err(Error::Invalid(m)) if m == "Custom Data storage IDs must be unique")
        );
        assert_eq!(e.entries().unwrap()[0].value().to_owned(), value);
        let c = e.commit().unwrap();
        let added = custom_publish(&p, &c);
        native_parts_retained(&p, &added);
        let snapshot = added.custom_data().unwrap();
        let entry = snapshot.find(STORAGE).unwrap();
        assert_eq!(entry.value().to_owned(), value);
        assert_eq!(snapshot.len(), 1);
        let names = [
            entry.properties_part_name().clone(),
            entry.data_part_name().clone(),
        ];
        let a = members(&p);
        let b = members(&added);
        let mut added_names = b
            .keys()
            .filter(|n| !a.contains_key(*n))
            .cloned()
            .collect::<Vec<_>>();
        let mut expected_names = vec![
            names[0].as_str().trim_start_matches('/').to_owned(),
            names[1].as_str().trim_start_matches('/').to_owned(),
            names[0]
                .rels_uri()
                .unwrap()
                .as_str()
                .trim_start_matches('/')
                .to_owned(),
        ];
        added_names.sort();
        expected_names.sort();
        assert_eq!(added_names, expected_names);
        let mut e = added.edit_custom_data().unwrap();
        assert!(!e.set(0, value.clone()).unwrap());
        assert!(e.commit().unwrap().patch().is_noop());
        let mut expected = value.clone();
        expected.properties.id = "Renamed 日本語".into();
        expected.data.extend_from_slice(b"changed");
        let mut e = added.edit_custom_data().unwrap();
        assert!(e.rename_storage(STORAGE, &expected.properties.id).unwrap());
        assert!(e.set_data(0, expected.data.clone()).unwrap());
        let c = e.commit().unwrap();
        let changed = custom_publish(&added, &c);
        let snapshot = changed.custom_data().unwrap();
        let entry = snapshot.find(&expected.properties.id).unwrap();
        assert_eq!(entry.value().to_owned(), expected);
        assert_eq!(
            [entry.properties_part_name(), entry.data_part_name()],
            [&names[0], &names[1]]
        );
        only_changed(
            &added,
            &changed,
            &[
                names[0].as_str().trim_start_matches('/'),
                names[1].as_str().trim_start_matches('/'),
            ],
        );
        assert!(
            matches!(changed.apply_custom_data(c.patch()), Err(Error::PatchConflict { part }) if part == "Custom Data source closure")
        );
        let mut e = changed.edit_custom_data().unwrap();
        assert!(
            e.remove_storage(
                &expected.properties.id,
                RemovalDisposition::RejectReferenced
            )
            .unwrap()
        );
        let c = e.commit().unwrap();
        let removed = custom_publish(&changed, &c);
        assert!(removed.custom_data().unwrap().is_empty());
        equivalent(&removed, &p);
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}

fn replace_part(p: &Package, name: &str, bytes: Vec<u8>) -> Package {
    let mut opc = p.clone().into_plain_opc();
    opc.get_part_mut(&PackURI::new(format!("/{name}")).unwrap())
        .unwrap()
        .set_blob(bytes);
    Package::from_opc(opc).unwrap()
}

#[test]
fn native_xlsx_custom_data_reference_rename_retarget_and_detach_preserve_connections() {
    let file = "excel-connections.xlsx";
    let bytes = native(file);
    let native = Package::from_bytes(&bytes).unwrap();
    let mut e = native.edit_custom_data().unwrap();
    e.insert(CustomData::new("storage-a", b"inert a".to_vec()))
        .unwrap();
    e.insert(CustomData::new("storage-b", b"inert b".to_vec()))
        .unwrap();
    let c = e.commit().unwrap();
    let with_storage = custom_publish(&native, &c);
    native_parts_retained(&native, &with_storage);
    let original = String::from_utf8(members(&native)[CONNECTIONS].clone()).unwrap();
    assert_eq!(original.matches("</connection>").count(), 2);
    let extension = format!(
        r#"<extLst><ext uri="{{D79990A0-CA42-45E3-83F4-45C500A0EAA5}}"><x:connection xmlns:x="{X14}" embeddedDataId="storage-a" culture="en-US"/></ext></extLst>"#
    );
    let mut connected_xml = original.clone();
    // Bind the native type-5 OLE DB connection, leaving its type-6 text connection intact.
    connected_xml.insert_str(connected_xml.rfind("</connection>").unwrap(), &extension);
    let p = reopen(&replace_part(
        &with_storage,
        CONNECTIONS,
        connected_xml.as_bytes().to_vec(),
    ));
    only_changed(&with_storage, &p, &[CONNECTIONS]);
    assert_eq!(
        p.custom_data().unwrap().connection_references("storage-a"),
        1
    );
    let mut e = p.edit_custom_data().unwrap();
    assert!(e.rename_storage("storage-a", "storage-renamed").unwrap());
    assert!(
        matches!(e.remove_storage("storage-renamed", RemovalDisposition::RejectReferenced), Err(Error::CustomDataReferenced { id, connections: 1 }) if id == "storage-renamed")
    );
    assert!(
        matches!(e.rename_storage("storage-renamed", ""), Err(Error::Unsupported { feature }) if feature == "renaming a referenced Custom Data storage to the unreferencable empty UID")
    );
    let c = e.commit().unwrap();
    let renamed = custom_publish(&p, &c);
    let renamed_xml = connected_xml.replace(
        "embeddedDataId=\"storage-a\"",
        "embeddedDataId=\"storage-renamed\"",
    );
    assert_eq!(members(&renamed)[CONNECTIONS], renamed_xml.as_bytes());
    let mut e = renamed.edit_custom_data().unwrap();
    assert!(
        e.remove_storage(
            "storage-renamed",
            RemovalDisposition::RetargetConnections("storage-b".into())
        )
        .unwrap()
    );
    let c = e.commit().unwrap();
    let retargeted = custom_publish(&renamed, &c);
    assert_eq!(
        retargeted
            .custom_data()
            .unwrap()
            .connection_references("storage-b"),
        1
    );
    let retargeted_xml = connected_xml.replace(
        "embeddedDataId=\"storage-a\"",
        "embeddedDataId=\"storage-b\"",
    );
    assert_eq!(members(&retargeted)[CONNECTIONS], retargeted_xml.as_bytes());
    let mut e = retargeted.edit_custom_data().unwrap();
    assert!(
        e.remove_storage("storage-b", RemovalDisposition::DetachConnections)
            .unwrap()
    );
    let c = e.commit().unwrap();
    let detached = custom_publish(&retargeted, &c);
    assert!(detached.custom_data().unwrap().is_empty());
    assert_eq!(
        members(&detached)[CONNECTIONS],
        connected_xml
            .replace("embeddedDataId=\"storage-a\"", "embeddedDataId=\"\"")
            .as_bytes()
    );
    only_changed(
        &native,
        &detached,
        &[
            CONNECTIONS,
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ],
    );
    assert!(native.to_bytes().unwrap() == bytes);
    assert!(self::native(file) == bytes);
}

fn tags() -> Collection {
    let mut active = Tag::new(7).unwrap();
    active.set_xml_based(true);
    active
        .add_property(Property::new("label", "日本語 🍎 & <saved>").unwrap())
        .unwrap();
    let mut deleted = Tag::new(32768).unwrap();
    deleted.set_deleted(true);
    Collection::new(vec![
        Cell::new("C3", vec![deleted]).unwrap(),
        Cell::new("A1", vec![active]).unwrap(),
    ])
    .unwrap()
}

fn tags_publish(p: &Package, c: &smart_tags::Commit) -> Package {
    let q = reopen(c.package());
    assert!(members(&p.apply_smart_tags(c.patch()).unwrap().into_package()) == members(&q));
    assert!(
        members(
            &q.apply_smart_tags(&c.patch().inverse())
                .unwrap()
                .into_package()
        ) == members(p)
    );
    only_changed(p, &q, &[SHEET]);
    q
}

#[test]
fn native_xlsx_cell_smart_tag_lifecycle_preserves_worksheet_and_inert_native_parts() {
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.smart_tags(0usize).unwrap().is_none());
        let expected = tags();
        let mut e = p.edit_smart_tags(0usize).unwrap();
        e.replace(expected.clone()).unwrap();
        let c = e.commit().unwrap();
        let added = tags_publish(&p, &c);
        assert_eq!(added.smart_tags(0usize).unwrap(), Some(expected.clone()));
        assert_eq!(
            added
                .workbook()
                .sheet(0usize)
                .unwrap()
                .unwrap()
                .smart_tags()
                .unwrap(),
            Some(expected.clone())
        );
        let mut e = added.edit_smart_tags(0usize).unwrap();
        e.replace(expected.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = added.edit_smart_tags(0usize).unwrap();
        assert_eq!(
            e.remove("C3").unwrap().as_ref(),
            expected.get("C3").unwrap()
        );
        let mut tag = Tag::new(17).unwrap();
        tag.add_property(Property::new("updated", "value").unwrap())
            .unwrap();
        e.set(Cell::new("A1", vec![tag]).unwrap()).unwrap();
        let expected = e.collection().unwrap();
        let c = e.commit().unwrap();
        let updated = tags_publish(&added, &c);
        assert_eq!(updated.smart_tags(0usize).unwrap(), Some(expected));
        assert!(matches!(
            updated.apply_smart_tags(c.patch()),
            Err(Error::PatchConflict { .. })
        ));
        let mut e = updated.edit_smart_tags(0usize).unwrap();
        assert!(e.clear().unwrap());
        assert!(!e.clear().unwrap());
        let c = e.commit().unwrap();
        let cleared = tags_publish(&updated, &c);
        assert!(cleared.smart_tags(0usize).unwrap().is_none());
        assert!(members(&cleared) == members(&p));
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}

fn phonetic(p: &Package, attributes: &str) -> Package {
    let xml = String::from_utf8(members(p)[SHEET].clone()).unwrap();
    // These native first worksheets have pageMargins directly after sheetData;
    // phoneticPr belongs before pageMargins in the worksheet content model.
    assert_eq!(xml.matches("<pageMargins").count(), 1);
    let changed = xml.replacen(
        "<pageMargins",
        &format!("<phoneticPr {attributes}/><pageMargins"),
        1,
    );
    replace_part(p, SHEET, changed.into_bytes())
}

#[test]
fn native_xlsx_phonetic_properties_read_defaults_and_survive_cell_publication() {
    for file in PARENTS {
        let bytes = native(file);
        let base = Package::from_bytes(&bytes).unwrap();
        assert!(
            base.workbook()
                .sheet(0usize)
                .unwrap()
                .unwrap()
                .phonetic_properties()
                .unwrap()
                .is_none()
        );
        for (attributes, kind, alignment) in [
            (r#"fontId="0""#, Type::FullWidthKatakana, Alignment::Left),
            (
                r#"fontId="0" type="Hiragana" alignment="center""#,
                Type::Hiragana,
                Alignment::Center,
            ),
            (
                r#"fontId="0" type="halfwidthKatakana" alignment="distributed""#,
                Type::HalfWidthKatakana,
                Alignment::Distributed,
            ),
            (
                r#"fontId="0" type="noConversion" alignment="noControl""#,
                Type::NoConversion,
                Alignment::NoControl,
            ),
        ] {
            let p = reopen(&phonetic(&base, attributes));
            only_changed(&base, &p, &[SHEET]);
            assert!(
                p.workbook()
                    .sheet(0usize)
                    .unwrap()
                    .unwrap()
                    .cell("Z121")
                    .unwrap()
                    .is_missing()
            );
            let expected = p
                .workbook()
                .sheet(0usize)
                .unwrap()
                .unwrap()
                .phonetic_properties()
                .unwrap()
                .unwrap();
            assert_eq!(
                (
                    expected.font_id(),
                    expected.phonetic_type(),
                    expected.alignment()
                ),
                (0, kind, alignment)
            );
            let mut e = p.edit().unwrap();
            e.sheet(0usize).unwrap().unwrap().set("Z121", 72).unwrap();
            let c = e.commit().unwrap();
            let q = reopen(c.package());
            cell_edit_preserved(&p, &q);
            let sheet = q.workbook().sheet(0usize).unwrap().unwrap();
            assert_eq!(sheet.phonetic_properties().unwrap(), Some(expected));
            assert_eq!(
                sheet.cell("Z121").unwrap().stored(),
                Some(&litchi_xlsx::Cell::Value(litchi_xlsx::Value::Number(
                    litchi_xlsx::cell::Number::new("72").unwrap()
                )))
            );
            let restored = q.apply(&c.patch().inverse()).unwrap().into_package();
            // Restoring the removed calculation-chain relationship can normalize
            // its XML. Compare the entire effective graph and exact data parts.
            equivalent(&restored, &p);
        }
        assert!(base.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}

#[test]
fn native_xlsx_malformed_phonetic_properties_return_specific_reader_errors() {
    let bytes = native("excel-metadata.xlsx");
    let p = Package::from_bytes(&bytes).unwrap();
    for (attributes, message) in [
        (r#"type="Hiragana""#, "worksheet phoneticPr requires fontId"),
        (
            r#"fontId="0" type="unknown""#,
            "invalid worksheet phonetic type 'unknown'",
        ),
        (
            r#"fontId="0" alignment="unknown""#,
            "invalid worksheet phonetic alignment 'unknown'",
        ),
    ] {
        let malformed = phonetic(&p, attributes);
        assert!(
            matches!(malformed.workbook().sheet(0usize).unwrap().unwrap().phonetic_properties(), Err(Error::Invalid(m)) if m == message)
        );
    }
    assert!(p.to_bytes().unwrap() == bytes);
    assert!(native("excel-metadata.xlsx") == bytes);
}
