//! Explicit inert MapInfo catalogs on native Transitional, Strict and macro parents.
use litchi_opc::OpcPackage;
use litchi_xlsx::{
    Error, Package,
    xml_maps::{
        self, XmlMap, XmlMapConformance as Conf, XmlMapDataBinding, XmlMapInfo, XmlMapSchema,
    },
};
use std::{collections::BTreeMap, path::Path};
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
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
type Graph = BTreeMap<String, (String, Vec<(String, String, String, bool)>)>;
fn graph(p: &OpcPackage) -> Graph {
    let rels = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    let mut out = BTreeMap::from([("/".into(), (String::new(), rels(p.rels())))]);
    for p in p.iter_parts() {
        out.insert(
            p.partname().to_string(),
            (p.content_type().into(), rels(p.rels())),
        );
    }
    out
}
fn equivalent(a: &Package, b: &Package) {
    assert_eq!(
        graph(&a.clone().into_plain_opc()),
        graph(&b.clone().into_plain_opc())
    );
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (n, v) in a {
        if n != "[Content_Types].xml" && n != "xl/_rels/workbook.xml.rels" {
            assert!(b.get(&n) == Some(&v), "changed {n}");
        }
    }
}
fn preserved(base: &Package, after: &Package) {
    let a = members(base);
    let b = members(after);
    for (n, v) in a {
        if n != "[Content_Types].xml" && n != "xl/_rels/workbook.xml.rels" {
            assert!(b.get(&n) == Some(&v), "native member changed {n}");
        }
    }
    let a = graph(&base.clone().into_plain_opc());
    let mut b = graph(&after.clone().into_plain_opc());
    for (name, expected) in a {
        let actual = b.get_mut(&name).unwrap();
        actual.1.retain(|r| {
            r.1 != litchi_ooxml_common::spreadsheet_xml_maps::REL
                && r.1 != litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL
        });
        assert_eq!(actual, &expected);
    }
}
fn snapshot(p: &Package) -> xml_maps::Snapshot {
    xml_maps::Snapshot::load(&p.clone().into_plain_opc()).unwrap()
}
fn info(conf: Conf) -> XmlMapInfo {
    let ns = if conf.is_strict() {
        "http://purl.oclc.org/ooxml/spreadsheetml/main"
    } else {
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    };
    XmlMapInfo {
        selection_namespaces: "xmlns:n='urn:litchi:native-map:071'".into(),
        schemas: vec![XmlMapSchema {
            id: "schema-71".into(),
            schema_reference: None,
            namespace: Some("urn:litchi:native-map:071".into()),
            payload_xml: Some(format!(r#"<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns="{ns}" targetNamespace="urn:litchi:native-map:071"><xs:import namespace="urn:external" schemaLocation="https://example.invalid/never-fetched.xsd"/><xs:element name="root" type="xs:string"/></xs:schema>"#).into_bytes()),
        }],
        maps: vec![XmlMap {
            id: 71,
            name: "Native Map 🍎".into(),
            root_element: "root".into(),
            schema_id: "schema-71".into(),
            show_import_export_validation_errors: true,
            auto_fit: false,
            append: false,
            preserve_sort_auto_filter_layout: true,
            preserve_format: true,
            data_binding: Some(XmlMapDataBinding {
                data_binding_name: Some("binding-71".into()),
                file_binding: Some(false),
                connection_id: None,
                file_binding_name: None,
                load_mode: 1,
                // The read view completes this child with the MapInfo default namespace.
                payload_xml: Some(br#"<n:retained xmlns:n="urn:litchi:binding" value="A &amp; B"/>"#.to_vec()),
            }),
        }],
    }
}
fn projected(conf: Conf) -> XmlMapInfo {
    let mut value = info(conf);
    let ns = if conf.is_strict() {
        "http://purl.oclc.org/ooxml/spreadsheetml/main"
    } else {
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    };
    value.maps[0].data_binding.as_mut().unwrap().payload_xml = Some(
        format!(r#"<n:retained xmlns:n="urn:litchi:binding" value="A &amp; B" xmlns="{ns}"/>"#)
            .into_bytes(),
    );
    value
}
fn publish(p: &Package, c: &xml_maps::Commit) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    c.package().save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    assert!(members(&q) == members(c.package()));
    equivalent(&p.apply_xml_maps(c.patch()).unwrap().into_package(), &q);
    equivalent(
        &q.apply_xml_maps(&c.patch().inverse())
            .unwrap()
            .into_package(),
        p,
    );
    q
}
#[test]
fn native_xlsx_xml_map_catalog_lifecycle_keeps_cells_parts_and_opaque_payloads() {
    for (file, conf) in [
        ("excel-metadata.xlsx", Conf::Transitional),
        ("excel-recorded-macro.xlsm", Conf::Transitional),
        ("excel-visual-analytics-strict.xlsx", Conf::Strict),
    ] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(snapshot(&p).is_empty());
        let value = info(conf);
        let mut e = p
            .edit_xml_maps_with_options(xml_maps::EditOptions::default().conformance(conf))
            .unwrap();
        assert!(e.set(value.clone()).unwrap());
        let c = e.commit().unwrap();
        let added = publish(&p, &c);
        let old = snapshot(&added);
        assert_eq!(old.info(), Some(&projected(conf)));
        assert!(
            std::str::from_utf8(old.source_xml().unwrap())
                .unwrap()
                .contains(
                    std::str::from_utf8(
                        value.maps[0]
                            .data_binding
                            .as_ref()
                            .unwrap()
                            .payload_xml
                            .as_ref()
                            .unwrap()
                    )
                    .unwrap()
                )
        );
        assert_eq!(old.conformance(), conf);
        preserved(&p, &added);
        let mut e = added.edit_xml_maps().unwrap();
        assert!(!e.set(old.info().unwrap().clone()).unwrap());
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = added.edit_xml_maps().unwrap();
        e.set(value.clone()).unwrap();
        let c = e.commit().unwrap();
        assert!(c.patch().is_noop());
        assert!(members(c.package()) == members(&added));
        let mut e = added.edit_xml_maps().unwrap();
        e.edit_map(71, |m| {
            m.name = "Renamed Map 日本語".into();
            m.auto_fit = true;
            m.data_binding.as_mut().unwrap().load_mode = 2;
            Ok(())
        })
        .unwrap();
        let expected = e.info().unwrap().unwrap().clone();
        let c = e.commit().unwrap();
        let updated = publish(&added, &c);
        let new = snapshot(&updated);
        assert_eq!(new.info(), Some(&expected));
        assert_eq!(new.part_name(), old.part_name());
        assert_eq!(new.info().unwrap().schemas, value.schemas);
        assert_eq!(
            new.info().unwrap().maps[0]
                .data_binding
                .as_ref()
                .unwrap()
                .payload_xml,
            old.info().unwrap().maps[0]
                .data_binding
                .as_ref()
                .unwrap()
                .payload_xml
        );
        preserved(&p, &updated);
        let a = members(&added);
        let b = members(&updated);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        assert_eq!(
            a.keys()
                .filter(|n| a.get(*n) != b.get(*n))
                .map(String::as_str)
                .collect::<Vec<_>>(),
            [old.part_name().unwrap().as_str().trim_start_matches('/')]
        );
        assert!(
            matches!(updated.apply_xml_maps(c.patch()),Err(Error::Invalid(m)) if m=="custom XML maps patch source is stale")
        );
        let mut e = updated.edit_xml_maps().unwrap();
        let mut schema = value.schemas[0].clone();
        schema.id = "schema-extra".into();
        e.insert_schema(schema).unwrap();
        let mut map = value.maps[0].clone();
        map.id = 72;
        map.name = "Extra map".into();
        map.schema_id = "schema-extra".into();
        e.insert_map(map).unwrap();
        let expanded_draft = e.info().unwrap().unwrap().clone();
        assert!(
            matches!(e.remove_schema("schema-extra"),Err(Error::Common(litchi_ooxml_common::Error::SpreadsheetXmlMaps(m))) if m=="Map references an unknown SchemaID")
        );
        assert_eq!(e.info().unwrap(), Some(&expanded_draft));
        let c = e.commit().unwrap();
        let expanded = publish(&updated, &c);
        assert_eq!(snapshot(&expanded).info().unwrap().maps.len(), 2);
        preserved(&p, &expanded);
        let mut e = expanded.edit_xml_maps().unwrap();
        e.edit(|info| {
            info.schemas.reverse();
            info.maps.reverse();
            Ok(())
        })
        .unwrap();
        let c = e.commit().unwrap();
        let reordered = publish(&expanded, &c);
        assert_eq!(snapshot(&reordered).info().unwrap().maps[0].id, 72);
        let mut e = reordered.edit_xml_maps().unwrap();
        e.remove_map(72).unwrap().unwrap();
        e.remove_schema("schema-extra").unwrap().unwrap();
        let c = e.commit().unwrap();
        let contracted = publish(&reordered, &c);
        assert_eq!(snapshot(&contracted).info(), Some(&expected));
        let mut e = contracted.edit_xml_maps().unwrap();
        e.remove().unwrap().unwrap();
        let c = e.commit().unwrap();
        let cleared = publish(&contracted, &c);
        assert!(snapshot(&cleared).is_empty());
        preserved(&p, &cleared);
        assert_eq!(
            members(&cleared).keys().collect::<Vec<_>>(),
            members(&p).keys().collect::<Vec<_>>()
        );
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_xlsx_xml_map_rejected_schema_removal_and_duplicate_id_keep_valid_staging() {
    let p = Package::from_bytes(&native("excel-metadata.xlsx")).unwrap();
    let mut e = p.edit_xml_maps().unwrap();
    let value = info(Conf::Transitional);
    e.set(value.clone()).unwrap();
    let error = e.remove_schema("schema-71").unwrap_err();
    assert!(
        matches!(error,Error::Common(litchi_ooxml_common::Error::SpreadsheetXmlMaps(m)) if m=="MapInfo requires 1..4096 Schema children")
    );
    assert_eq!(e.info().unwrap(), Some(&value));
    let mut duplicate = value.maps[0].clone();
    duplicate.name = "Duplicate ID".into();
    assert!(
        matches!(e.insert_map(duplicate),Err(Error::Common(litchi_ooxml_common::Error::SpreadsheetXmlMaps(m))) if m=="duplicate Map ID")
    );
    assert_eq!(e.info().unwrap(), Some(&value));
    assert!(
        matches!(e.edit_map(71,|m|{m.data_binding.as_mut().unwrap().file_binding=Some(true);Ok(())}),Err(Error::Common(litchi_ooxml_common::Error::SpreadsheetXmlMaps(m))) if m=="ConnectionID is required when FileBinding is true")
    );
    assert_eq!(e.info().unwrap(), Some(&value));
    let c = e.commit().unwrap();
    let q = publish(&p, &c);
    assert_eq!(snapshot(&q).info(), Some(&projected(Conf::Transitional)));
    preserved(&p, &q);
}

#[test]
fn native_xlsx_xml_map_conformance_change_preserves_opaque_namespace_context() {
    let p = Package::from_bytes(&native("excel-metadata.xlsx")).unwrap();
    let mut e = p.edit_xml_maps().unwrap();
    e.set(info(Conf::Transitional)).unwrap();
    let c = e.commit().unwrap();
    let added = publish(&p, &c);
    let before = snapshot(&added);
    let mut e = added.edit_xml_maps().unwrap();
    e.set_conformance(Conf::Strict).unwrap();
    let c = e.commit().unwrap();
    let strict = publish(&added, &c);
    assert_eq!(snapshot(&strict).info(), before.info());
    assert_eq!(snapshot(&strict).conformance(), Conf::Strict);
    preserved(&p, &strict);
    let mut e = strict.edit_xml_maps().unwrap();
    e.set_conformance(Conf::Transitional).unwrap();
    let c = e.commit().unwrap();
    let restored = publish(&strict, &c);
    assert_eq!(snapshot(&restored).info(), before.info());
    assert_eq!(snapshot(&restored).conformance(), Conf::Transitional);
    preserved(&p, &restored);
}
